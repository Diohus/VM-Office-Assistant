"""OCR chạy nội bộ và chuẩn bị ảnh. Cần `vmai[ocr]`. Không gửi gì ra ngoài.

RapidOCR đọc số tốt, Tesseract đọc tiếng Việt tốt hơn và độc lập về mô hình (Q25). Độc lập phải đo, không giả định.
Tesseract là chương trình riêng: đường dẫn lấy từ biến môi trường VMAI_TESSERACT, thư mục dữ liệu từ VMAI_TESSDATA.
Khi đo trên máy không chạy trực tiếp được, VMAI_TESSERACT_DOCKER là tên image, chạy không mạng (Q83).
Engine không chạy được thì báo OCR_ENGINE_FAILED, không tụt im lặng về một lần đọc.
"""
from __future__ import annotations

import csv
import hashlib
import io
import os
import subprocess
import tempfile
from dataclasses import dataclass
from pathlib import Path

from .errors import SourceFileError

# Ngưỡng tạm, cần hiệu chuẩn với ảnh chụp thật. Thang độ sắc nét phụ thuộc độ phân giải.
MIN_SHARPNESS = 25.0
MIN_BRIGHTNESS = 60.0
_PROJECTION_LIMIT = 10.0  # độ
_PROJECTION_STEP = 0.25
_PROJECTION_WIDTH = 1200
_PROJECTION_MIN_GAIN = 1.2  # giả định: phải rõ hơn hẳn so với không xoay mới xoay
_MAX_DIM = 2600
_LOCAL_BACKGROUND = 51
_STROKE_CONTRAST = 30
_HEIC_BRANDS = (b"heic", b"heix", b"hevc", b"heim", b"heis", b"mif1", b"msf1")


@dataclass(frozen=True)
class PreparedPage:
    rgb: object  # numpy RGB
    width: int
    height: int
    sha256: str
    rotation: int
    skew_degrees: float
    sharpness: float
    brightness: float
    edge_ratio: float = 99.0

    @property
    def low_quality(self) -> bool:
        # `edge_ratio` cố ý không dùng ở đây (Q102): trên bộ độc lập nó bật ở ảnh điện thoại, không bật ở ảnh độ phân giải thấp.
        return self.sharpness < MIN_SHARPNESS or self.brightness < MIN_BRIGHTNESS


def load_image(path: Path | str):
    """Đọc ảnh, xoay theo EXIF rồi bỏ EXIF (gồm cả vị trí GPS) bằng cách chép lại điểm ảnh."""
    import numpy as np
    from PIL import Image, ImageOps

    data = Path(path).read_bytes()
    if data[4:8] == b"ftyp" and data[8:12] in _HEIC_BRANDS:
        raise SourceFileError("IMAGE_HEIC_UNSUPPORTED", "Ảnh .heic chưa đọc được. Nhờ xuất sang JPEG rồi tải lại.")
    try:
        with Image.open(io.BytesIO(data)) as image:
            return np.array(ImageOps.exif_transpose(image).convert("RGB"))
    except Exception as exc:
        raise SourceFileError("IMAGE_CORRUPT", f"Không mở được ảnh: {type(exc).__name__}") from exc


def _upright_gray(gray):
    """Đưa về chữ tối trên nền sáng. So từng điểm với nền ngay quanh nó, không xét độ sáng cả ảnh:
    giấy chụp thiếu sáng hay có bóng đổ cũng tối nhưng chữ vẫn tối hơn nền, đảo màu sẽ làm hỏng."""
    import cv2
    import numpy as np

    difference = gray.astype(np.int16) - cv2.medianBlur(gray, _LOCAL_BACKGROUND)
    return 255 - gray if (difference > _STROKE_CONTRAST).sum() > (difference < -_STROKE_CONTRAST).sum() else gray


def _skew(rgb) -> float:
    import cv2
    import numpy as np

    gray = cv2.cvtColor(rgb, cv2.COLOR_RGB2GRAY)
    lines = cv2.HoughLinesP(cv2.Canny(gray, 50, 150), 1, np.pi / 720, threshold=120, minLineLength=gray.shape[1] // 3, maxLineGap=20)
    if lines is None:
        return 0.0
    angles = [a for x1, y1, x2, y2 in lines.reshape(-1, 4) if abs(a := float(np.degrees(np.arctan2(y2 - y1, x2 - x1)))) < 15]
    return float(np.median(angles)) if angles else 0.0


def _text_line_skew(rgb) -> float:
    """Độ nghiêng khi không có đường kẻ: góc làm các hàng chữ tách rõ nhất. Không chắc thì trả 0, không xoay bừa."""
    import cv2
    import numpy as np

    gray = cv2.cvtColor(rgb, cv2.COLOR_RGB2GRAY)
    h, w = gray.shape
    if w > _PROJECTION_WIDTH:
        gray = cv2.resize(gray, (_PROJECTION_WIDTH, int(h * _PROJECTION_WIDTH / w)), interpolation=cv2.INTER_AREA)
        h, w = gray.shape
    ink = (_upright_gray(gray) < gray.mean() - gray.std() * 0.5).astype(np.uint8) * 255

    def score(angle: float) -> float:
        turned = cv2.warpAffine(ink, cv2.getRotationMatrix2D((w / 2, h / 2), angle, 1.0), (w, h), flags=cv2.INTER_LINEAR)
        profile = turned.sum(axis=1, dtype=np.float64)
        return float(((profile[1:] - profile[:-1]) ** 2).sum())

    flat = score(0.0)
    best, best_score = 0.0, flat
    for angle in np.arange(-_PROJECTION_LIMIT, _PROJECTION_LIMIT + _PROJECTION_STEP, _PROJECTION_STEP):
        current = score(float(angle))
        if current > best_score:
            best, best_score = float(angle), current
    return best if best_score >= flat * _PROJECTION_MIN_GAIN else 0.0


def _edge_ratio(gray) -> float:
    """Độ dốc lớn nhất của nét so với biên độ sáng tối. Ảnh mờ, ảnh phóng to từ ảnh nhỏ có nét chuyển chậm dù tương phản còn."""
    import cv2
    import numpy as np

    gradient = cv2.magnitude(cv2.Sobel(gray, cv2.CV_64F, 1, 0), cv2.Sobel(gray, cv2.CV_64F, 0, 1))
    span = max(float(np.percentile(gray, 99.5)) - float(np.percentile(gray, 0.5)), 1.0)
    return float(np.percentile(gradient, 99.9)) / span


def prepare_image(path: Path | str, orientation_scorer=None) -> PreparedPage:
    """`orientation_scorer(rgb) -> (điểm, chắc chắn)`: chọn hướng có điểm cao nhất, chắc chắn thì dừng sớm vì OCR chậm.

    Sửa nghiêng trước khi xét hướng: đường kẻ bảng theo cả hai chiều nên đo được ở mọi hướng, và OCR đọc ảnh thẳng tốt hơn.
    """
    import cv2
    import numpy as np

    rgb = load_image(path)
    h, w = rgb.shape[:2]
    scale = _MAX_DIM / max(h, w)
    if scale < 1:
        rgb = cv2.resize(rgb, (int(w * scale), int(h * scale)), interpolation=cv2.INTER_AREA)
    angle = _skew(rgb)
    if abs(angle) >= 0.15:
        h, w = rgb.shape[:2]
        rgb = cv2.warpAffine(rgb, cv2.getRotationMatrix2D((w / 2, h / 2), angle, 1.0), (w, h), flags=cv2.INTER_CUBIC, borderValue=(255, 255, 255))
    else:
        angle = 0.0
    turns = 0
    if orientation_scorer is not None:
        scored = []
        for t in (0, 2, 1, 3):
            score, sure = orientation_scorer(np.rot90(rgb, t).copy())
            scored.append((score, t))
            if sure:
                break
        turns = max(scored, key=lambda item: item[0])[1]
    if turns:
        rgb = np.rot90(rgb, turns).copy()
    if angle == 0.0:
        # Bảng không kẻ không có đường để đo. Xét sau khi đã đúng hướng, vì hàng chữ chỉ rõ khi nằm ngang.
        angle = _text_line_skew(rgb)
        if abs(angle) >= 0.15:
            h, w = rgb.shape[:2]
            rgb = cv2.warpAffine(rgb, cv2.getRotationMatrix2D((w / 2, h / 2), angle, 1.0), (w, h), flags=cv2.INTER_CUBIC, borderValue=(255, 255, 255))
        else:
            angle = 0.0
    gray = _upright_gray(cv2.cvtColor(rgb, cv2.COLOR_RGB2GRAY))
    ok, encoded = cv2.imencode(".png", rgb[:, :, ::-1])
    return PreparedPage(
        rgb=rgb,
        width=rgb.shape[1],
        height=rgb.shape[0],
        sha256=hashlib.sha256(encoded.tobytes()).hexdigest() if ok else "",
        rotation=turns * 90,
        skew_degrees=angle,
        sharpness=float(cv2.Laplacian(gray, cv2.CV_64F).var()),
        brightness=float(gray.mean()),
        edge_ratio=_edge_ratio(gray),
    )


class RapidOcr:
    name = "rapidocr"

    def __init__(self) -> None:
        from importlib.metadata import version

        try:
            from rapidocr_onnxruntime import RapidOCR
        except ImportError as exc:
            raise SourceFileError("OCR_NOT_INSTALLED", "Chưa cài RapidOCR") from exc

        self.reader_id = f"ocr:rapidocr@{version('rapidocr-onnxruntime')}"
        self._ocr = RapidOCR()

    def read(self, rgb) -> list[dict]:
        result, _ = self._ocr(rgb[:, :, ::-1])
        out = []
        for box, text, _score in result or []:
            xs, ys = [p[0] for p in box], [p[1] for p in box]
            out.append({"text": text, "bbox": [float(min(xs)), float(min(ys)), float(max(xs)), float(max(ys))]})
        return out


class TesseractOcr:
    name = "tesseract"

    def __init__(self, languages: str = "vie+eng") -> None:
        self._docker = os.environ.get("VMAI_TESSERACT_DOCKER")
        exe = os.environ.get("VMAI_TESSERACT")
        self._tessdata = os.environ.get("VMAI_TESSDATA")
        if not self._docker and (not exe or not Path(exe).exists()):
            raise SourceFileError("OCR_NOT_INSTALLED", "Chưa cấu hình Tesseract")
        self._exe, self._languages = exe, languages
        out = self._run(["--version"], None).decode("utf-8", "replace")
        self.reader_id = f"ocr:tesseract@{out.split()[1]}"

    def _run(self, args: list[str], mount: str | None) -> bytes:
        if self._docker:
            prefix = ["docker", "run", "--rm", "--network", "none"]
            if mount:
                prefix += ["-v", f"{mount}:/work:ro"]
            command = [*prefix, self._docker, "tesseract", *args]
        else:
            command = [self._exe, *args]
        try:
            done = subprocess.run(command, capture_output=True, check=True, env={**os.environ, "OMP_THREAD_LIMIT": "2"})
        except (subprocess.CalledProcessError, OSError) as exc:
            # Không kèm đầu ra của engine: có thể chứa nội dung tài liệu.
            raise SourceFileError("OCR_ENGINE_FAILED", f"Tesseract không chạy được: {type(exc).__name__}") from exc
        return done.stdout + done.stderr if args == ["--version"] else done.stdout

    @staticmethod
    def _clean(rgb):
        """Cân lại độ sáng nền, xóa đường kẻ. Không phóng to vì đo thấy kém hơn. Chữ sáng trên nền tối thì đảo màu trước."""
        import cv2
        import numpy as np

        gray = _upright_gray(cv2.cvtColor(rgb, cv2.COLOR_RGB2GRAY))
        background = cv2.morphologyEx(gray, cv2.MORPH_CLOSE, cv2.getStructuringElement(cv2.MORPH_ELLIPSE, (41, 41)))
        flat = cv2.divide(gray, background, scale=255)
        binary = cv2.adaptiveThreshold(flat, 255, cv2.ADAPTIVE_THRESH_MEAN_C, cv2.THRESH_BINARY_INV, 31, 15)
        horizontal = cv2.morphologyEx(binary, cv2.MORPH_OPEN, cv2.getStructuringElement(cv2.MORPH_RECT, (80, 1)))
        vertical = cv2.morphologyEx(binary, cv2.MORPH_OPEN, cv2.getStructuringElement(cv2.MORPH_RECT, (1, 30)))
        flat[cv2.dilate(horizontal | vertical, np.ones((3, 3), np.uint8)) > 0] = 255
        return flat

    def read(self, rgb) -> list[dict]:
        from PIL import Image

        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "page.png"
            Image.fromarray(self._clean(rgb)).save(path)
            target = "/work/page.png" if self._docker else str(path)
            args = [target, "stdout", "-l", self._languages, "--psm", "11", "-c", "tessedit_create_tsv=1"]
            if self._tessdata and not self._docker:
                args[2:2] = ["--tessdata-dir", self._tessdata]
            out = self._run(args, tmp if self._docker else None)
        tokens = []
        for row in csv.DictReader(io.StringIO(out.decode("utf-8", "replace")), delimiter="\t", quoting=csv.QUOTE_NONE):
            text = (row.get("text") or "").strip()
            if text and row.get("level") == "5":
                x, y, w, h = int(row["left"]), int(row["top"]), int(row["width"]), int(row["height"])
                tokens.append({"text": text, "bbox": [float(x), float(y), float(x + w), float(y + h)]})
        return tokens
