"""Q117, T-AI-13 (c): mô tả đoán ra do backend lưu phải đọc chặt. Thiếu trường hoặc khác định dạng thì báo SPEC_OUTDATED,
backend coi như chưa có hồ sơ và hỏi lại. Dữ liệu giả."""
from __future__ import annotations

import json
from pathlib import Path

import pytest

import vmai
from vmai import infer_spec, load_spec, read_source
from vmai.errors import TemplateSpecError
from vmai.template import SPEC_FORMAT

ROOT = Path(__file__).resolve().parents[2]
L01 = ROOT / "fixtures" / "eval" / "generated" / "layouts" / "L01_hai_dong_tieu_de"


def _stored():
    guess = infer_spec(read_source(L01 / "template.xlsx"), read_source(L01 / "source.xlsx"))
    assert guess.spec is not None
    return json.loads(guess.spec.model_dump_json())  # đúng như backend lưu


def test_stored_spec_round_trips():
    spec = load_spec(_stored())
    assert spec.spec_format == SPEC_FORMAT and spec.origin == "inferred"


@pytest.mark.parametrize("field", ["spec_format", "user_chosen_fields", "level_answer_columns", "alternative_value_fields"])
def test_missing_field_is_refused(field):
    data = _stored()
    del data[field]
    with pytest.raises(TemplateSpecError) as err:
        load_spec(data)
    assert err.value.code == "SPEC_OUTDATED"


@pytest.mark.parametrize("version", [1, SPEC_FORMAT + 1, "2"])
def test_other_format_is_refused(version):
    data = _stored()
    data["spec_format"] = version
    with pytest.raises(TemplateSpecError) as err:
        load_spec(data)
    assert err.value.code == "SPEC_OUTDATED"


def test_declared_spec_files_are_not_affected():
    for path in (Path(vmai.__file__).parent / "specs").glob("*.yaml"):
        assert load_spec(path).origin == "declared"
