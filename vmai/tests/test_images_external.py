"""Công cụ đo bộ ảnh ngoài (Q96): đo mọi ca, không ghi vào vmai/out, không in tên ca. Token cấy, không OCR."""
from __future__ import annotations

import hashlib
import json
import shutil
import sys
from pathlib import Path

import pytest
from test_images import _template, _tokens

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "eval"))
import run_images_q90 as tool  # noqa: E402

SECRET = "ca_bi_mat_xyz"


def _make_set(root: Path) -> None:
    """Một ca có tên không thuộc loại ảnh nào đã biết."""
    case = root / "images" / SECRET
    case.mkdir(parents=True)
    _template(case / "template.xlsx")
    (case / "page.jpg").write_bytes(b"")
    cells = [{"a1": f"B{r}", "value": "5.10"} for r in range(5, 10)] + [{"a1": f"C{r}", "value": "16.20"} for r in range(5, 10)]
    body = json.dumps({"layout": "L1", "variant": "kieu_moi", "cells": cells})
    (case / "answer.json").write_text(body, encoding="utf-8")
    (root / "images" / "manifest.json").write_text(json.dumps({SECRET: hashlib.sha256(body.encode()).hexdigest()}), encoding="utf-8")


def _cache(out: Path) -> None:
    tokens = _tokens(headers=["Ngày", "Glucose Mức 1", "Glucose Mức 2"])
    (out / "ocr_cache").mkdir(parents=True)
    data = {"readers": ["ocr:a@1", "ocr:b@1"], "first": tokens, "second": tokens, "rotation": 0, "low_quality": False, "sharpness": 0, "brightness": 0, "seconds": [0, 0]}
    (out / "ocr_cache" / f"{SECRET}.json").write_text(json.dumps(data), encoding="utf-8")


def test_external_set_is_scored_and_prints_totals_only(tmp_path, capsys, monkeypatch):
    root, out, fake_out = tmp_path / "bo", tmp_path / "kq", tmp_path / "vmai_out"
    fake_out.mkdir()
    monkeypatch.setattr(tool, "OUT", fake_out)
    _make_set(root)
    _cache(out)
    code = tool.measure_external(root, out, "score")
    printed = capsys.readouterr().out
    assert code == 0
    assert SECRET not in printed and "kieu_moi" in printed and "L1" in printed
    assert "Đáp án đổi so với manifest: 0" in printed
    assert list(fake_out.iterdir()) == []
    result = json.loads((out / "eval_external.json").read_text(encoding="utf-8"))
    assert result["total"]["images"] == 1 and result["total"]["gold"] == 10 and result["total"]["green"] == 0


def test_changed_answer_is_counted(tmp_path, capsys, monkeypatch):
    root, out = tmp_path / "bo", tmp_path / "kq"
    monkeypatch.setattr(tool, "OUT", tmp_path / "vmai_out")
    _make_set(root)
    _cache(out)
    answer = root / "images" / SECRET / "answer.json"
    answer.write_text(answer.read_text(encoding="utf-8") + " ", encoding="utf-8")
    tool.measure_external(root, out, "score")
    assert "Đáp án đổi so với manifest: 1" in capsys.readouterr().out


def test_external_results_cannot_go_into_vmai_out(tmp_path, monkeypatch):
    root = tmp_path / "bo"
    monkeypatch.setattr(tool, "OUT", tmp_path / "vmai_out")
    _make_set(root)
    with pytest.raises(SystemExit):
        tool.measure_external(root, tmp_path / "vmai_out" / "x", "score")
    shutil.rmtree(root)


def test_model_profile_runs_on_an_external_set_and_records_cost(tmp_path, capsys, monkeypatch):
    root, out = tmp_path / "bo", tmp_path / "kq"
    monkeypatch.setattr(tool, "OUT", tmp_path / "vmai_out")
    _make_set(root)
    _cache(out)
    assert tool.measure_external(root, out, "score", "fake", True) == 0
    printed = capsys.readouterr().out
    assert "Model:" in printed and "bảng giá" in printed and SECRET not in printed
    model = json.loads((out / "eval_external.json").read_text(encoding="utf-8"))["model"]
    assert model["profile"] == "fake" and model["confirmed_network"] is True and {"calls", "tokens", "usd_by_price_table", "headers_would_send"} <= set(model)


def test_without_confirm_network_nothing_is_called(tmp_path, capsys, monkeypatch):
    root, out = tmp_path / "bo", tmp_path / "kq"
    monkeypatch.setattr(tool, "OUT", tmp_path / "vmai_out")
    _make_set(root)
    _cache(out)
    assert tool.measure_external(root, out, "score", "fake", False) == 0
    model = json.loads((out / "eval_external.json").read_text(encoding="utf-8"))["model"]
    assert model["confirmed_network"] is False and model["calls"] == 0
