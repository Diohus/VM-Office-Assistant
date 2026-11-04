"""Chạy: python -m pytest tools/snapshot -q"""
from __future__ import annotations

import pathlib
import zipfile

import pytest

import snapshot


def tree(tmp_path, monkeypatch):
    root = tmp_path / "repo"
    (root / "src").mkdir(parents=True)
    (root / "src" / "a.py").write_text("a")
    (root / ".venv").mkdir()
    (root / ".venv" / "x.py").write_text("x")
    (root / ".env").write_text("SECRET=1")
    (root / ".env.example").write_text("X=")
    (root / "~$khoa.xlsx").write_text("k")
    (root / ".local").mkdir()
    (root / ".local" / "machine-ca.pem").write_text("c")
    (root / "fixtures" / "generated").mkdir(parents=True)
    (root / "fixtures" / "generated" / "lon.xlsx").write_text("g")
    monkeypatch.setattr(snapshot, "ROOT", root)
    return root


def names(path: pathlib.Path) -> set[str]:
    with zipfile.ZipFile(path) as z:
        return set(z.namelist())


def test_loai_bi_mat_va_file_sinh_ra(tmp_path, monkeypatch):
    tree(tmp_path, monkeypatch)
    z = snapshot.create(tmp_path / "out")
    assert names(z) == {"src/a.py", ".env.example", snapshot.MANIFEST}


def test_khong_cho_de_trong_repo_hay_cloud(tmp_path, monkeypatch):
    root = tree(tmp_path, monkeypatch)
    with pytest.raises(SystemExit):
        snapshot.create(root / "bk")
    with pytest.raises(SystemExit):
        snapshot.create(tmp_path / "OneDrive" / "bk")


def test_diff_va_khoi_phuc(tmp_path, monkeypatch):
    root = tree(tmp_path, monkeypatch)
    a = snapshot.create(tmp_path / "out1")
    (root / "src" / "a.py").write_text("b")
    (root / "src" / "moi.py").write_text("n")
    b = snapshot.create(tmp_path / "out2")
    assert snapshot.diff(a, b) == {"them": ["src/moi.py"], "xoa": [], "doi": ["src/a.py"]}
    target = tmp_path / "khoi_phuc"
    assert snapshot.restore(b, target) == 3
    assert (target / "src" / "a.py").read_text() == "b"
    with pytest.raises(SystemExit):
        snapshot.restore(a, target)


def test_ban_chup_hong_bi_phat_hien(tmp_path, monkeypatch):
    tree(tmp_path, monkeypatch)
    z = snapshot.create(tmp_path / "out")
    bad = tmp_path / "bad.zip"
    with zipfile.ZipFile(z) as zi, zipfile.ZipFile(bad, "w") as zo:
        for n in zi.namelist():
            zo.writestr(n, b"sai" if n == "src/a.py" else zi.read(n))
    with pytest.raises(SystemExit):
        snapshot.restore(bad, tmp_path / "t")
