# -*- coding: utf-8 -*-
"""§15 pipeline 编排器回归测试：传播规则 / hash 检测 / 层依赖。"""
import json
import tempfile
from pathlib import Path

from conftest import load_script

pipe = load_script("exam_pipeline")


def test_propagation_rules():
    order = ["design", "content", "render"]
    # 改 L1 -> 全脏
    assert pipe.propagation(order, ["design"]) == {"design": True, "content": True, "render": True}
    # 改 L2 -> L2+L3
    assert pipe.propagation(order, ["content"]) == {"design": False, "content": True, "render": True}
    # 改 L3 -> 仅 L3
    assert pipe.propagation(order, ["render"]) == {"design": False, "content": False, "render": True}
    # 无改动
    assert pipe.propagation(order, []) == {"design": False, "content": False, "render": False}
    # 多处脏取最早
    assert pipe.propagation(order, ["content", "render"])["design"] is False


def test_hash_detects_change():
    with tempfile.TemporaryDirectory() as td:
        p = Path(td)
        (p / "a.json").write_text('{"answer": "B"}', encoding="utf-8")
        config = {
            "layer_order": ["design", "content", "render"],
            "design": {"inputs": []},
            "content": {"inputs": ["a.json"]},
            "render": {"inputs": []},
        }
        h1 = pipe._hash_inputs(config, str(p))
        (p / "a.json").write_text('{"answer": "C"}', encoding="utf-8")
        h2 = pipe._hash_inputs(config, str(p))
        assert h1["content"] != h2["content"]
        assert h1["design"] == h2["design"]


def test_config_workdir_resolution():
    with tempfile.TemporaryDirectory() as td:
        p = Path(td)
        (p / "pipeline.json").write_text(json.dumps({
            "layer_order": ["design", "content", "render"],
            "workdir": ".", "manifest": ".manifest.json",
            "design": {"inputs": []}, "content": {"inputs": []}, "render": {"inputs": []},
        }), encoding="utf-8")
        cfg = pipe.load_config(p / "pipeline.json")
        assert Path(cfg["_workdir"]).resolve() == p.resolve()
        assert cfg["_manifest"].endswith(".manifest.json")


def test_resolve_placeholders():
    cfg = {
        "build": [["{python}", "build.py", "{soffice}", "x.docx"]],
        "gates": [],
    }
    out = pipe._resolve_cmds(cfg, "/usr/bin/python3", "/opt/soffice")
    assert out["build"][0][0] == "/usr/bin/python3"
    assert out["build"][0][2] == "/opt/soffice"


def test_pdf_page_gate():
    gate = load_script("check_pdf_pages")
    # 无依赖时不应崩，返回 ERROR 而非异常
    rep = gate.audit("/nonexistent.pdf")
    assert rep["verdict"] in ("ERROR", "FAIL")
