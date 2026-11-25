# -*- coding: utf-8 -*-
"""试卷生成三层流水线编排器（exam-word-skill，跨课程通用）。

三层模型（顺序依赖，见 pitfalls §15）
------------------------------------
  L1 design  命题设计：知识点划分 / 双向细目 / 题量排序 / 分值配比 / A-B 平行
  L2 content 题目生成：题干 / 选项 / 答案 / 干扰项分布 / 无泄题 / 来源对齐
  L3 render  格式渲染：代码结构 / 模板规范 / 标签页 / 分页 / 缩进 / 字体

另有**正交门禁**（不随依赖传播，每层交付都要过）：内容安全、代码完整性、
A-B 一致性。它们由各层 gates 直接声明，本编排器不假设具体课程规则。

为什么需要
----------
真实事故（2026-11-24）：在 L2 改了答案分布后手工串联 6 步重生成，
没有任何机制保证下游（L3 代码排版等）被重新验证——靠人肉逐项检查，
漏一步就翻车。本脚本用「输入 hash + 层间传播 + fail-fast 门禁」把
「改一处 → 下游自动重跑核对」变成可执行的确定流程。

机制
----
1. 配置 pipeline.json 声明每层的 inputs（源文件 glob）、gates（门禁命令）、
   可选 build（生成命令），以及 layer_order。
2. `init` 记录各层 inputs 的 sha256 到 manifest。
3. `status` 对比当前 hash 与 manifest，标出变脏的层；按传播规则
   （L1 脏→全脏；L2 脏→L2+L3；L3 脏→L3）给出应从哪层重跑。
4. `run --from L2` 或 `run`（自动取最早脏层）依次执行每层的 build + gates，
   任一门禁非零退出即 fail-fast 停止，并写明失败层与命令。
5. `run` 成功后刷新 manifest（把当前 hash 记为新基线）。

设计原则
--------
* 配置驱动、课程无关：命令里用 {python} 占位当前解释器；路径相对 workdir。
* 门禁命令退出码即判据，编排器不解析自由文本，可靠且可移植。
* 只做编排，不复制任何生成/门禁逻辑——新增门禁只需改配置。

用法
----
    python exam_pipeline.py init    --config pipeline.json
    python exam_pipeline.py status  --config pipeline.json
    python exam_pipeline.py run     --config pipeline.json
    python exam_pipeline.py run     --config pipeline.json --from content
    python exam_pipeline.py run     --config pipeline.json --dry-run

配置格式见 assets/pipeline.template.json。
"""
import argparse
import glob
import hashlib
import json
import subprocess
import sys
from pathlib import Path

LAYER_ORDER_DEFAULT = ["design", "content", "render"]
LAYER_LABEL = {"design": "L1 命题设计", "content": "L2 题目生成", "render": "L3 格式渲染"}


def _hash_inputs(config, workdir):
    """返回 {layer: hash}；hash 覆盖每个匹配文件的内容+相对路径。"""
    result = {}
    for layer in config["layer_order"]:
        cfg = config[layer]
        h = hashlib.sha256()
        for pattern in cfg.get("inputs", []):
            for f in sorted(glob.glob(str(Path(workdir) / pattern), recursive=True)):
                p = Path(f)
                if p.is_file():
                    h.update(str(p.relative_to(workdir)).encode())
                    h.update(p.read_bytes())
        result[layer] = h.hexdigest()
    return result


def _resolve_cmds(cfg, python_exe, soffice=None):
    """把命令里的 {python} / {soffice} 占位符替换掉。"""
    soffice = soffice or "/Applications/LibreOffice.app/Contents/MacOS/soffice"
    def resolve(list_of_cmds):
        out = []
        for c in list_of_cmds:
            out.append([a.replace("{python}", python_exe).replace("{soffice}", soffice) for a in c])
        return out
    return {
        "build": resolve(cfg.get("build", [])),
        "gates": resolve(cfg.get("gates", [])),
    }


def load_config(path):
    cfg = json.loads(Path(path).read_text(encoding="utf-8-sig"))
    cfg.setdefault("layer_order", LAYER_ORDER_DEFAULT)
    # workdir 相对配置文件所在目录
    base = Path(path).resolve().parent
    cfg["_workdir"] = str((base / cfg.get("workdir", ".")).resolve())
    cfg["_manifest"] = str(base / cfg.get("manifest", ".pipeline-manifest.json"))
    return cfg


def propagation(order, dirty):
    """按依赖传播：某层脏则其所有下游也脏。返回 {layer: bool}。"""
    flags = {l: False for l in order}
    earliest = None
    for i, l in enumerate(order):
        if l in dirty and earliest is None:
            earliest = i
    if earliest is not None:
        for l in order[earliest:]:
            flags[l] = True
    return flags


def cmd_status(cfg, python_exe):
    order = cfg["layer_order"]
    cur = _hash_inputs(cfg, cfg["_workdir"])
    manifest_path = Path(cfg["_manifest"])
    old = json.loads(manifest_path.read_text(encoding="utf-8")).get("hashes", {}) if manifest_path.exists() else {}
    dirty = [l for l in order if old.get(l) != cur[l]]
    flags = propagation(order, dirty)
    print("工作目录:", cfg["_workdir"])
    print("基线:", "已建立" if manifest_path.exists() else "未建立（先跑 init）")
    print()
    for l in order:
        mark = "脏" if flags[l] else "净"
        direct = "（输入直接变化）" if l in dirty else ("（下游受影响）" if flags[l] else "")
        print(f"  {LAYER_LABEL.get(l,l):<12} {mark} {direct}")
    start = next((l for l in order if flags[l]), None)
    print()
    print("建议:", f"从 {LAYER_LABEL.get(start,start)} 重跑（python exam_pipeline.py run --from {start}）" if start else "全部干净，无需重跑")
    return 0


def cmd_init(cfg, python_exe):
    cur = _hash_inputs(cfg, cfg["_workdir"])
    Path(cfg["_manifest"]).write_text(
        json.dumps({"hashes": cur}, ensure_ascii=False, indent=2), encoding="utf-8")
    print("已建立基线 manifest:", cfg["_manifest"])
    for l, h in cur.items():
        print(f"  {LAYER_LABEL.get(l,l)}: {h[:12]}")
    return 0


def cmd_run(cfg, python_exe, from_layer=None, dry_run=False):
    order = cfg["layer_order"]
    cur = _hash_inputs(cfg, cfg["_workdir"])
    manifest_path = Path(cfg["_manifest"])
    old = json.loads(manifest_path.read_text(encoding="utf-8")).get("hashes", {}) if manifest_path.exists() else {}
    dirty = [l for l in order if old.get(l) != cur[l]]
    flags = propagation(order, dirty)
    if from_layer:
        if from_layer not in order:
            print(f"未知层: {from_layer}；应为 {order}", file=sys.stderr)
            return 2
        idx = order.index(from_layer)
        for l in order[idx:]:
            flags[l] = True

    print("=== 流水线执行计划 ===")
    for l in order:
        print(f"  {LAYER_LABEL.get(l,l)}: {'执行' if flags[l] else '跳过（未变脏）'}")
    if dry_run:
        print("（dry-run，不实际执行）")
        return 0

    workdir = Path(cfg["_workdir"])
    for l in order:
        if not flags[l]:
            continue
        cmds = _resolve_cmds(cfg[l], python_exe, cfg.get("soffice"))
        print(f"\n--- {LAYER_LABEL.get(l,l)} ---")
        for kind in ("build", "gates"):
            for c in cmds[kind]:
                print(f"  [{kind}] {' '.join(c)}")
                r = subprocess.run(c, cwd=workdir)
                if r.returncode != 0:
                    print(f"\n✘ 门禁失败 @ {LAYER_LABEL.get(l,l)}（退出码 {r.returncode}）")
                    print(f"  命令: {' '.join(c)}")
                    print("  已停止。修正后重跑；下游不会在失败态下继续。", file=sys.stderr)
                    return 1
        print(f"  ✔ {LAYER_LABEL.get(l,l)} 通过")

    Path(cfg["_manifest"]).write_text(
        json.dumps({"hashes": cur}, ensure_ascii=False, indent=2), encoding="utf-8")
    print("\n✔ 全流水线通过，manifest 已刷新为当前基线")
    return 0


def main(argv=None):
    ap = argparse.ArgumentParser(description="试卷生成三层流水线编排器")
    ap.add_argument("action", choices=["init", "status", "run"])
    ap.add_argument("--config", required=True, help="pipeline.json 路径")
    ap.add_argument("--from", dest="from_layer", help="从指定层开始重跑（design/content/render）")
    ap.add_argument("--dry-run", action="store_true", help="只打印计划不执行")
    args = ap.parse_args(argv)
    cfg = load_config(args.config)
    py = sys.executable
    if args.action == "init":
        return cmd_init(cfg, py)
    if args.action == "status":
        return cmd_status(cfg, py)
    return cmd_run(cfg, py, args.from_layer, args.dry_run)


if __name__ == "__main__":
    sys.exit(main())
