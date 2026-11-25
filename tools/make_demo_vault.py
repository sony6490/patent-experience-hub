# -*- coding: utf-8 -*-
"""生成完全虚构的演示知识库（用于公开仓库示例与文章截图，不含任何真实客户/项目信息）。"""
import sys
from pathlib import Path

from docx import Document
from openpyxl import Workbook

OUT = sys.argv[1] if len(sys.argv) > 1 else str(
    Path(__file__).resolve().parent.parent / "demo-vault")
VAULT = Path(OUT)

CLIENTS = {
    "启元机器人": {
        "projects": [
            ("PB-2101", "协作机械臂碰撞防护", "FTO", "B25J9/16;B25J13/08",
             ["(机械臂 OR 机器人手臂 OR ROBOT ARM) AND (碰撞检测 OR 力控 OR FORCE CONTROL)",
              "(关节 OR JOINT) AND ( torque sensor OR 力矩传感器 ) AND ( 安全停止 OR 急停 )",
              "B25J9/16 OR B25J13/08 OR G05B23/02"]),
            ("PB-2118", "移动机器人视觉定位", "检索", "G01S17/08;G06T7/70",
             ["(激光雷达 OR LiDAR OR LIDAR) AND ( SLAM OR 同步定位与建图 )",
              "(反射标 OR reflector) AND ( 定位精度 OR localisation accuracy ) AND ( 环境光干扰 )",
              "G01S17/89 OR G01S17/42 OR G06T7/73"]),
            ("PB-2204", "抓取夹具自适应", "无效", "B25J15/08",
             ["(夹爪 OR 夹具 OR GRIPPER) AND ( 自适应 OR 柔性 OR compliant ) AND ( 气动 OR pneumatic )"]),
        ],
    },
    "恒晟新能源": {
        "projects": [
            ("PB-2312", "动力电池液冷板", "FTO", "H01M10/0525;H01L23/36",
             ["( 液冷板 OR 冷板 OR cooling plate ) AND ( 电芯 OR cell ) AND ( 流道 OR channel )",
              "( 相变材料 OR PCM ) OR ( 相变冷却 ) AND ( 热失控 OR thermal runaway )",
              "H01M10/0482 OR H01M10/0525 OR B60L58/26"]),
            ("PB-2331", "极耳焊接结构", "查新", "B23K26/24;H01R43/10",
             ["( 极耳 OR tab ) AND ( 激光焊接 OR laser welding ) AND ( 卷绕 OR winding )"]),
            ("PB-2402", "模组集成 CTB", "检索", "H01M50/244;B60K13/04",
             ["( 电池模组 OR module ) AND ( 一体化 OR integrated ) AND ( 托盘 OR tray OR 箱体 )"]),
        ],
    },
    "朗矽半导体": {
        "projects": [
            ("PB-2408", "TSV 垂直封装", "FTO", "H01L23/31;H01L21/60",
             ["( TSV OR 硅通孔 OR through silicon via ) AND ( 键合 OR bonding ) AND ( 铜 OR copper )",
              "( 扇出 OR fan-out OR FOPLP ) AND ( 再布线 OR RDL ) AND ( 应力 OR warpage )"]),
            ("PB-2419", "功率器件散热", "无效", "H01L23/498;H05K7/20",
             ["( 焊层 OR solder ) AND ( 银烧结 OR silver sintering ) AND ( 基板 OR substrate OR DBC )"]),
        ],
    },
    "普微医疗器械": {
        "projects": [
            ("PB-2507", "介入导管涂层", "FTO", "A61M25/00;A61L29/04",
             ["( 导管 OR catheter ) AND ( 亲水涂层 OR hydrophilic coating ) AND ( 鞘 OR sheath )",
              "( 支架 OR stent ) AND ( 药物洗脱 OR drug eluting ) AND ( 高分子 OR polymer )"]),
            ("PB-2521", "超声探头阵列", "检索", "A61B8/00;G01N29/04",
             ["( 超声 OR ultrasound ) AND ( 换能器 OR transducer ) AND ( 面阵 OR matrix array )"]),
        ],
    },
}

SUMMARIES = {
    "FTO": "防侵权（Freedom to Operate）检索分析报告：确认产品量产前是否落入他人有效权利要求。",
    "检索": "专项检索报告：梳理相关申请人专利布局与技术路线，输出技术分支表与重点专利清单。",
    "无效": "专利权无效宣告检索：检索可用于无效请求的现有技术证据，逐条比对权利要求。",
    "查新": "查新检索报告：核对技术方案的新颖性要点，给出检索要素表与对比文件。",
}


def w(path: Path, text: str):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding="utf-8")


def docx(path: Path, title: str, lines):
    path.parent.mkdir(parents=True, exist_ok=True)
    d = Document()
    d.add_paragraph(title)
    for ln in lines:
        d.add_paragraph(ln)
    d.save(path)


def xlsx(path: Path, rows):
    path.parent.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws.title = "权利要求比对"
    for r in rows:
        ws.append(r)
    wb.save(path)


import os
import re
import time


def stamp(path: Path, ymd):
    """把文件 mtime 改成报告出具日期，界面年份筛选才不至于全是建库当年。"""
    t = time.mktime(time.strptime(ymd, "%Y-%m-%d"))
    os.utime(path, (t, t))


def stamp_tree(root: Path, ymd):
    for f in root.rglob("*"):
        if f.is_file():
            stamp(f, ymd)


def build():
    if VAULT.exists():
        import shutil
        shutil.rmtree(VAULT)
    (VAULT / "项目资料").mkdir(parents=True, exist_ok=True)
    n = 0
    for client, info in CLIENTS.items():
        for code, subject, ptype, ipc, queries in info["projects"]:
            year = code.split("-")[1][:2]
            year = "20" + year
            base = VAULT / "项目资料" / client / f"{code} {subject}"
            # 1. 报告正文 docx
            body = [
                f"项目编号：{code}    委托方：{client}",
                f"项目类型：{ptype}检索    出具日期：{year}-06-18",
                f"技术主题：{subject}",
                "一、检索要素表",
            ]
            for i, q in enumerate(queries, 1):
                body.append(f"要素{i}：{q}")
            body += [
                f"分类号限定：{ipc.replace(';', '、')}",
                "二、检索策略",
                f"以“{subject}”为核心，在 CN、EP、US、WO 四局库中分别执行高级检索，"
                "同族合并后按被引频次排序，命中 486 族，去重后保留 72 篇进入人工阅读。",
                "三、结论",
                "未发现足以阻断量产的未决权利要求；另有 3 篇同族专利已进入无效程序，需持续关注。",
            ]
            docx(base / f"{subject}-{ptype}检索报告.docx", f"{subject}{ptype}报告（{code}）V2", body)
            # 2. 检索历史 xlsx
            xlsx(base / f"{code} 检索历史.xlsx",
                 [["序号", "数据库", "检索式", "命中"], [1, "CN 高级", queries[0], 214],
                  [2, "EP 高级", queries[-1] if queries[-1] else queries[0], 96],
                  [3, "WO", " OR ".join(queries[:1]), 58]])
            # 4. 项目说明 md
            w(base / "项目说明.md",
              f"# {code} {subject}\n\n"
              f"- 客户：{client}\n- 类型：{ptype}检索\n- 年份：{year}\n"
              f"- 技术域：{subject}\n- IPC：{ipc}\n\n"
              f"{SUMMARIES[ptype]}\n\n交付物：报告一份、重点专利清单 72 篇、权利要求比对表。\n")
            # 5. 沟通纪要 html
            w(base / f"与客户沟通纪要（{code}）.html",
              f"<html><head><meta charset='utf-8'><title>{code} 沟通纪要</title></head><body>"
              f"<h1>{client} — {subject}</h1>"
              f"<p>产品于 {year} 年 Q3 定型，量产前需确认{ptype}结论。</p>"
              f"<p>客户关注点：{queries[0][:40]} 相关结构是否已被覆盖。</p>"
              f"<p>IPC 关注范围：{ipc}</p></body></html>")
            # 6. 报价书（含示例分类号，验证 IPC 不被污染）
            docx(VAULT / "公司资料" / "标准报价与结算" / f"{subject}服务简易报价书（{code}）.docx",
                 f"{subject}服务报价书（{code}）",
                 [f"委托方:{client}", f"服务类型:{ptype}检索",
                  "示例分类号仅用于说明计费口径：G06F17/00、G06Q10/06、H04W12/00",
                  "报价：人民币 3.6 万元，含两轮补充检索。"])
            # 7. 检索式 txt（另一目录，验证多目录扫描）
            w(VAULT / "归档文件" / client / f"{code} 检索式汇总.txt",
              "\n".join(queries) + f"\n分类号：{ipc}\n")
            stamp_tree(base, f"{year}-06-18")
            n += 7
    # 干扰样本：员工花名册目录应被排除
    w(VAULT / "公司资料" / "个人档案" / "张三 身份证扫描件.txt", "本目录应被 exclude_dirs 排除")
    w(VAULT / "公司资料" / "财务资料" / "2024 年报销明细.txt", "本目录应被 exclude_dirs 排除")
    # 报价书与归档检索式不在项目目录里，统一按项目编号补时间戳，
    # 否则它们的 mtime 落在建库当天，界面年份会出现 2021 项目配 2026 文档
    for f in VAULT.rglob("*"):
        if not f.is_file():
            continue
        m = re.search(r"PB-(\d{2})\d{2}", f.name + str(f.parent))
        if m:
            stamp(f, f"20{m.group(1)}-06-18")
    print("演示库文件数(不含排除):", n, "  根目录:", VAULT)


if __name__ == "__main__":
    build()
