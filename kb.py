#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
kb.py — 专利项目文件知识库 CLI（零第三方依赖，纯标准库）
基于 karpathy-llm-wiki 三层架构（sources / wiki / schema）。
命令：init / ingest / search / list / history / stats / refine / bulk

文本抽取：
  .txt/.md/.csv/.json/.py/.xml/.html → 直接读取
  .docx                            → zipfile + w:t 解析
  .xlsx                            → zipfile + sharedStrings/worksheet 单元格解析
  .pdf                             → zlib 解压 FlateDecode 流 + 内容流文本算子正则（尽力而为）
  .xls/.doc                        → 旧版二进制格式，标准库无法抽取（仅归档不抽全文）
YAML frontmatter：内置轻量解析/生成（覆盖本项目用到的标量 + 行内列表）。
"""
import argparse
import json
import os
import re
import shutil
import sys
import zipfile
import zlib
import threading
from collections import Counter
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parent
WIKI = ROOT / "wiki"
SOURCES = ROOT / "sources"
CONFIG = ROOT / "config.json"
SKIPPED_LARGE = WIKI / "_skipped_large.jsonl"
HISTORY = WIKI / "_search_history.jsonl"
INDEX = WIKI / "index.md"
LOG = WIKI / "log.md"

# type -> (子目录, 索引表头, 中文标签)
TYPES = {
    "project":   ("projects",   "Projects",   "项目"),
    "client":    ("clients",    "Clients",    "客户"),
    "applicant": ("applicants", "Applicants", "申请人/权利人"),
    "report":    ("reports",    "Reports",    "检索报告"),
    "comm":      ("comms",      "Comms",      "沟通记录"),
    "tech":      ("tech",       "Tech",       "技术域"),
}
SKIP = {"index.md", "log.md"}


# ---------------- 工具 ----------------
def now():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def today():
    return datetime.now().strftime("%Y-%m-%d")


def slugify(s):
    s = (s or "").strip().lower()
    s = re.sub(r"[^\w一-鿿]+", "-", s)
    s = re.sub(r"-+", "-", s).strip("-")
    return s or "item"


def load_config():
    try:
        return json.loads(CONFIG.read_text(encoding="utf-8"))
    except Exception:
        return {}


# ---------------- YAML frontmatter（轻量、自足） ----------------
_NEEDS_QUOTE = re.compile(r'[:#\[\],&*?|<>=!%@`"{}]')


def _unquote(v):
    v = v.strip()
    if len(v) >= 2 and v[0] == v[-1] and v[0] in "\"'":
        return v[1:-1]
    return v


def _quote(v):
    s = str(v)
    if s == "" or _NEEDS_QUOTE.search(s) or s.strip() != s:
        return '"' + s.replace("\\", "\\\\").replace('"', '\\"') + '"'
    return s


def _split_inline(s):
    # 按顶层逗号切分（不处理嵌套，本项目列表为简单字符串）
    out, buf, depth = [], "", 0
    for ch in s:
        if ch in "[{" :
            depth += 1
        elif ch in "]}":
            depth -= 1
        if ch == "," and depth == 0:
            out.append(buf.strip())
            buf = ""
        else:
            buf += ch
    if buf.strip():
        out.append(buf.strip())
    return out


def parse_frontmatter(text):
    m = re.match(r"^---\s*\n(.*?)\n---\s*\n?", text, re.S)
    if not m:
        return {}, text
    fm, body = m.group(1), text[m.end():]
    meta, i, lines = {}, 0, fm.splitlines()
    while i < len(lines):
        line = lines[i]
        if not line.strip() or line.lstrip().startswith("#"):
            i += 1
            continue
        if ":" not in line:
            i += 1
            continue
        key, _, val = line.partition(":")
        key, val = key.strip(), val.strip()
        if val == "":
            items, j = [], i + 1
            while j < len(lines) and re.match(r"^\s*-\s+", lines[j]):
                items.append(_unquote(lines[j].strip()[2:].strip()))
                j += 1
            meta[key] = items if items else ""
            i = j
            continue
        if val.startswith("[") and val.endswith("]"):
            inner = val[1:-1].strip()
            meta[key] = [] if inner == "" else [_unquote(x) for x in _split_inline(inner)]
        else:
            meta[key] = _unquote(val)
        i += 1
    return meta, body


def dump_frontmatter(meta, body):
    out = []
    for k, v in meta.items():
        if isinstance(v, (list, tuple)):
            out.append(f"{k}: [{', '.join(_quote(x) for x in v)}]")
        else:
            out.append(f"{k}: {_quote(v)}")
    return "---\n" + "\n".join(out) + "\n---\n\n" + body


# ---------------- 文本抽取（标准库） ----------------
def _decode_pdf_string(b):
    # 处理 \n \r \t \( \) \\ 以及 \ddd 八进制转义
    out, i = [], 0
    while i < len(b):
        c = b[i]
        if c == ord("\\") and i + 1 < len(b):
            nxt = b[i + 1]
            if nxt == ord("n"):
                out.append("\n"); i += 2; continue
            if nxt == ord("r"):
                out.append("\r"); i += 2; continue
            if nxt == ord("t"):
                out.append("\t"); i += 2; continue
            if nxt in (ord("("), ord(")"), ord("\\")):
                out.append(chr(nxt)); i += 2; continue
            if 48 <= nxt <= 55:  # 八进制 \ddd
                j = i + 1
                while j < i + 4 and j < len(b) and 48 <= b[j] <= 55:
                    j += 1
                try:
                    out.append(chr(int(b[i + 1:j].decode("ascii"), 8)))
                except Exception:
                    pass
                i = j; continue
            out.append(chr(nxt)); i += 2; continue
        out.append(chr(c) if 32 <= c < 127 else (chr(c) if c > 127 else " "))
        i += 1
    return "".join(out)


def _extract_pdf(path):
    data = path.read_bytes()
    texts = []
    try:
        for m in re.finditer(rb"stream\r?\n(.*?)\r?\nendstream", data, re.S):
            raw = m.group(1)
            try:
                dec = zlib.decompress(raw)
            except Exception:
                dec = raw
            for t in re.finditer(rb"\((?:[^()\\]|\\.)*\)\s*Tj", dec):
                s = t.group(0)
                inner = s[1:s.rfind(b")")]
                texts.append(_decode_pdf_string(inner))
            for t in re.finditer(rb"\[([^\]]*)\]\s*TJ", dec):
                for seg in re.finditer(rb"\((?:[^()\\]|\\.)*\)", t.group(1)):
                    texts.append(_decode_pdf_string(seg.group(0)[1:-1]))
    except Exception:
        pass
    return "\n".join(t for t in texts if t.strip())


def _extract_docx(path):
    try:
        with zipfile.ZipFile(path) as z:
            xml = z.read("word/document.xml")
    except Exception:
        return ""
    parts = re.findall(rb"<w:t[^>]*>(.*?)</w:t>", xml, re.S)
    return "\n".join(p.decode("utf-8", "ignore") for p in parts)


def _extract_xlsx(path, max_rows=200000, max_sheet_bytes=8_000_000,
                  small_threshold=5_000_000):
    """抽取 xlsx 全文。
    共享字符串(sharedStrings)承载绝大多数可搜索文本且体积小，优先抽取；
    仅对小文件(<small_threshold)额外做整表单元格解析（含内联字符串/数值）。
    大文件只抽 sharedStrings，避免对几十 MB 的 worksheet XML 做整表正则而卡死。
    """
    try:
        with zipfile.ZipFile(path) as z:
            names = z.namelist()
            texts = []
            if "xl/sharedStrings.xml" in names:
                data = z.read("xl/sharedStrings.xml")
                texts = [s.decode("utf-8", "ignore")
                         for s in re.findall(rb"<t[^>]*>(.*?)</t>", data, re.S)]
                if len(texts) > 100000:
                    texts = texts[:100000]
            if path.stat().st_size < small_threshold:
                shared = texts
                sheets = sorted(n for n in names
                                if re.match(r"xl/worksheets/sheet\d+\.xml", n))[:5]
                nrows = 0
                for sh in sheets:
                    xml = z.read(sh).decode("utf-8", "ignore")
                    if len(xml) > max_sheet_bytes:
                        xml = xml[:max_sheet_bytes]
                    for row in re.findall(r"<row[^>]*>(.*?)</row>", xml, re.S):
                        cells = []
                        for c in re.findall(r"<c\b[^>]*>(.*?)</c>", row, re.S):
                            if 't="s"' in c:
                                mv = re.search(r"<v>(.*?)</v>", c, re.S)
                                if mv:
                                    try:
                                        cells.append(shared[int(mv.group(1))])
                                    except Exception:
                                        pass
                            else:
                                mi = re.search(r"<is>.*?<t[^>]*>(.*?)</t>", c, re.S)
                                if mi:
                                    cells.append(mi.group(1))
                                else:
                                    mv = re.search(r"<v>(.*?)</v>", c, re.S)
                                    if mv:
                                        cells.append(mv.group(1))
                        texts.append("\t".join(cells))
                        nrows += 1
                        if nrows >= max_rows:
                            texts.append("...[截断: 超过 %d 行]" % max_rows)
                            break
            return "\n".join(texts)
    except Exception:
        return ""


def extract_text(path: Path):
    ext = path.suffix.lower()
    try:
        if ext in (".txt", ".md", ".csv", ".json", ".py", ".xml", ".html", ".htm"):
            return path.read_text(encoding="utf-8", errors="ignore")
        if ext == ".pdf":
            return _extract_pdf(path)
        if ext == ".docx":
            return _extract_docx(path)
        if ext == ".xlsx":
            return _extract_xlsx(path)
    except Exception as e:
        return f"[抽取失败: {e}]"
    return ""


def safe_extract(path: Path, timeout: float = 30.0):
    """带线程级超时兜底的全文抽取：单文件抽取超时就放弃，绝不拖垮整批。"""
    res = [None]

    def _run():
        try:
            res[0] = extract_text(path)
        except Exception as e:
            res[0] = f"[抽取失败: {e}]"

    th = threading.Thread(target=_run, daemon=True)
    th.start()
    th.join(timeout)
    if th.is_alive():
        return ""          # 超时：标记未抽取
    return res[0] or ""


# ---------------- 索引 / 日志 ----------------
def parse_index_summaries():
    out = {}
    if not INDEX.exists():
        return out
    for line in INDEX.read_text(encoding="utf-8").splitlines():
        m = re.match(r"\|\s*\[\[([^\]]+)\]\]\s*\|\s*(.*?)\s*\|", line)
        if m:
            out[m.group(1)] = m.group(2)
    return out


def rebuild_index():
    pages = []
    for p in WIKI.rglob("*.md"):
        rel = p.relative_to(WIKI)
        if p.name in SKIP or rel.name.startswith("_"):
            continue
        meta, _ = parse_frontmatter(p.read_text(encoding="utf-8"))
        if meta.get("type") in TYPES:
            pages.append((meta.get("type"), p, meta))
    existing = parse_index_summaries()
    lines = ["# Wiki Index", "", f"_Last updated: {today()} — {len(pages)} pages_", ""]
    for typ, (subdir, label, _) in TYPES.items():
        grp = [x for x in pages if x[0] == typ]
        lines.append(f"## {label}")
        lines.append("| Page | Summary | Updated |")
        lines.append("|------|---------|---------|")
        for _, p, meta in grp:
            slug = p.stem
            summ = existing.get(slug) or meta.get("summary") or meta.get("title") or ""
            lines.append(f"| [[{slug}]] | {summ} | {meta.get('date', today())} |")
        lines.append("")
    INDEX.write_text("\n".join(lines), encoding="utf-8")


def append_log(entry):
    LOG.parent.mkdir(parents=True, exist_ok=True)
    with LOG.open("a", encoding="utf-8") as f:
        f.write(f"## [{today()}] {entry}\n---\n\n")


def log_search(query, filters, hits):
    HISTORY.parent.mkdir(parents=True, exist_ok=True)
    rec = {"ts": now(), "query": query, "filters": filters, "hits": hits}
    with HISTORY.open("a", encoding="utf-8") as f:
        f.write(json.dumps(rec, ensure_ascii=False) + "\n")


# ---------------- 命令 ----------------
def _derive_year(name, mtime_year):
    cur = datetime.now().year
    for m in re.finditer(r"(19|20)\d{2}", name):
        y = int(m.group(0))
        if 1990 <= y <= cur:
            return str(y)
    return str(mtime_year)


def _derive_project_code(path_str):
    # 项目编号规则因所而异，优先用 config.json 的 code_pattern，这里是通用兜底
    m = re.search(r"\b([A-Z]{2,4}[- ]?\d{3,5})\b", path_str)
    return m.group(1) if m else ""


def _derive_jurisdiction(s):
    s_up = s.upper()
    for code, kws in (("US", ("美国", "USA", "U.S", "US-", "US ")),
                      ("DE", ("德国", "DE-", "DE ")),
                      ("EP", ("欧洲", "EP-", "EP ")),
                      ("JP", ("日本", "JP-", "JP ")),
                      ("CN", ("中国", "CN-", "CN "))):
        for kw in kws:
            if kw.upper() in s_up:
                return code
    return ""


def _derive_project_type(s):
    s_up = s.upper()
    if "FTO" in s_up:
        return "FTO"
    if "无效" in s:
        return "无效检索"
    if "新颖性" in s:
        return "新颖性检索"
    if "现有技术" in s:
        return "现有技术检索"
    if "防侵权" in s or ("侵权" in s and "防" in s):
        return "防侵权"
    if "标引" in s:
        return "专利标引"
    if "全景" in s or "分析" in s:
        return "专利分析"
    return ""


def _derive_tech_field(s):
    s_up = s.upper()
    for field, kws in (("机器人本体", ("机器人", "机械臂", "AGV", "ROBOT")),
                       ("动力电池", ("电池", "电芯", "模组", "BATTERY")),
                       ("半导体封装", ("封装", "晶圆", "WAFER", "PACKAGING")),
                       ("医疗器械", ("导管", "支架", "超声", "MEDICAL")),
                       ("外观设计", ("外观", "DESIGN"))):
        for kw in kws:
            if kw.upper() in s_up:
                return field
    return ""


def _bulk_page_path(typ, client_key, stem):
    subdir = TYPES[typ][0]
    slug = f"file-{slugify(client_key)}-{slugify(stem)}"[:90]
    return WIKI / subdir / f"{slug}.md"


def cmd_bulk(args):
    root = Path(args.root)
    include = {("." + x.strip().lower().lstrip(".")) for x in args.include.split(",") if x.strip()}
    folders = {f.strip() for f in (args.folders or "").split(",") if f.strip()}
    since = None
    if args.since:
        try:
            since = datetime.strptime(args.since, "%Y-%m-%d").timestamp()
        except Exception:
            try:
                since = datetime.strptime(args.since, "%Y").timestamp()
            except Exception:
                since = None
    typ = args.type or "report"
    if typ not in TYPES:
        print(f"未知类型 '{typ}'，可选：{', '.join(TYPES)}")
        sys.exit(1)
    subdir = TYPES[typ][0]

    files = []
    root_name = root.name
    custom_root = (str(root.resolve()) != str(ROOT.parent.resolve()))
    for dp, dn, fn in os.walk(root):
        parts = dp.split(os.sep)
        if "patent-kb" in parts:
            continue
        rel = Path(dp).relative_to(root)
        top = rel.parts[0] if rel.parts else "(root)"
        if folders and top not in folders:
            continue
        for f in fn:
            ext = os.path.splitext(f)[1].lower()
            if ext not in include:
                continue
            # client 归属：默认根目录取顶层文件夹；自定义根目录（如单客户）取根目录名
            client_key = root_name if custom_root else (top.lstrip("！!0123456789. ").strip() or top)
            files.append((Path(dp) / f, client_key))
    if args.limit:
        # 断点续跑：先剔除已生成知识页的文件，再取前 N 条，避免永远卡在已入库前缀
        remaining = []
        for f, ck in files:
            if not _bulk_page_path(typ, ck, f.stem).exists():
                remaining.append((f, ck))
        files = remaining[:args.limit]

    total = len(files)
    if args.dry:
        print(f"[dry-run] 命中 {total} 个文件（include={include}, folders={folders or '全部'}, since={args.since or '不限'}）")
        by_ext = Counter(os.path.splitext(f[0])[1].lower() for f in files)
        for k, v in sorted(by_ext.items()):
            print(f"  {k:7} {v}")
        return

    print(f"开始批量 ingest：{total} 个文件 | type={typ} | include={include} | folders={folders or '全部'}")
    ingested = extracted = skipped = 0
    for i, (f, client_key) in enumerate(files, 1):
        try:
            client = client_key
            fstat = f.stat()
            mtime_year = datetime.fromtimestamp(fstat.st_mtime).year
            year = _derive_year(f.name, mtime_year)
            pcode = _derive_project_code(str(f))
            jur = _derive_jurisdiction(str(f))
            ptype = _derive_project_type(str(f))
            tech = _derive_tech_field(str(f))
            dest_dir = SOURCES / year
            dest_dir.mkdir(parents=True, exist_ok=True)
            dest = dest_dir / f"{client_key}__{f.name}"
            if not dest.exists():
                shutil.copy2(f, dest)
            # 大文件：仅归档、跳过全文抽取（避免 30s 超时拖垮整批），记录待后续处理
            if fstat.st_size > args.max_mb * 1024 * 1024:
                with SKIPPED_LARGE.open("a", encoding="utf-8") as _sf:
                    _sf.write(json.dumps({"file": str(f), "size_mb": round(fstat.st_size / 1048576, 1),
                                          "ts": now()}, ensure_ascii=False) + "\n")
                if args.verbose:
                    print(f"  [跳过抽取] 过大: {f.name} ({round(fstat.st_size/1048576,1)}MB)")
                raw = ""
            else:
                raw = safe_extract(f, timeout=30)
            ok = bool(raw and not raw.startswith("[抽取失败"))
            if ok:
                sidecar = dest.parent / (dest.name + ".txt")
                sidecar.write_text(raw, encoding="utf-8")
                extracted += 1
            slug = f"file-{slugify(client_key)}-{slugify(f.stem)}"[:90]
            meta = {
                "type": typ,
                "title": f.name,
                "client": client,
                "ext": f.suffix.lower().lstrip("."),
                "date": year,
                "source_path": str(dest.relative_to(ROOT)),
            }
            if pcode:
                meta["project_code"] = pcode
            if jur:
                meta["jurisdiction"] = jur
            if ptype:
                meta["project_type"] = ptype
            if tech:
                meta["tech_field"] = tech
            tags = [x for x in (pcode, jur, ptype, tech) if x]
            if tags:
                meta["tags"] = tags
            body = f"# {f.name}\n\n> {client} / {f.name}（{f.suffix.lstrip('.')}）\n\n"
            body += f"- 源文件：`{dest.relative_to(ROOT)}`\n"
            body += f"- 归档时间：{now()}\n"
            body += f"- 客户/文件夹：{client}\n"
            if pcode:
                body += f"- 提取项目号：{pcode}\n"
            if jur:
                body += f"- 国家/地区：{jur}\n"
            if ptype:
                body += f"- 项目类型：{ptype}\n"
            if tech:
                body += f"- 技术领域：{tech}\n"
            body += f"- 全文抽取：{'已抽取 ' + str(len(raw)) + ' 字符' if ok else '未抽取（图片型PDF/旧版doc/xls，建议提供文本版）'}\n\n"
            body += "## 摘要\n\n（批量归档自动生成，请补充项目细节、检索结论、风险点等）\n\n"
            body += "## Related\n\n## Sources\n"
            page = WIKI / subdir / f"{slug}.md"
            if not page.exists():
                page.write_text(dump_frontmatter(meta, body), encoding="utf-8")
                ingested += 1
            else:
                skipped += 1
        except Exception as e:
            skipped += 1
            if args.verbose:
                print(f"  [跳过] {f}: {e}")
        if i % 250 == 0:
            print(f"  ... {i}/{total} 处理中（已抽全文 {extracted}）")

    rebuild_index()
    append_log(f"bulk | {ingested} 新增页 / {extracted} 抽全文 / {skipped} 跳过 | include={include} folders={folders or '全部'}")
    print(f"\n完成：新增知识页 {ingested} | 全文抽取 {extracted} | 跳过 {skipped} | 总计扫描 {total}")
    print(f"索引已重建：{INDEX.relative_to(ROOT)}")


def cmd_refine(args):
    """重新派生所有已入库页面的元数据维度：修正脏年份 + 补全 jurisdiction/project_type/tech_field/tags。
    仅在字段缺失或年份异常时更新，已正确的页面不动。依赖 bulk/ingest 写入的 source_path。"""
    pages = collect_pages()
    updated = 0
    cur = datetime.now().year
    for p, meta, body in pages:
        sp = meta.get("source_path")
        if not sp:
            # 旧版页面无 source_path：从 body 的「源文件：」行回退解析
            m = re.search(r"源文件：`([^`]+)`", body)
            sp = m.group(1) if m else None
        if not sp:
            continue
        name = Path(sp).name
        changed = False
        # 修正年份：始终用 _derive_year 重算（含范围校验 + 回退 mtime），
        # 仅当结果与原值不同才更新——可纠正早期误判的脏年份（如 1988/2098）
        try:
            mtime_year = datetime.fromtimestamp(Path(sp).stat().st_mtime).year
        except Exception:
            mtime_year = cur
        y = _derive_year(name, mtime_year)
        if str(y) != str(meta.get("date", "")):
            meta["date"] = str(y)
            changed = True
        # 补全可自动派生的维度（仅缺失时）
        if "jurisdiction" not in meta:
            jur = _derive_jurisdiction(name)
            if jur:
                meta["jurisdiction"] = jur
                changed = True
        if "project_type" not in meta:
            pt = _derive_project_type(name)
            if pt:
                meta["project_type"] = pt
                changed = True
        if "tech_field" not in meta:
            tf = _derive_tech_field(name)
            if tf:
                meta["tech_field"] = tf
                changed = True
        tags = [x for x in (meta.get("project_code"), meta.get("jurisdiction"),
                            meta.get("project_type"), meta.get("tech_field")) if x]
        if tags and meta.get("tags") != tags:
            meta["tags"] = tags
            changed = True
        if changed:
            p.write_text(dump_frontmatter(meta, body), encoding="utf-8")
            updated += 1
    rebuild_index()
    print(f"refine 完成：检查 {len(pages)} 页，更新 {updated} 页；索引已重建")


def cmd_init(args):
    for sub in TYPES.values():
        (WIKI / sub[0]).mkdir(parents=True, exist_ok=True)
    SOURCES.mkdir(parents=True, exist_ok=True)
    (SOURCES / ".gitkeep").touch()
    if not INDEX.exists():
        rebuild_index()
    print(f"知识库已初始化：{ROOT}")


def cmd_ingest(args):
    meta = {}
    for kv in args.set or []:
        if "=" not in kv:
            continue
        k, v = kv.split("=", 1)
        meta[k.strip()] = v.strip()
    typ = args.type or meta.get("type") or "report"
    if typ not in TYPES:
        print(f"未知类型 '{typ}'，可选：{', '.join(TYPES)}")
        sys.exit(1)
    subdir, label, _ = TYPES[typ]
    f = Path(args.file)
    if not f.exists():
        print(f"文件不存在：{f}")
        sys.exit(1)

    year = (meta.get("date") or today())[:4]
    dest_dir = SOURCES / year
    dest_dir.mkdir(parents=True, exist_ok=True)
    dest = dest_dir / f.name
    shutil.copy2(f, dest)

    raw = extract_text(f)
    extracted = bool(raw and not raw.startswith("[抽取失败"))
    if extracted:
        sidecar = dest.parent / (dest.name + ".txt")
        sidecar.write_text(raw, encoding="utf-8")

    if args.slug:
        slug = args.slug
    elif typ == "project" and meta.get("project_code"):
        slug = f"project-{slugify(meta['project_code'])}"
    else:
        slug = f"{typ}-{slugify(meta.get('title') or f.stem)}"

    if "type" not in meta:
        meta["type"] = typ
    if "title" not in meta:
        meta["title"] = f.name
    if "source_path" not in meta:
        meta["source_path"] = str(dest.relative_to(ROOT))

    summary = meta.get("summary") or meta.get("title")
    body = f"# {meta['title']}\n\n> {summary}\n\n"
    body += f"- 源文件：`{dest.relative_to(ROOT)}`\n"
    body += f"- 归档时间：{now()}\n"
    body += f"- 全文抽取：{'已抽取 ' + str(len(raw)) + ' 字符' if extracted else '未抽取（图片型PDF/不支持格式，建议提供文本版）'}\n\n"
    body += "## 摘要\n\n（由 ingest 自动生成，请补充项目细节、检索结论、风险点等）\n\n"
    body += "## Related\n\n## Sources\n"
    page = WIKI / subdir / f"{slug}.md"
    page.parent.mkdir(parents=True, exist_ok=True)
    page.write_text(dump_frontmatter(meta, body), encoding="utf-8")

    rebuild_index()
    append_log(f"ingest | {f.name} -> {page.relative_to(ROOT)}")
    print(f"已归档源文件：{dest.relative_to(ROOT)}")
    print(f"已生成知识页：{page.relative_to(ROOT)}")
    print(f"全文抽取：    {len(raw) if extracted else 0} 字符" if extracted else "全文抽取：    未抽取")
    print(f"已更新索引：  {INDEX.relative_to(ROOT)}")


def collect_pages():
    out = []
    for p in WIKI.rglob("*.md"):
        rel = p.relative_to(WIKI)
        if p.name in SKIP or rel.name.startswith("_"):
            continue
        text = p.read_text(encoding="utf-8")
        meta, body = parse_frontmatter(text)
        if meta.get("type") in TYPES:
            out.append((p, meta, body))
    return out


def match_filters(meta, f):
    if f["type"] and str(meta.get("type", "")).lower() != f["type"].lower():
        return False
    if f["project"] and f["project"].lower() not in str(meta.get("project_code", "")).lower():
        return False
    if f["client"] and f["client"].lower() not in str(meta.get("client", "")).lower():
        return False
    if f["applicant"] and f["applicant"].lower() not in str(meta.get("applicant", "")).lower():
        return False
    if f["jurisdiction"] and f["jurisdiction"].lower() not in str(meta.get("jurisdiction", "")).lower():
        return False
    if f["year"]:
        hay = " ".join(str(meta.get(k, "")) for k in ("date", "start_date"))
        if f["year"] not in hay:
            return False
    if f["tag"]:
        tags = meta.get("tags") or []
        if isinstance(tags, str):
            tags = [tags]
        if not any(f["tag"].lower() in str(t).lower() for t in tags):
            return False
    return True


def make_snippet(body, tokens, width=140):
    low = body.lower()
    pos = -1
    for tk in tokens:
        i = low.find(tk)
        if i != -1 and (pos == -1 or i < pos):
            pos = i
    if pos == -1:
        return body[:width].replace("\n", " ")
    start = max(0, pos - 40)
    return ("…" if start else "") + body[start:start + width].replace("\n", " ")


def cmd_search(args):
    query = args.query or ""
    tokens = re.findall(r"[\w一-鿿]+", query.lower())
    f = {
        "type": args.type, "project": args.project, "client": args.client,
        "applicant": args.applicant, "jurisdiction": args.jurisdiction,
        "year": args.year, "tag": args.tag,
    }
    results = []
    for p, meta, body in collect_pages():
        if not match_filters(meta, f):
            continue
        fm_str = " ".join(str(v) for v in meta.values()).lower()
        score = 0
        for tk in tokens:
            score += body.lower().count(tk) * 2
            score += fm_str.count(tk) * 3
        if tokens and score == 0:
            continue
        results.append((score, meta.get("title", p.stem), p, meta, make_snippet(body, tokens)))

    if args.sources:
        for sc in SOURCES.rglob("*.txt"):
            txt = sc.read_text(encoding="utf-8", errors="ignore").lower()
            score = sum(txt.count(tk) for tk in tokens)
            if tokens and score == 0:
                continue
            rel = sc.relative_to(ROOT)
            results.append((score, f"[源文件] {sc.stem}", None, {"source": str(rel)},
                             make_snippet(txt, tokens, 200)))

    results.sort(key=lambda x: -x[0])
    results = results[:args.top]

    print(f"检索：'{query}'  筛选：{ {k: v for k, v in f.items() if v} }  命中 {len(results)} 条\n")
    for i, (score, title, p, meta, snip) in enumerate(results, 1):
        loc = p.relative_to(ROOT) if p else meta.get("source", "")
        print(f"{i:>2}. [{score:>4}] {title}")
        print(f"     {loc}")
        if snip:
            print(f"     …{snip}…")
        print()
    log_search(query, {k: v for k, v in f.items() if v}, len(results))
    print(f"（本次检索已记入 {HISTORY.relative_to(ROOT)}）")


def cmd_list(args):
    f = {
        "type": args.type or None, "project": args.project, "client": args.client,
        "applicant": args.applicant, "jurisdiction": args.jurisdiction,
        "year": args.year, "tag": args.tag,
    }
    rows = [(p, meta) for p, meta, _ in collect_pages() if match_filters(meta, f)]
    if not rows:
        print("无匹配页面。")
        return
    print(f"{'类型':<9}{'标题':<28}{'项目':<10}{'客户':<14}{'类型2':<10}{'地区':<5}{'日期'}")
    print("-" * 92)
    for p, meta in rows:
        print(f"{meta.get('type',''):<9}{(meta.get('title','') or '')[:26]:<28}"
              f"{(str(meta.get('project_code','')) or '')[:10]:<10}"
              f"{(str(meta.get('client','')) or '')[:12]:<14}"
              f"{(str(meta.get('project_type','')) or '')[:10]:<10}"
              f"{(str(meta.get('jurisdiction','')) or '')[:5]:<5}{meta.get('date','')}")


def cmd_history(args):
    if not HISTORY.exists():
        print("尚无检索历史。")
        return
    lines = [l for l in HISTORY.read_text(encoding="utf-8").splitlines() if l.strip()]
    for line in lines[-args.n:]:
        try:
            r = json.loads(line)
        except Exception:
            continue
        flt = {k: v for k, v in (r.get("filters") or {}).items() if v}
        print(f"[{r['ts']}] 命中{r['hits']:>3} | {r['query']!r} | {flt}")


def cmd_stats(args):
    c = Counter()
    for _, meta, _ in collect_pages():
        c[meta.get("type", "unknown")] += 1
    print("页面统计：")
    for typ, (_, label, zh) in TYPES.items():
        print(f"  {label:<10}({zh}) : {c.get(typ,0)}")
    print(f"  合计        : {sum(c.values())}")
    src = sum(1 for _ in SOURCES.rglob("*") if _.is_file() and _.name != ".gitkeep")
    print(f"  源文件      : {src}")


# ---------------- CLI ----------------
def build_parser():
    ap = argparse.ArgumentParser(description="专利项目文件知识库 CLI")
    sub = ap.add_subparsers(dest="cmd")

    p = sub.add_parser("init", help="初始化目录结构")
    p.set_defaults(func=cmd_init)

    p = sub.add_parser("ingest", help="上传并归档一个项目文件")
    p.add_argument("--file", required=True, help="待归档文件路径")
    p.add_argument("--type", help="页面类型: " + "/".join(TYPES))
    p.add_argument("--slug", help="自定义页面 slug（默认自动生成）")
    p.add_argument("--set", action="append", metavar="key=value", help="元数据，可重复")
    p.set_defaults(func=cmd_ingest)

    p = sub.add_parser("search", help="全文 + 多维检索")
    p.add_argument("query", nargs="?", default="")
    p.add_argument("--type"); p.add_argument("--project"); p.add_argument("--client")
    p.add_argument("--applicant"); p.add_argument("--jurisdiction"); p.add_argument("--year")
    p.add_argument("--tag"); p.add_argument("--sources", action="store_true", help="同时检索源文件全文")
    p.add_argument("--top", type=int, default=20)
    p.set_defaults(func=cmd_search)

    p = sub.add_parser("list", help="按维度表格化列出")
    p.add_argument("--type"); p.add_argument("--project"); p.add_argument("--client")
    p.add_argument("--applicant"); p.add_argument("--jurisdiction"); p.add_argument("--year")
    p.add_argument("--tag")
    p.set_defaults(func=cmd_list)

    p = sub.add_parser("history", help="回溯检索历史")
    p.add_argument("--n", type=int, default=20)
    p.set_defaults(func=cmd_history)

    p = sub.add_parser("stats", help="统计")
    p.set_defaults(func=cmd_stats)

    p = sub.add_parser("refine", help="重派生全库元数据：修正脏年份+补全国家/类型/技术域/标签")
    p.set_defaults(func=cmd_refine)

    p = sub.add_parser("bulk", help="批量归档目录下所有匹配文件")
    p.add_argument("--root", default=str(ROOT.parent), help="待扫描根目录（默认工作区父目录）")
    p.add_argument("--include", default="xlsx,docx,pdf", help="逗号分隔扩展名，如 xlsx,docx,pdf,xls,doc")
    p.add_argument("--folders", help="仅限这些顶层文件夹（逗号分隔），留空=全部")
    p.add_argument("--since", help="仅收录该日期(YYYY-MM-DD)或年份(YYYY)之后修改的文件")
    p.add_argument("--type", default="report", help="页面类型，默认 report")
    p.add_argument("--limit", type=int, help="最多处理 N 个文件（试跑用）")
    p.add_argument("--dry", action="store_true", help="只统计不入库")
    p.add_argument("--verbose", action="store_true", help="打印跳过原因")
    p.add_argument("--max-mb", type=float, default=20.0, help="跳过大于该 MB 数的文件（仅归档不抽全文，记录到 _skipped_large.jsonl）")
    p.set_defaults(func=cmd_bulk)
    return ap


def main():
    ap = build_parser()
    args = ap.parse_args()
    if not args.cmd:
        ap.print_help()
        return
    args.func(args)


if __name__ == "__main__":
    main()
