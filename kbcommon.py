# -*- coding: utf-8 -*-
"""
kbcommon.py — 检索窗口共享库：配置、文件枚举、文本抽取、分词、SQLite 读写。
复用上级 patent-kb/kb.py 的零依赖抽取器与元数据派生函数。
"""
import json
import math
import os
import re
import shutil
import sqlite3
import subprocess
import sys
import tempfile
import time
import hashlib
from pathlib import Path

APP = Path(__file__).resolve().parent
PKB = APP.parent                      # .../patent-kb
sys.path.insert(0, str(PKB))
import kb                              # 零依赖抽取与派生函数

APP_DIR_NAME = "app"
TOP_DIRS = {"研发实体", "Sunyu", "律所&代理公司", "公司介绍及服务协议",
            "政府及事业单位", "金融机构", "！格式化文件", "patent-kb"}

# ---------------- 配置 ----------------
def load_config():
    cfg = json.loads((APP / "config.json").read_text(encoding="utf-8"))
    local = APP / "config.local.json"
    if local.exists():
        try:
            cfg["llm_cfg"] = json.loads(local.read_text(encoding="utf-8"))
        except Exception:
            cfg["llm_cfg"] = {}
    else:
        cfg["llm_cfg"] = {}
    return cfg

CFG = load_config()
ROOT = Path(CFG["root"])
DB_PATH = APP / CFG["db"]
TEXT_CACHE = APP / CFG["text_cache"]
TOMBSTONES = APP / CFG["tombstones"]
TEXT_CAP = int(CFG.get("text_cap_chars", 800000))

# ---------------- 停用词 ----------------
STOPWORDS = set("""
的 了 是 在 和 与 及 或 等 中 为 对 由 从 本 其 该 一 个 上 下 这 那 有 无 按 依 关于 对于
通过 进行 可以 应当 不得 如果 则 但 而 且 并 或 以及 因此 由于 根据 依据 按照 上述 所述
如下 如上 以下 以上 有关 可能 需要 应该 相关
a an the and or of to in for on with by is are be was were this that from at as it its
""".split())

# ---------------- 文件枚举 ----------------
def iter_files(cfg=None):
    """枚举所有待索引文件（相对 VAULT 根的 Path）。"""
    cfg = cfg or CFG
    root = Path(cfg["root"])
    text_exts = set(cfg["text_exts"])
    exclude_dirs = set(cfg["exclude_dirs"])
    prefixes = tuple(cfg.get("exclude_file_prefixes", []))
    max_bytes = int(cfg.get("max_file_mb", 60)) * 1024 * 1024
    pdf_kw = re.compile(cfg.get("pdf_whitelist_keywords", ""), re.I)
    patent_no = re.compile(r"\b(CN|US|EP|WO|JP|KR|TW|AU|CA|IN|BR|MX|RU|DE|GB|FR)\s*\d", re.I)
    skip_sources_txt = cfg.get("skip_sources_txt", True)

    for d in cfg["scan_dirs"]:
        base = root / d
        if not base.exists():
            continue
        for p in base.rglob("*"):
            if not p.is_file():
                continue
            rel = p.relative_to(root)
            parts = rel.parts
            if any(part in exclude_dirs for part in parts):
                continue
            name = p.name
            if name.startswith(prefixes):
                continue
            ext = p.suffix.lower()
            if ext not in text_exts:
                continue
            # sources 目录下的 .txt 是抽取缓存，不单独索引
            if skip_sources_txt and ext == ".txt" and parts[0] == "patent-kb":
                continue
            try:
                st = p.stat()
            except OSError:
                continue
            if st.st_size > max_bytes or st.st_size == 0:
                continue
            # PDF 只收报告类；专利号命名的 PDF 一律跳过
            if ext == ".pdf":
                if patent_no.search(name):
                    continue
                if not pdf_kw.search(name):
                    continue
            yield p, rel, st, ext

# ---------------- 文本抽取 ----------------
def _strip_html(data):
    data = re.sub(r"(?is)<(script|style)[^>]*>.*?</\1>", " ", data)
    data = re.sub(r"(?s)<[^>]+>", " ", data)
    data = re.sub(r"&nbsp;?", " ", data)
    data = re.sub(r"&[a-z]{2,8};", " ", data)
    return data

def _extract_pptx(path):
    import zipfile
    try:
        out = []
        with zipfile.ZipFile(path) as z:
            slides = sorted(n for n in z.namelist()
                            if re.match(r"ppt/slides/slide\d+\.xml$", n))
            for sh in slides:
                xml = z.read(sh).decode("utf-8", "ignore")
                out.extend(m for m in re.findall(r"<a:t>(.*?)</a:t>", xml, re.S))
        return "\n".join(out)
    except Exception:
        return ""


def _extract_docx_lib(path):
    """python-docx：正文段落 + 表格单元格。"""
    try:
        import docx
        d = docx.Document(str(path))
        parts = [p.text for p in d.paragraphs if p.text and p.text.strip()]
        for t in d.tables:
            for row in t.rows:
                cells = [c.text.strip() for c in row.cells if c.text and c.text.strip()]
                if cells:
                    parts.append("\t".join(cells))
        return "\n".join(parts)
    except Exception:
        return ""


def _extract_xlsx_lib(path):
    """openpyxl：逐工作表取值。"""
    try:
        import openpyxl
        wb = openpyxl.load_workbook(str(path), read_only=True, data_only=True)
        parts, total = [], 0
        for ws in wb.worksheets:
            parts.append("## " + str(ws.title))
            for row in ws.iter_rows(values_only=True):
                cells = [str(v).strip() for v in row
                         if v is not None and str(v).strip()]
                if cells:
                    line = "\t".join(cells)
                    parts.append(line)
                    total += len(line)
                if total >= 800000:
                    break
            if total >= 800000:
                break
        wb.close()
        return "\n".join(parts)
    except Exception:
        return ""


def _extract_pdf_lib(path):
    """PyMuPDF：逐页抽文本，快且质量好；失败回退到 kb.safe_extract。"""
    try:
        import fitz
        doc = fitz.open(str(path))
        parts, total = [], 0
        for page in doc:
            t = page.get_text("text")
            if t:
                parts.append(t)
                total += len(t)
            if total >= 800000:
                break
        doc.close()
        return "\n".join(parts)
    except Exception:
        return ""

_SOFFICE = Path("C:/Program Files/LibreOffice/program/soffice.exe")
_CONV_SEQ = [0]

def _soffice_convert(path, filt):
    """LibreOffice 转文本。复制到 ASCII 临时目录规避中文路径问题；worker 专属 profile 支持并行。"""
    if not _SOFFICE.exists():
        return ""
    _CONV_SEQ[0] += 1
    wid = os.environ.get("KB_WORKER_ID") or str(os.getpid())
    tmp = Path(tempfile.gettempdir()) / f"kbconv_{wid}"
    tmp.mkdir(exist_ok=True)
    src = tmp / f"f{_CONV_SEQ[0]}{path.suffix.lower()}"
    prof = f"-env:UserInstallation=file:///{(tmp / ('prof' + wid)).as_posix()}"
    try:
        shutil.copy2(path, src)
        outdir = tmp / "out"
        outdir.mkdir(exist_ok=True)
        cmd = [str(_SOFFICE), prof, "--headless", "--norestore", "--convert-to",
               filt, "--outdir", str(outdir), str(src)]
        subprocess.run(cmd, timeout=120, capture_output=True)
        for cand in outdir.iterdir():
            if cand.suffix.lower() in (".txt", ".csv"):
                for enc in ("utf-8-sig", "utf-8", "gb18030", "latin-1"):
                    try:
                        data = cand.read_bytes()
                        return data.decode(enc, errors="ignore")
                    except Exception:
                        continue
        return ""
    except Exception:
        return ""
    finally:
        try:
            src.unlink(missing_ok=True)
        except Exception:
            pass

def _extract_docx2(path):
    """修正版 docx 抽取：精确匹配 <w:t> / <w:t ...>，不误吞 <w:tcPr> 等样式标签。"""
    import zipfile
    try:
        with zipfile.ZipFile(path) as z:
            xml = z.read("word/document.xml")
    except Exception:
        return ""
    parts = re.findall(rb"<w:t(?:\s[^>]*)?>(.*?)</w:t>", xml, re.S)
    return "\n".join(p.decode("utf-8", "ignore") for p in parts)


def extract_any(path, ext):
    """按扩展名分发抽取，返回 (text, err_msg)。"""
    text, err = "", None
    try:
        if ext in (".txt", ".md", ".csv"):
            text = path.read_text(encoding="utf-8", errors="ignore")
        elif ext in (".html", ".htm"):
            text = _strip_html(path.read_text(encoding="utf-8", errors="ignore"))
        elif ext == ".pptx":
            text = _extract_pptx(path)
        elif ext == ".doc":
            text = _soffice_convert(path, "txt")
        elif ext == ".xls":
            text = _soffice_convert(path, "csv")
        elif ext == ".docx":
            text = _extract_docx_lib(path) or _extract_docx2(path)
        elif ext == ".xlsx":
            text = _extract_xlsx_lib(path) or kb.safe_extract(path, timeout=90.0)
        elif ext == ".pdf":
            text = _extract_pdf_lib(path) or kb.safe_extract(path, timeout=90.0)
        else:
            err = "unsupported"
    except Exception as e:
        err = str(e)[:200]
    if isinstance(text, str) and text.startswith("[抽取失败"):
        err = text[:200]
        text = ""
    return text or "", err

def clean_text(s):
    s = re.sub(r"[ \t\r\f\v]+", " ", s)
    s = re.sub(r"\n{3,}", "\n\n", s)
    return s.strip()

def cap_text(s, cap=None):
    cap = cap or TEXT_CAP
    return s[:cap]

def sha_text(s):
    return hashlib.sha256(s.encode("utf-8", "ignore")).hexdigest()

# ---------------- 分词 ----------------
import jieba
_jieba_ready = False

def _init_jieba():
    global _jieba_ready
    if not _jieba_ready:
        jieba.setLogLevel(60)
        jieba.initialize()
        _jieba_ready = True

def tokenize(text):
    _init_jieba()
    return " ".join(t for t in jieba.cut_for_search(text) if t.strip())

def query_terms(q):
    _init_jieba()
    terms = []
    for t in jieba.cut_for_search(q):
        t = t.strip()
        if not t or t in STOPWORDS:
            continue
        if re.fullmatch(r"[\d.\-/,]+", t):
            continue
        if len(t) == 1 and not t.isascii():
            continue
        terms.append(t)
    # 去重保序
    seen, out = set(), []
    for t in terms:
        if t not in seen:
            seen.add(t)
            out.append(t)
    return out

# ---------------- 元数据派生 ----------------
IPC_RE = re.compile(r"\b([A-H]\d{2}[A-Z]\s?\d{1,4}/\d{1,4})\b")

def derive_client(rel, title):
    """从路径推导客户：sources 副本取 __ 前缀；业务目录取二级目录名。"""
    parts = Path(str(rel)).parts
    if parts and parts[0] == "patent-kb":
        stem = Path(title).stem if title else ""
        c = stem.split("__", 1)[0] if "__" in stem else ""
        return "" if c in TOP_DIRS else c
    if len(parts) >= 2:
        cand = parts[1]
        # 排除顶层目录名本身、以及看起来像文件名的片段
        if cand not in TOP_DIRS and not Path(cand).suffix:
            return cand
    return ""


def derive_meta(rel, title, text):
    rel_s = str(rel)
    client = derive_client(rel_s, title)
    # 技术域/项目类型/地域只用路径+文件名判定，避免正文词（如合同里的 scope）误判
    hay = rel_s + " " + (title or "")
    code = kb._derive_project_code(hay)
    if not code and text:
        code = kb._derive_project_code(text[:3000])
    year = kb._derive_year(title or rel_s, time.localtime().tm_year)
    juris = kb._derive_jurisdiction(hay)
    ptype = kb._derive_project_type(hay)
    tech = kb._derive_tech_field(hay)
    ipcs = []
    if text:
        seen = set()
        for mm in IPC_RE.finditer(text[:500000]):
            v = re.sub(r"\s+", "", mm.group(1))
            if v not in seen:
                seen.add(v)
                ipcs.append(v)
            if len(ipcs) >= 40:
                break
    return {"client": client, "project_code": code, "year": year,
            "jurisdiction": juris, "project_type": ptype,
            "tech_field": tech, "ipcs": ipcs}

# ---------------- wiki 页映射 ----------------
def load_wiki_meta():
    """解析 7,305 张 wiki 卡片的 frontmatter，返回两层索引：
      by_pair{(client, filename): meta}   —— 精确匹配
      by_name{filename: meta}             —— 仅在文件名全局唯一时可用
    meta = {code, client, type, year, wiki}
    """
    by_pair, by_name, dup = {}, {}, set()
    wiki = PKB / "wiki"
    if not wiki.exists():
        return by_pair, by_name
    for p in wiki.rglob("*.md"):
        try:
            meta, _ = kb.parse_frontmatter(p.read_text(encoding="utf-8"))
        except Exception:
            continue
        sp = str(meta.get("source_path") or "")
        title = str(meta.get("title") or "")
        name = Path(sp).name if sp else title
        if not name:
            continue
        base, client_key = name, ""
        if "__" in base:
            client_key, base = base.split("__", 1)
        rel = p.relative_to(ROOT).as_posix()
        item = {"code": str(meta.get("project_code") or ""),
                "client": client_key or str(meta.get("client") or ""),
                "type": str(meta.get("project_type") or ""),
                "year": str(meta.get("date") or ""),
                "wiki": rel}
        if client_key:
            by_pair[(client_key, base)] = item
        if base in by_name:
            dup.add(base)
        else:
            by_name[base] = item
    for b in dup:
        by_name.pop(b, None)
    return by_pair, by_name


def enrich_from_wiki(conn):
    """用 wiki 卡片元数据回填 docs 中缺失的项目号/类型/年份/wiki 页。返回更新条数。"""
    by_pair, by_name = load_wiki_meta()
    if not by_pair and not by_name:
        return 0
    rows = conn.execute(
        "SELECT id, path, client, project_code, project_type, year, wiki_page "
        "FROM docs").fetchall()
    n = 0
    for did, path, client, code, ptype, year, wiki in rows:
        stem = Path(path).name
        item = by_pair.get((client, stem))
        if not item and "__" in stem:
            c2, base = stem.split("__", 1)
            item = by_pair.get((c2, base)) or by_pair.get((client, base))
        if not item:
            item = by_name.get(stem)
            if not item and "__" in stem:
                item = by_name.get(stem.split("__", 1)[1])
        if not item:
            continue
        nc = code or item["code"]
        nt = ptype or item["type"]
        ny = year or item["year"]
        nw = wiki or item["wiki"]
        if (nc, nt, ny, nw) != (code, ptype, year, wiki):
            conn.execute(
                "UPDATE docs SET project_code=?, project_type=?, year=?, wiki_page=? "
                "WHERE id=?", (nc, nt, ny, nw, did))
            n += 1
    conn.commit()
    return n


def load_wiki_map():
    """返回 {文件名或 client__文件名: wiki页相对 VAULT 根的 posix 路径}。"""
    m = {}
    wiki = PKB / "wiki"
    if not wiki.exists():
        return m
    for p in wiki.rglob("*.md"):
        try:
            meta, _ = kb.parse_frontmatter(p.read_text(encoding="utf-8"))
        except Exception:
            continue
        sp = str(meta.get("source_path") or "")
        if not sp:
            continue
        rel = p.relative_to(ROOT).as_posix()
        base = Path(sp).name                       # client__name 或 name
        m.setdefault(base, rel)
        if "__" in base:
            m.setdefault(base.split("__", 1)[1], rel)
    return m

# ---------------- SQLite ----------------
SCHEMA = """
CREATE TABLE IF NOT EXISTS meta(key TEXT PRIMARY KEY, value TEXT);
CREATE TABLE IF NOT EXISTS texts(
  tid INTEGER PRIMARY KEY AUTOINCREMENT,
  hash TEXT UNIQUE, chars INTEGER, text_path TEXT, added_at TEXT);
CREATE VIRTUAL TABLE IF NOT EXISTS texts_fts USING fts5(
  hash UNINDEXED, body, content='', contentless_delete=1);
CREATE TABLE IF NOT EXISTS docs(
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  text_hash TEXT, path TEXT UNIQUE, title TEXT, ext TEXT,
  size INTEGER, mtime REAL,
  year TEXT, client TEXT, project_code TEXT, jurisdiction TEXT,
  project_type TEXT, tech_field TEXT, ipc TEXT,
  wiki_page TEXT, status INTEGER DEFAULT 1,
  added_at TEXT, updated_at TEXT);
CREATE INDEX IF NOT EXISTS idx_docs_hash ON docs(text_hash);
CREATE INDEX IF NOT EXISTS idx_docs_project ON docs(project_code);
CREATE INDEX IF NOT EXISTS idx_docs_client ON docs(client);
CREATE TABLE IF NOT EXISTS projects(
  tid INTEGER PRIMARY KEY AUTOINCREMENT, code TEXT UNIQUE, name TEXT, clients TEXT,
  tech_fields TEXT, types TEXT, years TEXT, doc_count INTEGER, top_keywords TEXT,
  top_ipc TEXT, profile TEXT, updated_at TEXT);
CREATE VIRTUAL TABLE IF NOT EXISTS projects_fts USING fts5(
  code UNINDEXED, body, content='', contentless_delete=1);
CREATE TABLE IF NOT EXISTS queries(
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  project_code TEXT, text TEXT, doc_id INTEGER, src TEXT);
CREATE TABLE IF NOT EXISTS sync_log(
  id INTEGER PRIMARY KEY AUTOINCREMENT, ts TEXT, action TEXT, path TEXT, note TEXT);
"""

_SCHEMA_DONE = False

def get_db(building=False):
    global _SCHEMA_DONE
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(DB_PATH, timeout=120)
    if not _SCHEMA_DONE:
        conn.executescript(SCHEMA)
        conn.execute("PRAGMA journal_mode=WAL")
        _SCHEMA_DONE = True
    else:
        # 库文件被重建时兜底
        try:
            conn.execute("SELECT 1 FROM meta LIMIT 1")
        except sqlite3.OperationalError:
            conn.executescript(SCHEMA)
            conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA synchronous=%s" % ("OFF" if building else "NORMAL"))
    conn.execute("PRAGMA temp_store=MEMORY")
    conn.execute("PRAGMA cache_size=-64000")
    return conn

# ---------------- 项目画像 ----------------
_KW_STOP = STOPWORDS | {
    "深圳", "上海", "北京", "广州", "杭州", "苏州", "宁波", "温州",
    # 公司名/文档套话（不作为检索关键词推荐）
    "兰登", "紫金", "兰登紫金", "公司", "介绍", "项目", "报价", "报价单",
    "服务", "合同", "协议", "模板", "案例", "汇总", "附件", "版本", "更新",
    "文档", "文件", "内容", "客户", "专利", "知识产权", "报告", "分析",
    "检索", "研究", "方案", "技术", "有限公司", "我们", "贵司", "如下",
    "包括", "基于", "相关", "进行", "以及", "可以", "需要", "提供",
    "完成", "工作", "时间", "人员", "团队", "注意", "事项", "须知",
}

def _filter_token(t):
    if not t or t in _KW_STOP or t in STOPWORDS:
        return False
    if re.fullmatch(r"[\d.\-/,_%]+", t):
        return False
    if len(t) == 1 and not t.isascii():
        return False
    if t.isascii() and len(t) < 3:
        return False
    return True

def rebuild_project_profiles(conn):
    """全量重建 projects + projects_fts（画像为纯派生数据，先清空再生成）。"""
    conn.execute("DELETE FROM projects_fts")
    conn.execute("DELETE FROM projects")
    rows = conn.execute(
        "SELECT id, path, title, client, project_code, year, jurisdiction, "
        "project_type, tech_field, ipc FROM docs WHERE status=1").fetchall()
    groups = {}
    for (did, path, title, client, code, year, juris, ptype, tech, ipc) in rows:
        key = code if code else ("~" + (client or "未归类"))
        g = groups.setdefault(key, {"name": code or (client or "未归类"),
                                    "docs": [], "clients": set(), "tech": set(),
                                    "types": set(), "years": set(), "ipc": {}})
        g["docs"].append({"id": did, "title": title or "", "path": path})
        if client: g["clients"].add(client)
        if tech: g["tech"].add(tech)
        if ptype: g["types"].add(ptype)
        if year: g["years"].add(year)
        # IPC 只从报告类文档聚合，避免报价书/合同里的示例分类号污染
        report_like = bool(re.search(
            r"报告|检索|分析|FTO|无效|查新|意见|landscape|search|report|analysis|opinion",
            title or "", re.I))
        if ipc and report_like:
            for v in ipc.split(";"):
                if v: g["ipc"][v] = g["ipc"].get(v, 0) + 1

    # 语料级 df（按组）
    tfs = {}
    for key, g in groups.items():
        toks = [t for t in tokenize(
                    " ".join(d["title"] for d in g["docs"][:300])).split()
                if _filter_token(t)]
        tf = {}
        for t in toks:
            tf[t] = tf.get(t, 0) + 1
        tfs[key] = tf
    df = {}
    for tf in tfs.values():
        for t in tf:
            df[t] = df.get(t, 0) + 1
    n = max(len(groups), 1)

    for key, g in groups.items():
        tf = tfs[key]
        scored = [(t, f * math.log(1 + n / df[t])) for t, f in tf.items()]
        scored.sort(key=lambda x: -x[1])
        kws = [t for t, _ in scored[:30]]
        top_ipc = sorted(g["ipc"].items(), key=lambda x: -x[1])[:12]
        docs_sorted = sorted(g["docs"], key=lambda d: -len(d["title"]))[:8]
        profile = " ".join([
            g["name"], " ".join(sorted(g["tech"])), " ".join(sorted(g["types"])),
            " ".join(kws), " ".join(d["title"] for d in docs_sorted),
        ])
        clients_s = ";".join(sorted(g["clients"]))[:300]
        now_s = time.strftime("%Y-%m-%d %H:%M:%S")
        vals = (key, g["name"], clients_s, ";".join(sorted(g["tech"]))[:200],
                ";".join(sorted(g["types"]))[:200],
                "%s-%s" % (min(g["years"]) if g["years"] else "",
                           max(g["years"]) if g["years"] else ""),
                len(g["docs"]), json.dumps(kws, ensure_ascii=False),
                json.dumps(top_ipc, ensure_ascii=False), profile, now_s)
        cur = conn.execute(
            "INSERT INTO projects(code,name,clients,tech_fields,types,years,"
            "doc_count,top_keywords,top_ipc,profile,updated_at) "
            "VALUES(?,?,?,?,?,?,?,?,?,?,?)", vals)
        tid = cur.lastrowid
        conn.execute("INSERT INTO projects_fts(rowid,code,body) VALUES(?,?,?)",
                     (tid, key, tokenize(profile)))
    conn.commit()

# ---------------- 历史检索式抽取 ----------------
# 逐行线性扫描（禁用嵌套量词正则，避免灾难性回溯导致卡死）
_BOOL_RE = re.compile(r"\b(?:AND|OR|NOT)\b")  # 大小写敏感：专利检索式用大写布尔算符
_Q_TITLE_HINT = ("检索式", "检索策略", "关键词", "query", "strategy", "search")


def extract_queries(conn):
    """从标题含检索式/关键词/strategy 的文档中抽取检索式，写入 queries 表。"""
    conn.execute("DELETE FROM queries")
    like = " OR ".join(["title LIKE ?"] * len(_Q_TITLE_HINT))
    rows = conn.execute(
        f"SELECT id, project_code, title, text_hash FROM docs "
        f"WHERE status=1 AND ({like}) AND text_hash<>''",
        ["%" + h + "%" for h in _Q_TITLE_HINT]).fetchall()
    nq = 0
    for (did, code, title, th) in rows:
        tp = TEXT_CACHE / f"{th}.txt"
        if not tp.exists():
            continue
        try:
            data = tp.read_text(encoding="utf-8", errors="ignore")[:200000]
        except Exception:
            continue
        seen = set()
        for line in data.splitlines():
            if len(line) < 13 or len(line) > 400:
                continue
            if "(" not in line or ")" not in line:
                continue
            if not _BOOL_RE.search(line):
                continue
            s = line.strip()
            if "(" in s and ")" in s:
                s = s[:s.rfind(")") + 1].strip()
            s = re.sub(r"\s+", " ", s)
            if len(s) > 12 and s not in seen and len(seen) < 30:
                seen.add(s)
                conn.execute(
                    "INSERT INTO queries(project_code,text,doc_id,src) "
                    "VALUES(?,?,?,'title-match')", (code, s[:500], did))
                nq += 1
    conn.commit()
    return nq


def count_stats(conn):
    d = conn.execute("SELECT COUNT(*), SUM(size) FROM docs WHERE status=1").fetchone()
    t = conn.execute("SELECT COUNT(*) FROM texts").fetchone()
    p = conn.execute("SELECT COUNT(*) FROM projects").fetchone()
    q = conn.execute("SELECT COUNT(*) FROM queries").fetchone()
    return {"docs": d[0] or 0, "size_mb": round((d[1] or 0) / 1048576, 1),
            "texts": t[0], "projects": p[0], "queries": q[0]}
