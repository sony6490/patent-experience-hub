# -*- coding: utf-8 -*-
"""
server.py — 专利项目经验检索窗口（本地服务，仅监听 127.0.0.1）
启动：python server.py  →  http://127.0.0.1:8765
"""
import json
import os
import re
import sys
import time
import html
import threading
import urllib.parse
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import kbcommon as K

STATIC = K.APP / "static"
ASK_LOG = K.APP / "data" / "ask_log.jsonl"
CFG = K.CFG
VAULT_NAME = CFG.get("vault_name", "MyVault")
HOST = CFG.get("server_host", "127.0.0.1")
PORT = int(CFG.get("server_port", 8765))

LLM = (CFG.get("llm_cfg") or {}).get("llm", {})
LLM_ENABLED = bool(LLM.get("enabled") and LLM.get("api_key"))
ASK_CFG = (CFG.get("llm_cfg") or {}).get("ask", {})

MIME = {".html": "text/html; charset=utf-8", ".js": "text/javascript; charset=utf-8",
        ".css": "text/css; charset=utf-8", ".png": "image/png", ".jpg": "image/jpeg",
        ".svg": "image/svg+xml", ".ico": "image/x-icon"}

FILE_MIME = {".pdf": "application/pdf", ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
             ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
             ".pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
             ".doc": "application/msword", ".xls": "application/vnd.ms-excel",
             ".txt": "text/plain; charset=utf-8", ".md": "text/plain; charset=utf-8",
             ".csv": "text/plain; charset=utf-8", ".html": "text/html; charset=utf-8"}


# ================= 检索层 =================
def fts_query(terms, joiner=" "):
    return joiner.join('"%s"' % t.replace('"', '""') for t in terms)


def read_text_cached(h, head=600000):
    if not h:
        return ""
    p = K.TEXT_CACHE / f"{h}.txt"
    if not p.exists():
        return ""
    try:
        return p.read_text(encoding="utf-8", errors="ignore")[:head]
    except Exception:
        return ""


def make_snippet(text, terms, width=280):
    if not text:
        return ""
    low = text.lower()
    pos = -1
    for t in terms:
        i = low.find(t.lower())
        if i != -1 and (pos == -1 or i < pos):
            pos = i
    if pos == -1:
        return text[:width]
    start = max(0, pos - 60)
    seg = text[start:start + width]
    return ("…" if start else "") + seg


def _apply_filters(conn, sql, params, f):
    if f.get("client"):
        sql += " AND d.client = ?"
        params.append(f["client"])
    if f.get("project"):
        sql += " AND d.project_code LIKE ?"
        params.append("%" + f["project"] + "%")
    if f.get("tech"):
        sql += " AND d.tech_field = ?"
        params.append(f["tech"])
    if f.get("type"):
        sql += " AND d.project_type = ?"
        params.append(f["type"])
    if f.get("year"):
        sql += " AND d.year = ?"
        params.append(f["year"])
    return sql, params


def search_docs(conn, q, filters=None, limit=30, or_fallback=True):
    filters = filters or {}
    terms = K.query_terms(q)
    if not terms:
        return [], terms
    out = []
    for joiner in (" ", " OR ") if or_fallback else (" ",):
        sql = ("SELECT d.id, d.path, d.title, d.ext, d.year, d.client, d.project_code, "
               "d.jurisdiction, d.project_type, d.tech_field, d.ipc, d.wiki_page, d.text_hash, "
               "t.chars, bm25(texts_fts) AS rank "
               "FROM texts_fts f JOIN texts t ON f.rowid=t.tid "
               "JOIN docs d ON d.text_hash=t.hash AND d.status=1 WHERE texts_fts MATCH ?")
        params = [fts_query(terms, joiner)]
        sql, params = _apply_filters(conn, sql, params, filters)
        sql += " ORDER BY rank LIMIT ?"
        params.append(limit * 2)
        try:
            rows = conn.execute(sql, params).fetchall()
        except Exception:
            continue
        if rows:
            seen = set()
            for r in rows:
                (did, path, title, ext, year, client, code, juris, ptype,
                 tech, ipc, wiki, th, chars, rank) = r
                # 原文件 + sources 副本视为同一份，避免证据重复
                key = (re.sub(r"^.*?__", "", title or ""), client or "")
                if key in seen:
                    continue
                seen.add(key)
                text = read_text_cached(th)
                out.append({
                    "id": did, "path": path, "title": title, "ext": ext,
                    "year": year, "client": client, "project": code,
                    "jurisdiction": juris, "type": ptype, "tech": tech,
                    "ipc": [x for x in (ipc or "").split(";") if x][:8],
                    "wiki": wiki, "rank": round(rank, 3), "chars": chars or 0,
                    "snippet": make_snippet(text, terms),
                })
                if len(out) >= limit:
                    break
            break
    return out, terms


def browse_docs(conn, filters=None, limit=60):
    """仅按左侧筛选条件浏览文档（无需输入检索词）。"""
    filters = filters or {}
    sql = ("SELECT d.id, d.path, d.title, d.ext, d.year, d.client, d.project_code, "
           "d.jurisdiction, d.project_type, d.tech_field, d.ipc, d.wiki_page, d.text_hash "
           "FROM docs d WHERE d.status=1")
    params = []
    sql, params = _apply_filters(conn, sql, params, filters)
    sql += " ORDER BY d.year DESC, d.id DESC LIMIT ?"
    params.append(limit)
    rows = conn.execute(sql, params).fetchall()
    out = []
    for r in rows:
        (did, path, title, ext, year, client, code, juris, ptype,
         tech, ipc, wiki, th) = r
        text = read_text_cached(th, head=4000)
        out.append({
            "id": did, "path": path, "title": title, "ext": ext,
            "year": year, "client": client, "project": code,
            "jurisdiction": juris, "type": ptype, "tech": tech,
            "ipc": [x for x in (ipc or "").split(";") if x][:8],
            "wiki": wiki, "chars": len(text),
            "snippet": (text[:260] + ("…" if len(text) > 260 else "")).strip(),
        })
    return out


def similar_projects(conn, q, filters=None, topn=15):
    filters = filters or {}
    terms = K.query_terms(q)
    if not terms:
        return []
    rows = []
    for joiner in (" ", " OR "):
        sql = ("SELECT p.tid, p.code, p.name, p.clients, p.tech_fields, p.types, p.years, "
               "p.doc_count, p.top_keywords, p.top_ipc, p.profile, bm25(projects_fts) AS rank "
               "FROM projects_fts f JOIN projects p ON f.rowid=p.tid WHERE projects_fts MATCH ?")
        params = [fts_query(terms, joiner)]
        if filters.get("tech"):
            sql += " AND p.tech_fields LIKE ?"
            params.append("%" + filters["tech"] + "%")
        sql += " ORDER BY rank LIMIT ?"
        params.append(topn * 2)
        try:
            rows = conn.execute(sql, params).fetchall()
        except Exception:
            rows = []
        if rows:
            break
    qset = set(terms)
    out = []
    for (tid, code, name, clients, techs, types, years, dcount, kws_json,
         ipc_json, profile, rank) in rows:
        kws = json.loads(kws_json or "[]")
        overlap = len(qset & set(kws))
        score = -rank + 0.6 * min(overlap, 4)
        out.append({"code": code, "name": name, "clients": clients or "",
                    "tech": techs or "", "types": types or "", "years": years or "",
                    "doc_count": dcount, "keywords": kws[:15],
                    "ipc": json.loads(ipc_json or "[]"), "score": round(score, 2),
                    "overlap": overlap})
    out.sort(key=lambda x: -x["score"])
    res = out[:topn]
    # 每个项目附 2~3 个代表文档
    for p in res:
        docs = conn.execute(
            "SELECT id, path, title, ext, year, wiki_page, text_hash FROM docs "
            "WHERE status=1 AND project_code=? ORDER BY length(title) DESC LIMIT 3",
            (p["code"],)).fetchall()
        if not docs:
            client0 = (p["clients"] or "").split(";")[0]
            docs = conn.execute(
                "SELECT id, path, title, ext, year, wiki_page, text_hash FROM docs "
                "WHERE status=1 AND client=? ORDER BY length(title) DESC LIMIT 3",
                (client0,)).fetchall()
        p["docs"] = [{"id": d[0], "path": d[1], "title": d[2], "ext": d[3],
                      "year": d[4], "wiki": d[5]} for d in docs]
    return res


QUERY_FILE_RE = re.compile(r"检索式|关键词|检索策略|query", re.I)


def recommend_keywords(conn, hits, projects, q, topn=18):
    terms = K.query_terms(q)
    freq = {}
    srcs = {}
    for p in projects[:8]:
        for kw in p.get("keywords", []):
            freq[kw] = freq.get(kw, 0) + 1
            srcs.setdefault(kw, "项目画像")
    # 历史命中文档标题（高信噪比）
    for h in hits[:25]:
        toks = [t for t in K.tokenize(h.get("title", "")).split() if K._filter_token(t)]
        for t in toks:
            freq[t] = freq.get(t, 0) + 1
            srcs.setdefault(t, "历史标题")
    for t in terms:
        freq[t] = freq.get(t, 0) + 5
        srcs[t] = "你的描述"
    ranked = sorted(freq.items(), key=lambda x: -x[1])
    # 过滤客户名与项目号（不是检索关键词）
    client_tokens = set()
    for p in projects[:8]:
        for cl in (p.get("clients") or "").split(";"):
            if cl:
                client_tokens.update(K.tokenize(cl).split())
    ranked = [(k, n) for k, n in ranked
              if k not in client_tokens and not re.match(r"(?i)^lz\d+$", k)][:topn]
    out = [{"kw": k, "n": n, "en": k.isascii(), "src": srcs.get(k, "")} for k, n in ranked]
    return out


def _hash_of(conn, doc_id):
    r = conn.execute("SELECT text_hash FROM docs WHERE id=?", (doc_id,)).fetchone()
    return r[0] if r else ""


def recommend_ipc(projects, topn=10):
    """只从与查询关键词有重叠的项目聚合 IPC，且按项目排名 + 重叠度加权。"""
    agg = {}
    for i, p in enumerate(projects[:8]):
        ov = p.get("overlap", 0)
        if ov <= 0:
            continue
        w = max(1, 8 - i) * (1 + min(ov, 4))
        for code, n in p.get("ipc", []):
            agg[code] = agg.get(code, 0) + w * (n if isinstance(n, int) else 1)
    ranked = sorted(agg.items(), key=lambda x: -x[1])[:topn]
    return [{"code": c, "n": n, "src": "历史实证"} for c, n in ranked]


def build_queries(kws, ipcs, q):
    cn = [k["kw"] for k in kws if not k["en"]][:4]
    en = [k["kw"] for k in kws if k["en"]][:3]
    ipc = [i["code"] for i in ipcs][:3]
    core = " OR ".join(cn[:3]) if cn else "核心关键词"
    feature = cn[1] if len(cn) > 1 else "特征词"
    en_part = (" OR ".join(en[:2])) if en else ""
    ipc_part = " OR ".join(ipc[:2]) if ipc else ""
    out = []
    s1 = f"({core})"
    if en_part:
        s1 += f" OR ({en_part})"
    if ipc_part:
        s1 += f" AND ({ipc_part})"
    out.append({"name": "查全式", "expr": s1,
                "note": "关键词并集 + 分类号限定，先跑一遍看噪音量"})
    s2 = f"{cn[0] if cn else '核心词'} AND {feature}"
    if ipc_part:
        s2 += f" AND ({ipc_part})"
    out.append({"name": "查准式", "expr": s2,
                "note": "核心词与特征词同句限定，用于精读清单"})
    s3 = f"{core}" + (f" NOT ({cn[-1] if len(cn) > 2 else '无关词'})")
    out.append({"name": "排查式(无效检索)", "expr": s3,
                "note": "无效/现有技术检索时按特征逐组排查"})
    return out


def obsidian_uri(rel_posix):
    return ("obsidian://open?vault=" + urllib.parse.quote(VAULT_NAME, safe="")
            + "&file=" + urllib.parse.quote(rel_posix, safe="/"))


# ================= LLM =================
SYSTEM_PROMPT = """你是一名资深专利检索分析师，服务于一家专利咨询公司的检索团队。
用户会给出【新项目描述】以及从公司历史项目知识库检索到的【证据】（相似项目、文档片段、候选关键词、历史分类号、历史检索式）。
你的任务：输出一份新项目检索方案，严格遵循：
1. 只能基于【证据】作答，禁止编造不存在的项目、文件、分类号或数据。
2. 每个论断末尾标注引用编号，如 [1][2]，编号只能来自【证据文档】列表；禁止使用 [Top10 P2] 这类非证据编号。项目相关信息写在表格里，用项目号指代。
3. 若证据不足，直说"历史库中未找到足够相近的项目"，并基于候选关键词给出通用建议。
4. 分类号必须区分【历史实证】（证据中出现过）与【建议】（你的专业推断），建议要标注"建议"。
5. 检索式使用智慧芽语法（AND/OR/NOT、通配符*、同句算子按需说明）。
6. 相似项目表格中的项目号/客户/年份必须与【相似历史项目】列表一致，不得改写或编造；年份异常时如实标注"待核"。
7. 输出 Markdown，结构固定为：
## 相似历史项目
（表格：项目/客户/年份/技术域/相似理由，附引用）
## 推荐检索关键词
（分组：核心词/扩展词/英文同义词，每行末尾附引用）
## 推荐分类号（IPC/CPC）
（分【历史实证】与【建议】两栏）
## 检索式建议
（至少 3 条：查全式、查准式、无效/特征排查式，附简短说明）
## 风险与提醒
（3 条以内，基于证据）"""


def llm_ask(user_content):
    from openai import OpenAI
    client = OpenAI(api_key=LLM["api_key"], base_url=LLM.get("base_url", "https://api.deepseek.com"))
    resp = client.chat.completions.create(
        model=LLM.get("model", "deepseek-chat"),
        messages=[{"role": "system", "content": SYSTEM_PROMPT},
                  {"role": "user", "content": user_content}],
        temperature=float(LLM.get("temperature", 0.2)),
        max_tokens=int(LLM.get("max_tokens", 3500)),
    )
    usage = resp.usage
    return resp.choices[0].message.content, {
        "prompt_tokens": usage.prompt_tokens,
        "completion_tokens": usage.completion_tokens,
        "total_tokens": usage.total_tokens,
    }


def build_user_content(q, filters, projects, hits, kws, ipcs, queries, mode):
    max_docs = ASK_CFG.get("deep_max_docs", 24) if mode == "deep" else ASK_CFG.get("fast_max_docs", 10)
    lines = ["【新项目描述】", q or "（用户未填写，请按筛选条件推断）"]
    if any(filters.values()):
        lines.append("【筛选条件】" + json.dumps(filters, ensure_ascii=False))
    lines.append("\n【相似历史项目 Top%d】" % len(projects[:10]))
    for i, p in enumerate(projects[:10], 1):
        lines.append(f"P{i} {p['code']} | 客户:{p['clients'][:40]} | {p['years']} | "
                     f"技术域:{p['tech'][:40]} | 类型:{p['types'][:30]} | "
                     f"{p['doc_count']}份文档 | 关键词:{','.join(p['keywords'][:8])}")
    lines.append(f"\n【证据文档（引用编号 [n]）】")
    for i, h in enumerate(hits[:max_docs], 1):
        lines.append(f"[{i}] {h['title']} | {h['client']} | {h['year']} | "
                     f"{h['project'] or '无项目号'} | {h['tech']}")
        if h.get("snippet"):
            lines.append("    摘录：" + h["snippet"][:260].replace("\n", " "))
    lines.append("\n【候选关键词（按频次）】" + ", ".join(k["kw"] for k in kws[:18]))
    lines.append("【历史实证 IPC】" + ", ".join(i["code"] for i in ipcs[:8]) if ipcs else "【历史实证 IPC】（无）")
    if queries:
        lines.append("【历史检索式样例】")
        for qq in queries[:6]:
            lines.append("  " + qq[:180])
    return "\n".join(lines)


def local_answer(q, projects, hits, kws, ipcs, queries):
    """无 LLM 时的结构化回答。"""
    L = []
    L.append("## 相似历史项目\n")
    if projects:
        L.append("| 项目 | 客户 | 年份 | 技术域 | 文档数 |")
        L.append("|---|---|---|---|---|")
        for p in projects[:8]:
            L.append(f"| {p['code']} | {p['clients'][:30]} | {p['years']} | {p['tech'][:30]} | {p['doc_count']} |")
    else:
        L.append("未找到明显相似项目。")
    L.append("\n## 推荐检索关键词\n")
    if kws:
        L.append("、".join(k["kw"] for k in kws[:15]))
    L.append("\n## 推荐分类号（IPC）\n")
    if ipcs:
        L.append("历史实证：" + "、".join(i["code"] for i in ipcs[:6]))
    else:
        L.append("历史库中未提取到直接分类号，建议人工判断。")
    L.append("\n## 检索式建议（模板，未接 LLM）\n")
    for s in build_queries(kws, ipcs, q):
        L.append(f"- **{s['name']}**：`{s['expr']}`")
    L.append("\n> 当前为本地模式（未配置 DeepSeek Key），配置后可获得完整分析。")
    return "\n".join(L)


def run_ask(conn, payload):
    q = (payload.get("q") or "").strip()
    filters = payload.get("filters") or {}
    mode = payload.get("mode") or "fast"
    t0 = time.time()
    hits, terms = search_docs(conn, q, filters, limit=40)
    projects = similar_projects(conn, q, filters, topn=15)
    kws = recommend_keywords(conn, hits, projects, q)
    ipcs = recommend_ipc(projects)
    queries = []
    codes = [p["code"] for p in projects[:10]
             if p.get("code") and not p["code"].startswith("~")]
    if codes:
        marks = ",".join("?" * len(codes))
        queries = [r[0] for r in conn.execute(
            f"SELECT text FROM queries WHERE project_code IN ({marks}) "
            f"ORDER BY id DESC LIMIT 12", codes)]
    if not queries:
        queries = [r[0] for r in conn.execute(
            "SELECT text FROM queries ORDER BY id DESC LIMIT 8")]
    usage = None
    if LLM_ENABLED and q:
        try:
            content = build_user_content(q, filters, projects, hits, kws, ipcs, queries, mode)
            answer, usage = llm_ask(content)
        except Exception as e:
            answer = local_answer(q, projects, hits, kws, ipcs, queries)
            answer += f"\n\n> ⚠️ LLM 调用失败，已回退本地模式：{str(e)[:150]}"
    else:
        answer = local_answer(q, projects, hits, kws, ipcs, queries)
    ev_limit = ASK_CFG.get("deep_max_docs", 24) if mode == "deep" else ASK_CFG.get("fast_max_docs", 10)
    evidence = []
    for i, h in enumerate(hits[:ev_limit], 1):
        evidence.append({**h, "cite": i,
                         "obsidian": obsidian_uri(h["wiki"]) if h.get("wiki") else "",
                         "file_url": f"/api/file?id={h['id']}",
                         "reveal_url": f"/api/reveal?id={h['id']}"})
    for p in projects:
        for d in p.get("docs", []):
            d["obsidian"] = obsidian_uri(d["wiki"]) if d.get("wiki") else ""
            d["file_url"] = f"/api/file?id={d['id']}"
    result = {"answer": answer, "evidence": evidence, "projects": projects[:10],
              "keywords": kws, "ipcs": ipcs, "queries": queries,
              "elapsed": round(time.time() - t0, 2), "usage": usage,
              "llm": LLM_ENABLED}
    try:
        ASK_LOG.parent.mkdir(parents=True, exist_ok=True)
        with ASK_LOG.open("a", encoding="utf-8") as f:
            f.write(json.dumps({"ts": time.strftime("%Y-%m-%d %H:%M:%S"),
                                "q": q, "mode": mode, "usage": usage,
                                "elapsed": result["elapsed"]},
                               ensure_ascii=False) + "\n")
    except Exception:
        pass
    return result


# ================= HTTP =================
class Handler(BaseHTTPRequestHandler):
    protocol_version = "HTTP/1.1"
    server_version = "PatentKB/1.0"
    timeout = 60          # 单连接最长等待，防止僵死连接拖垮线程池

    def log_message(self, fmt, *args):
        pass

    def _json(self, obj, code=200):
        data = json.dumps(obj, ensure_ascii=False).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(data)

    def _read_body(self):
        n = int(self.headers.get("Content-Length") or 0)
        if n <= 0:
            return {}
        raw = self.rfile.read(n)
        try:
            return json.loads(raw.decode("utf-8"))
        except Exception:
            return {}

    def do_HEAD(self):
        self.do_GET()

    def do_GET(self):
        u = urllib.parse.urlparse(self.path)
        path = u.path
        qs = urllib.parse.parse_qs(u.query)
        if path in ("/", "/index.html"):
            return self._static("index.html")
        if path.startswith("/static/"):
            return self._static(path[len("/static/"):])
        if path == "/api/stats":
            with _db() as conn:
                s = K.count_stats(conn)
            s["llm"] = LLM_ENABLED
            s["llm_model"] = LLM.get("model", "")
            s["sync_log"] = _last_sync()
            return self._json(s)
        if path == "/api/filters":
            with _db() as conn:
                clients = conn.execute(
                    "SELECT client, COUNT(*) n FROM docs WHERE status=1 AND client<>'' "
                    "GROUP BY client ORDER BY n DESC LIMIT 120").fetchall()
                techs = conn.execute(
                    "SELECT DISTINCT tech_field FROM docs WHERE status=1 AND tech_field<>''"
                ).fetchall()
                types = conn.execute(
                    "SELECT DISTINCT project_type FROM docs WHERE status=1 AND project_type<>''"
                ).fetchall()
                years = conn.execute(
                    "SELECT DISTINCT year FROM docs WHERE status=1 AND year<>'' "
                    "ORDER BY year DESC").fetchall()
            return self._json({
                "clients": [{"name": c, "n": n} for c, n in clients],
                "techs": [t[0] for t in techs],
                "types": [t[0] for t in types],
                "years": [y[0] for y in years],
            })
        if path == "/api/file":
            return self._serve_file(qs.get("id", [""])[0])
        if path == "/api/reveal":
            return self._reveal(qs.get("id", [""])[0])
        if path == "/api/kb_folders":
            return self._get_kb_folders()
        return self._json({"error": "not found"}, 404)

    def do_POST(self):
        if self.path == "/api/search":
            body = self._read_body()
            with _db() as conn:
                hits, terms = search_docs(conn, body.get("q", ""),
                                          body.get("filters"), limit=30)
                projects = similar_projects(conn, body.get("q", ""), body.get("filters"), 10)
                for p in projects:
                    for d in p.get("docs", []):
                        d["obsidian"] = obsidian_uri(d["wiki"]) if d.get("wiki") else ""
                        d["file_url"] = f"/api/file?id={d['id']}"
            for h in hits:
                h["obsidian"] = obsidian_uri(h["wiki"]) if h.get("wiki") else ""
                h["file_url"] = f"/api/file?id={h['id']}"
                h["reveal_url"] = f"/api/reveal?id={h['id']}"
            return self._json({"hits": hits, "projects": projects, "terms": terms})
        if self.path == "/api/browse":
            body = self._read_body()
            with _db() as conn:
                hits = browse_docs(conn, body, limit=60)
            for h in hits:
                h["obsidian"] = obsidian_uri(h["wiki"]) if h.get("wiki") else ""
                h["file_url"] = f"/api/file?id={h['id']}"
                h["reveal_url"] = f"/api/reveal?id={h['id']}"
            return self._json({"hits": hits})
        if self.path == "/api/ask":
            body = self._read_body()
            try:
                with _db() as conn:
                    res = run_ask(conn, body)
            except Exception as e:
                return self._json({"error": str(e)[:300]}, 500)
            return self._json(res)
        if self.path == "/api/kb_folders":
            return self._set_kb_folders()
        if self.path == "/api/rescan":
            try:
                import sync
                r = sync.scan_once()
                return self._json(r)
            except Exception as e:
                return self._json({"error": str(e)[:200]}, 500)
        return self._json({"error": "not found"}, 404)

    def _get_kb_folders(self):
        root = Path(K.CFG["root"])
        scan = list(K.CFG.get("scan_dirs", []))
        excl = set(K.CFG.get("exclude_dirs", []))
        names = []
        try:
            names = [p.name for p in root.iterdir()
                     if p.is_dir() and not p.name.startswith(".") and p.name not in excl]
        except Exception:
            pass
        for s in scan:
            top = s.split("/", 1)[0]
            if top not in names and top not in excl:
                names.append(top)
        folders = []
        with _db() as conn:
            for name in sorted(set(names)):
                n = conn.execute(
                    "SELECT COUNT(*) FROM docs WHERE status=1 AND path LIKE ?",
                    (name + "\\%",)).fetchone()[0]
                folders.append({"path": name, "n": n})
        return self._json({"root": str(root), "scan_dirs": scan,
                           "exclude_dirs": sorted(excl), "folders": folders})

    def _set_kb_folders(self):
        body = self._read_body()
        new_scan = body.get("scan_dirs")
        force = bool(body.get("force"))
        if not isinstance(new_scan, list) or not all(isinstance(x, str) for x in new_scan):
            return self._json({"error": "scan_dirs 必须是字符串列表"}, 400)
        clean = []
        for s in new_scan:
            s = s.strip().replace("\\", "/").strip("/")
            if not s or s.startswith(".") or s.startswith("/") or ":" in s \
                    or ".." in s.split("/"):
                return self._json({"error": "非法文件夹路径: " + s}, 400)
            if s not in clean:
                clean.append(s)
        cfg_path = K.APP / "config.json"
        try:
            cfg = json.loads(cfg_path.read_text(encoding="utf-8"))
            cfg["scan_dirs"] = clean
            tmp = cfg_path.with_name("config.json.tmp")
            tmp.write_text(json.dumps(cfg, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
            os.replace(tmp, cfg_path)
            K.reload_config()
            import sync
            sync.reload_cfg()
            state = sync.scan_once(force=force)
        except Exception as e:
            return self._json({"error": str(e)[:300]}, 500)
        return self._json({"ok": True, "scan_dirs": clean, "sync": state})

    def _static(self, name):
        f = (STATIC / name).resolve()
        try:
            f.relative_to(STATIC.resolve())
        except ValueError:
            return self._json({"error": "not found"}, 404)
        if not f.exists() or not f.is_file():
            return self._json({"error": "not found"}, 404)
        data = f.read_bytes()
        self.send_response(200)
        self.send_header("Content-Type", MIME.get(f.suffix.lower(), "application/octet-stream"))
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def _serve_file(self, doc_id):
        try:
            doc_id = int(doc_id)
        except ValueError:
            return self._json({"error": "bad id"}, 400)
        with _db() as conn:
            r = conn.execute("SELECT path, ext, title, status FROM docs WHERE id=?",
                             (doc_id,)).fetchone()
        if not r:
            return self._json({"error": "doc not found"}, 404)
        path, ext, title, status = r
        if not status:
            return self._json({"error": "doc deleted (tombstone)"}, 410)
        fp = K.ROOT / path
        if not fp.exists():
            return self._json({"error": "file missing on disk"}, 404)
        data = fp.read_bytes()
        mime = FILE_MIME.get(ext.lower(), "application/octet-stream")
        fname = urllib.parse.quote(fp.name, safe="")
        self.send_response(200)
        self.send_header("Content-Type", mime)
        self.send_header("Content-Length", str(len(data)))
        if ext.lower() in (".pdf", ".txt", ".md", ".csv", ".html"):
            self.send_header("Content-Disposition", f'inline; filename*=UTF-8\'\'{fname}')
        else:
            self.send_header("Content-Disposition", f'attachment; filename*=UTF-8\'\'{fname}')
        self.end_headers()
        self.wfile.write(data)

    def _reveal(self, doc_id):
        try:
            doc_id = int(doc_id)
        except ValueError:
            return self._json({"error": "bad id"}, 400)
        with _db() as conn:
            r = conn.execute("SELECT path, status FROM docs WHERE id=?", (doc_id,)).fetchone()
        if not r or not r[1]:
            return self._json({"error": "not available"}, 404)
        fp = K.ROOT / r[0]
        if not fp.exists():
            return self._json({"error": "file missing"}, 404)
        import subprocess
        subprocess.Popen(["explorer", "/select,", str(fp)])
        return self._json({"ok": True})


class _db:
    def __enter__(self):
        self.conn = K.get_db()
        return self.conn

    def __exit__(self, *a):
        self.conn.close()


def _last_sync():
    try:
        p = K.APP / "data" / "sync_state.json"
        if p.exists():
            return json.loads(p.read_text(encoding="utf-8"))
    except Exception:
        pass
    return {}


def _write_pid(name="server.pid"):
    """写下当前进程 PID，供「停止服务.bat」精确结束（不误杀其他 python）。"""
    try:
        p = K.APP / "data" / name
        p.parent.mkdir(parents=True, exist_ok=True)
        p.write_text(str(os.getpid()), encoding="ascii")
    except Exception:
        pass


def main():
    ThreadingHTTPServer.daemon_threads = True
    addr = (HOST, PORT)
    httpd = ThreadingHTTPServer(addr, Handler)
    _write_pid("server.pid")
    print(f"[server] 专利项目经验检索台 by oc 检索窗口 http://{HOST}:{PORT}  (LLM: {'ON/' + LLM.get('model', '') if LLM_ENABLED else 'OFF(本地模式)'})",
          flush=True)
    httpd.serve_forever()


if __name__ == "__main__":
    main()
