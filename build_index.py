# -*- coding: utf-8 -*-
"""
build_index.py — 全量/增量建索引。
用法：
  python build_index.py --limit 150     # 冒烟测试
  python build_index.py                 # 全量（断点续跑：已入库且未变更的文件自动跳过）
"""
import os
import sys
import time
import json
import sqlite3
from concurrent.futures import ProcessPoolExecutor
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import kbcommon as K


def _worker_init(known_hashes=None):
    global _KNOWN
    _KNOWN = set(known_hashes or [])
    # 每个 worker 进程一个唯一 LibreOffice profile，避免并发锁死
    os.environ["KB_WORKER_ID"] = str(os.getpid())
    import jieba
    jieba.setLogLevel(60)
    jieba.initialize()


_KNOWN = set()


def process_file(args):
    """worker：抽取 + 元数据派生 + 分词，返回可 pickle 的 dict。"""
    path_str, rel_str, size, mtime, ext = args
    path = Path(path_str)
    rel = Path(rel_str)
    out = {"path": rel_str, "size": size, "mtime": mtime, "ext": ext,
           "text": "", "err": None, "meta": {}, "tokens": "", "chars": 0}
    # 统一用本地库重新抽取（旧 sources/*.txt 缓存质量差，含 XML 残渣与错位单元格）
    text, err = K.extract_any(path, ext)
    if err:
        out["err"] = err
    if not text.strip():
        return out
    text = K.cap_text(K.clean_text(text))
    meta = K.derive_meta(rel_str, path.stem, text, path.stat().st_mtime)
    out["text"] = text
    out["chars"] = len(text)
    out["meta"] = meta
    th = K.sha_text(text)
    # 检索索引只取前 20 万字分词（长篇报告后段信息密度低），全文仍完整缓存供摘录/LLM 使用
    out["tokens"] = "" if th in _KNOWN else K.tokenize(text[:200000])
    return out


def main():
    import argparse
    ap = argparse.ArgumentParser()
    ap.add_argument("--limit", type=int, default=0, help="只处理前 N 个文件（冒烟测试）")
    ap.add_argument("--workers", type=int, default=min(8, os.cpu_count() or 4))
    args = ap.parse_args()

    t0 = time.time()
    files = list(K.iter_files())
    print(f"[scan] 枚举到 {len(files)} 个待索引文件", flush=True)
    if args.limit:
        files = files[: args.limit]
        print(f"[scan] --limit 截取前 {len(files)} 个", flush=True)

    conn = K.get_db(building=True)
    done = {(r[0], r[1], r[2]): r[3] for r in
            conn.execute("SELECT path, size, mtime, text_hash FROM docs")}
    todo = [(str(K.ROOT / rel), str(rel), st.st_size, st.st_mtime, ext)
            for (p, rel, st, ext) in files
            if (str(rel), st.st_size, st.st_mtime) not in done]
    print(f"[scan] 已入库未变更 {len(files) - len(todo)} 个，待处理 {len(todo)} 个", flush=True)
    if not todo:
        print("[done] 无需更新", flush=True)
        return

    wiki_map = K.load_wiki_map()
    print(f"[wiki] wiki 页映射 {len(wiki_map)} 条", flush=True)

    known_hashes = [r[0] for r in conn.execute("SELECT hash FROM texts")]

    n_ok = n_empty = n_err = n_new_text = 0
    batch = []
    total = len(todo)
    last_profile_at = 0

    def flush_rows():
        nonlocal n_new_text
        for r in batch:
            text_hash = r["text_hash"]
            tokens = r["tokens"]
            tid = None
            if text_hash:
                row = conn.execute("SELECT tid FROM texts WHERE hash=?", (text_hash,)).fetchone()
                if not row:
                    if not tokens:            # 兜底：worker 判断失误时主进程补分词
                        tokens = K.tokenize(r["text"])
                    tp = str(K.TEXT_CACHE / f"{text_hash}.txt")
                    K.TEXT_CACHE.mkdir(parents=True, exist_ok=True)
                    (K.TEXT_CACHE / f"{text_hash}.txt").write_text(r["text"], encoding="utf-8")
                    cur = conn.execute(
                        "INSERT INTO texts(hash,chars,text_path,added_at) VALUES(?,?,?,?)",
                        (text_hash, r["chars"], tp, time.strftime("%Y-%m-%d %H:%M:%S")))
                    tid = cur.lastrowid
                    conn.execute("INSERT INTO texts_fts(rowid,hash,body) VALUES(?,?,?)",
                                 (tid, text_hash, tokens))
                    n_new_text += 1
                else:
                    tid = row[0]
            rel = r["path"]
            wiki_page = ""
            stem = Path(rel).name
            meta = r["meta"]
            client = meta.get("client", "")
            if client and "__" not in stem:
                wiki_page = wiki_map.get(f"{client}__{stem}", "") or wiki_map.get(stem, "")
            else:
                wiki_page = wiki_map.get(stem, "")
            m = meta
            conn.execute(
                "INSERT INTO docs(text_hash,path,title,ext,size,mtime,year,client,"
                "project_code,jurisdiction,project_type,tech_field,ipc,wiki_page,status,"
                "added_at,updated_at) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,1,?,?) "
                "ON CONFLICT(path) DO UPDATE SET text_hash=excluded.text_hash,"
                "title=excluded.title,size=excluded.size,mtime=excluded.mtime,"
                "year=excluded.year,client=excluded.client,"
                "project_code=excluded.project_code,jurisdiction=excluded.jurisdiction,"
                "project_type=excluded.project_type,tech_field=excluded.tech_field,"
                "ipc=excluded.ipc,wiki_page=excluded.wiki_page,updated_at=excluded.updated_at",
                (text_hash, rel, Path(rel).stem, r["ext"], r["size"], r["mtime"],
                 m.get("year", ""), m.get("client", ""), m.get("project_code", ""),
                 m.get("jurisdiction", ""), m.get("project_type", ""),
                 m.get("tech_field", ""), ";".join(m.get("ipcs", []))[:400],
                 wiki_page,
                 time.strftime("%Y-%m-%d %H:%M:%S"),
                 time.strftime("%Y-%m-%d %H:%M:%S")))

    def flush_with_retry(max_attempts=8):
        """批量写入；遇 database is locked 自动回滚重试，避免整库构建退出。"""
        for attempt in range(1, max_attempts + 1):
            try:
                flush_rows()
                conn.commit()
                return
            except sqlite3.OperationalError as e:
                msg = str(e).lower()
                if "locked" not in msg and "busy" not in msg:
                    raise
                print(f"[retry] 数据库被占用，5s 后重试 ({attempt}/{max_attempts})", flush=True)
                try:
                    conn.rollback()
                except Exception:
                    pass
                time.sleep(5)
        raise sqlite3.OperationalError("数据库持续被锁定，构建中止")

    slot = 0
    with ProcessPoolExecutor(max_workers=args.workers,
                             initializer=_worker_init,
                             initargs=(known_hashes,)) as ex:
        pending = {}
        it = iter(todo)
        # 提交队列保持约 workers*3 个在途任务，保证负载均衡
        in_flight = args.workers * 3
        futures = {}
        todo_iter = iter(todo)

        def submit_next(n):
            nonlocal slot
            for _ in range(n):
                try:
                    t = next(todo_iter)
                except StopIteration:
                    return
                slot += 1
                os.environ["KB_WORKER_SLOT"] = str(slot % args.workers)
                fut = ex.submit(process_file, t)
                futures[fut] = t

        submit_next(in_flight)
        done_n = 0
        while futures:
            # 轮询：取任意一个已完成的任务
            fut = None
            for f in list(futures):
                if f.done():
                    fut = f
                    break
            if fut is None:
                time.sleep(0.2)
                continue
            args_t = futures.pop(fut)
            try:
                r = fut.result()
            except Exception as e:
                r = {"path": str(args_t[1]), "size": args_t[2], "mtime": args_t[3],
                     "ext": args_t[4], "text": "", "err": f"worker:{e}"[:200],
                     "meta": {}, "tokens": "", "chars": 0}
            r["text_hash"] = K.sha_text(r["text"]) if r["text"] else ""
            batch.append(r)
            if r["err"]:
                n_err += 1
            elif not r["text"]:
                n_empty += 1
            else:
                n_ok += 1
            done_n += 1
            if len(batch) >= 100:
                flush_with_retry()
                batch.clear()
                el = time.time() - t0
                speed = done_n / el if el else 0
                eta = (total - done_n) / speed / 60 if speed else 0
                print(f"[{done_n}/{total}] ok={n_ok} empty={n_empty} err={n_err} "
                      f"新增文本={n_new_text} 速度={speed:.1f}/s ETA={eta:.0f}min", flush=True)
                if done_n - last_profile_at >= 2000:
                    last_profile_at = done_n
                    print("[profiles] 阶段性重建项目画像…", flush=True)
                    conn.commit()
                    K.rebuild_project_profiles(conn)
            submit_next(1)
        if batch:
            flush_with_retry()
            batch.clear()

    print("[wiki] 用卡片元数据回填项目号/类型/年份…", flush=True)
    n_enr = K.enrich_from_wiki(conn)
    print(f"[wiki] 回填 {n_enr} 条文档", flush=True)

    print(f"[profiles] 重建项目画像…", flush=True)
    K.rebuild_project_profiles(conn)

    # 历史检索式抽取
    print("[queries] 抽取历史检索式…", flush=True)
    nq = K.extract_queries(conn)
    print(f"[queries] 抽取 {nq} 条", flush=True)
    stats = K.count_stats(conn)
    el = time.time() - t0
    print(f"[done] 用时 {el/60:.1f}min | docs={stats['docs']} texts={stats['texts']} "
          f"projects={stats['projects']} queries={stats['queries']} "
          f"ok={n_ok} empty={n_empty} err={n_err}", flush=True)
    (K.APP / "data" / "build_result.json").write_text(
        json.dumps({"finished_at": time.strftime("%Y-%m-%d %H:%M:%S"),
                    "elapsed_min": round(el / 60, 1), **stats,
                    "ok": n_ok, "empty": n_empty, "err": n_err},
                   ensure_ascii=False, indent=2), encoding="utf-8")


import concurrent.futures as _cf


def concurrent_futures_wait(*a, **k):  # 未用占位
    return _cf.wait(*a, **k)


import re  # noqa: E402  (供 queries 抽取使用)

if __name__ == "__main__":
    main()
