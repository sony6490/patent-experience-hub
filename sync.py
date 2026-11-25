# -*- coding: utf-8 -*-
"""
sync.py — 增量同步：新增/修改/删除/复活。
用法：
  python sync.py once   # 单次增量（server 的 /api/rescan 也调用这里）
  python sync.py loop   # 常驻循环（start_sync.bat 用）
"""
import json
import os
import sys
import time
from datetime import datetime, timedelta
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import kbcommon as K

CFG = K.CFG
INTERVAL = int(CFG.get("sync_interval_sec", 300))
TOMBSTONE_DAYS = int(CFG.get("tombstone_days", 30))
STATE = K.APP / "data" / "sync_state.json"
BULK_DELETE_GUARD = 50


def _log_sync(conn, action, path, note=""):
    conn.execute("INSERT INTO sync_log(ts,action,path,note) VALUES(?,?,?,?)",
                 (time.strftime("%Y-%m-%d %H:%M:%S"), action, path, note))


def purge_old_tombstones(conn):
    """物理清理超过保留期的软删除记录。"""
    cutoff = (datetime.now() - timedelta(days=TOMBSTONE_DAYS)).strftime("%Y-%m-%d %H:%M:%S")
    rows = conn.execute(
        "SELECT id, path FROM docs WHERE status=0 AND updated_at < ?",
        (cutoff,)).fetchall()
    for did, path in rows:
        conn.execute("DELETE FROM docs WHERE id=?", (did,))
        _log_sync(conn, "purge", path, "tombstone 超过 %d 天" % TOMBSTONE_DAYS)
    return len(rows)


def _insert_text_and_doc(conn, r):
    text_hash = K.sha_text(r["text"]) if r["text"] else ""
    tokens = r["tokens"]
    if text_hash and tokens:
        row = conn.execute("SELECT tid FROM texts WHERE hash=?", (text_hash,)).fetchone()
        if not row:
            K.TEXT_CACHE.mkdir(parents=True, exist_ok=True)
            (K.TEXT_CACHE / f"{text_hash}.txt").write_text(r["text"], encoding="utf-8")
            cur = conn.execute("INSERT INTO texts(hash,chars,text_path,added_at) VALUES(?,?,?,?)",
                               (text_hash, r["chars"], str(K.TEXT_CACHE / f"{text_hash}.txt"),
                                time.strftime("%Y-%m-%d %H:%M:%S")))
            conn.execute("INSERT INTO texts_fts(rowid,hash,body) VALUES(?,?,?)",
                         (cur.lastrowid, text_hash, tokens))
    rel = r["path"]
    m = r["meta"]
    now = time.strftime("%Y-%m-%d %H:%M:%S")
    conn.execute(
        "INSERT INTO docs(text_hash,path,title,ext,size,mtime,year,client,project_code,"
        "jurisdiction,project_type,tech_field,ipc,wiki_page,status,added_at,updated_at) "
        "VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,COALESCE(NULLIF(?,''),"
        "(SELECT wiki_page FROM docs WHERE path=?),''),1,?,?) "
        "ON CONFLICT(path) DO UPDATE SET text_hash=excluded.text_hash,title=excluded.title,"
        "size=excluded.size,mtime=excluded.mtime,year=excluded.year,client=excluded.client,"
        "project_code=excluded.project_code,jurisdiction=excluded.jurisdiction,"
        "project_type=excluded.project_type,tech_field=excluded.tech_field,"
        "ipc=excluded.ipc,wiki_page=excluded.wiki_page,status=1,updated_at=excluded.updated_at",
        (text_hash, rel, Path(rel).stem, r["ext"], r["size"], r["mtime"],
         m.get("year", ""), m.get("client", ""), m.get("project_code", ""),
         m.get("jurisdiction", ""), m.get("project_type", ""), m.get("tech_field", ""),
         ";".join(m.get("ipcs", []))[:400], r.get("wiki_page", ""), rel, now, now))


def scan_once(force=False):
    import build_index as B
    B._worker_init([])
    t0 = time.time()
    conn = K.get_db()
    db_by_path = {r[0]: {"status": r[1], "size": r[2], "mtime": r[3]} for r in
                  conn.execute("SELECT path, status, size, mtime FROM docs")}
    disk = {}
    for (p, rel, st, ext) in K.iter_files():
        disk[str(rel)] = {"size": st.st_size, "mtime": st.st_mtime, "ext": ext}

    added = updated = deleted = revived = 0
    guarded = False
    now_s = time.strftime("%Y-%m-%d %H:%M:%S")

    # 1) 删除：库里有、磁盘没有（含改名后旧路径消失）
    gone = [rel for rel, d in db_by_path.items()
            if d["status"] == 1 and rel not in disk]
    if len(gone) > BULK_DELETE_GUARD and not force:
        guarded = True
    else:
        for rel in gone:
            conn.execute("UPDATE docs SET status=0, updated_at=? WHERE path=?", (now_s, rel))
            deleted += 1
            _log_sync(conn, "delete", rel, "tombstone %d天" % TOMBSTONE_DAYS)

    # 2) 分类任务：新增 / 复活 / 修改
    todo = []
    for rel, d in disk.items():
        db = db_by_path.get(rel)
        if db is None:
            todo.append(("add", rel, d))
        elif db["status"] == 0:
            if (db["size"], db["mtime"]) == (d["size"], d["mtime"]):
                todo.append(("revive", rel, d))       # 删除后又原样恢复
            else:
                todo.append(("update", rel, d))       # 删除期间被改过，重新抽取
        elif (db["size"], db["mtime"]) != (d["size"], d["mtime"]):
            todo.append(("update", rel, d))

    wiki_map = K.load_wiki_map()
    for action, rel, d in todo:
        if action == "revive":
            conn.execute("UPDATE docs SET status=1, updated_at=? WHERE path=?",
                         (now_s, rel))
            revived += 1
            _log_sync(conn, "revive", rel)
            continue
        path_str = str(K.ROOT / rel)
        r = B.process_file((path_str, rel, d["size"], d["mtime"], d["ext"]))
        if not r["text"]:
            r["meta"] = K.derive_meta(rel, Path(rel).stem, "", d.get("mtime"))
            r["text_hash"] = ""
            r["tokens"] = ""
        else:
            r["text_hash"] = K.sha_text(r["text"])
        # wiki 页映射（与 build_index 同规则）
        stem = Path(rel).name
        m0 = r["meta"]
        wiki_page = ""
        if m0.get("client") and "__" not in stem:
            wiki_page = wiki_map.get(f"{m0['client']}__{stem}", "") or wiki_map.get(stem, "")
        else:
            wiki_page = wiki_map.get(stem, "")
        r["wiki_page"] = wiki_page
        _insert_text_and_doc(conn, r)
        if action == "add":
            added += 1
        elif action == "update":
            updated += 1
        else:
            revived += 1
        _log_sync(conn, action, rel, r["err"] or "")

    if added or updated or deleted or revived:
        K.enrich_from_wiki(conn)
        K.rebuild_project_profiles(conn)

    purged = purge_old_tombstones(conn)

    conn.commit()
    state = {"last_scan": time.strftime("%Y-%m-%d %H:%M:%S"),
             "added": added, "updated": updated, "deleted": deleted,
             "revived": revived, "purged": purged, "guarded": guarded,
             "elapsed_sec": round(time.time() - t0, 1)}
    STATE.write_text(json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8")
    conn.close()
    print("[sync]", json.dumps(state, ensure_ascii=False), flush=True)
    return state


def reload_cfg():
    """热加载配置（网页改完文件夹后，常驻同步进程下一轮自动生效）。"""
    global INTERVAL, TOMBSTONE_DAYS
    K.reload_config()
    INTERVAL = int(K.CFG.get("sync_interval_sec", 300))
    TOMBSTONE_DAYS = int(K.CFG.get("tombstone_days", 30))


def main():
    mode = sys.argv[1] if len(sys.argv) > 1 else "once"
    force = "--force" in sys.argv
    if mode == "once":
        scan_once(force=force)
    elif mode == "loop":
        # 写下 PID，供「停止服务.bat」精确结束常驻同步进程
        try:
            spath = K.APP / "data" / "sync.pid"
            spath.parent.mkdir(parents=True, exist_ok=True)
            spath.write_text(str(os.getpid()), encoding="ascii")
        except Exception:
            pass
        print(f"[sync] 常驻模式：每 {INTERVAL}s 增量扫描一次", flush=True)
        while True:
            try:
                reload_cfg()
                scan_once(force=force)
            except Exception as e:
                print("[sync] error:", str(e)[:200], flush=True)
            time.sleep(INTERVAL)


if __name__ == "__main__":
    main()
