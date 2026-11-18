# -*- coding: utf-8 -*-
"""建库后收尾：修正非法年份 → wiki 元数据回填 → 重建项目画像 → 重抽历史检索式。"""
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import kbcommon as K


def fix_years(conn):
    cur = time.localtime().tm_year
    rows = conn.execute("SELECT id, year, mtime FROM docs").fetchall()
    n = 0
    for did, year, mtime in rows:
        bad = False
        if year:
            try:
                y = int(year)
            except Exception:
                y = 0
            bad = not (1990 <= y <= cur + 1)
        if bad:
            y2 = time.localtime(mtime or time.time()).tm_year
            conn.execute("UPDATE docs SET year=? WHERE id=?", (str(y2), did))
            n += 1
    conn.commit()
    return n


def fix_clients(conn):
    """按最新规则从路径重算客户（修正顶层目录名被当成客户的问题）。"""
    rows = conn.execute("SELECT id, path, title, client FROM docs").fetchall()
    n = 0
    for did, path, title, client in rows:
        new = K.derive_client(path, title)
        if new != client:
            conn.execute("UPDATE docs SET client=? WHERE id=?", (new, did))
            n += 1
    conn.commit()
    return n


if __name__ == "__main__":
    conn = K.get_db()
    print("重算客户字段:", fix_clients(conn), "条")
    print("修正非法年份:", fix_years(conn), "条")
    print("wiki 元数据回填:", K.enrich_from_wiki(conn), "条")
    K.rebuild_project_profiles(conn)
    print("项目画像重建完成:", conn.execute("SELECT count(*) FROM projects").fetchone()[0], "组")
    print("历史检索式:", K.extract_queries(conn), "条")
    stats = K.count_stats(conn)
    print("最终统计:", stats)
