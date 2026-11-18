# -*- coding: utf-8 -*-
"""单独重跑历史检索式抽取（不重建索引）。"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import kbcommon as K

if __name__ == "__main__":
    conn = K.get_db()
    n = K.extract_queries(conn)
    print(f"抽取历史检索式 {n} 条")
    rows = conn.execute(
        "SELECT project_code, substr(text,1,100) FROM queries "
        "ORDER BY id DESC LIMIT 8").fetchall()
    for r in rows:
        print("  ", r[0] or "(无项目号)", "|", r[1].replace("\n", " ")[:90])
