import os
import sqlite3
import sys
import tempfile


def check_python() -> None:
    print(f"Python: {sys.executable}")
    print(f"Version: {sys.version}")


def check_sqlite_write() -> None:
    tmp_dir = tempfile.mkdtemp()
    db_path = os.path.join(tmp_dir, "temp_check.db")
    try:
        conn = sqlite3.connect(db_path)
        conn.execute("CREATE TABLE test (id INTEGER PRIMARY KEY)")
        conn.execute("INSERT INTO test DEFAULT VALUES")
        conn.commit()
        conn.close()
        os.remove(db_path)
        os.rmdir(tmp_dir)
        print("SQLite write: OK")
    except Exception as exc:
        raise RuntimeError(
            "SQLite 无法写入临时数据库，请检查磁盘权限或路径可写性。"
        ) from exc


def check_imports() -> None:
    try:
        import pandas  # noqa: F401
        import openpyxl  # noqa: F401
        import docx  # noqa: F401
        import qrcode  # noqa: F401
        import PIL  # noqa: F401
    except Exception as exc:
        raise RuntimeError(
            "缺少 pandas/openpyxl/python-docx/qrcode/pillow，请执行: pip install pandas openpyxl python-docx qrcode[pil]"
        ) from exc

    try:
        import flask  # noqa: F401
    except Exception as exc:
        raise RuntimeError("缺少 Flask，请执行: pip install flask") from exc


def main() -> None:
    try:
        check_python()
        check_sqlite_write()
        check_imports()
        print("环境检查通过。")
    except Exception as exc:
        print(f"环境检查失败: {exc}")
        print("请先修复上述问题，然后重新运行 env_check.py。")
        sys.exit(1)


if __name__ == "__main__":
    main()
