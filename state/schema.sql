-- 自动报销系统 台账建表语句
-- 首次运行前执行本文件初始化 state/ledger.sqlite
CREATE TABLE entries (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        message_id TEXT UNIQUE,
        chat_id TEXT,
        driver TEXT,
        receipt_date TEXT,
        amount REAL,
        category TEXT,
        image_sha256 TEXT,
        phash TEXT,
        business_key TEXT,
        report_file TEXT,
        sheet TEXT,
        cell TEXT,
        old_value REAL,
        new_value REAL,
        run_id TEXT,
        created_at TEXT
    , task_id TEXT, task_url TEXT, task_status TEXT DEFAULT 'none', reimbursed INTEGER DEFAULT 0, reimbursed_at TEXT, no_report INTEGER DEFAULT 0, order_no TEXT, pay_time TEXT, voided INTEGER DEFAULT 0, voided_at TEXT, void_reason TEXT);

CREATE TABLE sqlite_sequence(name,seq);
