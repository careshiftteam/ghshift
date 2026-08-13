"""
GHShift v4 - 全事業所対応・事業所登録・総合メニュー
"""
import sqlite3, json, calendar, csv, re, io, os
from datetime import date, datetime, timedelta
import hashlib, uuid, secrets
from collections import defaultdict
from flask import Flask, request, jsonify, g, render_template, redirect

try:
    from flask_cors import CORS
    HAS_CORS = True
except ImportError:
    HAS_CORS = False

try:
    import openpyxl
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

app = Flask(__name__)
# このアプリはフロントエンド(HTML)とAPIを同じFlaskから同一オリジンで配信しており、
# 本来ブラウザからのCORSは不要（別ドメインのJSから叩かれる想定がない）。
# 全許可(CORS(app))は他サイト経由での不正な操作リスクを高めるため、有効化しない。

# ── シフト区分定数 ──────────────────────────────────────
# 非勤務扱いのシフト種別（就業日数カウント対象外）
NON_WORK_SHIFTS = ("休", "希休", "有給", "欠勤", None)
# 希望入力のreq_type → シフト種別マッピング
REQ_TYPE_MAP = {"rest": "希休", "hol": "有給", "absence": "欠勤"}

@app.after_request
def no_cache(response):
    response.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
    response.headers["Pragma"] = "no-cache"
    response.headers["Expires"] = "0"
    return response

DATABASE = os.environ.get("CARESHIFT_DB", "careshift.db")

def get_db():
    db = getattr(g, "_db", None)
    if db is None:
        db = g._db = sqlite3.connect(DATABASE)
        db.row_factory = sqlite3.Row
        db.execute("PRAGMA foreign_keys = ON")
    return db

@app.teardown_appcontext
def close_db(e):
    db = getattr(g, "_db", None)
    if db: db.close()

def qdb(sql, args=(), one=False):
    cur = get_db().execute(sql, args)
    rv = cur.fetchall()
    return (rv[0] if rv else None) if one else rv

def xdb(sql, args=()):
    db = get_db()
    cur = db.execute(sql, args)
    db.commit()
    return cur.lastrowid

SCHEMA = """
CREATE TABLE IF NOT EXISTS facilities (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    code        TEXT NOT NULL UNIQUE,
    name        TEXT NOT NULL,
    type        TEXT NOT NULL,
    color       TEXT DEFAULT '#5b8fff',
    auto_shift  INTEGER DEFAULT 1,
    zip_code    TEXT,
    address     TEXT,
    tel         TEXT,
    fax         TEXT,
    email       TEXT,
    manager     TEXT,
    note        TEXT,
    sort_order  INTEGER DEFAULT 0,
    created_at  TEXT DEFAULT (datetime('now','localtime'))
);

CREATE TABLE IF NOT EXISTS units (
    id               INTEGER PRIMARY KEY AUTOINCREMENT,
    facility_id      INTEGER NOT NULL REFERENCES facilities(id),
    unit_no          INTEGER NOT NULL,
    name             TEXT NOT NULL,
    residents        INTEGER DEFAULT 0,
    color            TEXT,
    night_staff_need INTEGER DEFAULT 1,  -- 1日の必要夜勤人数（加算取得時は2）
    UNIQUE(facility_id, unit_no)
);

CREATE TABLE IF NOT EXISTS staff (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    facility_id     INTEGER NOT NULL REFERENCES facilities(id),
    unit_id         INTEGER NOT NULL REFERENCES units(id),
    name            TEXT NOT NULL,
    role            TEXT,
    employment_type TEXT DEFAULT 'full_time',
    join_date       TEXT,
    leave_date      TEXT,
    is_active       INTEGER NOT NULL DEFAULT 1,
    monthly_limit   INTEGER NOT NULL DEFAULT 21,
    night_target    INTEGER DEFAULT 0,
    can_day         INTEGER DEFAULT 1,
    can_early       INTEGER DEFAULT 1,
    can_late        INTEGER DEFAULT 1,
    can_night       INTEGER DEFAULT 0,
    can_night_only  INTEGER DEFAULT 0,
    is_help         INTEGER DEFAULT 0,
    is_approver     INTEGER DEFAULT 0,    -- シフト承認者フラグ(1=リーダー承認可)
    note            TEXT,
    login_id        TEXT UNIQUE,          -- ログインID
    password_hash   TEXT,                 -- パスワード(ハッシュ)
    system_role     TEXT DEFAULT 'staff', -- staff/leader/admin
    created_at      TEXT DEFAULT (datetime('now','localtime'))
);

-- 保存前バックアップ（直前に戻す機能用）
CREATE TABLE IF NOT EXISTS shift_entries_backup (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    version_id  INTEGER NOT NULL REFERENCES shift_versions(id),
    staff_id    INTEGER NOT NULL,
    unit_id     INTEGER NOT NULL,
    date        TEXT NOT NULL,
    shift_type  TEXT,
    is_cross    INTEGER DEFAULT 0,
    from_unit_id INTEGER,
    backed_at   TEXT DEFAULT (datetime('now','localtime'))
);
CREATE TABLE IF NOT EXISTS shift_approvals (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    version_id  INTEGER NOT NULL REFERENCES shift_versions(id),
    staff_id    INTEGER NOT NULL REFERENCES staff(id),
    unit_id     INTEGER NOT NULL REFERENCES units(id),
    approved_at TEXT NOT NULL DEFAULT (datetime('now','localtime')),
    comment     TEXT,
    UNIQUE(version_id, staff_id)
);

-- 曜日固定設定
CREATE TABLE IF NOT EXISTS staff_fixed_days (
    id         INTEGER PRIMARY KEY AUTOINCREMENT,
    staff_id   INTEGER NOT NULL REFERENCES staff(id),
    dow        INTEGER NOT NULL, -- 0=月,1=火,...,6=日
    shift_type TEXT NOT NULL,    -- '早','遅','日','夜'
    UNIQUE(staff_id, dow)
);

CREATE TABLE IF NOT EXISTS staff_skills (
    id        INTEGER PRIMARY KEY AUTOINCREMENT,
    staff_id  INTEGER NOT NULL REFERENCES staff(id),
    unit_id   INTEGER NOT NULL REFERENCES units(id),
    level     TEXT NOT NULL DEFAULT 'ok',
    UNIQUE(staff_id, unit_id)
);

CREATE TABLE IF NOT EXISTS requests (
    id         INTEGER PRIMARY KEY AUTOINCREMENT,
    staff_id   INTEGER NOT NULL REFERENCES staff(id),
    year       INTEGER NOT NULL,
    month      INTEGER NOT NULL,
    day        INTEGER NOT NULL,
    req_type   TEXT NOT NULL,
    priority   TEXT DEFAULT 'prefer',
    note       TEXT,
    created_at TEXT DEFAULT (datetime('now','localtime')),
    UNIQUE(staff_id, year, month, day)
);

CREATE TABLE IF NOT EXISTS shift_versions (
    id            INTEGER PRIMARY KEY AUTOINCREMENT,
    facility_id   INTEGER REFERENCES facilities(id),
    year          INTEGER NOT NULL,
    month         INTEGER NOT NULL,
    version_no    INTEGER NOT NULL DEFAULT 1,
    status        TEXT NOT NULL DEFAULT 'draft',
    change_reason TEXT,
    created_by    TEXT,
    approved_by   TEXT,
    created_at    TEXT DEFAULT (datetime('now','localtime')),
    approved_at   TEXT,
    UNIQUE(facility_id, year, month, version_no)
);

CREATE TABLE IF NOT EXISTS shift_entries (
    id           INTEGER PRIMARY KEY AUTOINCREMENT,
    version_id   INTEGER NOT NULL REFERENCES shift_versions(id),
    staff_id     INTEGER NOT NULL REFERENCES staff(id),
    unit_id      INTEGER NOT NULL REFERENCES units(id),
    date         TEXT NOT NULL,
    shift_type   TEXT NOT NULL,
    is_cross     INTEGER DEFAULT 0,
    from_unit_id INTEGER REFERENCES units(id),
    is_manual    INTEGER DEFAULT 0,
    note         TEXT,
    UNIQUE(version_id, staff_id, date)
);

CREATE TABLE IF NOT EXISTS role_permissions (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    role        TEXT NOT NULL UNIQUE,  -- admin/manager/leader/staff
    password_hash TEXT,               -- ロール共通パスワード（任意）
    perm_req    INTEGER DEFAULT 1,    -- 希望日入力
    perm_gen    INTEGER DEFAULT 0,    -- 自動生成
    perm_create INTEGER DEFAULT 0,    -- シフト作成
    perm_change INTEGER DEFAULT 0,    -- シフト変更
    perm_approve INTEGER DEFAULT 0,   -- シフト承認
    perm_publish INTEGER DEFAULT 0,   -- シフト公開
    perm_staff   INTEGER DEFAULT 0,   -- 職員管理
    perm_facility INTEGER DEFAULT 0,  -- 事業所管理
    perm_settings INTEGER DEFAULT 0,  -- システム設定
    updated_at  TEXT DEFAULT (datetime('now','localtime'))
);


CREATE TABLE IF NOT EXISTS shift_dashboard_notes (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    facility_id INTEGER NOT NULL REFERENCES facilities(id),
    year        INTEGER NOT NULL,
    month       INTEGER NOT NULL,
    note        TEXT,
    updated_by  INTEGER REFERENCES staff(id),
    updated_at  TEXT DEFAULT (datetime('now','localtime')),
    UNIQUE(facility_id, year, month)
);

CREATE TABLE IF NOT EXISTS sessions (
    id         TEXT PRIMARY KEY,          -- セッションID(UUID)
    staff_id   INTEGER NOT NULL REFERENCES staff(id),
    created_at TEXT DEFAULT (datetime('now','localtime')),
    expires_at TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS absences (
    id               INTEGER PRIMARY KEY AUTOINCREMENT,
    staff_id         INTEGER NOT NULL REFERENCES staff(id),
    date             TEXT NOT NULL,
    type             TEXT NOT NULL,
    replace_staff_id INTEGER REFERENCES staff(id),
    note             TEXT,
    created_at       TEXT DEFAULT (datetime('now','localtime'))
);

-- ユニット別・シフト別の必要人員数
-- shift_type: '早'|'遅'|'夜'|'日'
-- required: その日そのユニットに最低何人必要か
-- is_admin_eligible: 人員不足時に管理ユニット職員を応援候補に含めるか
CREATE TABLE IF NOT EXISTS unit_required_staff (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    unit_id     INTEGER NOT NULL REFERENCES units(id),
    shift_type  TEXT NOT NULL CHECK(shift_type IN ('早','遅','夜','日')),
    required    INTEGER NOT NULL DEFAULT 1,
    is_admin_eligible INTEGER NOT NULL DEFAULT 0,
    UNIQUE(unit_id, shift_type)
);

-- 過去シフト（Excel/CSV）の評価ベースライン
CREATE TABLE IF NOT EXISTS shift_baselines (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    facility_id     INTEGER NOT NULL REFERENCES facilities(id),
    year            INTEGER NOT NULL,
    month           INTEGER NOT NULL,
    source_name     TEXT,             -- アップロードファイル名
    eval_json       TEXT,             -- _evaluate_schedule の結果
    unmatched_names TEXT,             -- マッチしなかった職員名(JSON配列)
    unknown_codes   TEXT,             -- 認識できなかったシフト記号(JSON: {記号:出現回数})
    created_at      TEXT DEFAULT (datetime('now','localtime')),
    UNIQUE(facility_id, year, month)
);

CREATE TABLE IF NOT EXISTS shift_baseline_entries (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    baseline_id INTEGER NOT NULL REFERENCES shift_baselines(id),
    staff_id    INTEGER NOT NULL REFERENCES staff(id),
    date        TEXT NOT NULL,
    shift_type  TEXT NOT NULL
);

-- 職員ごとのシフト種別希望回数（パート・派遣・夜勤専従）
-- shift_type: '早'|'遅'|'夜'|'日'
-- target: 月間の希望回数（極力満たす）
-- is_upper_limit: 1=上限として扱う(超えない), 0=目標として扱う(極力達成)
CREATE TABLE IF NOT EXISTS staff_shift_targets (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    staff_id    INTEGER NOT NULL REFERENCES staff(id),
    shift_type  TEXT NOT NULL CHECK(shift_type IN ('早','遅','夜','日')),
    target      INTEGER NOT NULL DEFAULT 0,
    is_upper_limit INTEGER NOT NULL DEFAULT 0,
    UNIQUE(staff_id, shift_type)
);

CREATE TABLE IF NOT EXISTS shift_actual (
    id           INTEGER PRIMARY KEY AUTOINCREMENT,
    version_id   INTEGER NOT NULL REFERENCES shift_versions(id),
    staff_id     INTEGER NOT NULL REFERENCES staff(id),
    date         INTEGER NOT NULL,
    actual_type  TEXT NOT NULL,
    time_from    TEXT NOT NULL,
    time_to      TEXT NOT NULL,
    work_hours   REAL NOT NULL DEFAULT 0,
    note         TEXT NOT NULL DEFAULT '',
    created_at   TEXT NOT NULL DEFAULT (datetime('now','localtime')),
    UNIQUE(version_id, staff_id, date)
);
"""

def init_db():
    with app.app_context():
        db = get_db()
        db.executescript(SCHEMA)
        # 既存DBへのカラム追加（既に存在する場合はエラーを無視）
        for sql in [
            "ALTER TABLE staff ADD COLUMN login_id TEXT",
            "ALTER TABLE staff ADD COLUMN password_hash TEXT",
            "ALTER TABLE staff ADD COLUMN system_role TEXT DEFAULT 'staff'",
            "ALTER TABLE staff ADD COLUMN is_approver INTEGER DEFAULT 0",
            "ALTER TABLE shift_versions ADD COLUMN eval_json TEXT",
            "ALTER TABLE shift_versions ADD COLUMN based_on_version_id INTEGER",
            "ALTER TABLE shift_versions ADD COLUMN diff_cells_json TEXT",
            "ALTER TABLE role_permissions ADD COLUMN perm_staff INTEGER DEFAULT 0",
            "ALTER TABLE role_permissions ADD COLUMN perm_facility INTEGER DEFAULT 0",
            "ALTER TABLE role_permissions ADD COLUMN perm_settings INTEGER DEFAULT 0",
        ]:
            try: db.execute(sql)
            except: pass
        db.commit()
        _seed()

def _ensure_default_permissions():
    """adminロールのみデフォルト権限を作成。他のロールはユーザーが追加する"""
    existing = qdb("SELECT id FROM role_permissions WHERE role='admin' LIMIT 1", one=True)
    if existing:
        return
    db = get_db()
    db.execute(
        "INSERT OR IGNORE INTO role_permissions"
        "(role,perm_req,perm_gen,perm_create,perm_change,perm_approve,perm_publish)"
        " VALUES('admin',1,1,1,1,1,1)")
    db.commit()

def _ensure_default_admin():
    """デフォルト管理者が存在しない場合に作成"""
    existing = qdb("SELECT id FROM staff WHERE login_id='admin' LIMIT 1", one=True)
    if existing:
        return
    first_staff = qdb("SELECT id FROM staff WHERE is_active=1 LIMIT 1", one=True)
    if first_staff:
        xdb("UPDATE staff SET login_id='admin',password_hash=?,system_role='admin' WHERE id=?",
            (hash_password("admin1234"), first_staff["id"]))
        get_db().commit()


def _seed():
    if qdb("SELECT id FROM facilities LIMIT 1"):
        _ensure_default_admin()
        _ensure_default_permissions()
        return
    db = get_db()
    facs = [
        ("GH1","GH1 認知症対応型共同生活介護","GH","#5b8fff",1,0),
        ("GH2","GH2 認知症対応型共同生活介護","GH","#a57eff",1,1),
        ("GH3","GH3 認知症対応型共同生活介護","GH","#3dd6c0",1,2),
        ("DS1","DS1 通所介護","DS","#f0a020",1,3),
        ("DS2","DS2 通所介護","DS","#f0a020",1,4),
        ("CPC1","CPC1 居宅介護支援事業所","CPC","#40c060",0,5),
        ("WGC1","WGC1 福祉用具貸与事業所","WGC","#f05060",0,6),
    ]
    for code,name,ftype,color,auto,sort in facs:
        db.execute("INSERT INTO facilities(code,name,type,color,auto_shift,sort_order) VALUES(?,?,?,?,?,?)",
                   (code,name,ftype,color,auto,sort))
    units = [
        ("GH1",1,"ユニット1",6,"#5b8fff"),
        ("GH2",1,"寿庵",8,"#5b8fff"),("GH2",2,"楽庵",8,"#a57eff"),
        ("GH3",1,"彩咲庵",9,"#5b8fff"),("GH3",2,"結咲庵",9,"#a57eff"),("GH3",3,"喜咲庵",9,"#3dd6c0"),("GH3",4,"GH3管理",0,"#888888"),
        ("DS1",1,"DS1",0,"#f0a020"),
        ("DS2",1,"DS2",0,"#f0a020"),
        ("CPC1",1,"居宅支援",0,"#40c060"),
        ("WGC1",1,"福祉用具",0,"#f05060"),
    ]
    for fc,un,nm,res,col in units:
        f=qdb("SELECT id FROM facilities WHERE code=?",(fc,),one=True)
        if f: db.execute("INSERT INTO units(facility_id,unit_no,name,residents,color) VALUES(?,?,?,?,?)",
                          (f["id"],un,nm,res,col))

    # GH系ユニットの必要人員初期設定
    # GH(グループホーム): 早1・遅1・夜1 が基本
    # 管理ユニット(unit_no=4, residents=0)は除外
    # is_admin_eligible=1: 人員不足時に管理ユニット職員を応援候補に含める
    gh_fac_codes = ["GH1","GH2","GH3"]
    for fc in gh_fac_codes:
        f = qdb("SELECT id FROM facilities WHERE code=?", (fc,), one=True)
        if not f: continue
        gh_units = qdb(
            "SELECT id FROM units WHERE facility_id=? AND residents > 0", (f["id"],))
        for u in gh_units:
            for shift_type, req, admin_ok in [("早",1,1),("遅",1,0),("夜",1,0)]:
                db.execute(
                    """INSERT OR IGNORE INTO unit_required_staff
                       (unit_id, shift_type, required, is_admin_eligible)
                       VALUES(?,?,?,?)""",
                    (u["id"], shift_type, req, admin_ok))
    staff_rows = [
        ("GH1",1,"山本 隆","ユニットリーダー","2020-04-01",21,1,1,1,0,0,0),
        ("GH1",1,"鈴木 良太","夜勤専従","2019-06-01",21,0,0,0,1,1,0),
        ("GH1",1,"田中 恵美","介護職員","2021-01-15",21,1,1,1,1,0,0),
        ("GH2",1,"伊藤 美穂","介護福祉士","2018-04-01",21,1,1,1,0,0,0),
        ("GH2",1,"小林 大輔","夜勤専従","2020-09-01",21,0,0,0,1,1,0),
        ("GH2",2,"加藤 さくら","介護職員","2022-04-01",21,1,1,1,1,0,0),
        ("GH2",2,"前田 雄介","介護職員","2021-07-01",21,1,1,1,0,0,0),
        ("GH2",2,"石川 由美","パート","2023-01-10",15,1,0,0,0,0,0),
        # GH3: R8年度1,2,4,5月のシフト実績(PDF)から抽出した実職員33名
        # unit_no: 1=彩咲庵, 2=結咲庵, 3=喜咲庵, 4=GH3管理(管理者/ケアマネ/事務)
        ("GH3", 1, "河野 聖也", "ユニットリーダー", None, 21, 1, 1, 1, 1, 0, 0),
        ("GH3", 4, "三坂 伸子", "計画作成 ケアマネ", None, 11, 1, 0, 0, 0, 0, 0),
        ("GH3", 1, "上村 真紀", "夜勤専従", None, 19, 0, 0, 0, 1, 1, 0),
        ("GH3", 4, "上林 尊行", "計画作成 ケアマネ", None, 4, 1, 0, 0, 0, 0, 0),
        ("GH3", 4, "北畑 奈緒美", "事務", None, 19, 1, 0, 0, 0, 0, 0),
        ("GH3", 1, "吉田 有紀子", "夜勤専従", None, 6, 0, 0, 0, 1, 1, 0),
        ("GH3", 1, "喜連 篤仁", "夜勤専従", None, 3, 0, 0, 0, 1, 1, 0),
        ("GH3", 1, "大槻 慶太", "介護職員", None, 20, 0, 1, 1, 1, 0, 0),
        ("GH3", 1, "寺坂 幸", "介護職員", None, 21, 0, 1, 1, 1, 0, 0),
        ("GH3", 4, "岩本 力", "管理者", None, 19, 1, 1, 0, 0, 0, 0),
        ("GH3", 1, "松友 えつこ", "介護職員", None, 21, 0, 1, 1, 1, 0, 0),
        ("GH3", 1, "津村 浩一", "夜勤専従", None, 20, 0, 0, 0, 1, 1, 0),
        ("GH3", 1, "豊田百合子", "介護職員", None, 16, 0, 1, 1, 0, 0, 0),
        ("GH3", 2, "山本 飛鳥", "ユニットリーダー", None, 19, 1, 1, 1, 1, 0, 0),
        ("GH3", 2, "中村 千恵", "介護職員", None, 15, 0, 1, 1, 0, 0, 0),
        ("GH3", 2, "十川 順平", "夜勤専従", None, 17, 0, 0, 0, 1, 1, 0),
        ("GH3", 2, "寺前 毬絵", "夜勤専従", None, 7, 0, 0, 0, 1, 1, 0),
        ("GH3", 2, "梛 陽美", "介護職員", None, 20, 0, 1, 1, 1, 0, 0),
        ("GH3", 2, "楠本 篤美", "夜勤専従", None, 12, 0, 0, 0, 1, 1, 0),
        ("GH3", 2, "濵端 直子", "介護職員", None, 20, 0, 1, 1, 0, 0, 0),
        ("GH3", 2, "薮野 香奈", "夜勤専従", None, 4, 0, 0, 0, 1, 1, 0),
        ("GH3", 2, "飛澤 京子", "夜勤専従", None, 14, 0, 0, 0, 1, 1, 0),
        ("GH3", 2, "龍田 淳一", "夜勤専従", None, 2, 0, 0, 0, 1, 1, 0),
        ("GH3", 3, "篠崎 理沙", "ユニットリーダー", None, 20, 1, 1, 1, 1, 0, 0),
        ("GH3", 3, "上田 美恵子", "介護職員", None, 18, 0, 1, 1, 0, 0, 0),
        ("GH3", 3, "中西 佐和", "夜勤専従", None, 21, 0, 0, 0, 1, 1, 0),
        ("GH3", 3, "伊藤 杏", "介護職員", None, 8, 0, 1, 1, 0, 0, 0),
        ("GH3", 3, "岡本 純子", "介護職員", None, 2, 0, 1, 1, 0, 0, 0),
        ("GH3", 3, "戸田 優香", "介護職員", None, 18, 0, 1, 1, 0, 0, 0),
        ("GH3", 3, "浦部 あゆみ", "介護職員", None, 10, 0, 1, 1, 1, 0, 0),
        ("GH3", 3, "田川 幸子", "介護職員", None, 20, 0, 1, 1, 1, 0, 0),
        ("GH3", 3, "畠山 典平", "夜勤専従", None, 10, 0, 0, 0, 1, 1, 0),
        ("GH3", 3, "高松 明彦", "介護職員", None, 4, 0, 1, 1, 0, 0, 0),
        ("DS1",1,"佐藤 花子","介護福祉士","2019-04-01",21,1,1,1,0,0,0),
        ("DS1",1,"高橋 真一","介護職員","2020-04-01",21,1,1,0,0,0,0),
        ("DS2",1,"林 由美","介護福祉士","2018-04-01",21,1,1,1,0,0,0),
        ("DS2",1,"清水 俊介","介護職員","2021-04-01",21,1,1,0,0,0,0),
        ("CPC1",1,"佐々木 優子","ケアマネジャー","2015-04-01",21,1,0,0,0,0,1),
        ("CPC1",1,"中島 浩二","ケアマネジャー","2018-04-01",21,1,0,0,0,0,1),
        ("WGC1",1,"岡田 誠","福祉用具専門員","2019-04-01",21,1,0,0,0,0,1),
    ]
    for fc,un,nm,role,jd,lim,cd,ce,cl,cn,cno,hlp in staff_rows:
        f=qdb("SELECT id FROM facilities WHERE code=?",(fc,),one=True)
        u=qdb("SELECT id FROM units WHERE facility_id=? AND unit_no=?",(f["id"],un),one=True)
        if f and u:
            db.execute("""INSERT INTO staff(facility_id,unit_id,name,role,join_date,monthly_limit,
                       can_day,can_early,can_late,can_night,can_night_only,is_help)
                       VALUES(?,?,?,?,?,?,?,?,?,?,?,?)""",
                       (f["id"],u["id"],nm,role,jd,lim,cd,ce,cl,cn,cno,hlp))
    # GH3: 夜勤専従/夜勤対応職員のnight_target（月平均夜勤回数、PDF実績から算出）
    gh3_night_targets = {
        "河野 聖也": 7, "上村 真紀": 9, "吉田 有紀子": 2, "喜連 篤仁": 1, "大槻 慶太": 2,
        "寺坂 幸": 6, "松友 えつこ": 4, "津村 浩一": 10, "山本 飛鳥": 4, "十川 順平": 8,
        "寺前 毬絵": 3, "梛 陽美": 2, "楠本 篤美": 6, "薮野 香奈": 2, "飛澤 京子": 7,
        "龍田 淳一": 1, "篠崎 理沙": 4, "中西 佐和": 11, "浦部 あゆみ": 3, "田川 幸子": 6,
        "畠山 典平": 5,
    }
    gh3 = qdb("SELECT id FROM facilities WHERE code='GH3'", one=True)
    if gh3:
        for nm, nt in gh3_night_targets.items():
            db.execute("UPDATE staff SET night_target=? WHERE facility_id=? AND name=?",
                       (nt, gh3["id"], nm))
    for s in qdb("SELECT id,unit_id FROM staff"):
        db.execute("INSERT OR IGNORE INTO staff_skills(staff_id,unit_id,level) VALUES(?,?,'ok')",
                   (s["id"],s["unit_id"]))
    db.commit()
    print("✓ 初期データ投入完了")

# ── ページ ───────────────────────────────

# ── 認証ヘルパー ──────────────────────────────────────
from werkzeug.security import generate_password_hash, check_password_hash

def hash_password(pw):
    """新規・変更されたパスワードは、ソルト付きの安全な方式(PBKDF2)でハッシュ化する。"""
    return generate_password_hash(pw)

def verify_password(stored_hash, pw):
    """保存済みハッシュとパスワードを照合する。
    新形式(werkzeugのpbkdf2:.../scrypt:...)と、
    移行期間中に残っている旧形式(単純なsha256)の両方に対応する。"""
    if not stored_hash:
        return False
    if stored_hash.startswith("pbkdf2:") or stored_hash.startswith("scrypt:"):
        return check_password_hash(stored_hash, pw)
    return stored_hash == hashlib.sha256(pw.encode()).hexdigest()

def get_session_staff():
    """現在のセッションからスタッフ情報を取得"""
    sid = request.cookies.get("session_id")
    if not sid: return None
    row = qdb(
        "SELECT s.*, f.code as fac_code FROM sessions se "
        "JOIN staff s ON se.staff_id=s.id "
        "JOIN facilities f ON s.facility_id=f.id "
        "WHERE se.id=? AND se.expires_at > datetime('now','localtime')",
        (sid,), one=True)
    return dict(row) if row else None

def require_login(min_role=None):
    """ログイン必須デコレータ用チェック。未ログインはNoneを返す"""
    u = get_session_staff()
    if not u: return None
    ROLE_ORDER = {"staff":1, "leader":2, "admin":3}
    if min_role and ROLE_ORDER.get(u.get("system_role","staff"),1) < ROLE_ORDER.get(min_role,1):
        return None
    return u

def require_perm(*perm_names):
    """ログイン必須 + 指定した権限（いずれか1つでも）を持つか確認する。
    admin ロールは常に許可。
    戻り値: (user, None) で許可 / (None, (response, status)) で拒否。"""
    u = get_session_staff()
    if not u:
        return None, (jsonify({"ok": False, "error": "ログインが必要です"}), 401)
    role = (u.get("system_role") or "staff").strip()
    if role == "admin":
        return u, None
    perm = qdb("SELECT * FROM role_permissions WHERE role=?", (role,), one=True)
    if not perm or not any(perm[p] for p in perm_names):
        return None, (jsonify({"ok": False, "error": "この操作を行う権限がありません"}), 403)
    return u, None

def require_any_login():
    """役職を問わず、ログインしていることだけを確認する。"""
    u = get_session_staff()
    if not u:
        return None, (jsonify({"ok": False, "error": "ログインが必要です"}), 401)
    return u, None

# ── 認証API ──────────────────────────────────────────
@app.route("/api/auth/login", methods=["POST"])
def api_login():
    d = request.get_json(silent=True) or {}
    login_id = d.get("login_id","").strip()
    password  = d.get("password","")

    row = qdb(
        "SELECT s.*, f.code as fac_code, rp.password_hash as role_password_hash "
        "FROM staff s "
        "JOIN facilities f ON s.facility_id=f.id "
        "LEFT JOIN role_permissions rp ON rp.role=s.system_role "
        "WHERE s.login_id=? AND s.is_active=1",
        (login_id,), one=True)

    staff_hash = row["password_hash"] if row else None
    role_hash  = row["role_password_hash"] if row else None

    # 認証順序:
    # 1. 職員個別パスワード staff.password_hash
    # 2. ロール共通パスワード role_permissions.password_hash
    # どちらか一致すればログイン成功。
    staff_ok = verify_password(staff_hash, password)
    role_ok  = verify_password(role_hash, password)
    if not row or not (staff_ok or role_ok):
        return jsonify({"ok":False, "error":"IDまたはパスワードが違います"}), 401

    # 旧形式（単純sha256）で一致した場合、次回以降のために新形式へ自動で移行する。
    def _is_legacy(h):
        return h and not (h.startswith("pbkdf2:") or h.startswith("scrypt:"))
    if staff_ok and _is_legacy(staff_hash):
        xdb("UPDATE staff SET password_hash=? WHERE id=?", (hash_password(password), row["id"]))
    if role_ok and _is_legacy(role_hash):
        xdb("UPDATE role_permissions SET password_hash=? WHERE role=?", (hash_password(password), row["system_role"]))

    # セッション作成（24時間有効）
    sid = secrets.token_hex(32)
    expires = (datetime.now() + __import__('datetime').timedelta(hours=24)).strftime("%Y-%m-%d %H:%M:%S")
    xdb("INSERT INTO sessions(id,staff_id,expires_at) VALUES(?,?,?)",
        (sid, row["id"], expires))
    get_db().commit()
    resp = jsonify({
        "ok": True,
        "user": {
            "id": row["id"], "name": row["name"],
            "system_role": row["system_role"] or "staff",
            "fac_code": row["fac_code"]
        }
    })
    # 注意: 現在はHTTP運用のためsecure=Trueは付けていない。
    # SSL化（Let's Encrypt等）が完了したら、必ず secure=True を追加すること。
    resp.set_cookie("session_id", sid, httponly=True, max_age=86400, samesite="Lax")
    return resp

@app.route("/api/auth/logout", methods=["POST"])
def api_logout():
    sid = request.cookies.get("session_id")
    if sid: xdb("DELETE FROM sessions WHERE id=?", (sid,))
    get_db().commit()
    resp = jsonify({"ok": True})
    resp.delete_cookie("session_id")
    return resp

@app.route("/api/auth/me")
def api_me():
    u = get_session_staff()
    if not u: return jsonify({"ok":False}), 401
    # ロールの権限設定を取得
    role = u.get("system_role", "staff")
    rp = qdb("SELECT * FROM role_permissions WHERE role=?", (role,), one=True)
    perms = {
        "perm_req":      int(rp["perm_req"]      if rp else 1),
        "perm_gen":      int(rp["perm_gen"]      if rp else 0),
        "perm_adjust":   int((rp["perm_create"] or rp["perm_change"]) if rp else 0),
        "perm_approve":  int(rp["perm_approve"]  if rp else 0),
        "perm_publish":  int(rp["perm_publish"]  if rp else 0),
        "perm_staff":    int(rp["perm_staff"]    if rp else 0),
        "perm_facility": int(rp["perm_facility"] if rp else 0),
        "perm_settings": int(rp["perm_settings"] if rp else 0),
    } if rp else {
        "perm_req":1,"perm_gen":0,"perm_adjust":0,"perm_approve":0,"perm_publish":0,
        "perm_staff":0,"perm_facility":0,"perm_settings":0
    }
    # adminは常に全権限
    if role == "admin":
        perms = {k:1 for k in perms}
    return jsonify({"ok":True, "user":{
        "id":u["id"], "name":u["name"],
        "system_role": role,
        "is_approver": int(u.get("is_approver", 0)),
        "fac_code":u.get("fac_code",""),
        "perms": perms
    }})


@app.route("/api/permissions", methods=["GET"])
def api_get_permissions():
    """ロール一覧。パスワードハッシュそのものは返さない。"""
    u, err = require_any_login()
    if err: return err
    rows = qdb("SELECT * FROM role_permissions ORDER BY CASE WHEN role='admin' THEN 0 ELSE 1 END, id")
    result = []
    for row in rows:
        d = dict(row)
        result.append({
            "id": d.get("id"),
            "role": d.get("role", ""),
            "has_password": bool(d.get("password_hash")),
            "perm_req": int(d.get("perm_req") or 0),
            "perm_gen": int(d.get("perm_gen") or 0),
            # 旧perm_create / perm_changeは、画面上では「シフト確認・調整」に統合する
            "perm_adjust": int(bool(d.get("perm_create")) or bool(d.get("perm_change"))),
            "perm_approve": int(d.get("perm_approve") or 0),
            "perm_publish": int(d.get("perm_publish") or 0),
            "perm_staff":    int(d.get("perm_staff")    or 0),
            "perm_facility": int(d.get("perm_facility") or 0),
            "perm_settings": int(d.get("perm_settings") or 0),
            "updated_at": d.get("updated_at"),
        })
    return jsonify(result)

@app.route("/api/permissions", methods=["POST"])
def api_save_permissions():
    """ロール別パスワード・権限設定を保存する。パスワード変更時は同一system_roleの職員ログインにも同期する。"""
    u = require_login("admin")
    if not u:
        return jsonify({"ok":False,"error":"権限がありません"}), 403

    items = request.get_json(silent=True) or []
    if isinstance(items, dict):
        items = [items]
    if not isinstance(items, list):
        return jsonify({"ok":False,"error":"保存データの形式が不正です"}), 400

    db = get_db()
    saved_roles = []
    for item in items:
        role = str(item.get("role", "")).strip()
        if not role:
            return jsonify({"ok":False,"error":"ロール名を入力してください"}), 400

        pw = str(item.get("password", "") or "").strip()
        adjust = int(item.get("perm_adjust", item.get("perm_change", 0)) or 0)
        vals = {
            "perm_req":      int(item.get("perm_req",      0) or 0),
            "perm_gen":      int(item.get("perm_gen",      0) or 0),
            "perm_create":   adjust,
            "perm_change":   adjust,
            "perm_approve":  int(item.get("perm_approve",  0) or 0),
            "perm_publish":  int(item.get("perm_publish",  0) or 0),
            "perm_staff":    int(item.get("perm_staff",    0) or 0),
            "perm_facility": int(item.get("perm_facility", 0) or 0),
            "perm_settings": int(item.get("perm_settings", 0) or 0),
        }

        existing = qdb("SELECT id,password_hash FROM role_permissions WHERE role=?", (role,), one=True)
        if existing:
            if pw:
                pw_hash = hash_password(pw)
                db.execute("""UPDATE role_permissions SET password_hash=?,perm_req=?,perm_gen=?,
                            perm_create=?,perm_change=?,perm_approve=?,perm_publish=?,
                            perm_staff=?,perm_facility=?,perm_settings=?,
                            updated_at=datetime('now','localtime') WHERE role=?""",
                           (pw_hash, vals["perm_req"], vals["perm_gen"], vals["perm_create"],
                            vals["perm_change"], vals["perm_approve"], vals["perm_publish"],
                            vals["perm_staff"], vals["perm_facility"], vals["perm_settings"], role))
                db.execute("UPDATE staff SET password_hash=? WHERE system_role=?", (pw_hash, role))
            else:
                db.execute("""UPDATE role_permissions SET perm_req=?,perm_gen=?,perm_create=?,
                            perm_change=?,perm_approve=?,perm_publish=?,
                            perm_staff=?,perm_facility=?,perm_settings=?,
                            updated_at=datetime('now','localtime') WHERE role=?""",
                           (vals["perm_req"], vals["perm_gen"], vals["perm_create"],
                            vals["perm_change"], vals["perm_approve"], vals["perm_publish"],
                            vals["perm_staff"], vals["perm_facility"], vals["perm_settings"], role))
        else:
            pw_hash = hash_password(pw) if pw else None
            db.execute("""INSERT INTO role_permissions
                        (role,password_hash,perm_req,perm_gen,perm_create,perm_change,
                         perm_approve,perm_publish,perm_staff,perm_facility,perm_settings)
                        VALUES(?,?,?,?,?,?,?,?,?,?,?)""",
                       (role, pw_hash, vals["perm_req"], vals["perm_gen"], vals["perm_create"],
                        vals["perm_change"], vals["perm_approve"], vals["perm_publish"],
                        vals["perm_staff"], vals["perm_facility"], vals["perm_settings"]))
            if pw_hash:
                db.execute("UPDATE staff SET password_hash=? WHERE system_role=?", (pw_hash, role))
        saved_roles.append(role)

    db.commit()
    return jsonify({"ok":True,"roles":saved_roles})

@app.route("/api/permissions/<path:role>", methods=["DELETE"])
def api_delete_permission(role):
    """未使用のカスタムロールを削除する。"""
    u = require_login("admin")
    if not u:
        return jsonify({"ok":False,"error":"権限がありません"}), 403
    if role == "admin":
        return jsonify({"ok":False,"error":"adminは削除できません"}), 400
    assigned = qdb("SELECT COUNT(*) cnt FROM staff WHERE system_role=?", (role,), one=True)
    if assigned and assigned["cnt"]:
        return jsonify({"ok":False,"error":"このロールは職員に割り当てられているため削除できません"}), 400
    xdb("DELETE FROM role_permissions WHERE role=?", (role,))
    return jsonify({"ok":True})

@app.route("/api/auth/set_password", methods=["POST"])
def api_set_password():
    """管理者がスタッフのlogin_id・password・system_roleを設定"""
    u = require_login("admin")
    if not u: return jsonify({"ok":False,"error":"権限がありません"}), 403
    d = request.get_json()
    sid  = d.get("staff_id")
    lid  = d.get("login_id","").strip() or None
    pw   = d.get("password") or None
    role = d.get("system_role","staff")
    if not lid:
        return jsonify({"ok":False,"error":"ログインIDは必須です"}), 400
    if pw:
        xdb("UPDATE staff SET login_id=?,password_hash=?,system_role=? WHERE id=?",
            (lid, hash_password(pw), role, sid))
    else:
        # パスワードなしの場合はlogin_idとroleのみ更新
        xdb("UPDATE staff SET login_id=?,system_role=? WHERE id=?",
            (lid, role, sid))
    get_db().commit()
    return jsonify({"ok":True})

from flask import redirect, url_for

@app.route("/login")
def login_page(): return render_template("login.html")

def _is_mobile_request():
    """User-AgentからPC/スマホを判定する。スマホ・タブレットならTrue。"""
    ua = (request.headers.get("User-Agent") or "").lower()
    return any(k in ua for k in ("iphone", "ipad", "ipod", "android", "mobile"))

def _staff_landing_path():
    """staffロールのログイン後・トップ着地先を、PC/スマホで出し分けて返す。"""
    return "/my_shift" if _is_mobile_request() else "/view_shift"

@app.route("/")
def menu():
    u = get_session_staff()
    if not u: return redirect("/login")
    # staffロールは「シフト表」（本人が見るだけの画面）に着地させる。
    # PC/スマホの振り分けはUser-Agentで行い、login.html側の決め打ちは廃止。
    # 希望日入力へはシフト表画面内のボタンから遷移する。
    if u.get("system_role","staff") == "staff":
        return redirect(_staff_landing_path())
    return render_template("menu.html")

@app.route("/my_shift")
def my_shift_page():
    u = get_session_staff()
    if not u: return redirect("/login")
    return render_template("my_shift.html")

@app.route("/view_shift")
def view_shift_page():
    if not get_session_staff(): return redirect("/login")
    return render_template("view_shift.html")

@app.route("/request")
def request_page():
    """希望日入力（本人のみ）。PC/スマホで別テンプレートを出し分ける。"""
    if not get_session_staff(): return redirect("/login")
    return render_template("request_mobile.html" if _is_mobile_request() else "request_pc.html")

@app.route("/permissions")
def permissions_page():
    u = require_login("admin")
    if not u: return redirect("/login")
    return render_template("permissions.html")

@app.route("/facility")
def facility_page():
    u = require_login("admin")
    if not u: return redirect("/login")
    return render_template("facility.html")

@app.route("/staff")
def staff_page():
    # adminが未設定の場合は初回セットアップとして許可
    admin_exists = qdb("SELECT id FROM staff WHERE system_role='admin' AND login_id IS NOT NULL LIMIT 1")
    if admin_exists:
        u = require_login("admin")
        if not u: return redirect("/login")
    return render_template("staff.html")

def _can_open_shift_page(user, step):
    """シフト画面の工程別アクセス可否を判定する。未ログインと権限不足を区別する。"""
    role = (user.get("system_role") or "staff").strip()
    if role == "admin":
        return True

    perm = qdb("SELECT * FROM role_permissions WHERE role=?", (role,), one=True)
    if perm:
        step = str(step or "0")
        if step == "1":
            return bool(perm["perm_req"])
        if step == "3":
            return bool(perm["perm_gen"] or perm["perm_create"] or perm["perm_change"])
        if step == "4":
            return bool(perm["perm_approve"] or perm["perm_publish"])
        return bool(perm["perm_req"] or perm["perm_create"] or perm["perm_change"])

    # 旧来ロールとの互換性。leader以上はシフト画面を利用可能。
    return role in ("leader", "admin")

@app.route("/shift")
def shift_page():
    u = get_session_staff()
    if not u:
        return redirect("/login")
    # 職員登録・事業所登録などの共通ナビから来た場合は、
    # /shift?facility=...&step=... の画面そのものを表示する。
    # ここでメニューへ戻すと、ボタン押下時に目的画面が開けないため、
    # 画面表示は許可し、各操作権限は画面内のボタン/API側で制御する。
    return render_template("index.html")


# ── API: メインメニュー右画面メモ ─────────────────────
@app.route("/api/dashboard_note/<int:year>/<int:month>", methods=["GET"])
def api_get_dashboard_note(year, month):
    """事業所＋対象年月ごとの作成メモ・申し送りを取得"""
    u = get_session_staff()
    if not u:
        return jsonify({"ok": False, "error": "未ログインです"}), 401

    fac_code = (request.args.get("facility_code") or "").strip()
    if not fac_code:
        return jsonify({"ok": False, "error": "facility_code が必要です"}), 400

    fac = qdb("SELECT id, code FROM facilities WHERE code=?", (fac_code,), one=True)
    if not fac:
        return jsonify({"ok": False, "error": "事業所が見つかりません"}), 404

    row = qdb(
        "SELECT note, updated_at FROM shift_dashboard_notes "
        "WHERE facility_id=? AND year=? AND month=?",
        (fac["id"], year, month), one=True)

    return jsonify({
        "ok": True,
        "facility_code": fac["code"],
        "year": year,
        "month": month,
        "note": row["note"] if row else "",
        "updated_at": row["updated_at"] if row else None
    })


@app.route("/api/dashboard_note", methods=["POST"])
def api_save_dashboard_note():
    """事業所＋対象年月ごとの作成メモ・申し送りを保存"""
    u = get_session_staff()
    if not u:
        return jsonify({"ok": False, "error": "未ログインです"}), 401

    d = request.get_json(silent=True) or {}
    fac_code = str(d.get("facility_code", "") or "").strip()
    year = int(d.get("year") or 0)
    month = int(d.get("month") or 0)
    note = str(d.get("note", "") or "")

    if not fac_code or not year or not month:
        return jsonify({"ok": False, "error": "facility_code / year / month が必要です"}), 400

    fac = qdb("SELECT id, code FROM facilities WHERE code=?", (fac_code,), one=True)
    if not fac:
        return jsonify({"ok": False, "error": "事業所が見つかりません"}), 404

    xdb(
        "INSERT INTO shift_dashboard_notes(facility_id, year, month, note, updated_by, updated_at) "
        "VALUES(?,?,?,?,?,datetime('now','localtime')) "
        "ON CONFLICT(facility_id, year, month) DO UPDATE SET "
        "note=excluded.note, updated_by=excluded.updated_by, updated_at=datetime('now','localtime')",
        (fac["id"], year, month, note, u["id"])
    )
    return jsonify({"ok": True, "facility_code": fac["code"], "year": year, "month": month})

# ── API: 全バージョン一覧（メインメニュー用） ──────────────────────────
@app.route("/api/shifts/versions/all")
def api_versions_all():
    u = get_session_staff()
    if not u:
        return jsonify([])
    rows = qdb(
        "SELECT sv.* FROM shift_versions sv "
        "INNER JOIN ("
        "  SELECT facility_id, year, month, MAX(version_no) AS max_ver "
        "  FROM shift_versions GROUP BY facility_id, year, month"
        ") latest ON sv.facility_id=latest.facility_id "
        "  AND sv.year=latest.year AND sv.month=latest.month "
        "  AND sv.version_no=latest.max_ver "
        "ORDER BY sv.facility_id, sv.year DESC, sv.month DESC"
    )
    return jsonify([dict(r) for r in rows])

# ── API: シフト充足率集計（メインメニュー用） ──────────────────────────
@app.route("/api/shift_summary", methods=["GET"])
def api_shift_summary():
    u = get_session_staff()
    if not u:
        return jsonify({"ok": False, "error": "未ログインです"}), 401
    fac_code = (request.args.get("facility_code") or "").strip()
    year  = int(request.args.get("year",  0) or 0)
    month = int(request.args.get("month", 0) or 0)
    if not fac_code or not year or not month:
        return jsonify({"ok": False, "error": "facility_code/year/month が必要です"}), 400
    fac = qdb("SELECT id FROM facilities WHERE code=?", (fac_code,), one=True)
    if not fac:
        return jsonify({"ok": False, "error": "事業所が見つかりません"}), 404
    fid = fac["id"]
    ver = qdb(
        "SELECT id FROM shift_versions "
        "WHERE facility_id=? AND year=? AND month=? "
        "ORDER BY version_no DESC LIMIT 1",
        (fid, year, month), one=True)
    if not ver:
        return jsonify({"ok": True, "summary": [], "unit_summary": [], "fac_summary": None})
    vid = ver["id"]
    import calendar
    from collections import defaultdict
    _, last_day = calendar.monthrange(year, month)
    units = qdb("SELECT id, name FROM units WHERE facility_id=? ORDER BY unit_no", (fid,))
    staff_rows = qdb(
        "SELECT id, name, monthly_limit, unit_id FROM staff "
        "WHERE facility_id=? AND is_active=1 ORDER BY unit_id, id", (fid,))
    entries = qdb(
        "SELECT staff_id, shift_type, is_cross FROM shift_entries WHERE version_id=?", (vid,))
    s_counts = defaultdict(lambda: {"早":0,"遅":0,"夜":0,"日":0,"明":0,"help":0,"total":0})
    for e in entries:
        st = e["shift_type"]
        sid = e["staff_id"]
        if st in ("早","遅","夜","日","明"):
            s_counts[sid]["total"] += 1
            s_counts[sid][st] = s_counts[sid].get(st, 0) + 1
            if e["is_cross"]:
                s_counts[sid]["help"] += 1
    summary = []
    for s in staff_rows:
        sid = s["id"]
        c = s_counts[sid]
        contract = s["monthly_limit"] or 0
        total = c["total"]
        rate = round(total / contract * 100) if contract else None
        summary.append({
            "name": s["name"], "unit_id": s["unit_id"],
            "early": c["早"], "late": c["遅"], "night": c["夜"],
            "ake": c["明"], "day": c["日"],
            "help": c["help"], "total": total,
            "contract": contract, "rate": rate,
        })
    unit_summary = []
    fac_contract = fac_total = fac_help = 0
    for u_row in units:
        uid = u_row["id"]
        unit_staff = [s for s in summary if s["unit_id"] == uid]
        contract = sum(s["contract"] for s in unit_staff)
        total    = sum(s["total"]    for s in unit_staff)
        help_cnt = sum(s["help"]     for s in unit_staff)
        early    = sum(s["early"]    for s in unit_staff)
        late     = sum(s["late"]     for s in unit_staff)
        night    = sum(s["night"]    for s in unit_staff)
        ake      = sum(s["ake"]      for s in unit_staff)
        day      = sum(s["day"]      for s in unit_staff)
        rate     = round(total / contract * 100) if contract else None
        surplus  = contract - total
        unit_summary.append({
            "unit_name": u_row["name"],
            "contract": contract, "total": total,
            "early": early, "late": late, "night": night,
            "ake": ake, "day": day,
            "help": help_cnt, "rate": rate, "surplus": surplus,
        })
        fac_contract += contract
        fac_total    += total
        fac_help     += help_cnt
    fac_summary = {
        "contract": fac_contract, "total": fac_total, "help": fac_help,
        "rate": round(fac_total / fac_contract * 100) if fac_contract else None,
        "surplus": fac_contract - fac_total,
    }
    return jsonify({"ok": True, "summary": summary, "unit_summary": unit_summary, "fac_summary": fac_summary})

# ── API: 勤務区分マスタ ──────────────────────────
@app.route("/api/facilities/<int:fid>/work_types", methods=["GET"])
def api_get_work_types(fid):
    u = get_session_staff()
    if not u:
        return jsonify({"ok": False, "error": "未ログインです"}), 401
    rows = qdb(
        "SELECT * FROM shift_work_types WHERE facility_id=? ORDER BY sort_order, id", (fid,))
    return jsonify([dict(r) for r in rows])

@app.route("/api/facilities/<int:fid>/work_types", methods=["POST"])
def api_save_work_types(fid):
    # 修正(2026-07-09): 他の事業所設定系（/api/facilities本体・ユニット等）と
    # 揃え、perm_facility（実運用上はadmin相当）必須にする。
    u, err = require_perm("perm_facility")
    if err: return err
    data = request.get_json() or {}
    items = data.get("items", [])
    db = get_db()
    db.execute("DELETE FROM shift_work_types WHERE facility_id=?", (fid,))
    for i, item in enumerate(items):
        db.execute(
            "INSERT INTO shift_work_types(facility_id,shift_type,label,start_time,end_time,break_minutes,sort_order) "
            "VALUES(?,?,?,?,?,?,?)",
            (fid, item.get("shift_type",""), item.get("label",""),
             item.get("start_time",""), item.get("end_time",""),
             int(item.get("break_minutes", 60)), i))
    db.commit()
    return jsonify({"ok": True})

# ── API: 事業所 ──────────────────────────
@app.route("/api/facilities", methods=["GET"])
def api_get_facilities():
    u, err = require_any_login()
    if err: return err
    rows = qdb("SELECT * FROM facilities ORDER BY sort_order,id")
    result = []
    for f in rows:
        units = qdb("SELECT * FROM units WHERE facility_id=? ORDER BY unit_no",(f["id"],))
        result.append({**dict(f), "units":[dict(u) for u in units]})
    return jsonify(result)

@app.route("/api/facilities", methods=["POST"])
def api_add_facility():
    u, err = require_perm("perm_facility")
    if err: return err
    d = request.get_json()
    if not d.get("code") or not d.get("name"):
        return jsonify({"error":"code と name は必須です"}), 400
    fid = xdb("""INSERT INTO facilities
        (code,name,type,color,auto_shift,zip_code,address,tel,fax,email,manager,note,sort_order)
        VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)""",
        (d["code"],d["name"],d.get("type","GH"),d.get("color","#5b8fff"),
         int(d.get("auto_shift",1)),d.get("zip_code",""),d.get("address",""),
         d.get("tel",""),d.get("fax",""),d.get("email",""),d.get("manager",""),
         d.get("note",""),d.get("sort_order",99)))
    # ユニット一括登録
    for u2 in d.get("units",[]):
        xdb("INSERT INTO units(facility_id,unit_no,name,residents,color) VALUES(?,?,?,?,?)",
            (fid,u2["unit_no"],u2["name"],u2.get("residents",0),u2.get("color","#5b8fff")))
    return jsonify({"ok":True,"id":fid})

@app.route("/api/facilities/<int:fid>", methods=["PUT"])
def api_update_facility(fid):
    u, err = require_perm("perm_facility")
    if err: return err
    d = request.get_json()
    xdb("""UPDATE facilities SET code=?,name=?,type=?,color=?,auto_shift=?,
           zip_code=?,address=?,tel=?,fax=?,email=?,manager=?,note=?,sort_order=?
           WHERE id=?""",
        (d["code"],d["name"],d.get("type","GH"),d.get("color","#5b8fff"),
         int(d.get("auto_shift",1)),d.get("zip_code",""),d.get("address",""),
         d.get("tel",""),d.get("fax",""),d.get("email",""),d.get("manager",""),
         d.get("note",""),d.get("sort_order",99),fid))
    return jsonify({"ok":True})

@app.route("/api/facilities/<int:fid>/units", methods=["POST"])
def api_add_unit(fid):
    u, err = require_perm("perm_facility")
    if err: return err
    d = request.get_json()
    uid = xdb("INSERT INTO units(facility_id,unit_no,name,residents,color) VALUES(?,?,?,?,?)",
              (fid,d["unit_no"],d["name"],d.get("residents",0),d.get("color","#5b8fff")))
    return jsonify({"ok":True,"id":uid})

@app.route("/api/facilities/<int:fid>/units/<int:uid>", methods=["PUT"])
def api_update_unit(fid,uid):
    u, err = require_perm("perm_facility")
    if err: return err
    d = request.get_json()
    xdb("UPDATE units SET name=?,residents=?,color=? WHERE id=? AND facility_id=?",
        (d["name"],d.get("residents",0),d.get("color","#5b8fff"),uid,fid))
    return jsonify({"ok":True})

@app.route("/api/facilities/<int:fid>/units/<int:uid>", methods=["DELETE"])
def api_delete_unit(fid,uid):
    u, err = require_perm("perm_facility")
    if err: return err
    xdb("DELETE FROM units WHERE id=? AND facility_id=?",(uid,fid))
    return jsonify({"ok":True})

# ── API: スタッフ ────────────────────────
@app.route("/api/staff", methods=["GET"])
def api_get_staff():
    u, err = require_any_login()
    if err: return err
    show_all = request.args.get("all","0")=="1"
    fac = request.args.get("facility")
    sql = """SELECT s.*,f.code fac_code,f.name fac_name,f.type fac_type,f.color fac_color,
                    u.name unit_name,u.unit_no,u.color unit_color
             FROM staff s
             JOIN facilities f ON s.facility_id=f.id
             JOIN units u ON s.unit_id=u.id WHERE 1=1"""
    args = []
    if not show_all:
        sql += " AND s.is_active=1"
    if fac:
        sql += " AND f.code=?"
        args.append(fac)
    sql += " ORDER BY f.sort_order,f.id,u.unit_no,s.id"
    rows = qdb(sql, args)
    result = []
    for r in rows:
        skills = qdb("""SELECT ss.unit_id,u.name uname,u.unit_no,f.code fcode,ss.level
                        FROM staff_skills ss
                        JOIN units u ON ss.unit_id=u.id
                        JOIN facilities f ON u.facility_id=f.id
                        WHERE ss.staff_id=?""",(r["id"],))
        result.append({**dict(r),"skills":[dict(s) for s in skills]})
    return jsonify(result)

@app.route("/api/staff", methods=["POST"])
def api_add_staff():
    u, err = require_perm("perm_staff")
    if err: return err
    d = request.get_json()
    system_role = str(d.get("system_role", "staff") or "staff").strip()
    if system_role != "staff" and (u.get("system_role") or "staff") != "admin":
        # staff以外のロール付与（leader/admin等の昇格）はadminのみ許可。
        system_role = "staff"
    role_row = qdb("SELECT password_hash FROM role_permissions WHERE role=?", (system_role,), one=True)
    role_pw_hash = role_row["password_hash"] if role_row else None
    sid = xdb("""INSERT INTO staff(facility_id,unit_id,name,role,employment_type,
               join_date,monthly_limit,night_target,
               can_day,can_early,can_late,can_night,can_night_only,is_help,is_approver,note,
               system_role,password_hash)
               VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
              (d["facility_id"],d["unit_id"],d["name"],d.get("role",""),
               d.get("employment_type","full_time"),
               d.get("join_date"),d.get("monthly_limit",21),
               int(d.get("night_target",0)),
               int(d.get("can_day",1)),int(d.get("can_early",1)),
               int(d.get("can_late",1)),int(d.get("can_night",0)),
               int(d.get("can_night_only",0)),int(d.get("is_help",0)),
               int(d.get("is_approver",0)),
               d.get("note",""), system_role, role_pw_hash))
    xdb("INSERT OR IGNORE INTO staff_skills(staff_id,unit_id,level) VALUES(?,?,'ok')",(sid,d["unit_id"]))
    return jsonify({"ok":True,"id":sid})

@app.route("/api/staff/<int:sid>", methods=["PUT"])
def api_update_staff(sid):
    u, err = require_perm("perm_staff")
    if err: return err
    is_admin = (u.get("system_role") or "staff") == "admin"
    d = request.get_json()
    leave = d.get("leave_date") or None
    xdb("""UPDATE staff SET facility_id=?,unit_id=?,name=?,role=?,employment_type=?,
           join_date=?,leave_date=?,monthly_limit=?,night_target=?,
           can_day=?,can_early=?,can_late=?,
           can_night=?,can_night_only=?,is_help=?,is_approver=?,is_active=?,note=? WHERE id=?""",
        (d["facility_id"],d["unit_id"],d["name"],d.get("role",""),
         d.get("employment_type","full_time"),
         d.get("join_date"),leave,d.get("monthly_limit",21),
         int(d.get("night_target",0)),
         int(d.get("can_day",1)),int(d.get("can_early",1)),
         int(d.get("can_late",1)),int(d.get("can_night",0)),
         int(d.get("can_night_only",0)),int(d.get("is_help",0)),
         int(d.get("is_approver",0)),
         0 if leave else 1,d.get("note",""),sid))
    # ログイン設定（入力があれば更新）。ロール・パスワードの変更は権限昇格に直結するためadmin限定。
    if is_admin:
        login_id    = d.get("login_id","").strip() or None
        system_role = d.get("system_role","staff")
        pw = d.get("password","").strip()
        if pw:
            xdb("UPDATE staff SET login_id=?,password_hash=?,system_role=? WHERE id=?",
                (login_id, hash_password(pw), system_role, sid))
        else:
            xdb("UPDATE staff SET login_id=?,system_role=? WHERE id=?",
                (login_id, system_role, sid))
    get_db().commit()
    return jsonify({"ok":True})

@app.route("/api/staff/<int:sid>", methods=["DELETE"])
def api_delete_staff(sid):
    u, err = require_perm("perm_staff")
    if err: return err
    xdb("UPDATE staff SET is_active=0,leave_date=? WHERE id=?",(date.today().isoformat(),sid))
    return jsonify({"ok":True})

# ── API: スキル ──────────────────────────
@app.route("/api/skills/<int:sid>")
def api_get_skills(sid):
    u, err = require_any_login()
    if err: return err
    rows = qdb("""SELECT ss.*,u.name uname,u.unit_no,f.code fcode,f.name fname
                  FROM staff_skills ss
                  JOIN units u ON ss.unit_id=u.id
                  JOIN facilities f ON u.facility_id=f.id
                  WHERE ss.staff_id=?""",(sid,))
    return jsonify([dict(r) for r in rows])

@app.route("/api/skills", methods=["POST"])
def api_upsert_skill():
    u, err = require_perm("perm_staff")
    if err: return err
    d = request.get_json()
    xdb("""INSERT INTO staff_skills(staff_id,unit_id,level) VALUES(?,?,?)
           ON CONFLICT(staff_id,unit_id) DO UPDATE SET level=excluded.level""",
        (d["staff_id"],d["unit_id"],d.get("level","ok")))
    return jsonify({"ok":True})

@app.route("/api/skills/<int:sid>/<int:uid>", methods=["DELETE"])
def api_delete_skill(sid,uid):
    u, err = require_perm("perm_staff")
    if err: return err
    xdb("DELETE FROM staff_skills WHERE staff_id=? AND unit_id=?",(sid,uid))
    return jsonify({"ok":True})

# ── API: 曜日固定設定 ────────────────────
@app.route("/api/fixed_days/<int:sid>", methods=["GET"])
def api_get_fixed_days(sid):
    u, err = require_any_login()
    if err: return err
    rows = qdb(
        "SELECT id, dow, shift_type FROM staff_fixed_days WHERE staff_id=? ORDER BY dow",
        (sid,))
    return jsonify([dict(r) for r in rows])

@app.route("/api/fixed_days/<int:sid>", methods=["POST"])
def api_save_fixed_days(sid):
    """曜日固定設定を全件置換"""
    u, err = require_perm("perm_staff")
    if err: return err
    data = request.get_json()
    db = get_db()
    db.execute("DELETE FROM staff_fixed_days WHERE staff_id=?", (sid,))
    for item in data:
        db.execute(
            "INSERT INTO staff_fixed_days(staff_id,dow,shift_type) VALUES(?,?,?)",
            (sid, item["dow"], item["shift_type"]))
    db.commit()
    return jsonify({"ok": True})


# ── API: 希望入力 ────────────────────────
@app.route("/api/requests/<int:year>/<int:month>")
def api_get_requests(year,month):
    u, err = require_any_login()
    if err: return err
    role = (u.get("system_role") or "staff").strip()
    sid = request.args.get("staff_id")
    if role == "staff":
        sid = str(u["id"])
    if sid:
        rows = qdb("SELECT * FROM requests WHERE staff_id=? AND year=? AND month=? ORDER BY day",(sid,year,month))
    else:
        rows = qdb("SELECT * FROM requests WHERE year=? AND month=? ORDER BY staff_id,day",(year,month))
    return jsonify([dict(r) for r in rows])

@app.route("/api/requests", methods=["POST"])
def api_save_request():
    u, err = require_any_login()
    if err: return err
    d = request.get_json()
    role = (u.get("system_role") or "staff").strip()
    if role == "staff" and int(d.get("staff_id", -1)) != u["id"]:
        return jsonify({"ok": False, "error": "自分以外の希望は入力できません"}), 403
    xdb("""INSERT INTO requests(staff_id,year,month,day,req_type,priority,note)
           VALUES(?,?,?,?,?,?,?)
           ON CONFLICT(staff_id,year,month,day)
           DO UPDATE SET req_type=excluded.req_type,priority=excluded.priority,note=excluded.note""",
        (d["staff_id"],d["year"],d["month"],d["day"],d["req_type"],d.get("priority","prefer"),d.get("note","")))
    if d["req_type"]=="night":
        last = calendar.monthrange(d["year"],d["month"])[1]
        if d["day"] < last:
            xdb("""INSERT INTO requests(staff_id,year,month,day,req_type,priority,note)
                   VALUES(?,?,?,?,'ake','must','夜勤明け（自動）')
                   ON CONFLICT(staff_id,year,month,day)
                   DO UPDATE SET req_type='ake',priority='must'""",
                (d["staff_id"],d["year"],d["month"],d["day"]+1))
    return jsonify({"ok":True})

@app.route("/api/requests/<int:sid>/<int:year>/<int:month>/<int:day>", methods=["DELETE"])
def api_delete_request(sid,year,month,day):
    u, err = require_any_login()
    if err: return err
    role = (u.get("system_role") or "staff").strip()
    if role == "staff" and sid != u["id"]:
        return jsonify({"ok": False, "error": "自分以外の希望は削除できません"}), 403
    req = qdb("SELECT req_type FROM requests WHERE staff_id=? AND year=? AND month=? AND day=?",(sid,year,month,day),one=True)
    if req and req["req_type"]=="night":
        xdb("DELETE FROM requests WHERE staff_id=? AND year=? AND month=? AND day=?",(sid,year,month,day+1))
    xdb("DELETE FROM requests WHERE staff_id=? AND year=? AND month=? AND day=?",(sid,year,month,day))
    return jsonify({"ok":True})

# ── API: シフト自動生成 ──────────────────
@app.route("/api/shifts/generate", methods=["POST"])
def api_generate():
    u, err = require_perm("perm_gen", "perm_create", "perm_change")
    if err: return err
    d = request.get_json()
    year      = d.get("year",  date.today().year)
    month     = d.get("month", date.today().month)
    fac_code  = d.get("facility_code")
    reason    = d.get("reason", "自動生成")
    preview   = d.get("preview", False)  # Trueの場合DBに書かずエントリを返す
    from_day  = int(d.get("from_day", 1) or 1)  # 基準日（この日以降を再生成）
    last_day  = calendar.monthrange(year, month)[1]

    if fac_code == "GH_ALL":
        fac_ids = [r["id"] for r in qdb("SELECT id FROM facilities WHERE type='GH'")]
    elif fac_code:
        f = qdb("SELECT id FROM facilities WHERE code=?", (fac_code,), one=True)
        fac_ids = [f["id"]] if f else []
    else:
        fac_ids = [r["id"] for r in qdb("SELECT id FROM facilities WHERE auto_shift=1")]

    results = []
    for fid in fac_ids:
        fac = qdb("SELECT code,type FROM facilities WHERE id=?", (fid,), one=True)

        if preview:
            # プレビューモード。
            # 修正(2026-06-20): 計算結果をDBに一切書かないままだと、
            # 自動生成を押しただけでは何も保存されず、後続のentries保存API
            # 経由でしか永続化されない＝change_reason/eval_jsonが更新されず
            # 古い情報のまま残る問題があったため、その年月の現在のドラフト
            # バージョンへ「計算結果を自動的に反映」するよう変更する。
            # （フロント側のレスポンス形は変えない。version_idはこれまで通りNone）
            fac_type = fac["type"]

            existing = qdb(
                """SELECT id FROM shift_versions
                   WHERE facility_id=? AND year=? AND month=?
                   ORDER BY version_no DESC LIMIT 1""",
                (fid, year, month), one=True)
            if existing:
                vid = existing["id"]
            else:
                vid = xdb(
                    "INSERT INTO shift_versions(facility_id,year,month,version_no,status,change_reason) "
                    "VALUES(?,?,?,?,'draft',?)",
                    (fid, year, month, 1, reason))

            if fac_type == "GH":
                # from_day指定時：基準日未満の既存シフトをlock_prefixに変換
                lock_prefix = {}
                if from_day > 1:
                    existing_entries = qdb(
                        "SELECT staff_id, date, shift_type, unit_id FROM shift_entries WHERE version_id=?",
                        (vid,))
                    for e in existing_entries:
                        day = int(e["date"].split("-")[2]) if "-" in str(e["date"]) else int(e["date"])
                        if day < from_day:
                            lock_prefix[(e["staff_id"], day)] = (e["shift_type"], e["unit_id"])
                entries, evaluation = _generate_gh(fid, vid, year, month, last_day,
                                                   lock_prefix=lock_prefix if lock_prefix else None)
            else:
                entries, evaluation = _generate_non_gh(fid, vid, year, month, last_day)

            db = get_db()
            db.execute("DELETE FROM shift_entries WHERE version_id=?", (vid,))
            db.executemany(
                "INSERT INTO shift_entries(version_id,staff_id,unit_id,date,shift_type,is_cross,from_unit_id) "
                "VALUES(:version_id,:staff_id,:unit_id,:date,:shift_type,:is_cross,:from_unit_id)",
                entries)
            if evaluation is not None:
                db.execute("UPDATE shift_versions SET eval_json=? WHERE id=?",
                           (json.dumps(evaluation, ensure_ascii=False), vid))
            # 通常の(欠勤対応ではない)全体再生成なので、累積差分ハイライトは
            # 無関係になる。古いdiff_cells_json/based_on_version_idが残っていると
            # 次の欠勤対応で無関係な差分が紛れ込むため、ここでリセットする
            # （2026-06-21追加）。
            # 修正(2026-07-08): 指定日更新（from_day>1）の場合、指定日より前は
            # 今回一切変更していないため、既存の変更マークのうち指定日より前の
            # 分はそのまま残す。指定日以降の分だけ「今回の再生成で無関係になった」
            # ものとして捨てる。全体再生成(from_day<=1)の場合は従来通り全消去。
            kept_diff_cells = []
            if from_day > 1:
                prev_ver = qdb("SELECT diff_cells_json FROM shift_versions WHERE id=?",
                               (vid,), one=True)
                if prev_ver and prev_ver["diff_cells_json"]:
                    try:
                        for c in json.loads(prev_ver["diff_cells_json"]):
                            try:
                                d = int(str(c.get("date", "")).split("-")[2])
                            except (IndexError, ValueError):
                                continue
                            if d < from_day:
                                kept_diff_cells.append(c)
                    except (TypeError, ValueError):
                        pass
            if kept_diff_cells:
                db.execute("UPDATE shift_versions SET diff_cells_json=? WHERE id=?",
                           (json.dumps(kept_diff_cells, ensure_ascii=False), vid))
            else:
                db.execute("UPDATE shift_versions SET diff_cells_json=NULL, based_on_version_id=NULL WHERE id=?",
                           (vid,))
            db.commit()

            # 実績（早退・遅刻等）マーク。指定日より前は今回変更していないため、
            # 画面側で自動生成直後も実績マークを再表示できるよう、そのまま
            # 一緒に返す（curGeneratedVidが未確定＝未保存の間も表示するため）。
            actual_rows = qdb(
                "SELECT staff_id, date, actual_type, time_from, time_to, work_hours, note "
                "FROM shift_actual WHERE version_id=?", (vid,))

            results.append({
                "facility_code": fac["code"],
                "version_id": None,
                "preview_vid": vid,
                "entries": [dict(e) for e in entries],
                "evaluation": evaluation,
                "actual": [dict(r) for r in actual_rows],
            })
        else:
            # 通常モード: DBにバージョン＋エントリを保存
            latest = qdb(
                "SELECT MAX(version_no) v FROM shift_versions WHERE facility_id=? AND year=? AND month=?",
                (fid, year, month), one=True)
            new_ver = (latest["v"] or 0) + 1
            vid = xdb(
                "INSERT INTO shift_versions(facility_id,year,month,version_no,status,change_reason) VALUES(?,?,?,?,'draft',?)",
                (fid, year, month, new_ver, reason))

            fac_type = fac["type"]
            if fac_type == "GH":
                # from_day指定時：基準日未満の既存シフトをlock_prefixに変換
                lock_prefix = {}
                if from_day > 1:
                    existing_entries = qdb(
                        "SELECT staff_id, date, shift_type, unit_id FROM shift_entries WHERE version_id=?",
                        (vid,))
                    for e in existing_entries:
                        day = int(e["date"].split("-")[2]) if "-" in str(e["date"]) else int(e["date"])
                        if day < from_day:
                            lock_prefix[(e["staff_id"], day)] = (e["shift_type"], e["unit_id"])
                entries, evaluation = _generate_gh(fid, vid, year, month, last_day,
                                                   lock_prefix=lock_prefix if lock_prefix else None)
            else:
                entries, evaluation = _generate_non_gh(fid, vid, year, month, last_day)

            db = get_db()
            db.executemany(
                "INSERT INTO shift_entries(version_id,staff_id,unit_id,date,shift_type,is_cross,from_unit_id) "
                "VALUES(:version_id,:staff_id,:unit_id,:date,:shift_type,:is_cross,:from_unit_id)",
                entries)
            if evaluation is not None:
                db.execute("UPDATE shift_versions SET eval_json=? WHERE id=?",
                           (json.dumps(evaluation, ensure_ascii=False), vid))
            db.commit()
            results.append({
                "facility_code": fac["code"],
                "version_id": vid, "version_no": new_ver, "entries": len(entries),
                "evaluation": evaluation,
            })
    return jsonify({"ok": True, "results": results})


def _get_base_entries_map(vid):
    """{(staff_id, 'YYYY-MM-DD'): {'shift_type':..., 'unit_id':...}} 全件を取得する。"""
    rows = qdb("SELECT staff_id, date, shift_type, unit_id FROM shift_entries WHERE version_id=?", (vid,))
    return {(r["staff_id"], r["date"]): {"shift_type": r["shift_type"], "unit_id": r["unit_id"]} for r in rows}

def _emap_state(emap, sid, date_str):
    e = emap.get((sid, date_str))
    return e["shift_type"] if e else "休"

def _date_add(date_str, n):
    y, m, d = (int(x) for x in date_str.split("-"))
    nd = date(y, m, d) + timedelta(days=n)
    return nd.strftime("%Y-%m-%d")

def _find_direct_candidates(emap, fid, abs_sid, date_str, shift_type, unit_id,
                             year, month, day, last_day, exclude_ids=None):
    """その日（夜勤なら翌日も）空いていて、能力的に対応可能な職員を探す（最小置換えの第1段階）。"""
    exclude_ids = exclude_ids or set()
    next_date = _date_add(date_str, 1) if (shift_type == "夜" and day < last_day) else None
    cap_col = {"早": "can_early", "遅": "can_late", "夜": "can_night", "日": "can_day"}[shift_type]
    cand_rows = qdb(
        f"SELECT id, name, employment_type, monthly_limit, unit_id, can_night_only "
        f"FROM staff WHERE facility_id=? AND is_active=1 AND {cap_col}=1 AND id<>?",
        (fid, abs_sid))
    skill_rows = qdb("SELECT staff_id, level FROM staff_skills WHERE unit_id=?", (unit_id,))
    skill_ok_ids = {r["staff_id"] for r in skill_rows if r["level"] in ("ok", "conditional")}

    out = []
    for c in cand_rows:
        if c["id"] in exclude_ids:
            continue
        if shift_type != "夜" and c["can_night_only"]:
            continue
        if c["unit_id"] != unit_id and c["id"] not in skill_ok_ids:
            continue
        if _emap_state(emap, c["id"], date_str) != "休":
            continue
        if next_date and _emap_state(emap, c["id"], next_date) != "休":
            continue
        req = qdb("SELECT req_type FROM requests WHERE staff_id=? AND year=? AND month=? AND day=?",
                  (c["id"], year, month, day), one=True)
        if req and req["req_type"] in ("rest", "hol"):
            continue
        worked = sum(1 for (s2, d2), v in emap.items() if s2 == c["id"] and v["shift_type"] not in ("休", None))
        out.append({
            "staff_id": c["id"], "name": c["name"], "employment_type": c["employment_type"],
            "monthly_limit": c["monthly_limit"], "worked_so_far": worked,
            "same_unit": c["unit_id"] == unit_id,
        })
    out.sort(key=lambda r: (0 if r["same_unit"] else 1, r["worked_so_far"]))
    return out, next_date

def _find_one_hop_candidates(emap, fid, abs_sid, date_str, shift_type, unit_id, year, month, day, last_day):
    """直接の空き要員がいない場合の第2段階。
    『その日すでに別の予定が入っている職員A』の予定を、別の空いている職員Bに
    1件だけ動かせばAを欠勤対応に回せる、という組み合わせを探す。
    2手目（Bの予定をさらに誰かに動かす等）には踏み込まない（玉突きの全体波及を防ぐため）。
    """
    next_date = _date_add(date_str, 1) if (shift_type == "夜" and day < last_day) else None
    cap_col = {"早": "can_early", "遅": "can_late", "夜": "can_night", "日": "can_day"}[shift_type]
    cand_rows = qdb(
        f"SELECT id, name, employment_type, monthly_limit, unit_id, can_night_only "
        f"FROM staff WHERE facility_id=? AND is_active=1 AND {cap_col}=1 AND id<>?",
        (fid, abs_sid))
    skill_rows = qdb("SELECT staff_id, level FROM staff_skills WHERE unit_id=?", (unit_id,))
    skill_ok_ids = {r["staff_id"] for r in skill_rows if r["level"] in ("ok", "conditional")}

    results = []
    for a in cand_rows:
        if shift_type != "夜" and a["can_night_only"]:
            continue
        if a["unit_id"] != unit_id and a["id"] not in skill_ok_ids:
            continue
        a_state = emap.get((a["id"], date_str))
        a_shift = a_state["shift_type"] if a_state else "休"
        if a_shift in ("休", "欠勤", "明", "有給", None):
            continue  # 空き・明け・有給は動かせない
        a_unit = a_state["unit_id"]
        if next_date and _emap_state(emap, a["id"], next_date) != "休":
            continue  # Aを翌日「明け」にできない
        req_a = qdb("SELECT req_type FROM requests WHERE staff_id=? AND year=? AND month=? AND day=?",
                    (a["id"], year, month, day), one=True)
        if req_a:
            continue  # 本人の希望が入っている日は崩さない
        b_options, _ = _find_direct_candidates(
            emap, fid, a["id"], date_str, a_shift, a_unit, year, month, day, last_day,
            exclude_ids={abs_sid})
        if b_options:
            worked_a = sum(1 for (s2, d2), v in emap.items() if s2 == a["id"] and v["shift_type"] not in ("休", None))
            results.append({
                "a_staff_id": a["id"], "a_name": a["name"], "a_current_shift_type": a_shift,
                "a_same_unit": a["unit_id"] == unit_id, "a_worked_so_far": worked_a,
                "b_options": b_options,
            })
    results.sort(key=lambda r: (0 if r["a_same_unit"] else 1, r["a_worked_so_far"]))
    return results


@app.route("/api/shifts/<int:vid>/absence/candidates", methods=["POST"])
def api_absence_candidates(vid):
    """欠勤発生時の対応プランを日付範囲で検索する。

    複数人欠勤対応に対応するため、フロントから現在の下書き状態(draft_entries)を
    受け取り、DBの保存済みデータではなくその下書き状態をベースに候補を探す。
    これにより「保存せずに複数回の欠勤対応を繰り返す」ことができる。
    draft_entriesが省略された場合はDBの保存済みデータを使う（後方互換）。

    各日について独立に、
      1. その日（夜勤なら翌日も）空いている職員＝直接候補
      2. いなければ、すでに別予定がある職員Aの予定を1件だけ空いている職員Bへ
         動かせば対応可能、という「1手先までの玉突き」候補
      3. それもなければ「代行者なし」（人員不足のまま欠勤のみ確定する想定）
    を返す。月全体やそれ以降の日への影響は一切計算しない（最小置換え方針）。
    """
    u, err = require_perm("perm_gen", "perm_create", "perm_change")
    if err: return err
    d = request.get_json() or {}
    sid = d.get("staff_id")
    start_date = d.get("date") or d.get("start_date")
    end_date = d.get("end_date") or start_date
    draft_entries = d.get("draft_entries")  # フロントの下書き状態（複数欠勤対応用）
    if not sid or not start_date:
        return jsonify({"error": "staff_id, dateが必要です"}), 400

    ver = qdb("SELECT * FROM shift_versions WHERE id=?", (vid,), one=True)
    if not ver:
        return jsonify({"error": "バージョンが見つかりません"}), 404

    year, month = ver["year"], ver["month"]
    last_day = calendar.monthrange(year, month)[1]
    sy, sm, sd = (int(x) for x in start_date.split("-"))
    ey, em_, ed = (int(x) for x in end_date.split("-"))
    if (sy, sm) != (year, month) or (ey, em_) != (year, month):
        return jsonify({"error": "対象月と異なる日付が指定されています"}), 400
    if sd > ed:
        return jsonify({"error": "終了日は開始日以降にしてください"}), 400

    # draft_entriesが渡された場合はそれをemapとして使う。
    # {(staff_id, date_str): {shift_type, unit_id}} 形式に変換する。
    if draft_entries:
        emap = {
            (e["staff_id"], e["date"]): {"shift_type": e["shift_type"], "unit_id": e["unit_id"]}
            for e in draft_entries
            if e.get("shift_type")
        }
    else:
        emap = _get_base_entries_map(vid)

    days_out = []
    for day in range(sd, ed + 1):
        date_str = f"{year:04d}-{month:02d}-{day:02d}"
        st = emap.get((sid, date_str))
        if not st or st["shift_type"] not in ("早", "遅", "夜", "日"):
            days_out.append({"date": date_str, "status": "no_action"})
            continue
        shift_type, uid = st["shift_type"], st["unit_id"]
        direct, next_date = _find_direct_candidates(
            emap, ver["facility_id"], sid, date_str, shift_type, uid, year, month, day, last_day)
        one_hop = []
        if direct:
            status = "direct"
        else:
            one_hop = _find_one_hop_candidates(
                emap, ver["facility_id"], sid, date_str, shift_type, uid, year, month, day, last_day)
            status = "one_hop" if one_hop else "no_candidate"
        days_out.append({
            "date": date_str, "shift_type": shift_type, "unit_id": uid, "next_date": next_date,
            "direct_candidates": direct, "one_hop_candidates": one_hop, "status": status,
        })

    return jsonify({"absent_staff_id": sid, "start_date": start_date, "end_date": end_date, "days": days_out})


@app.route("/api/shifts/<int:vid>/absence/confirm", methods=["POST"])
def api_absence_confirm(vid):
    """欠勤対応を確定する（最小置換え方式）。

    公開済みのvidは一切変更しない。新しいshift_versions行を作り、
    指定された日（複数可）について、欠勤者を「欠勤」、選ばれた代行を
    元のシフト種別で確定するだけの「最小限のセル置換え」を行う。
    月全体やそれ以降の日の再生成は一切行わない（影響範囲を最小化する方針、
    2026-06-21 設計変更）。
    """
    u, err = require_perm("perm_gen", "perm_create", "perm_change")
    if err: return err
    d = request.get_json() or {}
    sid = d.get("staff_id")
    days_in = d.get("days") or []
    draft_entries = d.get("draft_entries")  # フロントの下書き状態（複数欠勤対応用）
    if not sid or not days_in:
        return jsonify({"error": "staff_id, daysが必要です"}), 400

    ver = qdb("SELECT * FROM shift_versions WHERE id=?", (vid,), one=True)
    if not ver:
        return jsonify({"error": "バージョンが見つかりません"}), 404

    fid, year, month = ver["facility_id"], ver["year"], ver["month"]
    last_day = calendar.monthrange(year, month)[1]

    # draft_entriesが渡された場合はそれをベースにする（複数欠勤対応）。
    # 1回目の欠勤対応後の下書き状態を2回目以降の計算ベースに使うことで、
    # 保存せずに何回でも欠勤対応を繰り返せる。
    if draft_entries:
        emap = {
            (e["staff_id"], e["date"]): {"shift_type": e["shift_type"], "unit_id": e["unit_id"]}
            for e in draft_entries
            if e.get("shift_type")
        }
    else:
        emap = _get_base_entries_map(vid)
    original_emap = _get_base_entries_map(vid)  # 差分計算は常にDB保存済みベースで行う

    name_cache = {}
    def _name(staff_id):
        if staff_id not in name_cache:
            row = qdb("SELECT name FROM staff WHERE id=?", (staff_id,), one=True)
            name_cache[staff_id] = row["name"] if row else f"id={staff_id}"
        return name_cache[staff_id]

    applied_log = []
    for choice in days_in:
        date_str = choice.get("date")
        mode = choice.get("mode")
        if not date_str or mode in (None, "skip", "no_action"):
            continue
        st = emap.get((sid, date_str))
        if not st or st["shift_type"] not in ("早", "遅", "夜", "日"):
            continue
        shift_type, uid = st["shift_type"], st["unit_id"]
        y_, m_, day_ = (int(x) for x in date_str.split("-"))
        next_date = f"{year:04d}-{month:02d}-{day_+1:02d}" if (shift_type == "夜" and day_ < last_day) else None

        if mode == "direct":
            repl = choice.get("replacement_staff_id")
            if not repl:
                continue
            emap[(sid, date_str)] = {"shift_type": "欠勤", "unit_id": uid}
            emap[(repl, date_str)] = {"shift_type": shift_type, "unit_id": uid}
            if next_date:
                emap[(repl, next_date)] = {"shift_type": "明", "unit_id": uid}
                emap[(sid, next_date)] = {"shift_type": "休", "unit_id": None}
            applied_log.append(f"{date_str} {shift_type} {_name(sid)}が欠勤 → {_name(repl)}が代行")

        elif mode == "one_hop":
            a, b = choice.get("a_staff_id"), choice.get("b_staff_id")
            if not a or not b:
                continue
            a_entry = emap.get((a, date_str))
            if not a_entry or a_entry["shift_type"] in ("休", "欠勤", None):
                continue
            a_shift, a_unit = a_entry["shift_type"], a_entry["unit_id"]
            emap[(sid, date_str)] = {"shift_type": "欠勤", "unit_id": uid}
            emap[(a, date_str)] = {"shift_type": shift_type, "unit_id": uid}
            emap[(b, date_str)] = {"shift_type": a_shift, "unit_id": a_unit}
            if next_date:
                emap[(a, next_date)] = {"shift_type": "明", "unit_id": uid}
                emap[(sid, next_date)] = {"shift_type": "休", "unit_id": None}
            applied_log.append(
                f"{date_str} {shift_type} {_name(sid)}が欠勤 → {_name(a)}が代行"
                f"（{_name(a)}の本来の{a_shift}は{_name(b)}へ）")

        elif mode == "absence_only":
            emap[(sid, date_str)] = {"shift_type": "欠勤", "unit_id": uid}
            if next_date:
                emap[(sid, next_date)] = {"shift_type": "休", "unit_id": None}
            applied_log.append(f"{date_str} {shift_type} {_name(sid)}が欠勤（代行者なし、人員不足のまま残す）")

    if not applied_log:
        return jsonify({"error": "確定できる変更がありませんでした"}), 400

    # ホームユニット(is_cross判定用)
    staff_rows = qdb("SELECT id, unit_id FROM staff WHERE facility_id=?", (fid,))
    home_unit_of = {r["id"]: r["unit_id"] for r in staff_rows}

    # 変更されたセルだけを抽出して返す。DBへの保存はここでは一切行わない。
    # 画面の下書き状態（SHIFTS）に反映するだけにとどめ、実際の保存は
    # 本体の「保存」ボタン（PUT /api/shifts/<vid>/entries）で行う
    # （2026-06-21 設計変更：欠勤対応の確定操作とDB保存を分離）。
    changed_cells = []
    diff_keys = set()
    for key in set(original_emap) | set(emap):
        bv = original_emap.get(key, {"shift_type": "休"})["shift_type"]
        av_entry = emap.get(key, {"shift_type": "休", "unit_id": None})
        av = av_entry["shift_type"]
        if bv != av:
            diff_keys.add(key)
            s2, d2 = key
            uid2 = av_entry["unit_id"] if av_entry["unit_id"] is not None else home_unit_of.get(s2)
            home = home_unit_of.get(s2)
            is_cross = 1 if (uid2 is not None and home is not None and uid2 != home) else 0
            changed_cells.append({
                "staff_id": s2, "date": d2, "shift_type": av, "unit_id": uid2,
                "is_cross": is_cross, "from_unit_id": home if is_cross else None,
            })

    return jsonify({
        "base_version_id": vid, "applied": applied_log,
        "changed_cells": changed_cells,
        "diff_cells": [{"staff_id": k[0], "date": k[1]} for k in diff_keys],
    })


def _generate_non_gh(fid, vid, year, month, last_day):
    """DS/CPC/WGC: 土日=休、平日=日 デフォルト"""
    staff = qdb(
        "SELECT id,can_day,can_early,can_late,can_night,can_night_only,monthly_limit,unit_id "
        "FROM staff WHERE facility_id=? AND is_active=1", (fid,))
    reqs = qdb("SELECT staff_id,day,req_type FROM requests WHERE year=? AND month=?",
               (year, month))
    req_map = defaultdict(dict)
    for r in reqs:
        req_map[r["staff_id"]][r["day"]] = r["req_type"]
    entries = []
    for s in staff:
        sid = s["id"]; work_cnt = 0; day_shifts = {}
        for day in range(1, last_day + 1):
            req = req_map[sid].get(day)
            dow = date(year, month, day).weekday()
            is_we = dow >= 5
            if day > 1 and day_shifts.get(day - 1) == "夜":
                shift = "明"
            elif req == "ake" and day == 1:
                shift = "明"   # 初月稼働時のみ: 前月末夜勤者の1日明け
            elif work_cnt >= s["monthly_limit"] and req not in ("rest","hol","absence","ake"):
                shift = "休"
            elif req in REQ_TYPE_MAP: shift = REQ_TYPE_MAP[req]  # rest→希休, hol→有給, absence→欠勤
            elif req == "night" and s["can_night"]: shift = "夜"
            elif req == "early" and s["can_early"]: shift = "早"
            elif req == "late"  and s["can_late"]:  shift = "遅"
            elif req == "day"   and s["can_day"]:   shift = "日"
            elif is_we: shift = "休"
            else:       shift = "日"
            day_shifts[day] = shift
            if shift not in NON_WORK_SHIFTS: work_cnt += 1  # 明けも就業日数1日カウント
            entries.append({"version_id":vid,"staff_id":sid,"unit_id":s["unit_id"],
                            "date":f"{year}-{month:02d}-{day:02d}",
                            "shift_type":shift,"is_cross":0,"from_unit_id":None})
    # DS/CPC/WGCはシフト種別(早/遅/夜)の概念が薄いため評価対象外
    return entries, None


# ── シフト評価ロジック ────────────────────
def _rule_violations(entries, staff):
    """
    勤務表のルール違反を検出する（/api/rules/check と同じ判定基準を
    エントリ配列に対して直接適用できる形に切り出したもの）。

    entries: [{staff_id, date('YYYY-MM-DD'), shift_type}, ...]
    staff:   {staff_id: {name, monthly_limit, ...}}
    戻り値: [{type:'warning'|'error', staff_id, name, date, rule, message}, ...]
    """
    by_staff = defaultdict(list)
    for e in entries:
        by_staff[e["staff_id"]].append(e)

    violations = []
    for sid, rows in by_staff.items():
        rows = sorted(rows, key=lambda x: x["date"])
        s = staff.get(sid, {})
        name = s.get("name", str(sid))
        limit = s.get("monthly_limit", 21)
        work = 0
        nstreak = 0
        for i, r in enumerate(rows):
            v = r["shift_type"]
            cur = r["date"]
            if v == "夜":
                is_cont = False
                if i > 0:
                    prev_v = rows[i-1]["shift_type"]
                    prev_d = rows[i-1]["date"]
                    gap = (datetime.strptime(cur, "%Y-%m-%d") - datetime.strptime(prev_d, "%Y-%m-%d")).days
                    is_cont = (prev_v == "明" and gap == 1)
                if not is_cont:
                    nstreak = 0
                nstreak += 1
                if nstreak == 3:
                    violations.append({"type": "warning", "staff_id": sid, "name": name, "date": cur,
                                        "rule": "連続夜勤3回（警告）", "message": f"{cur} {name} 連続夜勤3回"})
                elif nstreak >= 4:
                    violations.append({"type": "error", "staff_id": sid, "name": name, "date": cur,
                                        "rule": "連続夜勤4回以上（禁止）", "message": f"{cur} {name} 連続夜勤{nstreak}回（禁止）"})
            elif v == "明":
                pass
            else:
                nstreak = 0
            if v != "休":
                work += 1
            if v == "早" and i > 0:
                prev_v = rows[i-1]["shift_type"]
                prev_d = rows[i-1]["date"]
                gap = (datetime.strptime(cur, "%Y-%m-%d") - datetime.strptime(prev_d, "%Y-%m-%d")).days
                if prev_v == "明" and gap == 1:
                    violations.append({"type": "error", "staff_id": sid, "name": name, "date": cur,
                                        "rule": "夜勤明け翌日の早出禁止", "message": f"{cur} {name} 夜勤明け→早出"})
        if work > limit:
            violations.append({"type": "warning", "staff_id": sid, "name": name, "date": None,
                                "rule": "就業日数超過", "message": f"{name} 就業{work}日/上限{limit}日"})
    return violations


def _evaluate_schedule(entries, staff, req_map, units, year, month, last_day):
    """
    完成したシフト表を4つの観点で評価し、0〜100点のスコアと診断詳細を返す。

    優先順位（重み）:
      A. 希望休・希望シフトの反映率        … 40点
      B. 人員不足日（候補者ゼロ）の発生防止 … 30点
      C. 公平性（早/遅/夜の回数の偏り）     … 20点
      D. ルール違反（連続夜勤・夜勤明け早出・就業日数超過等）… 10点

    entries: [{staff_id, unit_id, date('YYYY-MM-DD'), shift_type, ...}, ...]
    staff:   {staff_id: {name, can_early, can_late, can_night, can_night_only,
                          monthly_limit, ...}}
    req_map: {staff_id: {day(int): req_type}}
    units:   [{id, unit_no, name}, ...]
    """
    import statistics

    # 参照しやすい形へ変換: sched[(sid,day)] = shift_type / place[(sid,day)] = unit_id
    sched = {}
    place = {}
    for e in entries:
        d = int(e["date"][8:10])
        sched[(e["staff_id"], d)] = e["shift_type"]
        place[(e["staff_id"], d)] = e["unit_id"]

    # ── A. 希望休・希望シフトの反映率（40点） ──────────
    type_map = {"rest": "希休", "hol": "有給", "absence": "欠勤", "early": "早", "late": "遅",
                 "night": "夜", "day": "日", "ake": "明"}
    req_total = 0
    req_ok = 0
    unmet = []
    for sid, dmap in req_map.items():
        for day, rt in dmap.items():
            if not (1 <= day <= last_day) or rt not in type_map:
                continue
            req_total += 1
            expect = type_map[rt]
            actual = sched.get((sid, day), "休")
            # 後方互換：DBに古い「休」が残っている場合も希休・有給・欠勤と同等に扱う
            actual_norm = actual
            if expect in ("希休", "有給", "欠勤") and actual == "休":
                actual_norm = expect  # 旧データの「休」は期待値と一致したとみなす
            if actual_norm == expect:
                req_ok += 1
            else:
                unmet.append({
                    "staff_id": sid, "name": staff.get(sid, {}).get("name", str(sid)),
                    "date": f"{year}-{month:02d}-{day:02d}",
                    "requested": rt, "expected": expect, "actual": actual,
                })
    fulfillment_rate = (req_ok / req_total) if req_total else 1.0
    score_a = round(fulfillment_rate * 40, 1)

    # ── B. 人員不足日（候補者ゼロ）の発生（30点） ──────
    shortage_days = []
    total_slots = 0
    filled_slots = 0
    sids = list(staff.keys())

    # unit_required_staffテーブルから必要人員を取得
    # DBに設定がないユニット(管理ユニット等)はスキップ対象
    req_map_unit = {}
    for unit in units:
        uid = unit["id"]
        rows_req = qdb(
            "SELECT shift_type, required FROM unit_required_staff WHERE unit_id=?", (uid,))
        if rows_req:
            req_map_unit[uid] = {r["shift_type"]: r["required"] for r in rows_req}

    for unit in units:
        uid = unit["id"]
        uname = unit.get("name", str(uid))
        # 管理ユニット(residents=0またはunit_required_staffに設定なし)はスキップ
        if not req_map_unit.get(uid):
            continue
        for day in range(1, last_day + 1):
            for slot in ("早", "遅", "夜"):
                required = req_map_unit[uid].get(slot, 1)
                total_slots += required
                actual = sum(
                    1 for sid in sids
                    if sched.get((sid, day)) == slot and place.get((sid, day)) == uid)
                filled = min(actual, required)
                filled_slots += filled
                shortage = required - filled
                if shortage > 0:
                    shortage_days.append({
                        "unit_id": uid, "unit_name": uname,
                        "date": f"{year}-{month:02d}-{day:02d}",
                        "slot": slot,
                        "required": required, "actual": actual, "shortage": shortage,
                    })
    shortage_rate = (1 - filled_slots / total_slots) if total_slots else 0
    score_b = round((1 - shortage_rate) * 30, 1)

    # ── C. 公平性（早/遅/夜の回数の偏り）（20点） ──────
    fairness_detail = {}
    cvs = []
    for slot, can_key in (("早", "can_early"), ("遅", "can_late"), ("夜", "can_night")):
        eligible = [sid for sid, s in staff.items()
                    if s.get(can_key) and not (slot in ("早", "遅") and s.get("can_night_only"))]
        counts = {sid: sum(1 for day in range(1, last_day + 1) if sched.get((sid, day)) == slot)
                  for sid in eligible}
        values = list(counts.values())
        if len(values) >= 2 and sum(values) > 0:
            mean = sum(values) / len(values)
            cv = statistics.pstdev(values) / mean if mean > 0 else 0
        else:
            cv = 0
        cvs.append(cv)
        fairness_detail[slot] = {
            "counts": {staff.get(sid, {}).get("name", str(sid)): c for sid, c in counts.items()},
            "cv": round(cv, 2),
        }
    avg_cv = sum(cvs) / len(cvs) if cvs else 0
    score_c = round(max(0, 1 - avg_cv) * 20, 1)

    # ── D. ルール違反（警告）の最小化（10点） ─────────
    violations = _rule_violations(entries, staff)
    penalty = sum(2 if v["type"] == "error" else 1 for v in violations)
    score_d = round(max(0, 10 - penalty), 1)

    # ── E. 希望勤務回数達成率（早/遅/夜 shift_targets）（10点） ──
    # 対象: staff_shift_targetsに登録がある職員の目標回数に対する達成度
    target_items = []
    for sid, s in staff.items():
        tgts = s.get("shift_targets") or {}
        for shift_type, tinfo in tgts.items():
            if not tinfo or tinfo.get("target", 0) == 0:
                continue
            actual_cnt = sum(1 for day in range(1, last_day + 1) if sched.get((sid, day)) == shift_type)
            target_cnt = tinfo["target"]
            rate = min(actual_cnt / target_cnt, 1.0)  # 達成率(超過は100%扱い)
            gap = actual_cnt - target_cnt
            target_items.append({
                "staff_id": sid,
                "name": s.get("name", str(sid)),
                "shift_type": shift_type,
                "target": target_cnt,
                "actual": actual_cnt,
                "rate": round(rate * 100, 1),
                "gap": gap,  # 正=超過, 負=未達
            })
    if target_items:
        avg_target_rate = sum(t["rate"] for t in target_items) / len(target_items) / 100
    else:
        avg_target_rate = 1.0  # 設定なし = 満点
    score_e = round(avg_target_rate * 10, 1)

    total = round(score_a + score_b + score_c + score_d + score_e, 1)
    return {
        "total_score": total,
        "max_total": 110,
        "scores": {
            "request_fulfillment": score_a,
            "staffing_shortage": score_b,
            "fairness": score_c,
            "rule_violations": score_d,
            "shift_target_achievement": score_e,
        },
        "max_scores": {
            "request_fulfillment": 40, "staffing_shortage": 30,
            "fairness": 20, "rule_violations": 10,
            "shift_target_achievement": 10,
        },
        "detail": {
            "request_fulfillment": {
                "rate": round(fulfillment_rate * 100, 1),
                "total": req_total, "fulfilled": req_ok, "unmet": unmet,
            },
            "staffing_shortage": {
                "shortage_count": len(shortage_days),
                "total_slots": total_slots, "days": shortage_days,
            },
            "fairness": fairness_detail,
            "rule_violations": {
                "count": len(violations), "items": violations,
            },
            "shift_target_achievement": {
                "rate": round(avg_target_rate * 100, 1),
                "items": target_items,
            },
        },
    }


# ── 過去シフト（Excel/CSV）の取り込み・評価 ──────────────
# 1行目=ヘッダ(1列目=職員名の見出し、以降=日付/日番号)
# 2行目以降=1列目=職員名、以降=シフト記号
DEFAULT_SHIFT_CODE_MAP = {
    "早": "早", "遅": "遅", "夜": "夜", "明": "明", "休": "休", "日": "日",
    "希休": "希休", "有給": "有給", "欠勤": "欠勤",
    "早出": "早", "遅出": "遅", "夜勤": "夜", "明け": "明", "明番": "明", "日勤": "日",
    "公": "休", "公休": "休", "有": "有給", "有休": "有給", "休み": "休",
    "希": "希休", "希望休": "希休", "欠": "欠勤",
    "A": "早", "P": "遅", "N": "夜", "E": "早", "L": "遅",
    "AM": "早", "PM": "遅",
    "": "休", "-": "休", "ー": "休", "−": "休", "/": "休",
}


def _normalize_shift_code(raw, code_map=None):
    """シフト記号を早/遅/夜/明/休/日のいずれかに変換する。
    戻り値: (normalized_code, matched) matched=Falseの場合は'休'扱いだが
    元の記号は未知記号としてunknown_codesに記録すること。"""
    code_map = code_map or DEFAULT_SHIFT_CODE_MAP
    if raw is None:
        return "休", True
    key = str(raw).strip()
    if key in code_map:
        return code_map[key], True
    if key in ("早", "遅", "夜", "明", "休", "日"):
        return key, True
    return "休", False


def _parse_day_from_header(cell, year, month, last_day):
    """ヘッダ1セルから当月の日(1〜last_day)を推定する。該当しない場合はNone。"""
    if cell is None:
        return None
    if isinstance(cell, datetime):
        d = cell.day
        return d if 1 <= d <= last_day else None
    if isinstance(cell, (int, float)) and not isinstance(cell, bool):
        d = int(cell)
        return d if 1 <= d <= last_day else None
    s = str(cell).strip()
    if not s:
        return None
    m = re.search(r"(\d{1,2})\s*[/\-月]\s*(\d{1,2})", s)
    if m:
        mm, dd = int(m.group(1)), int(m.group(2))
        if mm == month and 1 <= dd <= last_day:
            return dd
        return None
    m = re.fullmatch(r"\d{1,2}", s)
    if m:
        d = int(s)
        return d if 1 <= d <= last_day else None
    return None


def _read_table_rows(file_storage):
    """アップロードされたCSV/Excelファイルを2次元配列(行のリスト)で返す。"""
    filename = (file_storage.filename or "").lower()
    if filename.endswith((".xlsx", ".xlsm")):
        if not HAS_OPENPYXL:
            raise ValueError("サーバーにopenpyxlが導入されていないため.xlsxを読み込めません。CSV形式で再度お試しください")
        wb = openpyxl.load_workbook(file_storage, data_only=True, read_only=True)
        ws = wb.active
        return [list(row) for row in ws.iter_rows(values_only=True)]
    # CSV（日本語環境のExcelはShift-JIS/cp932で保存されることが多い）
    raw = file_storage.read()
    for enc in ("utf-8-sig", "cp932", "utf-8"):
        try:
            text = raw.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    else:
        text = raw.decode("utf-8", errors="replace")
    return [row for row in csv.reader(io.StringIO(text))]


def _parse_baseline_table(rows, fid, year, month, last_day, code_map=None):
    """
    過去シフト表(2次元配列)を entries 形式に変換する。
    戻り値: (entries, unmatched_names, unknown_codes, matched_days)
      entries: [{staff_id, unit_id, date, shift_type}, ...]
      unmatched_names: 職員マスタに見つからなかった氏名のリスト
      unknown_codes: {認識できなかった記号: 出現回数}
      matched_days: ヘッダから認識できた日(int)のリスト
    """
    if not rows:
        return [], [], {}, []

    header = rows[0]
    day_cols = {}
    for i, cell in enumerate(header[1:], start=1):
        d = _parse_day_from_header(cell, year, month, last_day)
        if d:
            day_cols[i] = d

    staff_rows = qdb(
        "SELECT id,name,unit_id FROM staff WHERE facility_id=? AND is_active=1", (fid,))
    by_name = {s["name"].strip(): dict(s) for s in staff_rows}

    entries = []
    unmatched_names = []
    unknown_codes = defaultdict(int)

    for row in rows[1:]:
        if not row or row[0] is None:
            continue
        name = str(row[0]).strip()
        if not name:
            continue
        s = by_name.get(name)
        if not s:
            if name not in unmatched_names:
                unmatched_names.append(name)
            continue
        for col, day in day_cols.items():
            if col >= len(row):
                continue
            raw = row[col]
            norm, matched = _normalize_shift_code(raw, code_map)
            if not matched and str(raw).strip():
                unknown_codes[str(raw).strip()] += 1
            entries.append({
                "staff_id": s["id"], "unit_id": s["unit_id"],
                "date": f"{year}-{month:02d}-{day:02d}", "shift_type": norm,
            })

    return entries, unmatched_names, dict(unknown_codes), sorted(day_cols.values())


def _gh_load_data(fid, year, month, restrict_unit_id=None):
    """_generate_gh のデータ取得フェーズを切り出した関数。

    DBから職員・スキル・ユニット・必要人員・曜日固定・希望・シフト目標を読み込み、
    生成ロジックに必要な形式（dict/defaultdict）に変換して返す。
    ロジック（STEP0〜後処理）は一切含まない。
    """
    staff_rows = qdb(
        """SELECT s.id, s.name, s.employment_type, s.can_day, s.can_early, s.can_late,
                  s.can_night, s.can_night_only, s.monthly_limit,
                  s.night_target, s.is_help, s.unit_id, f.id as fac_id
           FROM staff s
           JOIN units u ON s.unit_id=u.id
           JOIN facilities f ON s.facility_id=f.id
           WHERE (s.facility_id=?
               OR EXISTS (SELECT 1 FROM staff_skills ss
                          JOIN units tu ON ss.unit_id=tu.id
                          WHERE ss.staff_id=s.id AND tu.facility_id=?
                          AND ss.level IN ('ok','conditional')))
           AND s.is_active=1""", (fid, fid))
    staff = {s["id"]: dict(s) for s in staff_rows}

    skill_rows = qdb(
        """SELECT ss.staff_id, ss.unit_id, ss.level
           FROM staff_skills ss JOIN units u ON ss.unit_id=u.id
           WHERE u.facility_id=?""", (fid,))
    skill_map = defaultdict(dict)
    for sk in skill_rows:
        skill_map[sk["staff_id"]][sk["unit_id"]] = sk["level"]
    if restrict_unit_id is not None:
        for sid_k in list(skill_map.keys()):
            skill_map[sid_k] = {uid_k: lv for uid_k, lv in skill_map[sid_k].items()
                                 if uid_k == restrict_unit_id}

    units = qdb(
        "SELECT id, unit_no, name, residents FROM units WHERE facility_id=? ORDER BY unit_no", (fid,))
    admin_unit_ids = {u["id"] for u in units if (u["residents"] or 0) <= 0}
    operational_units = [u for u in units if u["id"] not in admin_unit_ids]

    req_rows = qdb(
        """SELECT urs.unit_id, urs.shift_type, urs.required, urs.is_admin_eligible
           FROM unit_required_staff urs
           JOIN units u ON urs.unit_id=u.id
           WHERE u.facility_id=?""", (fid,))
    unit_req = {}
    for r in req_rows:
        unit_req.setdefault(r["unit_id"], {})[r["shift_type"]] = {
            "required": r["required"],
            "is_admin_eligible": r["is_admin_eligible"],
        }
    for u in operational_units:
        uid = u["id"]
        if uid not in unit_req:
            unit_req[uid] = {
                "早": {"required": 1, "is_admin_eligible": 1},
                "遅": {"required": 1, "is_admin_eligible": 0},
                "夜": {"required": 1, "is_admin_eligible": 0},
            }

    fixed_rows = qdb(
        """SELECT sfd.staff_id, sfd.dow, sfd.shift_type
           FROM staff_fixed_days sfd
           JOIN staff s ON sfd.staff_id=s.id
           WHERE s.facility_id=?""", (fid,))
    fixed_map = defaultdict(dict)
    for r in fixed_rows:
        fixed_map[r["staff_id"]][r["dow"]] = r["shift_type"]

    reqs = qdb(
        "SELECT staff_id, day, req_type FROM requests WHERE year=? AND month=?",
        (year, month))
    req_map = defaultdict(dict)
    for r in reqs:
        req_map[r["staff_id"]][r["day"]] = r["req_type"]

    shift_target_rows = qdb(
        """SELECT sst.staff_id, sst.shift_type, sst.target, sst.is_upper_limit
           FROM staff_shift_targets sst
           JOIN staff s ON sst.staff_id=s.id
           WHERE s.facility_id=?""", (fid,))
    for r in shift_target_rows:
        sid = r["staff_id"]
        if sid in staff:
            if "shift_targets" not in staff[sid]:
                staff[sid]["shift_targets"] = {}
            staff[sid]["shift_targets"][r["shift_type"]] = {
                "target": r["target"], "is_upper": r["is_upper_limit"]
            }
    for sid, s in staff.items():
        if "shift_targets" not in s:
            s["shift_targets"] = {}
        if s["can_night_only"] and s["night_target"] > 0 and "夜" not in s["shift_targets"]:
            s["shift_targets"]["夜"] = {"target": s["night_target"], "is_upper": 0}

    return {
        "staff": staff,
        "skill_map": skill_map,
        "units": units,
        "admin_unit_ids": admin_unit_ids,
        "operational_units": operational_units,
        "unit_req": unit_req,
        "fixed_map": fixed_map,
        "req_map": req_map,
    }


def _generate_gh(fid, vid, year, month, last_day, lock_prefix=None, restrict_unit_id=None):
    """
    GH シフト生成ロジック v8（4週分割法＋パート・派遣優先）
    人間のシフト作成手順を忠実に再現:

    lock_prefix: {(staff_id, day): shift_type または (shift_type, unit_id)} を渡すと、
    その日のその職員のシフトを「既に確定済み」として扱い、STEP0以降の自動割当の
    対象から外す（上書きしない）。欠勤発生時に「発生日より前は公開済みの第1版の
    まま、発生日以降だけ再生成する」ための機能（2026-06-20追加）。

    restrict_unit_id: 指定すると、このユニット以外へのクロスユニット（ヘルプ）
    割当を一切禁止する。lock_prefixで他ユニットを完全ロックしている場合、
    既に充足しているそのユニットへ余剰職員が「ヘルプ」として追加で送り込まれて
    しまう（必要数1人のところへロック済み1人＋ヘルプでダブつく）問題への対処。
    【割当順序（条件が厳しい順）】
    1. 曜日固定スタッフ（固定曜日に固定シフトを先入力）
    2. 夜勤専従（月X回夜勤目標を達成するよう3日サイクルで配分）
    3. 遅出専門スタッフ（週X日の遅出を配分）
    4. 早出専門スタッフ
    5. 夜勤可兼務スタッフ（週1〜2回夜勤 + 早/遅）
    6. フル勤務可スタッフ（残りを埋める）

    【1日の割当順序】
    夜勤 → 早出 → 遅出

    【禁止ルール（絶対）】
    - 遅→翌日早: 警告（割当は可能）
    - 夜/明→早: 禁止
    - 明→夜: OK（夜→明→夜→明→夜→明まで可能）
    - 夜→明→夜→明→夜→明→夜: 禁止（4連続夜勤）
    - 連続勤務5日超え: 禁止

    【警告ルール】
    - 連続夜勤3回: 警告
    - 夜→明→夜→明→夜→明の翌日: 必ず休み

    【埋まらない日】
    - 警告を出し、ヘルプ可能一覧に記録
    """
    from datetime import date as dt

    # ── 動作確認用マーカー（デバッグ目的、2026-06-20追加）──────
    # 新しいapp.pyが実際にサーバーへ反映されているかを切り分けるための目印。
    # 問題切り分けが完了したら削除してよい。
    if vid and vid > 0:
        try:
            _bg = get_db()
            _bg.execute(
                "UPDATE shift_versions SET change_reason=? WHERE id=?",
                (__import__('json').dumps(
                    [f"[build-check] _generate_gh実行開始 "
                     f"{__import__('datetime').datetime.now().isoformat(timespec='seconds')} "
                     f"build=post_process_v1"], ensure_ascii=False),
                 vid))
            _bg.commit()
        except Exception:
            pass

    # ── データ取得（_gh_load_dataに切り出し済み） ─────────────
    _data = _gh_load_data(fid, year, month, restrict_unit_id)
    staff          = _data["staff"]
    skill_map      = _data["skill_map"]
    units          = _data["units"]
    admin_unit_ids = _data["admin_unit_ids"]
    operational_units = _data["operational_units"]
    unit_req       = _data["unit_req"]
    fixed_map      = _data["fixed_map"]
    req_map        = _data["req_map"]

    # ── 状態管理 ──────────────────────────────────
    # ShiftState: S/place/slots/night_cntを1か所で管理するクラス。
    # 従来は4つの独立した変数を57個のネスト関数が直接書き換えていたため、
    # 「どこで何が変わったか」が追いにくく、不具合の温床になっていた
    # （assign_slotのカウント取り違えバグなど）。1か所にまとめることで、
    # 将来的に「状態を変更する唯一の窓口」を増やしていける土台を作る。
    # 今回は外から見たインターフェース（S[sid][day]の参照方法）は一切変えない。
    # ロジックは変えず、置き場所だけ整理するリファクタリング（2026-06-21）。
    class ShiftState:
        def __init__(self, unit_ids):
            self.S = defaultdict(dict)          # {sid: {day: shift_type}}
            self.place = defaultdict(dict)      # {sid: {day: unit_id}}
            self.slots = {uid: defaultdict(dict) for uid in unit_ids}  # {uid:{day:{shift_type:sid}}}
            self.night_cnt = defaultdict(int)   # {sid: 夜勤回数}

    _state = ShiftState([u["id"] for u in units])
    S = _state.S
    place = _state.place
    slots = _state.slots
    night_cnt = _state.night_cnt
    warnings = []                # 割当不可の警告

    # lock_prefixで指定された日は「確定済み」として事前に埋める。
    # 以降のSTEPはslots/Sが既に埋まっている日を自動的にスキップする
    # 作り（割当前に必ず現在の充足状況を見るため）になっているので、
    # 通常のSTEP0以降の処理に一切手を入れずに済む。
    # 値は shift_type の文字列、または (shift_type, unit_id) のタプルを許容する。
    # タプルでunit_idを明示しないと、他ユニットへヘルプ中だった職員が
    # 自分の所属ユニットに固定されてしまい、ヘルプ先ユニットが「不足」と
    # 誤認識される（2026-06-20 実データ検証で発覚）。
    locked_days = set()
    if lock_prefix:
        for (lp_sid, lp_day), lp_val in lock_prefix.items():
            if lp_val is None:
                continue
            if isinstance(lp_val, tuple):
                lp_shift, lp_uid = lp_val
            else:
                lp_shift, lp_uid = lp_val, staff[lp_sid]["unit_id"]
            locked_days.add((lp_sid, lp_day))
            if lp_shift == "休":
                S[lp_sid][lp_day] = "休"
                continue
            uid = lp_uid or staff[lp_sid]["unit_id"]
            S[lp_sid][lp_day] = lp_shift
            place[lp_sid][lp_day] = uid
            slots[uid][lp_day][lp_shift] = lp_sid
            if lp_shift == "夜":
                night_cnt[lp_sid] += 1

    # 管理ユニット(unit_no=4)・入居者0ユニットの職員は通常シフト割当から除外
    # → 人員不足時のフォールバック(assign_slotのfallback)にのみ登場する
    admin_staff_ids = {sid for sid, s in staff.items() if s["unit_id"] in admin_unit_ids}
    # 通常割当対象(管理ユニット除く)
    regular_staff = {sid: s for sid, s in staff.items() if sid not in admin_staff_ids}

    # ── 4週分割法 ─────────────────────────────────
    # 1～7日、8～14日、15～21日、22日～月末の4区分で管理する。
    # 第4区分だけ日数が長くなる月は、区分日数に応じて目標回数を配分する。
    week_ranges = [
        (1, min(7, last_day)),
        (8, min(14, last_day)),
        (15, min(21, last_day)),
        (22, last_day),
    ]
    week_ranges = [(a, b) for a, b in week_ranges if a <= b]

    def week_index(day):
        for idx, (start, end) in enumerate(week_ranges):
            if start <= day <= end:
                return idx
        return len(week_ranges) - 1

    def split_target(total):
        """月間回数を4区分へ、区分日数に比例しつつ差が大きくならないよう配分。"""
        total = max(0, int(total or 0))
        lengths = [end - start + 1 for start, end in week_ranges]
        days_total = sum(lengths) or 1
        raw = [total * ln / days_total for ln in lengths]
        base = [int(v) for v in raw]
        remain = total - sum(base)
        order = sorted(range(len(raw)), key=lambda i: (-(raw[i] - base[i]), i))
        for i in order[:remain]:
            base[i] += 1
        return base

    def weekly_count(sid, shift_type, widx):
        start, end = week_ranges[widx]
        return sum(1 for d in range(start, end + 1) if S[sid].get(d) == shift_type)

    def monthly_shift_target(sid, shift_type):
        s = staff[sid]
        target = s.get("shift_targets", {}).get(shift_type)
        if target and target.get("target", 0) > 0:
            return int(target["target"])
        if shift_type == "夜" and s.get("night_target", 0) > 0:
            return int(s["night_target"])
        if shift_type in ("早", "遅"):
            return _auto_early_late_target(sid, shift_type)
        return 0

    def _auto_early_late_target(sid, shift_type):
        """早・遅の月間目標が未設定の職員向けに、契約勤務日数から自動算出する。

        修正仕様（2026-06）対象は「契約上の勤務日数を守るべき」パート・派遣のみ。
        正社員は仕様書4節の方針どおり、上限に届かせる強制配分を行わない
        （＝target=0のまま、従来通り「少ない人を優先」の緩い均等化に留める）。
        正社員にも適用した初版は、本来ユニット余剰として残すべき正社員の空きを
        強制的に均等化してしまい、週次偏り悪化・修正範囲過大の原因となったため、
        対象をパート・派遣に限定する。

          配分日数 ＝ 月間契約日数(monthly_limit) − 夜勤回数×2
          （夜勤1回＝夜入り+明けの2日分のため）
          早出・遅出が両方可能な職員は、配分日数を半々に分ける
          （端数が出る場合は早出側へ+1）。
          片方しかできない職員は、できる方に全配分する。
        staff_shift_targets で早/遅のtargetが明示設定されている職員には
        適用しない（呼び出し元 monthly_shift_target 側で既にガード済み）。
        """
        s = staff[sid]
        if s.get("can_night_only"):
            return 0
        employment_type = (s.get("employment_type") or "full_time").strip()
        if employment_type not in ("part_time", "dispatch"):
            return 0
        monthly_limit = int(s.get("monthly_limit") or 0)
        night_t = monthly_shift_target(sid, "夜")
        alloc_days = max(0, monthly_limit - night_t * 2)
        can_e, can_l = bool(s.get("can_early")), bool(s.get("can_late"))
        if can_e and can_l:
            early_part = -(-alloc_days // 2)  # 端数は早出側へ+1（切り上げ）
            late_part = alloc_days - early_part
            return early_part if shift_type == "早" else late_part
        if can_e and shift_type == "早":
            return alloc_days
        if can_l and shift_type == "遅":
            return alloc_days
        return 0

    def weekly_target(sid, shift_type, widx):
        return split_target(monthly_shift_target(sid, shift_type))[widx]

    def dow_of(day):
        """day(1始まり)の曜日 0=月..6=日"""
        return dt(year, month, day).weekday()

    def prev(sid, day):
        return S[sid].get(day-1) if day > 1 else None

    def prev2(sid, day):
        return S[sid].get(day-2) if day > 2 else None

    def streak(sid, day):
        """day-1まで連続勤務日数"""
        count = 0
        for d in range(day-1, 0, -1):
            if S[sid].get(d) not in (None, "休"):
                count += 1
            else:
                break
        return count

    def night_streak(sid, day):
        """夜→明→夜→明→夜→明の判定（3連続夜勤）"""
        # day-6〜day-1で夜→明→夜→明→夜→明のパターン
        if day < 7: return False
        return all(S[sid].get(day-6+i) == s
                   for i, s in enumerate(["夜","明","夜","明","夜","明"]))

    def can_work(sid, day, shift_type, uid, relax_night=False, allow_overwork=False):
        """シフト割当可否チェック
        allow_overwork=True: 人員不足時に月間上限超過を例外許可
        戻り値: (bool, reason_str|None)
        """
        s = staff.get(sid)
        if not s: return False, "職員データなし"
        if S[sid].get(day) is not None: return False, f"当日={S[sid].get(day)}"
        req = req_map[sid].get(day)
        if req in ("rest","hol"): return False, "希望休"

        if streak(sid, day) >= 5: return False, "連続勤務5日"

        if shift_type == "夜":
            if not s["can_night"]: return False, "夜勤不可"
            p = prev(sid, day)
            if p in ("夜", "明"): return False, f"前日={p}"
            next_req = req_map[sid].get(day + 1) if day < last_day else None
            if next_req in ("rest", "hol"): return False, "翌日希望休"
            recent = [S[sid].get(day-i) for i in range(1,7) if day-i >= 1]
            if recent[:6] == ["明","夜","明","夜","明","夜"]:
                return False, "4連続夜勤禁止"
            if not relax_night:
                if recent[:3] == ["明","夜","明"] and recent[3:4] == ["夜"]:
                    return False, "2連続夜勤済み"
            used = sum(1 for d in range(1,day) if S[sid].get(d) not in (None,"休"))
            if used + 2 > s["monthly_limit"]:
                if not allow_overwork: return False, f"上限到達({used}/{s['monthly_limit']})"
            if s["can_night_only"] and s["night_target"] > 0 and night_cnt[sid] >= s["night_target"]:
                return False, f"夜勤上限({night_cnt[sid]}/{s['night_target']})"
            t_night = s["shift_targets"].get("夜")
            if t_night and t_night["is_upper"] and night_cnt[sid] >= t_night["target"]:
                return False, "夜勤target上限"
        elif shift_type == "早":
            if not s["can_early"] or s["can_night_only"]:
                return False, "夜勤専従" if s["can_night_only"] else "早出不可"
            p = prev(sid, day)
            if p in ("夜","明"): return False, f"前日={p}(早出禁止)"
            used = sum(1 for d in range(1,day) if S[sid].get(d) not in (None,"休"))
            if used >= s["monthly_limit"]:
                if not allow_overwork: return False, f"上限到達({used}/{s['monthly_limit']})"
            t_early = s["shift_targets"].get("早")
            if t_early and t_early["is_upper"]:
                early_cnt = sum(1 for d in range(1,day) if S[sid].get(d) == "早")
                if early_cnt >= t_early["target"]: return False, "早出target上限"
        elif shift_type == "遅":
            if not s["can_late"] or s["can_night_only"]:
                return False, "夜勤専従" if s["can_night_only"] else "遅出不可"
            p = prev(sid, day)
            if p in ("夜","明"): return False, f"前日={p}"
            used = sum(1 for d in range(1,day) if S[sid].get(d) not in (None,"休"))
            if used >= s["monthly_limit"]:
                if not allow_overwork: return False, f"上限到達({used}/{s['monthly_limit']})"
            t_late = s["shift_targets"].get("遅")
            if t_late and t_late["is_upper"]:
                late_cnt = sum(1 for d in range(1,day) if S[sid].get(d) == "遅")
                if late_cnt >= t_late["target"]: return False, "遅出target上限"
        elif shift_type == "遅_警告":
            if not s["can_late"] or s["can_night_only"]:
                return False, "夜勤専従" if s["can_night_only"] else "遅出不可"
            used = sum(1 for d in range(1,day) if S[sid].get(d) not in (None,"休"))
            if used >= s["monthly_limit"]:
                if not allow_overwork: return False, f"上限到達({used}/{s['monthly_limit']})"
        elif shift_type == "日":
            if not s["can_day"] or s["can_night_only"]: return False, "日勤不可"
            p = prev(sid, day)
            if p in ("夜","明"): return False, f"前日={p}"
            used = sum(1 for d in range(1,day) if S[sid].get(d) not in (None,"休"))
            if used >= s["monthly_limit"]:
                if not allow_overwork: return False, f"上限到達({used}/{s['monthly_limit']})"

        if s["unit_id"] != uid and skill_map[sid].get(uid,"no") not in ("ok","conditional"):
            return False, "スキル不足"
        return True, None

    def assign(sid, day, uid, shift_type):
        S[sid][day] = shift_type
        place[sid][day] = uid
        slots[uid][day][shift_type] = sid
        if shift_type == "夜":
            night_cnt[sid] += 1
            if day < last_day and S[sid].get(day+1) is None:
                S[sid][day+1] = "明"
                place[sid][day+1] = uid
                slots[uid][day+1]["明"] = sid
            # 3連続夜の翌々日（明けの翌日）は休み強制
            # 3連続夜勤（夜→明→夜→明→夜）の翌々日は必ず休み
        if day >= 5:
            pat = [S[sid].get(day-4), S[sid].get(day-3),
                   S[sid].get(day-2), S[sid].get(day-1)]
            if pat == ["夜","明","夜","明"]:  # 今日で3回目の夜勤
                rest_day = day + 2  # 明けの翌日
                if rest_day <= last_day and S[sid].get(rest_day) is None:
                    S[sid][rest_day] = "休"

    def contract_priority(sid):
        """パート・派遣の契約勤務日数を正社員より先に満たすための順位。

        戻り値: (雇用区分順位, 契約残日数の負値)
        ・契約日数未達のパート/派遣 = 0（最優先）
        ・その他                    = 1
        """
        s = staff[sid]
        used = sum(1 for d in range(1, last_day + 1)
                   if S[sid].get(d) not in (None, "休"))
        gap = max(0, int(s.get("monthly_limit") or 0) - used)
        employment_type = (s.get("employment_type") or "full_time").strip()
        is_hourly_contract = employment_type in ("part_time", "dispatch")
        return (0 if is_hourly_contract and gap > 0 else 1, -gap)

    def rank_for(sid, day, uid, prefer_types):
        s = staff[sid]
        widx = week_index(day)
        early_cnt = sum(1 for d in range(1, last_day + 1) if S[sid].get(d) == "早")
        late_cnt  = sum(1 for d in range(1, last_day + 1) if S[sid].get(d) == "遅")
        night_cnt_s = sum(1 for d in range(1, last_day + 1) if S[sid].get(d) == "夜")
        used = sum(1 for d in range(1, last_day + 1) if S[sid].get(d) not in (None, "休"))

        def target_score(shift_type, current_cnt):
            t = s["shift_targets"].get(shift_type)
            if t and t.get("target", 0) > 0:
                if t["is_upper"]:
                    return 2 if current_cnt >= t["target"] else 1
                target_val = int(t["target"])
            else:
                target_val = monthly_shift_target(sid, shift_type)
                if target_val <= 0:
                    return 1
            still_needed = max(0, target_val - current_cnt)
            return 1 if still_needed == 0 else -(still_needed / max(last_day, 1))

        if "早" in prefer_types or prefer_types == ["early","day"]:
            shift_type, current_cnt = "早", early_cnt
            req_types = ("early", "day")
        elif "遅" in prefer_types or prefer_types == ["late","day"]:
            shift_type, current_cnt = "遅", late_cnt
            req_types = ("late", "day")
        else:
            shift_type, current_cnt = "夜", night_cnt_s
            req_types = tuple(prefer_types)

        wcnt = weekly_count(sid, shift_type, widx)
        wtarget = weekly_target(sid, shift_type, widx)
        # 目標設定がない早・遅でも、その週の回数が少ない職員を優先する。
        weekly_gap = (wtarget - wcnt) if wtarget > 0 else -wcnt
        contract_gap = max(0, int(s["monthly_limit"] or 0) - used)
        ts = target_score(shift_type, current_cnt)

        employment_order, employment_gap = contract_priority(sid)

        return (
            employment_order,                                       # 契約日数未達のパート・派遣を最優先
            -weekly_gap,                                             # その中で週目標の不足が大きい人を優先
            employment_gap,                                         # 契約残日数が大きい人を優先
            ts,                                                      # 月間シフト目標未達を優先
            0 if req_map[sid].get(day) in req_types else 1,         # 希望勤務優先
            -contract_gap,                                          # その他職員も雇用日数の不足を考慮
            current_cnt,                                            # 月間回数の少ない人を優先
            used,                                                   # 総就業日数の少ない人を優先
            0 if s["unit_id"] == uid else 1,                    # 所属ユニット優先
            0 if s["fac_id"] == fid else 1,                     # 同一事業所優先
            sid,
        )

    # ── STEP 0: 希望休・有給・欠勤・前月明け・日勤希望を先に確定 ─────────
    for sid in staff:
        for day, rt in req_map[sid].items():
            if (sid, day) in locked_days:
                continue
            if rt in REQ_TYPE_MAP and 1 <= day <= last_day:
                S[sid][day] = REQ_TYPE_MAP[rt]  # rest→希休, hol→有給, absence→欠勤
            elif rt == "ake" and day == 1:
                # 前月夜勤明け
                uid = staff[sid]["unit_id"]
                S[sid][1] = "明"
                place[sid][1] = uid
                slots[uid][1]["明"] = sid
            elif rt == "day" and 1 <= day <= last_day:
                # 日勤希望：can_dayを満たす職員のみ、この時点で確定させる。
                # ここで確定しておかないと、後続STEPで休みや夜勤に上書きされてしまう
                # （2026-07-10修正：STEP0にこの分岐が無く、日勤希望が無視される不具合）。
                s = staff[sid]
                if s["can_day"]:
                    uid = s["unit_id"]
                    S[sid][day] = "日"
                    place[sid][day] = uid
                    slots[uid][day]["日"] = sid

    # ── STEP 1: 曜日固定スタッフを先に割当 ──────────
    for sid, dow_map in fixed_map.items():
        s = staff.get(sid)
        if not s: continue
        uid = s["unit_id"]
        for day in range(1, last_day+1):
            dw = dow_of(day)
            if dw in dow_map and S[sid].get(day) is None:
                ft = dow_map[dw]
                ok, _ = can_work(sid, day, ft, uid)
                if ok:
                    assign(sid, day, uid, ft)

    # ── STEP 2: 夜勤専従の夜勤を配分（目標回数まで）──
    # シミュレーション方式:
    #   先に「2連続夜勤まで」の厳格モードで全日を割当てる
    #   割当不足日が出た場合のみ「3連続夜勤まで」の緩和モードで再割当
    night_only_staff = [s for s in staff.values()
                        if s["can_night_only"] and s["fac_id"] == fid]
    for s in night_only_staff:
        if s["night_target"] == 0:
            # night_targetが未設定(0)の場合のみ月間上限//2をデフォルトにする
            # ※雇用契約で回数が決まっている場合は職員登録で必ず設定すること
            s["night_target"] = s["monthly_limit"] // 2

    def fill_nights(relax_night=False):
        """夜勤を4週に分け、各週の目標不足が大きい職員から割り当てる。"""
        shortage_days = []
        for unit in operational_units:
            uid = unit["id"]
            req_count = unit_req.get(uid, {}).get("夜", {}).get("required", 1)
            unit_night_only = [s for s in night_only_staff if s["unit_id"] == uid]
            for widx, (week_start, week_end) in enumerate(week_ranges):
                for day in range(week_start, week_end + 1):
                    already_night = sum(
                        1 for sid2 in staff
                        if S[sid2].get(day) == "夜" and place[sid2].get(day) == uid)
                    if already_night >= req_count:
                        continue
                    cands = sorted(
                        [s for s in unit_night_only
                         if can_work(s["id"], day, "夜", uid, relax_night)[0]
                         and night_cnt[s["id"]] < s["night_target"]],
                        key=lambda s: (
                            contract_priority(s["id"])[0],
                            -(weekly_target(s["id"], "夜", widx) - weekly_count(s["id"], "夜", widx)),
                            contract_priority(s["id"])[1],
                            night_cnt[s["id"]] / max(s["night_target"], 1),
                            night_cnt[s["id"]],
                            s["id"],
                        ))
                    if cands:
                        assign(cands[0]["id"], day, uid, "夜")
                    else:
                        shortage_days.append((uid, day))
        return shortage_days

    # 厳格モードで試行（2連続まで）
    shortage = fill_nights(relax_night=False)

    # 不足日がある場合、緩和モードで再試行（3連続まで、警告付き）
    if shortage:
        for uid, day in shortage:
            req_count = unit_req.get(uid, {}).get("夜", {}).get("required", 1)
            already_night = sum(
                1 for sid2 in staff
                if S[sid2].get(day) == "夜" and place[sid2].get(day) == uid)
            if already_night >= req_count:
                continue
            unit_night_only = [s for s in night_only_staff if s["unit_id"] == uid]
            cands = sorted(
                [s for s in unit_night_only
                 if can_work(s["id"], day, "夜", uid, relax_night=True)[0]
                 and night_cnt[s["id"]] < s["night_target"]],
                key=lambda s: (
                    contract_priority(s["id"])[0],
                    contract_priority(s["id"])[1],
                    night_cnt[s["id"]],
                    s["id"]
                ))
            if cands:
                assign(cands[0]["id"], day, uid, "夜")
                warnings.append({
                    "type": "warning", "unit_id": uid, "day": day,
                    "slot": "夜", "staff_id": cands[0]["id"],
                    "msg": f"{day}日 {cands[0]['name']} 連続夜勤3回（警告）"
                })

    # ── STEP 3: 夜勤可兼務の夜勤を4週均等に配分 ──────
    night_capable = [s for s in staff.values()
                     if s["can_night"] and not s["can_night_only"]]

    for unit in operational_units:
        uid = unit["id"]
        req_count = unit_req.get(uid, {}).get("夜", {}).get("required", 1)
        unit_night_capable = [s for s in night_capable if s["unit_id"] == uid]
        unit_early_capable = [s for s in regular_staff.values()
                              if s["unit_id"] == uid and s["can_early"] and not s["can_night_only"]]
        for widx, (week_start, week_end) in enumerate(week_ranges):
            for day in range(week_start, week_end + 1):
                already_night = sum(
                    1 for sid2 in staff
                    if S[sid2].get(day) == "夜" and place[sid2].get(day) == uid)
                if already_night >= req_count:
                    continue
                cands = sorted(
                    [s for s in unit_night_capable if can_work(s["id"], day, "夜", uid)[0]],
                    key=lambda s: (
                        contract_priority(s["id"])[0],
                        -(weekly_target(s["id"], "夜", widx) - weekly_count(s["id"], "夜", widx)),
                        contract_priority(s["id"])[1],
                        night_cnt[s["id"]] / max(s["night_target"], 1),
                        night_cnt[s["id"]],
                        0 if req_map[s["id"]].get(day) == "night" else 1,
                        sum(1 for d in range(1, last_day + 1) if S[s["id"]].get(d) not in (None, "休")),
                        s["id"],
                    ))
                if cands:
                    chosen = cands[0]
                    next_day = day + 1
                    if next_day <= last_day:
                        early_next = [
                            s for s in unit_early_capable
                            if s["id"] != chosen["id"]
                            and S[s["id"]].get(next_day) is None
                            and req_map[s["id"]].get(next_day) not in ("rest", "hol")
                            and S[s["id"]].get(day) not in ("夜", "明")
                        ]
                        if not early_next:
                            for alt in cands[1:]:
                                early_next_alt = [
                                    s for s in unit_early_capable
                                    if s["id"] != alt["id"]
                                    and S[s["id"]].get(next_day) is None
                                    and req_map[s["id"]].get(next_day) not in ("rest", "hol")
                                    and S[s["id"]].get(day) not in ("夜", "明")
                                ]
                                if early_next_alt:
                                    chosen = alt
                                    break
                    assign(chosen["id"], day, uid, "夜")

    # ── STEP 4: 早出・遅出を割当（1日ずつ早→遅の順）──
    def _unit_ok(sid, uid):
        s = staff[sid]
        if s["unit_id"] == uid:
            return True
        return skill_map.get(sid, {}).get(uid) in ("ok", "conditional")

    def assign_slot(uid, day, slot_type):
        """
        早出/遅出を必要人員数(unit_req)が満たされるまで割当。
        1. regular_staffから通常候補を探す
        2. ゼロ → is_admin_eligible=1なら管理職も候補に加える
        3. ゼロ → 上限超過を例外許可してもう一度探す（allow_overwork=True）
        4. ゼロ → 失敗理由ログ付きでパターン不足警告を記録
        """
        req_info = unit_req.get(uid, {}).get(slot_type, {"required": 1, "is_admin_eligible": 0})
        required = req_info["required"]
        admin_ok = req_info["is_admin_eligible"]
        req_names = ["early","day"] if slot_type == "早" else ["late","day"]

        # 既存バグ修正(2026-06-21): slots[uid][day]は {shift_type: staff_id} 構造。
        # 旧コードは items() を (sid2, sh) として取り違えて比較しており、
        # 「staff_id文字列 == shift_type」が常にFalseになるため、充足済みでも
        # 常に already=0 と誤認識していた。結果、欠勤対応の再生成中、既に
        # ロックで充足している他ユニットへも「不足」と誤判定してヘルプを
        # 重複投入してしまっていた（実データ検証で発覚）。
        already = 1 if slot_type in slots[uid][day] else 0
        remaining = required - already
        if remaining <= 0:
            return

        for _ in range(remaining):
            # ① regular_staffから候補（明け以外）
            normal = sorted(
                [s for s in regular_staff.values()
                 if can_work(s["id"], day, slot_type, uid)[0]
                 and S[s["id"]].get(day-1) != "明"],
                key=lambda s: rank_for(s["id"], day, uid, req_names))

            if normal:
                chosen = normal[0]
                if slot_type == "早" and S[chosen["id"]].get(day-1) == "遅":
                    warnings.append({
                        "type": "warning", "unit_id": uid, "day": day,
                        "slot": slot_type, "is_help_needed": False,
                        "staff_id": chosen["id"],
                        "msg": f"{day}日 {chosen['name']} 遅出→早出（警告）"
                    })
                assign(chosen["id"], day, uid, slot_type)
                continue

            # ② フォールバック候補（admin_ok=1なら管理職も含む）
            # この経路はユニットを問わず最終手段として拾う設計のため、通常生成
            # ではskill_map制限をかけない。ただしrestrict_unit_id指定時（欠勤
            # 対応の部分再生成）だけは、対象ユニット以外への漏れを防ぐため
            # _unit_okでも絞る（2026-06-21追加）。
            pool = staff.values() if admin_ok else regular_staff.values()
            fallback = sorted(
                [s for s in pool
                 if S[s["id"]].get(day) is None
                 and req_map[s["id"]].get(day) not in ("rest","hol")
                 and not s["can_night_only"]
                 and (slot_type == "早" and s["can_early"]
                      or slot_type == "遅" and s["can_late"])
                 and S[s["id"]].get(day-1) != "夜"
                 and (slot_type != "早" or S[s["id"]].get(day-1) != "明")
                 and sum(1 for d in range(1,day) if S[s["id"]].get(d) not in (None,"休")) < s["monthly_limit"]
                 and (restrict_unit_id is None or _unit_ok(s["id"], uid))],
                key=lambda s: (
                    1 if s["id"] in admin_staff_ids else 0,
                    rank_for(s["id"], day, uid, req_names)
                ))

            if fallback:
                chosen = fallback[0]
                is_admin = chosen["id"] in admin_staff_ids
                warnings.append({
                    "type": "warning", "unit_id": uid, "day": day,
                    "slot": slot_type, "is_help_needed": True,
                    "staff_id": chosen["id"],
                    "msg": f"{day}日 {uid} {slot_type} ヘルプ必要: {chosen['name']}{'(管理者応援)' if is_admin else ''}"
                })
                assign(chosen["id"], day, uid, slot_type)
                continue

            # ③ 上限超過を例外許可して再チェック（allow_overwork=True）
            overwork_cands = sorted(
                [s for s in regular_staff.values()
                 if can_work(s["id"], day, slot_type, uid, allow_overwork=True)[0]
                 and S[s["id"]].get(day-1) != "明"
                 and (slot_type != "早" or S[s["id"]].get(day-1) not in ("夜","明"))],
                key=lambda s: rank_for(s["id"], day, uid, req_names))

            if overwork_cands:
                chosen = overwork_cands[0]
                used = sum(1 for d in range(1, day) if S[chosen["id"]].get(d) not in (None, "休"))
                warnings.append({
                    "type": "warning", "unit_id": uid, "day": day,
                    "slot": slot_type, "is_help_needed": False,
                    "staff_id": chosen["id"],
                    "msg": f"{day}日 {chosen['name']} {slot_type}【上限超過例外】({used}/{chosen['monthly_limit']}日)"
                })
                assign(chosen["id"], day, uid, slot_type)
                continue

            # ④ 全候補ゼロ → 失敗理由ログ付きでパターン不足警告
            # 全職員の除外理由を収集してログに記録
            failure_reasons = []
            for s in staff.values():
                if s["id"] in admin_staff_ids: continue
                ok, reason = can_work(s["id"], day, slot_type, uid, allow_overwork=True)
                if not ok:
                    failure_reasons.append({"name": s["name"], "reason": reason})
            warnings.append({
                "type": "error", "unit_id": uid, "day": day,
                "slot": slot_type, "is_help_needed": True,
                "staff_id": None,
                "msg": f"{day}日 {uid} {slot_type} 【パターン不足】候補者なし",
                "failure_reasons": failure_reasons,
            })

    # 早出を4週均等に割当てた後、遅出を4週均等に割り当てる。
    for slot_type in ("早", "遅"):
        for widx, (week_start, week_end) in enumerate(week_ranges):
            for day in range(week_start, week_end + 1):
                for unit in operational_units:
                    uid = unit["id"]
                    req_count = unit_req.get(uid, {}).get(slot_type, {}).get("required", 1)
                    # 同様のカウント取り違えバグ修正(2026-06-21)。
                    # slots[uid][day]は{shift_type: staff_id}なので、valuesは
                    # staff_id。文字列のslot_typeと比較しても常にFalseだった。
                    filled = 1 if slot_type in slots[uid][day] else 0
                    if filled < req_count:
                        assign_slot(uid, day, slot_type)

    # ── STEP 5: 未割当日は休みにする（自動日勤は作らない） ──
    # 日勤は、希望日入力で req_type='day' と指定された日だけ STEP 0 で確定する。
    # 月間上限に届かせる目的で自動的に「日」を追加しない。
    # 上限未満は許容し、配置上不要な余剰勤務を作らない。
    for sid, s in staff.items():
        for day in range(1, last_day + 1):
            if S[sid].get(day) is not None:
                continue
            S[sid][day] = "休"

    # ── STEP 5b: 就業日数超過の自動調整・警告 ────────
    for sid, s in staff.items():
        worked = sum(1 for d in range(1,last_day+1) if S[sid].get(d) not in (None,"休"))
        over = worked - s["monthly_limit"]
        if over > 0:
            # 超過分を月末側から「日」→「休」に変換して調整
            adjusted = 0
            for day in range(last_day, 0, -1):
                if adjusted >= over: break
                if S[sid].get(day) == "日" and req_map[sid].get(day) != "day":
                    S[sid][day] = "休"
                    adjusted += 1
            remaining_over = over - adjusted
            if remaining_over > 0:
                worked2 = sum(1 for d in range(1,last_day+1) if S[sid].get(d) not in (None,"休"))
                warnings.append({
                    "type": "warning", "unit_id": s["unit_id"], "day": 0,
                    "staff_id": sid,
                    "msg": f"{s['name']} 就業日数超過: {worked2}日（上限{s['monthly_limit']}日）。日勤での調整不可、シフト見直し必要"
                })

    # ── STEP 6: entries生成 ──────────────────────────
    entries = []
    for sid in staff:
        s = staff[sid]
        home_uid = s["unit_id"]
        for day in range(1, last_day+1):
            shift = S[sid].get(day, "休")
            assigned_uid = place[sid].get(day, home_uid)
            is_cross = 1 if assigned_uid != home_uid else 0
            entries.append({
                "version_id":   vid,
                "staff_id":     sid,
                "unit_id":      assigned_uid,
                "date":         f"{year}-{month:02d}-{day:02d}",
                "shift_type":   shift,
                "is_cross":     is_cross,
                "from_unit_id": home_uid if is_cross else None,
            })

    # 警告をshift_versionsのnoteに保存（既存の内容には追記する。
    # 冒頭の動作確認マーカーを上書きで消さないため）
    if warnings:
        _db4 = get_db()
        _row4 = qdb("SELECT change_reason FROM shift_versions WHERE id=?", (vid,), one=True)
        _existing4 = []
        if _row4 and _row4["change_reason"]:
            try:
                _existing4 = __import__('json').loads(_row4["change_reason"])
                if not isinstance(_existing4, list):
                    _existing4 = [str(_existing4)]
            except Exception:
                _existing4 = [_row4["change_reason"]]
        warn_json = __import__('json').dumps(
            _existing4 + [w["msg"] for w in warnings], ensure_ascii=False)
        _db4.execute(
            "UPDATE shift_versions SET change_reason=? WHERE id=?",
            (warn_json, vid))
        _db4.commit()

    # ── STEP 7: 評価（スコア計算・診断） ─────────────
    units_list = [dict(u) for u in operational_units]
    evaluation = _evaluate_schedule(entries, staff, req_map, units_list, year, month, last_day)

    # ── STEP 8: 反復改善ループ（作る→崩す→再配置→評価） ──────────
    # 初期案をそのまま返さず、必要配置を維持したまま職員間・日付間の
    # 入替えを繰り返し、雇用条件・4週均等・連勤/連休のバランスを改善する。
    # 終了条件: 最大5秒 / 最大3000試行 / 最高点更新なし400回。
    #
    # 【このブロックで定義される内部関数の一覧】
    # ─ 変換ユーティリティ ─────────────────────────────────
    #   _entries_to_sched(ents)          entries → (sc, pl) 辞書に変換
    #   _sched_to_entries(sc, pl)        (sc, pl) → entries に変換
    # ─ ロック・能力判定（状態変数に依存しない純粋関数） ───────────
    #   _is_part_or_dispatch(s)          パート・派遣かどうか
    #   _is_locked(sid, day, shift)      動かしてはいけない日か（希望・固定・明け）
    #   _is_protected(sid, day)          保護対象か（明けは除く：夜勤ペア調整用）
    #   _capable(sid, shift)             そのシフトに就ける能力があるか
    #   _adjacent_ok(sc, sid, day, new)  夜明け連続など近接ルールを満たすか
    # ─ 評価関数（sc/plだけで計算、外部データへの依存はクロージャ経由） ─
    #   _coverage_shortage(sc, pl)       必須配置の不足数
    #   _hard_signature(sc, pl)          絶対に悪化させない必須条件の要約
    #   _score(sc, pl)                   総合スコア（小さいほど良い）
    # ─ 改善操作（スワップ・移管の試行） ───────────────────────
    #   _try_transfer(sc, pl, rng)       職員間のシフト移管を試みる
    #   _try_cross_day_swap(sc, pl, rng) 日付をまたいだシフト交換を試みる
    #   _worked_count(sc, sid)           月間勤務日数
    #   _work_run_before(sc, sid, day)   指定日前の連勤日数
    #   _rest_run_around(sc, sid, day)   指定日前後の連休日数
    #   _missing_slots(sc, pl)           充足できていないスロットの一覧
    # ─ カバレッジ修復・系統的リバランス ────────────────────────
    #   _repair_coverage(sc, pl)         不足スロットを埋める修復パス
    #   _systematic_contract_rebalance   契約日数の偏りを系統的に修正
    #   _systematic_night_pattern_balance 夜勤パターンの均等化
    #   _systematic_pattern_balance      連勤・連休の偏りを修正
    # ─ ユニット余剰・後処理パス ─────────────────────────────
    #   _compute_unit_surplus(entries)   ユニット別の余剰日数を計算
    #   _post_process_late_then_early    遅出翌日早出の後処理
    #   _post_process_balance_early_late 正社員の早出・遅出比率均等化
    #
    # 注: これらは全て _generate_gh のクロージャ内で定義されており、
    # staff / req_map / fixed_map / unit_req / operational_units / last_day /
    # week_ranges / skill_map / locked_days などの変数を直接参照する。
    # 将来の切り出し時はこれらを明示的な引数として渡す必要がある。
    import random
    import time
    import math

    MAX_TRIALS = 1200
    NO_IMPROVE_LIMIT = 220
    TIME_LIMIT_SEC = 3.0

    def _entries_to_sched(ents):
        sc, pl = {}, {}
        for e in ents:
            d = int(e["date"][8:10])
            sc[(e["staff_id"], d)] = e["shift_type"]
            pl[(e["staff_id"], d)] = e["unit_id"]
        return sc, pl

    def _sched_to_entries(sc, pl):
        ents = []
        for sid in staff:
            home_uid = staff[sid]["unit_id"]
            for day in range(1, last_day + 1):
                stype = sc.get((sid, day), "休")
                uid = pl.get((sid, day), home_uid)
                is_cross = 1 if uid != home_uid else 0
                ents.append({
                    "version_id": vid, "staff_id": sid, "unit_id": uid,
                    "date": f"{year}-{month:02d}-{day:02d}",
                    "shift_type": stype, "is_cross": is_cross,
                    "from_unit_id": home_uid if is_cross else None,
                })
        return ents

    def _is_part_or_dispatch(s):
        return s.get("employment_type") in ("part_time", "dispatch")

    def _is_locked(sid, day, shift=None):
        """希望勤務・希望休・曜日固定・明けは通常の反復改善で動かさない。
        lock_prefix（欠勤対応の部分再生成）で固定された日も同様に保護する。"""
        if (sid, day) in locked_days:
            return True
        cur = shift if shift is not None else current_sc.get((sid, day), "休")
        req = req_map.get(sid, {}).get(day)
        expected = {"rest":"希休", "hol":"有給", "absence":"欠勤", "early":"早", "late":"遅",
                    "night":"夜", "day":"日", "ake":"明"}.get(req)
        if expected is not None:
            return True
        if cur == "明":
            return True
        fixed = fixed_map.get(sid, {}).get(dow_of(day))
        if fixed:
            return True
        return False

    def _is_protected(sid, day):
        """希望入力・曜日固定・lock_prefixで固定された日を保護する。

        自動生成された「明」は夜勤と一体で移動可能とする。従来は `_is_locked`
        が全ての明けを固定扱いしたため、夜勤＋明けの直接移管が一度も成立しない
        場合があった。夜勤ペアの系統的調整ではこちらを使う。
        """
        if (sid, day) in locked_days:
            return True
        req = req_map.get(sid, {}).get(day)
        if req is not None:
            return True
        return bool(fixed_map.get(sid, {}).get(dow_of(day)))

    def _capable(sid, shift):
        s = staff[sid]
        if shift == "早": return bool(s.get("can_early")) and not s.get("can_night_only")
        if shift == "遅": return bool(s.get("can_late")) and not s.get("can_night_only")
        if shift == "夜": return bool(s.get("can_night"))
        if shift == "日": return bool(s.get("can_day"))
        return True

    def _adjacent_ok(sc, sid, day, new_shift):
        """局所的な夜勤・明け整合性。4/5連勤は許容し、採点で抑える。"""
        prev_shift = sc.get((sid, day - 1), "休") if day > 1 else "休"
        next_shift = sc.get((sid, day + 1), "休") if day < last_day else "休"
        if new_shift == "早" and prev_shift in ("夜", "明"):
            return False
        if new_shift == "夜":
            if prev_shift in ("夜", "明"):
                return False
            if day < last_day and next_shift not in ("休", "明"):
                return False
        if new_shift == "休" and next_shift == "明":
            # 前日夜勤を外す操作以外で孤立した明けを作らない
            old = sc.get((sid, day), "休")
            if old != "夜":
                return False
        return True

    def _coverage_shortage(sc, pl):
        """早・遅・夜・明の必須配置不足数を返す。

        unit_required_staff には通常「明」の行がないため、明け必要数は
        同ユニットの夜勤必要数と同数として扱う。1日も前月末夜勤の明けを
        必須とし、各ユニットに夜勤必要数分の明けが存在することを確認する。
        """
        shortage = 0
        for u in operational_units:
            uid = u["id"]
            reqs = unit_req.get(uid, {})
            night_required = int(reqs.get("夜", {}).get("required", 1))
            for day in range(1, last_day + 1):
                for slot in ("早", "遅", "夜"):
                    rinfo = reqs.get(slot)
                    if not rinfo:
                        continue
                    required = int(rinfo.get("required", 1))
                    actual = sum(1 for sid in staff
                                 if sc.get((sid, day)) == slot and pl.get((sid, day)) == uid)
                    shortage += max(0, required - actual)

                # 明けは夜勤と同じユニット・同じ必要人数を必須とする
                actual_ake = sum(1 for sid in staff
                                 if sc.get((sid, day)) == "明" and pl.get((sid, day)) == uid)
                shortage += max(0, night_required - actual_ake)
        return shortage

    def _hard_signature(sc, pl):
        """反復改善で絶対に悪化させない必須条件の要約。"""
        shortage = _coverage_shortage(sc, pl)
        overwork = 0
        broken_night_ake = 0
        broken_requests = 0
        broken_fixed = 0
        broken_capability = 0
        broken_unit = 0

        for sid, st in staff.items():
            worked = sum(1 for d in range(1, last_day + 1)
                         if sc.get((sid, d), "休") not in ("休", None))
            overwork += max(0, worked - int(st.get("monthly_limit") or 0))
            for day in range(1, last_day + 1):
                sh = sc.get((sid, day), "休")
                req = req_map.get(sid, {}).get(day)
                expected = {"rest":"希休", "hol":"有給", "absence":"欠勤", "early":"早", "late":"遅",
                            "night":"夜", "day":"日", "ake":"明"}.get(req)
                if expected is not None and sh != expected:
                    broken_requests += 1
                fixed = fixed_map.get(sid, {}).get(dow_of(day))
                if fixed and sh != fixed:
                    broken_fixed += 1
                if sh in ("早", "遅", "夜", "日") and not _capable(sid, sh):
                    broken_capability += 1
                if sh not in ("休", "明") and not _unit_ok(sid, pl.get((sid, day), st["unit_id"])):
                    broken_unit += 1
                if sh == "夜" and day < last_day and sc.get((sid, day + 1)) != "明":
                    broken_night_ake += 1
                if sh == "明" and day > 1 and sc.get((sid, day - 1)) != "夜":
                    broken_night_ake += 1
                if sh == "早" and day > 1 and sc.get((sid, day - 1)) in ("夜", "明"):
                    broken_night_ake += 1

        return (shortage, overwork, broken_night_ake, broken_requests,
                broken_fixed, broken_capability, broken_unit)

    def _score(sc, pl):
        """小さいほど良い。必須条件は極大ペナルティ、次に雇用条件、最後に均等性。"""
        hard = 0.0
        shortage = _coverage_shortage(sc, pl)
        hard += shortage * 100000.0

        contract_penalty = 0.0
        balance_penalty = 0.0
        pattern_penalty = 0.0
        cross_penalty = 0.0
        detail = {"shortage": shortage, "overwork": 0, "contract_gap": 0,
                  "long_streaks": 0, "long_rests": 0}

        for sid, s in staff.items():
            worked_days = [d for d in range(1, last_day + 1)
                           if sc.get((sid, d), "休") not in ("休", None)]
            worked = len(worked_days)
            limit = int(s.get("monthly_limit") or 0)
            over = max(0, worked - limit)
            detail["overwork"] += over
            hard += over * 50000.0

            # 希望・固定・能力・夜→明を必須条件として確認
            for day in range(1, last_day + 1):
                sh = sc.get((sid, day), "休")
                req = req_map.get(sid, {}).get(day)
                expected = {"rest":"希休", "hol":"有給", "absence":"欠勤", "early":"早", "late":"遅",
                            "night":"夜", "day":"日", "ake":"明"}.get(req)
                if expected is not None and sh != expected:
                    hard += 100000.0
                fixed = fixed_map.get(sid, {}).get(dow_of(day))
                if fixed and sh != fixed:
                    hard += 100000.0
                if sh in ("早", "遅", "夜", "日") and not _capable(sid, sh):
                    hard += 100000.0
                if sh not in ("休", "明") and not _unit_ok(sid, pl.get((sid, day), s["unit_id"])):
                    hard += 100000.0
                if sh == "夜" and day < last_day and sc.get((sid, day + 1)) != "明":
                    hard += 100000.0
                if sh == "明" and day > 1 and sc.get((sid, day - 1)) != "夜":
                    # 月初明けは前月夜勤のため例外
                    hard += 100000.0
                if sh == "早" and day > 1 and sc.get((sid, day - 1)) in ("夜", "明"):
                    hard += 100000.0
                if pl.get((sid, day), s["unit_id"]) != s["unit_id"] and sh not in ("休", "明"):
                    cross_penalty += 0.5

            # パート・派遣は契約日数を必達に近い強い評価。
            # 正社員は不足を埋めるための日勤を作らないため、未達は罰しない。
            if _is_part_or_dispatch(s):
                # パート・派遣のmonthly_limitは契約勤務日数として扱い、
                # 不足・超過のどちらも全職員共通の基準で強く評価する。
                contract_diff = abs(limit - worked)
                detail["contract_gap"] += contract_diff
                contract_penalty += contract_diff * 1800.0

                # 月初詰め・月末連休の根本原因を抑えるため、月途中の累積勤務数を評価する。
                cumulative = 0
                for d in range(1, last_day + 1):
                    if sc.get((sid, d), "休") not in ("休", None):
                        cumulative += 1
                    expected_cumulative = limit * d / last_day
                    balance_penalty += (cumulative - expected_cumulative) ** 2 * 1.8

            # シフト別目標（特に夜勤専従）
            for stype, tinfo in s.get("shift_targets", {}).items():
                target = int(tinfo.get("target") or 0)
                actual = sum(1 for d in range(1, last_day + 1) if sc.get((sid, d)) == stype)
                if tinfo.get("is_upper"):
                    contract_penalty += max(0, actual - target) * 500.0
                else:
                    contract_penalty += abs(actual - target) * (120.0 if stype == "夜" else 60.0)

            # 4区分の勤務数を、区分日数に比例した期待値へ近づける
            lengths = [b - a + 1 for a, b in week_ranges]
            total_len = sum(lengths) or 1
            week_counts = []
            for a, b in week_ranges:
                week_counts.append(sum(1 for d in range(a, b + 1)
                                       if sc.get((sid, d), "休") not in ("休", None)))
            expected_counts = [worked * ln / total_len for ln in lengths]
            balance_penalty += sum((c - e) ** 2 for c, e in zip(week_counts, expected_counts)) * 6.0

            # 連勤・連休。4連勤/5連勤は禁止せず段階的に減点する。
            work_run = rest_run = 0
            for day in range(1, last_day + 1):
                sh = sc.get((sid, day), "休")
                if sh not in ("休", None):
                    work_run += 1
                    rest_run = 0
                    if work_run == 4: pattern_penalty += 6.0
                    elif work_run == 5: pattern_penalty += 20.0; detail["long_streaks"] += 1
                    elif work_run >= 6: pattern_penalty += 80.0; detail["long_streaks"] += 1
                else:
                    rest_run += 1
                    work_run = 0
                    # 連休の許容幅は契約日数に応じて変える。
                    target_days = max(1, limit)
                    natural_gap = max(2, int(round(last_day / target_days)))
                    soft_limit = natural_gap + 1
                    if rest_run > soft_limit:
                        excess_rest = rest_run - soft_limit
                        pattern_penalty += excess_rest * excess_rest * 4.0
                        if excess_rest >= 2:
                            detail["long_rests"] += 1

            # 月末7日だけ極端に休みが多い場合を軽く減点
            tail_start = max(1, last_day - 6)
            tail_work = sum(1 for d in range(tail_start, last_day + 1)
                            if sc.get((sid, d), "休") not in ("休", None))
            expected_tail = worked * (last_day - tail_start + 1) / last_day
            if tail_work < expected_tail:
                balance_penalty += (expected_tail - tail_work) ** 2 * 4.0

        total = hard + contract_penalty + balance_penalty + pattern_penalty + cross_penalty
        detail.update({
            "hard": round(hard, 3), "contract": round(contract_penalty, 3),
            "balance": round(balance_penalty, 3), "pattern": round(pattern_penalty, 3),
            "cross": round(cross_penalty, 3), "total": round(total, 3),
        })
        return total, detail

    def _try_transfer(sc, pl, rng):
        """同じ日・同じ必要枠をAからBへ移す。必要配置数は変わらない。"""
        day = rng.randint(1, last_day)
        slot = rng.choice(("早", "遅", "夜"))
        donors = [sid for sid in regular_staff if sc.get((sid, day)) == slot]
        if not donors:
            return None
        def donor_key(sid):
            s = staff[sid]
            worked = _worked_count(sc, sid)
            limit = int(s.get("monthly_limit") or 0)
            over = max(0, worked - limit)
            # 超過者、次に正社員、週内過多・連勤が長い人から勤務を外す
            widx = week_index(day)
            week_work = sum(1 for d in range(week_ranges[widx][0], week_ranges[widx][1] + 1)
                            if sc.get((sid, d), "休") not in ("休", None))
            return (-over, 0 if not _is_part_or_dispatch(s) else 1, -week_work, -_work_run_before(sc, sid, day), rng.random())
        donors.sort(key=donor_key)
        donor = rng.choice(donors[:min(4, len(donors))])
        uid = pl.get((donor, day), staff[donor]["unit_id"])
        if _is_locked(donor, day, slot):
            return None

        receivers = []
        for sid, s in regular_staff.items():
            if sid == donor or sc.get((sid, day), "休") != "休":
                continue
            if _is_locked(sid, day, "休") or not _capable(sid, slot) or not _unit_ok(sid, uid):
                continue
            if not _adjacent_ok(sc, sid, day, slot):
                continue
            if slot == "夜" and day < last_day:
                if sc.get((sid, day + 1), "休") != "休":
                    continue
                if _is_locked(sid, day + 1, "休"):
                    continue
                if sc.get((donor, day + 1)) != "明" or _is_locked(donor, day + 1, "明"):
                    continue
            receivers.append(sid)
        if not receivers:
            return None

        # 契約未達のパート・派遣、次に週内不足の大きい人を優先しつつランダム性を残す
        def receiver_key(sid):
            s = staff[sid]
            worked = sum(1 for d in range(1, last_day + 1) if sc.get((sid, d)) not in (None, "休"))
            gap = max(0, int(s.get("monthly_limit") or 0) - worked)
            widx = week_index(day)
            wc = sum(1 for d in range(week_ranges[widx][0], week_ranges[widx][1] + 1)
                     if sc.get((sid, d)) not in (None, "休"))
            rest_run = _rest_run_around(sc, sid, day)
            expected = int(s.get("monthly_limit") or 0) * ((week_ranges[widx][1] - week_ranges[widx][0] + 1) / last_day)
            week_gap = expected - wc
            return (0 if _is_part_or_dispatch(s) and gap > 0 else 1,
                    0 if gap > 0 else 1, -gap, -rest_run, -week_gap, wc, rng.random())
        receivers.sort(key=receiver_key)
        receiver = rng.choice(receivers[:min(3, len(receivers))])

        nsc, npl = dict(sc), dict(pl)
        nsc[(donor, day)] = "休"
        npl[(donor, day)] = staff[donor]["unit_id"]
        nsc[(receiver, day)] = slot
        npl[(receiver, day)] = uid
        if slot == "夜" and day < last_day:
            nsc[(donor, day + 1)] = "休"
            npl[(donor, day + 1)] = staff[donor]["unit_id"]
            nsc[(receiver, day + 1)] = "明"
            npl[(receiver, day + 1)] = uid
        return nsc, npl, {"type":"transfer", "day":day, "slot":slot,
                           "from":donor, "to":receiver}

    def _try_cross_day_swap(sc, pl, rng):
        """AとBの同一シフト勤務日を交換し、各人の総勤務日数と必要配置を維持する。"""
        slot = rng.choice(("早", "遅"))
        d1, d2 = rng.sample(range(1, last_day + 1), 2)
        a_list = [sid for sid in regular_staff
                  if sc.get((sid, d1)) == slot and sc.get((sid, d2), "休") == "休"]
        if not a_list:
            return None
        rng.shuffle(a_list)
        for a in a_list[:5]:
            if _is_locked(a, d1, slot) or _is_locked(a, d2, "休"):
                continue
            b_list = [sid for sid in regular_staff
                      if sid != a and sc.get((sid, d1), "休") == "休" and sc.get((sid, d2)) == slot]
            rng.shuffle(b_list)
            for b in b_list[:8]:
                if _is_locked(b, d1, "休") or _is_locked(b, d2, slot):
                    continue
                uid1 = pl.get((a, d1), staff[a]["unit_id"])
                uid2 = pl.get((b, d2), staff[b]["unit_id"])
                if not (_capable(a, slot) and _capable(b, slot)
                        and _unit_ok(a, uid2) and _unit_ok(b, uid1)):
                    continue
                if not (_adjacent_ok(sc, a, d2, slot) and _adjacent_ok(sc, b, d1, slot)):
                    continue
                nsc, npl = dict(sc), dict(pl)
                nsc[(a, d1)] = "休"; npl[(a, d1)] = staff[a]["unit_id"]
                nsc[(b, d2)] = "休"; npl[(b, d2)] = staff[b]["unit_id"]
                nsc[(a, d2)] = slot; npl[(a, d2)] = uid2
                nsc[(b, d1)] = slot; npl[(b, d1)] = uid1
                return nsc, npl, {"type":"cross_day", "slot":slot,
                                   "a":a, "b":b, "d1":d1, "d2":d2}
        return None

    def _worked_count(sc, sid):
        return sum(1 for d in range(1, last_day + 1)
                   if sc.get((sid, d), "休") not in ("休", None))

    def _work_run_before(sc, sid, day):
        run = 0
        for d in range(day - 1, 0, -1):
            if sc.get((sid, d), "休") not in ("休", None):
                run += 1
            else:
                break
        return run

    def _rest_run_around(sc, sid, day):
        """dayを含む連休の長さ。長い連休の中央を優先して勤務へ変えるために使う。"""
        if sc.get((sid, day), "休") != "休":
            return 0
        run = 1
        d = day - 1
        while d >= 1 and sc.get((sid, d), "休") == "休":
            run += 1; d -= 1
        d = day + 1
        while d <= last_day and sc.get((sid, d), "休") == "休":
            run += 1; d += 1
        return run

    def _missing_slots(sc, pl):
        missing = []
        for u in operational_units:
            uid = u["id"]
            reqs = unit_req.get(uid, {})
            night_required = int(reqs.get("夜", {}).get("required", 1))
            for day in range(1, last_day + 1):
                for slot in ("夜", "早", "遅"):
                    rinfo = reqs.get(slot)
                    if not rinfo:
                        continue
                    required = int(rinfo.get("required", 1))
                    actual = sum(1 for sid in staff
                                 if sc.get((sid, day)) == slot and pl.get((sid, day)) == uid)
                    for _ in range(max(0, required - actual)):
                        missing.append((uid, day, slot))
                actual_ake = sum(1 for sid in staff
                                 if sc.get((sid, day)) == "明" and pl.get((sid, day)) == uid)
                for _ in range(max(0, night_required - actual_ake)):
                    missing.append((uid, day, "明"))
        return missing

    def _repair_coverage(sc, pl, max_passes=8):
        """生成後の不足枠を補完する。

        夜勤は翌日の明けと一体で追加する。明け不足は前日の夜勤者を最優先で復元する。
        契約未達のパート・派遣、長い連休中、週内勤務不足の順で候補を選ぶ。
        """
        sc, pl = dict(sc), dict(pl)
        repaired = []
        for _pass in range(max_passes):
            changed = False
            missing = _missing_slots(sc, pl)
            if not missing:
                break

            # 明けは夜勤との対応復元を先に行う
            missing.sort(key=lambda x: (0 if x[2] == "明" else 1 if x[2] == "夜" else 2, x[1], x[0]))
            for uid, day, slot in missing:
                # 既に別の修復で埋まっていればスキップ
                required = (int(unit_req.get(uid, {}).get("夜", {}).get("required", 1))
                            if slot == "明" else
                            int(unit_req.get(uid, {}).get(slot, {}).get("required", 1)))
                actual = sum(1 for sid in staff
                             if sc.get((sid, day)) == slot and pl.get((sid, day)) == uid)
                if actual >= required:
                    continue

                if slot == "明":
                    if day > 1:
                        prev_night = [sid for sid in regular_staff
                                      if sc.get((sid, day - 1)) == "夜"
                                      and pl.get((sid, day - 1)) == uid
                                      and sc.get((sid, day), "休") == "休"
                                      and not _is_locked(sid, day, "休")]
                        if prev_night:
                            sid = prev_night[0]
                            sc[(sid, day)] = "明"; pl[(sid, day)] = uid
                            repaired.append({"type":"restore_ake","sid":sid,"day":day,"unit":uid})
                            changed = True
                            continue
                    # 月初明けは前月夜勤扱い。夜勤可能者の休から補完する。
                    if day == 1:
                        cands = [sid for sid in regular_staff
                                 if staff[sid].get("can_night")
                                 and sc.get((sid, day), "休") == "休"
                                 and not _is_locked(sid, day, "休")
                                 and _unit_ok(sid, uid)]
                        if cands:
                            cands.sort(key=lambda sid: (_worked_count(sc, sid), sid))
                            sid = cands[0]
                            sc[(sid, day)] = "明"; pl[(sid, day)] = uid
                            repaired.append({"type":"month_start_ake","sid":sid,"day":day,"unit":uid})
                            changed = True
                    continue

                candidates = []
                for sid, s in regular_staff.items():
                    if sc.get((sid, day), "休") != "休":
                        continue
                    if _is_locked(sid, day, "休") or not _capable(sid, slot) or not _unit_ok(sid, uid):
                        continue
                    if not _adjacent_ok(sc, sid, day, slot):
                        continue
                    if slot == "夜" and day < last_day:
                        if sc.get((sid, day + 1), "休") != "休" or _is_locked(sid, day + 1, "休"):
                            continue
                    worked = _worked_count(sc, sid)
                    limit = int(s.get("monthly_limit") or 0)
                    gap = limit - worked
                    is_hourly = _is_part_or_dispatch(s)
                    widx = week_index(day)
                    week_work = sum(1 for d in range(week_ranges[widx][0], week_ranges[widx][1] + 1)
                                    if sc.get((sid, d), "休") not in ("休", None))
                    expected_week = limit * ((week_ranges[widx][1] - week_ranges[widx][0] + 1) / last_day)
                    week_gap = expected_week - week_work
                    candidates.append((
                        0 if is_hourly and gap > 0 else 1,
                        0 if gap > 0 else 1,
                        -max(gap, 0),
                        -_rest_run_around(sc, sid, day),
                        -week_gap,
                        _work_run_before(sc, sid, day),
                        0 if s["unit_id"] == uid else 1,
                        worked, sid
                    ))
                if not candidates:
                    continue
                candidates.sort()
                sid = candidates[0][-1]
                sc[(sid, day)] = slot; pl[(sid, day)] = uid
                if slot == "夜" and day < last_day:
                    sc[(sid, day + 1)] = "明"; pl[(sid, day + 1)] = uid
                repaired.append({"type":"fill_slot","sid":sid,"day":day,"slot":slot,"unit":uid})
                changed = True

            if not changed:
                break
        return sc, pl, repaired, _missing_slots(sc, pl)

    def _systematic_contract_rebalance(sc, pl, max_passes=10):
        """契約日数・シフト目標の差を、同一枠の直接移管で解消する。

        特定職員名には依存しない。まず夜勤専従・夜勤目標者について、
        夜勤＋翌日明けを一組として、目標超過者または正社員の夜勤から
        目標未達者へ移す。月末夜勤は1日分として扱う。
        次に早・遅の通常勤務を、契約超過者から契約未達者へ移す。
        変更前後で必須配置数は変えない。
        """
        sc, pl = dict(sc), dict(pl)
        changes = []

        def worked_map():
            return {sid: _worked_count(sc, sid) for sid in regular_staff}

        def night_count(sid):
            return sum(1 for d in range(1, last_day + 1) if sc.get((sid, d)) == "夜")

        def target_nights(sid):
            st = staff[sid]
            t = int(st.get("night_target") or 0)
            if t > 0:
                return t
            info = st.get("shift_targets", {}).get("夜")
            return int(info.get("target") or 0) if info else 0

        # Phase 1: 夜勤目標の不足を、同じ日の夜勤枠そのものを移して解消する。
        for _ in range(max_passes):
            worked = worked_map()
            under = [sid for sid in regular_staff
                     if staff[sid].get("can_night")
                     and target_nights(sid) > night_count(sid)]
            if not under:
                break
            under.sort(key=lambda sid: (-(target_nights(sid) - night_count(sid)), sid))
            moved = False

            for receiver in under:
                rec_limit = int(staff[receiver].get("monthly_limit") or 0)
                rec_gap = rec_limit - worked[receiver]
                if rec_gap <= 0:
                    continue

                best = None
                for day in range(1, last_day + 1):
                    gain = 1 if day == last_day else 2
                    if gain > rec_gap:
                        continue
                    if sc.get((receiver, day), "休") != "休" or _is_locked(receiver, day, "休"):
                        continue
                    if not _adjacent_ok(sc, receiver, day, "夜"):
                        continue
                    if day < last_day:
                        if sc.get((receiver, day + 1), "休") != "休" or _is_locked(receiver, day + 1, "休"):
                            continue

                    for donor in regular_staff:
                        if donor == receiver or sc.get((donor, day)) != "夜" or _is_locked(donor, day, "夜"):
                            continue
                        uid = pl.get((donor, day), staff[donor]["unit_id"])
                        if not _unit_ok(receiver, uid):
                            continue
                        if day < last_day and (sc.get((donor, day + 1)) != "明" or _is_protected(donor, day + 1)):
                            continue

                        donor_target = target_nights(donor)
                        donor_nights = night_count(donor)
                        donor_limit = int(staff[donor].get("monthly_limit") or 0)
                        donor_over_days = worked[donor] - donor_limit
                        donor_surplus_nights = donor_nights - donor_target if donor_target > 0 else 0

                        # パート・派遣からは契約超過または夜勤目標超過の時だけ移す。
                        # 正社員からは月間勤務を下限扱いしないため移管可能。
                        if _is_part_or_dispatch(staff[donor]) and donor_over_days < gain and donor_surplus_nights <= 0:
                            continue

                        # 受取側の長期連休を分断し、月末夜勤1日で奇数差も埋める。
                        key = (
                            0 if donor_over_days >= gain else 1,
                            0 if donor_surplus_nights > 0 else 1,
                            0 if day == last_day and rec_gap % 2 == 1 else 1,
                            -_rest_run_around(sc, receiver, day),
                            abs(day - (last_day * (night_count(receiver) + 1) / max(target_nights(receiver), 1))),
                            day,
                            donor,
                        )
                        if best is None or key < best[0]:
                            best = (key, donor, day, uid, gain)

                if best is None:
                    continue
                _, donor, day, uid, gain = best
                sc[(donor, day)] = "休"; pl[(donor, day)] = staff[donor]["unit_id"]
                sc[(receiver, day)] = "夜"; pl[(receiver, day)] = uid
                if day < last_day:
                    sc[(donor, day + 1)] = "休"; pl[(donor, day + 1)] = staff[donor]["unit_id"]
                    sc[(receiver, day + 1)] = "明"; pl[(receiver, day + 1)] = uid
                changes.append({"type":"night_pair", "from":donor, "to":receiver,
                                "day":day, "slot":"夜", "unit":uid, "gain_days":gain})
                moved = True
                break
            if not moved:
                break

        # Phase 2: 契約日数不足者へ早・遅を直接移管する。
        for _ in range(max_passes):
            worked = worked_map()
            under = [sid for sid, st in regular_staff.items()
                     if _is_part_or_dispatch(st)
                     and worked[sid] < int(st.get("monthly_limit") or 0)]
            if not under:
                break
            under.sort(key=lambda sid: (-(int(staff[sid].get("monthly_limit") or 0) - worked[sid]), sid))
            moved = False
            for receiver in under:
                best = None
                for day in range(1, last_day + 1):
                    if sc.get((receiver, day), "休") != "休" or _is_locked(receiver, day, "休"):
                        continue
                    for slot in ("早", "遅"):
                        if not _capable(receiver, slot) or not _adjacent_ok(sc, receiver, day, slot):
                            continue
                        for donor in regular_staff:
                            if donor == receiver or sc.get((donor, day)) != slot or _is_locked(donor, day, slot):
                                continue
                            uid = pl.get((donor, day), staff[donor]["unit_id"])
                            if not _unit_ok(receiver, uid):
                                continue
                            donor_limit = int(staff[donor].get("monthly_limit") or 0)
                            donor_over = worked[donor] - donor_limit
                            if _is_part_or_dispatch(staff[donor]) and donor_over <= 0:
                                continue
                            key = (0 if donor_over > 0 else 1, -donor_over,
                                   -_rest_run_around(sc, receiver, day), day, donor)
                            if best is None or key < best[0]:
                                best = (key, donor, day, slot, uid)
                if best is None:
                    continue
                _, donor, day, slot, uid = best
                sc[(donor, day)] = "休"; pl[(donor, day)] = staff[donor]["unit_id"]
                sc[(receiver, day)] = slot; pl[(receiver, day)] = uid
                changes.append({"type":"day_transfer", "from":donor, "to":receiver,
                                "day":day, "slot":slot, "unit":uid, "gain_days":1})
                moved = True
                break
            if not moved:
                break
        return sc, pl, changes

    def _systematic_night_pattern_balance(sc, pl, max_passes=6, max_checks=240):
        """夜勤＋明けの配置日を職員間で交換し、長期連休を分断する。

        各職員の勤務日数・夜勤回数、各日の必要配置は一切変えず、夜勤ペアの
        日付だけを交換する。希望入力・曜日固定・月初明けは動かさない。
        """
        sc, pl = dict(sc), dict(pl)
        changes = []
        for _ in range(max_passes):
            base_score, _ = _score(sc, pl)
            best = None
            best_score = base_score
            checks = 0

            # 長い連休を持つ夜勤可能者を優先する。
            def longest_rest(sid):
                longest = run = 0
                for d in range(1, last_day + 1):
                    if sc.get((sid, d), "休") == "休":
                        run += 1; longest = max(longest, run)
                    else:
                        run = 0
                return longest

            focus = [sid for sid in regular_staff if staff[sid].get("can_night")]
            focus.sort(key=lambda sid: (-longest_rest(sid), sid))
            focus = focus[:16]

            for a in focus:
                a_nights = [d for d in range(1, last_day)
                            if sc.get((a, d)) == "夜" and sc.get((a, d + 1)) == "明"
                            and not _is_protected(a, d) and not _is_protected(a, d + 1)]
                # a の長期連休の中央付近にある休・休を優先する。
                a_rest_starts = [d for d in range(1, last_day)
                                 if sc.get((a, d), "休") == "休"
                                 and sc.get((a, d + 1), "休") == "休"
                                 and not _is_protected(a, d) and not _is_protected(a, d + 1)]
                a_rest_starts.sort(key=lambda d: (-(_rest_run_around(sc, a, d) + _rest_run_around(sc, a, d + 1)), d))
                for d_old in a_nights[:10]:
                    for d_new in a_rest_starts[:14]:
                        if d_old == d_new:
                            continue
                        for b in regular_staff:
                            if b == a or not staff[b].get("can_night"):
                                continue
                            if sc.get((b, d_old), "休") != "休" or sc.get((b, d_old + 1), "休") != "休":
                                continue
                            if sc.get((b, d_new)) != "夜" or sc.get((b, d_new + 1)) != "明":
                                continue
                            if (_is_protected(b, d_old) or _is_protected(b, d_old + 1)
                                    or _is_protected(b, d_new) or _is_protected(b, d_new + 1)):
                                continue
                            uid_old = pl.get((a, d_old), staff[a]["unit_id"])
                            uid_new = pl.get((b, d_new), staff[b]["unit_id"])
                            if not (_unit_ok(a, uid_new) and _unit_ok(b, uid_old)):
                                continue

                            nsc, npl = dict(sc), dict(pl)
                            # a の旧ペアを b へ、b の旧ペアを a へ交換
                            nsc[(a, d_old)] = "休"; npl[(a, d_old)] = staff[a]["unit_id"]
                            nsc[(a, d_old + 1)] = "休"; npl[(a, d_old + 1)] = staff[a]["unit_id"]
                            nsc[(b, d_new)] = "休"; npl[(b, d_new)] = staff[b]["unit_id"]
                            nsc[(b, d_new + 1)] = "休"; npl[(b, d_new + 1)] = staff[b]["unit_id"]
                            nsc[(a, d_new)] = "夜"; npl[(a, d_new)] = uid_new
                            nsc[(a, d_new + 1)] = "明"; npl[(a, d_new + 1)] = uid_new
                            nsc[(b, d_old)] = "夜"; npl[(b, d_old)] = uid_old
                            nsc[(b, d_old + 1)] = "明"; npl[(b, d_old + 1)] = uid_old

                            if _hard_signature(nsc, npl) != _hard_signature(sc, pl):
                                continue
                            score, _ = _score(nsc, npl)
                            checks += 1
                            if score < best_score - 1e-6:
                                best_score = score
                                best = (nsc, npl, {"type":"night_date_swap", "a":a, "b":b,
                                                   "a_from":d_old, "a_to":d_new})
                            if checks >= max_checks:
                                break
                        if checks >= max_checks:
                            break
                    if checks >= max_checks:
                        break
                if checks >= max_checks:
                    break
            if best is None:
                break
            sc, pl, info = best
            changes.append(info)
        return sc, pl, changes

    def _systematic_pattern_balance(sc, pl, max_passes=4, max_checks=160):
        """同一シフトの日付交換で、勤務日数と必要配置を保ったまま月内偏りを改善する。

        全組合せ探索はPC負荷が大きいため、偏りの大きい職員・日だけに候補を絞る。
        """
        sc, pl = dict(sc), dict(pl)
        changes = []
        for _ in range(max_passes):
            base_score, _ = _score(sc, pl)
            best = None
            best_score = base_score
            checks = 0

            # 月後半に休みが偏る職員、または前半に休みが偏る職員を優先対象にする。
            def imbalance(sid):
                counts = []
                for a, b in week_ranges:
                    counts.append(sum(1 for d in range(a, b + 1)
                                      if sc.get((sid, d), "休") not in ("休", None)))
                if not counts:
                    return 0
                return max(counts) - min(counts)

            focus_staff = sorted(regular_staff, key=lambda sid: (-imbalance(sid), sid))[:14]
            for slot in ("早", "遅"):
                for a in focus_staff:
                    work_days = [d for d in range(1, last_day + 1)
                                 if sc.get((a, d)) == slot and not _is_locked(a, d, slot)]
                    rest_days = [d for d in range(1, last_day + 1)
                                 if sc.get((a, d), "休") == "休" and not _is_locked(a, d, "休")]
                    # 長い連休内の休みを先に候補にする。
                    rest_days.sort(key=lambda d: (-_rest_run_around(sc, a, d), d))
                    for d1 in work_days[:10]:
                        for d2 in rest_days[:12]:
                            if d1 == d2:
                                continue
                            for b in regular_staff:
                                if b == a or sc.get((b,d1), "休") != "休" or sc.get((b,d2)) != slot:
                                    continue
                                if _is_locked(b,d1,"休") or _is_locked(b,d2,slot):
                                    continue
                                uid1 = pl.get((a,d1), staff[a]["unit_id"])
                                uid2 = pl.get((b,d2), staff[b]["unit_id"])
                                if not (_unit_ok(a,uid2) and _unit_ok(b,uid1)):
                                    continue
                                if not (_adjacent_ok(sc,a,d2,slot) and _adjacent_ok(sc,b,d1,slot)):
                                    continue
                                nsc, npl = dict(sc), dict(pl)
                                nsc[(a,d1)]="休"; npl[(a,d1)]=staff[a]["unit_id"]
                                nsc[(b,d2)]="休"; npl[(b,d2)]=staff[b]["unit_id"]
                                nsc[(a,d2)]=slot; npl[(a,d2)]=uid2
                                nsc[(b,d1)]=slot; npl[(b,d1)]=uid1
                                if _hard_signature(nsc,npl) != _hard_signature(sc,pl):
                                    continue
                                score,_ = _score(nsc,npl)
                                checks += 1
                                if score < best_score - 1e-6:
                                    best_score = score
                                    best = (nsc,npl,{"a":a,"b":b,"d1":d1,"d2":d2,"slot":slot})
                                if checks >= max_checks:
                                    break
                            if checks >= max_checks:
                                break
                        if checks >= max_checks:
                            break
                    if checks >= max_checks:
                        break
                if checks >= max_checks:
                    break
            if best is None:
                break
            sc, pl, info = best
            changes.append(info)
        return sc, pl, changes

    initial_sc, initial_pl = _entries_to_sched(entries)
    # 初期案に不足がある場合は、反復改善の前に必ず補完する。
    initial_sc, initial_pl, initial_repairs, initial_unresolved = _repair_coverage(initial_sc, initial_pl)
    initial_sc, initial_pl, initial_contract_rebalances = _systematic_contract_rebalance(initial_sc, initial_pl)
    initial_sc, initial_pl, initial_night_pattern_rebalances = _systematic_night_pattern_balance(initial_sc, initial_pl)
    initial_sc, initial_pl, initial_pattern_rebalances = _systematic_pattern_balance(initial_sc, initial_pl)
    initial_sc, initial_pl, post_balance_repairs, initial_unresolved = _repair_coverage(initial_sc, initial_pl)
    current_sc, current_pl = dict(initial_sc), dict(initial_pl)
    current_score, current_detail = _score(current_sc, current_pl)
    initial_hard_signature = _hard_signature(current_sc, current_pl)
    current_hard_signature = initial_hard_signature
    best_sc, best_pl = dict(current_sc), dict(current_pl)
    best_score, best_detail = current_score, dict(current_detail)
    best_hard_signature = initial_hard_signature

    # ── 複数シードで反復改善を実行し最高スコアを採用 ──────────────
    # N_RUNS回、異なるシードで生成して最良結果を採用する。
    # ベースシードに現在時刻(ns)を混ぜることで、同じバージョンで
    # 何度自動生成しても毎回異なる結果が得られる。
    import time as _time
    N_RUNS = 4          # 試行回数
    TIME_PER_RUN = 2.5  # 1回あたりの時間制限（秒）
    MAX_TRIALS_PER_RUN = 800
    NO_IMPROVE_LIMIT_PER_RUN = 180

    # 現在時刻を混ぜることで毎回異なるシードになる
    time_factor = int(_time.time_ns() % (10**9))
    base_seed = ((fid << 16) ^ (year << 8) ^ month ^ time_factor) & 0x7FFFFFFF

    global_best_sc = dict(current_sc)
    global_best_pl = dict(current_pl)
    global_best_score = current_score
    global_best_detail = dict(current_detail)
    all_logs = []

    for run_i in range(N_RUNS):
        # 各runで異なるシード（base_seed + run番号）
        seed = base_seed + run_i * 997
        rng = random.Random(seed)
        started = _time.perf_counter()
        no_improve = 0
        accepted = 0
        improved_count = 0
        trials = 0
        log = []

        # 各runの初期値はglobal_bestから出発（前runの成果を引き継ぐ）
        current_sc2 = dict(global_best_sc)
        current_pl2 = dict(global_best_pl)
        current_score2, current_detail2 = _score(current_sc2, current_pl2)
        current_hard_signature2 = _hard_signature(current_sc2, current_pl2)
        run_best_sc = dict(current_sc2)
        run_best_pl = dict(current_pl2)
        run_best_score = current_score2
        run_best_detail = dict(current_detail2)
        run_best_hard = current_hard_signature2

        while trials < MAX_TRIALS_PER_RUN and no_improve < NO_IMPROVE_LIMIT_PER_RUN:
            if _time.perf_counter() - started >= TIME_PER_RUN:
                break
            trials += 1
            move = _try_transfer(current_sc2, current_pl2, rng) if rng.random() < 0.58 \
                   else _try_cross_day_swap(current_sc2, current_pl2, rng)
            if move is None:
                no_improve += 1
                continue
            cand_sc, cand_pl, move_info = move
            cand_hard_signature = _hard_signature(cand_sc, cand_pl)

            if any(c > b for c, b in zip(cand_hard_signature, current_hard_signature2)):
                no_improve += 1
                continue

            cand_score, cand_detail = _score(cand_sc, cand_pl)

            elapsed_ratio = min(1.0, trials / MAX_TRIALS_PER_RUN)
            temperature = max(0.05, 4.0 * (1.0 - elapsed_ratio))
            delta = cand_score - current_score2
            accept = delta < -1e-9
            if not accept and cand_detail["hard"] <= current_detail2["hard"] and delta < 20.0:
                accept = rng.random() < math.exp(-max(0.0, delta) / temperature) * 0.03

            if accept:
                current_sc2, current_pl2 = cand_sc, cand_pl
                current_score2, current_detail2 = cand_score, cand_detail
                current_hard_signature2 = cand_hard_signature
                accepted += 1

            if (all(c <= b for c, b in zip(cand_hard_signature, initial_hard_signature))
                    and cand_score < run_best_score - 1e-9):
                run_best_sc, run_best_pl = dict(cand_sc), dict(cand_pl)
                run_best_score, run_best_detail = cand_score, dict(cand_detail)
                run_best_hard = cand_hard_signature
                improved_count += 1
                no_improve = 0
                if len(log) < 50:
                    log.append({"run":run_i, "trial":trials, "score":round(run_best_score, 3)})
            else:
                no_improve += 1

        # このrunの最良をglobal_bestと比較
        if run_best_score < global_best_score - 1e-9:
            global_best_sc = dict(run_best_sc)
            global_best_pl = dict(run_best_pl)
            global_best_score = run_best_score
            global_best_detail = dict(run_best_detail)
        all_logs.extend(log)

    # グローバル最良を採用
    best_sc, best_pl = global_best_sc, global_best_pl
    best_score, best_detail = global_best_score, global_best_detail
    log = all_logs
    seed = base_seed  # ログ記録用

    # 最終段階でも一般則による契約再配分と勤休平準化を行う。
    best_sc, best_pl, final_contract_rebalances = _systematic_contract_rebalance(best_sc, best_pl)
    best_sc, best_pl, final_night_pattern_rebalances = _systematic_night_pattern_balance(best_sc, best_pl)
    best_sc, best_pl, final_pattern_rebalances = _systematic_pattern_balance(best_sc, best_pl)
    best_sc, best_pl, final_repairs, final_unresolved = _repair_coverage(best_sc, best_pl)
    final_hard_signature = _hard_signature(best_sc, best_pl)
    fallback_to_initial = False
    if final_hard_signature[0] > 0:
        # 最良案の補完に失敗した場合は、補完済み初期案へ戻す。
        best_sc, best_pl = dict(initial_sc), dict(initial_pl)
        best_score, best_detail = _score(best_sc, best_pl)
        final_hard_signature = _hard_signature(best_sc, best_pl)
        fallback_to_initial = True
    else:
        best_score, best_detail = _score(best_sc, best_pl)

    def _compute_unit_surplus(entries, last_day):
        """ユニット毎の人員余剰（経費削減余地）を可視化する。

        修正仕様（2026-06）:
          「ユニット毎のスタッフ数の余剰がある。余剰は空きと考える。
           空きは余剰で経費削減に繋がる。余剰資源の見える化に繋げる。
           社員の余剰はそのユニットの余剰とする。」

        余剰日数 ＝ Σ(所属職員のmonthly_limit) − Σ(実際の配置勤務日数)
                    − Σ(希望休・有給による意図的な休み日数)
        希望休・有給は「本人都合の不在」であって人員余剰ではないため除外する。
        正の値が大きいほど、そのユニットは契約上の稼働力に対して実配置が
        少ない＝人員に余剰があり、コスト最適化（増員抑制・配置転換等）の
        余地があることを示す参考指標。診断目的であり、自動生成の必須条件・
        スコアには一切影響しない（評価結果に追加情報として付与するのみ）。
        """
        worked_days = defaultdict(int)
        for e in entries:
            if e["shift_type"] not in ("休", None):
                worked_days[e["staff_id"]] += 1

        surplus_by_unit = {}
        for u in units_list:
            uid = u["id"]
            unit_staff = [s for s in staff.values() if s["unit_id"] == uid]
            total_contract = sum(int(s.get("monthly_limit") or 0) for s in unit_staff)
            total_worked = sum(worked_days.get(s["id"], 0) for s in unit_staff)
            intentional_off = 0
            for s in unit_staff:
                for day in range(1, last_day + 1):
                    if req_map.get(s["id"], {}).get(day) in ("rest", "hol"):
                        intentional_off += 1
            surplus_days = max(0, total_contract - total_worked - intentional_off)
            surplus_by_unit[str(uid)] = {
                "unit_name": u.get("name"),
                "staff_count": len(unit_staff),
                "contract_days_total": total_contract,
                "worked_days_total": total_worked,
                "intentional_off_days": intentional_off,
                "surplus_days": surplus_days,
                "surplus_fte": round(surplus_days / max(last_day, 1), 2),
            }
        return surplus_by_unit

    def _post_process_late_then_early(sc, pl):
        """安定版（STEP1〜8）の完成結果に対し、遅出翌日早出パターンのみを
        個別に検査し、すべての安全条件を満たす場合に限りスワップで解消する。

        レポート「修正版シフト生成結果の評価」4節の採用条件に準拠:
          ・必須条件（shortage/overwork/希望/固定/能力/所属ユニット/夜→明）を
            1件も悪化させない（_hard_signatureで判定）
          ・契約日数(contract)・週次/前後半バランス(balance)・連勤連休(pattern)・
            他ユニット勤務(cross)のいずれも悪化させない（_score()のdetail単位で判定。
            _score合計自体には遅出翌日早出の項目がないため、合計の比較ではなく
            各項目を個別に比較する）
          ・交代先は「既に休み」の職員に限定し、他職員の既存シフトは動かさない
            （連鎖的な再配置をしないことで、ランキング/スコア関数自体を
            変更した場合のような波及を防ぐ）
          ・希望休・希望勤務・曜日固定のある日は、本人・交代先ともに触らない
          ・交代先で新たな遅出翌日早出パターンを作らない
        1件ずつ独立に判定するため、改善できない箇所はそのまま残る
        （安定版の結果を壊してまで解消しようとはしない）。
        """
        sc = dict(sc)
        pl = dict(pl)
        base_hard = _hard_signature(sc, pl)
        applied = []
        skipped = []

        targets = [(sid, day) for sid in staff for day in range(2, last_day + 1)
                   if sc.get((sid, day - 1)) == "遅" and sc.get((sid, day)) == "早"]

        for sid, day in targets:
            # 直前のスワップで既に解消済みなら再判定不要
            if not (sc.get((sid, day - 1)) == "遅" and sc.get((sid, day)) == "早"):
                continue
            if _is_protected(sid, day):
                skipped.append({"day": day, "staff_id": sid, "name": staff[sid]["name"],
                                 "reason": "本人が希望/固定のため対象外"})
                continue

            uid = pl.get((sid, day), staff[sid]["unit_id"])
            _, cur_detail = _score(sc, pl)
            best_choice = None
            best_score = None

            candidates = [
                s["id"] for s in staff.values()
                if s["id"] != sid
                and s.get("can_early") and not s.get("can_night_only")
                and sc.get((s["id"], day), "休") == "休"
                and not _is_protected(s["id"], day)
                and sc.get((s["id"], day - 1)) not in ("遅", "夜", "明")
                and (staff[s["id"]]["unit_id"] == uid
                     or skill_map.get(s["id"], {}).get(uid) in ("ok", "conditional"))
            ]

            if not candidates:
                skipped.append({"day": day, "staff_id": sid, "name": staff[sid]["name"],
                                 "reason": "交代候補なし（早出可能・休み・条件適合の職員が不在）"})
                continue

            # パート・派遣の契約勤務日数（1日のズレで1800点相当の重み）が絡む
            # スワップは、わずかな改善でもこの基準では弾かれてしまい、かつ
            # 契約日数遵守の優先順位の方が高いため、ここでは正社員同士の
            # スワップのみを対象とする（パート・派遣側の遅出翌日早出は、
            # 契約日数を壊さない形で別途対応する）。
            if _is_part_or_dispatch(staff[sid]):
                skipped.append({"day": day, "staff_id": sid, "name": staff[sid]["name"],
                                 "reason": "本人がパート・派遣のため対象外（契約日数優先）"})
                continue
            candidates = [psid for psid in candidates if not _is_part_or_dispatch(staff[psid])]
            if not candidates:
                skipped.append({"day": day, "staff_id": sid, "name": staff[sid]["name"],
                                 "reason": "交代候補なし（正社員の早出可能・休み・条件適合の職員が不在）"})
                continue

            # 許容誤差。厳密ゼロだと、週次バランス・連勤連休側のわずかな
            # 揺らぎ（スワップで2人分のパターンが変わるため必ず生じる）で
            # ほぼ全件が弾かれてしまうため、小さな許容幅を持たせる。
            # 契約日数(contract)は正社員同士なら常に0なので許容は不要。
            TOL = 3.0
            for psid in candidates:
                trial_sc = dict(sc)
                trial_pl = dict(pl)
                trial_sc[(sid, day)] = "休"
                trial_pl.pop((sid, day), None)
                trial_sc[(psid, day)] = "早"
                trial_pl[(psid, day)] = uid

                trial_hard = _hard_signature(trial_sc, trial_pl)
                if any(c > b for c, b in zip(trial_hard, base_hard)):
                    continue
                trial_score, trial_detail = _score(trial_sc, trial_pl)
                # _score合計の改善ではなく、レポート4節の各項目を個別にチェックする。
                # _score()自体には遅出翌日早出を評価する項目がないため、合計の
                # 厳密改善だけを条件にすると、この後処理パスがほぼ機能しなくなる。
                not_worse = (
                    trial_detail["contract"] <= cur_detail["contract"] + 1e-6
                    and trial_detail["balance"] <= cur_detail["balance"] + TOL
                    and trial_detail["pattern"] <= cur_detail["pattern"] + TOL
                    and trial_detail["cross"] <= cur_detail["cross"] + 1e-6
                )
                if not not_worse:
                    continue
                if best_score is None or trial_score < best_score:
                    best_score = trial_score
                    best_choice = (psid, trial_sc, trial_pl)

            if best_choice:
                psid, sc, pl = best_choice
                applied.append({"day": day, "from_staff_id": sid, "to_staff_id": psid,
                                 "from_name": staff[sid]["name"], "to_name": staff[psid]["name"]})
            else:
                skipped.append({"day": day, "staff_id": sid, "name": staff[sid]["name"],
                                 "reason": f"候補{len(candidates)}名いたが、必須条件か契約/週次/連勤連休/他ユニットのいずれかが悪化するため不採用"})

        return sc, pl, applied, skipped

    def _post_process_balance_early_late_fulltime(sc, pl):
        """正社員（夜勤専従・パート派遣を除く）の「一人ひとりの中の早出:遅出比率」
        をユニット内で整える後処理パス。

        前バージョン（人と人の合計を比べて多い方から少ない方へ1日分移す方式）は
        「ユニット全員が同じように早出に偏っている/遅出に偏っている」ケースを
        検出できなかった（誰と比べても差がないため）。修正仕様⑧が求めているのは
        各個人の早出・遅出比率の均等化であり、これには「Aの遅出1日とBの早出1日を
        入れ替える」クロストレードが必要。

        設計方針は他の後処理パスと同じ:
          ・ランキング/スコア関数(_score, rank_for, assign_slot)は変更しない
          ・パート・派遣は対象外（契約日数優先）
          ・契約(contract)・他ユニット(cross)は悪化させない
          ・週次バランス(balance)・連勤連休(pattern)の悪化は許容誤差以内
          ・必須条件(hard_signature)は1件も悪化させない
        """
        sc = dict(sc)
        pl = dict(pl)
        applied = []
        skipped_summary = {"no_valid_trade": 0, "rejected_by_tolerance": 0}
        TOL = 3.0
        MAX_ROUNDS = 300

        def counts(sid):
            e = sum(1 for d in range(1, last_day + 1) if sc.get((sid, d)) == "早")
            l = sum(1 for d in range(1, last_day + 1) if sc.get((sid, d)) == "遅")
            return e, l

        for unit in operational_units:
            uid = unit["id"]
            members = [s["id"] for s in staff.values()
                       if s["unit_id"] == uid
                       and not s.get("can_night_only")
                       and not _is_part_or_dispatch(s)
                       and s.get("can_early") and s.get("can_late")]
            if len(members) < 2:
                continue

            for _round in range(MAX_ROUNDS):
                skew = {sid: counts(sid)[0] - counts(sid)[1] for sid in members}  # 早-遅。負=遅に偏り
                a_sid = min(skew, key=skew.get)   # 最も遅に偏っている人(早を増やしたい)
                b_sid = max(skew, key=skew.get)   # 最も早に偏っている人(遅を増やしたい)
                if a_sid == b_sid or skew[b_sid] - skew[a_sid] < 4:
                    break  # これ以上整えても誤差レベル

                base_hard = _hard_signature(sc, pl)
                _, cur_detail = _score(sc, pl)

                a_late_days = [d for d in range(1, last_day + 1)
                               if sc.get((a_sid, d)) == "遅" and not _is_protected(a_sid, d)]
                b_early_days = [d for d in range(1, last_day + 1)
                                if sc.get((b_sid, d)) == "早" and not _is_protected(b_sid, d)]

                traded = False
                for dx in a_late_days:      # aの遅出日 → bが引き継ぐ
                    if _is_protected(b_sid, dx) or sc.get((b_sid, dx), "休") != "休":
                        continue
                    prev_b = sc.get((b_sid, dx - 1)) if dx > 1 else None
                    if prev_b in ("夜", "明"):
                        continue
                    for dy in b_early_days:  # bの早出日 → aが引き継ぐ
                        if dy == dx or _is_protected(a_sid, dy) or sc.get((a_sid, dy), "休") != "休":
                            continue
                        prev_a = sc.get((a_sid, dy - 1)) if dy > 1 else None
                        if prev_a in ("遅", "夜", "明"):
                            continue

                        trial_sc = dict(sc)
                        trial_pl = dict(pl)
                        trial_sc[(a_sid, dx)] = "休"
                        trial_pl.pop((a_sid, dx), None)
                        trial_sc[(b_sid, dx)] = "遅"
                        trial_pl[(b_sid, dx)] = uid
                        trial_sc[(b_sid, dy)] = "休"
                        trial_pl.pop((b_sid, dy), None)
                        trial_sc[(a_sid, dy)] = "早"
                        trial_pl[(a_sid, dy)] = uid

                        trial_hard = _hard_signature(trial_sc, trial_pl)
                        if any(c > b for c, b in zip(trial_hard, base_hard)):
                            continue
                        _, trial_detail = _score(trial_sc, trial_pl)
                        not_worse = (
                            trial_detail["contract"] <= cur_detail["contract"] + 1e-6
                            and trial_detail["cross"] <= cur_detail["cross"] + 1e-6
                            and trial_detail["balance"] <= cur_detail["balance"] + TOL
                            and trial_detail["pattern"] <= cur_detail["pattern"] + TOL
                        )
                        if not not_worse:
                            skipped_summary["rejected_by_tolerance"] += 1
                            continue

                        sc, pl = trial_sc, trial_pl
                        applied.append({
                            "a_staff_id": a_sid, "a_name": staff[a_sid]["name"],
                            "b_staff_id": b_sid, "b_name": staff[b_sid]["name"],
                            "a_late_day_given_to_b": dx, "b_early_day_given_to_a": dy,
                        })
                        traded = True
                        break
                    if traded:
                        break

                if not traded:
                    skipped_summary["no_valid_trade"] += 1
                    break

        return sc, pl, applied, skipped_summary

    best_sc, best_pl, late_then_early_swaps, late_then_early_skipped = \
        _post_process_late_then_early(best_sc, best_pl)

    best_sc, best_pl, balance_swaps, balance_skip_summary = \
        _post_process_balance_early_late_fulltime(best_sc, best_pl)

    entries = _sched_to_entries(best_sc, best_pl)
    evaluation = _evaluate_schedule(entries, staff, req_map, units_list, year, month, last_day)
    evaluation["late_then_early_post_process"] = {
        "applied_count": len(late_then_early_swaps),
        "applied": late_then_early_swaps,
        "skipped_count": len(late_then_early_skipped),
        "skipped": late_then_early_skipped,
    }
    evaluation["balance_early_late_post_process"] = {
        "applied_count": len(balance_swaps),
        "applied": balance_swaps,
        "skip_summary": balance_skip_summary,
    }
    evaluation["unit_surplus"] = _compute_unit_surplus(entries, last_day)

    # change_reasonへ追記（eval_jsonは別画面の再評価で上書きされるため、
    # 後処理パスの結果はnote欄にも残し、追跡できるようにする）.
    if vid and vid > 0:
        pp_lines = []
        if restrict_unit_id is not None:
            cross_leak = [e for e in entries
                          if e["unit_id"] != restrict_unit_id
                          and staff.get(e["staff_id"], {}).get("unit_id") == restrict_unit_id
                          and e["shift_type"] not in ("休", "欠勤", None)]
            pp_lines.append(
                f"[診断] restrict_unit_id={restrict_unit_id}"
                f"(type={type(restrict_unit_id).__name__}) "
                f"/ skill_map保持エントリ数={sum(len(v) for v in skill_map.values())} "
                f"/ 制限後も残る他ユニット配置={len(cross_leak)}件")
            for e in cross_leak[:15]:
                pp_lines.append(f"  漏れ: staff_id={e['staff_id']} {e['date']} "
                                 f"{e['shift_type']} -> unit_id={e['unit_id']}")
        pp_lines.append(f"[後処理] 遅出翌日早出 検出{len(late_then_early_swaps) + len(late_then_early_skipped)}件 "
                    f"/ 解消{len(late_then_early_swaps)}件 / 見送り{len(late_then_early_skipped)}件")
        for a in late_then_early_swaps:
            pp_lines.append(f"  解消: {a['day']}日 {a['from_name']}→{a['to_name']}")
        for s in late_then_early_skipped:
            pp_lines.append(f"  見送り: {s['day']}日 {s['name']} ({s['reason']})")
        pp_lines.append(f"[後処理] 正社員早出/遅出比率均等化 トレード{len(balance_swaps)}件 "
                         f"/ 有効な交換先なしで打ち切り{balance_skip_summary['no_valid_trade']}回 "
                         f"/ 許容誤差超過で却下{balance_skip_summary['rejected_by_tolerance']}回")
        for a in balance_swaps:
            pp_lines.append(f"  交換: {a['a_name']}の遅出{a['a_late_day_given_to_b']}日→{a['b_name']}へ / "
                             f"{a['b_name']}の早出{a['b_early_day_given_to_a']}日→{a['a_name']}へ")
        db = get_db()
        row = qdb("SELECT change_reason FROM shift_versions WHERE id=?", (vid,), one=True)
        existing = []
        if row and row["change_reason"]:
            try:
                existing = __import__('json').loads(row["change_reason"])
                if not isinstance(existing, list):
                    existing = [str(existing)]
            except Exception:
                existing = [row["change_reason"]]
        combined = __import__('json').dumps(existing + pp_lines, ensure_ascii=False)
        db.execute("UPDATE shift_versions SET change_reason=? WHERE id=?", (combined, vid))
        db.commit()
    evaluation["iterative_balance"] = {
        "enabled": True,
        "logic_version": "rootcause_v3_night_pair_unlock",
        "seed": seed,
        "trials": trials,
        "accepted": accepted,
        "improvements": improved_count,
        "stopped_by": (
            "time_limit" if time.perf_counter() - started >= TIME_LIMIT_SEC else
            "no_improvement" if no_improve >= NO_IMPROVE_LIMIT else
            "max_trials"
        ),
        "elapsed_sec": round(time.perf_counter() - started, 3),
        "initial_score": round(_score(initial_sc, initial_pl)[0], 3),
        "final_score": round(best_score, 3),
        "initial_detail": _score(initial_sc, initial_pl)[1],
        "final_detail": best_detail,
        "initial_hard_signature": list(initial_hard_signature),
        "final_hard_signature": list(final_hard_signature),
        "fallback_to_initial": fallback_to_initial,
        "initial_repairs": initial_repairs,
        "initial_contract_rebalances": initial_contract_rebalances,
        "initial_night_pattern_rebalances": initial_night_pattern_rebalances,
        "initial_pattern_rebalances": initial_pattern_rebalances,
        "post_balance_repairs": post_balance_repairs,
        "initial_unresolved": [list(x) for x in initial_unresolved],
        "final_contract_rebalances": final_contract_rebalances,
        "final_night_pattern_rebalances": final_night_pattern_rebalances,
        "final_pattern_rebalances": final_pattern_rebalances,
        "final_repairs": final_repairs,
        "final_unresolved": [list(x) for x in final_unresolved],
        "hard_guard": "早・遅・夜・明・希望・固定・能力・夜勤明け・上限の悪化を禁止",
        "log": log,
    }

    return entries, evaluation

import html as _html_mod

WEEKDAY_JA = ["月","火","水","木","金","土","日"]
# シフト種別ごとの色（画面表示の配色に合わせる）
SHIFT_COLORS = {
    "日":  ("#e8f0ff", "#1d4ed8"),
    "早":  ("#ecfdf3", "#166534"),
    "遅":  ("#fdf2f8", "#be185d"),
    "夜":  ("#f0e9ff", "#6d28d9"),
    "明":  ("#eef2ff", "#4338ca"),
    "休":  ("#f3f4f6", "#4b5563"),
    "希休": ("#fdf0ff", "#8b35b5"),
    "欠勤": ("#fee2e2", "#b91c1c"),
    "有給": ("#fff7ed", "#9a3412"),
    "ヘ":  ("#ecfeff", "#0f766e"),
}

@app.route("/print_shift")
def print_shift():
    """A4横向きでの印刷用シフト表を表示する。画面と同じ配色で色分けし、
    実績（早退・遅刻）マーク・欠勤対応（前版から変更）マークも反映する。
    クエリ: year, month, facility(コード)
    """
    u = get_session_staff()
    if not u:
        return redirect("/login")

    try:
        year = int(request.args.get("year"))
        month = int(request.args.get("month"))
    except (TypeError, ValueError):
        return "年月の指定が不正です", 400
    fac_code = request.args.get("facility", "")

    f = qdb("SELECT * FROM facilities WHERE code=?", (fac_code,), one=True)
    if not f:
        return f"事業所 '{_html_mod.escape(fac_code)}' が見つかりません", 404

    ver = qdb(
        "SELECT * FROM shift_versions WHERE facility_id=? AND year=? AND month=? "
        "ORDER BY version_no DESC LIMIT 1",
        (f["id"], year, month), one=True)
    if not ver:
        return f"{year}年{month}月のシフトはまだ作成されていません", 404

    last_day = calendar.monthrange(year, month)[1]
    days = list(range(1, last_day + 1))

    # 職員の「所属ユニット」を基準に取得する（応援等で日によって配置ユニットが
    # 変わっても、1人1行にまとまるようにするため。se.unit_idでグルーピングすると
    # 応援日だけ別ユニットの行に分裂し、大部分が空欄の行ができてしまうため避ける）。
    rows = qdb("""SELECT se.staff_id, se.date, se.shift_type,
                         s.name sname, s.employment_type,
                         hu.id uid, hu.name uname, hu.unit_no
                  FROM shift_entries se
                  JOIN staff s ON se.staff_id=s.id
                  JOIN units hu ON s.unit_id=hu.id
                  WHERE se.version_id=?
                  ORDER BY hu.unit_no, s.id, se.date""", (ver["id"],))

    # 実績（早退・遅刻）マーク
    actual_rows = qdb(
        "SELECT staff_id, date, actual_type FROM shift_actual WHERE version_id=?", (ver["id"],))
    actual_map = {}
    for r in actual_rows:
        d = datetime.strptime(r["date"], "%Y-%m-%d").day
        actual_map[(r["staff_id"], d)] = r["actual_type"]

    # 欠勤対応等で「前版から変更」されたセル（画面のオレンジ丸マークに相当）
    diff_cells = json.loads(ver["diff_cells_json"]) if ver["diff_cells_json"] else []
    diff_set = set()
    for c in diff_cells:
        try:
            dd = datetime.strptime(c["date"], "%Y-%m-%d").day
            diff_set.add((int(c["staff_id"]), dd))
        except (KeyError, TypeError, ValueError):
            pass

    units = {}
    for r in rows:
        uid = r["uid"]
        if uid not in units:
            units[uid] = {"name": r["uname"], "unit_no": r["unit_no"], "staff": {}}
        sid = r["staff_id"]
        if sid not in units[uid]["staff"]:
            units[uid]["staff"][sid] = {"name": r["sname"], "days": {}}
        d = datetime.strptime(r["date"], "%Y-%m-%d").day
        units[uid]["staff"][sid]["days"][d] = r["shift_type"] or ""

    esc = _html_mod.escape

    def cell_style(v):
        if v in SHIFT_COLORS:
            bg, fg = SHIFT_COLORS[v]
            return f'background:{bg};color:{fg};font-weight:700'
        return ''

    day_headers = "".join(
        f'<th class="daycol">{d}<br><span class="dow">{WEEKDAY_JA[date(year,month,d).weekday()]}</span></th>'
        for d in days
    )

    body_rows = []
    for uid, udata in sorted(units.items(), key=lambda x: x[1]["unit_no"] or 0):
        body_rows.append(
            f'<tr class="unit-row"><td colspan="{len(days)+1}">{esc(udata["name"])}</td></tr>'
        )
        for sid, sdata in udata["staff"].items():
            cells = []
            for d in days:
                v = sdata["days"].get(d, "")
                marks = ""
                at = actual_map.get((sid, d))
                if at == "早退":
                    marks += '<span class="mk mk-early">早</span>'
                elif at == "遅刻":
                    marks += '<span class="mk mk-late">遅</span>'
                if (sid, d) in diff_set:
                    marks += '<span class="mk-chg"></span>'
                cells.append(
                    f'<td class="daycol" style="{cell_style(v)}"><span class="cellwrap">{esc(v)}</span>{marks}</td>'
                )
            body_rows.append(f'<tr><td class="namecol">{esc(sdata["name"])}</td>{"".join(cells)}</tr>')

    legend = "".join(
        f'<span class="legend-item"><span class="legend-sw" style="background:{bg};color:{fg}"></span>{esc(k)}</span>'
        for k, (bg, fg) in SHIFT_COLORS.items()
    )

    html_out = f"""<!DOCTYPE html>
<html lang="ja"><head><meta charset="utf-8">
<title>{esc(f["name"])} {year}年{month}月 シフト表</title>
<style>
  @page {{ size: A4 landscape; margin: 8mm; }}
  body {{ font-family: "Hiragino Sans","Yu Gothic",sans-serif; font-size: 10px; margin:0; }}
  h1 {{ font-size: 14px; margin: 4px 8px; }}
  table {{ border-collapse: collapse; width: 100%; table-layout: fixed; }}
  th, td {{ border: 1px solid #999; text-align: center; padding: 2px; }}
  .namecol {{ width: 90px; text-align: left; white-space: nowrap; font-size: 10px; background:#fff!important;color:#111!important;font-weight:400!important }}
  .daycol {{ width: 24px; font-size: 9px; position:relative; overflow:visible }}
  .cellwrap{{ display:inline-block }}
  .dow {{ color: #888; font-size: 8px; }}
  .unit-row td {{ background:#eef1f5!important; font-weight:bold; text-align:left; color:#172033!important }}
  .print-bar {{ margin: 8px; display:flex; align-items:center; gap:14px; flex-wrap:wrap }}
  .legend-item {{ display:inline-flex; align-items:center; gap:3px; font-size:10px; margin-right:8px }}
  .legend-sw {{ display:inline-block; width:14px; height:14px; border-radius:3px; border:1px solid #ccc }}
  .mk{{position:absolute;right:1px;bottom:1px;width:9px;height:9px;border-radius:50%;background:#111;color:#fff;font-size:6px;line-height:9px;text-align:center;font-weight:700}}
  .mk-chg{{position:absolute;top:1px;right:1px;width:6px;height:6px;border-radius:50%;background:#d97706;box-shadow:0 0 0 1px #fff}}
  @media print {{ .print-bar {{ display:none; }} }}
</style>
</head>
<body>
  <div class="print-bar">
    <button onclick="window.print()">🖨️ 印刷</button>
    <button onclick="window.close()">閉じる</button>
    <span style="display:inline-flex;flex-wrap:wrap;gap:4px">{legend}</span>
    <span class="legend-item"><span class="mk" style="position:static;display:inline-block">早</span>早退</span>
    <span class="legend-item"><span class="mk" style="position:static;display:inline-block">遅</span>遅刻</span>
    <span class="legend-item"><span class="mk-chg" style="position:static;display:inline-block"></span>前版から変更（欠勤対応等）</span>
  </div>
  <h1>{esc(f["name"])}　{year}年{month}月　シフト表（第{ver["version_no"]}版）</h1>
  <table>
    <thead><tr><th class="namecol">氏名</th>{day_headers}</tr></thead>
    <tbody>{"".join(body_rows)}</tbody>
  </table>
</body></html>"""
    return html_out

# ── API: シフト取得 ──────────────────────
@app.route("/api/shifts/<int:year>/<int:month>")
def api_get_shifts(year,month):
    u, err = require_any_login()
    if err: return err
    fac_code = request.args.get("facility")
    ver_no   = request.args.get("version")
    fid = None
    if fac_code:
        f = qdb("SELECT id FROM facilities WHERE code=?",(fac_code,),one=True)
        if f: fid = f["id"]

    if ver_no and fid:
        ver = qdb("SELECT id FROM shift_versions WHERE facility_id=? AND year=? AND month=? AND version_no=?",(fid,year,month,ver_no),one=True)
    elif fid:
        ver = qdb("SELECT id FROM shift_versions WHERE facility_id=? AND year=? AND month=? ORDER BY version_no DESC LIMIT 1",(fid,year,month),one=True)
    else:
        ver = qdb("SELECT id FROM shift_versions WHERE year=? AND month=? ORDER BY version_no DESC LIMIT 1",(year,month),one=True)

    if not ver: return jsonify({"error":"not found"}),404
    rows = qdb("""SELECT se.id, se.version_id, se.staff_id, se.unit_id,
                         se.date, se.shift_type,
                         se.is_cross, se.from_unit_id, se.is_manual,
                         s.name sname, s.monthly_limit, s.can_night,
                         u.name uname, u.unit_no,
                         f.code fcode, f.name fname
                  FROM shift_entries se
                  JOIN staff s ON se.staff_id=s.id
                  JOIN units u ON se.unit_id=u.id
                  JOIN facilities f ON u.facility_id=f.id
                  WHERE se.version_id=?
                  ORDER BY f.sort_order,u.unit_no,s.id,se.date""",(ver["id"],))
    return jsonify([dict(r) for r in rows])

@app.route("/api/shifts/<int:vid>/diff_cells")
def api_get_diff_cells(vid):
    """この版が「欠勤対応」等で前の版から変更されたセルの一覧を返す。
    画面リロード後もハイライト表示を復元するために使う（2026-06-21追加）。
    """
    u, err = require_any_login()
    if err: return err
    ver = qdb("SELECT based_on_version_id, diff_cells_json FROM shift_versions WHERE id=?",
              (vid,), one=True)
    if not ver:
        return jsonify({"based_on_version_id": None, "diff_cells": []})
    diff_cells = json.loads(ver["diff_cells_json"]) if ver["diff_cells_json"] else []
    return jsonify({"based_on_version_id": ver["based_on_version_id"], "diff_cells": diff_cells})

@app.route("/api/shifts/<int:vid>/entry", methods=["PUT"])
def api_update_entry(vid):
    u, err = require_perm("perm_create", "perm_change")
    if err: return err
    d = request.get_json()
    sid,dt,shift = d["staff_id"],d["date"],d["shift_type"]
    dt_obj = datetime.strptime(dt, "%Y-%m-%d")

    old = qdb("SELECT shift_type FROM shift_entries WHERE version_id=? AND staff_id=? AND date=?",(vid,sid,dt),one=True)
    old_shift = old["shift_type"] if old else None

    # 【夜→別シフト】翌日の明けを休に変える
    if old_shift == "夜" and shift != "夜":
        ndt = (dt_obj + timedelta(days=1)).strftime("%Y-%m-%d")
        xdb("UPDATE shift_entries SET shift_type='休',is_manual=1 WHERE version_id=? AND staff_id=? AND date=?",(vid,sid,ndt))

    # 【明→クリア(休)】前日が夜勤なら前日も休に変える（1日目は前月夜勤のため除外）
    if old_shift == "明" and shift == "休" and dt_obj.day > 1:
        pdt = (dt_obj - timedelta(days=1)).strftime("%Y-%m-%d")
        prev = qdb("SELECT shift_type FROM shift_entries WHERE version_id=? AND staff_id=? AND date=?",(vid,sid,pdt),one=True)
        if prev and prev["shift_type"] == "夜":
            xdb("UPDATE shift_entries SET shift_type='休',is_manual=1 WHERE version_id=? AND staff_id=? AND date=?",(vid,sid,pdt))

    xdb("UPDATE shift_entries SET shift_type=?,is_manual=1 WHERE version_id=? AND staff_id=? AND date=?",(shift,vid,sid,dt))

    # 【夜を新たに入力】翌日を明けにする
    if shift == "夜":
        ndt = (dt_obj + timedelta(days=1)).strftime("%Y-%m-%d")
        xdb("UPDATE shift_entries SET shift_type='明',is_manual=1 WHERE version_id=? AND staff_id=? AND date=?",(vid,sid,ndt))

    return jsonify({"ok":True})

def _validate_saved_gh_entries(vid, entries):
    """画面保存直前のGHシフトを必須条件で検証する。

    早・遅・夜・明の不足、夜→翌日明の不整合、希望入力違反を検出する。
    不正なシフトはDBへ書き込まない。
    """
    ver = qdb(
        "SELECT sv.facility_id,sv.year,sv.month,f.type "
        "FROM shift_versions sv JOIN facilities f ON sv.facility_id=f.id "
        "WHERE sv.id=?", (vid,), one=True)
    if not ver:
        return False, {"error":"version not found"}
    if ver["type"] != "GH":
        return True, {"shortages":[], "broken_night_ake":[], "broken_requests":[]}

    fid, year, month = ver["facility_id"], ver["year"], ver["month"]
    last_day = calendar.monthrange(year, month)[1]
    units = qdb(
        "SELECT id,name,unit_no,residents FROM units "
        "WHERE facility_id=? AND residents>0 AND unit_no<>4 ORDER BY unit_no", (fid,))
    req_rows = qdb(
        "SELECT unit_id,shift_type,required FROM unit_required_staff "
        "WHERE unit_id IN (SELECT id FROM units WHERE facility_id=?)", (fid,))
    req = defaultdict(dict)
    for r in req_rows:
        req[r["unit_id"]][r["shift_type"]] = int(r["required"] or 0)
    for u in units:
        if not req.get(u["id"]):
            req[u["id"]] = {"早":1,"遅":1,"夜":1}

    sched = {}
    place = {}
    duplicate_keys = []
    for e in entries:
        sid = int(e["staff_id"])
        dt_str = str(e["date"])
        key = (sid, dt_str)
        if key in sched:
            duplicate_keys.append({"staff_id":sid,"date":dt_str})
        sched[key] = e["shift_type"]
        place[key] = int(e["unit_id"])

    shortages = []
    for u in units:
        uid = u["id"]
        night_need = int(req[uid].get("夜", 1))
        for day in range(1, last_day+1):
            dt_str = f"{year}-{month:02d}-{day:02d}"
            for slot in ("早","遅","夜"):
                need = int(req[uid].get(slot, 1))
                actual = sum(1 for key, sh in sched.items()
                             if key[1] == dt_str and sh == slot and place.get(key) == uid)
                if actual < need:
                    shortages.append({"date":dt_str,"unit_id":uid,"unit_name":u["name"],
                                      "slot":slot,"required":need,"actual":actual})
            actual_ake = sum(1 for key, sh in sched.items()
                             if key[1] == dt_str and sh == "明" and place.get(key) == uid)
            if actual_ake < night_need:
                shortages.append({"date":dt_str,"unit_id":uid,"unit_name":u["name"],
                                  "slot":"明","required":night_need,"actual":actual_ake})

    broken_night_ake = []
    for (sid, dt_str), sh in sched.items():
        d = datetime.strptime(dt_str, "%Y-%m-%d")
        day = d.day
        if sh == "夜" and day < last_day:
            ndt = (d + timedelta(days=1)).strftime("%Y-%m-%d")
            if sched.get((sid, ndt)) != "明" or place.get((sid, ndt)) != place.get((sid, dt_str)):
                broken_night_ake.append({"staff_id":sid,"night_date":dt_str,"ake_date":ndt})
        if sh == "明" and day > 1:
            pdt = (d - timedelta(days=1)).strftime("%Y-%m-%d")
            if sched.get((sid, pdt)) != "夜" or place.get((sid, pdt)) != place.get((sid, dt_str)):
                broken_night_ake.append({"staff_id":sid,"ake_date":dt_str,"night_date":pdt})

    type_map = {"rest":"希休","hol":"有給","absence":"欠勤","early":"早","late":"遅",
                "night":"夜","day":"日","ake":"明"}
    broken_requests = []
    rows = qdb(
        "SELECT staff_id,day,req_type FROM requests WHERE year=? AND month=?", (year,month))
    for r in rows:
        expected = type_map.get(r["req_type"])
        if expected is None:
            continue
        dt_str = f"{year}-{month:02d}-{int(r['day']):02d}"
        actual = sched.get((int(r["staff_id"]), dt_str), "休")
        if actual != expected:
            broken_requests.append({"staff_id":r["staff_id"],"date":dt_str,
                                    "expected":expected,"actual":actual})

    ok = not shortages and not broken_night_ake and not broken_requests and not duplicate_keys
    return ok, {
        "shortages": shortages,
        "broken_night_ake": broken_night_ake,
        "broken_requests": broken_requests,
        "duplicate_entries": duplicate_keys,
    }


@app.route("/api/shifts/<int:vid>/entries", methods=["PUT"])
def api_save_entries(vid):
    """ドラフト画面からの一括保存。検証結果にかかわらず保存し、検証内容を記録する。

    absence_diff_cells（任意）: 欠勤対応のプレビュー適用で変更されたセルの
    一覧。指定されていれば、既存のdiff_cells_json（累積差分）にマージしてから
    保存する（欠勤対応自体はDBへ即時保存しない設計のため、ここで合流させる）。
    """
    u, err = require_perm("perm_create", "perm_change")
    if err: return err
    d = request.get_json(silent=True) or {}
    entries = d.get("entries", [])
    absence_diff_cells = d.get("absence_diff_cells") or []
    if not isinstance(entries, list) or not entries:
        return jsonify({"ok":False,"error":"保存対象のシフトがありません"}), 400

    # 生成ロジック検証中は、不足・不整合があっても保存を止めない。
    # 検証結果はeval_jsonとレスポンスへ残し、DBエクスポート後の比較に使用する。
    ok, validation = _validate_saved_gh_entries(vid, entries)

    db = get_db()
    try:
        db.execute("BEGIN")
        # 保存前に現在のentriesをバックアップ（直前に戻す機能用）
        existing = db.execute(
            "SELECT staff_id,unit_id,date,shift_type,is_cross,from_unit_id FROM shift_entries WHERE version_id=?",
            (vid,)).fetchall()
        if existing:
            db.execute("DELETE FROM shift_entries_backup WHERE version_id=?", (vid,))
            db.executemany(
                "INSERT INTO shift_entries_backup(version_id,staff_id,unit_id,date,shift_type,is_cross,from_unit_id) VALUES(?,?,?,?,?,?,?)",
                [(vid, e[0], e[1], e[2], e[3], e[4], e[5]) for e in existing])
        db.execute("DELETE FROM shift_entries WHERE version_id=?", (vid,))
        db.executemany(
            """INSERT INTO shift_entries
               (version_id,staff_id,unit_id,date,shift_type,is_cross,from_unit_id,is_manual)
               VALUES(?,?,?,?,?,?,?,1)""",
            [(vid, e["staff_id"], e["unit_id"], e["date"], e["shift_type"],
              e.get("is_cross", 0), e.get("from_unit_id")) for e in entries]
        )
        save_meta = {
            "saved_entry_count": len(entries),
            "validation_ok": bool(ok),
            "save_validation": validation,
            "validated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        }
        old = qdb("SELECT eval_json FROM shift_versions WHERE id=?", (vid,), one=True)
        eval_data = {}
        if old and old["eval_json"]:
            try:
                eval_data = json.loads(old["eval_json"])
            except Exception:
                eval_data = {}
        eval_data["save_guard"] = save_meta
        db.execute("UPDATE shift_versions SET status='draft',eval_json=? WHERE id=?",
                   (json.dumps(eval_data, ensure_ascii=False), vid))

        if absence_diff_cells:
            diff_keys = {(c["staff_id"], c["date"]) for c in absence_diff_cells}
            old_diff = qdb("SELECT diff_cells_json FROM shift_versions WHERE id=?", (vid,), one=True)
            if old_diff and old_diff["diff_cells_json"]:
                for c in json.loads(old_diff["diff_cells_json"]):
                    diff_keys.add((c["staff_id"], c["date"]))
            merged = [{"staff_id": k[0], "date": k[1]} for k in diff_keys]
            db.execute("UPDATE shift_versions SET diff_cells_json=? WHERE id=?",
                       (json.dumps(merged, ensure_ascii=False), vid))
        db.commit()
    except Exception:
        db.rollback()
        raise
    return jsonify({
        "ok": True,
        "saved": len(entries),
        "validation_ok": bool(ok),
        "validation": validation,
    })


@app.route("/api/shifts/versions/<int:vid>/undo", methods=["POST"])
def api_undo_entries(vid):
    """直前の保存状態に戻す。shift_entries_backupから復元する。"""
    u, err = require_perm("perm_create", "perm_change")
    if err: return err
    backup = qdb(
        "SELECT staff_id,unit_id,date,shift_type,is_cross,from_unit_id FROM shift_entries_backup WHERE version_id=?",
        (vid,))
    if not backup:
        return jsonify({"ok": False, "error": "バックアップがありません。保存後のみ使用できます。"}), 400
    db = get_db()
    db.execute("DELETE FROM shift_entries WHERE version_id=?", (vid,))
    db.executemany(
        "INSERT INTO shift_entries(version_id,staff_id,unit_id,date,shift_type,is_cross,from_unit_id,is_manual) VALUES(?,?,?,?,?,?,?,1)",
        [(vid, e["staff_id"], e["unit_id"], e["date"], e["shift_type"], e["is_cross"], e["from_unit_id"]) for e in backup])
    # 変更差分ハイライトもリセット
    db.execute("UPDATE shift_versions SET diff_cells_json=NULL WHERE id=?", (vid,))
    # バックアップを削除（1回のみ戻せる）
    db.execute("DELETE FROM shift_entries_backup WHERE version_id=?", (vid,))
    db.commit()
    return jsonify({"ok": True, "restored": len(backup)})


@app.route("/api/shifts/new_version", methods=["POST"])
def api_new_version():
    """公開済みシフトから新版（version_no+1）を作成する。

    バージョン番号は「公開済み版が存在する場合」にのみインクリメントする。
    draft段階での保存は同一version_noのままentries上書き（api_save_entries）で行う。

    オプション: copy_entries=true を渡すと、直前版のentriesをそのままコピーして
    編集ベースにする（公開後の「シフト変更（新版作成）」ボタン用）。
    """
    u, err = require_perm("perm_create", "perm_change")
    if err: return err
    d = request.get_json()
    year      = d.get("year")
    month     = d.get("month")
    fac_code  = d.get("facility_code")
    reason    = d.get("reason", "新版作成")
    copy_entries = d.get("copy_entries", False)

    f = qdb("SELECT id FROM facilities WHERE code=?", (fac_code,), one=True)
    if not f:
        return jsonify({"ok": False, "error": "facility not found"}), 404
    fid = f["id"]

    # 最新版を取得
    latest = qdb(
        "SELECT id, version_no, status FROM shift_versions "
        "WHERE facility_id=? AND year=? AND month=? ORDER BY version_no DESC LIMIT 1",
        (fid, year, month), one=True)

    # 既にdraftが存在する場合はそのまま返す（重複作成防止）
    if latest and latest["status"] == "draft":
        return jsonify({"ok": True, "version_id": latest["id"],
                        "version_no": latest["version_no"], "reused": True})

    new_ver = (latest["version_no"] if latest else 0) + 1
    vid = xdb(
        "INSERT INTO shift_versions(facility_id,year,month,version_no,status,change_reason) "
        "VALUES(?,?,?,?,'draft',?)",
        (fid, year, month, new_ver, reason))

    # 前版のentriesをコピー（公開後シフト変更の起点として使う）
    if copy_entries and latest:
        prev_entries = qdb(
            "SELECT staff_id,unit_id,date,shift_type,is_cross,from_unit_id "
            "FROM shift_entries WHERE version_id=?", (latest["id"],))
        if prev_entries:
            get_db().executemany(
                "INSERT INTO shift_entries(version_id,staff_id,unit_id,date,shift_type,is_cross,from_unit_id) "
                "VALUES(?,?,?,?,?,?,?)",
                [(vid, e["staff_id"], e["unit_id"], e["date"], e["shift_type"],
                  e["is_cross"], e["from_unit_id"]) for e in prev_entries])

    get_db().commit()
    return jsonify({"ok": True, "version_id": vid, "version_no": new_ver, "reused": False})


# ── API: バージョン ──────────────────────
@app.route("/api/shifts/versions/<int:vid>", methods=["DELETE"])
def api_delete_version(vid):
    """シフトバージョンとそのエントリを削除"""
    u, err = require_perm("perm_create", "perm_change")
    if err: return err
    db = get_db()
    db.execute("DELETE FROM shift_entries WHERE version_id=?", (vid,))
    db.execute("DELETE FROM shift_versions WHERE id=?", (vid,))
    db.commit()
    return jsonify({"ok": True})

@app.route("/api/shifts/versions/<int:year>/<int:month>")
def api_versions(year,month):
    u, err = require_any_login()
    if err: return err
    fac_code = request.args.get("facility")
    if fac_code:
        f = qdb("SELECT id FROM facilities WHERE code=?",(fac_code,),one=True)
        rows = qdb("SELECT * FROM shift_versions WHERE facility_id=? AND year=? AND month=? ORDER BY version_no DESC",(f["id"],year,month)) if f else []
    else:
        rows = qdb("SELECT * FROM shift_versions WHERE year=? AND month=? ORDER BY version_no DESC",(year,month))
    return jsonify([dict(r) for r in rows])

@app.route("/api/shifts/versions/<int:vid>/status", methods=["PUT"])
def api_update_status(vid):
    """シフト版のステータスを更新する。

    status遷移:
      draft → leader_approved  : シフト責任者承認（perm_approve_leader）
      leader_approved → published : 管理者が確定・公開（perm_approve_admin）
      published → draft         : 差し戻し（管理者のみ）
      * draft → published も許可（管理者が直接公開する小規模運用向け）

    published への遷移時:
      ・diff_cells_jsonをリセット（公開後の変更差分をゼロから積み上げ直す）
      ・approved_by / approved_at を記録

    バージョン番号インクリメントはここでは行わない。
    公開後に「シフト変更（新版作成）」ボタンを押したとき（/api/shifts/new_version）に
    初めて新しいversion_noを作る設計（publish=版確定、新版作成=次の編集開始）。
    """
    u, err = require_any_login()
    if err: return err
    d = request.get_json()
    status = d.get("status")
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    ver = qdb("SELECT * FROM shift_versions WHERE id=?", (vid,), one=True)
    if not ver:
        return jsonify({"ok": False, "error": "version not found"}), 404

    allowed_transitions = {
        "draft":            ("leader_approved", "published"),  # リーダー承認 or 管理者直接公開
        "leader_approved":  ("published", "draft"),            # 管理者公開 or 差し戻し
        "published":        ("draft",),                        # 差し戻し（再編集）
    }
    current = ver["status"] or "draft"
    allowed = allowed_transitions.get(current, ())
    if status not in allowed:
        return jsonify({"ok": False,
                        "error": f"ステータス '{current}' から '{status}' への遷移は許可されていません"}), 400

    # 遷移先ごとに必要な権限を確認する（承認=perm_approve、公開・差し戻し=perm_publish）
    role = (u.get("system_role") or "staff").strip()
    if role != "admin":
        perm = qdb("SELECT * FROM role_permissions WHERE role=?", (role,), one=True)
        needed = "perm_approve" if status == "leader_approved" else "perm_publish"
        if not perm or not perm[needed]:
            return jsonify({"ok": False, "error": "この操作を行う権限がありません"}), 403

    db = get_db()
    if status == "published":
        # 公開確定：承認者・承認日時を記録、差分ハイライトをリセット
        approved_by = d.get("approved_by") or "管理者"
        db.execute(
            "UPDATE shift_versions SET status=?,approved_by=?,approved_at=?,diff_cells_json=NULL WHERE id=?",
            (status, approved_by, now, vid))
    elif status == "leader_approved":
        approved_by = d.get("approved_by") or "リーダー"
        db.execute(
            "UPDATE shift_versions SET status=?,approved_by=?,approved_at=? WHERE id=?",
            (status, approved_by, now, vid))
    else:
        # draft（差し戻し）
        db.execute("UPDATE shift_versions SET status=? WHERE id=?", (status, vid))
    db.commit()
    return jsonify({"ok": True, "status": status})

@app.route("/api/shifts/versions/<int:vid>/approvals")
def api_get_approvals(vid):
    """版の承認状況を返す。
    各ユニットの承認者一覧と、誰が承認済みかを返す。
    全承認者が承認済みかどうかの判定も含む。
    """
    u, err = require_any_login()
    if err: return err
    ver = qdb("SELECT facility_id FROM shift_versions WHERE id=?", (vid,), one=True)
    if not ver:
        return jsonify({"ok": False, "error": "not found"}), 404
    fid = ver["facility_id"]

    # 事業所の全承認者（is_approver=1、在職中）をユニット別に取得
    approvers = qdb(
        """SELECT s.id, s.name, s.unit_id, u.name as unit_name
           FROM staff s JOIN units u ON s.unit_id=u.id
           WHERE s.facility_id=? AND s.is_approver=1 AND s.is_active=1
           ORDER BY u.unit_no, s.id""", (fid,))

    # 承認済みレコードを取得
    done = qdb(
        "SELECT staff_id, unit_id, approved_at, comment FROM shift_approvals WHERE version_id=?",
        (vid,))
    done_map = {r["staff_id"]: dict(r) for r in done}

    # ユニット別に整理
    units_status = {}
    for a in approvers:
        uid = a["unit_id"]
        if uid not in units_status:
            units_status[uid] = {"unit_id": uid, "unit_name": a["unit_name"], "approvers": []}
        rec = done_map.get(a["id"])
        units_status[uid]["approvers"].append({
            "staff_id": a["id"], "name": a["name"],
            "approved": rec is not None,
            "approved_at": rec["approved_at"] if rec else None,
            "comment": rec["comment"] if rec else None,
        })

    units_list = list(units_status.values())
    # 全承認者が承認済みか
    all_approved = (
        len(approvers) > 0
        and all(done_map.get(a["id"]) is not None for a in approvers)
    )
    return jsonify({
        "ok": True,
        "units": units_list,
        "total_approvers": len(approvers),
        "approved_count": len(done_map),
        "all_approved": all_approved,
    })


@app.route("/api/shifts/versions/<int:vid>/approvals", methods=["POST"])
def api_add_approval(vid):
    """承認者がユニット分を承認する。
    全承認者が承認したら version の status を leader_approved（シフト責任者承認済）に更新する。
    """
    u = get_session_staff()
    if not u:
        return jsonify({"ok": False, "error": "未ログインです"}), 401

    ver = qdb("SELECT facility_id, status FROM shift_versions WHERE id=?", (vid,), one=True)
    if not ver:
        return jsonify({"ok": False, "error": "version not found"}), 404
    if ver["status"] not in ("draft", "leader_approved"):
        return jsonify({"ok": False, "error": f"ステータス '{ver['status']}' では承認できません"}), 400

    d = request.get_json() or {}
    staff_id = d.get("staff_id") or u["id"]
    comment  = d.get("comment", "")

    # admin以外は自分自身のstaff_idでしか承認できない（なりすまし防止）
    if (u.get("system_role") or "staff") != "admin" and int(staff_id) != int(u["id"]):
        return jsonify({"ok": False, "error": "他の職員として承認することはできません"}), 403

    # 承認者か確認
    s = qdb("SELECT id, unit_id, is_approver FROM staff WHERE id=? AND is_active=1",
            (staff_id,), one=True)
    if not s or not s["is_approver"]:
        return jsonify({"ok": False, "error": "承認権限がありません"}), 403

    db = get_db()
    db.execute(
        """INSERT INTO shift_approvals(version_id, staff_id, unit_id, comment)
           VALUES(?,?,?,?)
           ON CONFLICT(version_id, staff_id) DO UPDATE SET
           approved_at=datetime('now','localtime'), comment=excluded.comment""",
        (vid, staff_id, s["unit_id"], comment))
    db.commit()

    # 全承認者が承認済みか確認
    fid = ver["facility_id"]
    approvers = qdb(
        "SELECT id FROM staff WHERE facility_id=? AND is_approver=1 AND is_active=1", (fid,))
    done = qdb(
        "SELECT staff_id FROM shift_approvals WHERE version_id=?", (vid,))
    done_ids = {r["staff_id"] for r in done}
    all_approved = len(approvers) > 0 and all(a["id"] in done_ids for a in approvers)

    if all_approved and ver["status"] not in ("leader_approved", "published"):
        db.execute("UPDATE shift_versions SET status='leader_approved' WHERE id=?", (vid,))
        db.commit()

    return jsonify({"ok": True, "all_approved": all_approved,
                    "approved_count": len(done_ids), "total": len(approvers)})


@app.route("/api/shifts/versions/<int:vid>/approvals/<int:staff_id>", methods=["DELETE"])
def api_delete_approval(vid, staff_id):
    """承認を取り消す（差し戻し時などに使用）。statusもdraftに戻す。"""
    u, err = require_perm("perm_approve", "perm_publish")
    if err: return err
    db = get_db()
    db.execute("DELETE FROM shift_approvals WHERE version_id=? AND staff_id=?", (vid, staff_id))
    db.execute("UPDATE shift_versions SET status='draft' WHERE id=?", (vid,))
    db.commit()
    return jsonify({"ok": True})


@app.route("/api/shifts/versions/<int:vid>/approvals", methods=["DELETE"])
def api_clear_approvals(vid):
    """全承認を取り消してdraftに戻す（管理者による差し戻し）。"""
    u, err = require_perm("perm_approve", "perm_publish")
    if err: return err
    db = get_db()
    db.execute("DELETE FROM shift_approvals WHERE version_id=?", (vid,))
    db.execute("UPDATE shift_versions SET status='draft' WHERE id=?", (vid,))
    db.commit()
    return jsonify({"ok": True})
def api_evaluate_version(vid):
    """
    保存済みシフト（自動生成直後・ドラフト手動編集後どちらでも可）を
    現在のentriesから再評価し、スコア・診断を返す（DBのeval_jsonも更新）。
    """
    ver = qdb("SELECT * FROM shift_versions WHERE id=?", (vid,), one=True)
    if not ver:
        return jsonify({"ok": False, "error": "version not found"}), 404

    year, month, fac_id = ver["year"], ver["month"], ver["facility_id"]
    last_day = calendar.monthrange(year, month)[1]

    fac = qdb("SELECT type FROM facilities WHERE id=?", (fac_id,), one=True)
    if fac and fac["type"] != "GH":
        return jsonify({"ok": True, "evaluation": None,
                        "note": "GH以外の事業所は評価対象外です"})

    rows = qdb("SELECT staff_id,unit_id,date,shift_type FROM shift_entries WHERE version_id=?", (vid,))
    entries = [dict(r) for r in rows]
    if not entries:
        return jsonify({"ok": False, "error": "no entries for this version"}), 404

    staff_rows = qdb(
        """SELECT id,name,can_day,can_early,can_late,can_night,can_night_only,
                  monthly_limit,night_target,unit_id
           FROM staff WHERE facility_id=? AND is_active=1""", (fac_id,))
    staff = {s["id"]: dict(s) for s in staff_rows}

    reqs = qdb("SELECT staff_id,day,req_type FROM requests WHERE year=? AND month=?", (year, month))
    req_map = defaultdict(dict)
    for r in reqs:
        req_map[r["staff_id"]][r["day"]] = r["req_type"]

    units_list = [dict(u) for u in
                   qdb("SELECT id,unit_no,name FROM units WHERE facility_id=? ORDER BY unit_no", (fac_id,))]

    evaluation = _evaluate_schedule(entries, staff, req_map, units_list, year, month, last_day)

    xdb("UPDATE shift_versions SET eval_json=? WHERE id=?",
        (json.dumps(evaluation, ensure_ascii=False), vid))
    return jsonify({"ok": True, "evaluation": evaluation})

# ── API: 割当候補 ────────────────────────
@app.route("/api/candidates")
def api_candidates():
    u, err = require_any_login()
    if err: return err
    unit_id = request.args.get("unit_id",type=int)
    target_date = request.args.get("date")
    vid = request.args.get("version_id",type=int)
    if not unit_id: return jsonify({"error":"unit_id required"}),400
    cands = qdb("""SELECT s.id,s.name,s.role,s.can_night,s.monthly_limit,
                          s.unit_id home_uid,u.name home_uname,f.code fcode,ss.level
                   FROM staff_skills ss
                   JOIN staff s ON ss.staff_id=s.id
                   JOIN units u ON s.unit_id=u.id
                   JOIN facilities f ON s.facility_id=f.id
                   WHERE ss.unit_id=? AND s.unit_id!=? AND s.is_active=1
                   ORDER BY CASE WHEN f.type='GH' THEN 0 ELSE 1 END,f.sort_order,s.id""",
                (unit_id,unit_id))
    result = []
    for c in cands:
        item = dict(c)
        if target_date and vid:
            e = qdb("SELECT shift_type FROM shift_entries WHERE version_id=? AND staff_id=? AND date=?",(vid,c["id"],target_date),one=True)
            item["current_shift"] = e["shift_type"] if e else None
        if vid:
            w = qdb("SELECT COUNT(*) cnt FROM shift_entries WHERE version_id=? AND staff_id=? AND shift_type NOT IN ('休','明')",(vid,c["id"]),one=True)
            item["work_days"] = w["cnt"] if w else 0
        result.append(item)
    return jsonify(result)

@app.route("/api/shifts/assign", methods=["POST"])
def api_assign():
    u, err = require_perm("perm_create", "perm_change")
    if err: return err
    d = request.get_json()
    vid,sid,uid,dt,shift = d["version_id"],d["staff_id"],d["unit_id"],d["date"],d.get("shift_type","日")
    home = qdb("SELECT unit_id FROM staff WHERE id=?",(sid,),one=True)
    is_cross = 1 if home and home["unit_id"]!=uid else 0
    xdb("UPDATE shift_entries SET shift_type=?,unit_id=?,is_cross=?,is_manual=1 WHERE version_id=? AND staff_id=? AND date=?",(shift,uid,is_cross,vid,sid,dt))
    if shift=="夜":
        ndt=(datetime.strptime(dt,"%Y-%m-%d")+timedelta(days=1)).strftime("%Y-%m-%d")
        xdb("UPDATE shift_entries SET shift_type='明',unit_id=?,is_cross=?,is_manual=1 WHERE version_id=? AND staff_id=? AND date=?",(uid,is_cross,vid,sid,ndt))
    return jsonify({"ok":True,"is_cross":bool(is_cross)})

# ── API: ルール違反チェック ──────────────
@app.route("/api/rules/check/<int:year>/<int:month>")
def api_check_rules(year,month):
    u, err = require_any_login()
    if err: return err
    fac_code = request.args.get("facility")
    if fac_code:
        f = qdb("SELECT id FROM facilities WHERE code=?",(fac_code,),one=True)
        ver = qdb("SELECT id FROM shift_versions WHERE facility_id=? AND year=? AND month=? ORDER BY version_no DESC LIMIT 1",(f["id"],year,month),one=True) if f else None
    else:
        ver = qdb("SELECT id FROM shift_versions WHERE year=? AND month=? ORDER BY version_no DESC LIMIT 1",(year,month),one=True)
    if not ver: return jsonify([])
    rows = qdb("""SELECT se.staff_id,se.date,se.shift_type,s.name,s.monthly_limit,u.name uname,f.code fcode
                  FROM shift_entries se
                  JOIN staff s ON se.staff_id=s.id
                  JOIN units u ON se.unit_id=u.id
                  JOIN facilities f ON u.facility_id=f.id
                  WHERE se.version_id=? ORDER BY se.staff_id,se.date""",(ver["id"],))
    shifts = defaultdict(list)
    for r in rows: shifts[r["staff_id"]].append(dict(r))
    violations = []
    for sid,ss in shifts.items():
        ss.sort(key=lambda x:x["date"])
        info=ss[0]; work=0; nstreak=0; limit=info["monthly_limit"]
        for i,sh in enumerate(ss):
            v=sh["shift_type"]
            cur_date=sh["date"]
            if v=="夜":
                # 直前エントリが「明」かつgap==1(翌日)のみ連続カウント
                # それ以外（間に休みあり・直前が他シフト）はリセット
                is_cont = False
                if i>0:
                    prev_v=ss[i-1]["shift_type"]
                    prev_d=ss[i-1]["date"]
                    gap=(datetime.strptime(cur_date,"%Y-%m-%d")-datetime.strptime(prev_d,"%Y-%m-%d")).days
                    is_cont = (prev_v=="明" and gap==1)
                if not is_cont:
                    nstreak=0
                nstreak+=1
                if nstreak==3:
                    violations.append({"type":"warning","staff_name":info["name"],"facility":info["fcode"],"unit":info["uname"],"date":cur_date,"rule":"連続夜勤3回（警告）","message":f"{cur_date} 連続夜勤3回"})
                elif nstreak>=4:
                    violations.append({"type":"error","staff_name":info["name"],"facility":info["fcode"],"unit":info["uname"],"date":cur_date,"rule":"連続夜勤4回以上（禁止）","message":f"{cur_date} 連続夜勤{nstreak}回（禁止）"})
            elif v=="明":
                pass  # 明けはカウント変更なし
            else:
                nstreak=0  # 早・遅・日・休でリセット
            if v!="休": work+=1
            if v=="早" and i>0:
                prev_v=ss[i-1]["shift_type"]
                prev_d=ss[i-1]["date"]
                gap=(datetime.strptime(cur_date,"%Y-%m-%d")-datetime.strptime(prev_d,"%Y-%m-%d")).days
                if prev_v=="明" and gap==1:
                    violations.append({"type":"error","staff_name":info["name"],"facility":info["fcode"],"unit":info["uname"],"date":cur_date,"rule":"夜勤明け翌日の早出禁止","message":f"{cur_date} 夜勤明け→早出"})
        if work>limit:
            violations.append({"type":"warning","staff_name":info["name"],"facility":info["fcode"],"unit":info["uname"],"date":None,"rule":"就業日数超過","message":f"就業{work}日/上限{limit}日"})
    return jsonify(violations)

# ── API: 過去シフト（Excel/CSV）の評価ベースライン ────────
@app.route("/api/shifts/baselines/import", methods=["POST"])
def api_import_baseline():
    """
    過去のシフト表(Excel .xlsx または CSV)をアップロードし、評価関数で
    スコアを算出してベースラインとして保存する。

    フォーム項目:
      facility_code: 事業所コード（必須）
      year, month:   対象年月（必須）
      file:          1行目=日付ヘッダ、1列目=職員名のシフト表（必須）

    注意: 過去データには希望休・希望シフトの記録がないため、
    「希望休・希望シフトの反映率」は満点(40点)固定で返る。
    ベースラインと自動生成案の比較は、人員不足/公平性/ルール違反の
    3項目で行うこと。
    """
    u, err = require_perm("perm_facility", "perm_gen", "perm_create")
    if err: return err
    facility_code = request.form.get("facility_code")
    year = request.form.get("year", type=int)
    month = request.form.get("month", type=int)
    file = request.files.get("file")
    if not (facility_code and year and month and file and file.filename):
        return jsonify({"ok": False, "error": "facility_code, year, month, file は必須です"}), 400

    f = qdb("SELECT id FROM facilities WHERE code=?", (facility_code,), one=True)
    if not f:
        return jsonify({"ok": False, "error": "facility not found"}), 404
    fid = f["id"]
    last_day = calendar.monthrange(year, month)[1]

    try:
        rows = _read_table_rows(file)
    except ValueError as e:
        return jsonify({"ok": False, "error": str(e)}), 400

    entries, unmatched_names, unknown_codes, matched_days = _parse_baseline_table(
        rows, fid, year, month, last_day)

    if not entries:
        return jsonify({
            "ok": False,
            "error": "シフトデータを読み取れませんでした。1行目に日付（または日番号）、"
                     "1列目に職員名が入っているかご確認ください。",
            "unmatched_names": unmatched_names,
            "matched_days": matched_days,
        }), 400

    staff_rows = qdb(
        """SELECT id,name,can_day,can_early,can_late,can_night,can_night_only,
                  monthly_limit,night_target,unit_id
           FROM staff WHERE facility_id=? AND is_active=1""", (fid,))
    staff = {s["id"]: dict(s) for s in staff_rows}
    units = [dict(u) for u in
             qdb("SELECT id,unit_no,name FROM units WHERE facility_id=? ORDER BY unit_no", (fid,))]

    evaluation = _evaluate_schedule(entries, staff, {}, units, year, month, last_day)
    evaluation["note"] = (
        "過去データには希望休・希望シフトの記録がないため、"
        "「希望休・希望シフトの反映率」は満点(40点)固定です。"
        "比較は人員不足・公平性・ルール違反の3項目で行ってください。"
    )

    db = get_db()
    existing = qdb("SELECT id FROM shift_baselines WHERE facility_id=? AND year=? AND month=?",
                    (fid, year, month), one=True)
    if existing:
        db.execute("DELETE FROM shift_baseline_entries WHERE baseline_id=?", (existing["id"],))
        db.execute("DELETE FROM shift_baselines WHERE id=?", (existing["id"],))

    cur = db.execute(
        """INSERT INTO shift_baselines(facility_id,year,month,source_name,eval_json,unmatched_names,unknown_codes)
           VALUES(?,?,?,?,?,?,?)""",
        (fid, year, month, file.filename, json.dumps(evaluation, ensure_ascii=False),
         json.dumps(unmatched_names, ensure_ascii=False), json.dumps(unknown_codes, ensure_ascii=False)))
    bid = cur.lastrowid
    db.executemany(
        "INSERT INTO shift_baseline_entries(baseline_id,staff_id,date,shift_type) VALUES(?,?,?,?)",
        [(bid, e["staff_id"], e["date"], e["shift_type"]) for e in entries])
    db.commit()

    return jsonify({
        "ok": True, "baseline_id": bid,
        "entries": len(entries), "days_matched": matched_days,
        "unmatched_names": unmatched_names, "unknown_codes": unknown_codes,
        "evaluation": evaluation,
    })


@app.route("/api/shifts/baselines")
def api_list_baselines():
    """事業所ごとのベースライン一覧（年月・スコア・取込時の警告）を返す"""
    u, err = require_any_login()
    if err: return err
    facility_code = request.args.get("facility")
    if not facility_code:
        return jsonify({"ok": False, "error": "facility is required"}), 400
    f = qdb("SELECT id FROM facilities WHERE code=?", (facility_code,), one=True)
    if not f:
        return jsonify([])
    rows = qdb(
        """SELECT id,year,month,source_name,eval_json,unmatched_names,unknown_codes,created_at
           FROM shift_baselines WHERE facility_id=? ORDER BY year,month""", (f["id"],))
    result = []
    for r in rows:
        ev = json.loads(r["eval_json"]) if r["eval_json"] else None
        result.append({
            "id": r["id"], "year": r["year"], "month": r["month"],
            "source_name": r["source_name"],
            "total_score": ev["total_score"] if ev else None,
            "scores": ev["scores"] if ev else None,
            "unmatched_names": json.loads(r["unmatched_names"]) if r["unmatched_names"] else [],
            "unknown_codes": json.loads(r["unknown_codes"]) if r["unknown_codes"] else {},
            "created_at": r["created_at"],
        })
    return jsonify(result)


@app.route("/api/shifts/baselines/<int:bid>")
def api_get_baseline(bid):
    """ベースライン1件の詳細（評価結果・取り込んだentries）を返す"""
    u, err = require_any_login()
    if err: return err
    row = qdb("SELECT * FROM shift_baselines WHERE id=?", (bid,), one=True)
    if not row:
        return jsonify({"ok": False, "error": "not found"}), 404
    entries = qdb(
        "SELECT staff_id,date,shift_type FROM shift_baseline_entries WHERE baseline_id=?", (bid,))
    return jsonify({
        "ok": True,
        "facility_id": row["facility_id"], "year": row["year"], "month": row["month"],
        "source_name": row["source_name"],
        "evaluation": json.loads(row["eval_json"]) if row["eval_json"] else None,
        "unmatched_names": json.loads(row["unmatched_names"]) if row["unmatched_names"] else [],
        "unknown_codes": json.loads(row["unknown_codes"]) if row["unknown_codes"] else {},
        "entries": [dict(e) for e in entries],
    })


@app.route("/api/shifts/baselines/<int:bid>", methods=["DELETE"])
def api_delete_baseline(bid):
    u, err = require_perm("perm_facility", "perm_gen", "perm_create")
    if err: return err
    db = get_db()
    db.execute("DELETE FROM shift_baseline_entries WHERE baseline_id=?", (bid,))
    db.execute("DELETE FROM shift_baselines WHERE id=?", (bid,))
    db.commit()
    return jsonify({"ok": True})

# ── API: 必要人員テーブル CRUD ──────────────────────────
@app.route("/api/units/<int:uid>/required_staff")
def api_get_required_staff(uid):
    """ユニットの必要人員設定を返す"""
    cur_user, err = require_any_login()
    if err: return err
    rows = qdb(
        "SELECT shift_type,required,is_admin_eligible FROM unit_required_staff WHERE unit_id=?", (uid,))
    return jsonify([dict(r) for r in rows])

@app.route("/api/units/<int:uid>/required_staff", methods=["PUT"])
def api_put_required_staff(uid):
    """
    ユニットの必要人員を一括更新する。
    リクエストボディ: [{shift_type, required, is_admin_eligible?}, ...]
    例: [{"shift_type":"早","required":1,"is_admin_eligible":1},
         {"shift_type":"遅","required":1},{"shift_type":"夜","required":1}]
    """
    cur_user, err = require_perm("perm_facility")
    if err: return err
    unit_row = qdb("SELECT id,name FROM units WHERE id=?", (uid,), one=True)
    if not unit_row:
        return jsonify({"ok": False, "error": "unit not found"}), 404
    items = request.get_json()
    if not isinstance(items, list):
        return jsonify({"ok": False, "error": "array required"}), 400
    db = get_db()
    db.execute("DELETE FROM unit_required_staff WHERE unit_id=?", (uid,))
    saved = []
    for item in items:
        st = item.get("shift_type")
        req = int(item.get("required", 1))
        admin_ok = int(item.get("is_admin_eligible", 0))
        if st not in ("早", "遅", "夜", "日"):
            continue
        if req < 0:
            continue
        db.execute(
            "INSERT INTO unit_required_staff(unit_id,shift_type,required,is_admin_eligible) VALUES(?,?,?,?)",
            (uid, st, req, admin_ok))
        saved.append({"shift_type": st, "required": req, "is_admin_eligible": admin_ok})
    db.commit()
    return jsonify({"ok": True, "unit_id": uid, "name": unit_row["name"], "saved": saved})

@app.route("/api/facilities/<fcode>/required_staff")
def api_get_facility_required_staff(fcode):
    """事業所全ユニットの必要人員設定一覧を返す（設定画面用）"""
    cur_user, err = require_any_login()
    if err: return err
    f = qdb("SELECT id FROM facilities WHERE code=?", (fcode,), one=True)
    if not f:
        return jsonify({"ok": False, "error": "facility not found"}), 404
    rows = qdb(
        """SELECT u.id as unit_id, u.unit_no, u.name as unit_name,
                  urs.shift_type, urs.required, urs.is_admin_eligible
           FROM units u
           LEFT JOIN unit_required_staff urs ON urs.unit_id=u.id
           WHERE u.facility_id=? AND u.residents > 0
           ORDER BY u.unit_no, urs.shift_type""", (f["id"],))
    by_unit = {}
    for r in rows:
        uid = r["unit_id"]
        if uid not in by_unit:
            by_unit[uid] = {"unit_id": uid, "unit_no": r["unit_no"],
                             "unit_name": r["unit_name"], "required": []}
        if r["shift_type"]:
            by_unit[uid]["required"].append({
                "shift_type": r["shift_type"],
                "required": r["required"],
                "is_admin_eligible": r["is_admin_eligible"],
            })
    return jsonify(list(by_unit.values()))


@app.route("/api/staff/<int:sid>/shift_targets")
def api_get_shift_targets(sid):
    """職員のシフト種別希望回数一覧を返す"""
    u, err = require_any_login()
    if err: return err
    rows = qdb("SELECT shift_type,target,is_upper_limit FROM staff_shift_targets WHERE staff_id=?", (sid,))
    return jsonify([dict(r) for r in rows])

@app.route("/api/staff/<int:sid>/shift_targets", methods=["PUT"])
def api_put_shift_targets(sid):
    """
    職員のシフト種別希望回数を一括更新する。
    リクエストボディ: [{shift_type, target, is_upper_limit?}, ...]
    例: [{"shift_type":"早","target":8},{"shift_type":"遅","target":4},{"shift_type":"夜","target":0}]
    target=0または未指定のshift_typeは削除(設定なし)として扱う。
    """
    u, err = require_perm("perm_staff")
    if err: return err
    s = qdb("SELECT id,name FROM staff WHERE id=?", (sid,), one=True)
    if not s:
        return jsonify({"ok": False, "error": "staff not found"}), 404
    items = request.get_json()
    if not isinstance(items, list):
        return jsonify({"ok": False, "error": "array required"}), 400

    db = get_db()
    db.execute("DELETE FROM staff_shift_targets WHERE staff_id=?", (sid,))
    saved = []
    for item in items:
        st = item.get("shift_type")
        tgt = int(item.get("target", 0))
        is_upper = int(item.get("is_upper_limit", 0))
        if st not in ("早", "遅", "夜", "日"):
            continue
        if tgt <= 0:
            continue  # 0は「設定なし」= 登録不要
        db.execute(
            "INSERT INTO staff_shift_targets(staff_id,shift_type,target,is_upper_limit) VALUES(?,?,?,?)",
            (sid, st, tgt, is_upper))
        saved.append({"shift_type": st, "target": tgt, "is_upper_limit": is_upper})
    db.commit()
    return jsonify({"ok": True, "saved": saved, "staff_id": sid, "name": s["name"]})

@app.route("/api/facilities/<fcode>/shift_targets")
def api_get_facility_shift_targets(fcode):
    """事業所全職員のシフト種別希望回数を一覧で返す（設定画面用）"""
    u, err = require_any_login()
    if err: return err
    f = qdb("SELECT id FROM facilities WHERE code=?", (fcode,), one=True)
    if not f:
        return jsonify({"ok": False, "error": "facility not found"}), 404
    rows = qdb(
        """SELECT s.id as staff_id, s.name, s.employment_type,
                  sst.shift_type, sst.target, sst.is_upper_limit
           FROM staff s
           LEFT JOIN staff_shift_targets sst ON sst.staff_id=s.id
           WHERE s.facility_id=? AND s.is_active=1
           ORDER BY s.id, sst.shift_type""", (f["id"],))
    by_staff = {}
    for r in rows:
        sid = r["staff_id"]
        if sid not in by_staff:
            by_staff[sid] = {"staff_id": sid, "name": r["name"],
                              "employment_type": r["employment_type"], "targets": []}
        if r["shift_type"]:
            by_staff[sid]["targets"].append({
                "shift_type": r["shift_type"], "target": r["target"],
                "is_upper_limit": r["is_upper_limit"],
            })
    return jsonify(list(by_staff.values()))

@app.route("/api/health")
def health():
    return jsonify({"status":"ok","time":datetime.now().isoformat()})

# ── API: シフト実績 ──────────────────────────
@app.route("/api/shifts/versions/<int:vid>/actual", methods=["GET"])
def api_get_actual(vid):
    u = get_session_staff()
    if not u:
        return jsonify({"ok": False, "error": "未ログインです"}), 401
    rows = qdb("SELECT * FROM shift_actual WHERE version_id=? ORDER BY date, staff_id", (vid,))
    return jsonify([dict(r) for r in rows])

@app.route("/api/shifts/versions/<int:vid>/actual", methods=["POST"])
def api_save_actual(vid):
    # 修正(2026-07-09): 他のシフト編集系（entries等）と揃え、
    # perm_create/perm_change必須にする（admin・リーダー・事務員ロール等）。
    u, err = require_perm("perm_create", "perm_change")
    if err: return err
    d = request.get_json() or {}
    staff_id    = d.get("staff_id")
    date        = d.get("date")
    actual_type = d.get("actual_type", "")
    time_from   = d.get("time_from", "")
    time_to     = d.get("time_to", "")
    work_hours  = float(d.get("work_hours", 0) or 0)
    note        = d.get("note", "")
    if not staff_id or not date or not actual_type:
        return jsonify({"ok": False, "error": "必須項目が不足しています"}), 400
    xdb(
        "INSERT INTO shift_actual(version_id,staff_id,date,actual_type,time_from,time_to,work_hours,note) "
        "VALUES(?,?,?,?,?,?,?,?) "
        "ON CONFLICT(version_id,staff_id,date) DO UPDATE SET "
        "actual_type=excluded.actual_type, time_from=excluded.time_from, "
        "time_to=excluded.time_to, work_hours=excluded.work_hours, "
        "note=excluded.note, created_at=datetime('now','localtime')",
        (vid, staff_id, date, actual_type, time_from, time_to, work_hours, note)
    )
    return jsonify({"ok": True})

@app.route("/api/shifts/versions/<int:vid>/actual/<int:staff_id>/<int:date>", methods=["DELETE"])
def api_delete_actual(vid, staff_id, date):
    # 修正(2026-07-09): 保存APIと同様にperm_create/perm_change必須にする。
    u, err = require_perm("perm_create", "perm_change")
    if err: return err
    ver = qdb("SELECT year, month FROM shift_versions WHERE id=?", (vid,), one=True)
    if not ver:
        return jsonify({"ok": False, "error": "version not found"}), 404
    date_str = f"{ver['year']:04d}-{ver['month']:02d}-{date:02d}"
    xdb("DELETE FROM shift_actual WHERE version_id=? AND staff_id=? AND date=?",
        (vid, staff_id, date_str))
    return jsonify({"ok": True})

@app.route("/api/shifts/versions/<int:vid>/validate")
def api_validate_shift(vid):
    """validate_shifts.pyのロジックをAPI経由で呼び出す。"""
    u, err = require_any_login()
    if err: return err
    try:
        import importlib.util, os
        vpath = os.path.join(os.path.dirname(os.path.abspath(__file__)), "validate_shifts.py")
        spec = importlib.util.spec_from_file_location("validate_shifts", vpath)
        vm = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(vm)

        db_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), DATABASE)
        conn = vm.connect(db_path)
        report = vm.validate_version(conn, vid)
        conn.close()

        checks = [{"name": name, "level": level, "ok": ok, "detail": detail or ""}
                  for name, level, ok, detail in report.checks]
        fail_count = sum(1 for c in checks if not c["ok"] and c["level"] == "fail")
        warn_count = sum(1 for c in checks if not c["ok"] and c["level"] == "warn")
        pass_count = sum(1 for c in checks if c["ok"])

        return jsonify({
            "ok": True, "version_id": vid, "checks": checks,
            "summary": {
                "all_ok": report.ok(),
                "fail_count": fail_count,
                "warn_count": warn_count,
                "pass_count": pass_count,
            }
        })
    except FileNotFoundError:
        return jsonify({"ok": False,
                        "error": "validate_shifts.py が見つかりません。app.pyと同じフォルダに配置してください。"}), 500
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500

if __name__=="__main__":
    init_db()
    print("\n"+"="*55)
    print("  GHShift v4")
    print("  http://127.0.0.1:5000          総合メニュー")
    print("  http://127.0.0.1:5000/facility 事業所登録")
    print("  http://127.0.0.1:5000/staff    職員登録")
    print("  http://127.0.0.1:5000/shift    シフト管理")
    print("="*55+"\n")
    # デバッグモードは既定でOFF。ローカルで一時的に使いたい場合のみ
    # 環境変数 CARESHIFT_DEBUG=1 を設定して起動する。
    # 本番（さくらVPS等）では絶対に有効にしないこと。
    _debug = os.environ.get("CARESHIFT_DEBUG") == "1"
    _host = os.environ.get("CARESHIFT_HOST", "127.0.0.1")
    app.run(debug=_debug, host=_host, port=5000)
