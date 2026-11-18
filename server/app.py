"""
NutraPLM 后端服务
- FastAPI + SQLite
- 账号登录 / 会话 Token / 角色权限（服务端强制）
- 实体库（原辅料、配方、产品 BOM、变更、项目、文档、供应商）+ 乐观并发提交
- 调价申请与多级审批（批准后自动生效并写入价格历史）
- BOM 版本快照与差异对比
- Excel(.xlsx) 导入解析
"""
import os, json, sqlite3, hashlib, secrets, time
from datetime import datetime, timedelta
from typing import Optional

from fastapi import FastAPI, Header, HTTPException, UploadFile, File, Body
from fastapi.responses import FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles

BASE = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.environ.get("PLM_DB", os.path.join(BASE, "plm.db"))
WEB_DIR = os.path.join(os.path.dirname(BASE), "web")
SEED_PATH = os.path.join(BASE, "seed.json")

TOKEN_TTL_DAYS = 7

ROLES = {
    "admin": "系统管理员",
    "rd": "研发工程师",
    "qa": "质量经理",
    "ra": "法规专员",
    "mfg": "生产主管",
    "viewer": "访客（只读）",
}

# 集合 -> 需要的权限
KIND_PERM = {
    "materials": "materials.write",
    "libs": "materials.write",
    "suppliers": "materials.write",
    "formulas": "formulas.write",
    "products": "products.write",
    "bomVersions": "products.write",
    "changes": "changes.write",
    "projects": "projects.write",
    "docs": "docs.write",
    "priceReqs": "price.request",
    "meta": "admin",
}

ROLE_PERMS = {
    "admin": ["*"],
    "rd": ["materials.write", "formulas.write", "products.write", "changes.write",
           "projects.write", "price.request", "import", "docs.write"],
    "qa": ["materials.write", "docs.write", "changes.write", "price.approve", "import"],
    "ra": ["docs.write", "changes.write", "price.approve"],
    "mfg": ["projects.write"],
    "viewer": [],
}


def now():
    return datetime.now().strftime("%Y-%m-%d %H:%M")


def today():
    return datetime.now().strftime("%Y-%m-%d")


# ---------------------------------------------------------------- 数据库
def conn():
    c = sqlite3.connect(DB_PATH)
    c.row_factory = sqlite3.Row
    c.execute("PRAGMA journal_mode=WAL")
    return c


def hash_pwd(pwd: str, salt: str) -> str:
    return hashlib.pbkdf2_hmac("sha256", pwd.encode(), salt.encode(), 120000).hex()


SCHEMA = """
CREATE TABLE IF NOT EXISTS users(
  id TEXT PRIMARY KEY, username TEXT UNIQUE NOT NULL, name TEXT, role TEXT,
  salt TEXT, pwd TEXT, active INTEGER DEFAULT 1, created_at TEXT);
CREATE TABLE IF NOT EXISTS sessions(
  token TEXT PRIMARY KEY, user_id TEXT, created_at TEXT, expires_at TEXT);
CREATE TABLE IF NOT EXISTS entities(
  kind TEXT, id TEXT, code TEXT, name TEXT, ord INTEGER, payload TEXT, updated_at TEXT,
  PRIMARY KEY(kind,id));
CREATE TABLE IF NOT EXISTS price_history(
  id INTEGER PRIMARY KEY AUTOINCREMENT, material_id TEXT, date TEXT,
  old_price REAL, new_price REAL, source TEXT, user TEXT, request_id TEXT);
CREATE TABLE IF NOT EXISTS price_requests(
  id TEXT PRIMARY KEY, code TEXT, material_id TEXT, old_price REAL, new_price REAL,
  reason TEXT, status TEXT, created_by TEXT, created_at TEXT, decided_at TEXT,
  approvals TEXT, history TEXT);
CREATE TABLE IF NOT EXISTS bom_versions(
  id TEXT PRIMARY KEY, product_id TEXT, ver TEXT, note TEXT, created_by TEXT,
  created_at TEXT, bom TEXT, cost REAL, unit_cost REAL);
CREATE TABLE IF NOT EXISTS audit_log(
  id INTEGER PRIMARY KEY AUTOINCREMENT, ts TEXT, user TEXT, role TEXT,
  action TEXT, kind TEXT, detail TEXT);
CREATE TABLE IF NOT EXISTS meta(k TEXT PRIMARY KEY, v TEXT);
"""

DEFAULT_USERS = [
    ("U0", "admin", "系统管理员", "admin", "admin123"),
    ("U1", "rd01", "张工", "rd", "rd123"),
    ("U2", "qa01", "陈工", "qa", "qa123"),
    ("U3", "ra01", "周工", "ra", "ra123"),
    ("U4", "mfg01", "赵工", "mfg", "mfg123"),
    ("U5", "viewer", "访客", "viewer", "view123"),
]


def init_db():
    fresh = not os.path.exists(DB_PATH)
    c = conn()
    c.executescript(SCHEMA)
    c.execute("INSERT OR IGNORE INTO meta(k,v) VALUES('version','1')")
    if fresh or c.execute("SELECT COUNT(*) n FROM users").fetchone()["n"] == 0:
        seed_all(c)
    c.commit()
    c.close()


def audit(c, user, role, action, kind="", detail=""):
    c.execute("INSERT INTO audit_log(ts,user,role,action,kind,detail) VALUES(?,?,?,?,?,?)",
              (now(), user, role, action, kind, detail))


def seed_all(c):
    data = json.load(open(SEED_PATH, encoding="utf-8"))

    for uid, uname, name, role, pwd in DEFAULT_USERS:
        salt = secrets.token_hex(8)
        c.execute("INSERT OR REPLACE INTO users(id,username,name,role,salt,pwd,active,created_at) VALUES(?,?,?,?,?,?,1,?)",
                  (uid, uname, name, role, salt, hash_pwd(pwd, salt), now()))

    # 实体
    kinds = ["libs", "materials", "suppliers", "formulas", "products", "changes", "projects", "docs"]
    for k in kinds:
        for i, item in enumerate(data.get(k, [])):
            c.execute("INSERT OR REPLACE INTO entities(kind,id,code,name,ord,payload,updated_at) VALUES(?,?,?,?,?,?,?)",
                      (k, str(item.get("id")), item.get("code", ""), item.get("name", "") or item.get("title", ""),
                       i, json.dumps(item, ensure_ascii=False), now()))
    c.execute("INSERT OR REPLACE INTO entities(kind,id,code,name,ord,payload,updated_at) VALUES('meta','meta','','',0,?,?)",
              (json.dumps(data.get("meta", {}), ensure_ascii=False), now()))

    # 价格历史（演示：近 6 个月 5 个点，末点为当前价）
    base = datetime.now()
    for m in data.get("materials", []):
        price = float(m.get("price") or 0)
        if price <= 0:
            continue
        hist = []
        for i, back in enumerate([150, 110, 70, 35]):
            d = (base - timedelta(days=back)).strftime("%Y-%m-%d")
            f = 0.90 + 0.035 * i
            hist.append((d, round(price * f, 2)))
        hist.append((base.strftime("%Y-%m-%d"), price))
        prev = None
        for d, p in hist:
            c.execute("INSERT INTO price_history(material_id,date,old_price,new_price,source,user,request_id) VALUES(?,?,?,?,?,?,?)",
                      (m["id"], d, prev, p, "期初导入" if prev is None else "供应商调价", "系统", ""))
            prev = p

    # 演示：一条审批中的调价单 + 一条已通过的
    m1 = next((m for m in data.get("materials", []) if m["id"] == "M001"), None)
    if m1:
        old = float(m1.get("price") or 0)
        c.execute("""INSERT INTO price_requests
          (id,code,material_id,old_price,new_price,reason,status,created_by,created_at,decided_at,approvals,history)
          VALUES(?,?,?,?,?,?,?,?,?,?,?,?)""",
                  ("PR001", "PR-2026-003", "M001", old, round(old * 1.06, 2),
                   "上游原料减产，供应商报价上调 6%，需评估对 FG-1001 成本影响后决定是否接受。",
                   "审批中", "张工", now(), "",
                   json.dumps([{"role": "质量", "user": "陈工", "state": "已通过", "date": now(), "cmt": "小样检验合格，同意调价"},
                               {"role": "法规", "user": "周工", "state": "待审批", "date": "", "cmt": ""}], ensure_ascii=False),
                   json.dumps([{"t": now(), "u": "张工", "a": "提交调价申请"}], ensure_ascii=False)))
        c.execute("""INSERT INTO price_requests
          (id,code,material_id,old_price,new_price,reason,status,created_by,created_at,decided_at,approvals,history)
          VALUES(?,?,?,?,?,?,?,?,?,?,?,?)""",
                  ("PR000", "PR-2026-001", "M005", 38, 32,
                   "年度框架协议重新议价，填充剂单价下调。", "已批准", "张工",
                   (base - timedelta(days=45)).strftime("%Y-%m-%d %H:%M"),
                   (base - timedelta(days=42)).strftime("%Y-%m-%d %H:%M"),
                   json.dumps([{"role": "质量", "user": "陈工", "state": "已通过", "date": (base - timedelta(days=43)).strftime("%Y-%m-%d %H:%M"), "cmt": "同意"},
                               {"role": "法规", "user": "周工", "state": "已通过", "date": (base - timedelta(days=42)).strftime("%Y-%m-%d %H:%M"), "cmt": "无合规影响"}], ensure_ascii=False),
                   json.dumps([{"t": (base - timedelta(days=45)).strftime("%Y-%m-%d %H:%M"), "u": "张工", "a": "提交调价申请"},
                               {"t": (base - timedelta(days=42)).strftime("%Y-%m-%d %H:%M"), "u": "系统", "a": "审批通过，单价 38 → 32 已生效"}], ensure_ascii=False)))

    # BOM 版本（P001 两个版本，用于对比演示）
    p1 = next((p for p in data.get("products", []) if p["id"] == "P001"), None)
    if p1:
        mats = {m["id"]: m for m in data.get("materials", [])}
        v1 = json.loads(json.dumps(p1["bom"]))
        for n in v1:
            if n.get("mid") == "M002":
                n["qty"] = 0.06
        v2 = json.loads(json.dumps(p1["bom"]))
        c.execute("INSERT INTO bom_versions(id,product_id,ver,note,created_by,created_at,bom,cost,unit_cost) VALUES(?,?,?,?,?,?,?,?,?)",
                  ("BV001", "P001", "V1.0", "初版 BOM（中间体 0.06kg/千片）", "张工",
                   (base - timedelta(days=90)).strftime("%Y-%m-%d %H:%M"), json.dumps(v1, ensure_ascii=False), 0, 0))
        c.execute("INSERT INTO bom_versions(id,product_id,ver,note,created_by,created_at,bom,cost,unit_cost) VALUES(?,?,?,?,?,?,?,?,?)",
                  ("BV002", "P001", "V2.0", "中间体下调至 0.05kg/千片，片重调整", "张工",
                   (base - timedelta(days=30)).strftime("%Y-%m-%d %H:%M"), json.dumps(v2, ensure_ascii=False), 0, 0))

    audit(c, "系统", "system", "初始化演示数据", "system", "创建用户、实体与价格历史")


# ---------------------------------------------------------------- 鉴权
def perms_of(role: str):
    p = ROLE_PERMS.get(role, [])
    return ["*"] if "*" in p else p


def has_perm(role: str, perm: str) -> bool:
    p = perms_of(role)
    return "*" in p or perm in p


def current_user(authorization: Optional[str]):
    if not authorization or not authorization.lower().startswith("bearer "):
        raise HTTPException(401, "未登录")
    token = authorization.split(" ", 1)[1].strip()
    c = conn()
    row = c.execute("""SELECT u.*, s.expires_at FROM sessions s JOIN users u ON u.id=s.user_id
                       WHERE s.token=?""", (token,)).fetchone()
    if not row:
        c.close()
        raise HTTPException(401, "会话已失效，请重新登录")
    if row["expires_at"] < now():
        c.close()
        raise HTTPException(401, "会话已过期，请重新登录")
    if not row["active"]:
        c.close()
        raise HTTPException(403, "账号已停用")
    c.close()
    return {"id": row["id"], "username": row["username"], "name": row["name"], "role": row["role"]}


# ---------------------------------------------------------------- 应用
app = FastAPI(title="NutraPLM API", version="3.0")


@app.post("/api/login")
def login(body: dict = Body(...)):
    username = (body.get("username") or "").strip()
    pwd = body.get("password") or ""
    c = conn()
    u = c.execute("SELECT * FROM users WHERE username=?", (username,)).fetchone()
    if not u or not u["active"] or hash_pwd(pwd, u["salt"]) != u["pwd"]:
        c.close()
        raise HTTPException(401, "账号或密码不正确")
    token = secrets.token_urlsafe(32)
    exp = (datetime.now() + timedelta(days=TOKEN_TTL_DAYS)).strftime("%Y-%m-%d %H:%M")
    c.execute("INSERT INTO sessions(token,user_id,created_at,expires_at) VALUES(?,?,?,?)",
              (token, u["id"], now(), exp))
    audit(c, u["name"], u["role"], "登录", "auth", f"{u['username']} 登录成功")
    c.commit()
    c.close()
    return {"token": token, "user": {"id": u["id"], "username": u["username"], "name": u["name"],
                                     "role": u["role"], "roleName": ROLES.get(u["role"], u["role"]),
                                     "perms": perms_of(u["role"])}}


@app.post("/api/logout")
def logout(authorization: Optional[str] = Header(None)):
    user = current_user(authorization)
    token = authorization.split(" ", 1)[1].strip()
    c = conn()
    c.execute("DELETE FROM sessions WHERE token=?", (token,))
    audit(c, user["name"], user["role"], "退出登录", "auth")
    c.commit()
    c.close()
    return {"ok": True}


@app.get("/api/me")
def me(authorization: Optional[str] = Header(None)):
    u = current_user(authorization)
    return {**u, "roleName": ROLES.get(u["role"], u["role"]), "perms": perms_of(u["role"])}


@app.get("/api/snapshot")
def snapshot(authorization: Optional[str] = Header(None)):
    u = current_user(authorization)
    c = conn()
    db = {}
    for k in ["libs", "materials", "suppliers", "formulas", "products", "changes", "projects", "docs"]:
        rows = c.execute("SELECT payload FROM entities WHERE kind=? ORDER BY ord", (k,)).fetchall()
        db[k] = [json.loads(r["payload"]) for r in rows]
    meta = c.execute("SELECT payload FROM entities WHERE kind='meta'").fetchone()
    db["meta"] = json.loads(meta["payload"]) if meta else {}
    db["log"] = [dict(r) for r in c.execute(
        "SELECT ts AS t, user AS u, action AS a, kind AS k FROM audit_log ORDER BY id DESC LIMIT 40").fetchall()]
    db["priceReqs"] = []
    for r in c.execute("SELECT * FROM price_requests ORDER BY created_at DESC").fetchall():
        db["priceReqs"].append({
            "id": r["id"], "code": r["code"], "mid": r["material_id"], "oldPrice": r["old_price"],
            "newPrice": r["new_price"], "reason": r["reason"], "status": r["status"],
            "createdBy": r["created_by"], "createdAt": r["created_at"], "decidedAt": r["decided_at"],
            "approvals": json.loads(r["approvals"] or "[]"), "history": json.loads(r["history"] or "[]")})
    hist = {}
    for r in c.execute("SELECT * FROM price_history ORDER BY date").fetchall():
        hist.setdefault(r["material_id"], []).append({"date": r["date"], "price": r["new_price"],
                                                      "src": r["source"], "user": r["user"]})
    db["priceHist"] = hist
    db["bomVersions"] = [dict(r) for r in c.execute(
        "SELECT id,product_id AS productId,ver,note,created_by AS createdBy,created_at AS createdAt,cost,unit_cost AS unitCost FROM bom_versions ORDER BY created_at DESC").fetchall()]
    ver = int(c.execute("SELECT v FROM meta WHERE k='version'").fetchone()["v"])
    users = [{"id": r["id"], "username": r["username"], "name": r["name"], "role": r["role"],
              "roleName": ROLES.get(r["role"], r["role"]), "active": bool(r["active"])}
             for r in c.execute("SELECT * FROM users ORDER BY id").fetchall()]
    c.close()
    return {"version": ver, "db": db, "users": users, "user": u,
            "roleName": ROLES.get(u["role"], u["role"]), "perms": perms_of(u["role"]), "serverTime": now()}


@app.post("/api/commit")
def commit(body: dict = Body(...), authorization: Optional[str] = Header(None)):
    """整库提交：服务端按集合做权限校验 + 冲突检测 + 审计。"""
    u = current_user(authorization)
    incoming = body.get("db") or {}
    base_version = int(body.get("version") or 0)
    c = conn()
    ver = int(c.execute("SELECT v FROM meta WHERE k='version'").fetchone()["v"])
    if base_version and base_version != ver:
        c.close()
        return JSONResponse(status_code=409, content={"error": "conflict", "version": ver,
                                                     "message": "数据已被他人更新，请刷新后重试"})

    changed = []
    for kind, perm in KIND_PERM.items():
        if kind in ("priceReqs", "meta", "bomVersions"):
            continue
        rows = c.execute("SELECT id,payload FROM entities WHERE kind=?", (kind,)).fetchall()
        old = {r["id"]: json.loads(r["payload"]) for r in rows}
        new_items = incoming.get(kind)
        if not isinstance(new_items, list):
            continue
        new = {str(x.get("id")): x for x in new_items}
        diff_ids = [i for i in set(old) | set(new)
                    if json.dumps(old.get(i), sort_keys=True) != json.dumps(new.get(i), sort_keys=True)]
        if diff_ids:
            if not has_perm(u["role"], perm):
                c.close()
                raise HTTPException(403, f"当前角色（{ROLES.get(u['role'])}）无权修改「{kind}」")
            changed.append((kind, diff_ids, old, new_items))

    now_s = now()
    for kind, diff_ids, old, new_items in changed:
        c.execute("DELETE FROM entities WHERE kind=?", (kind,))
        for i, item in enumerate(new_items):
            c.execute("INSERT OR REPLACE INTO entities(kind,id,code,name,ord,payload,updated_at) VALUES(?,?,?,?,?,?,?)",
                      (kind, str(item.get("id")), item.get("code", ""),
                       item.get("name", "") or item.get("title", ""), i,
                       json.dumps(item, ensure_ascii=False), now_s))
    if incoming.get("meta"):
        c.execute("INSERT OR REPLACE INTO entities(kind,id,code,name,ord,payload,updated_at) VALUES('meta','meta','','',0,?,?)",
                  (json.dumps(incoming["meta"], ensure_ascii=False), now_s))

    for kind, diff_ids, old, new_items in changed:
        add = len([i for i in diff_ids if i not in old])
        rem = len([i for i in diff_ids if i not in {str(x.get("id")) for x in new_items}])
        upd = len(diff_ids) - add - rem
        detail = f"新增 {add} / 更新 {upd} / 删除 {rem}" if (add or rem) else f"更新 {upd} 条"
        audit(c, u["name"], u["role"], "保存数据", kind, detail)

    new_ver = ver + 1
    c.execute("UPDATE meta SET v=? WHERE k='version'", (str(new_ver),))
    c.commit()
    c.close()
    return {"version": new_ver, "changed": [k for k, *_ in changed]}


# ---------------------------------------------------------------- 调价审批
def _req_row(c, rid):
    r = c.execute("SELECT * FROM price_requests WHERE id=?", (rid,)).fetchone()
    if not r:
        raise HTTPException(404, "调价单不存在")
    return r


@app.post("/api/price-requests")
def create_price_request(body: dict = Body(...), authorization: Optional[str] = Header(None)):
    u = current_user(authorization)
    if not has_perm(u["role"], "price.request"):
        raise HTTPException(403, "当前角色无权发起调价")
    mid = body.get("mid")
    new_price = float(body.get("newPrice") or 0)
    reason = (body.get("reason") or "").strip()
    if not mid or new_price <= 0 or not reason:
        raise HTTPException(400, "请填写物料、新单价与调价原因")

    c = conn()
    m = c.execute("SELECT payload FROM entities WHERE kind='materials' AND id=?", (mid,)).fetchone()
    if not m:
        c.close()
        raise HTTPException(404, "物料不存在")
    mat = json.loads(m["payload"])
    old_price = float(mat.get("price") or 0)
    if abs(new_price - old_price) < 1e-9:
        c.close()
        raise HTTPException(400, "新单价与当前单价相同")
    seq = c.execute("SELECT COUNT(*) n FROM price_requests").fetchone()["n"] + 1
    rid = "PR" + secrets.token_hex(4)
    code = f"PR-{datetime.now().strftime('%Y')}-{seq:03d}"
    approvals = [{"role": "质量", "user": "陈工", "state": "待审批", "date": "", "cmt": ""},
                 {"role": "法规", "user": "周工", "state": "待审批", "date": "", "cmt": ""}]
    history = [{"t": now(), "u": u["name"], "a": f"提交调价申请：{old_price:g} → {new_price:g} 元"}]
    c.execute("""INSERT INTO price_requests(id,code,material_id,old_price,new_price,reason,status,created_by,created_at,decided_at,approvals,history)
                 VALUES(?,?,?,?,?,?,?,?,?,?,?,?)""",
              (rid, code, mid, old_price, new_price, reason, "审批中", u["name"], now(), "",
               json.dumps(approvals, ensure_ascii=False), json.dumps(history, ensure_ascii=False)))
    audit(c, u["name"], u["role"], "发起调价", "price", f"{code} {mat.get('code')} {old_price:g}→{new_price:g}")
    c.commit()
    c.close()
    return {"id": rid, "code": code}


@app.post("/api/price-requests/{rid}/decide")
def decide_price_request(rid: str, body: dict = Body(...), authorization: Optional[str] = Header(None)):
    u = current_user(authorization)
    if not has_perm(u["role"], "price.approve"):
        raise HTTPException(403, "当前角色无审批权限")
    action = body.get("action")
    cmt = (body.get("comment") or "").strip()
    if action not in ("approve", "reject"):
        raise HTTPException(400, "参数错误")
    c = conn()
    r = _req_row(c, rid)
    if r["status"] in ("已批准", "已驳回"):
        c.close()
        raise HTTPException(400, "该调价单已结束")

    approvals = json.loads(r["approvals"] or "[]")
    history = json.loads(r["history"] or "[]")
    role_label = {"qa": "质量", "ra": "法规", "admin": "管理员"}.get(u["role"], u["role"])
    step = next((a for a in approvals if a["state"] == "待审批" and (a["role"] == role_label or u["role"] == "admin")), None)
    if not step:
        c.close()
        raise HTTPException(403, f"当前审批节点不属于你的角色（下一节点：{next((a['role'] for a in approvals if a['state']=='待审批'), '无')}）")

    step["state"] = "已通过" if action == "approve" else "已驳回"
    step["date"] = now()
    step["cmt"] = cmt
    step["user"] = u["name"]
    history.append({"t": now(), "u": u["name"], "a": f"{step['role']}审批{'通过' if action=='approve' else '驳回'}" + (f"：{cmt}" if cmt else "")})

    status = r["status"]
    decided = ""
    if action == "reject":
        status = "已驳回"
        decided = now()
        history.append({"t": now(), "u": "系统", "a": "调价申请已驳回，原单价保持不变"})
    else:
        if all(a["state"] == "已通过" for a in approvals):
            status = "已批准"
            decided = now()
            mat_row = c.execute("SELECT payload FROM entities WHERE kind='materials' AND id=?", (r["material_id"],)).fetchone()
            if mat_row:
                mat = json.loads(mat_row["payload"])
                mat["price"] = r["new_price"]
                mat["date"] = today()
                c.execute("UPDATE entities SET payload=?, updated_at=? WHERE kind='materials' AND id=?",
                          (json.dumps(mat, ensure_ascii=False), now(), r["material_id"]))
            c.execute("""INSERT INTO price_history(material_id,date,old_price,new_price,source,user,request_id)
                         VALUES(?,?,?,?,?,?,?)""",
                      (r["material_id"], today(), r["old_price"], r["new_price"], "调价审批通过", u["name"], rid))
            history.append({"t": now(), "u": "系统", "a": f"审批通过，单价 {r['old_price']:g} → {r['new_price']:g} 已生效"})
            c.execute("UPDATE meta SET v=CAST(v AS INTEGER)+1 WHERE k='version'")
        else:
            status = "审批中"

    c.execute("UPDATE price_requests SET status=?, approvals=?, history=?, decided_at=? WHERE id=?",
              (status, json.dumps(approvals, ensure_ascii=False), json.dumps(history, ensure_ascii=False), decided, rid))
    audit(c, u["name"], u["role"], "调价审批" + ("通过" if action == "approve" else "驳回"), "price", r["code"])
    c.commit()
    c.close()
    return {"status": status}


@app.get("/api/price-history/{mid}")
def price_history(mid: str, authorization: Optional[str] = Header(None)):
    current_user(authorization)
    c = conn()
    rows = c.execute("SELECT date,old_price,new_price,source,user FROM price_history WHERE material_id=? ORDER BY date", (mid,)).fetchall()
    c.close()
    return {"items": [{"date": r["date"], "price": r["new_price"], "src": r["source"], "user": r["user"]} for r in rows]}


# ---------------------------------------------------------------- BOM 版本
def bom_cost(bom, mats):
    def unit(node):
        if node.get("children"):
            return sum(unit(c) * (float(c.get("qty") or 0)) for c in node["children"]) * (1 + float(node.get("loss") or 0) / 100)
        m = mats.get(node.get("mid"))
        return float((m or {}).get("price") or 0)

    def line(node):
        v = float(node.get("qty") or 0) * unit(node)
        return v
    return sum(line(n) for n in bom)


@app.post("/api/bom/{pid}/snapshot")
def bom_snapshot(pid: str, body: dict = Body(...), authorization: Optional[str] = Header(None)):
    u = current_user(authorization)
    if not has_perm(u["role"], "products.write"):
        raise HTTPException(403, "无权存档 BOM")
    c = conn()
    p = c.execute("SELECT payload FROM entities WHERE kind='products' AND id=?", (pid,)).fetchone()
    if not p:
        c.close()
        raise HTTPException(404, "产品不存在")
    prod = json.loads(p["payload"])
    mats = {r["id"]: json.loads(r["payload"]) for r in
            c.execute("SELECT id,payload FROM entities WHERE kind='materials'").fetchall()}
    cost = bom_cost(prod.get("bom") or [], mats)
    n = c.execute("SELECT COUNT(*) n FROM bom_versions WHERE product_id=?", (pid,)).fetchone()["n"] + 1
    ver = f"V{n}.0"
    vid = "BV" + secrets.token_hex(4)
    c.execute("""INSERT INTO bom_versions(id,product_id,ver,note,created_by,created_at,bom,cost,unit_cost)
                 VALUES(?,?,?,?,?,?,?,?,?)""",
              (vid, pid, ver, (body.get("note") or "").strip() or "手动存档", u["name"], now(),
               json.dumps(prod.get("bom") or [], ensure_ascii=False), round(cost, 2), 0))
    audit(c, u["name"], u["role"], "BOM 存档", "product", f"{prod.get('code')} {ver}")
    c.commit()
    c.close()
    return {"id": vid, "ver": ver, "cost": round(cost, 2)}


@app.get("/api/bom/{pid}/versions")
def bom_versions(pid: str, authorization: Optional[str] = Header(None)):
    current_user(authorization)
    c = conn()
    rows = c.execute("SELECT id,ver,note,created_by,created_at,bom,cost FROM bom_versions WHERE product_id=? ORDER BY created_at", (pid,)).fetchall()
    c.close()
    return {"items": [{"id": r["id"], "ver": r["ver"], "note": r["note"], "createdBy": r["created_by"],
                       "createdAt": r["created_at"], "bom": json.loads(r["bom"]), "cost": r["cost"]} for r in rows]}


# ---------------------------------------------------------------- Excel 导入
@app.post("/api/import/xlsx")
async def import_xlsx(file: UploadFile = File(...), authorization: Optional[str] = Header(None)):
    u = current_user(authorization)
    if not has_perm(u["role"], "import"):
        raise HTTPException(403, "当前角色无导入权限")
    raw = await file.read()
    try:
        import io
        from openpyxl import load_workbook
        wb = load_workbook(io.BytesIO(raw), read_only=True, data_only=True)
    except Exception as e:
        raise HTTPException(400, f"无法解析 Excel 文件：{e}")
    ws = wb[wb.sheetnames[0]]
    rows, empty_streak = [], 0
    for row in ws.iter_rows(values_only=True):
        vals = ["" if v is None else str(v).strip() for v in row]
        if not any(vals):
            empty_streak += 1
            if empty_streak > 5:
                break
            continue
        empty_streak = 0
        rows.append(vals)
        if len(rows) >= 500:
            break
    wb.close()
    if not rows:
        raise HTTPException(400, "表格内容为空")
    width = max(len(r) for r in rows)
    rows = [r + [""] * (width - len(r)) for r in rows]
    audit_conn = conn()
    audit(audit_conn, u["name"], u["role"], "Excel 导入解析", "material", f"{file.filename} · {len(rows)-1} 行")
    audit_conn.commit()
    audit_conn.close()
    return {"filename": file.filename, "sheet": ws.title, "rows": rows}


# ---------------------------------------------------------------- 用户与审计（管理员）
@app.get("/api/users")
def list_users(authorization: Optional[str] = Header(None)):
    u = current_user(authorization)
    if u["role"] != "admin":
        raise HTTPException(403, "仅管理员可查看用户")
    c = conn()
    rows = c.execute("SELECT id,username,name,role,active FROM users ORDER BY id").fetchall()
    c.close()
    return {"items": [{"id": r["id"], "username": r["username"], "name": r["name"],
                       "role": r["role"], "roleName": ROLES.get(r["role"], r["role"]),
                       "active": bool(r["active"])} for r in rows]}


@app.post("/api/users")
def create_user(body: dict = Body(...), authorization: Optional[str] = Header(None)):
    u = current_user(authorization)
    if u["role"] != "admin":
        raise HTTPException(403, "仅管理员可新建用户")
    uname = (body.get("username") or "").strip()
    pwd = body.get("password") or ""
    role = body.get("role") or "viewer"
    name = (body.get("name") or uname).strip()
    if not uname or len(pwd) < 6:
        raise HTTPException(400, "账号必填，密码至少 6 位")
    if role not in ROLES:
        raise HTTPException(400, "角色不合法")
    c = conn()
    if c.execute("SELECT 1 FROM users WHERE username=?", (uname,)).fetchone():
        c.close()
        raise HTTPException(400, "账号已存在")
    uid = "U" + secrets.token_hex(4)
    salt = secrets.token_hex(8)
    c.execute("INSERT INTO users(id,username,name,role,salt,pwd,active,created_at) VALUES(?,?,?,?,?,?,1,?)",
              (uid, uname, name, role, salt, hash_pwd(pwd, salt), now()))
    audit(c, u["name"], u["role"], "新建用户", "user", f"{uname} / {ROLES.get(role)}")
    c.commit()
    c.close()
    return {"id": uid}


@app.patch("/api/users/{uid}")
def update_user(uid: str, body: dict = Body(...), authorization: Optional[str] = Header(None)):
    u = current_user(authorization)
    if u["role"] != "admin":
        raise HTTPException(403, "仅管理员可修改用户")
    c = conn()
    row = c.execute("SELECT * FROM users WHERE id=?", (uid,)).fetchone()
    if not row:
        c.close()
        raise HTTPException(404, "用户不存在")
    role = body.get("role", row["role"])
    active = 1 if body.get("active", bool(row["active"])) else 0
    if role not in ROLES:
        c.close()
        raise HTTPException(400, "角色不合法")
    if uid == u["id"] and (active == 0 or role != "admin"):
        c.close()
        raise HTTPException(400, "不能停用或降级自己")
    c.execute("UPDATE users SET role=?, active=? WHERE id=?", (role, active, uid))
    if body.get("password"):
        salt = secrets.token_hex(8)
        c.execute("UPDATE users SET salt=?, pwd=? WHERE id=?", (salt, hash_pwd(body["password"], salt), uid))
    audit(c, u["name"], u["role"], "修改用户", "user", f"{row['username']} → {ROLES.get(role)} / {'启用' if active else '停用'}")
    c.commit()
    c.close()
    return {"ok": True}


@app.get("/api/audit")
def get_audit(limit: int = 100, authorization: Optional[str] = Header(None)):
    current_user(authorization)
    c = conn()
    rows = c.execute("SELECT ts,user,role,action,kind,detail FROM audit_log ORDER BY id DESC LIMIT ?", (min(limit, 500),)).fetchall()
    c.close()
    return {"items": [dict(r) for r in rows]}


@app.get("/api/health")
def health():
    return {"ok": True, "time": now()}


@app.get("/api/version")
def version(authorization: Optional[str] = Header(None)):
    current_user(authorization)
    c = conn()
    v = int(c.execute("SELECT v FROM meta WHERE k='version'").fetchone()["v"])
    c.close()
    return {"version": v}


# ---------------------------------------------------------------- 静态资源
@app.get("/")
def index():
    return FileResponse(os.path.join(WEB_DIR, "app.html"))


@app.get("/app.html")
def index2():
    return FileResponse(os.path.join(WEB_DIR, "app.html"))


app.mount("/", StaticFiles(directory=WEB_DIR, html=True), name="web")

init_db()
