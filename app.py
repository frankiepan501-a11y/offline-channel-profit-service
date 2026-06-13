# -*- coding: utf-8 -*-
"""国内分销 c渠道 月度重算云服务 (Zeabur)。
按钮触发 → POST /recompute?month=2026-05 → 下载上传台账单 → 顺丰API+账单匹配 →
订单级分摊填物流 → 完整度 → 净毛利报表 → 飞书汇报 Frankie。
密钥全走 env: FEISHU_APP_ID/FEISHU_APP_SECRET/SF_PARTNER/SF_CHECKWORD/AUTH_TOKEN
"""
import os, io, json, time, hashlib, base64, datetime, tempfile
from collections import defaultdict
import requests, openpyxl
from fastapi import FastAPI, Request, HTTPException

FEISHU_APP_ID = os.environ["FEISHU_APP_ID"]
FEISHU_APP_SECRET = os.environ["FEISHU_APP_SECRET"]
SF_CC = os.environ.get("SF_PARTNER", "ADEDZLPVZYMO")
SF_CW = os.environ["SF_CHECKWORD"]
AUTH_TOKEN = os.environ.get("AUTH_TOKEN", "")
APP = "JqZwbSi7uaDlw0sjEFPcTDlenMf"; O = "tblJ7Z9cUGTz8fsu"; UP = "tblSdk9uVmDRgzJq"
DX = "tblgTajhPMenNmI6"; PROD = "tbl4d30zXJRFQbPD"; FRANKIE = "ou_629ce01f4bc31de078e10fcb038dbf78"
FIN_APP = "P9awbhG9faFstxsO1KZc9b9Qnxb"; OVERVIEW_TBL = "tbltFK8vwdcrlfBa"  # 全渠道销售总览
FEISHU = "https://open.feishu.cn/open-apis"
app = FastAPI()

def tok():
    r = requests.post(f"{FEISHU}/auth/v3/tenant_access_token/internal",
                      json={"app_id": FEISHU_APP_ID, "app_secret": FEISHU_APP_SECRET}, timeout=20)
    return r.json()["tenant_access_token"]

def gv(f, k):
    v = f.get(k)
    if v is None: return ""
    if isinstance(v, list): return " ".join(str(x.get("text", x.get("value", x)) if isinstance(x, dict) else x) for x in v)
    if isinstance(v, dict): return str(v.get("text") or v.get("value") or "")
    return str(v)

def num(x):
    try: return float(x)
    except: return 0.0

def dg(s): return "".join(c for c in str(s) if c.isdigit())

def getall(T, tbl):
    items = []; pt = None
    while True:
        u = f"{FEISHU}/bitable/v1/apps/{APP}/tables/{tbl}/records?page_size=500" + (f"&page_token={pt}" if pt else "")
        d = requests.get(u, headers={"Authorization": f"Bearer {T}"}, timeout=30).json()["data"]
        items += (d.get("items") or []); pt = d.get("page_token")
        if not d.get("has_more"): break
    return items

def download_bills(T, month):
    mk = dg(month); out = {"中通": [], "顺丰": []}
    for r in getall(T, UP):
        f = r["fields"]
        if dg(gv(f, "账单月份(如2026-05)")) != mk: continue
        carrier = gv(f, "物流商")
        for a in (f.get("账单文件") or []):
            ftok = a.get("file_token")
            if not ftok: continue
            try:
                resp = requests.get(f"{FEISHU}/drive/v1/medias/{ftok}/download",
                                    headers={"Authorization": f"Bearer {T}"}, timeout=60)
                p = os.path.join(tempfile.gettempdir(), f"{mk}_{carrier}_{ftok}.xlsx")
                open(p, "wb").write(resp.content)
                out.setdefault(carrier, []).append(p)
            except Exception as e:
                print("download fail", e)
    return out

def load_zto(paths):
    d = {}
    for fp in paths:
        try:
            wb = openpyxl.load_workbook(fp, read_only=True, data_only=True); ws = wb.worksheets[0]; hdr = None
            for r in ws.iter_rows(values_only=True):
                vals = [("" if v is None else str(v).strip()) for v in r]
                if hdr is None: hdr = vals; continue
                row = dict(zip(hdr, vals)); wn = row.get("运单号", "").strip(); fee = row.get("合计") or row.get("金额") or ""
                if wn and fee:
                    try: d[wn] = float(fee)
                    except: pass
            wb.close()
        except Exception as e: print("zto", e)
    return d

def load_sf(paths):
    d = {}
    for fp in paths:
        try:
            wb = openpyxl.load_workbook(fp, data_only=True)
            ws = wb["账单明细"] if "账单明细" in wb.sheetnames else wb.worksheets[0]
            for rr in range(3, ws.max_row + 1):
                wn = ws.cell(rr, 3).value; fee = ws.cell(rr, 12).value
                if wn and fee not in (None, ""):
                    try: d[str(wn).strip()] = float(fee)
                    except: pass
            wb.close()
        except Exception as e: print("sf", e)
    return d

_sf = {}
def sf_freight(wb, sfx):
    if wb in _sf: return _sf[wb]
    v = None
    try:
        msg = json.dumps({"trackingNum": wb, "trackingType": "2"}, ensure_ascii=False, separators=(",", ":"))
        ts = str(int(time.time() * 1000)); dig = base64.b64encode(hashlib.md5((msg + ts + SF_CW).encode()).digest()).decode()
        form = {"partnerID": SF_CC, "requestID": f"r{ts}", "serviceCode": "EXP_RECE_QUERY_SFWAYBILL",
                "timestamp": ts, "msgDigest": dig, "msgData": msg}
        res = requests.post("https://sfapi.sf-express.com/std/service", data=form, timeout=20).json()
        if res.get("apiResultCode") == "A1000":
            data = res.get("apiResultData"); data = json.loads(data) if isinstance(data, str) else data
            fees = (data.get("msgData") or {}).get("waybillFeeList") or []
            if fees: v = round(sum(float(f.get("value", 0)) for f in fees), 2)
    except Exception as e: print("sfapi", e)
    time.sleep(0.15)
    if v is None and wb in sfx: v = sfx[wb]
    _sf[wb] = v; return v

def report(T, month):
    """净毛利报表(B口径: 买断+赠样 from 订单台, 代销/寄售 from 动销表, 铺货运费渠道扣)"""
    od = getall(T, O); dx = getall(T, DX); mk = dg(month)
    rows = []
    for r in od:
        f = r["fields"]; d = gv(f, "下单/出货日期")
        if not d.isdigit(): continue
        if datetime.datetime.utcfromtimestamp(int(d) / 1000).strftime("%Y-%m") != month: continue
        way = gv(f, "合作方式"); qty = num(gv(f, "数量")); cg = num(gv(f, "单位成本(自动)")); wl = num(gv(f, "物流成本(待接)"))
        if way in ("经销买断", "赠样"):
            rev = num(gv(f, "订单金额")) if way == "经销买断" else 0
            ml = num(gv(f, "毛利(自动)")) if gv(f, "毛利(自动)") else (rev - qty * cg - wl)
            rows.append({"cust": gv(f, "关联经销商"), "prod": gv(f, "产品名"), "rev": rev, "wl": wl, "ml": ml})
        elif way in ("代销月结", "寄售") and wl:
            rows.append({"cust": gv(f, "关联经销商"), "prod": gv(f, "产品名"), "rev": 0, "wl": wl, "ml": -wl})
    for r in dx:
        f = r["fields"]
        if dg(gv(f, "动销月份")) != mk: continue
        rev = num(gv(f, "动销金额")); ml = num(gv(f, "动销毛利(自动)"))
        rows.append({"cust": gv(f, "关联经销商"), "prod": gv(f, "产品名"), "rev": rev, "wl": num(gv(f, "物流成本(待接)")), "ml": ml})
    tr = sum(x["rev"] for x in rows); tm = sum(x["ml"] for x in rows); tw = sum(x["wl"] for x in rows)
    C = defaultdict(lambda: [0, 0])
    for x in rows:
        c = C[x["cust"]]; c[0] += x["rev"]; c[1] += x["ml"]
    top = sorted(C.items(), key=lambda x: -x[1][1])[:6]
    lines = [f"总收入 {tr:.0f} | 净毛利 {tm:.0f} ({tm/tr*100 if tr else 0:.1f}%) | 物流 {tw:.0f}"]
    for c, v in top:
        lines.append(f"  {c}: 收入{v[0]:.0f} 净毛利{v[1]:.0f} ({v[1]/v[0]*100 if v[0] else 0:.1f}%)")
    return "\n".join(lines), {"rev": tr, "ml": tm, "wl": tw}


def upsert_overview(T, month, t):
    """灌全渠道总表 c国内线下汇总行(幂等, 对齐现有行格式: 销售/物流/毛利/毛利率)。"""
    if t["rev"] <= 0 and t["ml"] == 0:
        return "skip(无数据)"
    fields = {"月份": month, "渠道大类": "国内线下", "平台": "线下",
              "店铺": "经销分销汇总(买断+代销/寄售+赠样)",
              "销售额RMB": round(t["rev"], 2), "物流费RMB": round(t["wl"], 2),
              "全额毛利RMB": round(t["ml"], 2),
              "毛利率": round(t["ml"] / t["rev"], 4) if t["rev"] else 0}
    found = None
    for rec in _ovw_recs(T):   # 总表在 FIN_APP, getall() 用的是c渠道APP, 故单独拉
        f = rec["fields"]
        if gv(f, "月份") == month and gv(f, "平台") == "线下":
            found = rec["record_id"]; break
    H = {"Authorization": f"Bearer {T}", "Content-Type": "application/json"}
    if found:
        requests.put(f"{FEISHU}/bitable/v1/apps/{FIN_APP}/tables/{OVERVIEW_TBL}/records/{found}",
                     headers=H, json={"fields": fields}, timeout=30)
        return "updated"
    requests.post(f"{FEISHU}/bitable/v1/apps/{FIN_APP}/tables/{OVERVIEW_TBL}/records",
                  headers=H, json={"fields": fields}, timeout=30)
    return "created"


def _ovw_recs(T):
    items = []; pt = None
    while True:
        u = f"{FEISHU}/bitable/v1/apps/{FIN_APP}/tables/{OVERVIEW_TBL}/records?page_size=500" + (f"&page_token={pt}" if pt else "")
        d = requests.get(u, headers={"Authorization": f"Bearer {T}"}, timeout=30).json().get("data", {})
        items += (d.get("items") or []); pt = d.get("page_token")
        if not d.get("has_more"): break
    return items

def do_recompute(month):
    T = tok()
    bills = download_bills(T, month)
    zto = load_zto(bills.get("中通", [])); sfx = load_sf(bills.get("顺丰", []))
    od = getall(T, O)
    orders = defaultdict(list)
    for r in od:
        f = r["fields"]; d = gv(f, "下单/出货日期")
        if not d.isdigit(): continue
        if datetime.datetime.utcfromtimestamp(int(d) / 1000).strftime("%Y-%m") != month: continue
        orders[gv(f, "订单号").strip()].append({"rid": r["record_id"], "wn": gv(f, "运单号").strip(),
            "logi": gv(f, "物流公司").strip(), "qty": num(gv(f, "数量")), "cur": gv(f, "物流成本(待接)").strip(),
            "way": gv(f, "合作方式")})
    updates = {}; unresolved = []; no_wn = []; MANUAL = {"信丰", "跨越", "韵达", "其他", "百世"}
    for onum, rs in orders.items():
        wbset = {r["wn"]: r["logi"] for r in rs if r["wn"]}
        needship = any(r["way"] in ("经销买断", "代销月结", "寄售", "赠样") for r in rs)
        selfpick = any(r["logi"] == "自提/无物流" for r in rs)
        if not wbset:
            if needship and not selfpick: no_wn.append(onum)
            continue
        total = 0.0; ok = True; miss = []
        for wn, logi in wbset.items():
            fr = None
            if logi == "顺丰" or wn.upper().startswith("SF"): fr = sf_freight(wn, sfx)
            elif logi == "中通": fr = zto.get(wn)
            elif logi in MANUAL or logi == "":
                man = sum(num(r["cur"]) for r in rs if r["wn"] == wn and r["cur"]); fr = man if man > 0 else None
            if fr is None: ok = False; miss.append(f"{wn}({logi})")
            else: total += fr
        if not ok: unresolved.append((onum, "; ".join(miss))); continue
        tq = sum(r["qty"] for r in rs) or len(rs)
        for r in rs: updates[r["rid"]] = round(total * (r["qty"] / tq), 2) if tq else round(total / len(rs), 2)
    if updates:
        recs = [{"record_id": rid, "fields": {"物流成本(待接)": v}} for rid, v in updates.items()]
        for i in range(0, len(recs), 400):
            requests.post(f"{FEISHU}/bitable/v1/apps/{APP}/tables/{O}/records/batch_update",
                          headers={"Authorization": f"Bearer {T}", "Content-Type": "application/json"},
                          json={"records": recs[i:i + 400]}, timeout=30)
    zto_miss = [o for o, d in unresolved if "中通" in d]
    sf_miss = [o for o, d in unresolved if "顺丰" in d or "SF" in d]
    other_miss = [o for o, d in unresolved if not ("中通" in d or "顺丰" in d or "SF" in d)]
    total_ship = sum(1 for o, rs in orders.items() if any(r["way"] in ("经销买断", "代销月结", "寄售", "赠样") for r in rs) and not any(r["logi"] == "自提/无物流" for r in rs))
    resolved = total_ship - len(unresolved) - len(no_wn)
    complete = (len(no_wn) == 0 and len(unresolved) == 0)
    summ, totals = report(T, month)
    ovw = upsert_overview(T, month, totals)   # 自动灌全渠道总表 c行(本次新增)
    lines = [f"📊 国内分销c渠道 {month} 月度重算完成", summ, f"运费覆盖: {resolved}/{total_ship}单已解析",
             f"已{ovw}总表(国内线下行)"]
    if no_wn: lines.append(f"⚠️ {len(no_wn)}单缺运单号: {', '.join(no_wn[:6])}")
    if zto_miss: lines.append(f"⚠️ {len(zto_miss)}单中通运费缺(账单未传/未出): {', '.join(zto_miss[:6])}")
    if sf_miss: lines.append(f"⚠️ {len(sf_miss)}单顺丰运费缺(单号疑错): {', '.join(sf_miss[:6])}")
    if other_miss: lines.append(f"⚠️ {len(other_miss)}单其他物流缺: {', '.join(other_miss[:6])}")
    lines.append("✅ 物流账单已齐,该月成本完整" if complete else "🟡 物流尚未补齐(上方为当前可算,补齐更准)")
    msg = "\n".join(lines)
    requests.post(f"{FEISHU}/im/v1/messages?receive_id_type=open_id",
                  headers={"Authorization": f"Bearer {T}", "Content-Type": "application/json"},
                  json={"receive_id": FRANKIE, "msg_type": "text", "content": json.dumps({"text": msg}, ensure_ascii=False)}, timeout=20)
    return {"complete": complete, "resolved": resolved, "total": total_ship, "overview": ovw, "msg": msg}

C3_APP_ID = os.environ.get("C3_APP_ID", "")
C3_APP_SECRET = os.environ.get("C3_APP_SECRET", "")
REMINDER_JOB = "国内渠道商务专员"

def c3_tok():
    r = requests.post(f"{FEISHU}/auth/v3/tenant_access_token/internal",
                      json={"app_id": C3_APP_ID, "app_secret": C3_APP_SECRET}, timeout=20)
    return r.json()["tenant_access_token"]

def resolve_target():
    """按职务实时查(聪哥1号)→ email → 聪哥3号 open_id。不硬编码人名。"""
    T = tok()
    deps = requests.get(f"{FEISHU}/contact/v3/departments?parent_department_id=0&fetch_child=true&page_size=50&department_id_type=open_department_id",
                        headers={"Authorization": f"Bearer {T}"}, timeout=20).json()
    dep_ids = ["0"] + [d["open_department_id"] for d in deps.get("data", {}).get("items", [])]
    seen = set(); email = None; mobile = None; name = None
    for did in dep_ids:
        pt = None
        while True:
            u = f"{FEISHU}/contact/v3/users?department_id={did}&page_size=50&user_id_type=open_id&department_id_type=open_department_id" + (f"&page_token={pt}" if pt else "")
            try: d = requests.get(u, headers={"Authorization": f"Bearer {T}"}, timeout=20).json()
            except: break
            for usr in d.get("data", {}).get("items", []):
                if usr.get("open_id") in seen: continue
                seen.add(usr.get("open_id"))
                if REMINDER_JOB in (usr.get("job_title") or ""):
                    email = usr.get("email") or usr.get("enterprise_email")
                    mobile = usr.get("mobile"); name = usr.get("name")
            if d.get("data", {}).get("has_more"): pt = d["data"]["page_token"]
            else: break
        if email or mobile: break
    if not (email or mobile): return None, None
    # 聪哥3号 namespace open_id by mobile(优先,同事多无email)或 email
    T3 = c3_tok()
    body = {}
    if mobile: body["mobiles"] = [mobile]
    if email: body["emails"] = [email]
    r = requests.post(f"{FEISHU}/contact/v3/users/batch_get_id?user_id_type=open_id",
                      headers={"Authorization": f"Bearer {T3}", "Content-Type": "application/json"},
                      json=body, timeout=20).json()
    lst = r.get("data", {}).get("user_list", [])
    oid = next((x.get("user_id") for x in lst if x.get("user_id")), None)
    return oid, name

def build_card(month_hint):
    return {"config": {"wide_screen_mode": True},
            "header": {"title": {"tag": "plain_text", "content": f"📦 国内分销 物流账单提醒 · {month_hint}账单"}, "template": "blue"},
            "elements": [
                {"tag": "div", "text": {"tag": "lark_md", "content": f"威哥，**{month_hint}** 的中通/顺丰账单出了，请上传到「物流账单上传台」。\n两个顺丰账号(957/956)+中通账单都传上后，点下面按钮，我自动算出当月净毛利。"}},
                {"tag": "action", "actions": [
                    {"tag": "button", "text": {"tag": "plain_text", "content": "✅ 账单已传完，开始算"},
                     "type": "primary", "value": {"action": "recompute_offline"}}]}]}

@app.get("/health")
def health(): return {"ok": True}

@app.post("/send-reminder")
async def send_reminder(request: Request):
    if AUTH_TOKEN and request.headers.get("Authorization") != f"Bearer {AUTH_TOKEN}":
        raise HTTPException(401, "unauthorized")
    last = datetime.date.today().replace(day=1) - datetime.timedelta(days=1)
    mh = last.strftime("%Y-%m")
    oid, name = resolve_target()
    if not oid: return {"ok": False, "err": "职务未解析到人"}
    T3 = c3_tok()
    r = requests.post(f"{FEISHU}/im/v1/messages?receive_id_type=open_id",
                      headers={"Authorization": f"Bearer {T3}", "Content-Type": "application/json"},
                      json={"receive_id": oid, "msg_type": "interactive", "content": json.dumps(build_card(mh), ensure_ascii=False)}, timeout=20).json()
    return {"ok": r.get("code") == 0, "target": name, "month": mh, "code": r.get("code")}

@app.post("/recompute")
async def recompute(request: Request, month: str = ""):
    if AUTH_TOKEN and request.headers.get("Authorization") != f"Bearer {AUTH_TOKEN}":
        raise HTTPException(401, "unauthorized")
    if not month:
        last = datetime.date.today().replace(day=1) - datetime.timedelta(days=1)
        month = last.strftime("%Y-%m")
    return do_recompute(month)
