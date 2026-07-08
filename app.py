"""
ממשק ווב — מחולל תזרים התחדשות עירונית (תקן 21), רב-מתחמי
גוררים קובץ תחשיב (גיליונות "תחשיב מתחם N") → קובץ תזרים לכל מתחם + אחוז מימון.
"""
import os, sys, tempfile, traceback, json
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from fastapi import FastAPI, UploadFile, File, Form
from fastapi.responses import HTMLResponse, FileResponse, JSONResponse

import write_multi as WM

app = FastAPI(title="מחולל תזרים — התחדשות עירונית")

DEFAULTS = dict(months=42, rate=5.5, equity=25, down=20, promo=1, track20=20,
                permit_share=50, fee_acc=0.5, fee_sale=0.8, fee_rent=3.5,
                fee_own=0.7, fee_nu=0.4, bs=2, bd=9, ss=10, sd=12, fs=20, fd=23)


def _to_engine(form, use_track20):
    """שדות טופס → מפתחות CFG של המנוע (אחוזים→עשרוני)."""
    pct = {"rate", "equity", "down", "track20", "permit_share",
           "fee_acc", "fee_sale", "fee_rent", "fee_own", "fee_nu"}
    v = {}
    for k, dv in DEFAULTS.items():
        raw = form.get(k)
        val = float(raw) if raw not in (None, "") else float(dv)
        if k in pct: val /= 100.0
        elif k in ("months", "promo", "bs", "bd", "ss", "sd", "fs", "fd"): val = int(round(val))
        v[k] = val
    return dict(
        months=v["months"], annual_rate=v["rate"], equity_pct=v["equity"],
        down_payment=v["down"], promo=v["promo"], track20_share=v["track20"],
        planning_permit_share=v["permit_share"], use_track20=use_track20,
        fee_accompaniment=v["fee_acc"], fee_sale_law=v["fee_sale"],
        fee_rent=v["fee_rent"], fee_owners=v["fee_own"], fee_non_util=v["fee_nu"],
        basement=(v["bs"], v["bd"]), skeleton=(v["ss"], v["sd"]), finishing=(v["fs"], v["fd"]),
    )


@app.get("/", response_class=HTMLResponse)
def index():
    return INDEX


@app.post("/generate")
async def generate(
    file: UploadFile = File(...),
    months: str = Form(None), rate: str = Form(None), equity: str = Form(None),
    down: str = Form(None), promo: str = Form(None), track20: str = Form(None),
    permit_share: str = Form(None), fee_acc: str = Form(None), fee_sale: str = Form(None),
    fee_rent: str = Form(None), fee_own: str = Form(None), fee_nu: str = Form(None),
    bs: str = Form(None), bd: str = Form(None), ss: str = Form(None), sd: str = Form(None),
    fs: str = Form(None), fd: str = Form(None), use_track20: str = Form("on"),
):
    form = dict(months=months, rate=rate, equity=equity, down=down, promo=promo,
                track20=track20, permit_share=permit_share, fee_acc=fee_acc,
                fee_sale=fee_sale, fee_rent=fee_rent, fee_own=fee_own, fee_nu=fee_nu,
                bs=bs, bd=bd, ss=ss, sd=sd, fs=fs, fd=fd)
    ov = _to_engine(form, use_track20 in ("on", "true", "1", True))

    tin = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
    tout = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
    try:
        tin.write(await file.read()); tin.close(); tout.close()
        results = WM.build(tin.name, tout.name, overrides=ov)
        if not results:
            return JSONResponse(status_code=400,
                                content={"detail": "לא נמצאו גיליונות בשם 'תחשיב מתחם N' בקובץ"})
        summary = [dict(sheet=r["sheet"], pidyon=round(r["pidyon"]), owners=round(r["owners"]),
                        cost=round(r["cost"]), financing=round(r["financing"]),
                        pct=round(r["pct"], 4), reconciled=r["reconciled"]) for r in results]
        return FileResponse(
            tout.name,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename="cashflow.xlsx",
            headers={"X-Results": json.dumps(summary, ensure_ascii=True)},
        )
    except Exception as e:
        traceback.print_exc()
        return JSONResponse(status_code=400, content={"detail": f"שגיאה בעיבוד: {e}"})
    finally:
        try: os.unlink(tin.name)
        except Exception: pass


def _field(k, label, unit=""):
    return (f'<label>{label}{unit}<input name="{k}" type="number" step="any" '
            f'value="{DEFAULTS[k]}"></label>')


ASSUM = "".join([
    "<div class='grp'><div class='gh'>כללי</div>",
    _field("months", "חודשי בנייה"), _field("rate", "ריבית שנתית", " %"),
    _field("equity", "הון עצמי", " %"), _field("promo", "תקופת מבצע"),
    "</div><div class='grp'><div class='gh'>הכנסות</div>",
    "<label class='chk'><input type='checkbox' name='use_track20' checked> מסלול 20%/80%</label>",
    _field("down", "מקדמה", " %"), _field("track20", "מסלול 20%", " %"),
    _field("permit_share", "תכנון בהיתר", " %"),
    "</div><div class='grp'><div class='gh'>ערבויות ועמלות</div>",
    _field("fee_acc", "ליווי", " %"), _field("fee_sale", "חוק מכר", " %"),
    _field("fee_rent", "שכ\"ד", " %"), _field("fee_own", "בעלים", " %"),
    _field("fee_nu", "אי-ניצול", " %"),
    "</div><div class='grp'><div class='gh'>חלונות ביצוע (התחלה / משך)</div>",
    _field("bs", "מרתף — התחלה"), _field("bd", "מרתף — משך"),
    _field("ss", "שלד — התחלה"), _field("sd", "שלד — משך"),
    _field("fs", "גמר — התחלה"), _field("fd", "גמר — משך"),
    "</div>",
])

INDEX = """<!doctype html><html lang="he" dir="rtl"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>מחולל תזרים — תקן 21</title>
<style>
 :root{--purple:#4F2D7F;--orange:#F57C00}
 *{box-sizing:border-box;font-family:'Heebo','Rubik',Arial,sans-serif}
 body{margin:0;background:#f5f3f8;color:#222;padding:24px}
 .card{background:#fff;border-radius:18px;box-shadow:0 10px 40px rgba(79,45,127,.15);padding:32px;max-width:720px;margin:0 auto}
 h1{color:var(--purple);font-size:22px;margin:0 0 4px}
 p.sub{color:#777;margin:0 0 22px;font-size:14px}
 #drop{border:2.5px dashed #c9b8e0;border-radius:14px;padding:32px 20px;text-align:center;cursor:pointer;background:#faf8fd}
 #drop.hover{border-color:var(--orange);background:#fff5ec}
 #drop .big{font-size:16px;color:var(--purple);font-weight:700}
 #drop .small{font-size:13px;color:#999;margin-top:6px}
 details{margin-top:16px;border:1px solid #eee;border-radius:12px;padding:6px 14px;background:#fbfafc}
 summary{cursor:pointer;font-weight:700;color:var(--purple);padding:8px 0}
 .grp{margin:10px 0}.gh{font-size:12px;color:var(--orange);font-weight:700;margin:8px 0 6px}
 label{display:inline-flex;flex-direction:column;font-size:12px;color:#555;margin:0 8px 8px 0;width:118px}
 label.chk{flex-direction:row;align-items:center;width:auto;gap:6px;font-weight:700;color:var(--purple)}
 input[type=number]{margin-top:3px;padding:7px;border:1px solid #ddd;border-radius:7px;font-size:14px;width:100%}
 .btn{margin-top:18px;width:100%;border:0;border-radius:10px;padding:14px;background:var(--purple);color:#fff;font-size:15px;font-weight:700;cursor:pointer}
 .btn:disabled{opacity:.5}
 .spin{display:none;margin-top:14px;color:var(--purple)}
 .err{display:none;color:#c0392b;margin-top:14px}
 .result{display:none;margin-top:22px}
 table{width:100%;border-collapse:collapse;font-size:13px;margin-bottom:14px}
 th{background:var(--purple);color:#fff;padding:8px;font-weight:700}
 td{padding:8px;border-bottom:1px solid #eee;text-align:center}
 td.pct{font-weight:800;color:var(--orange)}
 .ok{color:#2e7d32}.bad{color:#c62828;font-weight:700}
 .dl{display:inline-block;background:var(--orange);color:#fff;text-decoration:none;padding:12px 24px;border-radius:9px;font-weight:700}
</style></head><body>
<div class="card">
 <h1>מחולל תזרים הכנסות והוצאות</h1>
 <p class="sub">התחדשות עירונית · תקן 21 · לשכת שמאות לנדאו</p>
 <div id="drop">
   <div class="big">גררו לכאן קובץ תחשיב (xlsx.)</div>
   <div class="small">הקובץ חייב לכלול גיליונות בשם "תחשיב מתחם 1", "תחשיב מתחם 2"...</div>
   <input id="file" type="file" accept=".xlsx" hidden>
 </div>
 <details><summary>הנחות עבודה (לחיצה לעריכה)</summary>__ASSUM__</details>
 <button id="go" class="btn" disabled>הפקת תזרים</button>
 <div class="spin" id="spin">מעבד…</div>
 <div class="err" id="err"></div>
 <div class="result" id="result">
   <table id="tbl"><thead><tr><th>מתחם</th><th>פדיון</th><th>עלות לפני מימון</th><th>מימון</th><th>אחוז מימון</th><th>בקרה</th></tr></thead><tbody id="tb"></tbody></table>
   <div style="text-align:center"><a class="dl" id="dl">הורדת קובץ התזרים</a></div>
 </div>
</div>
<script>
const $=id=>document.getElementById(id);
const drop=$('drop'),inp=$('file'),go=$('go'),res=$('result'),tb=$('tb'),dl=$('dl'),err=$('err'),spin=$('spin');
let chosen=null;
const fmt=n=>Number(n).toLocaleString('he-IL');
drop.onclick=()=>inp.click();
['dragover','dragenter'].forEach(e=>drop.addEventListener(e,ev=>{ev.preventDefault();drop.classList.add('hover')}));
['dragleave','drop'].forEach(e=>drop.addEventListener(e,ev=>{ev.preventDefault();drop.classList.remove('hover')}));
drop.addEventListener('drop',ev=>{chosen=ev.dataTransfer.files[0];show()});
inp.onchange=()=>{chosen=inp.files[0];show()};
function show(){if(chosen){drop.querySelector('.big').textContent=chosen.name;go.disabled=false;res.style.display='none';err.style.display='none'}}
go.onclick=async()=>{
 if(!chosen)return;
 go.disabled=true;spin.style.display='block';err.style.display='none';res.style.display='none';
 const fd=new FormData();fd.append('file',chosen);
 document.querySelectorAll('details input[name]').forEach(i=>{
   if(i.type==='checkbox') fd.append(i.name, i.checked?'on':'off');
   else fd.append(i.name,i.value);
 });
 try{
   const r=await fetch('/generate',{method:'POST',body:fd});
   if(!r.ok){const j=await r.json();throw new Error(j.detail||'שגיאה');}
   const results=JSON.parse(r.headers.get('X-Results'));
   const blob=await r.blob();
   tb.innerHTML=results.map(x=>`<tr><td>${x.sheet}</td><td>${fmt(x.pidyon)}</td><td>${fmt(x.cost)}</td><td>${fmt(x.financing)}</td><td class="pct">${(x.pct*100).toFixed(2)}%</td><td class="${x.reconciled?'ok':'bad'}">${x.reconciled?'תקין ✓':'פער — בדוק'}</td></tr>`).join('');
   dl.href=URL.createObjectURL(blob);dl.download='תחשיב_עם_תזרים.xlsx';
   res.style.display='block';
 }catch(e){err.textContent=e.message;err.style.display='block';}
 finally{spin.style.display='none';go.disabled=false;}
};
</script></body></html>""".replace("__ASSUM__", ASSUM)
