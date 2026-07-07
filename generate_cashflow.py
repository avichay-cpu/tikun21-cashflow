"""
ממשק ווב — מחולל תזרים התחדשות עירונית (תקן 21)
גוררים קובץ תחשיב + קובעים הנחות עבודה → אחוז מימון + קובץ תזרים (נוסחאות).

הרצה מקומית:
    pip install -r requirements.txt
    uvicorn app:app --reload
    http://127.0.0.1:8000
"""
import os, sys, tempfile, traceback
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from fastapi import FastAPI, UploadFile, File, Form
from fastapi.responses import HTMLResponse, FileResponse, JSONResponse

import generate_cashflow as G
import write_cashflow_formulas as W

app = FastAPI(title="מחולל תזרים — התחדשות עירונית")

DEFAULTS = dict(months=42, rate=5.5, equity=25, down=20, promo=1, track20=20,
                permit_share=50, fee_acc=0.5, fee_sale=0.8, fee_rent=3.5,
                fee_own=0.7, fee_nu=0.4, bs=2, bd=9, ss=10, sd=12, fs=20, fd=23)


def _canonical(form):
    pct = {"rate", "equity", "down", "track20", "permit_share",
           "fee_acc", "fee_sale", "fee_rent", "fee_own", "fee_nu"}
    ov = {}
    for k, dv in DEFAULTS.items():
        v = form.get(k)
        v = float(v) if v not in (None, "") else float(dv)
        if k in pct:
            v = v / 100.0
        elif k in ("months", "promo", "bs", "bd", "ss", "sd", "fs", "fd"):
            v = int(round(v))
        ov[k] = v
    return ov


def _to_engine(ov):
    return dict(
        months=ov["months"], annual_rate=ov["rate"], equity_pct=ov["equity"],
        down_payment=ov["down"], promo=ov["promo"], track20_share=ov["track20"],
        planning_permit_share=ov["permit_share"],
        fee_accompaniment=ov["fee_acc"], fee_sale_law=ov["fee_sale"],
        fee_rent=ov["fee_rent"], fee_owners=ov["fee_own"], fee_non_util=ov["fee_nu"],
        basement=(ov["bs"], ov["bd"]), skeleton=(ov["ss"], ov["sd"]),
        finishing=(ov["fs"], ov["fd"]),
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
    fs: str = Form(None), fd: str = Form(None),
):
    form = dict(months=months, rate=rate, equity=equity, down=down, promo=promo,
                track20=track20, permit_share=permit_share, fee_acc=fee_acc,
                fee_sale=fee_sale, fee_rent=fee_rent, fee_own=fee_own, fee_nu=fee_nu,
                bs=bs, bd=bd, ss=ss, sd=sd, fs=fs, fd=fd)
    ov = _canonical(form)

    tmp_in = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
    tmp_out = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
    try:
        tmp_in.write(await file.read()); tmp_in.close(); tmp_out.close()
        R = G.generate(tmp_in.name, overrides=_to_engine(ov))
        W.build(tmp_in.name, tmp_out.name, overrides=ov)
        return FileResponse(
            tmp_out.name,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename="cashflow.xlsx",
            headers={"X-Financing-Pct": f"{R['financing_pct']:.4f}"},
        )
    except KeyError:
        return JSONResponse(status_code=400,
                            content={"detail": "הקובץ אינו כולל גיליון 'מאוחד' תקין"})
    except Exception as e:
        traceback.print_exc()
        return JSONResponse(status_code=400, content={"detail": f"שגיאה בעיבוד: {e}"})
    finally:
        try: os.unlink(tmp_in.name)
        except Exception: pass


def _field(k, label, unit=""):
    return (f'<label>{label}{unit}<input name="{k}" type="number" step="any" '
            f'value="{DEFAULTS[k]}"></label>')


ASSUM_HTML = "".join([
    "<div class='grp'><div class='gh'>כללי</div>",
    _field("months", "חודשי בנייה"), _field("rate", "ריבית שנתית", " %"),
    _field("equity", "הון עצמי", " %"), _field("promo", "תקופת מבצע"),
    "</div><div class='grp'><div class='gh'>הכנסות</div>",
    _field("down", "מקדמה", " %"), _field("track20", "מסלול 20%", " %"),
    _field("permit_share", "תכנון בהיתר", " %"),
    "</div><div class='grp'><div class='gh'>ערבויות ועמלות</div>",
    _field("fee_acc", "ליווי", " %"), _field("fee_sale", "חוק מכר", " %"),
    _field("fee_rent", "שכ\"ד", " %"), _field("fee_own", "בעלים", " %"),
    _field("fee_nu", "אי-ניצול", " %"),
    "</div><div class='grp'><div class='gh'>חלונות ביצוע (חודש התחלה / משך)</div>",
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
 body{margin:0;background:#f5f3f8;color:#222;padding:28px}
 .card{background:#fff;border-radius:18px;box-shadow:0 10px 40px rgba(79,45,127,.15);
       padding:34px;max-width:640px;margin:0 auto}
 h1{color:var(--purple);font-size:22px;margin:0 0 4px}
 p.sub{color:#777;margin:0 0 22px;font-size:14px}
 #drop{border:2.5px dashed #c9b8e0;border-radius:14px;padding:34px 20px;text-align:center;
       cursor:pointer;transition:.2s;background:#faf8fd}
 #drop.hover{border-color:var(--orange);background:#fff5ec}
 #drop .big{font-size:16px;color:var(--purple);font-weight:700}
 #drop .small{font-size:13px;color:#999;margin-top:6px}
 details{margin-top:16px;border:1px solid #eee;border-radius:12px;padding:6px 14px;background:#fbfafc}
 summary{cursor:pointer;font-weight:700;color:var(--purple);padding:8px 0}
 .grp{margin:10px 0}
 .gh{font-size:12px;color:var(--orange);font-weight:700;margin:8px 0 6px}
 label{display:inline-flex;flex-direction:column;font-size:12px;color:#555;margin:0 8px 8px 0;width:120px}
 input[type=number]{margin-top:3px;padding:7px;border:1px solid #ddd;border-radius:7px;font-size:14px;width:100%}
 .btn{margin-top:18px;width:100%;border:0;border-radius:10px;padding:14px;background:var(--purple);
      color:#fff;font-size:15px;font-weight:700;cursor:pointer}
 .btn:disabled{opacity:.5;cursor:default}
 .result{margin-top:22px;display:none;border-radius:12px;padding:22px;background:#fff5ec;
         border:1px solid #ffd9b0;text-align:center}
 .pct{font-size:46px;font-weight:800;color:var(--orange);line-height:1}
 .pct-lbl{font-size:13px;color:#a15a00;margin-top:4px}
 .dl{margin-top:16px;display:inline-block;background:var(--orange);color:#fff;text-decoration:none;
     padding:11px 22px;border-radius:9px;font-weight:700}
 .err{color:#c0392b;margin-top:14px;font-size:14px;display:none}
 .spin{display:none;margin-top:14px;color:var(--purple);font-size:14px}
</style></head><body>
<div class="card">
 <h1>מחולל תזרים הכנסות והוצאות</h1>
 <p class="sub">התחדשות עירונית · תקן 21 · לשכת שמאות לנדאו</p>
 <div id="drop">
   <div class="big">גררו לכאן קובץ תחשיב (xlsx.)</div>
   <div class="small">או לחצו לבחירה — הקובץ חייב לכלול גיליון "מאוחד"</div>
   <input id="file" type="file" accept=".xlsx" hidden>
 </div>
 <details>
   <summary>הנחות עבודה (לחיצה לעריכה)</summary>
   __ASSUM__
 </details>
 <button id="go" class="btn" disabled>הפקת תזרים</button>
 <div class="spin" id="spin">מעבד…</div>
 <div class="err" id="err"></div>
 <div class="result" id="result">
   <div class="pct" id="pct">—</div><div class="pct-lbl">אחוז מימון</div>
   <a class="dl" id="dl">הורדת קובץ התזרים</a>
 </div>
</div>
<script>
const drop=document.getElementById('drop'),inp=document.getElementById('file'),
      go=document.getElementById('go'),res=document.getElementById('result'),
      pct=document.getElementById('pct'),dl=document.getElementById('dl'),
      err=document.getElementById('err'),spin=document.getElementById('spin');
let chosen=null;
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
 document.querySelectorAll('details input[name]').forEach(i=>fd.append(i.name,i.value));
 try{
   const r=await fetch('/generate',{method:'POST',body:fd});
   if(!r.ok){const j=await r.json();throw new Error(j.detail||'שגיאה');}
   const p=r.headers.get('X-Financing-Pct');
   const blob=await r.blob();
   pct.textContent=(parseFloat(p)*100).toFixed(2)+'%';
   const url=URL.createObjectURL(blob);
   dl.href=url;dl.download='תחשיב_עם_תזרים.xlsx';
   res.style.display='block';
 }catch(e){err.textContent=e.message;err.style.display='block';}
 finally{spin.style.display='none';go.disabled=false;}
};
</script></body></html>""".replace("__ASSUM__", ASSUM_HTML)
