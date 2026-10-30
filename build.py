"""Usage: python3 build.py [YYYY-MM-DD path/to/slate.docx]
With args: (re)render that issue into issues/DATE.html. Always regenerates index.html tabs from issues/*.html."""
import mammoth, os, re, html, sys, json, datetime
sys.path.insert(0,'/workspace/risk-watcher'); from tvv_nav import nav, NAV_CSS
T="TechVestVentures Daily Risk Alert Digest"
def render(D,src):
    imgdir=f"issues/img/{D}"; os.makedirs(imgdir,exist_ok=True); n=[0]
    # 2026-10-30 standing method: mirror Michael's docx graph widths (wp:extent cx / usable text width, cap 100%)
    # and apply any a:srcRect crop -- see /workspace/risk-watcher/docx_img_widths.py
    from docx_img_widths import pictures, crop_bytes
    import hashlib
    P=pictures(src); used=[False]*len(P)
    def img(i):
        n[0]+=1; ext=i.content_type.split("/")[-1]; fn=f"g{n[0]:02d}.{ext}"
        with i.open() as f: data=f.read()
        h=hashlib.md5(data).hexdigest(); k=n[0]-1
        if not (k<len(P) and not used[k] and P[k]["md5"]==h):   # fall back to first unused picture with same bytes
            k=next((j for j,q in enumerate(P) if not used[j] and q["md5"]==h),None)
        a={"src":f"img/{D}/{fn}","alt":f"Graph {n[0]}","loading":"lazy"}
        if k is not None:
            used[k]=True; q=P[k]
            if q["crop"]: data=crop_bytes(data,q["crop"])
            a["style"]=f"width:{q['pct']:g}%"
        else: print("WARN no width match for",fn)
        open(f"{imgdir}/{fn}","wb").write(data)
        return a
    b=mammoth.convert_to_html(open(src,"rb"),convert_image=mammoth.images.img_element(img)).value
    # reader-facing: drop the internal star legend line
    b=re.sub(r'<p>(?:<[^>]+>)*\s*★\s*=\s*printed in today[^<]*(?:</[^>]+>)*</p>','',b)
    b=re.sub(r'<a href="(https?://[^"]+)"',r'<a href="\1" target="_blank" rel="noopener"',b)
    # desk headings -> h2
    b=re.sub(r'<p><strong>(\d{2} [A-Z][A-Z &/-]*|\d{1,2}\. [^<]{3,80})</strong></p>',r'<h2 class="desk">\1</h2>',b)
    # desk headings typed as Word headings -> same h2
    b=re.sub(r'<h[1-3]>(?:<strong>)?(\d{1,2}\. [^<]{3,80})(?:</strong>)?</h[1-3]>',r'<h2 class="desk">\1</h2>',b)
    # ACTION-ALERT block (through the paragraph before Graph Sources) -> red-bordered box
    b=re.sub(r'(<p>(?:(?!<p>).)*?ACTION-ALERT.*?)(?=<p><strong>Graph Sources|$)',r'<div class="action-box">\1</div>',b,count=1,flags=re.S)
    # ● paragraphs -> list items, grouped into <ul>
    b=re.sub(r'<p>\s*●\s*(.*?)</p>',r'<li>\1</li>',b,flags=re.S)
    b=re.sub(r'((?:<li>.*?</li>)+)',r'<ul class="items">\1</ul>',b,flags=re.S)
    page=f'''<!doctype html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>{T} — {D}</title><link rel="stylesheet" href="../assets/style.css"><script>if(!/[?&]raw/.test(location.search))location.replace("../index.html#{D}")</script></head><body><main id="issue">{b}</main></body></html>'''
    open(f"issues/{D}.html","w").write(page); print(D,n[0],"images")
CSS="""body{margin:0;font:16px/1.5 Arial,Helvetica,sans-serif;color:#111;background:#fafafa}
#top{position:sticky;top:0;z-index:9;background:#fff;box-shadow:0 2px 6px #0002}
header{border-bottom:3px solid #e8063c;display:flex;align-items:center;gap:14px;padding:8px 16px}
header img{height:56px}header h1{font-size:1.25rem;margin:0}
#tabs{display:flex;gap:6px;overflow-x:auto;white-space:nowrap;padding:6px 12px;background:#f3f3f3;border-bottom:1px solid #ddd;-webkit-overflow-scrolling:touch}
#tabs button{flex:0 0 auto;font:600 14px Arial,sans-serif;padding:6px 12px;border:1px solid #bbb;border-radius:16px;background:#fff;color:#222;cursor:pointer}
#tabs button[aria-selected=true]{background:#e8063c;border-color:#e8063c;color:#fff}
main{max-width:1040px;margin:0 auto;padding:16px;background:#fff}
main img{width:100%;max-width:1000px;height:auto;display:block;margin:14px auto}
a{color:#0645ad;text-decoration:underline}h2{border-top:2px solid #e8063c;padding-top:8px;margin-top:22px;font-size:1.2rem;font-weight:700;scroll-margin-top:130px}
main table{width:100%;border-collapse:collapse}main td{display:block;width:100%;padding:0}main td:empty{display:none}
ul.items{margin-left:24px;padding-left:1.1em}
.action-box{border:3px solid #e8063c;padding:10px 14px;margin:20px 0}ul.items li{margin:6px 0}
@media(max-width:600px){header img{height:40px}header h1{font-size:1rem}}"""
if len(sys.argv)==3: render(sys.argv[1],sys.argv[2])
open("assets/style.css","w").write(CSS+NAV_CSS)
issues=sorted([f[:-5] for f in os.listdir("issues") if re.fullmatch(r"\d{4}-\d{2}-\d{2}\.html",f)],reverse=True)
lab={d:datetime.date.fromisoformat(d).strftime("%a %b %-d, %Y") for d in issues}
tabs="".join(f'<button role="tab" data-d="{d}" aria-selected="false">{lab[d]}</button>' for d in issues)
JS="""const tabs=[...document.querySelectorAll('#tabs button')],m=document.getElementById('content'),ds=tabs.map(t=>t.dataset.d);
async function show(d){if(!ds.includes(d))d=ds[0];tabs.forEach(t=>t.setAttribute('aria-selected',t.dataset.d==d));
document.querySelector(`#tabs [data-d="${d}"]`).scrollIntoView({inline:'nearest',block:'nearest'});
m.innerHTML='<p>Loading…</p>';try{const r=await fetch('issues/'+d+'.html',{cache:'no-cache'});const doc=new DOMParser().parseFromString(await r.text(),'text/html');
const c=doc.querySelector('#issue');c.querySelectorAll('img').forEach(i=>i.setAttribute('src','issues/'+i.getAttribute('src')));
c.querySelectorAll('a[href^="http"]').forEach(a=>{a.target='_blank';a.rel='noopener'});m.innerHTML=c.innerHTML;document.title=TT+' — '+d;window.scrollTo(0,0)}
catch(e){m.innerHTML='<p>Could not load issue '+d+'.</p>'}}
const TT=document.title;tabs.forEach(t=>t.onclick=()=>{location.hash=t.dataset.d});
addEventListener('hashchange',()=>show(location.hash.slice(1)));show(location.hash.slice(1));"""
idx=f'''<!doctype html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>{T}</title><link rel="stylesheet" href="assets/style.css"></head><body>{nav("digest")}<div id="top"><header><img src="assets/logo.png" alt="TVV logo"><h1>{T}</h1></header><nav id="tabs" role="tablist" aria-label="Issue dates">{tabs}</nav></div><main id="content"></main><script>{JS}</script></body></html>'''
open("index.html","w").write(idx); open(".nojekyll","w").write(""); print("tabs:",issues)
