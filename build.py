import mammoth, os, re, html
D="2026-10-14"; src=f"../issues/{D}/TVV Slate {D} with Graphs.docx"
imgdir=f"issues/img/{D}"; os.makedirs(imgdir,exist_ok=True); n=[0]
def img(i):
    n[0]+=1; ext=i.content_type.split("/")[-1]; fn=f"g{n[0]:02d}.{ext}"
    with i.open() as f, open(f"{imgdir}/{fn}","wb") as o: o.write(f.read())
    return {"src":f"img/{D}/{fn}","alt":f"Graph {n[0]}","loading":"lazy"}
body=mammoth.convert_to_html(open(src,"rb"),convert_image=mammoth.images.img_element(img)).value
body=re.sub(r'<a href="(https?://[^"]+)"',r'<a href="\1" target="_blank" rel="noopener"',body)
CSS="""body{margin:0;font:16px/1.5 Arial,Helvetica,sans-serif;color:#111;background:#fafafa}
header{position:sticky;top:0;z-index:9;background:#fff;border-bottom:3px solid #e8063c;display:flex;align-items:center;gap:14px;padding:8px 16px;box-shadow:0 2px 6px #0002}
header img{height:56px}header h1{font-size:1.25rem;margin:0}header a{color:inherit;text-decoration:none}
main{max-width:860px;margin:0 auto;padding:16px;background:#fff}
main img{max-width:100%;height:auto;display:block;margin:12px auto}
a{color:#0645ad;text-decoration:underline}h2,h3{border-bottom:1px solid #ddd;padding-top:8px}
nav ul{padding-left:1.2em}@media(max-width:600px){header img{height:40px}header h1{font-size:1rem}}"""
open("assets/style.css","w").write(CSS)
T="TechVestVentures Daily Risk Alert Digest"
def page(title,content,pre):
    return f'''<!doctype html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>{html.escape(title)}</title><link rel="stylesheet" href="{pre}assets/style.css"></head><body><header><a href="{pre}index.html"><img src="{pre}assets/logo.png" alt="TVV logo"></a><h1><a href="{pre}index.html">{T}</a></h1></header><main>{content}</main></body></html>'''
open(f"issues/{D}.html","w").write(page(f"{T} — {D}",f'<p><a href="../index.html">← All issues</a></p>'+body,"../"))
issues=sorted([f[:-5] for f in os.listdir("issues") if f.endswith(".html")],reverse=True)
lst="".join(f'<li><a href="issues/{d}.html">{d}</a></li>' for d in issues)
latest=body.replace('src="img/','src="issues/img/')
open("index.html","w").write(page(T,f'<nav><h2>Issues</h2><ul>{lst}</ul></nav><h2>Latest issue: {issues[0]}</h2>'+latest,""))
open(".nojekyll","w").write(""); print(n[0],"images")
