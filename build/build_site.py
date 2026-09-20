"""Build the QuotaKit static site (Cloudflare Pages) and a one-page artifact preview."""
import os, json, shutil
from PIL import Image

HERE = os.path.dirname(os.path.abspath(__file__))
DIST = os.environ.get("QK_DIST", os.path.join(HERE, "dist"))
IMG_SRC = os.environ.get("QK_IMG_SRC", os.path.join(HERE, "..", "exp1-quotakit-templates", "images", "listing"))
SHOP = "https://quotakit.gumroad.com"                 # Gumroad profile (placeholder until the account exists)
GUMROAD = "https://quotakit.gumroad.com/l/salesops"   # skills pack (placeholder slug)
GITHUB = "https://github.com/quotakit/salesops-skills"

FONTS = '<link rel="preconnect" href="https://fonts.googleapis.com"><link rel="preconnect" href="https://fonts.gstatic.com" crossorigin><link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Bricolage+Grotesque:opsz,wght@12..96,500;12..96,700;12..96,800&family=Source+Sans+3:ital,wght@0,400;0,600;1,400&family=JetBrains+Mono:wght@400;600&display=swap">'

CSS = r"""
:root{--ink:#1F2A44;--teal:#0E9F8E;--teal-2:#0B7F72;--paper:#F7F8F5;--surface:#FFFFFF;--text:#28324A;--muted:#5E6A80;--line:#D6DCD9;--input:#FFF6D6;--input-ink:#0B3FD6;--amber:#F2B134;--good:#1B7F3B;--warn:#B45309;--bad:#B42318;--shadow:0 10px 30px rgba(31,42,68,.10);color-scheme:light}
@media (prefers-color-scheme: dark){:root:not([data-theme="light"]){--ink:#E9EEF6;--teal:#2FC4B0;--teal-2:#7BE0D2;--paper:#121826;--surface:#1A2234;--text:#DCE3EE;--muted:#9AA7BC;--line:#2C3750;--input:#3A3417;--input-ink:#9DB8FF;--shadow:0 10px 30px rgba(0,0,0,.35);color-scheme:dark}}
:root[data-theme="dark"]{--ink:#E9EEF6;--teal:#2FC4B0;--teal-2:#7BE0D2;--paper:#121826;--surface:#1A2234;--text:#DCE3EE;--muted:#9AA7BC;--line:#2C3750;--input:#3A3417;--input-ink:#9DB8FF;--shadow:0 10px 30px rgba(0,0,0,.35);color-scheme:dark}
*{box-sizing:border-box}html{scroll-behavior:smooth}
body{margin:0;background:var(--paper);color:var(--text);font:400 17px/1.55 "Source Sans 3",system-ui,-apple-system,Segoe UI,sans-serif;padding-inline:clamp(16px,4vw,48px);padding-block:0 64px}
h1,h2,h3{font-family:"Bricolage Grotesque","Source Sans 3",system-ui,sans-serif;color:var(--ink);text-wrap:balance;line-height:1.08;margin:0}
h1{font-size:clamp(2.1rem,5.2vw,3.6rem);font-weight:800;letter-spacing:-.02em}
h2{font-size:clamp(1.5rem,3vw,2.1rem);font-weight:700;letter-spacing:-.015em}
h3{font-size:1.15rem;font-weight:700}
p{margin:0}
.measure{max-width:62ch}
a{color:var(--teal-2);text-decoration-thickness:1px;text-underline-offset:3px}
a:focus-visible,button:focus-visible,input:focus-visible,select:focus-visible{outline:3px solid var(--amber);outline-offset:2px}
.wrap{max-width:1120px;margin-inline:auto}
.eyebrow{font:600 .78rem/1 "JetBrains Mono",monospace;letter-spacing:.12em;text-transform:uppercase;color:var(--teal-2)}
nav.top{display:flex;align-items:center;justify-content:space-between;gap:16px;padding-block:18px;border-bottom:1px solid var(--line);flex-wrap:wrap}
.brand{font-family:"Bricolage Grotesque",sans-serif;font-weight:800;font-size:1.25rem;color:var(--ink);text-decoration:none;letter-spacing:-.01em}
.brand span{color:var(--teal)}
nav.top ul{list-style:none;display:flex;gap:20px;margin:0;padding:0;flex-wrap:wrap}
nav.top a{color:var(--text);text-decoration:none;font-weight:600;font-size:.95rem}
nav.top a:hover{color:var(--teal-2)}
.hero{display:grid;grid-template-columns:1.1fr .9fr;gap:40px;align-items:center;padding-block:56px 40px}
.hero .lede{font-size:1.2rem;color:var(--muted);margin-top:18px;max-width:52ch}
.cta-row{display:flex;gap:12px;flex-wrap:wrap;margin-top:28px}
.btn{display:inline-flex;align-items:center;gap:8px;padding:12px 18px;border-radius:8px;font-weight:600;text-decoration:none;border:1px solid transparent;cursor:pointer;font-size:1rem;font-family:inherit}
.btn.primary{background:var(--teal);color:#fff}.btn.primary:hover{background:var(--teal-2)}
.btn.ghost{border-color:var(--line);color:var(--ink);background:var(--surface)}.btn.ghost:hover{border-color:var(--teal)}
.curve-card{background:var(--surface);border:1px solid var(--line);border-radius:12px;padding:18px;box-shadow:var(--shadow)}
.curve-card .cap{display:flex;justify-content:space-between;align-items:baseline;gap:12px;margin-bottom:8px}
.curve-card svg{width:100%;height:auto;display:block}
section{padding-block:48px;border-top:1px solid var(--line)}
section > .wrap > header{display:flex;flex-direction:column;gap:10px;margin-bottom:28px}
.tools{display:grid;grid-template-columns:repeat(auto-fit,minmax(240px,1fr));gap:18px}
.tool{background:var(--surface);border:1px solid var(--line);border-top:4px solid var(--teal);border-radius:10px;padding:20px;display:flex;flex-direction:column;gap:10px;text-decoration:none;color:inherit}
.tool:hover{border-color:var(--teal)}
.tool .mono{font:600 .8rem/1 "JetBrains Mono",monospace;color:var(--muted)}
.sheet{background:var(--surface);border:1px solid var(--line);border-radius:12px;box-shadow:var(--shadow);display:grid;grid-template-columns:minmax(280px,420px) 1fr;overflow:hidden}
.sheet .inputs{padding:24px;border-right:1px solid var(--line);display:flex;flex-direction:column;gap:14px}
.sheet .results{padding:24px;display:flex;flex-direction:column;gap:18px;background:color-mix(in srgb,var(--surface) 92%,var(--teal) 8%)}
.field{display:flex;flex-direction:column;gap:6px}
.field label{font-weight:600;font-size:.92rem;color:var(--ink)}
.field small{color:var(--muted);font-size:.82rem}
input[type=number],input[type=text],select{font:600 1rem "JetBrains Mono",monospace;padding:9px 10px;border:1px solid var(--line);border-radius:6px;background:var(--input);color:var(--input-ink);width:100%}
select{font-family:"Source Sans 3",sans-serif}
.tiers{display:grid;grid-template-columns:auto 1fr 1fr;gap:8px 10px;align-items:center}
.tiers .h{font:600 .72rem/1 "JetBrains Mono",monospace;letter-spacing:.08em;text-transform:uppercase;color:var(--muted)}
.kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(140px,1fr));gap:12px}
.kpi{border:1px solid var(--line);border-radius:8px;padding:12px 14px;background:var(--surface)}
.kpi b{display:block;font:600 .72rem/1 "JetBrains Mono",monospace;letter-spacing:.08em;text-transform:uppercase;color:var(--muted);margin-bottom:8px}
.kpi output{font:600 1.6rem/1 "JetBrains Mono",monospace;color:var(--ink);font-variant-numeric:tabular-nums;display:block}
.kpi .sub{font-size:.85rem;color:var(--muted);margin-top:6px}
table.grid{width:100%;border-collapse:collapse;font-variant-numeric:tabular-nums;font-size:.95rem}
table.grid th{text-align:left;font:600 .72rem/1 "JetBrains Mono",monospace;letter-spacing:.08em;text-transform:uppercase;color:var(--muted);padding:8px 6px;border-bottom:1px solid var(--line)}
table.grid td{padding:8px 6px;border-bottom:1px solid var(--line)}
table.grid td.n{text-align:right;font-family:"JetBrains Mono",monospace}
.verdict{border-left:4px solid var(--teal);padding:10px 14px;background:var(--surface);border-radius:0 8px 8px 0}
.verdict.warn{border-color:var(--warn)}.verdict.bad{border-color:var(--bad)}.verdict.good{border-color:var(--good)}
.prose{max-width:66ch;display:flex;flex-direction:column;gap:14px}
.prose h3{margin-top:10px}
.formula{font:500 .92rem/1.5 "JetBrains Mono",monospace;background:var(--surface);border:1px solid var(--line);border-radius:8px;padding:12px 14px;overflow-x:auto}
.catalog{display:grid;grid-template-columns:repeat(auto-fit,minmax(300px,1fr));gap:20px}
.product{background:var(--surface);border:1px solid var(--line);border-radius:12px;overflow:hidden;display:flex;flex-direction:column}
.product img{width:100%;height:auto;aspect-ratio:1/1;object-fit:cover;display:block;border-bottom:1px solid var(--line)}
.product .body{padding:16px 18px 18px;display:flex;flex-direction:column;gap:8px;flex:1}
.product .price{font:600 1rem "JetBrains Mono",monospace;color:var(--ink)}
.product .price s{color:var(--muted);font-weight:400;margin-left:8px}
.product .body a.btn{margin-top:auto;align-self:flex-start}
.skills{display:grid;grid-template-columns:1.2fr 1fr;gap:32px;align-items:start}
.skill-list{display:grid;grid-template-columns:1fr 1fr;gap:10px 18px;margin:0;padding:0;list-style:none}
.skill-list li{padding:10px 12px;border:1px solid var(--line);border-radius:8px;background:var(--surface);font-size:.95rem}
.skill-list li code{font:600 .82rem "JetBrains Mono",monospace;color:var(--teal-2)}
.faq{display:grid;gap:12px;max-width:70ch}
.faq details{border:1px solid var(--line);border-radius:8px;background:var(--surface);padding:12px 16px}
.faq summary{cursor:pointer;font-weight:600;color:var(--ink)}
.faq details p{margin-top:8px;color:var(--text)}
footer{padding-top:32px;border-top:1px solid var(--line);color:var(--muted);font-size:.9rem;display:flex;justify-content:space-between;gap:16px;flex-wrap:wrap}
.disclosure{font-size:.85rem;color:var(--muted)}
@media (max-width:820px){.hero{grid-template-columns:1fr}.sheet{grid-template-columns:1fr}.sheet .inputs{border-right:0;border-bottom:1px solid var(--line)}.skills{grid-template-columns:1fr}.skill-list{grid-template-columns:1fr}}
@media (prefers-reduced-motion:no-preference){.tool,.product{transition:transform .15s ease,border-color .15s ease}.tool:hover{transform:translateY(-2px)}}
"""

# ------------------------------------------------------------------ shared JS
JS_COMMON = r"""
const $=(s,r=document)=>r.querySelector(s);const $$=(s,r=document)=>[...r.querySelectorAll(s)];
const fmt=(n,d=0)=>isFinite(n)?n.toLocaleString(undefined,{maximumFractionDigits:d,minimumFractionDigits:d}):'—';
const pct=(n,d=1)=>isFinite(n)?(n*100).toFixed(d)+'%':'—';
function interp(att,pts,cap,step){pts=pts.filter(p=>isFinite(p[0])&&isFinite(p[1])).sort((a,b)=>a[0]-b[0]);if(!pts.length||att<pts[0][0])return 0;let k=0;for(let i=0;i<pts.length;i++)if(att>=pts[i][0])k=i;let p;if(step||k===pts.length-1)p=pts[k][1];else{const[a0,p0]=pts[k],[a1,p1]=pts[k+1];p=p0+(att-a0)/(a1-a0)*(p1-p0)}return Math.min(cap,p)}
function curveSVG(el,pts,cap,step,marker){const W=520,H=260,L=48,R=16,T=16,B=36;const xs=a=>L+(a/2)*(W-L-R);const ymax=Math.max(cap,...pts.map(p=>p[1]),1)*1.05;const ys=p=>T+(1-p/ymax)*(H-T-B);let d='';for(let a=0;a<=2.0001;a+=0.01){const p=interp(a,pts,cap,step);d+=(d?' L':'M')+xs(a).toFixed(1)+' '+ys(p).toFixed(1)}let grid='';for(let a=0;a<=2;a+=0.25){grid+=`<line x1="${xs(a)}" y1="${T}" x2="${xs(a)}" y2="${H-B}" stroke="var(--line)" stroke-width="1"/><text x="${xs(a)}" y="${H-12}" text-anchor="middle" font-size="11" fill="var(--muted)" font-family="JetBrains Mono,monospace">${Math.round(a*100)}%</text>`}for(let p=0;p<=ymax;p+=0.5){grid+=`<line x1="${L}" y1="${ys(p)}" x2="${W-R}" y2="${ys(p)}" stroke="var(--line)" stroke-width="1"/><text x="${L-6}" y="${ys(p)+4}" text-anchor="end" font-size="11" fill="var(--muted)" font-family="JetBrains Mono,monospace">${Math.round(p*100)}%</text>`}let mk='';if(marker!=null){const a=Math.min(2,Math.max(0,marker)),p=interp(a,pts,cap,step);mk=`<line x1="${xs(a)}" y1="${T}" x2="${xs(a)}" y2="${H-B}" stroke="var(--amber)" stroke-width="2" stroke-dasharray="4 4"/><circle cx="${xs(a)}" cy="${ys(p)}" r="6" fill="var(--amber)" stroke="var(--ink)" stroke-width="2"/>`}el.innerHTML=`<svg viewBox="0 0 ${W} ${H}" role="img" aria-label="Payout curve: attainment on the horizontal axis, payout as a percentage of target incentive on the vertical axis">${grid}<path d="${d}" fill="none" stroke="var(--teal)" stroke-width="3" stroke-linejoin="round"/>${mk}</svg>`}
"""

# ------------------------------------------------------------------ components
def nav(one_page):
    links = [("Tools", "#tools" if one_page else "/#tools"), ("Templates", "#templates" if one_page else "/#templates"),
             ("Claude skills", "#skills" if one_page else "/#skills"), ("FAQ", "#faq" if one_page else "/#faq")]
    return f'<nav class="top wrap" aria-label="Main"><a class="brand" href="{"#top" if one_page else "/"}">Quota<span>Kit</span></a><ul>' + "".join(f'<li><a href="{h}">{t}</a></li>' for t, h in links) + f'<li><a class="btn ghost" href="{SHOP}" rel="noopener">Shop</a></li></ul></nav>'


def hero():
    return f'''<div class="hero wrap" id="top"><div><p class="eyebrow">Sales ops, without the four-hour spreadsheet</p><h1 style="margin-top:12px">Commission, quota and pipeline maths that just works.</h1><p class="lede">Free calculators for sales managers, spreadsheet templates that do the heavy lifting, and Claude skills that write the review for you. Built by a sales-force-effectiveness practitioner, tested cell by cell.</p><div class="cta-row"><a class="btn primary" href="#tools">Use the free tools</a><a class="btn ghost" href="#templates">See the templates</a></div></div><div class="curve-card"><div class="cap"><span class="eyebrow">Payout curve</span><span class="mono" style="font:600 .8rem 'JetBrains Mono',monospace;color:var(--muted)">80% threshold · 100% target · 200% cap</span></div><div id="heroCurve"></div><p class="disclosure" style="margin-top:8px">Attainment on the x-axis, payout as % of target incentive on the y-axis. The same curve the Incentive Plan Simulator template draws.</p></div></div>
<script>curveSVG(document.getElementById('heroCurve'),[[0.8,0.5],[1,1],[1.2,1.5],[1.3,1.75]],2,false,1.12);</script>'''


def tools_index(one_page):
    tools = [("Sales commission calculator", "Tiered rates, marginal or whole-amount, with the breakdown by tier.", "commission", "sales-commission-calculator"),
             ("Incentive payout calculator", "Attainment → payout % of target incentive on your own curve, with the chart.", "incentive", "incentive-payout-calculator"),
             ("Quota attainment calculator", "Where you are, the gap, and the run-rate you need to close it.", "quota", "quota-attainment-calculator"),
             ("Pipeline coverage calculator", "Coverage ratio, weighted forecast and how many new deals you must source.", "pipeline", "pipeline-coverage-calculator")]
    cards = "".join(f'<a class="tool" href="{"#" + a if one_page else "/" + s + "/"}"><span class="mono">FREE · NO SIGN-UP</span><h3>{t}</h3><p>{d}</p></a>' for t, d, a, s in tools)
    return f'<section id="tools"><div class="wrap"><header><p class="eyebrow">Free tools</p><h2>Four calculators sales managers reach for every month</h2><p class="measure" style="color:var(--muted)">Nothing is stored, nothing is tracked, nothing to sign up for. Each one uses the same formulas as the matching spreadsheet template.</p></header><div class="tools">{cards}</div></div></section>'


def tool_commission():
    return r'''<div class="sheet" id="commission"><div class="inputs"><p class="eyebrow">Inputs</p><div class="field"><label for="c_sales">Monthly sales (any currency)</label><input type="number" id="c_sales" value="42500" min="0" step="100"></div><div class="field"><label for="c_method">Tier method</label><select id="c_method"><option value="marginal">Marginal — each slice at its own rate</option><option value="whole">Whole amount — all sales at the top tier reached</option></select></div><div class="tiers"><span class="h">Tier</span><span class="h">From (sales ≥)</span><span class="h">Rate %</span>
<span>1</span><input type="number" id="c_f1" value="0" min="0"><input type="number" id="c_r1" value="3" min="0" step="0.5">
<span>2</span><input type="number" id="c_f2" value="25000" min="0"><input type="number" id="c_r2" value="5" min="0" step="0.5">
<span>3</span><input type="number" id="c_f3" value="50000" min="0"><input type="number" id="c_r3" value="7" min="0" step="0.5">
<span>4</span><input type="number" id="c_f4" value="100000" min="0"><input type="number" id="c_r4" value="10" min="0" step="0.5">
<span>5</span><input type="number" id="c_f5" placeholder="blank" min="0"><input type="number" id="c_r5" placeholder="blank" min="0" step="0.5"></div><div class="field"><label for="c_quota">Monthly quota (optional, for the bonus)</label><input type="number" id="c_quota" value="40000" min="0"></div><div class="field"><label for="c_bonus">Quota bonus at 100% (optional)</label><input type="number" id="c_bonus" value="500" min="0"></div></div>
<div class="results"><p class="eyebrow">Results</p><div class="kpis"><div class="kpi"><b>Commission</b><output id="c_out"></output></div><div class="kpi"><b>Effective rate</b><output id="c_rate"></output></div><div class="kpi"><b>Quota bonus</b><output id="c_bon"></output><span class="sub" id="c_att"></span></div><div class="kpi"><b>Total payout</b><output id="c_tot"></output></div></div><table class="grid"><thead><tr><th>Tier</th><th>From</th><th>To</th><th style="text-align:right">Sales in tier</th><th style="text-align:right">Rate</th><th style="text-align:right">Commission</th></tr></thead><tbody id="c_rows"></tbody></table><p class="disclosure">Marginal works like income-tax brackets. Whole-amount pays the entire month at the highest tier reached, which creates a cliff: a rep at 49,999 earns less than one at 50,000.</p></div></div>
<script>(function(){const ids=['c_sales','c_method','c_quota','c_bonus'];for(let i=1;i<=5;i++){ids.push('c_f'+i,'c_r'+i)}function calc(){const S=+$('#c_sales').value||0,m=$('#c_method').value;const tiers=[];for(let i=1;i<=5;i++){const f=$('#c_f'+i).value,r=$('#c_r'+i).value;if(f!==''&&r!=='')tiers.push([+f,+r/100])}tiers.sort((a,b)=>a[0]-b[0]);let rows='',total=0,top=-1;for(let i=0;i<tiers.length;i++)if(S>=tiers[i][0])top=i;for(let i=0;i<tiers.length;i++){const[f,r]=tiers[i],next=i<tiers.length-1?tiers[i+1][0]:Infinity;const slice=Math.max(0,Math.min(S,next)-f);let c=0;if(m==='marginal')c=slice*r;else if(i===top)c=S*r;total+=c;rows+=`<tr><td>Tier ${i+1}</td><td class="n">${fmt(f)}</td><td class="n">${isFinite(next)?fmt(next):'and above'}</td><td class="n">${m==='marginal'?fmt(slice):(i===top?fmt(S):'—')}</td><td class="n">${pct(r,1)}</td><td class="n"><strong>${fmt(c)}</strong></td></tr>`}$('#c_rows').innerHTML=rows;const q=+$('#c_quota').value||0,b=+$('#c_bonus').value||0;const att=q>0?S/q:NaN;const bon=(q>0&&b>0&&att>=1)?b:0;$('#c_out').value=fmt(total);$('#c_rate').value=S>0?pct(total/S,2):'—';$('#c_bon').value=fmt(bon);$('#c_att').textContent=q>0?'attainment '+pct(att,1):'no quota set';$('#c_tot').value=fmt(total+bon)}ids.forEach(id=>$('#'+id).addEventListener('input',calc));calc()})();</script>'''


def tool_incentive():
    return r'''<div class="sheet" id="incentive"><div class="inputs"><p class="eyebrow">Inputs</p><div class="field"><label for="i_target">Target for the period</label><input type="number" id="i_target" value="500000" min="0"></div><div class="field"><label for="i_actual">Actual</label><input type="number" id="i_actual" value="560000" min="0"></div><div class="field"><label for="i_ti">Target incentive (payout at 100%)</label><input type="number" id="i_ti" value="25000" min="0"><small>What the rep earns at exactly 100% attainment.</small></div><div class="field"><label for="i_type">Curve type</label><select id="i_type"><option value="linear">Linear between points</option><option value="step">Step (slabs)</option></select></div><div class="tiers"><span class="h">Point</span><span class="h">Attainment %</span><span class="h">Payout % of TI</span>
<span>1</span><input type="number" id="i_a1" value="80"><input type="number" id="i_p1" value="50">
<span>2</span><input type="number" id="i_a2" value="100"><input type="number" id="i_p2" value="100">
<span>3</span><input type="number" id="i_a3" value="120"><input type="number" id="i_p3" value="150">
<span>4</span><input type="number" id="i_a4" value="130"><input type="number" id="i_p4" value="175">
<span>5</span><input type="number" id="i_a5" placeholder="blank"><input type="number" id="i_p5" placeholder="blank"></div><div class="field"><label for="i_cap">Cap (max payout % of TI)</label><input type="number" id="i_cap" value="200" min="0"></div></div>
<div class="results"><p class="eyebrow">Results</p><div class="kpis"><div class="kpi"><b>Attainment</b><output id="i_att"></output></div><div class="kpi"><b>Payout % of TI</b><output id="i_pp"></output></div><div class="kpi"><b>Payout</b><output id="i_pay"></output></div></div><div id="i_curve" class="curve-card" style="box-shadow:none"></div><div class="verdict" id="i_note"></div></div></div>
<script>(function(){const ids=['i_target','i_actual','i_ti','i_type','i_cap'];for(let i=1;i<=5;i++)ids.push('i_a'+i,'i_p'+i);function calc(){const T=+$('#i_target').value||0,A=+$('#i_actual').value||0,TI=+$('#i_ti').value||0,cap=(+$('#i_cap').value||0)/100,step=$('#i_type').value==='step';const pts=[];for(let i=1;i<=5;i++){const a=$('#i_a'+i).value,p=$('#i_p'+i).value;if(a!==''&&p!=='')pts.push([+a/100,+p/100])}const att=T>0?A/T:0,pp=interp(att,pts,cap,step),pay=Math.round(pp*TI);$('#i_att').value=pct(att,1);$('#i_pp').value=pct(pp,1);$('#i_pay').value=fmt(pay);curveSVG($('#i_curve'),pts,cap,step,att);const th=pts.length?Math.min(...pts.map(p=>p[0])):0;const n=$('#i_note');if(att<th){n.className='verdict bad';n.textContent=`Below the ${Math.round(th*100)}% threshold: no payout. The rep needs ${fmt(th*T-A)} more to enter the plan.`}else if(pp>=cap){n.className='verdict warn';n.textContent=`At the cap (${Math.round(cap*100)}% of TI). Sales above this point earn nothing extra — check whether that is what you want at ${pct(att,0)} attainment.`}else{n.className='verdict good';n.textContent=`Paid at ${pct(pp,1)} of target incentive. Each extra 1% of attainment here is worth about ${fmt((interp(att+0.01,pts,cap,step)-pp)*TI)}.`}}ids.forEach(id=>$('#'+id).addEventListener('input',calc));calc()})();</script>'''


def tool_quota():
    return r'''<div class="sheet" id="quota"><div class="inputs"><p class="eyebrow">Inputs</p><div class="field"><label for="q_quota">Quota for the period</label><input type="number" id="q_quota" value="1200000" min="0"></div><div class="field"><label for="q_actual">Actual so far</label><input type="number" id="q_actual" value="640000" min="0"></div><div class="field"><label for="q_elapsed">Working days elapsed</label><input type="number" id="q_elapsed" value="34" min="0"></div><div class="field"><label for="q_total">Working days in the period</label><input type="number" id="q_total" value="63" min="1"><small>A quarter is usually 62–65 working days; a month 20–22.</small></div></div>
<div class="results"><p class="eyebrow">Results</p><div class="kpis"><div class="kpi"><b>Attainment</b><output id="q_att"></output><span class="sub" id="q_pace"></span></div><div class="kpi"><b>Gap to quota</b><output id="q_gap"></output></div><div class="kpi"><b>Run-rate so far</b><output id="q_rr"></output><span class="sub">per working day</span></div><div class="kpi"><b>Needed per day</b><output id="q_need"></output><span class="sub" id="q_days"></span></div><div class="kpi"><b>Projected finish</b><output id="q_proj"></output><span class="sub">at current run-rate</span></div></div><div class="verdict" id="q_note"></div><p class="disclosure">"On pace" compares attainment with the share of the period that has elapsed. A rep at 53% with 54% of the days gone is on pace; the same 53% with 70% of days gone is not.</p></div></div>
<script>(function(){const ids=['q_quota','q_actual','q_elapsed','q_total'];function calc(){const Q=+$('#q_quota').value||0,A=+$('#q_actual').value||0,e=+$('#q_elapsed').value||0,t=+$('#q_total').value||1;const att=Q>0?A/Q:0,gap=Math.max(0,Q-A),rr=e>0?A/e:0,left=Math.max(0,t-e),need=left>0?gap/left:gap,proj=rr*t,pace=e/t;$('#q_att').value=pct(att,1);$('#q_pace').textContent=`${pct(pace,0)} of the period elapsed`;$('#q_gap').value=fmt(gap);$('#q_rr').value=fmt(rr);$('#q_need').value=fmt(need);$('#q_days').textContent=`${left} working days left`;$('#q_proj').value=pct(Q>0?proj/Q:0,0);const n=$('#q_note'),ratio=rr>0?need/rr:Infinity;if(gap===0){n.className='verdict good';n.textContent='Quota achieved. Everything from here is accelerator territory.'}else if(ratio<=1){n.className='verdict good';n.textContent=`On pace: the required daily rate (${fmt(need)}) is at or below the current run-rate (${fmt(rr)}).`}else if(ratio<=1.25){n.className='verdict warn';n.textContent=`Slightly behind: the daily rate must rise ${pct(ratio-1,0)} from ${fmt(rr)} to ${fmt(need)} for the remaining ${left} days. Recoverable with a focused close plan.`}else{n.className='verdict bad';n.textContent=`Behind: the daily rate must rise ${pct(ratio-1,0)} to ${fmt(need)}. At the current pace the period ends at ${pct(Q>0?proj/Q:0,0)}. Decide now which deals or accounts close the ${fmt(gap)} gap.`}}ids.forEach(id=>$('#'+id).addEventListener('input',calc));calc()})();</script>'''


def tool_pipeline():
    return r'''<div class="sheet" id="pipeline"><div class="inputs"><p class="eyebrow">Inputs</p><div class="field"><label for="p_target">Target for the period</label><input type="number" id="p_target" value="2500000" min="0"></div><div class="field"><label for="p_won">Closed-won so far</label><input type="number" id="p_won" value="600000" min="0"></div><div class="field"><label for="p_open">Open pipeline closing in the period</label><input type="number" id="p_open" value="3100000" min="0"></div><div class="field"><label for="p_wr">Historical win rate %</label><input type="number" id="p_wr" value="30" min="0" max="100"></div><div class="field"><label for="p_deal">Average deal size</label><input type="number" id="p_deal" value="45000" min="0"></div><div class="field"><label for="p_cov">Healthy coverage ratio</label><input type="number" id="p_cov" value="3" min="1" step="0.5"><small>3× for new business at 25–35% win rates; 1.5–2× for renewals.</small></div></div>
<div class="results"><p class="eyebrow">Results</p><div class="kpis"><div class="kpi"><b>Remaining target</b><output id="p_rem"></output></div><div class="kpi"><b>Coverage</b><output id="p_covout"></output><span class="sub" id="p_covsub"></span></div><div class="kpi"><b>Expected from pipeline</b><output id="p_exp"></output><span class="sub">open × win rate</span></div><div class="kpi"><b>Forecast gap</b><output id="p_gap"></output></div><div class="kpi"><b>New deals to source</b><output id="p_need"></output><span class="sub" id="p_needsub"></span></div></div><div class="verdict" id="p_note"></div><p class="disclosure">Coverage = open pipeline ÷ remaining target. Expected value assumes your historical win rate applies to what is open now; stale deals inflate it, which is why the pipeline review matters.</p></div></div>
<script>(function(){const ids=['p_target','p_won','p_open','p_wr','p_deal','p_cov'];function calc(){const T=+$('#p_target').value||0,W=+$('#p_won').value||0,O=+$('#p_open').value||0,wr=(+$('#p_wr').value||0)/100,D=+$('#p_deal').value||0,H=+$('#p_cov').value||3;const rem=Math.max(0,T-W),cov=rem>0?O/rem:Infinity,exp=O*wr,gap=Math.max(0,rem-exp),need=(D>0&&wr>0)?Math.ceil(gap/(D*wr)):0,pipeNeed=Math.max(0,rem*H-O);$('#p_rem').value=fmt(rem);$('#p_covout').value=isFinite(cov)?cov.toFixed(1)+'×':'done';$('#p_covsub').textContent=isFinite(cov)?`target ${H}× · ${pipeNeed>0?fmt(pipeNeed)+' more pipeline':'covered'}`:'target already met';$('#p_exp').value=fmt(exp);$('#p_gap').value=fmt(gap);$('#p_need').value=fmt(need);$('#p_needsub').textContent=`at ${fmt(D)} avg and ${pct(wr,0)} win rate`;const n=$('#p_note');if(rem===0){n.className='verdict good';n.textContent='Target met. Shift the conversation to next period\'s coverage.'}else if(cov>=H&&gap===0){n.className='verdict good';n.textContent=`Coverage ${cov.toFixed(1)}× and the expected value clears the remaining target. The risk now is pipeline quality, not quantity — review stale and slipped deals.`}else if(cov>=H){n.className='verdict warn';n.textContent=`Coverage looks fine at ${cov.toFixed(1)}× but expected value falls ${fmt(gap)} short — the win rate is the problem, not the volume. Qualify harder or improve conversion before sourcing more.`}else{n.className='verdict bad';n.textContent=`Under-covered: ${cov.toFixed(1)}× against a healthy ${H}×. Sourcing is the first job — about ${need} new opportunities (${fmt(need*D)} of pipeline) at your average deal size.`}}ids.forEach(id=>$('#'+id).addEventListener('input',calc));calc()})();</script>'''


PRODUCTS = [
    ("Sales-Commission-Calculator", "Sales Commission Calculator", "Tiered rates, quota bonus, printable rep statements, dashboard.", "9", "14"),
    ("Incentive-Plan-Simulator", "Sales Incentive Plan Simulator", "Payout curve designer with Base / Upside / Downside cost model.", "14", "19"),
    ("Sales-KPI-Tracker", "Sales KPI Dashboard & Tracker", "Seven inputs a month; funnel, win rate, RAG scorecard.", "9", "12"),
    ("Pipeline-Deal-Tracker", "Sales Pipeline & Deal Tracker", "300 deals, weighted 6-month forecast, overdue flags.", "9", "14"),
    ("Territory-Coverage-Planner", "Territory & Account Coverage Planner", "A/B/C visit plans, coverage %, rep league table.", "9", "14"),
]
SLUGS = {"Sales-Commission-Calculator": "commission", "Incentive-Plan-Simulator": "incentive", "Sales-KPI-Tracker": "kpi", "Pipeline-Deal-Tracker": "pipeline", "Territory-Coverage-Planner": "territory"}


def catalog():
    cards = "".join(f'<article class="product"><img src="img/{k}.jpg" alt="{n} — dashboard preview" loading="lazy" width="800" height="800"><div class="body"><h3>{n}</h3><p style="color:var(--muted)">{d}</p><p class="price">USD {p}<s>{o}</s></p><a class="btn ghost" href="{SHOP}/l/{SLUGS[k]}" rel="noopener">Buy — USD {p}</a></div></article>' for k, n, d, p, o in PRODUCTS)
    bundle = f'<article class="product" style="border-color:var(--teal)"><div class="body" style="justify-content:center;gap:12px"><p class="eyebrow">Bundle</p><h3>Sales Ops Toolkit — all five</h3><p style="color:var(--muted)">Every template above, blank plus sample, in one download. Less than half the price of buying separately.</p><p class="price">USD 29<s>49</s></p><a class="btn primary" href="{SHOP}/l/toolkit" rel="noopener">Get the bundle</a></div></article>'
    return f'<section id="templates"><div class="wrap"><header><p class="eyebrow">Spreadsheet templates</p><h2>The spreadsheets behind the calculators</h2><p class="measure" style="color:var(--muted)">Excel and Google Sheets. Formulas only, no macros. Each file ships blank and with sample data, plus a read-me. Yellow cells with blue text are yours; everything else calculates.</p></header><div class="catalog">{cards}{bundle}</div></div></section>'


def skills():
    items = [("incentive-plan-designer", "payout curve + cost simulation"), ("commission-plan-auditor", "loopholes, cost exposure, fixes"), ("sales-report-analyst", "MBR commentary from your export"), ("territory-planner", "A/B/C tiers, workload, call plan"),
             ("pipeline-reviewer", "stale, slipped, coverage, 1:1 questions"), ("quota-setter", "fair allocation with guardrails"), ("field-force-sizing", "headcount build-up + sensitivity"), ("sales-review-prep", "MBR/QBR slides, 1:1 half-pages")]
    lis = "".join(f'<li><code>{k}</code><br>{d}</li>' for k, d in items)
    return f'<section id="skills"><div class="wrap"><header><p class="eyebrow">Claude skills</p><h2>Give Claude a sales-ops analyst\'s playbooks</h2></header><div class="skills"><div class="prose"><p>Eight skills that tell Claude exactly how a good sales-operations analyst works — what to ask for, which checks to run, what the output looks like — with small Python scripts so the numbers are computed from your file rather than estimated.</p><p>Two are free on GitHub (<code>sales-report-analyst</code>, <code>pipeline-reviewer</code>). The full pack is a one-time purchase with free updates.</p><div class="cta-row"><a class="btn primary" href="{GUMROAD}" rel="noopener">Pro pack — USD 19</a><a class="btn ghost" href="{GITHUB}" rel="noopener">Free starter on GitHub</a></div></div><ul class="skill-list">{lis}</ul></div></div></section>'


FAQ = [
    ("Do the calculators store my numbers?", "No. Everything runs in your browser; nothing is sent anywhere. Refresh the page and it resets."),
    ("Marginal or whole-amount commission — which should I use?", "Marginal (each slice at its own rate) is fairer and has no cliffs, so reps do not sandbag a sale to jump a tier. Whole-amount is simpler to explain but creates a cliff at every threshold. If you must use whole-amount, keep the rate steps small."),
    ("Will the templates work in Google Sheets?", "Yes. Upload the .xlsx to Google Drive and open it with Sheets, then File > Save as Google Sheets. The templates use only functions that Excel, Sheets, Numbers and LibreOffice all support — no macros, no XLOOKUP."),
    ("Who builds these?", "QuotaKit is run by a sales-force-effectiveness practitioner. The templates and skills are drafted with AI assistance, then designed, curated and tested by hand — every formula is checked against an independent calculation before release."),
    ("Can I use a template for my clients?", "Yes — the licence covers use inside your own company and for your own clients. What it does not cover is reselling or redistributing the files themselves."),
]


def faq_section():
    items = "".join(f'<details><summary>{q}</summary><p>{a}</p></details>' for q, a in FAQ)
    ld = json.dumps({"@context": "https://schema.org", "@type": "FAQPage", "mainEntity": [{"@type": "Question", "name": q, "acceptedAnswer": {"@type": "Answer", "text": a}} for q, a in FAQ]})
    return f'<section id="faq"><div class="wrap"><header><p class="eyebrow">FAQ</p><h2>Questions we get</h2></header><div class="faq">{items}</div></div></section><script type="application/ld+json">{ld}</script>'


def footer():
    return f'<footer class="wrap"><span>© 2026 QuotaKit · Sales-ops tools, templates and Claude skills.</span><span class="disclosure">Created with AI assistance; designed, curated and tested by QuotaKit. Not legal, tax or HR advice.</span></footer>'


TOOL_PAGES = {
    "sales-commission-calculator": {
        "title": "Sales Commission Calculator", "desc": "Free tiered sales commission calculator: marginal or whole-amount tiers, quota bonus, effective rate and a breakdown by tier. Same formulas as the QuotaKit Excel template.",
        "h1": "Sales commission calculator", "intro": "Type a month's sales and your tier table. The calculator pays it out two ways — marginal (each slice at its own rate) and whole-amount (everything at the top tier reached) — and shows the breakdown a rep would want to see.",
        "tool": tool_commission,
        "how": [("The marginal formula", "For each tier, commission = (sales in that tier) × (tier rate). Sales in a tier = min(sales, next tier's threshold) − this tier's threshold, floored at zero. Add the tiers up. In a spreadsheet: <span class=\"formula\">=SUMPRODUCT((S&gt;From)*(S−From)*(Rate−PrevRate))</span>"),
                ("The whole-amount formula", "Find the highest tier whose threshold the month's sales reach, and pay the whole amount at that tier's rate: <span class=\"formula\">=INDEX(Rates, MATCH(S, Froms, 1)) × S</span>. Simple to explain, but it creates a cliff at every threshold."),
                ("Quota bonus", "A flat amount paid when sales ÷ quota reaches 100%. Keep it under 15% of a typical month's commission or it starts to dominate behaviour.")],
        "faq": [("Which method do most companies use?", "Marginal, for the same reason tax brackets are marginal: nobody earns less by selling more. Whole-amount survives in small teams because it fits on one line."),
                ("How do I handle a cap?", "Apply it after the tier maths: payout = min(commission, cap). The Excel template has a cap cell; set it to 0 for no cap."),
                ("What if a rep's sales include returns?", "Pay on net invoiced sales and deduct credits issued within 60–90 days from the period in which they are issued. Say so in the plan document.")],
        "template": ("Sales-Commission-Calculator", "Sales Commission Calculator template", "50 reps, 1,000-row sales log, monthly matrices, printable statements, dashboard.")},
    "incentive-payout-calculator": {
        "title": "Incentive Payout Calculator", "desc": "Free sales incentive payout calculator: attainment to payout % of target incentive on a step or linear curve with threshold, accelerators and cap. Draws your payout curve.",
        "h1": "Incentive payout calculator", "intro": "Enter target, actual and the target incentive, then shape the payout curve — threshold, 100% point, accelerators, cap. The chart redraws as you type and the marker shows where this rep lands.",
        "tool": tool_incentive,
        "how": [("Target incentive", "The amount a rep earns at exactly 100% attainment. Every payout is expressed as a percentage of it, which keeps plans comparable across roles with different pay."),
                ("Linear vs step", "Linear interpolates between the points you set — smooth, fair, hard to game. Step (slabs) jumps to the next payout at each point; simple, but a rep one sale short of a slab has a reason to hold the order back."),
                ("Threshold and cap", "Below the first point nothing is paid. Above the cap nothing extra is paid. Set the threshold where 25–35% of reps historically land below it, and the cap at 200–250% of target incentive for monthly plans.")],
        "faq": [("What is a good accelerator?", "Above 100%, make each attainment point worth 1.5–3× what it was worth below 100%. In the default curve, 80→100% is 2.5 points of payout per attainment point; 100→120% is also 2.5 — raise the 120% payout to 160% to make it a real accelerator."),
                ("How do I know what the plan will cost?", "Run the whole team through the curve at Base, +10% and −10% attainment and look at cost ÷ sales in the upside case. The Incentive Plan Simulator template does exactly this for up to 100 reps."),
                ("Should a plan ever be uncapped?", "Only annual plans with a windfall clause (unusually large deals are reviewed) and clawback. Monthly uncapped plans get gamed by shipment timing.")],
        "template": ("Incentive-Plan-Simulator", "Sales Incentive Plan Simulator template", "100 reps, three scenarios, cost ÷ sales, attainment distribution, payout curve chart.")},
    "quota-attainment-calculator": {
        "title": "Quota Attainment Calculator", "desc": "Free quota attainment calculator: attainment %, gap to quota, run-rate so far, required daily rate for the days left, and projected finish — with an on-pace verdict.",
        "h1": "Quota attainment calculator", "intro": "Where a rep or a team stands against quota, and what it takes to close the gap in the working days that are left. Use it mid-month or mid-quarter, not just at the end.",
        "tool": tool_quota,
        "how": [("Attainment and pace", "Attainment = actual ÷ quota. Pace = working days elapsed ÷ working days in the period. On pace means attainment ≥ pace, not attainment ≥ 100%."),
                ("Required daily rate", "(Quota − actual) ÷ working days left. Compare it with the run-rate so far (actual ÷ days elapsed). A required rate more than 25% above the current one rarely happens without a specific close plan."),
                ("Projected finish", "Run-rate so far × total working days, as a % of quota. It assumes the period is linear, which most are not — treat it as a floor, not a forecast, if the period ends with a push.")],
        "faq": [("Should I use calendar days or working days?", "Working days. Sales do not happen on Sundays, and a month with a festival week has fewer selling days than the calendar suggests."),
                ("How do I use this for a team?", "Enter the team's total quota and actual. Then run the reps who are below pace one at a time — the gap is usually concentrated in two or three of them."),
                ("What is the KPI tracker for?", "The same maths every month, plus funnel conversion, win rate and a red/amber/green scorecard, for a rep or a team. Seven inputs a month.")],
        "template": ("Sales-KPI-Tracker", "Sales KPI Dashboard & Tracker template", "Monthly inputs, conversion rates, attainment, run-rate, RAG scorecard, charts.")},
    "pipeline-coverage-calculator": {
        "title": "Pipeline Coverage Calculator", "desc": "Free pipeline coverage calculator: coverage ratio against remaining target, expected value at your win rate, forecast gap and the number of new opportunities you need to source.",
        "h1": "Pipeline coverage calculator", "intro": "Is there enough in the pipeline to make the number? Enter the target, what has closed, what is open, your win rate and average deal size. The calculator says whether the problem is volume or conversion, and how many deals to source.",
        "tool": tool_pipeline,
        "how": [("Coverage ratio", "Open pipeline closing in the period ÷ remaining target. Healthy is about 3× for new business at 25–35% win rates and 1.5–2× for renewals. Coverage is a volume check only."),
                ("Expected value", "Open pipeline × historical win rate. If this is below the remaining target while coverage looks healthy, the pipeline is fat but not good — fix qualification before sourcing more."),
                ("Deals to source", "Forecast gap ÷ (average deal size × win rate), rounded up. It is the number of *new* opportunities, not closed deals — most of them will be lost.")],
        "faq": [("Why does my CRM's weighted pipeline differ?", "CRMs weight each deal by its stage probability; this calculator uses one blended win rate. Both are estimates; the stage-weighted version is better when stage probabilities are calibrated from history."),
                ("What counts as 'open pipeline closing in the period'?", "Deals whose expected close date falls inside the period. Anything overdue (close date in the past) should be re-dated or removed first, or it inflates coverage."),
                ("How do I review the pipeline itself?", "Look for deals with no activity in 30+ days, deals in the same stage for twice the median, and close dates that keep moving. The Pipeline & Deal Tracker template flags all three.")],
        "template": ("Pipeline-Deal-Tracker", "Sales Pipeline & Deal Tracker template", "300 deals, stage probabilities, weighted 6-month forecast, overdue flags, top 10 open deals.")},
}


def head(title, desc, canonical, one_page=False):
    ld = ""
    return (f'<title>{title}</title><meta name="description" content="{desc}"><meta property="og:title" content="{title}"><meta property="og:description" content="{desc}">'
            f'{"" if one_page else f"<link rel=canonical href={canonical}>"}{FONTS}<style>{CSS}</style><script>{JS_COMMON}</script>')


def tool_page(slug, spec):
    how = "".join(f'<h3>{h}</h3><p>{b}</p>' for h, b in spec["how"])
    fq = "".join(f'<details><summary>{q}</summary><p>{a}</p></details>' for q, a in spec["faq"])
    ld = json.dumps({"@context": "https://schema.org", "@type": "FAQPage", "mainEntity": [{"@type": "Question", "name": q, "acceptedAnswer": {"@type": "Answer", "text": a}} for q, a in spec["faq"]]})
    k, tn, td = spec["template"]
    body = (nav(False) + f'<div class="wrap" style="padding-block:40px 24px"><p class="eyebrow">Free tool</p><h1 style="margin-top:12px">{spec["h1"]}</h1><p class="lede measure" style="margin-top:16px;color:var(--muted);font-size:1.15rem">{spec["intro"]}</p></div>'
            + f'<div class="wrap">{spec["tool"]()}</div>'
            + f'<section><div class="wrap"><header><p class="eyebrow">How it works</p><h2>The maths, so you can check it</h2></header><div class="prose">{how}</div></div></section>'
            + f'<section><div class="wrap"><header><p class="eyebrow">Spreadsheet version</p><h2>Do this for the whole team</h2></header><div class="catalog" style="grid-template-columns:minmax(260px,420px)"><article class="product"><img src="../img/{k}.jpg" alt="{tn} preview" width="800" height="800" loading="lazy"><div class="body"><h3>{tn}</h3><p style="color:var(--muted)">{td}</p><a class="btn primary" href="{SHOP}/l/{SLUGS[k]}" rel="noopener">Get the template</a></div></article></div></div></section>'
            + f'<section id="faq"><div class="wrap"><header><p class="eyebrow">FAQ</p><h2>Questions about this calculator</h2></header><div class="faq">{fq}</div></div></section><script type="application/ld+json">{ld}</script>' + footer())
    return f'<!doctype html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">{head(spec["title"] + " — QuotaKit", spec["desc"], f"https://quotakit.github.io/{slug}/")}</head><body>{body}</body></html>'


def home(one_page=False):
    tool_blocks = ""
    if one_page:
        tool_blocks = "".join(f'<section><div class="wrap"><header><p class="eyebrow">Free tool</p><h2>{s["h1"]}</h2><p class="measure" style="color:var(--muted)">{s["intro"]}</p></header>{s["tool"]()}</div></section>' for s in TOOL_PAGES.values())
    body = nav(one_page) + hero() + tools_index(one_page) + tool_blocks + catalog() + skills() + faq_section() + footer()
    if one_page:
        body = body.replace('src="img/', 'src="img/')
        return f'{head("QuotaKit", "Free sales-ops calculators, spreadsheet templates and Claude skills for sales managers.", "", True)}{body}'
    return f'<!doctype html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">{head("QuotaKit — sales ops calculators, templates and Claude skills", "Free commission, incentive, quota and pipeline calculators; Excel & Google Sheets templates; Claude skills for sales managers and RevOps.", "https://quotakit.github.io/")}</head><body>{body}</body></html>'


def main():
    shutil.rmtree(DIST, ignore_errors=True)
    os.makedirs(os.path.join(DIST, "img"))
    for k, *_ in PRODUCTS:
        im = Image.open(os.path.join(IMG_SRC, f"{k}_01_main.png")).convert("RGB")
        im.thumbnail((800, 800)); im.save(os.path.join(DIST, "img", f"{k}.jpg"), quality=82, optimize=True)
    open(os.path.join(DIST, "index.html"), "w").write(home())
    for slug, spec in TOOL_PAGES.items():
        os.makedirs(os.path.join(DIST, slug))
        open(os.path.join(DIST, slug, "index.html"), "w").write(tool_page(slug, spec))
    open(os.path.join(DIST, "robots.txt"), "w").write("User-agent: *\nAllow: /\nSitemap: https://quotakit.github.io/sitemap.xml\n")
    urls = ["https://quotakit.github.io/"] + [f"https://quotakit.github.io/{s}/" for s in TOOL_PAGES]
    open(os.path.join(DIST, "sitemap.xml"), "w").write('<?xml version="1.0" encoding="UTF-8"?><urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">' + "".join(f"<url><loc>{u}</loc></url>" for u in urls) + "</urlset>")
    open(os.path.join(DIST, ".nojekyll"), "w").write("")
    if os.environ.get("QK_NO_ARTIFACT"):
        print("built", os.listdir(DIST)); return
    # artifact preview (one page, no skeleton)
    os.makedirs(os.path.join(HERE, "artifact", "img"), exist_ok=True)
    for f in os.listdir(os.path.join(DIST, "img")):
        shutil.copy(os.path.join(DIST, "img", f), os.path.join(HERE, "artifact", "img", f))
    open(os.path.join(HERE, "artifact", "quotakit-preview.html"), "w").write(home(one_page=True))
    print("built", os.listdir(DIST))


if __name__ == "__main__":
    main()
