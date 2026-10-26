<!doctype html>
<html lang="pt-BR">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Controle de Histogramas — COSAMPA</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<style>
:root{
  --or:#F7931E; --or2:#E86F00; --or-soft:#FFF4E6;
  --bk:#0E0E10; --bk2:#1C1C20; --bg:#F3F4F6; --card:#FFFFFF;
  --tx:#1F2937; --mut:#6B7280; --line:#E5E7EB; --line2:#F1F2F4;
  --ok:#16A34A; --err:#DC2626; --warn:#D97706; --info:#0369A1;
  --r:12px; --sh:0 1px 2px rgba(16,24,40,.06),0 4px 14px rgba(16,24,40,.06);
}
*{box-sizing:border-box}
[hidden]{display:none !important}
html,body{margin:0}
body{font-family:Arial,Helvetica,sans-serif;background:var(--bg);color:var(--tx);font-size:14px;-webkit-font-smoothing:antialiased}
button,input,select{font-family:inherit;font-size:13px}
/* ---------- layout: menu lateral ---------- */
.app{display:grid;grid-template-columns:236px minmax(0,1fr);min-height:100vh}
.side{background:var(--bk);color:#fff;position:sticky;top:0;height:100vh;display:flex;flex-direction:column;overflow-y:auto;border-right:3px solid var(--or);z-index:45}
.side-brand{padding:18px 18px 16px;border-bottom:1px solid #26262B}
.logo{min-height:34px;display:flex;align-items:center;cursor:pointer}
.logo img{height:auto;width:100%;max-width:180px;display:block}
.logo-text{font-weight:900;letter-spacing:2px;font-size:20px;color:#fff}
.logo-text b{color:var(--or)}
.side-title{margin-top:10px;font-weight:700;font-size:14px}
.side-title small{display:block;color:#A1A1AA;font-weight:400;font-size:11px;margin-top:2px}
.side-nav{display:flex;flex-direction:column;gap:4px;padding:14px 10px}
.side-nav .cap{font-size:10px;color:#71717A;font-weight:900;letter-spacing:1px;text-transform:uppercase;padding:4px 10px 6px}
.side-nav button{display:flex;align-items:center;gap:12px;width:100%;border:0;background:transparent;color:#D4D4D8;padding:11px 12px;border-radius:10px;font-weight:700;font-size:14px;cursor:pointer;text-align:left;position:relative;transition:background .15s,color .15s}
.side-nav button svg{width:20px;height:20px;flex:none;opacity:.85}
.side-nav button:hover{background:#1C1C20;color:#fff}
.side-nav button.on{background:linear-gradient(90deg,rgba(247,147,30,.22),rgba(247,147,30,.04));color:#fff}
.side-nav button.on svg{color:var(--or);opacity:1}
.side-nav button.on:before{content:"";position:absolute;left:-10px;top:8px;bottom:8px;width:4px;border-radius:0 4px 4px 0;background:var(--or)}
.side-foot{margin-top:auto;padding:14px;border-top:1px solid #26262B;display:flex;flex-direction:column;gap:8px}
.side-btn{display:flex;align-items:center;gap:8px;background:#1C1C20;border:1px solid #2E2E34;color:#E4E4E7;border-radius:9px;padding:9px 11px;font-weight:700;cursor:pointer;font-size:13px}
.side-btn:hover{border-color:var(--or);color:var(--or)}
.side-note{font-size:11px;color:#A1A1AA;line-height:1.4}
.side-nav .badge{margin-left:auto;background:var(--or);color:#fff;font-style:normal;font-size:10px;font-weight:900;border-radius:999px;padding:1px 7px}
.pn-grid{display:grid;grid-template-columns:minmax(0,1.35fr) minmax(0,1fr);gap:14px}
.pn-grid > .card{margin-bottom:14px;min-width:0}
.vg-wk{display:flex;flex-direction:column;font-size:10px;font-weight:900;letter-spacing:.6px;text-transform:uppercase;color:var(--mut);gap:3px}
.vg-wk select{border:1px solid var(--line);border-radius:8px;padding:7px 10px;font-size:13px;text-transform:none;letter-spacing:0;font-weight:400;color:var(--tx);background:#fff}
.vg-sum{display:flex;gap:10px;align-items:center;flex-wrap:wrap;margin-bottom:12px}
.vg-pill{display:flex;align-items:center;gap:7px;background:#fff;border:1px solid var(--line);border-left:5px solid var(--c);border-radius:999px;padding:7px 14px;font-size:13px;box-shadow:var(--sh)}
.vg-pill b{font-size:18px}.vg-rule{font-size:11.5px;margin-left:auto}
.vg-dot{display:inline-block;vertical-align:-1px;margin:0 4px 0 2px;width:12px;height:12px;border-radius:50%;background:var(--c);box-shadow:0 0 0 4px color-mix(in srgb,var(--c) 18%,transparent);flex:none}
.vg-card{padding:0;overflow:hidden}
.vg-tbl{display:flex;flex-direction:column}
.vg-row{display:grid;grid-template-columns:minmax(200px,1.5fr) minmax(170px,1.2fr) minmax(120px,.9fr) minmax(120px,.9fr) minmax(140px,1fr) minmax(170px,1.2fr);gap:14px;align-items:center;padding:14px 18px;border:0;border-bottom:1px solid var(--line2);background:#fff;text-align:left;font:inherit;color:inherit;cursor:pointer;border-left:5px solid var(--c,transparent);width:100%}
.vg-row:hover{background:#FFFBF5}
.vg-h{cursor:default;background:#0E0E10 !important;color:#fff;font-size:11px;font-weight:900;letter-spacing:.6px;text-transform:uppercase;padding:10px 18px;border-left-color:#0E0E10}
.vg-ob{display:flex;gap:12px;align-items:center}.vg-ob b{font-size:15px;display:block}.vg-ob small{display:block;color:var(--mut);font-size:12px;margin-top:2px}
.vg-av{display:flex;flex-direction:column;gap:4px}.vg-avn{display:flex;align-items:baseline;gap:8px}.vg-avn b{font-size:20px}.vg-avn span{font-size:12px;color:var(--mut)}
.vg-av small{font-size:11.5px;font-weight:700;color:#16A34A}.vg-av small.neg{color:#DC2626}
.kpi .vg-av .vg-avn b{color:var(--or);font-size:26px}.kpi.dark .vg-av .vg-avn span{color:#A1A1AA}
.vg-ef b{font-size:20px}.vg-ef small{color:var(--mut);font-size:11px}.vg-sk{height:26px;width:100%;max-width:150px;margin-top:2px}.vg-sk svg{width:100%;height:100%;display:block}
.vg-d{font-size:11.5px;font-weight:900;margin-left:4px;color:var(--mut)}.vg-d.up{color:#16A34A}.vg-d.dn{color:#DC2626}
.vg-pd{display:flex;flex-direction:column;gap:3px;font-size:12.5px}.vg-pd .bad{color:#DC2626}.vg-pd .warn{color:#B45309}
.vg-nx b{display:block;font-size:14px}.vg-nx small{display:block;font-size:12px;line-height:1.3;overflow:hidden;display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical}
.vg-k4{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:12px;margin-bottom:16px}
.vg-k-ef{display:flex;flex-direction:column;gap:2px;margin-top:4px;font-size:13px}.vg-k-ef b{font-size:22px}
.vg-h3{font-size:16px;margin:4px 0 10px}
.vg-att3{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:14px;margin-bottom:2px}
.vg-att{border-top:4px solid var(--line)}.vg-att.err{border-top-color:#DC2626}.vg-att.warn{border-top-color:#D97706}
.vg-ul{list-style:none;margin:0;padding:0}.vg-ul li{display:flex;justify-content:space-between;gap:10px;padding:8px 0;border-bottom:1px solid var(--line2);font-size:13px}
.vg-ul li:last-child{border-bottom:0}.vg-ul li small{display:block;color:var(--mut);font-size:11.5px}.vg-ul .r{text-align:right;white-space:nowrap;font-size:12px}.vg-ul .r small{margin-top:2px}
@media (max-width:1100px){.vg-row{grid-template-columns:repeat(2,minmax(0,1fr))}.vg-h{display:none}.vg-row > div[data-l]::before{content:attr(data-l);display:block;font-size:10px;font-weight:900;letter-spacing:.6px;text-transform:uppercase;color:var(--mut);margin-bottom:3px}.vg-ob{grid-column:1/-1}.vg-att3{grid-template-columns:minmax(0,1fr)}.vg-k4{grid-template-columns:repeat(2,minmax(0,1fr))}.vg-rule{margin-left:0}}
@media (max-width:560px){.vg-row{grid-template-columns:minmax(0,1fr)}}
.h3-grid{display:grid;grid-template-columns:minmax(0,1fr) minmax(0,1fr);gap:18px}
@media (max-width:900px){.h3-grid{grid-template-columns:minmax(0,1fr)}}
.h3-box{min-width:0}.h3-top{display:flex;align-items:center;gap:10px;flex-wrap:wrap;margin-bottom:8px}.h3-top .h3-tt{margin:0}.h3-read{min-height:38px;padding:5px 10px}.h3-read .rv{font-size:20px}.h3-tt{font-weight:900;font-size:14px;margin:0 0 6px;padding-left:8px;border-left:4px solid var(--or)}
.h3t{width:100%;border-collapse:collapse;margin-top:10px;font-size:12.5px}
.h3t th{background:#F3F4F6 !important;color:#374151 !important;font-weight:800;padding:6px 8px;text-align:right;position:static;border-bottom:2px solid #E5E7EB;font-size:12px}
.h3t td{padding:5px 8px;border-bottom:1px solid #F0F1F4;text-align:right}
.h3t .l{text-align:left}.h3t tr.tot td{font-weight:900;border-top:2px solid #111;background:#FFF7ED}
.pbar{position:relative;height:9px;background:#EEF0F3;border-radius:99px;overflow:visible;margin:2px 0 3px}
.pbar i{position:absolute;left:0;top:0;bottom:0;border-radius:99px}
.pbar b{position:absolute;top:-3px;width:2px;height:15px;background:#111827;border-radius:2px;transform:translateX(-1px)}
.la-board{display:grid;grid-auto-flow:column;grid-auto-columns:minmax(210px,1fr);gap:10px;overflow-x:auto;padding-bottom:6px}
.la-col{background:#F7F8FA;border:1px solid var(--line);border-radius:12px;padding:8px;min-height:120px}
.la-col.venc{background:#FEF2F2;border-color:#FECACA}
.la-h{display:flex;justify-content:space-between;align-items:center;font-weight:900;font-size:12px;margin-bottom:8px;gap:6px}
.la-h small{display:block;font-weight:400;color:var(--mut);font-size:10.5px}
.la-card{background:#fff;border:1px solid var(--line);border-left:4px solid var(--or);border-radius:9px;padding:8px;margin-bottom:8px;font-size:12.5px;box-shadow:0 1px 2px rgba(0,0,0,.04)}
.la-col.venc .la-card{border-left-color:var(--err)}
.la-card small{display:block;color:var(--mut);margin-top:2px}
.la-meta{display:flex;gap:6px;flex-wrap:wrap;margin-top:6px;font-size:11px;color:#4B5563;align-items:center}
.la-acts{display:flex;gap:6px;margin-top:6px}
select.sel-ok{background:#DCFCE7;border-color:#86EFAC}select.sel-err{background:#FEE2E2;border-color:#FCA5A5}
td select{border:1px solid var(--line);border-radius:7px;padding:4px 6px;font-size:12px;max-width:230px}
@media (max-width:1100px){.pn-grid{grid-template-columns:1fr}}
.cfg-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(340px,1fr));gap:14px}
.cfg-grid .card{margin:0;display:flex;flex-direction:column}
.cfg-txt{font-size:13px;color:#374151;line-height:1.55;flex:1}
.cfg-txt a{color:var(--or2);word-break:break-all}
.cfg-acts{display:flex;gap:8px;flex-wrap:wrap;margin-top:12px}
.av-sec{margin-bottom:14px}
.av-item{display:flex;gap:10px;align-items:flex-start;padding:9px 12px;border-bottom:1px solid var(--line2);font-size:13px}
.av-item:last-child{border-bottom:0}
.av-item .btn{margin-left:8px}
.av-item .ic{width:22px;height:22px;border-radius:50%;display:grid;place-items:center;font-size:12px;font-weight:900;flex:none;background:#E0F2FE;color:#075985}
.av-item.warn .ic{background:#FEF3C7;color:#92400E}.av-item.err .ic{background:#FEE2E2;color:#991B1B}.av-item.ok .ic{background:#DCFCE7;color:#166534}
.av-item .tm{margin-left:auto;color:var(--mut);font-size:11px;white-space:nowrap}
.side-note b{color:var(--or)}
.content{min-width:0;display:flex;flex-direction:column}
.topbar{background:var(--bk);color:#fff;display:flex;align-items:center;gap:12px;padding:12px 22px;position:sticky;top:0;z-index:40;border-bottom:3px solid var(--or)}
.topbar .crumb{font-size:12px;color:#A1A1AA}
.topbar .crumb b{display:block;color:#fff;font-size:17px;margin-top:1px}
.logo-m{display:none}
.top-actions{margin-left:auto;display:none;gap:6px}
.top-actions .btn{background:#26262B;color:#fff;border-color:#3A3A40;padding:7px 10px}
.sync{display:flex;align-items:center;gap:10px;flex-wrap:wrap;padding:8px 22px;font-size:12.5px;border-bottom:1px solid var(--line);background:#fff}
.sync .dot{width:9px;height:9px;border-radius:50%;background:#9CA3AF;flex:none}
.sync.ok .dot{background:var(--ok)} .sync.chg{background:#FFF7ED} .sync.chg .dot{background:var(--or);box-shadow:0 0 0 4px rgba(247,147,30,.2)} .sync.err{background:#FEF2F2} .sync.err .dot{background:var(--err)}
.sync.load .dot{animation:pulse 1s infinite}@keyframes pulse{50%{opacity:.3}}
.sync .acts{margin-left:auto;display:flex;gap:6px}
.sync.err{padding-top:6px;padding-bottom:6px}
.notice{background:var(--or-soft);color:#7C3E00;border-bottom:1px solid #FBD9B0;padding:7px 22px;font-size:12px;display:flex;gap:8px;align-items:center}
.notice b{color:#5C2E00}
main{max-width:1560px;width:100%;margin:0 auto;padding:18px 22px 28px}
.tab{display:none}.tab.on{display:block;animation:fade .25s ease}
@keyframes fade{from{opacity:0;transform:translateY(4px)}to{opacity:1;transform:none}}
/* ---------- generic ---------- */
.card{background:var(--card);border-radius:var(--r);box-shadow:var(--sh);padding:14px;margin-bottom:14px;border:1px solid var(--line2)}
.card-h{display:flex;align-items:center;gap:10px;flex-wrap:wrap;margin-bottom:10px}
.card-h h3{margin:0;font-size:15px}
.card-h .hint{color:var(--mut);font-size:12px}
.card-h .sp{flex:1}
.btn{border:1px solid var(--line);background:#fff;color:var(--tx);padding:7px 12px;border-radius:8px;cursor:pointer;font-weight:700;display:inline-flex;align-items:center;gap:6px;white-space:nowrap;transition:all .15s}
.btn:hover{border-color:var(--or);color:var(--or2)}
.btn.primary{background:var(--or);border-color:var(--or);color:#fff}
.btn.primary:hover{background:var(--or2);color:#fff}
.btn.dark{background:var(--bk);border-color:var(--bk);color:#fff}
.btn.dark:hover{color:var(--or)}
.btn.ghost{background:transparent}
.btn.danger{color:var(--err)}
.btn.sm{padding:4px 8px;font-size:12px}
.btn:disabled{opacity:.45;cursor:not-allowed}
.toolbar{display:flex;align-items:center;gap:10px;flex-wrap:wrap;margin-bottom:12px}
.toolbar h2{margin:0;font-size:20px}
.toolbar .sub{color:var(--mut);font-size:12px}
.toolbar .acts{margin-left:auto;display:flex;gap:6px;flex-wrap:wrap}
.filters{display:grid;grid-template-columns:repeat(auto-fit,minmax(170px,1fr));gap:10px;align-items:end}
.filters label,.fld{display:flex;flex-direction:column;gap:4px;font-size:11px;font-weight:700;color:var(--mut);text-transform:uppercase;letter-spacing:.3px}
.filters input,.filters select,.fld input,.fld select{border:1px solid var(--line);border-radius:8px;padding:7px 9px;background:#fff;color:var(--tx);text-transform:none;letter-spacing:0;font-weight:400;min-height:34px;width:100%}
.filters input:focus,.filters select:focus,.fld input:focus,.fld select:focus{outline:2px solid #FBD9B0;border-color:var(--or)}
.chip{display:inline-flex;align-items:center;gap:4px;padding:2px 8px;border-radius:999px;font-size:11px;font-weight:700;background:#F3F4F6;color:#374151;white-space:nowrap}
.chip.ok{background:#DCFCE7;color:#166534}.chip.err{background:#FEE2E2;color:#991B1B}.chip.warn{background:#FEF3C7;color:#92400E}.chip.or{background:var(--or-soft);color:#9A4A00}.chip.info{background:#E0F2FE;color:#075985}
.muted{color:var(--mut)}
.seg{display:inline-flex;background:#F3F4F6;border-radius:9px;padding:3px;gap:2px;flex-wrap:wrap}
.seg button{border:0;background:transparent;padding:5px 10px;border-radius:7px;font-weight:700;color:#4B5563;cursor:pointer;font-size:12px}
.seg button.on{background:#fff;color:var(--bk);box-shadow:0 1px 3px rgba(0,0,0,.12)}
/* ---------- multiselect ---------- */
.ms{position:relative}
.ms-btn{width:100%;display:flex;align-items:center;justify-content:space-between;gap:6px;border:1px solid var(--line);border-radius:8px;background:#fff;padding:7px 9px;min-height:34px;cursor:pointer;text-transform:none;font-weight:400;color:var(--tx);letter-spacing:0;text-align:left}
.ms-btn .ms-txt{overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.ms-btn .cnt{background:var(--or);color:#fff;border-radius:999px;font-size:10px;padding:1px 6px;font-weight:700}
.ms-pop{position:absolute;z-index:60;top:calc(100% + 4px);left:0;min-width:100%;width:max-content;max-width:min(340px,calc(100vw - 24px));background:#fff;border:1px solid var(--line);border-radius:10px;box-shadow:0 12px 30px rgba(0,0,0,.15);padding:8px;text-transform:none;letter-spacing:0;font-weight:400}
.ms-pop input{width:100%;margin-bottom:6px}
.ms-quick{display:flex;gap:4px;flex-wrap:wrap;margin-bottom:6px}
.ms-quick button{border:1px solid var(--line);background:#F9FAFB;border-radius:6px;padding:3px 7px;font-size:11px;cursor:pointer;font-weight:700}
.ms-quick button:hover{border-color:var(--or);color:var(--or2)}
.ms-list{max-height:260px;overflow:auto}
.ms-list label{display:flex;flex-direction:row;align-items:center;gap:8px;padding:6px;border-radius:6px;cursor:pointer;font-size:13px;color:var(--tx);text-transform:none;font-weight:400;letter-spacing:0}
.ms-list label:hover{background:var(--or-soft)}
.ms-list label small{color:var(--mut);font-size:11px}
.ms-list input{width:auto;min-height:0;margin:0;accent-color:var(--or)}
/* ---------- KPIs ---------- */
.kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:10px;margin-bottom:14px}
.kpis.k4{grid-template-columns:repeat(4,minmax(0,1fr))}
@media (max-width:640px){.kpis.k4{grid-template-columns:repeat(2,minmax(0,1fr))}}
.kpi{background:var(--card);border-radius:var(--r);padding:12px 14px;box-shadow:var(--sh);border:1px solid var(--line2);position:relative;overflow:hidden}
.kpi:before{content:"";position:absolute;left:0;top:0;bottom:0;width:4px;background:var(--c,var(--or))}
.kpi .l{font-size:11px;color:var(--mut);font-weight:700;text-transform:uppercase;letter-spacing:.3px}
.kpi .v{font-size:26px;font-weight:900;margin-top:4px;font-variant-numeric:tabular-nums}
.kpi .s{font-size:11px;color:var(--mut);margin-top:2px}
.kpi.dark{background:linear-gradient(135deg,#111 0%,#26262B 100%);color:#fff;border:0}
.kpi.dark .l,.kpi.dark .s{color:#A1A1AA}
.kpi.dark .v{color:var(--or)}
/* ---------- charts ---------- */
.chart-scroll{overflow-x:auto;overflow-y:hidden;-webkit-overflow-scrolling:touch;padding-bottom:4px}
.chart-scroll::-webkit-scrollbar{height:8px}.chart-scroll::-webkit-scrollbar-thumb{background:#D1D5DB;border-radius:9px}
.chart-frame{border:1px solid var(--line);border-radius:12px;background:linear-gradient(180deg,#FCFCFD,#fff);padding:12px 10px 6px}
.chart-inner{position:relative;height:360px;margin:0 auto}
.empty{padding:40px 10px;text-align:center;color:var(--mut)}
.empty b{display:block;color:var(--tx);font-size:15px;margin-bottom:4px}
.legend{display:flex;gap:6px;flex-wrap:wrap}
.legend button{display:inline-flex;align-items:center;gap:6px;border:1px solid var(--line);background:#fff;border-radius:999px;padding:4px 10px;font-size:12px;font-weight:700;cursor:pointer;color:var(--tx)}
.legend button i{width:10px;height:10px;border-radius:3px;background:var(--c)}
.legend button.off{opacity:.4;text-decoration:line-through}
/* ---------- journey ---------- */
.journey{background:#fff}
.jr-grid{display:grid;grid-template-columns:250px minmax(0,1fr) 210px;gap:14px;align-items:stretch}
.jr-side{border:1px solid var(--line);border-radius:12px;background:#FAFAFB;padding:10px;display:flex;flex-direction:column;gap:8px;max-height:452px;min-height:0}
.jr-side-h{font-size:11px;font-weight:900;text-transform:uppercase;letter-spacing:.5px;color:var(--mut);display:flex;justify-content:space-between;align-items:center;padding:2px 2px 0}
.jr-q{border:1px solid var(--line);border-radius:8px;padding:7px 9px;width:100%}
.jr-list{overflow:auto;display:flex;flex-direction:column;gap:6px;min-height:0;padding-right:2px}
.jr-item{display:grid;grid-template-columns:minmax(0,1fr) auto;grid-template-areas:"nm lv" "sk lv";gap:3px 10px;text-align:left;border:1px solid var(--line);background:#fff;border-radius:10px;padding:8px 10px;cursor:pointer;transition:border-color .15s,box-shadow .15s}
.jr-item .nm{grid-area:nm;font-weight:700;font-size:12.5px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;color:var(--tx)}
.jr-item .sk{grid-area:sk;height:22px}
.jr-item svg{width:100%;height:22px;display:block}
.jr-item .lv{grid-area:lv;align-self:center;text-align:right;font-weight:900;font-size:16px;color:var(--tx);font-variant-numeric:tabular-nums}
.jr-item .lv small{display:block;font-size:10px;color:var(--mut);font-weight:400}
.jr-item:hover{border-color:#FBC98A}
.jr-item.on{border-color:var(--or);background:var(--or-soft);box-shadow:inset 4px 0 0 var(--or)}
.jr-select{display:none;border:1px solid var(--line);border-radius:8px;padding:8px 10px;width:100%;font-weight:700}
.jr-main{min-width:0;border:1px solid var(--line);border-radius:12px;background:#fff;padding:12px 14px;display:flex;flex-direction:column}
.jr-top{display:flex;align-items:center;gap:12px;flex-wrap:wrap;margin-bottom:10px}
.jr-obra{font-weight:900;font-size:16px;min-width:0}
.jr-obra small{display:block;font-weight:400;color:var(--mut);font-size:11px;margin-top:2px}
.jr-read{margin-left:auto;display:flex;align-items:center;gap:10px;background:var(--bk);color:#fff;border-radius:10px;padding:7px 12px;flex-wrap:wrap;min-height:44px}
.jr-read .rv{font-size:24px;font-weight:900;color:var(--or);font-variant-numeric:tabular-nums;line-height:1}
.jr-read .ru{font-size:11px;color:#A1A1AA;margin-left:-6px}
.jr-read .rw{font-size:12px;font-weight:700}
.jr-read .rd{font-size:11px;color:#D4D4D8}
.jr-read .rtag{padding:3px 9px;border-radius:999px;font-size:10.5px;font-weight:900;letter-spacing:.5px;background:#2A2A30;color:#fff;white-space:nowrap}
.jr-frame{border-radius:10px;background:linear-gradient(180deg,#FFFBF5 0%,#fff 70%);border:1px solid var(--line2);flex:1}
.jr-chart{position:relative;height:318px}
.jr-ctrl{display:flex;align-items:center;gap:10px;margin-top:10px}
.jr-ctrl input[type=range]{flex:1;accent-color:var(--or);min-width:0}
.jr-miles{display:grid;grid-template-columns:1fr;grid-auto-rows:1fr;gap:10px}
.mile{border-radius:12px;padding:10px 12px;background:#fff;border:1px solid var(--line);position:relative;transition:transform .25s,box-shadow .25s;display:flex;flex-direction:column;justify-content:center}
.mile .t{display:flex;align-items:center;gap:6px;font-size:10.5px;font-weight:900;letter-spacing:.6px;text-transform:uppercase;color:var(--c)}
.mile .t i{width:9px;height:9px;border-radius:50%;background:var(--c);box-shadow:0 0 0 4px color-mix(in srgb,var(--c) 18%,transparent);flex:none}
.mile .v{font-size:24px;font-weight:900;margin-top:4px;font-variant-numeric:tabular-nums}
.mile .w{font-size:11.5px;color:var(--mut)}
.mile .d{position:absolute;right:10px;top:9px}
.mile.hit{animation:hit .9s ease}
@keyframes hit{0%{transform:scale(1)}30%{transform:scale(1.04);box-shadow:0 0 0 4px color-mix(in srgb,var(--c) 30%,transparent),0 10px 24px rgba(0,0,0,.12)}100%{transform:scale(1)}}
.play{width:38px;height:38px;border-radius:50%;border:0;background:var(--or);color:#fff;cursor:pointer;display:grid;place-items:center;box-shadow:0 4px 12px rgba(247,147,30,.45);flex:none}
.play:hover{background:var(--or2)}
.play svg{width:16px;height:16px}
/* ---------- tables ---------- */
.tbl-wrap{overflow:auto;border:1px solid var(--line);border-radius:10px;max-height:620px}
table{border-collapse:collapse;width:100%}
th,td{padding:7px 9px;border-bottom:1px solid var(--line2);text-align:center;white-space:nowrap;font-size:12.5px}
th{background:#18181B;color:#fff;font-size:11px;text-transform:uppercase;letter-spacing:.3px;position:sticky;top:0;z-index:2}
tbody tr:nth-child(even){background:#FAFAFB}
tbody tr:hover{background:var(--or-soft)}
td.l{text-align:left}
tr.tot td{background:#111 !important;color:#fff;font-weight:900}
tr.tot td.hl{color:var(--or)}
td.num{font-variant-numeric:tabular-nums;font-weight:700}
.pager{display:flex;align-items:center;gap:8px;justify-content:flex-end;margin-top:8px;font-size:12px;color:var(--mut)}
.alert{border-radius:10px;padding:10px 12px;font-size:13px;margin-bottom:10px;border:1px solid}
.alert.info{background:#F0F9FF;border-color:#BAE6FD;color:#075985}
.alert.warn{background:#FFFBEB;border-color:#FDE68A;color:#92400E}
.alert.err{background:#FEF2F2;border-color:#FECACA;color:#991B1B}
.alert.ok{background:#F0FDF4;border-color:#BBF7D0;color:#166534}
/* ---------- resumo ---------- */
.sub-tabs{margin-bottom:12px}
.obra-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(360px,1fr));gap:12px}
.obra-card h4{margin:0 0 8px;font-size:14px;display:flex;align-items:center;gap:8px}
.obra-card .pair{display:grid;grid-template-columns:1fr 1fr;gap:8px}
.obra-card .pair > div{position:relative;height:210px}
.obra-card .cap{font-size:11px;font-weight:700;color:var(--mut);text-align:center;text-transform:uppercase;margin-bottom:2px}
/* ---------- modal ---------- */
.modal-bg{position:fixed;inset:0;background:rgba(10,10,12,.55);z-index:100;display:flex;align-items:flex-start;justify-content:center;padding:28px 12px;overflow:auto;animation:fade .15s}
.modal{background:#fff;border-radius:14px;width:100%;max-width:720px;box-shadow:0 20px 50px rgba(0,0,0,.3);overflow:hidden}
.modal.wide{max-width:1180px}
.modal-h{background:var(--bk);color:#fff;padding:12px 16px;display:flex;align-items:center;gap:10px;border-bottom:3px solid var(--or)}
.modal-h h3{margin:0;font-size:15px;flex:1}
.modal-h .x{background:none;border:0;color:#fff;font-size:22px;cursor:pointer;line-height:1}
.modal-b{padding:16px;max-height:calc(100vh - 190px);overflow:auto}
.modal-f{padding:12px 16px;border-top:1px solid var(--line);display:flex;gap:8px;justify-content:flex-end;flex-wrap:wrap;background:#FAFAFA}
.form-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(200px,1fr));gap:12px}
.fld .req{color:var(--err)}
.fld .err-msg{color:var(--err);font-size:11px;text-transform:none;font-weight:700;letter-spacing:0}
.fld.bad input,.fld.bad select{border-color:var(--err);background:#FEF2F2}
.week-in{display:flex;gap:6px}
.week-in input[type=date]{max-width:44px;padding:6px 4px;color:transparent}
.week-in input[type=date]::-webkit-datetime-edit{display:none}
.stat-row{display:flex;gap:8px;flex-wrap:wrap;margin:10px 0}
.err-list{max-height:180px;overflow:auto;font-size:12px;border:1px solid #FECACA;background:#FEF2F2;border-radius:8px;padding:8px 10px;margin:0;list-style:none}
.err-list li{padding:2px 0}
.upd-list{max-height:150px;overflow:auto;font-size:12px;border:1px solid #FDE68A;background:#FFFBEB;border-radius:8px;padding:8px 10px;margin:0;list-style:none}
.drop{border:2px dashed #D1D5DB;border-radius:12px;padding:18px;text-align:center;color:var(--mut);cursor:pointer;transition:all .15s}
.drop:hover,.drop.over{border-color:var(--or);background:var(--or-soft);color:#7C3E00}
.cols2{display:grid;grid-template-columns:1fr 1fr;gap:14px}
.toast-box{position:fixed;right:16px;bottom:16px;z-index:200;display:flex;flex-direction:column;gap:8px;max-width:380px}
.toast{background:#111;color:#fff;padding:10px 14px;border-radius:10px;box-shadow:0 10px 25px rgba(0,0,0,.25);font-size:13px;border-left:4px solid var(--or);animation:fade .2s}
.toast.err{border-left-color:var(--err)}.toast.ok{border-left-color:var(--ok)}
#print-area,#offscreen{display:none}
#offscreen{position:fixed;left:-20000px;top:0}
/* ---------- responsive ---------- */
@media (max-width:1280px){
  .jr-grid{grid-template-columns:220px minmax(0,1fr)}
  .jr-miles{grid-column:1 / -1;grid-template-columns:repeat(4,1fr);grid-auto-rows:auto}
}
@media (max-width:900px){
  .app{display:block;padding-bottom:66px}
  .side{position:fixed;top:auto;bottom:0;left:0;right:0;height:62px;flex-direction:row;border-right:0;border-top:3px solid var(--or)}
  .side-brand,.side-foot,.side-nav .cap{display:none}
  .side-nav{flex-direction:row;padding:4px 6px;width:100%;gap:2px;overflow-x:auto;-webkit-overflow-scrolling:touch}
  .side-nav button{flex:0 0 auto;width:auto;min-width:74px;max-width:96px;font-size:10.5px;white-space:normal;line-height:1.15;text-align:center}
  .side-nav button{flex-direction:column;gap:3px;font-size:11px;padding:6px 4px;justify-content:center}
  .side-nav button.on:before{left:20%;right:20%;top:-7px;bottom:auto;height:3px;width:auto;border-radius:0 0 3px 3px}
  .logo-m{display:flex;align-items:center;cursor:pointer}
  .logo-m img{height:22px;width:auto;display:block}
  .logo-m .logo-text{font-size:16px}
  .top-actions{display:flex}
  .topbar{padding:9px 12px}
  .topbar .crumb{border-left:1px solid #3a3a3a;padding-left:10px}
  .topbar .crumb b{font-size:15px}
  .jr-grid{grid-template-columns:1fr}
  .jr-side{max-height:none;padding:0;border:0;background:none}
  .jr-side-h,.jr-list,.jr-q{display:none}
  .jr-select{display:block}
  .jr-read{margin-left:0;width:100%}
  .jr-miles{grid-template-columns:repeat(2,1fr)}
  .cols2{grid-template-columns:1fr}
}
@media (max-width:640px){
  .topbar .crumb span{display:none}
  main{padding:10px}
  .toolbar .acts{margin-left:0;width:100%}
  .toolbar .acts .btn{flex:1 1 auto;justify-content:center}
  .filters{grid-template-columns:1fr 1fr;gap:8px}
  .filters label:nth-last-of-type(1),.filters [data-a=clear]{grid-column:1 / -1}
  .chart-inner{height:330px}
  .jr-chart{height:250px}
  .chart-frame{padding:8px 4px 4px}
  .kpi .v{font-size:22px}
  .obra-grid{grid-template-columns:1fr}
  .obra-card .pair > div{height:180px}
  .modal-bg{padding:0}
  .modal{border-radius:0;min-height:100vh}
  .modal-b{max-height:none}
}
@media (prefers-reduced-motion:reduce){*{animation:none !important;transition:none !important}}
/* ---------- print ---------- */
@media print{
  @page{size:A4 landscape;margin:8mm}
  body{background:#fff}
  body > *:not(#print-area){display:none !important}
  #print-area{display:block !important}
  *{-webkit-print-color-adjust:exact;print-color-adjust:exact}
}
#print-area{font-family:Arial,Helvetica,sans-serif;color:#111}
.p-page{page-break-after:always;break-after:page}
.p-page:last-child{page-break-after:auto;break-after:auto}
.p-head{background:#0E0E10;color:#fff;display:flex;align-items:center;gap:12px;padding:8px 12px;border-bottom:3px solid #F7931E;margin-bottom:6px}
.p-head .lg{font-weight:900;letter-spacing:2px;font-size:16px}.p-head .lg b{color:#F7931E}
.p-head img{height:24px;width:auto}
.p-head .tt{font-weight:700;font-size:13px}
.p-head .dt{margin-left:auto;font-size:10px;color:#ccc}
.p-filt{font-size:9.5px;color:#333;margin-bottom:6px;border:1px solid #ddd;border-radius:6px;padding:5px 8px;background:#FAFAFA}
.p-filt b{color:#000}
.p-img{width:100%;max-height:150mm;object-fit:contain;display:block;margin:0 auto}
.p-tbl{width:100%;border-collapse:collapse;margin-bottom:6px}
.p-tbl th{background:#111 !important;color:#fff;font-size:7.5px;padding:3px 3px;position:static}
.p-tbl td{font-size:7.5px;padding:2px 3px;border-bottom:1px solid #e5e5e5;white-space:normal}
.p-tbl tr.tot td{background:#111 !important;color:#fff}
.p-sec{font-size:12px;font-weight:900;margin:6px 0 4px;border-left:4px solid #F7931E;padding-left:6px}
.p-note{font-size:9px;color:#555;margin:4px 0}
/* relatório de atividades */
.pdot{display:inline-block;width:9px;height:9px;border-radius:50%;margin-right:5px;vertical-align:middle}
.p-vg td{height:34px}
.p-k4{display:grid;grid-template-columns:repeat(4,1fr);gap:8px;margin:4px 0 8px}
.p-k4 .pk{border:1px solid #E5E7EB;border-left:4px solid #F7931E;border-radius:8px;padding:6px 10px}
.p-k4 .lb{font-size:8px;font-weight:700;letter-spacing:.5px;text-transform:uppercase;color:#6B7280}.p-k4 .vl{font-size:20px;font-weight:900;line-height:1.1}.p-k4 .vl small{font-size:9px;font-weight:400;color:#6B7280}.p-k4 .sb{font-size:8.5px;color:#374151}
.p-att3{display:grid;grid-template-columns:repeat(3,1fr);gap:10px}
.p-ul{list-style:none;margin:0;padding:0}.p-ul li{display:flex;justify-content:space-between;gap:8px;font-size:8.5px;padding:2px 0;border-bottom:1px solid #F0F1F3}.p-ul li em{display:block;font-style:normal;color:#6B7280;font-size:7.5px}.p-ul li b{white-space:nowrap}
.p-h3{display:flex;gap:16px;align-items:flex-start}.p-h3b{flex:1;min-width:0}.p-h3b img{width:100%;display:block}
.p-h3top{display:flex;align-items:center;gap:8px;margin-bottom:4px}.p-h3top .p-h3t{margin:0}#print-area .jr-read{display:flex;align-items:center;gap:8px;background:#0E0E10 !important;color:#fff;border-radius:8px;padding:4px 10px;margin-left:auto}#print-area .jr-read .rv{font-size:16px;font-weight:900;color:#F7931E}#print-area .jr-read .ru,#print-area .jr-read .rd{font-size:8.5px;color:#D4D4D8}#print-area .jr-read .rw{font-size:9px;font-weight:700}#print-area .jr-read .rtag{font-size:8px;font-weight:900;padding:2px 7px;border-radius:999px;color:#fff}
.p-h3t{font-weight:900;font-size:12px;margin:2px 0 4px;padding-left:6px;border-left:4px solid #F7931E}
#print-area .h3t{font-size:9.5px;margin-top:6px}#print-area .h3t th{font-size:9px;padding:3px 6px}#print-area .h3t td{padding:3px 6px;font-size:9.5px}#print-area .h3t + div{font-size:9px !important}
.p-kpi{display:flex;align-items:center;gap:18px;border:1px solid #E5E7EB;border-left:5px solid #F7931E;border-radius:8px;padding:10px 16px;margin:4px 0 8px;background:#FFFBF5}
.p-kpi .lb{font-size:9px;font-weight:700;letter-spacing:.6px;text-transform:uppercase;color:#6B7280}
.p-kpi .vl{font-size:30px;font-weight:900;color:#E86F00;line-height:1.05}
.p-kpi .sb{font-size:11px;color:#374151;margin-top:2px}
.p-kpi .bar{flex:1;height:12px;background:#EEF0F3;border-radius:6px;position:relative;overflow:hidden}
.p-kpi .bar i{position:absolute;left:0;top:0;bottom:0;background:#F7931E;border-radius:6px}
.p-kpi .bar u{position:absolute;top:-2px;bottom:-2px;width:2px;background:#111}
.p-at{width:100%;border-collapse:collapse;table-layout:fixed}
.p-at th{white-space:normal;background:#111 !important;color:#fff;font-size:7.2px;font-weight:700;text-transform:uppercase;height:22px;padding:0 3px;text-align:center;vertical-align:middle;line-height:1.15;border-right:1px solid #333}
.p-at td{height:24px;padding:0 4px;white-space:normal;font-size:7.6px;line-height:1.2;vertical-align:middle;text-align:center;border-bottom:1px solid #E5E7EB;border-right:1px solid #F0F1F3;overflow:hidden}
.p-at td.l{text-align:left}
.p-at tbody tr:nth-child(even) td{background:#F8F9FA}
.p-at tr{page-break-inside:avoid;break-inside:avoid}
.p-at thead{display:table-header-group}
.p-at .c2{display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;overflow:hidden;word-break:break-word}
.p-at .d{display:block;font-size:6.6px;color:#DC2626;font-weight:700}
.p-at .st-err{color:#DC2626;font-weight:700}.p-at .st-ok{color:#16A34A;font-weight:700}.p-at .st-or{color:#B45309;font-weight:700}
</style>
</head>
<body>
<div class="app">
<aside class="side">
  <div class="side-brand"><div class="logo" id="logoBox" title="Clique para alterar o logo"></div><div class="side-title">Controle de Histogramas<small>Planejamento e controle de obras</small></div></div>
  <nav class="side-nav" id="tabs">
    <div class="cap">Gerencial</div>
    <button data-tab="vg"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M3 3v18h18"/><path d="M7 15l4-4 3 3 5-6"/><circle cx="19" cy="8" r="1.4"/></svg><span>Visão gerencial</span></button>
    <div class="cap" style="margin-top:10px">Painéis</div>
    <button data-tab="labor" class="on"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="9" cy="8" r="3.2"/><path d="M3 20c0-3.3 2.7-6 6-6s6 2.7 6 6"/><circle cx="17" cy="9" r="2.6"/><path d="M15.5 14.2c3 .2 5.5 2.6 5.5 5.8"/></svg><span>Mão de obra</span></button>
    <button data-tab="equip"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M2 17V11h7l2 3h4v3"/><path d="M9 11V7h4l3 4"/><path d="M15 14l4-6 3 1-2 5"/><circle cx="6" cy="18" r="2"/><circle cx="13" cy="18" r="2"/></svg><span>Equipamentos</span></button>
    <button data-tab="resumo"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="3" width="7" height="9" rx="1.5"/><rect x="14" y="3" width="7" height="5" rx="1.5"/><rect x="14" y="12" width="7" height="9" rx="1.5"/><rect x="3" y="16" width="7" height="5" rx="1.5"/></svg><span>Resumo</span></button>
    <div class="cap" style="margin-top:10px">Acompanhamento</div>
    <button data-tab="ativ"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M9 6h11M9 12h11M9 18h11"/><path d="M3 6l1.5 1.5L7 5M3 12l1.5 1.5L7 11"/><circle cx="5" cy="18" r="1.3"/></svg><span>Atividades</span></button>
    <div class="cap" style="margin-top:10px">Sistema</div>
    <button data-tab="config"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="3"/><path d="M19.4 15a1.7 1.7 0 0 0 .3 1.8l.1.1a2 2 0 1 1-2.8 2.8l-.1-.1a1.7 1.7 0 0 0-1.8-.3 1.7 1.7 0 0 0-1 1.5V21a2 2 0 1 1-4 0v-.1a1.7 1.7 0 0 0-1.1-1.5 1.7 1.7 0 0 0-1.8.3l-.1.1a2 2 0 1 1-2.8-2.8l.1-.1a1.7 1.7 0 0 0 .3-1.8 1.7 1.7 0 0 0-1.5-1H3a2 2 0 1 1 0-4h.1a1.7 1.7 0 0 0 1.5-1.1 1.7 1.7 0 0 0-.3-1.8l-.1-.1a2 2 0 1 1 2.8-2.8l.1.1a1.7 1.7 0 0 0 1.8.3H9a1.7 1.7 0 0 0 1-1.5V3a2 2 0 1 1 4 0v.1a1.7 1.7 0 0 0 1 1.5 1.7 1.7 0 0 0 1.8-.3l.1-.1a2 2 0 1 1 2.8 2.8l-.1.1a1.7 1.7 0 0 0-.3 1.8V9a1.7 1.7 0 0 0 1.5 1H21a2 2 0 1 1 0 4h-.1a1.7 1.7 0 0 0-1.5 1z"/></svg><span>Configurações</span></button>
    <button data-tab="avisos"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M18 8a6 6 0 0 0-12 0c0 7-3 9-3 9h18s-3-2-3-9"/><path d="M13.7 21a2 2 0 0 1-3.4 0"/></svg><span>Avisos</span><i class="badge" id="avisosBadge" hidden></i></button>
  </nav>
  <div class="side-foot">
    <button class="side-btn" id="btnSync" style="border-color:#7C3E00">⟳ Sincronizar planilha</button>
    <div class="side-note" id="sideMini"></div>
  </div>
</aside>
<div class="content">
<header class="topbar">
  <div class="logo-m" id="logoBoxM"></div>
  <div class="crumb"><span>Controle de Histogramas</span><b id="pageTitle">Mão de obra</b></div>
  <div class="top-actions"><button class="btn" id="btnSyncM" title="Sincronizar planilha">⟳</button></div>
</header>
<div id="syncBar"></div>
<main>
  <section class="tab on" id="tab-labor"></section>
  <section class="tab" id="tab-equip"></section>
  <section class="tab" id="tab-resumo"></section>
  <section class="tab" id="tab-painel"></section>
  <section class="tab" id="tab-ativ"></section>
  <section class="tab" id="tab-vg"></section>
  <section class="tab" id="tab-lps"></section>
  <section class="tab" id="tab-restr"></section>
  <section class="tab" id="tab-config">
    <div class="toolbar"><div><h2>Configurações</h2><div class="sub">Conexões, dados e backup do aplicativo</div></div></div>
    <div class="cfg-grid">
      <div class="card"><div class="card-h"><h3>📄 Planilha conectada</h3></div><div id="cfgSheet" class="cfg-txt"></div>
        <div class="cfg-acts"><button class="btn primary" id="btnSync2">⟳ Sincronizar agora</button><button class="btn" id="btnSyncCfg">⚙ Conexão com a planilha</button></div></div>
      <div class="card"><div class="card-h"><h3>🗄 Banco de dados</h3></div><div id="dbStatus" class="cfg-txt"></div>
        <div class="cfg-acts"><button class="btn" data-db="cfg">🗄 Banco de dados</button></div></div>
      <div class="card"><div class="card-h"><h3>💾 Backup</h3></div><div class="cfg-txt">Cópia de segurança de todos os cadastros, semanas e lançamentos manuais (arquivo .json). Restaurar substitui os dados atuais, com confirmação.</div>
        <div class="cfg-acts"><button class="btn" id="btnBackup">⬇ Baixar backup</button><button class="btn" id="btnRestore">⬆ Restaurar backup</button></div></div>
      <div class="card"><div class="card-h"><h3>🧹 Limpar dados</h3></div><div class="cfg-txt">Apaga os dados do app para recomeçar a importação — tudo ou só uma obra, por tipo de dado. Antes de apagar é criado um ponto de restauração. A planilha conectada não é alterada.</div>
        <div class="cfg-acts"><button class="btn danger" id="btnClearData">🧹 Limpar dados…</button></div></div>
      <div class="card"><div class="card-h"><h3>🕘 Pontos de restauração</h3></div><div class="cfg-txt">Criados automaticamente antes de cada importação e restauração (os 5 mais recentes, neste navegador). Use para desfazer uma importação.<div id="snapList" style="margin-top:8px"></div></div></div>
      <div class="card"><div class="card-h"><h3>📦 Armazenamento</h3></div><div class="cfg-txt"><span id="sideLocal">Dados salvos <b>somente neste navegador e neste endereço</b> (<b id="curHost"></b>). Outro navegador, outro computador ou outro endereço do app (inclusive os links de publicação da Vercel) <b>não</b> enxergam estes dados. Baixe o backup regularmente.</span><div id="storeUse" style="margin-top:6px"></div></div></div>
      <div class="card"><div class="card-h"><h3>🖼 Logo</h3></div><div class="cfg-txt">Logo exibido no menu, no celular e nos relatórios PDF. A troca vale para este navegador.</div>
        <div class="cfg-acts"><button class="btn" id="btnLogo">Alterar logo</button><button class="btn ghost" id="btnLogoReset">Voltar ao logo COSAMPA</button></div></div>
    </div>
  </section>
  <section class="tab" id="tab-avisos">
    <div class="toolbar"><div><h2>Avisos</h2><div class="sub">Mensagens e orientações do aplicativo, registradas aqui para não ocupar as telas de trabalho</div></div>
      <div class="acts"><button class="btn" id="btnClearLog">Limpar histórico</button></div></div>
    <div id="avisosBody"></div>
  </section>
</main>
</div>
</div>
<div class="toast-box" id="toasts"></div>
<div id="print-area"></div>
<div id="offscreen"></div>
<input type="file" id="fileRestore" accept=".json,application/json" hidden>
<input type="file" id="fileLogo" accept="image/*" hidden>

<script>
'use strict';
/* =====================================================================
   UTILITÁRIOS
   ===================================================================== */
const $ = (s, r = document) => r.querySelector(s);
const $$ = (s, r = document) => [...r.querySelectorAll(s)];
const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const uid = () => Date.now().toString(36) + Math.random().toString(36).slice(2, 9);
const norm = s => String(s ?? '').normalize('NFD').replace(/[̀-ͯ]/g, '').toUpperCase().replace(/\s+/g, ' ').trim();
const normHeader = s => String(s ?? '').normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().replace(/[^a-z0-9]/g, '');
const digits = s => String(s ?? '').replace(/\D/g, '');
const clean = s => String(s ?? '').replace(/\s+/g, ' ').trim();
const fmtN = n => n == null ? '—' : Number(n).toLocaleString('pt-BR');
const pad2 = n => String(n).padStart(2, '0');
const isEmptyVal = v => v == null || String(v).trim() === '';
const DASHES = new Set(['-', '–', '—', '--', '---']);

function fmtBR(ymd){ if(!ymd) return ''; const [y,m,d] = ymd.split('-'); return `${d}/${m}/${y}`; }
function ymdUTC(dt){ return `${dt.getUTCFullYear()}-${pad2(dt.getUTCMonth()+1)}-${pad2(dt.getUTCDate())}`; }
function validYMD(y,m,d){ const dt = new Date(Date.UTC(y, m-1, d)); return dt.getUTCFullYear()===y && dt.getUTCMonth()===m-1 && dt.getUTCDate()===d ? ymdUTC(dt) : null; }
function serialToYMD(n){ if(!(n > 1000 && n < 120000)) return null; const dt = new Date(Date.UTC(1899,11,30) + Math.round(n)*86400000); return ymdUTC(dt); }
/** Converte valor de data. Retorna {v:'YYYY-MM-DD'|'' , err?:string, dash?:true} */
function parseDate(raw, order){
  if (raw == null) return {v:''};
  if (typeof raw === 'number') { const v = serialToYMD(raw); return v ? {v} : {v:'', err:'Data inválida'}; }
  const s = String(raw).trim();
  if (!s) return {v:''};
  if (DASHES.has(s)) return {v:'', dash:true};
  let m;
  if ((m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})(?:[T\s].*)?$/))) { const v = validYMD(+m[1], +m[2], +m[3]); return v ? {v} : {v:'', err:`Data inválida "${s}"`}; }
  if ((m = s.match(/^(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{2}|\d{4})(?:\s.*)?$/))) { let y = +m[3]; if (y < 100) y += 2000; const v = order === 'mdy' ? validYMD(y, +m[1], +m[2]) : validYMD(y, +m[2], +m[1]); return v ? {v} : {v:'', err:`Data inválida "${s}"`}; }
  if (/^\d{5}(\.\d+)?$/.test(s)) { const v = serialToYMD(parseFloat(s)); if (v) return {v}; }
  return {v:'', err:`Data inválida "${s}"`};
}
/* ---------- semanas ISO ---------- */
function isoWeek(ymd){
  const [y,m,d] = ymd.split('-').map(Number);
  const dt = new Date(Date.UTC(y, m-1, d));
  const day = (dt.getUTCDay() + 6) % 7;
  dt.setUTCDate(dt.getUTCDate() - day + 3);
  const wy = dt.getUTCFullYear();
  const jan4 = new Date(Date.UTC(wy, 0, 4));
  const w1 = new Date(jan4); w1.setUTCDate(jan4.getUTCDate() - ((jan4.getUTCDay()+6)%7) + 3);
  const wk = 1 + Math.round((dt - w1) / 604800000);
  return `${wy}-W${pad2(wk)}`;
}
const _wrCache = new Map();
function weekRange(w){
  if (_wrCache.has(w)) return _wrCache.get(w);
  const m = /^(\d{4})-W(\d{2})$/.exec(w || ''); if (!m) return null;
  const y = +m[1], wk = +m[2];
  const jan4 = new Date(Date.UTC(y, 0, 4));
  const mon = new Date(jan4); mon.setUTCDate(jan4.getUTCDate() - ((jan4.getUTCDay()+6)%7) + (wk-1)*7);
  const sun = new Date(mon); sun.setUTCDate(mon.getUTCDate() + 6);
  const r = {start: ymdUTC(mon), end: ymdUTC(sun)};
  if (isoWeek(r.start) !== w) { _wrCache.set(w, null); return null; }
  _wrCache.set(w, r); return r;
}
function parseWeek(raw){
  if (raw == null) return {v:''};
  if (typeof raw === 'number') { const d = serialToYMD(raw); return d ? {v: isoWeek(d)} : {v:'', err:'Semana inválida'}; }
  const s = String(raw).trim().toUpperCase(); if (!s) return {v:''};
  let m, y, w;
  if ((m = s.match(/^(\d{4})\s*[-\/ ]?\s*(?:W|S|SEM|SEMANA)\s*(\d{1,2})$/))) { y = +m[1]; w = +m[2]; }
  else if ((m = s.match(/^(?:W|S|SEM\.?|SEMANA)?\s*(\d{1,2})\s*[-\/ ]\s*(\d{4})$/))) { w = +m[1]; y = +m[2]; }
  if (y) { const v = `${y}-W${pad2(w)}`; return weekRange(v) ? {v} : {v:'', err:`Semana inexistente "${raw}"`}; }
  const d = parseDate(s); if (d.v) return {v: isoWeek(d.v)};
  return {v:'', err:`Semana inválida "${raw}" (use 2026-W40)`};
}
function weekLabel(w, long){ const r = weekRange(w); if (!r) return w; const a = fmtBR(r.start).slice(0,5), b = fmtBR(r.end).slice(0,5); return long ? `${w} (${a} a ${b})` : `${w} · ${a}–${b}`; }
function weekShort(w){ const m = /^(\d{4})-W(\d{2})$/.exec(w); return m ? `S${m[2]}/${m[1].slice(2)}` : w; }
function wrapText(str, max = 14, lines = 3){
  const words = String(str).split(/\s+/); const out = []; let cur = '';
  for (const wd of words) { if ((cur + ' ' + wd).trim().length > max && cur) { out.push(cur); cur = wd; } else cur = (cur + ' ' + wd).trim(); }
  if (cur) out.push(cur);
  if (out.length > lines) { const keep = out.slice(0, lines); keep[lines-1] = keep[lines-1].slice(0, max-1) + '…'; return keep; }
  return out;
}
/* ---------- Pontos de restauração (automáticos antes de cada importação/restauração) ---------- */
const LS_SNAP = 'cosampa_histogramas_pontos';
const Snap = {
  list(){ try { return JSON.parse(localStorage.getItem(LS_SNAP) || '[]'); } catch (e) { return []; } },
  take(label){
    const item = {t: new Date().toISOString(), label, counts: Object.fromEntries(ALLK.map(k => [k, Repo.data[k].length])), data: Object.fromEntries(ALLK.map(k => [k, Repo.pack(k)]))};
    let l = [item, ...this.list()].slice(0, 5);
    while (l.length) { try { localStorage.setItem(LS_SNAP, JSON.stringify(l)); return true; } catch (e) { l = l.slice(0, -1); } }
    return false;
  },
  async restore(i){
    const it = this.list()[i]; if (!it) return;
    const ok = await confirmBox('Restaurar ponto', `<p>Voltar os dados para <b>${new Date(it.t).toLocaleString('pt-BR')}</b> (${esc(it.label)})?</p><p>${it.counts.labor} cadastro(s) de mão de obra, ${it.counts.equip} de equipamentos e ${it.counts.manual} lançamento(s) manuais.</p><div class="alert warn">Os dados atuais${Repo.mode === 'db' ? ' do BANCO COMPARTILHADO' : ''} serão substituídos. Antes disso, um novo ponto com a situação atual é criado.</div>`, 'Restaurar', true);
    if (!ok) return;
    this.take('Antes de restaurar ponto');
    const nd = Object.fromEntries(ALLK.map(k => [k, Repo.unpack(it.data[k])]));
    if (Repo.replaceAll(nd)) { toast('Ponto de restauração aplicado.', 'ok'); refreshAll(); }
  }
};
const LS_LOG = 'cosampa_histogramas_avisos';
const Log = { items: (() => { try { return JSON.parse(localStorage.getItem(LS_LOG) || '[]'); } catch (e) { return []; } })(), unread: 0,
  add(msg, type){ this.items.unshift({t: new Date().toISOString(), msg, type}); this.items = this.items.slice(0, 200); this.unread++; try { localStorage.setItem(LS_LOG, JSON.stringify(this.items)); } catch (e) {} if (typeof updateBadge === 'function') updateBadge(); if (typeof UI !== 'undefined' && UI.tab === 'avisos') renderAvisos(); },
  clear(){ this.items = []; this.unread = 0; try { localStorage.removeItem(LS_LOG); } catch (e) {} } };
function toast(msg, type = ''){ Log.add(msg, type); const t = document.createElement('div'); t.className = 'toast ' + type; t.textContent = msg; $('#toasts').appendChild(t); setTimeout(() => t.remove(), type === 'err' ? 7000 : 4000); }
function download(name, blob){ const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = name; document.body.appendChild(a); a.click(); setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 500); }
const reduceMotion = window.matchMedia && window.matchMedia('(prefers-reduced-motion: reduce)').matches;
function countUp(el, to, dur = 850){
  if (to == null || isNaN(to)) { el.textContent = '—'; return; }
  const from = Number(el.dataset.v || 0); el.dataset.v = to;
  if (reduceMotion || from === to) { el.textContent = fmtN(to); return; }
  const t0 = performance.now();
  const tick = now => { const p = Math.min(1, (now - t0) / dur); const e = 1 - Math.pow(1 - p, 3); el.textContent = fmtN(Math.round(from + (to - from) * e)); if (p < 1) requestAnimationFrame(tick); };
  requestAnimationFrame(tick);
}
function todayWeek(){ const d = new Date(); return isoWeek(`${d.getFullYear()}-${pad2(d.getMonth()+1)}-${pad2(d.getDate())}`); }
const libsOk = () => typeof Chart !== 'undefined' && typeof XLSX !== 'undefined';

/* =====================================================================
   CONFIGURAÇÃO DOS TIPOS DE REGISTRO
   ===================================================================== */
const KINDS = {
  labor: {
    key:'labor', title:'Mão de obra', unit:'pessoas', catField:'tipo', start:'admissao', end:'demissao',
    catTitle:'Categoria', startLabel:'Admissão', idLabel:'Matrícula',
    cats:[ {k:'MOD', l:'MOD', full:'Direta — MOD', color:'#F7931E'}, {k:'MOI', l:'MOI', full:'Indireta — MOI', color:'#334155'}, {k:'TER', l:'Terceiros', full:'Terceiros', color:'#14B8A6'} ],
    fields:[
      {k:'obra', l:'Obra', req:1, al:['obra','obras','empreendimento','centrodecusto']},
      {k:'matricula', l:'Matrícula', id:1, al:['matricula','mat','chapa','registro']},
      {k:'nome', l:'Nome', req:1, al:['nome','colaborador','funcionario','nomecompleto','nomedocolaborador']},
      {k:'situacao', l:'Situação', al:['situacao','status']},
      {k:'semana', l:'Semana', req:1, t:'week', al:['semana','semanaiso','sem','week']},
      {k:'admissao', l:'Admissão', t:'date', al:['admissao','dataadmissao','dtadmissao','datadeadmissao']},
      {k:'demissao', l:'Demissão', t:'date', al:['demissao','datademissao','dtdemissao','desligamento','datadesligamento']},
      {k:'funcao', l:'Função', al:['funcao','cargo']},
      {k:'cpf', l:'CPF', id:1, al:['cpf']},
      {k:'sexo', l:'Sexo', al:['sexo','genero']},
      {k:'contratacao', l:'Contratação', al:['contratacao','tipodecontratacao','tipocontratacao','regime','vinculo']},
      {k:'cidade', l:'Cidade', al:['cidade','municipio']},
      {k:'estado', l:'Estado', al:['estado','uf']},
      {k:'telefone', l:'Telefone', id:1, al:['telefone','celular','fone','contato']},
      {k:'tipo', l:'Tipo de mão de obra', req:1, t:'cat', al:['tipodemaodeobra','tipomaodeobra','tipomo','tipo','categoria','maodeobra','classificacao']},
      {k:'disciplina', l:'Disciplina', al:['disciplina','setor']},
    ]
  },
  equip: {
    key:'equip', title:'Equipamentos', unit:'equipamentos', catField:'classificacao', start:'mobilizacao', end:'desmobilizacao',
    catTitle:'Classificação', startLabel:'Mobilização', idLabel:'Prefixo',
    cats:[ {k:'LA', l:'Linha amarela', full:'Linha amarela', color:'#EAB308'}, {k:'LB', l:'Linha branca', full:'Linha branca', color:'#64748B'}, {k:'PP', l:'P. porte', full:'P. porte', color:'#0EA5E9'} ],
    fields:[
      {k:'obra', l:'Obra', req:1, al:['obra','obras','empreendimento','centrodecusto']},
      {k:'prefixo', l:'Prefixo', req:1, id:1, al:['prefixo','codigo','frota']},
      {k:'descricao', l:'Descrição', al:['descricao','equipamento','descricaodoequipamento']},
      {k:'classificacao', l:'Classificação', req:1, t:'cat', al:['classificacao','linha','classe','categoria']},
      {k:'modelo', l:'Modelo', al:['modelo']},
      {k:'semana', l:'Semana', req:1, t:'week', al:['semana','semanaiso','sem','week']},
      {k:'mobilizacao', l:'Mobilização', t:'date', al:['mobilizacao','datamobilizacao','dtmobilizacao','datademobilizacao','entrada']},
      {k:'desmobilizacao', l:'Desmobilização', t:'date', al:['desmobilizacao','datadesmobilizacao','dtdesmobilizacao','saida']},
      {k:'proprietario', l:'Proprietário', al:['proprietario','locador','dono','fornecedor']},
      {k:'telefone', l:'Telefone', id:1, al:['telefone','celular','fone','contato']},
      {k:'renavam', l:'Renavam', id:1, al:['renavam']},
      {k:'placa', l:'Placa', id:1, al:['placa']},
      {k:'fabricante', l:'Fabricante', al:['fabricante','marca']},
      {k:'motor', l:'Nº motor', id:1, al:['nmotor','nomotor','numeromotor','numerodomotor','motor']},
      {k:'disciplina', l:'Disciplina', al:['disciplina','setor','frente','area']},
    ]
  }
};
const SD = {k:'SD', l:'Sem detalhamento', color:'#9CA3AF'};

/* =====================================================================
   ACOMPANHAMENTO DA OBRA — esquemas (Atividades, Avanço, Compromissos, Restrições)
   Base: Last Planner System (PPC, causas, lookahead de restrições) e curva S
   (avanço físico ponderado planejado × realizado, IDP).
   ===================================================================== */
const CAUSAS = ['Material / suprimentos', 'Mão de obra (falta ou deslocamento)', 'Equipamento', 'Projeto / mudança de escopo', 'Clima', 'Atividade predecessora não concluída', 'Planejamento / produtividade superestimada', 'Segurança / liberação', 'Cliente / contratual', 'Outros'];
const TIPOS_RESTR = ['Material', 'Mão de obra', 'Equipamento', 'Projeto / engenharia', 'Licença / autorização', 'Segurança', 'Atividade anterior', 'Contratual / financeiro', 'Clima', 'Outros'];
const STATUS_LPS = ['Planejado', 'Concluído', 'Não concluído'];
const STATUS_RESTR = ['Aberta', 'Removida'];
const TRACK = {
  ativ: {title: 'Atividades', sheet: 'Atividades', table: 'atividades', key: r => obraKey(r.obra) + '|' + norm(r.trecho || '') + '|' + norm(r.nome),
    tpl: ['Obra', 'Trecho', 'Serviço', 'Tipo', 'Und.', 'Qtd. Prevista', 'Qtd. Realizada', 'Qtd. Restante', 'Início Previsto', 'Término Previsto', 'Início Revisado', 'Término Revisado', 'Peso', 'Impedimentos'],
    fields: [
      {k: 'obra', l: 'Obra', req: 1, al: ['obra', 'obras', 'empreendimento']},
      {k: 'trecho', l: 'Trecho', al: ['trecho', 'local', 'frente', 'area', 'setor', 'disciplina', 'etapa']},
      {k: 'nome', l: 'Serviço', req: 1, al: ['servico', 'servicos', 'atividade', 'descricao', 'nome', 'tarefa', 'descricaodaatividade', 'descricaodoservico']},
      {k: 'tipo', l: 'Tipo', al: ['tipo', 'tipodeservico', 'classificacao']},
      {k: 'unidade', l: 'Und.', al: ['und', 'unidade', 'un', 'unid']},
      {k: 'qtd', l: 'Qtd. Prevista', t: 'num', al: ['qtdprevista', 'quantidadeprevista', 'qtdeprevista', 'quantidade', 'qtd', 'qtde', 'quant']},
      {k: 'realizada', l: 'Qtd. Realizada', t: 'num', al: ['qtdrealizada', 'quantidaderealizada', 'qtderealizada', 'realizada', 'realizado', 'executado', 'qtdexecutada', 'executadoacumulado']},
      {k: 'inicio', l: 'Início Previsto', t: 'date', al: ['inicioprevisto', 'inicio', 'datainicio', 'inicioplanejado', 'datadeinicio']},
      {k: 'fim', l: 'Término Previsto', t: 'date', al: ['terminoprevisto', 'termino', 'fim', 'datafim', 'fimprevisto', 'terminoplanejado', 'datadetermino']},
      {k: 'inicio_rev', l: 'Início Revisado', t: 'date', al: ['iniciorevisado', 'novoinicio', 'inicioreprogramado', 'reinicio']},
      {k: 'revisado', l: 'Término Revisado', t: 'date', al: ['terminorevisado', 'fimrevisado', 'novotermino', 'terminoreprogramado', 'reprogramado']},
      {k: 'peso', l: 'Peso', t: 'num', al: ['peso', 'pesopercentual', 'ponderacao', 'pesodaatividade', 'hh']},
      {k: 'impedimentos', l: 'Impedimentos', al: ['impedimentos', 'impedimento', 'restricao', 'restricoes', 'pendencia', 'pendencias']},
      {k: 'codigo', l: 'Código', hidden: 1, al: ['codigo', 'cod', 'item', 'eap', 'wbs', 'edt']},
    ]},
  avanco: {title: 'Avanço', sheet: 'Avanço', table: 'avanco', key: r => obraKey(r.obra) + '|' + norm(r.codigo) + '|' + r.semana,
    fields: [
      {k: 'obra', l: 'Obra', req: 1, al: ['obra', 'obras']},
      {k: 'codigo', l: 'Código', req: 1, al: ['codigo', 'cod', 'item', 'id', 'eap', 'wbs', 'edt']},
      {k: 'semana', l: 'Semana', req: 1, t: 'week', al: ['semana', 'semanaiso', 'sem', 'week']},
      {k: 'acumulado', l: 'Qtd executada acumulada', t: 'num', al: ['qtdexecutadaacumulada', 'executadoacumulado', 'quantidadeexecutadaacumulada', 'acumulado', 'qtdacumulada', 'executado', 'qtdexecutada']},
      {k: 'pct', l: '% acumulado', t: 'pct', al: ['acumuladopercentual', 'percentualacumulado', 'pctacumulado', 'avanco', 'avancoacumulado', 'percentual', 'avancofisico']},
      {k: 'obs', l: 'Observação', al: ['observacao', 'obs', 'comentario']},
      {k: 'inicio_rev', l: 'Início Revisado', t: 'date', hidden: 1, al: []},
      {k: 'revisado', l: 'Término Revisado', t: 'date', hidden: 1, al: []},
    ]},
  lps: {title: 'Compromissos semanais', sheet: 'Compromissos', table: 'compromissos', key: r => obraKey(r.obra) + '|' + r.semana + '|' + norm(r.codigo || '') + '|' + norm(r.atividade),
    fields: [
      {k: 'obra', l: 'Obra', req: 1, al: ['obra', 'obras']},
      {k: 'semana', l: 'Semana', req: 1, t: 'week', al: ['semana', 'semanaiso', 'sem', 'week']},
      {k: 'codigo', l: 'Código da atividade', al: ['codigo', 'cod', 'item', 'codigodaatividade']},
      {k: 'atividade', l: 'Compromisso / pacote de trabalho', req: 1, al: ['compromisso', 'atividade', 'tarefa', 'pacotedetrabalho', 'descricao', 'servico']},
      {k: 'responsavel', l: 'Responsável', al: ['responsavel', 'encarregado', 'equipe', 'lider']},
      {k: 'meta', l: 'Meta da semana', al: ['meta', 'metadasemana', 'quantidade', 'qtd']},
      {k: 'status', l: 'Status', t: 'opt', opts: STATUS_LPS, al: ['status', 'situacao', 'concluido', 'cumprido']},
      {k: 'causa', l: 'Causa do não cumprimento', t: 'opt', opts: CAUSAS, al: ['causa', 'causadonaocumprimento', 'motivo', 'causanaocumprimento']},
      {k: 'obs', l: 'Observação', al: ['observacao', 'obs', 'comentario']},
    ]},
  restr: {title: 'Restrições', sheet: 'Restrições', table: 'restricoes', key: r => obraKey(r.obra) + '|' + norm(r.descricao) + '|' + norm(r.codigo || r.atividade || ''),
    fields: [
      {k: 'obra', l: 'Obra', req: 1, al: ['obra', 'obras']},
      {k: 'codigo', l: 'Código da atividade', al: ['codigo', 'cod', 'item', 'codigodaatividade']},
      {k: 'atividade', l: 'Atividade', al: ['atividade', 'servico', 'tarefa']},
      {k: 'descricao', l: 'Restrição', req: 1, al: ['restricao', 'descricao', 'descricaodarestricao', 'impedimento', 'pendencia']},
      {k: 'tipo', l: 'Tipo', t: 'opt', opts: TIPOS_RESTR, al: ['tipo', 'categoria', 'tipoderestricao']},
      {k: 'responsavel', l: 'Responsável', al: ['responsavel', 'dono', 'encarregado']},
      {k: 'prazo', l: 'Data necessária', t: 'date', al: ['datanecessaria', 'prazo', 'necessidade', 'datalimite', 'datadenecessidade']},
      {k: 'status', l: 'Status', t: 'opt', opts: STATUS_RESTR, al: ['status', 'situacao']},
      {k: 'remocao', l: 'Data de remoção', t: 'date', al: ['dataderemocao', 'remocao', 'dataremocao', 'removidaem', 'datadesolucao']},
      {k: 'obs', l: 'Observação', al: ['observacao', 'obs', 'comentario']},
    ]},
};
const TRK = Object.keys(TRACK);
const TRK_ON = ['ativ']; // abas de acompanhamento ativas (Compromissos/Restrições ocultos)
const ALLK = ['labor', 'equip', 'manual', ...TRK];
const emptyData = () => Object.fromEntries(ALLK.map(k => [k, []]));
const catLabel = (kind, k) => (KINDS[kind].cats.find(c => c.k === k) || {}).l || k || '';
function normCat(kind, v){
  const s = norm(v).replace(/[^A-Z ]/g, ' ').replace(/\s+/g, ' ').trim();
  if (!s) return '';
  if (kind === 'labor') {
    if (/\bMOI\b|INDIRET/.test(s)) return 'MOI';
    if (/\bMOD\b|DIRET/.test(s)) return 'MOD';
    if (/TERC|\bTER\b|\bTERCEIRO/.test(s)) return 'TER';
    return null;
  }
  if (/AMAREL|\bLA\b/.test(s)) return 'LA';
  if (/BRANC|\bLB\b/.test(s)) return 'LB';
  if (/^PP$|\bP PORTE\b|PEQUENO|\bPEQ\b|PPORTE/.test(s)) return 'PP';
  return null;
}

/* =====================================================================
   CAMADA DE DADOS (adaptador local — substituível por API/banco futuramente)
   ===================================================================== */
const LS_KEY = 'cosampa_histogramas_v1', LS_LOGO = 'cosampa_histogramas_logo';
const LocalAdapter = {
  read(){ try { return JSON.parse(localStorage.getItem(LS_KEY) || 'null'); } catch (e) { return null; } },
  write(txt){ localStorage.setItem(LS_KEY, txt); }
};
const Repo = {
  adapter: LocalAdapter,
  data: emptyData(),
  lastBytes: 0,
  cols(kind){ const base = kind === 'manual' ? ['obra','semana','MOD','MOI','TER','moTotal','LA','LB','PP','eqTotal'] : TRACK[kind] ? TRACK[kind].fields.map(f => f.k) : KINDS[kind].fields.map(f => f.k); return ['id', ...base, 'createdAt', 'updatedAt']; },
  pack(kind){ const cols = this.cols(kind); return {cols, rows: this.data[kind].map(r => { const a = cols.map(c => r[c] == null || r[c] === '' ? null : r[c]); while (a.length && a[a.length - 1] === null) a.pop(); return a; })}; },
  unpack(t){ if (Array.isArray(t)) return t; if (!t || !Array.isArray(t.rows)) return []; return t.rows.map(a => { const o = {}; t.cols.forEach((c, i) => { if (a[i] != null) o[c] = a[i]; }); return o; }); },
  load(){ const d = this.adapter.read(); if (d && typeof d === 'object') this.data = Object.fromEntries(ALLK.map(k => [k, this.unpack(d[k])])); },
  mode: 'local',
  save(){
    if (this.mode === 'db') { DB.schedule(); return true; }
    try { const txt = JSON.stringify({version: 2, savedAt: new Date().toISOString(), ...Object.fromEntries(ALLK.map(k => [k, this.pack(k)]))}); this.adapter.write(txt); this.lastBytes = txt.length * 2; paintStorage(); return true; }
    catch (e) { const full = /quota|exceed/i.test(e.name + ' ' + e.message); toast(full ? 'O armazenamento deste navegador está cheio — os dados NÃO foram salvos. Baixe um backup e exclua semanas antigas.' : 'Não foi possível salvar no navegador: ' + e.message, 'err'); return false; }
  },
  /** Executa uma alteração e grava; se a gravação falhar, desfaz tudo (nada fica “meio salvo”). */
  tx(fn){ const snap = JSON.stringify(this.data); let out; try { out = fn(); } catch (e) { this.data = JSON.parse(snap); toast('Erro: ' + e.message, 'err'); return {ok: false}; } if (!this.save()) { this.data = JSON.parse(snap); return {ok: false}; } return {ok: true, out}; },
  replaceAll(d){ const old = this.data; this.data = Object.fromEntries(ALLK.map(k => [k, d[k] || []])); if (!this.save()) { this.data = old; return false; } return true; }
};

/* =====================================================================
   BANCO DE DADOS COMPARTILHADO (Supabase — REST + Auth, sem biblioteca extra)
   Só a URL do projeto e a chave pública "anon" ficam no navegador. A proteção
   dos dados é feita no banco (login + lista de e-mails autorizados + RLS).
   ===================================================================== */
// Opcional: preencha para que todos os usuários já abram o app conectado ao banco.
const DB_DEFAULT = {url: '', key: ''};
const LS_DB = 'cosampa_histogramas_db', LS_SESS = 'cosampa_histogramas_sessao';
const TABLES = {labor: 'mao_de_obra', equip: 'equipamentos', manual: 'totais_manuais', ...Object.fromEntries(Object.entries(TRACK).map(([k, v]) => [k, v.table]))};
const MANUAL_COLS = {MOD: 'mod', MOI: 'moi', TER: 'ter', moTotal: 'mo_total', LA: 'la', LB: 'lb', PP: 'pp', eqTotal: 'eq_total'};
const DB = {
  sess: null, queue: Promise.resolve(), pending: 0, synced: null, lastOk: null,
  cfg(){ try { const c = JSON.parse(localStorage.getItem(LS_DB) || 'null'); if (c && c.url && c.key) return c; } catch (e) {} return DB_DEFAULT.url && DB_DEFAULT.key ? DB_DEFAULT : null; },
  setCfg(c){ try { c ? localStorage.setItem(LS_DB, JSON.stringify(c)) : localStorage.removeItem(LS_DB); } catch (e) {} },
  loadSess(){ try { this.sess = JSON.parse(localStorage.getItem(LS_SESS) || 'null'); } catch (e) { this.sess = null; } },
  saveSess(x){ this.sess = x; try { x ? localStorage.setItem(LS_SESS, JSON.stringify(x)) : localStorage.removeItem(LS_SESS); } catch (e) {} },
  base(){ return this.cfg().url.trim().replace(/\/+$/, '').replace(/\/rest\/v1$/, ''); },
  async auth(body, grant){
    let r; try { r = await fetch(`${this.base()}/auth/v1/token?grant_type=${grant}`, {method: 'POST', headers: {apikey: this.cfg().key, 'Content-Type': 'application/json'}, body: JSON.stringify(body)}); }
    catch (e) { throw new Error('não foi possível acessar o banco (verifique a URL e a internet)'); }
    const j = await r.json().catch(() => ({}));
    if (!r.ok) { const m = j.error_description || j.msg || j.message || ('HTTP ' + r.status); throw new Error(/invalid login|invalid_grant|credentials/i.test(m) ? 'e-mail ou senha incorretos' : m); }
    this.saveSess({access: j.access_token, refresh: j.refresh_token, exp: Date.now() + (j.expires_in || 3600) * 1000 - 60000, email: (j.user && j.user.email) || (this.sess && this.sess.email) || body.email});
  },
  login(email, pw){ return this.auth({email, password: pw}, 'password'); },
  async ensure(){ if (!this.sess) throw new Error('sessão expirada — entre novamente'); if (Date.now() > this.sess.exp) await this.auth({refresh_token: this.sess.refresh}, 'refresh_token'); },
  logout(){ const x = this.sess; this.saveSess(null); if (x && this.cfg()) fetch(`${this.base()}/auth/v1/logout`, {method: 'POST', headers: {apikey: this.cfg().key, Authorization: 'Bearer ' + x.access}}).catch(() => {}); },
  async req(method, path, body, extra = {}, retry = true){
    await this.ensure();
    let r; try { r = await fetch(`${this.base()}/rest/v1/${path}`, {method, headers: {apikey: this.cfg().key, Authorization: 'Bearer ' + this.sess.access, 'Content-Type': 'application/json', ...extra}, body: body ? JSON.stringify(body) : undefined}); }
    catch (e) { throw new Error('sem conexão com o banco'); }
    if (r.status === 401 && retry) { await this.auth({refresh_token: this.sess.refresh}, 'refresh_token'); return this.req(method, path, body, extra, false); }
    if (!r.ok) { const j = await r.json().catch(() => ({})); const m = j.message || j.hint || ('HTTP ' + r.status); throw new Error(/row-level security/i.test(m) ? 'seu e-mail não está autorizado no banco (tabela usuarios_autorizados)' : /does not exist|schema cache/i.test(m) ? 'tabelas não encontradas — execute o script supabase_tabelas.sql no SQL Editor' : m); }
    return method === 'GET' ? r.json() : null;
  },
  toRow(kind, r){
    const o = {id: r.id};
    if (kind === 'manual') { o.obra = r.obra ?? null; o.semana = r.semana ?? null; for (const [k, c] of Object.entries(MANUAL_COLS)) o[c] = r[k] == null || r[k] === '' ? null : Number(r[k]); }
    else if (TRACK[kind]) TRACK[kind].fields.forEach(f => o[f.k] = r[f.k] == null || r[f.k] === '' ? null : (f.t === 'num' || f.t === 'pct') ? Number(r[f.k]) : String(r[f.k]));
    else KINDS[kind].fields.forEach(f => o[f.k] = r[f.k] == null || r[f.k] === '' ? null : String(r[f.k]));
    o.created_at = r.createdAt || null; o.updated_at = r.updatedAt || null; o.updated_by = (this.sess && this.sess.email) || null;
    return o;
  },
  fromRow(kind, row){
    const o = {id: row.id};
    if (kind === 'manual') { o.obra = row.obra; o.semana = row.semana; for (const [k, c] of Object.entries(MANUAL_COLS)) o[k] = row[c] ?? null; }
    else (TRACK[kind] || KINDS[kind]).fields.forEach(f => { if (row[f.k] != null) o[f.k] = (f.t === 'num' || f.t === 'pct') ? Number(row[f.k]) : row[f.k]; });
    if (row.created_at) o.createdAt = row.created_at; if (row.updated_at) o.updatedAt = row.updated_at;
    return o;
  },
  async loadAll(){
    const out = {};
    this.missing = [];
    for (const kind of ALLK) {
      const rows = [];
      try { for (let off = 0; ; off += 1000) { const part = await this.req('GET', `${TABLES[kind]}?select=*&order=id&limit=1000&offset=${off}`); rows.push(...part); if (part.length < 1000) break; } }
      catch (e) { if (TRACK[kind] && /tabelas não encontradas/.test(e.message)) { this.missing.push(TABLES[kind]); out[kind] = []; continue; } throw e; }
      out[kind] = rows.map(r => this.fromRow(kind, r));
    }
    if (this.missing.length) Log.add('Banco: tabelas do acompanhamento ainda não criadas (' + this.missing.join(', ') + '). Execute de novo o script supabase_tabelas.sql. Até lá, atividades, avanço, compromissos e restrições não são gravados no banco.', 'warn');
    return out;
  },
  snapshot(data){ const x = {}; for (const k of ALLK) x[k] = new Map((data[k] || []).map(r => [r.id, JSON.stringify(r)])); return x; },
  diff(){
    const d = {}; let n = 0;
    for (const k of ALLK) {
      const prev = this.synced[k] || new Map(); const cur = new Map((Repo.data[k] || []).map(r => [r.id, JSON.stringify(r)]));
      const up = []; for (const [id, js] of cur) if (prev.get(id) !== js) up.push(JSON.parse(js));
      const del = [...prev.keys()].filter(id => !cur.has(id));
      d[k] = {up, del}; n += up.length + del.length;
    }
    return {d, n};
  },
  async push(d){
    for (const k of ALLK) {
      if ((this.missing || []).includes(TABLES[k])) continue;
      const {up, del} = d[k];
      for (let i = 0; i < up.length; i += 500) await this.req('POST', TABLES[k], up.slice(i, i + 500).map(r => this.toRow(k, r)), {Prefer: 'resolution=merge-duplicates,return=minimal'});
      for (let i = 0; i < del.length; i += 100) await this.req('DELETE', `${TABLES[k]}?id=in.(${del.slice(i, i + 100).map(x => '"' + String(x).replace(/["(),]/g, '') + '"').join(',')})`);
    }
  },
  /** Grava no banco só o que mudou (envios em fila, um de cada vez). Se falhar, recarrega do banco. */
  schedule(){
    if (!this.synced) return;
    const {d, n} = this.diff(); if (!n) return;
    this.synced = this.snapshot(Repo.data); this.pending++; paintDbStatus();
    this.queue = this.queue.then(() => this.push(d)).then(() => { this.pending--; this.lastOk = new Date(); paintDbStatus(); })
      .catch(async e => { this.pending--; toast('Não foi possível salvar no banco: ' + e.message + '. Os dados foram recarregados do banco.', 'err'); await reloadFromDb(); });
  }
};
async function reloadFromDb(silent){
  try { const d = await DB.loadAll(); Repo.data = d; DB.synced = DB.snapshot(d); paintDbStatus(); refreshAll(); if (!silent) toast('Dados atualizados do banco.', 'ok'); return true; }
  catch (e) { if (/sessão|entre novamente|refresh|invalid/i.test(e.message)) { DB.saveSess(null); showLogin('Sua sessão expirou. Entre novamente.'); } else toast('Erro ao ler o banco: ' + e.message, 'err'); paintDbStatus(e.message); return false; }
}
function paintDbStatus(err){
  const el = document.getElementById('dbStatus'); if (!el) return; paintSideMini(err);
  const st = document.getElementById('storeUse'); const nt = document.querySelector('.notice span');
  if (Repo.mode === 'db') {
    const saving = DB.pending > 0;
    el.innerHTML = `<div style="display:flex;align-items:center;gap:6px;margin-bottom:4px"><span style="width:8px;height:8px;border-radius:50%;background:${err ? '#EF4444' : saving ? '#F7931E' : '#22C55E'}"></span><b>Banco compartilhado</b></div>Conectado como <b>${esc((DB.sess && DB.sess.email) || '')}</b><br>${err ? `<span style="color:#FCA5A5">${esc(err)}</span>` : saving ? 'Salvando no banco…' : 'Tudo salvo no banco'} · <a href="#" data-db="reload">recarregar</a> · <a href="#" data-db="logout">sair</a>`;
    if (st) st.style.display = 'none'; const sl = document.getElementById('sideLocal'); if (sl) sl.style.display = 'none';
    if (nt) nt.innerHTML = `Dados salvos no <b>banco compartilhado</b> (Supabase) — todos os usuários autorizados veem as mesmas informações. A planilha conectada é só lida.`;
  } else {
    el.innerHTML = `<b>Modo local</b> — dados só neste navegador. <a href="#" data-db="cfg">Conectar banco</a>`;
    if (st) st.style.display = ''; const sl = document.getElementById('sideLocal'); if (sl) sl.style.display = '';
    if (nt) nt.innerHTML = `Os lançamentos ficam salvos <b>somente neste navegador/dispositivo</b> (armazenamento local). A planilha conectada é só lida (nada é gravado nela). Use <b>Baixar backup</b> regularmente.`;
  }
}
function showLogin(msg){
  if (document.getElementById('loginBg')) return;
  const bg = document.createElement('div'); bg.className = 'modal-bg'; bg.id = 'loginBg'; bg.style.alignItems = 'center';
  bg.innerHTML = `<div class="modal" style="max-width:400px"><div class="modal-h" style="justify-content:center;padding:18px"><img src="${getLogo() || LOGO_COSAMPA}" alt="Grupo COSAMPA" style="height:30px"></div>
    <div class="modal-b"><h3 style="margin:0 0 4px">Controle de Histogramas</h3><p class="muted" style="margin:0 0 14px;font-size:13px">Entre com o usuário cadastrado no banco de dados.</p>
      ${msg ? `<div class="alert warn" style="font-size:12px">${esc(msg)}</div>` : ''}
      <form data-login><div class="fld"><span>E-mail</span><input type="email" name="email" autocomplete="username" required value="${esc((DB.sess && DB.sess.email) || '')}"></div>
      <div class="fld" style="margin-top:10px"><span>Senha</span><input type="password" name="pw" autocomplete="current-password" required></div>
      <div data-lerr></div>
      <button class="btn primary" style="width:100%;justify-content:center;margin-top:14px;padding:10px">Entrar</button></form>
      <div style="display:flex;justify-content:space-between;margin-top:14px;font-size:12px"><a href="#" data-db="cfg" style="color:var(--or2)">Configurar banco</a><a href="#" data-db="local" style="color:var(--mut)">Usar sem banco (só neste navegador)</a></div>
    </div></div>`;
  document.body.appendChild(bg);
  $('[data-login]', bg).onsubmit = async ev => {
    ev.preventDefault(); const f = ev.target; const b = $('button', f); b.disabled = true; b.textContent = 'Entrando…'; $('[data-lerr]', bg).innerHTML = '';
    try { await DB.login(f.email.value.trim(), f.pw.value); bg.remove(); await enterDbMode(); }
    catch (e) { $('[data-lerr]', bg).innerHTML = `<div class="alert err" style="margin-top:10px;font-size:12px">${esc(e.message)}</div>`; b.disabled = false; b.textContent = 'Entrar'; }
  };
  setTimeout(() => { const i = $(DB.sess && DB.sess.email ? 'input[name=pw]' : 'input[name=email]', bg); i && i.focus(); }, 50);
}
async function enterDbMode(){
  Repo.mode = 'db'; paintDbStatus();
  let d; try { d = await DB.loadAll(); }
  catch (e) { paintDbStatus(e.message); toast('Erro ao ler o banco: ' + e.message, 'err'); if (/sessão|entre novamente/i.test(e.message)) { DB.saveSess(null); showLogin(e.message); } return false; }
  Repo.data = d; DB.synced = DB.snapshot(d); paintDbStatus(); refreshAll();
  // Migração: banco vazio e dados antigos neste navegador
  const loc = LocalAdapter.read();
  if (!d.labor.length && !d.equip.length && !d.manual.length && loc) {
    const L = Object.fromEntries(ALLK.map(k => [k, Repo.unpack(loc[k])]));
    const n = ALLK.reduce((a, k) => a + L[k].length, 0);
    if (n && await confirmBox('Enviar dados para o banco', `<p>O banco está vazio e este navegador tem <b>${L.labor.length}</b> cadastro(s) de mão de obra, <b>${L.equip.length}</b> de equipamentos e <b>${L.manual.length}</b> lançamento(s) manuais.</p><p>Deseja enviá-los para o banco compartilhado?</p>`, 'Enviar para o banco')) {
      Repo.data = L; DB.schedule(); refreshAll(); toast('Enviando dados para o banco…');
    }
  }
  Sync.auto();
  return true;
}
function openDbConfig(){
  const c = DB.cfg() || {url: '', key: ''};
  const md = openModal({title: 'Banco de dados compartilhado (Supabase)', body: `
    <div class="alert info" style="font-size:12px">No Supabase: <b>Project Settings → API</b> (ou <b>Data API / API Keys</b>). Copie a <b>Project URL</b> e a <b>Publishable key</b> (<i>sb_publishable_…</i>; em projetos antigos, a <i>anon public</i>). <b>Nunca</b> use a <i>Secret key</i> (<i>sb_secret_…</i>), a <i>service_role</i> nem a senha do banco.</div>
    <div class="fld"><span>Project URL</span><input data-c="url" placeholder="https://xxxx.supabase.co" value="${esc(c.url)}"></div>
    <div class="fld" style="margin-top:10px"><span>Publishable key (ou anon public)</span><input data-c="key" placeholder="sb_publishable_…" value="${esc(c.key)}"></div>
    <p class="muted" style="font-size:12px">Antes, execute o script <b>supabase_tabelas.sql</b> no SQL Editor do Supabase e crie os usuários em Authentication → Users.</p>`,
    foot: `${Repo.mode === 'db' || DB.cfg() ? '<button class="btn danger" data-off style="margin-right:auto">Desconectar banco</button>' : ''}${Repo.mode === 'db' && DB.sess ? `<button class="btn" data-db="logout">Sair (${esc(DB.sess.email || '')})</button>` : ''}<button class="btn" data-x>Cancelar</button><button class="btn primary" data-ok>Salvar e entrar</button>`});
  $('[data-x]', md.el).onclick = md.close;
  const off = $('[data-off]', md.el); if (off) off.onclick = () => { DB.logout(); DB.setCfg(null); location.reload(); };
  $('[data-ok]', md.el).onclick = () => {
    const url = $('[data-c=url]', md.el).value.trim(), key = $('[data-c=key]', md.el).value.trim();
    if (!/^https:\/\/[\w.-]+/.test(url) || key.length < 20) return toast('Informe a Project URL (https://…) e a Publishable key.', 'err');
    if (/^sb_secret_/i.test(key) || /service_role/.test(atobSafe(key))) return toast('Esta é uma chave SECRETA — não use no app. Copie a “Publishable key” (sb_publishable_…) ou a “anon public”.', 'err');
    DB.setCfg({url, key}); DB.saveSess(null); md.close(); const lb = document.getElementById('loginBg'); if (lb) lb.remove(); showLogin();
  };
}
function atobSafe(k){ try { return atob(k.split('.')[1].replace(/-/g, '+').replace(/_/g, '/')); } catch (e) { return ''; } }
document.addEventListener('click', async e => {
  const a = e.target.closest('[data-db]'); if (!a) return; e.preventDefault();
  const act = a.dataset.db;
  if (act === 'cfg') openDbConfig();
  else if (act === 'reload') reloadFromDb();
  else if (act === 'logout') { const mb = a.closest('.modal-bg'); if (mb) mb.remove(); if (DB.pending) return toast('Aguarde terminar de salvar no banco.', 'err'); DB.logout(); Repo.data = emptyData(); DB.synced = null; refreshAll(); showLogin('Você saiu.'); }
  else if (act === 'local') { const lb = document.getElementById('loginBg'); if (lb) lb.remove(); Repo.mode = 'local'; Repo.load(); paintDbStatus(); refreshAll(); toast('Modo local: dados só neste navegador.'); Sync.auto(); }
});

/* =====================================================================
   REGRAS DE NEGÓCIO
   ===================================================================== */
const obraKey = o => norm(o);
function canonObra(name){
  const n = clean(name); const k = obraKey(n); if (!k) return n;
  for (const kind of ['labor','equip']) { const r = Repo.data[kind].find(x => obraKey(x.obra) === k); if (r) return r.obra; }
  const m = Repo.data.manual.find(x => obraKey(x.obra) === k); return m ? m.obra : n;
}
const isZeroMat = m => { const s = String(m ?? '').trim(); return !s || /^0+$/.test(s); };
function personKey(r){
  if (!isZeroMat(r.matricula)) return 'MAT:' + norm(r.matricula);
  const c = digits(r.cpf); if (c) return 'CPF:' + c;
  return 'NOME:' + norm(r.nome);
}
const vincKey = (kind, r) => obraKey(r.obra) + '|' + (kind === 'labor' ? personKey(r) : 'PRE:' + norm(r.prefixo));
const recKey = (kind, r) => r.semana + '|' + vincKey(kind, r);
function describeKey(kind, r){
  if (kind === 'equip') return `${r.semana} · ${r.obra} · Prefixo ${r.prefixo}`;
  if (!isZeroMat(r.matricula)) return `${r.semana} · ${r.obra} · Mat. ${r.matricula} — ${r.nome}`;
  return `${r.semana} · ${r.obra} · PJ ${digits(r.cpf) ? 'CPF ' + r.cpf : ''} — ${r.nome}`;
}
function activeAtClose(kind, r, weekEnd){
  const K = KINDS[kind]; const s = r[K.start], e = r[K.end];
  if (s && s > weekEnd) return false;      // ainda não iniciou no fechamento
  if (e && e < weekEnd) return false;      // encerrou antes do fechamento (datas inclusivas)
  return true;
}
/** Lançamento sem detalhamento → grupo (mão de obra ou equipamentos) */
function manualGroup(m, kind){
  const ks = KINDS[kind].cats.map(c => c.k); const tk = kind === 'labor' ? 'moTotal' : 'eqTotal';
  if (ks.some(k => !isEmptyVal(m[k]))) { const o = {unknown:false}; ks.forEach(k => o[k] = Number(m[k] || 0)); o.total = ks.reduce((s, k) => s + o[k], 0); return o; }
  if (!isEmptyVal(m[tk])) { const o = {unknown:true, total:Number(m[tk])}; ks.forEach(k => o[k] = null); return o; }
  return null;
}
function manualMap(kind){
  const mp = new Map();
  for (const m of Repo.data.manual) { const g = manualGroup(m, kind); if (g) mp.set(m.semana + '|' + obraKey(m.obra), {...g, src:m}); }
  return mp;
}
function blankFilter(){ return {obras:[], weeks:[], from:'', to:'', cat:'', sit:'', q:'', basis:'foto'}; }
function passCommon(obra, week, f){
  if (f.obras && f.obras.length && !f.obras.includes(obraKey(obra))) return false;
  if (f.weeks && f.weeks.length && !f.weeks.includes(week)) return false;
  if (f.from || f.to) { const r = weekRange(week); if (!r) return false; if (f.from && r.end < f.from) return false; if (f.to && r.start > f.to) return false; }
  return true;
}
function displayVal(kind, r, fd){ const v = r[fd.k]; if (fd.t === 'date') return fmtBR(v); if (fd.t === 'cat') return catLabel(kind, v); return v ?? ''; }
/** Registros filtrados. Os grupos substituídos por totais manuais ficam ocultos. */
function filterRecords(kind, f, opt = {}){
  const K = KINDS[kind]; const mm = opt.mm || manualMap(kind);
  const terms = f.q ? norm(f.q).split(' ').filter(Boolean) : [];
  const out = [], hidden = [];
  for (const r of Repo.data[kind]) {
    if (!passCommon(r.obra, r.semana, f)) continue;
    if (f.cat && r[K.catField] !== f.cat) continue;
    if (kind === 'labor' && f.sit) { const s = clean(r.situacao); if (f.sit === '__none__' ? s !== '' : norm(s) !== f.sit) continue; }
    if (terms.length) { const hay = norm(K.fields.map(fd => displayVal(kind, r, fd)).join(' ')); if (!terms.every(t => hay.includes(t))) continue; }
    out.push(r);
  }
  return {rows: out, hidden};
}
/** Efetivo no fechamento por semana × obra (sem somar dias; sem duplicar pessoa). */
function computeCells(kind, f){
  const K = KINDS[kind]; const ks = K.cats.map(c => c.k); const mm = manualMap(kind);
  const {rows, hidden} = filterRecords(kind, f, {mm});
  const cells = new Map(), labels = new Map(), seen = new Set();
  const cell = (w, o, label) => { const key = w + '||' + o; if (!labels.has(o)) labels.set(o, label); let c = cells.get(key); if (!c) { c = {w, o, total:0, SD:0, manual:false, unknown:false}; ks.forEach(k => c[k] = 0); cells.set(key, c); } return c; };
  for (const r of rows) {
    const wr = weekRange(r.semana); if (!wr) continue;
    const c = cell(r.semana, obraKey(r.obra), r.obra);
    if (!activeAtClose(kind, r, wr.end)) continue;
    const vk = r.semana + '|' + vincKey(kind, r); if (seen.has(vk)) continue; seen.add(vk);
    const cat = r[K.catField]; if (ks.includes(cat)) c[cat]++; c.total++;
  }
  if (f.basis === 'datas') estimateByDates(kind, f, mm, cells, cell);
  // Lançamento manual COMPLEMENTA os cadastros: só vale nas categorias sem nenhum cadastro na obra/semana.
  const mnotes = []; const tk = kind === 'labor' ? 'moTotal' : 'eqTotal';
  for (const [, g] of mm) {
    const m = g.src; if (!passCommon(m.obra, m.semana, f)) continue;
    const c = cell(m.semana, obraKey(m.obra), m.obra); let added = false;
    if (ks.some(k => !isEmptyVal(m[k]))) {
      for (const k of ks) { if (isEmptyVal(m[k]) || (f.cat && f.cat !== k)) continue; if (!c[k]) { c[k] = Number(m[k]); added = true; } }
    } else if (!isEmptyVal(m[tk]) && !c.total) {
      if (f.cat) mnotes.push(m); else { c.SD = Number(m[tk]); ks.forEach(k => c[k] = null); c.unknown = true; added = true; }
    }
    if (added) { c.manual = true; c.total = ks.reduce((sm, k) => sm + (c[k] || 0), 0) + (c.SD || 0); }
  }
  return {cells, labels, rows, hidden, mm, mnotes};
}
/** Semanas reconstruídas a partir das datas de admissão/demissão (ou mobilização/desmobilização).
 *  Só preenche semanas SEM fotografia da obra, até a última fotografia existente (nunca projeta o futuro).
 *  Para cada vínculo usa a fotografia mais próxima posterior à semana (a que tem as datas mais atualizadas). */
function weeksBetween(a, b, cap = 156){ const out = []; let w = a, g = 0; while (w && w <= b && g++ < 2000) { out.push(w); w = nextWeek(w); } return out.slice(-cap); }
function obraTimeline(kind, f){
  const K = KINDS[kind];
  const all = filterRecords(kind, {...f, weeks: [], from: '', to: ''}).rows;
  const byObra = new Map();
  for (const r of all) {
    const o = obraKey(r.obra); let ob = byObra.get(o);
    if (!ob) { ob = {o, label: r.obra, v: new Map(), snaps: new Set(), first: null, last: null}; byObra.set(o, ob); }
    ob.snaps.add(r.semana);
    if (!ob.last || r.semana > ob.last) ob.last = r.semana;
    const sw = r[K.start] ? isoWeek(r[K.start]) : r.semana; if (!ob.first || sw < ob.first) ob.first = sw;
    const vk = vincKey(kind, r); if (!ob.v.has(vk)) ob.v.set(vk, []); ob.v.get(vk).push(r);
  }
  byObra.forEach(ob => ob.v.forEach(list => list.sort((a, b) => a.semana.localeCompare(b.semana))));
  return byObra;
}
function estimateWeeksList(f, byObra){
  if (f.weeks && f.weeks.length) return [...f.weeks].filter(w => passCommon('', w, {from: f.from, to: f.to})).sort();
  let a = null, b = null; byObra.forEach(ob => { if (!a || ob.first < a) a = ob.first; if (!b || ob.last > b) b = ob.last; });
  if (f.from) a = isoWeek(f.from); if (f.to) b = isoWeek(f.to) < b || !b ? isoWeek(f.to) : b;
  return a && b ? weeksBetween(a, b, f.from ? 156 : 26) : [];
}
function estimateByDates(kind, f, mm, cells, cell){
  const K = KINDS[kind]; const ks = K.cats.map(c => c.k);
  const byObra = obraTimeline(kind, f);
  for (const w of estimateWeeksList(f, byObra)) {
    const end = weekRange(w).end;
    for (const ob of byObra.values()) {
      if (!passCommon(ob.label, w, f)) continue;
      if (ob.snaps.has(w)) continue;                                 // fotografia real tem prioridade
      if (w > ob.last || w < ob.first) continue;                    // não projeta além da última fotografia
      const c = cell(w, ob.o, ob.label); c.est = true;
      for (const list of ob.v.values()) {
        const r = list.find(x => x.semana >= w) || list[list.length - 1];
        if (!r[K.start] && w < list[0].semana) continue;            // sem data inicial: não presume presença antes da 1ª fotografia
        if (!activeAtClose(kind, r, end)) continue;
        const cat = r[K.catField]; if (ks.includes(cat)) c[cat]++; c.total++;
      }
    }
  }
}
/** Há semanas que poderiam ser reconstruídas pelas datas? (para sugerir ao usuário) */
function missingWeeksHint(kind, f){
  const byObra = obraTimeline(kind, f); const out = [];
  byObra.forEach(ob => { const span = weeksBetween(ob.first, ob.last, 26); const miss = span.filter(w => !ob.snaps.has(w)).length; if (miss) out.push({label: ob.label, miss, first: ob.first, snaps: ob.snaps.size}); });
  return out;
}
function sortedWeeks(cells){ return [...new Set([...cells.values()].map(c => c.w))].sort(); }
function sortedObras(cells, labels){ return [...new Set([...cells.values()].map(c => c.o))].sort((a, b) => String(labels.get(a)).localeCompare(String(labels.get(b)), 'pt-BR')); }
function aggCells(kind, list){
  const ks = KINDS[kind].cats.map(c => c.k); const a = {total:0, SD:0, n:list.length, allUnknown: list.length > 0 && list.every(c => c.unknown), anyManual: list.some(c => c.manual), anyEst: list.some(c => c.est)};
  ks.forEach(k => a[k] = 0);
  for (const c of list) { a.total += c.total; a.SD += c.SD || 0; ks.forEach(k => a[k] += c[k] || 0); }
  if (a.allUnknown) ks.forEach(k => a[k] = null);
  return a;
}

/* =====================================================================
   ESTADO DA INTERFACE
   ===================================================================== */
const UI = {
  tab: 'labor',
  f: {labor: blankFilter(), equip: blankFilter()},
  mode: {labor:'auto', equip:'auto'},
  page: {labor:1, equip:1},
  ms: {},
  resumo: {obras:[], week:'__last__', sit:'', sub:'tabela', jr:'labor'},
};
const charts = {};
function killChart(id){ if (charts[id]) { try { charts[id].destroy(); } catch (e) {} delete charts[id]; } }
const journeys = {};

/* =====================================================================
   COMPONENTE: COMBOBOX DE SELEÇÃO MÚLTIPLA
   ===================================================================== */
function MultiSelect(el, cfg){
  el.classList.add('ms');
  el.innerHTML = `<button type="button" class="ms-btn"><span class="ms-txt"></span><span class="cnt" hidden></span><span>▾</span></button>
    <div class="ms-pop" hidden><input type="search" class="ms-s" placeholder="Pesquisar…"><div class="ms-quick"></div><div class="ms-list"></div></div>`;
  const btn = $('.ms-btn', el), pop = $('.ms-pop', el), list = $('.ms-list', el), search = $('.ms-s', el), quick = $('.ms-quick', el);
  let value = [...(cfg.value || [])];
  const api = {
    get value(){ return [...value]; },
    set(v, silent){ value = [...v]; paint(); if (!silent) cfg.onChange(api.value); },
    refresh(){ const opts = cfg.options(); value = value.filter(v => opts.some(o => o.v === v)); paint(); }
  };
  function paint(){
    const opts = cfg.options();
    const txt = !value.length ? cfg.all : value.length === 1 ? (opts.find(o => o.v === value[0]) || {}).l || value[0] : `${value.length} selecionadas`;
    $('.ms-txt', el).textContent = txt;
    const cnt = $('.cnt', el); cnt.hidden = value.length < 2; cnt.textContent = value.length;
    const q = norm(search.value);
    list.innerHTML = opts.filter(o => !q || norm(o.l + ' ' + (o.sub || '')).includes(q)).map(o =>
      `<label><input type="checkbox" value="${esc(o.v)}" ${value.includes(o.v) ? 'checked' : ''}><span>${esc(o.l)} ${o.sub ? `<small>${esc(o.sub)}</small>` : ''}</span></label>`).join('') || '<div class="muted" style="padding:6px">Nenhuma opção</div>';
  }
  quick.innerHTML = (cfg.quick || []).map((q, i) => `<button type="button" data-i="${i}">${esc(q.l)}</button>`).join('');
  quick.addEventListener('click', e => { const b = e.target.closest('button'); if (!b) return; api.set(cfg.quick[+b.dataset.i].fn(cfg.options().map(o => o.v))); });
  btn.addEventListener('click', e => { e.stopPropagation(); const open = pop.hidden; $$('.ms-pop').forEach(p => p.hidden = true); pop.hidden = !open; if (open) { search.value = ''; paint(); pop.style.left = '0'; pop.style.right = 'auto'; const rc = pop.getBoundingClientRect(); if (rc.right > window.innerWidth - 8) { pop.style.left = 'auto'; pop.style.right = '0'; } search.focus(); } });
  pop.addEventListener('click', e => e.stopPropagation());
  search.addEventListener('input', paint);
  list.addEventListener('change', e => { const v = e.target.value; if (e.target.checked) { if (!value.includes(v)) value.push(v); } else value = value.filter(x => x !== v); paint(); cfg.onChange(api.value); });
  paint();
  return api;
}
document.addEventListener('click', () => $$('.ms-pop').forEach(p => p.hidden = true));

/* =====================================================================
   GRÁFICOS — tema moderno, plugins próprios
   ===================================================================== */
function setupChartDefaults(){
  Chart.defaults.font.family = 'Arial, Helvetica, sans-serif';
  Chart.defaults.font.size = 12;
  Chart.defaults.color = '#4B5563';
  Chart.defaults.plugins.tooltip.backgroundColor = 'rgba(17,17,20,.94)';
  Chart.defaults.plugins.tooltip.titleColor = '#F7931E';
  Chart.defaults.plugins.tooltip.padding = 10;
  Chart.defaults.plugins.tooltip.cornerRadius = 10;
  Chart.defaults.plugins.tooltip.boxPadding = 5;
  Chart.defaults.plugins.tooltip.usePointStyle = true;
  Chart.register(valueLabelPlugin, journeyPlugin, centerTextPlugin);
}
function shade(hex, amt){ const n = parseInt(hex.slice(1), 16); let r = (n >> 16) + amt, g = ((n >> 8) & 255) + amt, b = (n & 255) + amt; r = Math.max(0, Math.min(255, r)); g = Math.max(0, Math.min(255, g)); b = Math.max(0, Math.min(255, b)); return '#' + ((1 << 24) + (r << 16) + (g << 8) + b).toString(16).slice(1); }
function rgba(hex, a){ const n = parseInt(hex.slice(1), 16); return `rgba(${n >> 16},${(n >> 8) & 255},${n & 255},${a})`; }
function barGradient(color){
  return ctx => { const ch = ctx.chart, a = ch.chartArea; if (!a) return color; const g = ch.ctx.createLinearGradient(0, a.bottom, 0, a.top); g.addColorStop(0, shade(color, -18)); g.addColorStop(1, shade(color, 38)); return g; };
}
let _stripe;
function stripePattern(){
  if (_stripe) return _stripe; const c = document.createElement('canvas'); c.width = c.height = 10; const x = c.getContext('2d');
  x.fillStyle = '#D1D5DB'; x.fillRect(0, 0, 10, 10); x.strokeStyle = '#9CA3AF'; x.lineWidth = 3; x.beginPath(); x.moveTo(-2, 12); x.lineTo(12, -2); x.stroke();
  _stripe = x.createPattern(c, 'repeat'); return _stripe;
}
/** Números acima das barras (acompanham a animação) e "—" para categorias desconhecidas */
const valueLabelPlugin = {
  id: 'valueLabels',
  afterDatasetsDraw(chart, args, opts){
    if (!opts || !opts.enabled) return;
    const {ctx} = chart;
    ctx.save(); ctx.textAlign = 'center'; ctx.textBaseline = 'bottom';
    chart.data.datasets.forEach((ds, i) => {
      if (ds.noLabel) return;
      if (!chart.isDatasetVisible(i)) return;
      const meta = chart.getDatasetMeta(i);
      meta.data.forEach((el, j) => {
        const v = ds.data[j]; const unk = ds.unknown && ds.unknown[j];
        if (v == null && !unk) return;
        const {x, y, base} = el.getProps(['x', 'y', 'base'], false);
        if (unk) { ctx.font = 'bold 12px Arial'; ctx.fillStyle = '#9CA3AF'; ctx.fillText('—', x, (base ?? y) - 3); return; }
        const txt = fmtN(v);
        ctx.font = `bold ${opts.size || 11}px Arial`; ctx.fillStyle = opts.color || '#111827';
        ctx.fillText(txt, x, y - 4);
      });
    });
    ctx.restore();
  }
};
const centerTextPlugin = {
  id: 'centerText',
  afterDraw(chart, args, opts){
    if (!opts || opts.text == null) return; const {ctx, chartArea: a} = chart; if (!a) return;
    const cx = (a.left + a.right) / 2, cy = (a.top + a.bottom) / 2;
    ctx.save(); ctx.textAlign = 'center'; ctx.textBaseline = 'middle';
    ctx.font = 'bold 22px Arial'; ctx.fillStyle = '#111827'; ctx.fillText(opts.text, cx, cy - 6);
    ctx.font = '10px Arial'; ctx.fillStyle = '#6B7280'; ctx.fillText(opts.sub || '', cx, cy + 13);
    ctx.restore();
  }
};
function rrect(ctx, x, y, w, h, r){ ctx.beginPath(); ctx.moveTo(x + r, y); ctx.arcTo(x + w, y, x + w, y + h, r); ctx.arcTo(x + w, y + h, x, y + h, r); ctx.arcTo(x, y + h, x, y, r); ctx.arcTo(x, y, x + w, y, r); ctx.closePath(); }
const easeOutBack = t => { const c1 = 1.70158, c3 = c1 + 1; return 1 + c3 * Math.pow(t - 1, 3) + c1 * Math.pow(t - 1, 2); };
const MILE = {start:{l:'INÍCIO', c:'#16A34A'}, peak:{l:'PICO', c:'#DC2626'}, end:{l:'FINAL', c:'#2563EB'}};
/** Percurso: revela a linha da esquerda p/ direita, cursor e marcos Início/Pico/Final */
const journeyPlugin = {
  id: 'journey',
  beforeDatasetsDraw(chart){
    const j = chart.$j; if (!j) return;
    const {ctx, chartArea: a, scales: {x}} = chart;
    ctx.save();
    if (j.reveal < j.n - 1 - 1e-6) { const px = x.getPixelForValue(Math.max(0, j.reveal)); ctx.beginPath(); ctx.rect(a.left - 12, a.top - 30, px - a.left + 12, a.bottom - a.top + 60); ctx.clip(); }
  },
  afterDatasetsDraw(chart){
    const j = chart.$j; if (!j) return;
    const {ctx, chartArea: a, scales: {x, y}} = chart;
    ctx.restore();
    const data = chart.data.datasets[0].data;
    if (!chart.isDatasetVisible(0)) return;
    // cursor
    if (j.cursor) {
      const px = x.getPixelForValue(j.reveal);
      ctx.save(); ctx.strokeStyle = 'rgba(17,17,20,.35)'; ctx.setLineDash([4, 4]); ctx.lineWidth = 1; ctx.beginPath(); ctx.moveTo(px, a.top); ctx.lineTo(px, a.bottom); ctx.stroke(); ctx.restore();
      const i0 = Math.floor(j.reveal), i1 = Math.min(j.n - 1, i0 + 1), t = j.reveal - i0;
      const v0 = data[i0], v1 = data[i1];
      if (v0 != null && (v1 != null || t < 1e-6)) {
        const v = v1 == null ? v0 : v0 + (v1 - v0) * t; const py = y.getPixelForValue(v);
        ctx.save(); const g = ctx.createRadialGradient(px, py, 0, px, py, 18); g.addColorStop(0, 'rgba(247,147,30,.55)'); g.addColorStop(1, 'rgba(247,147,30,0)');
        ctx.fillStyle = g; ctx.beginPath(); ctx.arc(px, py, 18, 0, Math.PI * 2); ctx.fill();
        ctx.fillStyle = '#F7931E'; ctx.strokeStyle = '#fff'; ctx.lineWidth = 3; ctx.beginPath(); ctx.arc(px, py, 7, 0, Math.PI * 2); ctx.fill(); ctx.stroke(); ctx.restore();
      }
    }
    // marcos
    const meta = chart.getDatasetMeta(0);
    for (const m of j.miles) {
      if (j.reveal < m.idx - 1e-6) continue;
      const p = j.reveal >= j.n - 1 - 1e-6 && !j.cursor ? 1 : Math.min(1, (j.reveal - m.idx) / 0.35 + 0.05);
      const s = Math.max(0, easeOutBack(p)); const el = meta.data[m.idx]; if (!el) continue;
      const px = el.x, py = el.y; const col = MILE[m.kinds[m.kinds.length - 1]].c;
      ctx.save();
      ctx.fillStyle = rgba(col, .16); ctx.beginPath(); ctx.arc(px, py, 13 * s, 0, Math.PI * 2); ctx.fill();
      ctx.fillStyle = col; ctx.strokeStyle = '#fff'; ctx.lineWidth = 2.5; ctx.beginPath(); ctx.arc(px, py, 6 * s, 0, Math.PI * 2); ctx.fill(); ctx.stroke();
      const label = m.kinds.map(k => MILE[k].l).join(' · ') + '  ' + fmtN(m.value);
      ctx.font = 'bold 11px Arial'; const w = ctx.measureText(label).width + 16, h = 22;
      let bx = Math.min(Math.max(px - w / 2, a.left), a.right - w), by = py - 20 - h;
      if (by < a.top - 6) by = py + 16;
      ctx.globalAlpha = Math.min(1, p * 1.6);
      ctx.translate(px, py); ctx.scale(s, s); ctx.translate(-px, -py);
      ctx.shadowColor = 'rgba(0,0,0,.18)'; ctx.shadowBlur = 8; ctx.shadowOffsetY = 2;
      ctx.fillStyle = col; rrect(ctx, bx, by, w, h, 11); ctx.fill();
      ctx.shadowColor = 'transparent'; ctx.fillStyle = '#fff'; ctx.textAlign = 'left'; ctx.textBaseline = 'middle'; ctx.fillText(label, bx + 8, by + h / 2 + 0.5);
      ctx.restore();
    }
  }
};

/** Histograma moderno com barras agrupadas */
function histogramModel(kind, f, mode){
  const K = KINDS[kind]; const {cells, labels, rows, hidden, mm, mnotes} = computeCells(kind, f);
  const weeks = sortedWeeks(cells), obras = sortedObras(cells, labels);
  let m = mode;
  if (m === 'auto' || m === 'semana') m = weeks.length <= 1 ? 'obra' : 'semanaobra';
  const groups = []; let note = '';
  if (m === 'obra') {
    const lat = latestPerObra({weeks, obras, cells});
    if (weeks.length > 1) note = 'Por obra mostra a semana mais recente de CADA obra dentro do filtro — efetivos de semanas diferentes da mesma obra não são somados.';
    for (const x of lat) groups.push({label: [...wrapText(labels.get(x.o)), weekShort(x.w)], title: `${labels.get(x.o)} — ${weekLabel(x.w, true)}`, a: aggCells(kind, [x.c])});
  } else if (m === 'semana') {
    for (const w of weeks) { const list = obras.map(o => cells.get(w + '||' + o)).filter(Boolean); groups.push({label: [weekShort(w), `${list.length} obra${list.length > 1 ? 's' : ''}`], title: weekLabel(w, true), a: aggCells(kind, list)}); }
  } else {
    for (const w of weeks) for (const o of obras) { const c = cells.get(w + '||' + o); if (c) groups.push({label: (obras.length === 1 ? [weekShort(w), fmtBR(weekRange(w).end).slice(0, 5)] : [weekShort(w), ...wrapText(labels.get(o), 14, 2)]).concat(c.est ? ['(est.)'] : []), title: `${weekLabel(w, true)} — ${labels.get(o)}`, a: aggCells(kind, [c])}); }
  }
  const series = K.cats.filter(c => !f.cat || f.cat === c.k);
  const datasets = series.map(c => ({
    key: c.k, label: c.l, color: c.color,
    data: groups.map(g => g.a[c.k]), unknown: groups.map(g => g.a[c.k] == null && g.a.SD > 0)
  }));
  if (groups.some(g => g.a.SD > 0)) datasets.push({key:'SD', label:SD.l, color:SD.color, stripe:true, data: groups.map(g => g.a.SD > 0 ? g.a.SD : null), unknown: groups.map(() => false)});
  return {mode: m, groups, datasets, weeks, obras, cells, labels, rows, hidden, mm, mnotes, note};
}
function barConfig(model, opts = {}){
  const anim = opts.animate === false || reduceMotion ? false : {
    duration: 900, easing: 'easeOutQuart',
    delay: ctx => ctx.type === 'data' && ctx.mode === 'default' ? Math.min(700, ctx.dataIndex * 45) + ctx.datasetIndex * 110 : 0
  };
  return {
    type: 'bar',
    data: {
      labels: model.groups.map(g => g.label),
      datasets: model.datasets.map(d => ({
        label: d.label, data: d.data, unknown: d.unknown,
        backgroundColor: d.stripe ? (opts.print ? '#C4C8CE' : stripePattern()) : (ctx => model.groups[ctx.dataIndex] && model.groups[ctx.dataIndex].a.anyEst ? rgba(d.color, .5) : (opts.print ? d.color : barGradient(d.color)(ctx))),
        hoverBackgroundColor: d.stripe ? '#9CA3AF' : shade(d.color, -25),
        borderRadius: {topLeft: 6, topRight: 6}, borderSkipped: 'bottom',
        maxBarThickness: 38, categoryPercentage: .78, barPercentage: .88
      }))
    },
    options: {
      responsive: !opts.print, maintainAspectRatio: false, animation: anim,
      devicePixelRatio: opts.print ? 2 : undefined,
      layout: {padding: {top: 26, right: 8, left: 4}},
      interaction: {mode: 'index', intersect: false},
      scales: {
        x: {grid: {display: false}, border: {display: false}, ticks: {autoSkip: false, maxRotation: 0, font: {size: 11, weight: 'bold'}, color: '#374151'}},
        y: {beginAtZero: true, grace: '12%', border: {display: false}, grid: {color: '#EEF0F3'}, ticks: {precision: 0, color: '#9CA3AF'}}
      },
      plugins: {
        legend: {display: !!opts.print, position: 'top', align: 'end', labels: {usePointStyle: true, pointStyle: 'rectRounded', font: {weight: 'bold'}}},
        valueLabels: {enabled: true, size: opts.print ? 10 : 11},
        tooltip: {callbacks: {
          title: items => model.groups[items[0].dataIndex].title,
          label: it => ` ${it.dataset.label}: ${it.raw == null ? '—' : fmtN(it.raw)}`,
          footer: items => { const g = model.groups[items[0].dataIndex]; return `Total: ${fmtN(g.a.total)}${g.a.anyManual ? '  (inclui lançamento sem detalhamento)' : ''}${g.a.anyEst ? '  (estimado pelas datas)' : ''}`; }
        }}
      }
    }
  };
}
function legendHTML(datasets){ return datasets.map((d, i) => `<button type="button" data-i="${i}" style="--c:${d.color}"><i ${d.stripe ? 'style="background:repeating-linear-gradient(45deg,#9CA3AF 0 3px,#D1D5DB 3px 6px)"' : ''}></i>${esc(d.label)}</button>`).join(''); }
function bindLegend(box, chartId){
  box.onclick = e => { const b = e.target.closest('button'); if (!b) return; const ch = charts[chartId]; if (!ch) return; const i = +b.dataset.i; const vis = ch.isDatasetVisible(i); ch.setDatasetVisibility(i, !vis); b.classList.toggle('off', vis); ch.update(); };
}
function perGroupWidth(model){ const n = model.datasets.length; return Math.max(model.mode === 'semana' ? 70 : 92, n * 30 + 26); }

/* ---------- Percurso (início → pico → final) ---------- */
function journeyModel(weeks, totals, cats){
  const idx = totals.map((v, i) => v == null ? -1 : i).filter(i => i >= 0);
  if (idx.length < 2) return null;
  const s = idx[0], e = idx[idx.length - 1];
  let p = s; for (const i of idx) if (totals[i] > totals[p]) p = i;
  const byIdx = new Map(); const add = (i, k) => { if (!byIdx.has(i)) byIdx.set(i, {idx:i, kinds:[], value: totals[i]}); byIdx.get(i).kinds.push(k); };
  add(s, 'start'); add(p, 'peak'); add(e, 'end');
  return {weeks, totals, cats, s, p, e, miles: [...byIdx.values()].sort((a, b) => a.idx - b.idx)};
}
function pct(a, b){ if (a == null || b == null) return null; if (a === 0) return b === 0 ? 0 : null; return (b - a) / a * 100; }
function pctChip(v, inverse){ if (v == null) return ''; const up = v > 0, z = Math.abs(v) < .05; const cls = z ? '' : (up !== !!inverse ? 'ok' : 'err'); return `<span class="chip ${cls}">${z ? '=' : up ? '▲' : '▼'} ${Math.abs(v).toLocaleString('pt-BR', {maximumFractionDigits: 1})}%</span>`; }
function nextWeek(w){ const [y, m, d] = weekRange(w).start.split('-').map(Number); return isoWeek(ymdUTC(new Date(Date.UTC(y, m - 1, d + 7)))); }
function weekSpan(a, b){ const out = []; let w = a, g = 0; while (w <= b && g++ < 600) { out.push(w); w = nextWeek(w); } return out; }
/** Percurso por obra — cada obra tem sua própria série; obras diferentes nunca são somadas. */
function obraJourneys(kind, model, f){
  const K = KINDS[kind];
  return model.obras.map(o => {
    const own = model.weeks.filter(w => model.cells.has(w + '||' + o)); if (!own.length) return null;
    const weeks = f.weeks && f.weeks.length ? [...f.weeks].filter(w => passCommon('', w, {from: f.from, to: f.to})).sort() : weekSpan(own[0], own[own.length - 1]);
    const cell = w => model.cells.get(w + '||' + o);
    const totals = weeks.map(w => cell(w) ? cell(w).total : null);
    const cats = K.cats.filter(c => !f.cat || f.cat === c.k).map(c => ({k: c.k, l: c.l, color: c.color, data: weeks.map(w => cell(w) ? cell(w)[c.k] : null)}));
    const sd = weeks.map(w => cell(w) ? (cell(w).SD || 0) : null);
    if (sd.some(v => v > 0)) cats.push({k: 'SD', l: SD.l, color: '#9CA3AF', data: sd});
    const nn = totals.filter(v => v != null);
    return {o, label: model.labels.get(o), weeks, totals, n: nn.length, last: [...totals].reverse().find(v => v != null), jm: journeyModel(weeks, totals, cats)};
  }).filter(Boolean);
}
function sparkSVG(vals){
  const nn = vals.filter(v => v != null); if (!nn.length) return '';
  const max = Math.max(...nn, 1), n = vals.length, X = i => n === 1 ? 50 : i / (n - 1) * 100, Y = v => 25 - v / max * 21;
  const segs = []; let cur = [];
  vals.forEach((v, i) => { if (v == null) { if (cur.length) segs.push(cur); cur = []; } else cur.push([X(i), Y(v)]); }); if (cur.length) segs.push(cur);
  const pts = sg => sg.map(p => p[0].toFixed(1) + ',' + p[1].toFixed(1)).join(' ');
  return `<svg viewBox="0 0 100 28" preserveAspectRatio="none">${segs.map(sg => sg.length > 1
    ? `<polygon points="${sg[0][0].toFixed(1)},28 ${pts(sg)} ${sg[sg.length - 1][0].toFixed(1)},28" fill="rgba(247,147,30,.16)"/><polyline points="${pts(sg)}" fill="none" stroke="#F7931E" stroke-width="2" vector-effect="non-scaling-stroke" stroke-linejoin="round"/>`
    : `<line x1="${sg[0][0] - 1.5}" y1="${sg[0][1]}" x2="${sg[0][0] + 1.5}" y2="${sg[0][1]}" stroke="#F7931E" stroke-width="3" vector-effect="non-scaling-stroke"/>`).join('')}</svg>`;
}
UI.jrSel = {};
function renderJourney(box, id, list, opts){
  if (journeys[id]) cancelAnimationFrame(journeys[id].raf);
  killChart(id);
  const head = `<div class="card-h"><h3>Percurso do efetivo por obra</h3><span class="hint">${esc(opts.subtitle || '')}</span><span class="sp"></span>${opts.extra || ''}<div class="legend jr-leg"></div></div>`;
  if (!list.length) { box.innerHTML = head + `<div class="jr-frame"><div class="empty"><b>Sem dados no filtro</b>Cadastre ou importe registros para acompanhar a evolução semanal de cada obra.</div></div>`; return; }
  const sel = list.find(x => x.o === UI.jrSel[id]) || list.find(x => x.jm) || list[0];
  UI.jrSel[id] = sel.o;
  const jm = sel.jm;
  const mileCard = (k, title, val, wk, extra) => `<div class="mile" ${k ? `data-m="${k}"` : ''} style="--c:${k ? MILE[k].c : '#F7931E'}"><div class="t"><i></i>${title}</div><div class="v" data-n="${val}">0</div><div class="w">${wk}</div>${extra ? `<div class="d">${extra}</div>` : ''}</div>`;
  let miles = '';
  if (jm) {
    const {weeks, totals: T, s, p, e} = jm;
    miles = mileCard('start', 'Início', T[s], esc(weekLabel(weeks[s])))
      + mileCard('peak', 'Pico', T[p], esc(weekLabel(weeks[p])), pctChip(pct(T[s], T[p])))
      + mileCard('end', 'Final', T[e], esc(weekLabel(weeks[e])), p !== e ? pctChip(pct(T[p], T[e])) : '')
      + mileCard('', 'Início → final', T[e] - T[s], `${esc(opts.unit)} ${pctChip(pct(T[s], T[e]))}`);
  }
  box.innerHTML = head + `<div class="jr-grid">
    <aside class="jr-side">
      <div class="jr-side-h">Selecione a obra <span class="chip">${list.length}</span></div>
      ${list.length > 6 ? '<input type="search" class="jr-q" placeholder="Buscar obra…">' : ''}
      <div class="jr-list">${list.map(x => `<button type="button" class="jr-item ${x.o === sel.o ? 'on' : ''}" data-o="${esc(x.o)}" title="${esc(x.label)}"><span class="nm">${esc(x.label)}</span><span class="sk">${sparkSVG(x.totals)}</span><span class="lv">${x.last == null ? '—' : fmtN(x.last)}<small>${x.n} sem.</small></span></button>`).join('')}</div>
      <select class="jr-select" aria-label="Obra">${list.map(x => `<option value="${esc(x.o)}" ${x.o === sel.o ? 'selected' : ''}>${esc(x.label)}</option>`).join('')}</select>
    </aside>
    <section class="jr-main">
      <div class="jr-top"><div class="jr-obra">${esc(sel.label)}<small>${sel.weeks.length} semana(s) · ${esc(weekLabel(sel.weeks[0]))} até ${esc(weekLabel(sel.weeks[sel.weeks.length - 1]))}</small></div>
        ${jm ? `<div class="jr-read"><span class="rtag"></span><span class="rv">—</span><span class="ru">${esc(opts.unit)}</span><span class="rw"></span><span class="rd"></span></div>` : ''}</div>
      <div class="jr-frame">${jm ? '<div class="jr-chart"><canvas></canvas></div>' : `<div class="empty"><b>Somente uma semana com registro nesta obra</b>O percurso (início → pico → final) precisa de duas ou mais semanas. Ajuste o filtro de semanas ou escolha outra obra.</div>`}</div>
      ${jm ? `<div class="jr-ctrl"><button class="play" title="Reproduzir percurso"></button><input type="range" min="0" max="${sel.weeks.length - 1}" step="0.01" value="0" aria-label="Linha do tempo"><div class="seg jr-speed"><button data-s="0.5">0,5×</button><button data-s="1" class="on">1×</button><button data-s="2">2×</button></div></div>` : ''}
    </section>
    ${jm ? `<aside class="jr-miles">${miles}</aside>` : ''}
  </div>`;
  const pick = o => { UI.jrSel[id] = o; renderJourney(box, id, list, opts); };
  $$('.jr-item', box).forEach(b => b.onclick = () => pick(b.dataset.o));
  $('.jr-select', box).onchange = ev => pick(ev.target.value);
  const q = $('.jr-q', box); if (q) q.oninput = () => { const t = norm(q.value); $$('.jr-item', box).forEach(b => b.hidden = !!t && !norm(b.textContent).includes(t)); };
  const selBtn = $('.jr-item.on', box); if (selBtn) selBtn.scrollIntoView({block: 'nearest'});
  if (!jm) { $('.jr-leg', box).innerHTML = ''; return; }
  const {weeks, totals, s} = jm;
  $$('.mile .v', box).forEach(el => countUp(el, Number(el.dataset.n), 1100));
  const canvas = $('canvas', box);
  const ds = [{
    label: 'Total', data: totals, borderColor: '#111827', borderWidth: 3.5, tension: 0, spanGaps: false,
    pointRadius: 3.5, pointBackgroundColor: '#fff', pointBorderColor: '#111827', pointBorderWidth: 2, pointHoverRadius: 6,
    fill: true, backgroundColor: c => { const a = c.chart.chartArea; if (!a) return 'rgba(247,147,30,.15)'; const g = c.chart.ctx.createLinearGradient(0, a.top, 0, a.bottom); g.addColorStop(0, 'rgba(247,147,30,.30)'); g.addColorStop(1, 'rgba(247,147,30,0.02)'); return g; },
    order: 0
  }];
  (jm.cats || []).forEach(c => ds.push({label: c.l, data: c.data, borderColor: c.color, backgroundColor: c.color, borderWidth: 2, tension: 0, pointRadius: 2, pointHoverRadius: 5, borderDash: c.k === 'SD' ? [5, 4] : undefined, fill: false, spanGaps: false, order: 1}));
  const leg = $('.jr-leg', box); leg.innerHTML = legendHTML(ds.map(d => ({label: d.label, color: d.borderColor})));
  const ch = new Chart(canvas, {
    type: 'line',
    plugins: [{id: 'jrInit', beforeInit(c){ c.$j = {n: weeks.length, reveal: reduceMotion ? weeks.length - 1 : 0, cursor: !reduceMotion, miles: jm.miles}; }}],
    data: {labels: weeks.map(w => [weekShort(w), fmtBR(weekRange(w).end).slice(0, 5)]), datasets: ds},
    options: {
      responsive: true, maintainAspectRatio: false, animation: false,
      layout: {padding: {top: 36, right: 14, left: 8, bottom: 4}},
      interaction: {mode: 'index', intersect: false},
      scales: {x: {offset: true, grid: {display: false}, border: {color: '#E5E7EB'}, ticks: {maxRotation: 0, autoSkip: true, font: {size: 11, weight: 'bold'}}},
               y: {beginAtZero: true, grace: '20%', border: {display: false}, grid: {color: '#F0F1F4'}, ticks: {precision: 0, color: '#9CA3AF', maxTicksLimit: 6}}},
      plugins: {legend: {display: false}, tooltip: {callbacks: {title: it => `${sel.label} — ${weekLabel(weeks[it[0].dataIndex], true)}`, label: it => ` ${it.dataset.label}: ${it.raw == null ? 'sem registro' : fmtN(it.raw)}`}}}
    }
  });
  charts[id] = ch; bindLegend(leg, id);
  const J = journeys[id] = {raf: 0, playing: false, speed: 1, hold: 0};
  const read = $('.jr-read', box), range = $('input[type=range]', box), playBtn = $('.play', box);
  const ICON_PLAY = '<svg viewBox="0 0 24 24" fill="currentColor"><path d="M7 4.5v15l13-7.5z"/></svg>';
  const ICON_PAUSE = '<svg viewBox="0 0 24 24" fill="currentColor"><path d="M6 4h4.5v16H6zM13.5 4H18v16h-4.5z"/></svg>';
  const ICON_REPLAY = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.6" stroke-linecap="round"><path d="M4 12a8 8 0 1 0 2.4-5.7"/><path d="M4 4v4.5h4.5"/></svg>';
  playBtn.innerHTML = ICON_PLAY;
  let lastTag = null;
  function paintReadout(){
    const i = Math.max(0, Math.min(weeks.length - 1, Math.round(ch.$j.reveal)));
    const val = totals[i];
    $('.rv', read).textContent = val == null ? '—' : fmtN(val);
    $('.rw', read).textContent = weekLabel(weeks[i]);
    const d = pct(totals[s], val);
    $('.rd', read).textContent = val == null ? 'sem registro (não transportado)' : (i === s ? 'ponto de partida' : `vs. início ${val - totals[s] >= 0 ? '+' : ''}${fmtN(val - totals[s])}${d != null ? ` (${d >= 0 ? '+' : ''}${d.toLocaleString('pt-BR', {maximumFractionDigits: 1})}%)` : ''}`);
    const m = jm.miles.find(mm => mm.idx === i);
    const tag = $('.rtag', read);
    if (m) { const k = m.kinds[m.kinds.length - 1]; tag.textContent = m.kinds.map(x => MILE[x].l).join(' · '); tag.style.background = MILE[k].c; }
    else { tag.textContent = `SEMANA ${i + 1}/${weeks.length}`; tag.style.background = '#2A2A30'; }
    if (m && lastTag !== i) m.kinds.forEach(k => { const c = $(`.mile[data-m="${k}"]`, box); if (c) { c.classList.remove('hit'); void c.offsetWidth; c.classList.add('hit'); } });
    lastTag = i;
    if (!J.scrubbing) range.value = ch.$j.reveal;
  }
  function stop(final){ cancelAnimationFrame(J.raf); J.playing = false; playBtn.innerHTML = ch.$j.reveal >= weeks.length - 1 - 1e-6 ? ICON_REPLAY : ICON_PLAY; if (final) { ch.$j.cursor = false; ch.draw(); } }
  function run(totalMs, fromStart){
    cancelAnimationFrame(J.raf); J.playing = true; playBtn.innerHTML = ICON_PAUSE; ch.$j.cursor = true;
    if (fromStart || ch.$j.reveal >= weeks.length - 1 - 1e-6) { ch.$j.reveal = 0; lastTag = null; }
    let last = performance.now(); J.hold = totalMs ? 0 : 500 / J.speed;
    const msPer = totalMs ? totalMs / Math.max(1, weeks.length - 1) : 1000 / J.speed;
    const step = now => {
      if (!box.isConnected) return;
      const dt = Math.min(64, now - last); last = now;
      if (J.hold > 0) J.hold -= dt;
      else {
        const prev = ch.$j.reveal; ch.$j.reveal = Math.min(weeks.length - 1, prev + dt / msPer);
        if (!totalMs) { const hit = jm.miles.find(m => prev < m.idx && ch.$j.reveal >= m.idx); if (hit) { ch.$j.reveal = hit.idx; J.hold = 750 / J.speed; } }
      }
      ch.draw(); paintReadout();
      if (ch.$j.reveal >= weeks.length - 1 && J.hold <= 0) { stop(true); paintReadout(); return; }
      J.raf = requestAnimationFrame(step);
    };
    J.raf = requestAnimationFrame(step);
  }
  playBtn.onclick = () => { if (J.playing) stop(false); else run(0); };
  range.oninput = () => { stop(false); J.scrubbing = true; ch.$j.cursor = true; ch.$j.reveal = Number(range.value); ch.draw(); paintReadout(); J.scrubbing = false; playBtn.innerHTML = ICON_PLAY; };
  $('.jr-speed', box).onclick = ev => { const b = ev.target.closest('button'); if (!b) return; J.speed = Number(b.dataset.s); $$('.jr-speed button', box).forEach(x => x.classList.toggle('on', x === b)); };
  if (reduceMotion) { ch.$j.reveal = weeks.length - 1; ch.$j.cursor = false; ch.draw(); paintReadout(); playBtn.innerHTML = ICON_REPLAY; }
  else run(Math.min(2200, 350 * weeks.length), true);
}

/* =====================================================================
   ABAS MÃO DE OBRA / EQUIPAMENTOS
   ===================================================================== */
function weekOptions(kind){
  const set = new Set(Repo.data[kind].map(r => r.semana));
  for (const m of Repo.data.manual) if (manualGroup(m, kind)) set.add(m.semana);
  if (UI.f[kind] && UI.f[kind].basis === 'datas') { const t = obraTimeline(kind, {...UI.f[kind], obras: []}); t.forEach(ob => weeksBetween(ob.first, ob.last, 104).forEach(w => set.add(w))); }
  return [...set].filter(weekRange).sort().reverse().map(w => ({v: w, l: w, sub: weekLabel(w).split('·')[1]}));
}
function obraOptions(kinds){
  const mp = new Map();
  for (const k of kinds) for (const r of Repo.data[k]) if (!mp.has(obraKey(r.obra))) mp.set(obraKey(r.obra), r.obra);
  for (const m of Repo.data.manual) if (!mp.has(obraKey(m.obra))) mp.set(obraKey(m.obra), m.obra);
  return [...mp].map(([v, l]) => ({v, l})).sort((a, b) => a.l.localeCompare(b.l, 'pt-BR'));
}
function buildKindTab(kind){
  const K = KINDS[kind]; const sec = $('#tab-' + kind);
  sec.innerHTML = `
  <div class="toolbar"><div><h2>${K.title}</h2><div class="sub">Efetivo no fechamento de cada semana ISO (segunda a domingo)</div></div>
    <div class="acts">
      <button class="btn primary" data-a="new">＋ Novo cadastro</button>
      <button class="btn" data-a="import">⇪ Importar planilha</button>
      <button class="btn" data-a="template">⬇ Modelo Excel</button>
      <button class="btn" data-a="conf">✓ Conferir com a planilha</button>
      <button class="btn" data-a="excel">⊞ Exportar Excel</button>
      <button class="btn dark" data-a="pdf">⎙ PDF</button>
    </div></div>
  <div class="card filters">
    <label>Obra<div data-ms="obras"></div></label>
    <label>Semanas<div data-ms="weeks"></div></label>
    <label>Período — de<input type="date" data-f="from"></label>
    <label>até<input type="date" data-f="to"></label>
    <label>Base do cálculo<select data-f="basis"><option value="foto">Fotografias semanais importadas</option><option value="datas">Reconstruir semanas pelas datas (${kind === 'labor' ? 'admissão/demissão' : 'mobilização/desmob.'})</option></select></label>
    <label>${K.catTitle}<select data-f="cat"><option value="">Todas</option>${K.cats.map(c => `<option value="${c.k}">${esc(c.full)}</option>`).join('')}</select></label>
    ${kind === 'labor' ? '<label>Situação<select data-f="sit"></select></label>' : ''}
    <label>Busca<input type="search" data-f="q" placeholder="Nome, função, prefixo…"></label>
    <button class="btn ghost" data-a="clear">Limpar filtros</button>
  </div>
  <div class="kpis" data-kpis></div>
  <div class="card" data-hist>
    <div class="card-h"><h3>Histograma</h3><span class="hint" data-hint></span><span class="sp"></span>
      <div class="seg" data-mode>
        <button data-m="auto">Automático</button><button data-m="obra">Por obra (semana mais recente)</button><button data-m="semanaobra">Semana × obra</button>
      </div>
    </div>
    <div class="legend" data-leg style="margin-bottom:6px"></div>
    <div data-note></div>
    <div class="chart-frame"><div class="chart-scroll"><div class="chart-inner"><canvas></canvas></div></div></div>
  </div>
  <div class="card journey" data-journey></div>
  <div class="card">
    <div class="card-h"><h3>Relação de cadastros</h3><span class="hint" data-count></span></div>
    <div data-manual></div>
    <div class="tbl-wrap"><table><thead></thead><tbody></tbody></table></div>
    <div class="pager" data-pager></div>
  </div>`;
  const f = UI.f[kind];
  UI.ms[kind] = {
    obras: MultiSelect($('[data-ms=obras]', sec), {all: 'Todas as obras', value: f.obras, options: () => obraOptions([kind]),
      quick: [{l: 'Todas', fn: () => []}], onChange: v => { f.obras = v; UI.page[kind] = 1; renderKind(kind); }}),
    weeks: MultiSelect($('[data-ms=weeks]', sec), {all: 'Todas as semanas', value: f.weeks, options: () => weekOptions(kind),
      quick: [{l: 'Todas', fn: () => []}, {l: 'Últimas 4', fn: a => a.slice(0, 4)}, {l: 'Últimas 8', fn: a => a.slice(0, 8)}, {l: 'Últimas 12', fn: a => a.slice(0, 12)}],
      onChange: v => { f.weeks = v; UI.page[kind] = 1; renderKind(kind); }})
  };
  let qT;
  sec.addEventListener('input', e => { const k = e.target.dataset.f; if (!k) return; if (k === 'q') { clearTimeout(qT); qT = setTimeout(() => { f.q = e.target.value; UI.page[kind] = 1; renderKind(kind); }, 220); } });
  sec.addEventListener('change', e => { const k = e.target.dataset.f; if (!k || k === 'q') return; f[k] = e.target.value; UI.page[kind] = 1; renderKind(kind); });
  sec.addEventListener('click', e => {
    const b = e.target.closest('[data-a]'); if (b) {
      const a = b.dataset.a;
      if (a === 'new') openRecordForm(kind);
      else if (a === 'import') openImport(kind);
      else if (a === 'template') downloadTemplate();
      else if (a === 'excel') exportExcel(kind);
      else if (a === 'conf') { switchTab('avisos'); if (!Sync.lastWb) Sync.run(false); setTimeout(() => { const h = [...document.querySelectorAll('#avisosBody h3')].find(x => /Conferência/.test(x.textContent)); h && h.scrollIntoView({behavior: 'smooth'}); }, 80); }
      else if (a === 'pdf') printKind(kind);
      else if (a === 'clear') { Object.assign(f, blankFilter()); $$('[data-f]', sec).forEach(i => i.value = ''); UI.ms[kind].obras.set([], true); UI.ms[kind].weeks.set([], true); renderKind(kind); }
      else if (a === 'edit') openRecordForm(kind, Repo.data[kind].find(r => r.id === b.dataset.id));
      else if (a === 'del') deleteRecord(kind, b.dataset.id);
      else if (a === 'medit') openManualForm(Repo.data.manual.find(m => m.id === b.dataset.id));
      else if (a === 'mdrop') dropManualAsk(kind, b.dataset.id);
      else if (a === 'pg') { UI.page[kind] = +b.dataset.p; renderTable(kind, UI._last[kind]); }
      return;
    }
    const m = e.target.closest('[data-mode] button'); if (m) { UI.mode[kind] = m.dataset.m; renderKind(kind); }
  });
}
UI._last = {};
function renderKind(kind){
  const K = KINDS[kind]; const sec = $('#tab-' + kind); const f = UI.f[kind];
  UI.ms[kind].obras.refresh(); UI.ms[kind].weeks.refresh();
  f.obras = UI.ms[kind].obras.value; f.weeks = UI.ms[kind].weeks.value;
  if (kind === 'labor') {
    const sel = $('[data-f=sit]', sec); const vals = new Map();
    for (const r of Repo.data.labor) { const s = clean(r.situacao); if (s && !vals.has(norm(s))) vals.set(norm(s), s); }
    sel.innerHTML = `<option value="">Todas as situações</option>${[...vals].sort((a, b) => a[1].localeCompare(b[1], 'pt-BR')).map(([k, l]) => `<option value="${esc(k)}">${esc(l)}</option>`).join('')}<option value="__none__">Não informada</option>`;
    sel.value = f.sit;
  }
  $('[data-f=basis]', sec).value = f.basis || 'foto';
  $$('[data-mode] button', sec).forEach(b => b.classList.toggle('on', b.dataset.m === UI.mode[kind]));
  const model = histogramModel(kind, f, UI.mode[kind]);
  UI._last[kind] = model;
  renderKpis(kind, model);
  renderHist(kind, model);
  renderKindJourney(kind, model);
  renderTable(kind, model);
}
/** Para cada obra: a semana mais recente (dentro do filtro) que tem dados */
function latestPerObra(model){
  const out = [];
  for (const o of model.obras) { for (let i = model.weeks.length - 1; i >= 0; i--) { const c = model.cells.get(model.weeks[i] + '||' + o); if (c) { out.push({o, w: model.weeks[i], c}); break; } } }
  return out;
}
function renderKpis(kind, model){
  const K = KINDS[kind]; const box = $(`#tab-${kind} [data-kpis]`);
  // Efetivo atual = soma, por obra, da semana mais recente de cada obra no período filtrado (não soma semanas)
  const lat = latestPerObra(model);
  const a = lat.length ? aggCells(kind, lat.map(x => x.c)) : null;
  const ws = [...new Set(lat.map(x => x.w))].sort();
  const wTxt = !ws.length ? 'sem dados no filtro' : ws.length === 1 ? weekLabel(ws[0]) : 'última semana de cada obra: ' + lat.map(x => `${model.labels.get(x.o)} ${weekShort(x.w)}`).join(' · ');
  const distinct = new Set(model.rows.map(r => vincKey(kind, r))).size;
  const items = [
    {l: lat.length <= 1 ? `Efetivo no fechamento` : `Efetivo atual — ${lat.length} obras`, v: a ? a.total : null, s: wTxt, dark: true},
    ...K.cats.map(c => ({l: c.full, v: a ? a[c.k] : null, s: a && a.SD ? `+ ${fmtN(a.SD)} sem detalhamento` : (ws.length === 1 ? ws[0] : `${ws.length ? weekShort(ws[0]) + ' a ' + weekShort(ws[ws.length - 1]) : ''}`), c: c.color})),
    {l: kind === 'labor' ? 'Pessoas distintas no filtro' : 'Equipamentos distintos no filtro', v: distinct, s: `${fmtN(model.rows.length)} registros semanais · ${model.weeks.length} semana(s) · ${model.obras.length} obra(s)`, c: '#111'}
  ];
  box.innerHTML = items.map(i => `<div class="kpi ${i.dark ? 'dark' : ''}" style="--c:${i.c || '#F7931E'}"><div class="l">${esc(i.l)}</div><div class="v" data-n="${i.v ?? ''}">0</div><div class="s">${esc(i.s)}</div></div>`).join('');
  $$('.kpi .v', box).forEach(el => countUp(el, el.dataset.n === '' ? null : Number(el.dataset.n)));
}
function renderHist(kind, model){
  const sec = $(`#tab-${kind} [data-hist]`); const id = 'hist-' + kind; killChart(id);
  const inner = $('.chart-inner', sec), leg = $('[data-leg]', sec);
  const modeName = {obra: 'por obra', semanaobra: model.obras.length === 1 ? 'evolução semanal da obra' : 'por semana e obra (cada obra separada)'}[model.mode];
  $('[data-hint]', sec).textContent = `Efetivo no fechamento · ${modeName}`;
  $('[data-note]', sec).innerHTML = ''; updateBadge();
  if (!model.groups.length) { leg.innerHTML = ''; inner.style.width = '100%'; inner.innerHTML = '<div class="empty"><b>Nenhum dado para os filtros</b>Cadastre, importe ou ajuste os filtros.</div>'; return; }
  inner.innerHTML = '<canvas></canvas>';
  const wrapW0 = ($('.chart-scroll', sec).clientWidth || 800) - 2;
  const n0 = model.groups.length, natural0 = n0 * perGroupWidth(model) + 60, comfy0 = n0 * (model.datasets.length * 58 + 64) + 90;
  inner.style.width = (natural0 >= wrapW0 ? natural0 : Math.max(Math.min(wrapW0, comfy0), Math.min(wrapW0, 620))) + 'px';
  charts[id] = new Chart($('canvas', inner), barConfig(model));
  leg.innerHTML = legendHTML(model.datasets); bindLegend(leg, id);
}
/** Mensagens/orientações do histograma (exibidas na página Avisos) */
function histNotes(kind, model){
  const notes = [];
  if (model.note) notes.push(model.note);
  if (model.mnotes.length) notes.push(`${model.mnotes.length} lançamento(s) sem detalhamento não entram no filtro de categoria (categorias desconhecidas).`);
  if (model.datasets.some(d => d.key === 'SD')) notes.push('Barras hachuradas = total lançado sem detalhamento; “—” indica categoria desconhecida.');
  const f = UI.f[kind]; let extra = '';
  if (f.basis === 'datas') extra = `<div class="alert warn" style="padding:6px 10px;font-size:12px">Semanas sem fotografia importada foram <b>reconstruídas pelas datas</b> (${kind === 'labor' ? 'admissão/demissão' : 'mobilização/desmobilização'}) — barras mais claras e marcadas com “est.”. Quem saiu da obra e não consta em nenhuma fotografia não pode ser contado.</div>`;
  else { const miss = missingWeeksHint(kind, f); if (miss.length) extra = `<div class="alert warn" style="padding:7px 10px;font-size:12px;display:flex;gap:10px;align-items:center;flex-wrap:wrap"><span>${miss.length === 1 ? `<b>${esc(miss[0].label)}</b> tem fotografia em ${miss[0].snaps} semana(s), mas há ${kind === 'labor' ? 'admissões' : 'mobilizações'} desde ${esc(weekShort(miss[0].first))}` : `<b>${miss.length} obras</b> têm semanas sem fotografia importada`}. O histograma mostra somente as semanas importadas.</span><button class="btn sm primary" data-basis="datas">Ver semana a semana pelas datas</button></div>`; }
  return {extra, notes};
}
function renderKindJourney(kind, model){
  const f = UI.f[kind]; const box = $(`#tab-${kind} [data-journey]`);
  renderJourney(box, 'jr-' + kind, obraJourneys(kind, model, f), {unit: KINDS[kind].unit, subtitle: 'Efetivo no fechamento, semana a semana, de cada obra (obras não são somadas)'});
}
const PAGE = 50;
function renderTable(kind, model){
  const K = KINDS[kind]; const sec = $('#tab-' + kind);
  const rows = [...model.rows].sort((a, b) => b.semana.localeCompare(a.semana) || String(a.obra).localeCompare(String(b.obra), 'pt-BR') || String(a.nome || a.prefixo).localeCompare(String(b.nome || b.prefixo), 'pt-BR'));
  const pages = Math.max(1, Math.ceil(rows.length / PAGE)); const pg = Math.min(UI.page[kind], pages); UI.page[kind] = pg;
  $('[data-count]', sec).textContent = `${fmtN(rows.length)} cadastro(s) no filtro`;
  // Totais sem detalhamento que afetam o filtro
  const f = UI.f[kind]; const ms = [];
  for (const [, g] of model.mm) { if (passCommon(g.src.obra, g.src.semana, f)) ms.push(g); }
  const hid = model.hidden.length;
  $('[data-manual]', sec).innerHTML = ms.length ? `<div class="alert warn"><b>Sem detalhamento</b> — ${ms.length} lançamento(s) manual(is) complementam as categorias sem cadastro nas obras/semanas abaixo${hid ? ` (${hid} cadastro(s) individuais preservados, porém ocultos)` : ''}:
    <div style="margin-top:8px;display:flex;flex-direction:column;gap:6px">${ms.sort((a, b) => b.src.semana.localeCompare(a.src.semana)).slice(0, 30).map(g => { const nh = model.hidden.filter(r => r.semana === g.src.semana && obraKey(r.obra) === obraKey(g.src.obra)).length;
      return `<div style="display:flex;gap:8px;align-items:center;flex-wrap:wrap;background:#fff;border:1px solid #FDE68A;border-radius:8px;padding:6px 8px"><span class="chip warn">${esc(g.src.semana)}</span><b>${esc(g.src.obra)}</b><span>${K.cats.map(c => `${c.l}: ${g.unknown ? '—' : fmtN(g[c.k])}`).join(' · ')} · <b>Total ${fmtN(g.total)}</b></span>${nh ? `<span class="chip">${nh} cadastro(s) oculto(s)</span>` : ''}<span style="margin-left:auto;display:flex;gap:6px"><button class="btn sm" data-a="medit" data-id="${g.src.id}">✎ Editar</button><button class="btn sm danger" data-a="mdrop" data-id="${g.src.id}">✕ Remover</button></span></div>`; }).join('')}</div></div>` : '';
  const cols = K.fields;
  $('thead', sec).innerHTML = `<tr><th>Ações</th>${cols.map(c => `<th>${esc(c.l)}</th>`).join('')}<th>No fechamento</th></tr>`;
  const slice = rows.slice((pg - 1) * PAGE, pg * PAGE);
  $('tbody', sec).innerHTML = slice.length ? slice.map(r => {
    const act = activeAtClose(kind, r, weekRange(r.semana).end);
    return `<tr><td><button class="btn sm" data-a="edit" data-id="${r.id}" title="Editar">✎</button> <button class="btn sm danger" data-a="del" data-id="${r.id}" title="Excluir">✕</button></td>${cols.map(c => {
      const v = displayVal(kind, r, c);
      if (c.t === 'cat') return `<td><span class="chip" style="background:${rgba((K.cats.find(x => x.k === r[c.k]) || {color: '#999999'}).color, .15)}">${esc(v)}</span></td>`;
      return `<td class="${['nome', 'descricao', 'obra', 'funcao'].includes(c.k) ? 'l' : ''}">${esc(v)}</td>`;
    }).join('')}<td>${act ? '<span class="chip ok">Sim</span>' : '<span class="chip">Não</span>'}</td></tr>`;
  }).join('') : `<tr><td colspan="${cols.length + 2}" class="muted" style="padding:20px">Nenhum cadastro para os filtros.</td></tr>`;
  $('[data-pager]', sec).innerHTML = pages > 1 ? `<button class="btn sm" data-a="pg" data-p="${Math.max(1, pg - 1)}" ${pg === 1 ? 'disabled' : ''}>‹ Anterior</button> Página ${pg} de ${pages} <button class="btn sm" data-a="pg" data-p="${Math.min(pages, pg + 1)}" ${pg === pages ? 'disabled' : ''}>Próxima ›</button>` : '';
}

/* =====================================================================
   MODAIS
   ===================================================================== */
function openModal({title, body, foot = '', wide = false, onClose}){
  const bg = document.createElement('div'); bg.className = 'modal-bg';
  bg.innerHTML = `<div class="modal ${wide ? 'wide' : ''}" role="dialog" aria-modal="true"><div class="modal-h"><h3>${esc(title)}</h3><button class="x" aria-label="Fechar">×</button></div><div class="modal-b">${body}</div><div class="modal-f">${foot}</div></div>`;
  document.body.appendChild(bg);
  const close = () => { bg.remove(); onClose && onClose(); };
  $('.x', bg).onclick = close;
  bg.addEventListener('mousedown', e => { if (e.target === bg) close(); });
  return {el: bg, close};
}
function confirmBox(title, html, okLabel = 'Confirmar', danger = false){
  return new Promise(res => {
    const m = openModal({title, body: html, foot: `<button class="btn" data-x>Cancelar</button><button class="btn ${danger ? 'dark' : 'primary'}" data-ok>${esc(okLabel)}</button>`, onClose: () => res(false)});
    $('[data-x]', m.el).onclick = () => m.close();
    $('[data-ok]', m.el).onclick = () => { m.el.remove(); res(true); };
  });
}
function weekInputHTML(name, val){
  return `<div class="week-in"><input data-k="${name}" value="${esc(val || '')}" placeholder="2026-W40" autocomplete="off"><input type="date" data-pick="${name}" title="Escolher pelo calendário"></div>`;
}
function bindWeekPickers(root){
  $$('[data-pick]', root).forEach(p => p.addEventListener('change', () => { if (p.value) { const t = $(`[data-k="${p.dataset.pick}"]`, root); t.value = isoWeek(p.value); t.dispatchEvent(new Event('input', {bubbles: true})); } }));
}

/* ---------- cadastro / edição ---------- */
function validateRecord(kind, raw, opt = {}){
  const K = KINDS[kind]; const rec = {}; const errors = []; const warns = [];
  for (const fd of K.fields) {
    const v = raw[fd.k];
    if (fd.t === 'date') { const d = parseDate(v, opt.order); rec[fd.k] = d.v; if (d.err) errors.push({field: fd.l, msg: d.err}); if (d.dash) warns.push(`${fd.l} "-" tratado como não informado`); }
    else if (fd.t === 'week') { const w = parseWeek(v); rec[fd.k] = w.v; if (w.err) errors.push({field: fd.l, msg: w.err}); }
    else if (fd.t === 'cat') { const c = normCat(kind, v); rec[fd.k] = c || ''; if (c === null) errors.push({field: fd.l, msg: `Valor não reconhecido "${clean(v)}"`}); }
    else rec[fd.k] = clean(v);
    if (fd.req && !rec[fd.k] && !errors.some(e => e.field === fd.l)) errors.push({field: fd.l, msg: 'Obrigatório'});
  }
  if (kind === 'labor') {
    if (rec.cpf) { const d = digits(rec.cpf); if (d && d.length < 11 && /^\d+$/.test(rec.cpf)) rec.cpf = d.padStart(11, '0'); }
    if (isZeroMat(rec.matricula) && !digits(rec.cpf) && !rec.nome) errors.push({field: 'Matrícula', msg: 'PJ (matrícula zero) exige CPF ou Nome'});
    if (rec.estado) rec.estado = rec.estado.toUpperCase();
  } else if (rec.placa) rec.placa = rec.placa.toUpperCase();
  if (rec[K.start] && rec[K.end] && rec[K.end] < rec[K.start]) errors.push({field: KINDS[kind].fields.find(x => x.k === K.end).l, msg: 'Anterior à data inicial'});
  rec.obra = rec.obra ? canonObra(rec.obra) : '';
  return {rec, errors, warns};
}
/** Ao criar a próxima fotografia semanal do mesmo vínculo, preserva a data inicial original */
function previousSnapshot(kind, rec){
  const vk = vincKey(kind, rec); let best = null;
  for (const r of Repo.data[kind]) if (r.semana < rec.semana && vincKey(kind, r) === vk && (!best || r.semana > best.semana)) best = r;
  return best;
}
function inheritStart(kind, rec){
  const K = KINDS[kind]; const prev = previousSnapshot(kind, rec);
  if (prev && prev[K.start] && !prev[K.end] && prev[K.start] !== rec[K.start]) { const old = rec[K.start]; rec[K.start] = prev[K.start]; return old ? `${K.startLabel} original preservada (${fmtBR(prev[K.start])}; informado ${fmtBR(old)})` : `${K.startLabel} herdada da semana ${prev.semana}`; }
  return '';
}
function openRecordForm(kind, existing){
  const K = KINDS[kind]; const r = existing || {};
  const body = `<div class="form-grid">${K.fields.map(fd => {
    const v = r[fd.k] ?? '';
    let input;
    if (fd.t === 'date') input = `<input type="date" data-k="${fd.k}" value="${esc(v)}">`;
    else if (fd.t === 'week') input = weekInputHTML(fd.k, v || (existing ? '' : todayWeek()));
    else if (fd.t === 'cat') input = `<select data-k="${fd.k}"><option value="">Selecione…</option>${K.cats.map(c => `<option value="${c.k}" ${v === c.k ? 'selected' : ''}>${esc(c.full)}</option>`).join('')}</select>`;
    else if (fd.k === 'obra') input = `<input data-k="obra" value="${esc(v)}" list="dl-obras">`;
    else input = `<input data-k="${fd.k}" value="${esc(v)}" ${fd.k === 'matricula' ? 'placeholder="0 para PJ"' : ''}>`;
    return `<div class="fld" data-fd="${fd.k}"><span>${esc(fd.l)} ${fd.req ? '<span class="req">*</span>' : ''}</span>${input}<span class="err-msg"></span></div>`;
  }).join('')}</div><datalist id="dl-obras">${obraOptions(['labor', 'equip']).map(o => `<option value="${esc(o.l)}">`).join('')}</datalist>
  <p class="muted" style="font-size:12px;margin:12px 0 0">Chave: ${kind === 'labor' ? 'Semana + Obra + Matrícula (PJ com matrícula 0: Semana + Obra + CPF ou Nome)' : 'Semana + Obra + Prefixo'}.</p><div data-msg></div>`;
  const m = openModal({title: (existing ? 'Editar ' : 'Novo cadastro — ') + K.title, body, foot: `<button class="btn" data-x>Cancelar</button><button class="btn primary" data-ok>Salvar</button>`});
  bindWeekPickers(m.el);
  $('[data-x]', m.el).onclick = m.close;
  $('[data-ok]', m.el).onclick = () => {
    const raw = {}; $$('[data-k]', m.el).forEach(i => raw[i.dataset.k] = i.value);
    const {rec, errors} = validateRecord(kind, raw);
    $$('.fld', m.el).forEach(x => { x.classList.remove('bad'); $('.err-msg', x).textContent = ''; });
    errors.forEach(er => { const fd = K.fields.find(x => x.l === er.field); const box = fd && $(`[data-fd="${fd.k}"]`, m.el); if (box) { box.classList.add('bad'); $('.err-msg', box).textContent = er.msg; } });
    if (errors.length) return;
    const key = recKey(kind, rec);
    const clash = Repo.data[kind].find(x => x.id !== r.id && recKey(kind, x) === key);
    if (clash) { $('[data-msg]', m.el).innerHTML = `<div class="alert err" style="margin-top:10px">Bloqueado: já existe cadastro com esta chave (${esc(describeKey(kind, clash))}). Edite o cadastro existente — não é permitido duplicar nem sobrescrever outra semana.</div>`; return; }
    let info = '';
    const res = Repo.tx(() => {
      if (existing) { const cur = Repo.data[kind].find(x => x.id === r.id); Object.assign(cur, rec, {updatedAt: new Date().toISOString()}); }
      else { info = inheritStart(kind, rec); Repo.data[kind].push({id: uid(), ...rec, createdAt: new Date().toISOString()}); }
    });
    if (res.ok) { toast('Cadastro salvo.' + (info ? ' ' + info + '.' : ''), 'ok'); m.close(); refreshAll(); }
  };
}
async function deleteRecord(kind, id){
  const r = Repo.data[kind].find(x => x.id === id); if (!r) return;
  if (!(await confirmBox('Excluir cadastro', `<p>Excluir <b>${esc(describeKey(kind, r))}</b>?</p><p class="muted">Somente esta fotografia semanal será excluída; as demais semanas são preservadas.</p>`, 'Excluir', true))) return;
  if (Repo.tx(() => { Repo.data[kind] = Repo.data[kind].filter(x => x.id !== id); }).ok) toast('Cadastro excluído.'); refreshAll();
}

/* ---------- importação ---------- */
const GS_URL = 'https://docs.google.com/spreadsheets/d/e/2PACX-1vT1_F5MtPLrJTkP6All5p6W0rkPNkKG-D_aiX_zfrUSmJ07QguGEMH6xDBS-2-JyyaLHvaA4-7yNxNV/pubhtml';
const GS_URL_OLD = ['https://docs.google.com/spreadsheets/d/1UULPop5fY9MYWFMIJt5YWGMJdKvTBbvu9R2qDIJjiCc'];
function sheetGrid(ws){
  if (!ws || !ws['!ref']) return [];
  const rg = XLSX.utils.decode_range(ws['!ref']); const out = [];
  for (let R = rg.s.r; R <= rg.e.r; R++) { const row = []; for (let C = rg.s.c; C <= rg.e.c; C++) { const c = ws[XLSX.utils.encode_cell({r: R, c: C})]; row.push(c ? {t: c.t, v: c.v, w: c.w} : null); } out.push({n: R + 1, cells: row}); }
  return out;
}
function mapHeaders(kind, grid){
  const K = KINDS[kind]; let best = null;
  for (const row of grid.slice(0, 15)) {
    const map = {}; let hits = 0;
    row.cells.forEach((c, i) => { if (!c) return; const h = normHeader(c.w ?? c.v); const fd = K.fields.find(fd => fd.al.includes(h)); if (fd && map[fd.k] == null) { map[fd.k] = i; hits++; } });
    if (hits >= 2 && (!best || hits > best.hits)) best = {row: row.n, map, hits};
  }
  return best;
}
/** Detecta se as datas em texto do arquivo estão como dia/mês ou mês/dia (ex.: CSV do Google em inglês) */
function detectDateOrder(kind, grid, hdr){
  let dmy = 0, mdy = 0; const cols = KINDS[kind].fields.filter(f => f.t === 'date').map(f => hdr.map[f.k]).filter(i => i != null);
  for (const row of grid) { if (row.n <= hdr.row) continue; for (const i of cols) { const c = row.cells[i]; if (!c || c.t === 'n') continue; const m = String(c.w ?? c.v ?? '').trim().match(/^(\d{1,2})[\/.\-](\d{1,2})[\/.\-]\d{2,4}/); if (!m) continue; if (+m[1] > 12) dmy++; if (+m[2] > 12) mdy++; } }
  return mdy && !dmy ? 'mdy' : dmy && !mdy ? 'dmy' : null;
}
/** Formato de data do arquivo inteiro (usa as duas abas como evidência) */
function detectWbOrder(wb){
  let o = null;
  for (const k of ['labor', 'equip']) { const sh = guessSheet(k, wb); if (!sh) continue; const g = sheetGrid(wb.Sheets[sh]); const h = mapHeaders(k, g); if (!h) continue; const d = detectDateOrder(k, g, h); if (d && o && d !== o) return null; o = d || o; }
  return o;
}
function cellValue(fd, c){
  if (!c) return '';
  if (fd.t === 'date' || fd.t === 'week') { if (c.t === 'n') return c.v; if (c.t === 'd') return ymdUTC(c.v); return c.w ?? String(c.v ?? ''); }
  if (c.t === 'n') { if (fd.id && Number.isInteger(c.v) && !(c.w && /^0\d/.test(c.w))) return String(c.v); return c.w != null && c.w !== 'General' ? c.w : String(c.v); }
  return c.w ?? String(c.v ?? '');
}
function analyzeImport(kind, grid, hdr, weekOverride, forceWeek, order){
  const K = KINDS[kind]; const out = []; const seenKey = new Map();
  const existing = new Map(Repo.data[kind].map(r => [recKey(kind, r), r]));
  const hasWeekCol = hdr.map.semana != null;
  for (const row of grid) {
    if (row.n <= hdr.row) continue;
    const raw = {};
    for (const fd of K.fields) { const i = hdr.map[fd.k]; raw[fd.k] = i == null ? '' : cellValue(fd, row.cells[i]); }
    if (K.fields.every(fd => isEmptyVal(raw[fd.k]))) continue;
    if (forceWeek || !hasWeekCol || isEmptyVal(raw.semana)) raw.semana = weekOverride || raw.semana;
    const {rec, errors, warns} = validateRecord(kind, raw, {order});
    const res = {line: row.n, rec, errors, warns, existing: null, changed: false};
    if (!errors.length) {
      const key = recKey(kind, rec);
      if (seenKey.has(key)) errors.push({field: kind === 'labor' ? 'Matrícula/CPF/Nome' : 'Prefixo', msg: `Duplicado na planilha (mesma chave da linha ${seenKey.get(key)})`});
      else { seenKey.set(key, row.n); res.existing = existing.get(key) || null; if (res.existing) res.changed = K.fields.some(f => f.k !== K.start && clean(res.existing[f.k]) !== clean(rec[f.k])) || (!res.existing[K.start] && !!rec[K.start]); }
    }
    out.push(res);
  }
  return out;
}
/** Lançamentos sem detalhamento que esconderiam os registros importados (mesma obra e semana) */
function manualConflicts(kind, results){
  return [];   // lançamentos manuais apenas complementam; não há mais conflito com importação
  const mm = manualMap(kind); const out = new Map();
  for (const r of results) { if (r.errors.length) continue; const k = r.rec.semana + '|' + obraKey(r.rec.obra); if (mm.has(k)) { const g = mm.get(k); if (!out.has(k)) out.set(k, {src: g.src, n: 0}); out.get(k).n++; } }
  return [...out.values()];
}
/** Remove só o grupo (mão de obra OU equipamentos) do lançamento; apaga o lançamento se nada sobrar */
function dropManualGroup(kind, id){
  const m = Repo.data.manual.find(x => x.id === id); if (!m) return;
  (kind === 'labor' ? ['MOD', 'MOI', 'TER', 'moTotal'] : ['LA', 'LB', 'PP', 'eqTotal']).forEach(k => m[k] = null);
  if (!manualGroup(m, 'labor') && !manualGroup(m, 'equip')) Repo.data.manual = Repo.data.manual.filter(x => x.id !== id);
}
function commitImport(kind, results, dropManual, removeIds){
  const K = KINDS[kind]; let created = 0, updated = 0, dropped = 0; const notes = [];
  if (dropManual) for (const c of manualConflicts(kind, results)) { dropManualGroup(kind, c.src.id); dropped++; }
  let removed = 0; if (removeIds && removeIds.size) { const n0 = Repo.data[kind].length; Repo.data[kind] = Repo.data[kind].filter(x => !removeIds.has(x.id)); removed = n0 - Repo.data[kind].length; }
  const idx = new Map(Repo.data[kind].map(r => [recKey(kind, r), r]));
  const now = new Date().toISOString();
  const valid = results.filter(r => !r.errors.length).sort((a, b) => a.rec.semana.localeCompare(b.rec.semana));
  for (const r of valid) {
    const rec = {...r.rec}; rec.obra = canonObra(rec.obra);
    const key = recKey(kind, rec); const ex = idx.get(key);
    if (ex) { if (r.existing && !r.changed) continue; const keep = ex[K.start] || rec[K.start]; Object.assign(ex, rec, {[K.start]: keep, updatedAt: now}); updated++; }
    else { const n = inheritStart(kind, rec); if (n) notes.push(n); const nr = {id: uid(), ...rec, createdAt: now}; Repo.data[kind].push(nr); idx.set(key, nr); created++; }
  }
  return {created, updated, dropped, removed, preserved: notes.length};
}
const LS_SHEET = 'cosampa_histogramas_planilha';
function sheetCfg(){ const d = {url: GS_URL, gidL: '', gidE: '', auto: true}; try { const c = {...d, ...JSON.parse(localStorage.getItem(LS_SHEET) || '{}')}; if (!c.url || GS_URL_OLD.some(u => c.url.startsWith(u))) { c.url = GS_URL; c.gidL = ''; c.gidE = ''; } return c; } catch (e) { return d; } }
function saveSheetCfg(c){ try { localStorage.setItem(LS_SHEET, JSON.stringify(c)); } catch (e) {} }
/** Lê a planilha publicada: tenta o arquivo inteiro (.xlsx, todas as abas); se falhar, lê cada aba pelo gid (CSV). Somente leitura. */
const SHARE_HINT = 'Verifique se a planilha está compartilhada como “Qualquer pessoa com o link — Leitor” (ou publicada na Web).';
function isHtml(txt){ return /^\s*<(!doctype|html)/i.test(txt); }
/** Lê a planilha do Google (somente leitura).
 *  • Link de edição/compartilhamento (…/spreadsheets/d/ID/…): exige “Qualquer pessoa com o link — Leitor”. Lê as abas pelo NOME (MO e Equipamento).
 *  • Link publicado (…/d/e/…/pubhtml): lê o arquivo inteiro; se falhar, cada aba pelo gid. */
async function fetchPublishedWorkbook(raw, gids){
  raw = String(raw || '').trim();
  const pub = raw.match(/docs\.google\.com\/spreadsheets\/d\/e\/[\w-]+/);
  const idm = !pub && raw.match(/docs\.google\.com\/spreadsheets\/d\/([\w-]{20,})/);
  if (!pub && !idm) throw new Error('link inválido — cole o link da planilha do Google Sheets');
  const getXlsx = async url => { const r = await fetch(url, {cache: 'no-store'}); if (!r.ok) throw new Error('HTTP ' + r.status); const buf = await r.arrayBuffer(); const head = new TextDecoder().decode(buf.slice(0, 64)); if (isHtml(head)) throw new Error('acesso negado'); return XLSX.read(buf, {type: 'array', cellDates: false, cellText: true}); };
  const getCsv = async (url, label) => { let r; try { r = await fetch(url, {cache: 'no-store'}); } catch (e) { throw new Error(`aba ${label}: acesso bloqueado. ${SHARE_HINT}`); } if (r.status === 401 || r.status === 403) throw new Error(`aba ${label}: acesso negado (planilha privada). ${SHARE_HINT}`); if (!r.ok) throw new Error(`aba ${label}: HTTP ${r.status}`); const txt = await r.text(); if (isHtml(txt)) throw new Error(`aba ${label}: acesso negado. ${SHARE_HINT}`); return txt; };
  const addCsv = (wb, txt, name) => { const w2 = XLSX.read(txt, {type: 'string', raw: true}); XLSX.utils.book_append_sheet(wb, w2.Sheets[w2.SheetNames[0]], name); };
  if (idm) {
    const id = idm[1], base = `https://docs.google.com/spreadsheets/d/${id}`;
    try { return await getXlsx(`${base}/export?format=xlsx`); }
    catch (e) {
      const wb = XLSX.utils.book_new(); let n = 0, lastErr = null; const seenTxt = new Set();
      for (const k of ['labor', 'equip', 'manual']) {
        const gid = (gids || {})[k];
        const url = gid ? `${base}/gviz/tq?tqx=out:csv&gid=${encodeURIComponent(gid)}` : `${base}/gviz/tq?tqx=out:csv&sheet=${encodeURIComponent(IMPORT_SHEET[k])}`;
        try { const txt = await getCsv(url, IMPORT_SHEET[k]); if (seenTxt.has(txt)) continue; /* aba inexistente: o Google devolve a 1ª aba */ seenTxt.add(txt); addCsv(wb, txt, IMPORT_SHEET[k]); n++; } catch (er) { if (k !== 'manual') lastErr = er; }
      }
      for (const k of TRK_ON) { try { const txt = await getCsv(`${base}/gviz/tq?tqx=out:csv&sheet=${encodeURIComponent(TRACK[k].sheet)}`, TRACK[k].sheet); if (seenTxt.has(txt)) continue; seenTxt.add(txt); addCsv(wb, txt, TRACK[k].sheet); } catch (er) {} }
      if (!n) throw lastErr || new Error(SHARE_HINT);
      return wb;
    }
  }
  const base = raw.replace(/\/pub(html)?(\?.*)?(#.*)?$/, '/pub');
  try { return await getXlsx(base + '?output=xlsx'); }
  catch (e) {
    const wb = XLSX.utils.book_new(); let n = 0;
    for (const [k, gid] of Object.entries(gids || {})) { if (!gid) continue; addCsv(wb, await getCsv(`${base}?output=csv&single=true&gid=${encodeURIComponent(gid)}`, IMPORT_SHEET[k]), IMPORT_SHEET[k]); n++; }
    if (!n) throw new Error(e.message + '. ' + SHARE_HINT);
    return wb;
  }
}
/* Planilha única com duas abas: “MO” → Mão de obra · “Equipamento” → Equipamentos */
const IMPORT_SHEET = {labor: 'MO', equip: 'Equipamento', manual: 'Resumo'};
/* Aba “Resumo”: totais por obra e semana (lançamentos sem detalhamento / complementos) */
const RESUMO_COLS = [
  {k: 'obra', l: 'Obra', al: ['obra', 'obras', 'empreendimento']},
  {k: 'semana', l: 'Semana', al: ['semana', 'semanaiso', 'sem', 'week']},
  {k: 'MOD', l: 'MOD', al: ['mod', 'direta', 'direto', 'maodeobradireta']},
  {k: 'MOI', l: 'MOI', al: ['moi', 'indireta', 'indireto', 'maodeobraindireta']},
  {k: 'TER', l: 'Terceiros', al: ['terceiros', 'terceiro', 'terceirizados', 'terceirizado', 'ter']},
  {k: 'moTotal', l: 'Total MO', al: ['totalmo', 'totalmaodeobra', 'mototal', 'totaldemaodeobra']},
  {k: 'LA', l: 'Linha amarela', al: ['linhaamarela', 'la']},
  {k: 'LB', l: 'Linha branca', al: ['linhabranca', 'lb']},
  {k: 'PP', l: 'P. porte', al: ['pporte', 'pequenoporte', 'pp']},
  {k: 'eqTotal', l: 'Total equipamentos', al: ['totalequipamentos', 'totalequip', 'totalequipamento', 'eqtotal', 'totaldeequipamentos']},
];
const NUMK = ['MOD', 'MOI', 'TER', 'moTotal', 'LA', 'LB', 'PP', 'eqTotal'];
function mapResumoHeaders(grid){
  let best = null;
  for (const row of grid.slice(0, 15)) {
    const map = {}; let hits = 0, nums = 0;
    row.cells.forEach((c, i) => { if (!c) return; const h = normHeader(c.w ?? c.v); const fd = RESUMO_COLS.find(x => x.al.includes(h)); if (fd && map[fd.k] == null) { map[fd.k] = i; hits++; if (NUMK.includes(fd.k)) nums++; } });
    if (map.obra != null && nums >= 1 && (!best || hits > best.hits)) best = {row: row.n, map, hits};
  }
  return best;
}
function parseCount(raw){
  if (raw == null) return {v: null};
  if (typeof raw === 'number') return Number.isInteger(raw) && raw >= 0 ? {v: raw} : {v: null, err: `Quantidade inválida "${raw}"`};
  const t = String(raw).trim(); if (!t || DASHES.has(t)) return {v: null};
  const n = Number(t.replace(/\s/g, '').replace(/\.(?=\d{3}(\D|$))/g, '').replace(',', '.'));
  return Number.isInteger(n) && n >= 0 ? {v: n} : {v: null, err: `Quantidade inválida "${t}"`};
}
function analyzeResumo(grid, hdr, weekOverride, forceWeek){
  const out = []; const seen = new Map();
  const existing = new Map(Repo.data.manual.map(m => [m.semana + '|' + obraKey(m.obra), m]));
  const cell = (row, k) => { const i = hdr.map[k]; if (i == null) return ''; const c = row.cells[i]; if (!c) return ''; return c.t === 'n' ? c.v : (c.w ?? String(c.v ?? '')); };
  for (const row of grid) {
    if (row.n <= hdr.row) continue;
    const obraRaw = clean(cell(row, 'obra')); let wkRaw = cell(row, 'semana');
    const nums = {}; const errors = []; const warns = [];
    for (const k of NUMK) { const r = parseCount(cell(row, k)); nums[k] = r.v; if (r.err) errors.push({field: RESUMO_COLS.find(x => x.k === k).l, msg: r.err}); }
    if (!obraRaw && NUMK.every(k => nums[k] == null) && isEmptyVal(wkRaw)) continue;          // linha vazia
    if (forceWeek || hdr.map.semana == null || isEmptyVal(wkRaw)) wkRaw = weekOverride || wkRaw;
    const w = parseWeek(wkRaw);
    if (!obraRaw) errors.push({field: 'Obra', msg: 'Obrigatório'});
    if (!w.v) errors.push({field: 'Semana', msg: w.err || 'Obrigatório'});
    if (NUMK.every(k => nums[k] == null) && obraRaw) errors.push({field: 'Quantitativos', msg: 'Nenhum valor informado'});
    // categorias + total informados e diferentes: vale a soma das categorias
    [['MOD', 'MOI', 'TER', 'moTotal', 'mão de obra'], ['LA', 'LB', 'PP', 'eqTotal', 'equipamentos']].forEach(([a, b, c, t, nm]) => {
      const cats = [a, b, c].filter(k => nums[k] != null);
      if (cats.length && nums[t] != null) { const sm = [a, b, c].reduce((x, k) => x + (nums[k] || 0), 0); if (sm !== nums[t]) warns.push(`Total de ${nm} (${nums[t]}) diferente da soma das categorias (${sm}) — vale a soma`); nums[t] = null; }
    });
    const rec = {obra: obraRaw ? canonObra(obraRaw) : '', semana: w.v, ...nums};
    const res = {line: row.n, rec, errors, warns, existing: null, changed: false};
    if (!errors.length) {
      const key = rec.semana + '|' + obraKey(rec.obra);
      if (seen.has(key)) errors.push({field: 'Obra/Semana', msg: `Duplicado na planilha (mesma obra e semana da linha ${seen.get(key)})`});
      else { seen.set(key, row.n); res.existing = existing.get(key) || null; if (res.existing) res.changed = NUMK.some(k => rec[k] != null && (res.existing[k] ?? null) !== rec[k]); }   // célula vazia não altera o que já existe
    }
    out.push(res);
  }
  return out;
}
function commitResumo(results){
  let created = 0, updated = 0; const now = new Date().toISOString();
  const idx = new Map(Repo.data.manual.map(m => [m.semana + '|' + obraKey(m.obra), m]));
  for (const r of results) {
    if (r.errors.length) continue;
    const key = r.rec.semana + '|' + obraKey(r.rec.obra); const ex = idx.get(key);
    if (ex) { if (!r.changed) continue; NUMK.forEach(k => { if (r.rec[k] != null) ex[k] = r.rec[k]; }); ex.updatedAt = now; updated++; }   // só substitui o que foi preenchido
    else { const n = {id: uid(), obra: r.rec.obra, semana: r.rec.semana, createdAt: now}; NUMK.forEach(k => n[k] = r.rec[k] ?? null); Repo.data.manual.push(n); idx.set(key, n); created++; }
  }
  return {created, updated};
}
const DISTINCT = {labor: ['nome', 'matricula', 'tipo', 'admissao'], equip: ['prefixo', 'classificacao', 'mobilizacao']};
function sheetNameMatch(kind, n){ const h = normHeader(n); return kind === 'labor' ? (h === 'mo' || /^mo\d|maodeobra|efetivo|colaborador/.test(h)) : kind === 'manual' ? /resumo|totais/.test(h) : /equip/.test(h); }
function guessSheet(kind, wb){
  const byName = wb.SheetNames.find(n => sheetNameMatch(kind, n)); if (byName) return byName;
  if (kind === 'manual') return wb.SheetNames.find(n => !/instru/i.test(n) && mapResumoHeaders(sheetGrid(wb.Sheets[n]))) || '';
  let best = '', bestHits = 0;
  for (const n of wb.SheetNames) { if (/instru/i.test(n)) continue; const h = mapHeaders(kind, sheetGrid(wb.Sheets[n])); if (h && DISTINCT[kind].some(k => h.map[k] != null) && h.hits > bestHits) { best = n; bestHits = h.hits; } }
  return best;
}
function openImport(startKind, preWb, preName){
  if (!libsOk()) return toast('Bibliotecas não carregadas (verifique a conexão com a internet).', 'err');
  const KS = ['labor', 'equip'];
  const st = {wb: null, source: '', part: {labor: {sheet: '', results: [], hdr: null}, equip: {sheet: '', results: [], hdr: null}, manual: {sheet: '', results: [], hdr: null}}};
  const body = `
  <div class="alert info" style="font-size:12px">A planilha de importação tem <b>três abas</b>: <b>MO</b> → <b>Mão de obra</b> · <b>Equipamento</b> → <b>Equipamentos</b> · <b>Resumo</b> → totais por obra e semana (lançamentos sem detalhamento / complementos, ex.: obras sem cadastro individual). As três são lidas e importadas juntas. Baixe o <b>Modelo Excel</b> para ver o formato.</div>
  <div class="cols2">
    <div><div class="drop" data-drop><b>Arraste o arquivo aqui</b><br>ou clique para escolher (.xlsx, .xls, .csv)<input type="file" accept=".xlsx,.xls,.xlsm,.csv,.txt" hidden data-file></div></div>
    <div class="card" style="margin:0;box-shadow:none"><div style="font-weight:700;margin-bottom:6px">Planilha do Google Sheets — somente leitura</div>
      <div class="fld"><span>Link da planilha</span><input data-gsurl value="${esc(sheetCfg().url)}"></div>
      <div style="display:flex;gap:8px;align-items:end;margin-top:8px;flex-wrap:wrap">
        <div class="fld" style="flex:1;min-width:110px"><span>gid aba MO (opcional)</span><input data-gid-labor value="${esc(sheetCfg().gidL)}"></div>
        <div class="fld" style="flex:1;min-width:110px"><span>gid aba Equipamento</span><input data-gid-equip value="${esc(sheetCfg().gidE)}" placeholder="opcional"></div>
        <button class="btn dark" data-gs>Ler planilha</button></div>
      <div class="muted" style="font-size:11px;margin-top:6px">Lê as abas MO, Equipamento e Resumo pelo nome. A planilha precisa estar compartilhada como “Qualquer pessoa com o link — Leitor” (ou publicada). Nada é gravado no Google Sheets. Se o navegador bloquear, baixe a planilha em Excel (.xlsx) e importe pelo arquivo.</div>
    </div>
  </div>
  <div data-gserr></div>
  <div data-step2 hidden style="margin-top:14px">
    <div class="form-grid" style="grid-template-columns:repeat(auto-fit,minmax(240px,1fr))">
      <div class="fld"><span>Semana para linhas sem semana (vale para todas as abas)</span>${weekInputHTML('wk', '')}</div>
      <label class="fld" style="flex-direction:row;align-items:center;gap:8px;text-transform:none;font-size:13px;color:var(--tx)"><input type="checkbox" data-force style="width:auto;min-height:0"> Aplicar esta semana a todas as linhas</label>
      <div class="fld"><span>Formato das datas em texto</span><select data-dord><option value="auto">Automático (detecta pelo arquivo)</option><option value="dmy">dia/mês/ano (dd/mm/aaaa)</option><option value="mdy">mês/dia/ano (mm/dd/aaaa — Google em inglês)</option></select></div>
    </div>
    <div data-wkerr></div>
    ${KS.map(k => `<div class="card" data-part="${k}" style="margin-top:12px;box-shadow:none;border-color:var(--line)">
      <div class="card-h" style="margin-bottom:6px"><h3>${k === 'labor' ? 'Aba MO → Mão de obra' : 'Aba Equipamento → Equipamentos'}</h3><span class="sp"></span>
        <div class="fld" style="flex-direction:row;align-items:center;gap:8px"><span>Aba do arquivo</span><select data-sheet="${k}" style="min-width:180px"></select></div></div>
      <div data-sum="${k}"></div></div>`).join('')}
    <div class="card" data-part="manual" style="margin-top:12px;box-shadow:none;border-color:var(--line)">
      <div class="card-h" style="margin-bottom:6px"><h3>Aba Resumo → Totais por obra e semana</h3><span class="sp"></span>
        <div class="fld" style="flex-direction:row;align-items:center;gap:8px"><span>Aba do arquivo</span><select data-sheet="manual" style="min-width:180px"></select></div></div>
      <div data-sum="manual"></div></div>
  </div>`;
  const m = openModal({title: 'Importar planilha (MO + Equipamento + Resumo)', body, wide: true, foot: `<label data-confirm-wrap hidden style="margin-right:auto;display:flex;gap:6px;align-items:center;font-size:13px"><input type="checkbox" data-confirm> Confirmo a atualização dos cadastros existentes listados</label><button class="btn" data-x>Cancelar</button><button class="btn primary" data-ok disabled>Importar</button>`});
  const el = m.el; bindWeekPickers(el);
  $('[data-x]', el).onclick = m.close;
  const drop = $('[data-drop]', el), fileIn = $('[data-file]', el);
  drop.onclick = () => fileIn.click();
  drop.ondragover = e => { e.preventDefault(); drop.classList.add('over'); };
  drop.ondragleave = () => drop.classList.remove('over');
  drop.ondrop = e => { e.preventDefault(); drop.classList.remove('over'); if (e.dataTransfer.files[0]) loadFile(e.dataTransfer.files[0]); };
  fileIn.onchange = () => fileIn.files[0] && loadFile(fileIn.files[0]);
  async function loadFile(file){
    try {
      const isText = /\.(csv|txt)$/i.test(file.name);
      const wb = isText ? XLSX.read(await file.text(), {type: 'string', raw: true}) : XLSX.read(await file.arrayBuffer(), {type: 'array', cellDates: false, cellText: true});
      setWb(wb, file.name);
    } catch (e) { toast('Não foi possível ler o arquivo: ' + e.message, 'err'); }
  }
  $('[data-gs]', el).onclick = async () => {
    const b = $('[data-gs]', el); b.disabled = true; b.textContent = 'Lendo…'; $('[data-gserr]', el).innerHTML = '';
    try { setWb(await fetchPublishedWorkbook($('[data-gsurl]', el).value, {labor: $('[data-gid-labor]', el).value.trim(), equip: $('[data-gid-equip]', el).value.trim()}), 'Google Sheets publicado'); }
    catch (e) { $('[data-gserr]', el).innerHTML = `<div class="alert err" style="margin-top:10px"><b>O navegador não conseguiu ler a planilha publicada</b> (${esc(e.message)}). Isso costuma ocorrer por bloqueio de acesso entre sites (CORS) ou gid incorreto.<br>Alternativa: abra a planilha, use <i>Arquivo → Fazer download → Microsoft Excel (.xlsx)</i> e importe o arquivo aqui — as abas MO e Equipamento são lidas juntas.</div>`; }
    finally { b.disabled = false; b.textContent = 'Ler planilha'; }
  };
  function setWb(wb, name){
    st.wb = wb; st.source = name; Sync.lastWb = wb; Sync.lastSrc = name; Sync.lastAt = new Date();
    for (const k of [...KS, 'manual']) {
      const sel = $(`[data-sheet="${k}"]`, el);
      sel.innerHTML = `<option value="">(não importar)</option>` + wb.SheetNames.map(n => `<option>${esc(n)}</option>`).join('');
      sel.value = guessSheet(k, wb);
    }
    // arquivo de uma aba só (CSV) sem nome reconhecível: usa a aba do app de onde o usuário abriu
    if (wb.SheetNames.length === 1 && !$('[data-sheet="labor"]', el).value && !$('[data-sheet="equip"]', el).value) $(`[data-sheet="${startKind}"]`, el).value = wb.SheetNames[0];
    $('[data-step2]', el).hidden = false; analyzeAll();
  }
  [...KS, 'manual'].forEach(k => $(`[data-sheet="${k}"]`, el).onchange = analyzeAll);
  el.addEventListener('input', e => { if (e.target.dataset.k === 'wk') analyzeAll(); });
  $('[data-force]', el).onchange = analyzeAll;
  $('[data-dord]', el).onchange = analyzeAll;
  $('[data-confirm]', el).onchange = updateBtn;
  el.addEventListener('change', e => { if (e.target.dataset.mirror) updateBtn(); });
  function analyzeAll(){
    if (!st.wb) return;
    const wkRaw = $('[data-k=wk]', el).value.trim(); const wk = wkRaw ? parseWeek(wkRaw) : {v: ''};
    $('[data-wkerr]', el).innerHTML = wkRaw && wk.err ? `<div class="alert err">${esc(wk.err)}</div>` : '';
    const sameSheet = $('[data-sheet="labor"]', el).value && $('[data-sheet="labor"]', el).value === $('[data-sheet="equip"]', el).value;
    for (const k of KS) analyzePart(k, wk.v, sameSheet);
    analyzeResumoPart(wk.v);
    const anyUpd = [...KS, 'manual'].some(k => st.part[k].results.some(r => !r.errors.length && r.changed));
    $('[data-confirm-wrap]', el).hidden = !anyUpd; $('[data-confirm]', el).checked = false;
    updateBtn();
  }
  function analyzePart(kind, wkV, sameSheet){
    const K = KINDS[kind]; const P = st.part[kind]; const sum = $(`[data-sum="${kind}"]`, el);
    P.sheet = $(`[data-sheet="${kind}"]`, el).value; P.results = []; P.hdr = null;
    if (!P.sheet) { sum.innerHTML = `<div class="muted" style="font-size:13px">Nenhuma aba selecionada — ${K.title.toLowerCase()} não serão importados${st.wb.SheetNames.length > 1 ? ` (esperada a aba “${IMPORT_SHEET[kind]}”)` : ''}.</div>`; return; }
    if (sameSheet) { sum.innerHTML = `<div class="alert err">A mesma aba foi escolhida para Mão de obra e Equipamentos. Selecione abas diferentes.</div>`; return; }
    const grid = sheetGrid(st.wb.Sheets[P.sheet]); P.hdr = mapHeaders(kind, grid);
    if (!P.hdr || !DISTINCT[kind].some(k => P.hdr.map[k] != null)) { P.hdr = null; sum.innerHTML = `<div class="alert err">Cabeçalhos de ${K.title.toLowerCase()} não reconhecidos na aba “${esc(P.sheet)}”. Colunas esperadas: ${K.fields.map(f => f.l).join(', ')}.</div>`; return; }
    const dsel = $('[data-dord]', el).value; const det = detectDateOrder(kind, grid, P.hdr) || detectWbOrder(st.wb); P.order = dsel === 'auto' ? (det || 'dmy') : dsel;
    P.results = analyzeImport(kind, grid, P.hdr, wkV, $('[data-force]', el).checked, P.order);
    const valid = P.results.filter(r => !r.errors.length), bad = P.results.filter(r => r.errors.length), upd = valid.filter(r => r.changed), same = valid.filter(r => r.existing && !r.changed), news = valid.filter(r => !r.existing);
    // cadastros do app (mesmas obras/semanas da planilha) que não constam mais na planilha
    const pairs = new Set(valid.map(r => r.rec.semana + '|' + obraKey(r.rec.obra))), keys = new Set(valid.map(r => recKey(kind, r.rec)));
    P.missing = Repo.data[kind].filter(x => pairs.has(x.semana + '|' + obraKey(x.obra)) && !keys.has(recKey(kind, x)));
    const missing = K.fields.filter(fd => P.hdr.map[fd.k] == null).map(fd => fd.l);
    const dashN = P.results.reduce((s2, r) => s2 + r.warns.filter(w => w.includes('"-"')).length, 0);
    const errs = bad.flatMap(r => r.errors.map(e => `<li><b>Linha ${r.line}</b> — ${esc(e.field)}: ${esc(e.msg)}</li>`));
    const cols = K.fields;
    sum.innerHTML = `
      <div class="stat-row"><span class="chip info">${esc(st.source)} · aba “${esc(P.sheet)}”</span><span class="chip">Cabeçalho na linha ${P.hdr.row}</span><span class="chip">${P.results.length} linha(s)</span><span class="chip ok">${valid.length} válida(s)</span><span class="chip err">${bad.length} pendente(s)</span><span class="chip or">${news.length} nova(s)</span><span class="chip warn">${upd.length} alterada(s)</span><span class="chip">${same.length} sem alteração</span><span class="chip info">Datas: ${P.order === 'mdy' ? 'mês/dia/ano' : 'dia/mês/ano'}${dsel === 'auto' ? (det ? ' (detectado)' : ' (padrão)') : ''}</span></div>
      ${P.missing.length ? `<label class="alert warn" style="display:flex;gap:8px;align-items:center;font-size:12.5px"><input type="checkbox" data-mirror="${kind}" style="width:auto;min-height:0"> <span><b>${P.missing.length} cadastro(s) do app não constam mais na planilha</b> nas mesmas obras/semanas (ex.: ${esc(P.missing.slice(0, 3).map(x => describeKey(kind, x)).join('; '))}). Marque para excluí-los e deixar o app igual à planilha.</span></label>` : ''}
      ${P.hdr.map.semana == null ? `<div class="alert warn">A aba não tem coluna <b>Semana</b>. Informe a semana acima.</div>` : ''}
      ${missing.length ? `<div class="alert info" style="font-size:12px">Colunas não encontradas (ficarão vazias): ${esc(missing.join(', '))}</div>` : ''}
      ${dashN ? `<div class="alert info" style="font-size:12px">${dashN} data(s) com “-” não foram aceitas como data e serão gravadas como não informadas.</div>` : ''}
      ${upd.length ? `<div class="alert warn"><b>${upd.length} cadastro(s) existente(s) na mesma obra e semana serão atualizados</b> (todos os campos, inclusive vazios; a ${K.startLabel.toLowerCase()} original é preservada):<ul class="upd-list" style="margin-top:6px">${upd.slice(0, 300).map(r => `<li>Linha ${r.line}: ${esc(describeKey(kind, r.existing))}</li>`).join('')}</ul></div>` : ''}
      ${errs.length ? `<div style="font-weight:700;margin:8px 0 4px">Erros por linha e campo (${errs.length})</div><ul class="err-list">${errs.slice(0, 500).join('')}</ul>` : ''}
      <details ${P.results.length <= 30 ? 'open' : ''} style="margin-top:8px"><summary style="cursor:pointer;font-weight:700">Prévia ${P.results.length > 100 ? '(primeiras 100 linhas)' : ''}</summary>
      <div class="tbl-wrap" style="max-height:300px;margin-top:6px"><table><thead><tr><th>Linha</th><th>Status</th>${cols.map(c => `<th>${esc(c.l)}</th>`).join('')}</tr></thead><tbody>
      ${P.results.slice(0, 100).map(r => `<tr><td>${r.line}</td><td>${r.errors.length ? '<span class="chip err">Pendente</span>' : r.changed ? '<span class="chip warn">Alterar</span>' : r.existing ? '<span class="chip">Sem alteração</span>' : '<span class="chip ok">Novo</span>'}</td>${cols.map(c => `<td>${esc(displayVal(kind, r.rec, c))}</td>`).join('')}</tr>`).join('')}
      </tbody></table></div></details>`;
  }
  function analyzeResumoPart(wkV){
    const P = st.part.manual; const sum = $('[data-sum="manual"]', el);
    P.sheet = $('[data-sheet="manual"]', el).value; P.results = []; P.hdr = null;
    if (!P.sheet) { sum.innerHTML = `<div class="muted" style="font-size:13px">Nenhuma aba selecionada — totais por obra não serão importados${st.wb.SheetNames.length > 1 ? ' (esperada a aba “Resumo”)' : ''}.</div>`; return; }
    if ([...KS].some(k => $(`[data-sheet="${k}"]`, el).value === P.sheet)) { sum.innerHTML = `<div class="alert err">A aba “${esc(P.sheet)}” já foi escolhida para outra parte. Selecione a aba Resumo.</div>`; return; }
    const grid = sheetGrid(st.wb.Sheets[P.sheet]); P.hdr = mapResumoHeaders(grid);
    if (!P.hdr) { sum.innerHTML = `<div class="alert err">Cabeçalhos da aba Resumo não reconhecidos em “${esc(P.sheet)}”. Colunas esperadas: ${RESUMO_COLS.map(c => c.l).join(', ')}.</div>`; return; }
    P.results = analyzeResumo(grid, P.hdr, wkV, $('[data-force]', el).checked);
    const valid = P.results.filter(r => !r.errors.length), bad = P.results.filter(r => r.errors.length), upd = valid.filter(r => r.changed), same = valid.filter(r => r.existing && !r.changed), news = valid.filter(r => !r.existing);
    const errs = bad.flatMap(r => r.errors.map(e => `<li><b>Linha ${r.line}</b> — ${esc(e.field)}: ${esc(e.msg)}</li>`));
    const warns = P.results.flatMap(r => r.warns.map(w => `<li><b>Linha ${r.line}</b> — ${esc(w)}</li>`));
    const v = x => x == null ? '<span class="muted">—</span>' : fmtN(x);
    sum.innerHTML = `
      <div class="stat-row"><span class="chip info">${esc(st.source)} · aba “${esc(P.sheet)}”</span><span class="chip">${P.results.length} linha(s)</span><span class="chip ok">${valid.length} válida(s)</span><span class="chip err">${bad.length} pendente(s)</span><span class="chip or">${news.length} nova(s)</span><span class="chip warn">${upd.length} alterada(s)</span><span class="chip">${same.length} sem alteração</span></div>
      <div class="alert info" style="font-size:12px">Os totais <b>complementam</b> os cadastros: em cada obra/semana, as categorias que já têm cadastros individuais continuam vindo dos cadastros; os valores desta aba preenchem apenas as que não têm. Se a obra/semana já existe no app, só os valores <b>preenchidos</b> na planilha substituem os atuais — células vazias mantêm o que já existe. Lançamentos que não estão na aba Resumo continuam no app.</div>
      ${upd.length ? `<div class="alert warn"><b>${upd.length} lançamento(s) existente(s) serão atualizados:</b> ${esc(upd.slice(0, 20).map(r => r.existing.obra + ' ' + r.existing.semana).join(' · '))}</div>` : ''}
      ${warns.length ? `<ul class="upd-list">${warns.slice(0, 100).join('')}</ul>` : ''}
      ${errs.length ? `<div style="font-weight:700;margin:8px 0 4px">Erros por linha e campo (${errs.length})</div><ul class="err-list">${errs.slice(0, 300).join('')}</ul>` : ''}
      <details ${P.results.length <= 30 ? 'open' : ''} style="margin-top:8px"><summary style="cursor:pointer;font-weight:700">Prévia</summary>
      <div class="tbl-wrap" style="max-height:300px;margin-top:6px"><table><thead><tr><th>Linha</th><th>Status</th>${RESUMO_COLS.map(c => `<th>${esc(c.l)}</th>`).join('')}</tr></thead><tbody>
      ${P.results.slice(0, 100).map(r => `<tr><td>${r.line}</td><td>${r.errors.length ? '<span class="chip err">Pendente</span>' : r.changed ? '<span class="chip warn">Alterar</span>' : r.existing ? '<span class="chip">Sem alteração</span>' : '<span class="chip ok">Novo</span>'}</td><td class="l">${esc(r.rec.obra)}</td><td>${esc(r.rec.semana)}</td>${NUMK.map(k => `<td>${v(r.rec[k])}</td>`).join('')}</tr>`).join('')}
      </tbody></table></div></details>`;
  }
  function updateBtn(){
    const v = {}; KS.forEach(k => v[k] = st.part[k].results.filter(r => !r.errors.length));
    const act = k => v[k].filter(r => !r.existing || r.changed).length;
    const rm = k => { const cb = $(`[data-mirror="${k}"]`, el); return cb && cb.checked ? (st.part[k].missing || []).length : 0; };
    v.manual = st.part.manual.results.filter(r => !r.errors.length);
    const total = act('labor') + act('equip') + act('manual') + rm('labor') + rm('equip'); const upd = [...KS, 'manual'].some(k => v[k].some(r => r.changed)) || rm('labor') + rm('equip') > 0;
    const b = $('[data-ok]', el);
    b.disabled = !total || (upd && !$('[data-confirm]', el).checked);
    b.textContent = total ? `Importar: ${act('labor')} mão de obra · ${act('equip')} equipamentos · ${act('manual')} totais${rm('labor') + rm('equip') ? ` · excluir ${rm('labor') + rm('equip')}` : ''}` : ([...KS, 'manual'].some(k => v[k].length) ? 'Nada a alterar — app já está igual à planilha' : 'Importar');
  }
  $('[data-ok]', el).onclick = () => {
    Snap.take('Antes de importar planilha (' + st.source + ')');
    const res = Repo.tx(() => { const o = {}; for (const k of KS) { const cb = $(`[data-mirror="${k}"]`, el); const ids = cb && cb.checked ? new Set(st.part[k].missing.map(x => x.id)) : null; if (st.part[k].results.length) o[k] = commitImport(k, st.part[k].results, false, ids); } if (st.part.manual.results.length) o.manual = commitResumo(st.part.manual.results); return o; });
    Sync.markImported();
    if (!res.ok) return;
    const t = k => { const r = res.out[k]; return r ? `${KINDS[k].title}: ${r.created} novo(s), ${r.updated} alterado(s)${r.removed ? `, ${r.removed} excluído(s)` : ''}${r.preserved ? `, ${r.preserved} com data inicial preservada` : ''}` : `${KINDS[k].title}: não importado`; };
    const tm = res.out.manual ? `Totais (Resumo): ${res.out.manual.created} novo(s), ${res.out.manual.updated} alterado(s)` : 'Totais (Resumo): não importado';
    toast(`Importação salva — ${t('labor')} · ${t('equip')} · ${tm}.`, 'ok');
    m.close(); refreshAll();
  };
  if (preWb) setWb(preWb, preName || 'Google Sheets publicado');
}

/* ---------- totais sem detalhamento ---------- */
async function dropManualAsk(kind, id){
  const m = Repo.data.manual.find(x => x.id === id); if (!m) return;
  if (!(await confirmBox('Remover total manual', `<p>Remover o total sem detalhamento de <b>${esc(KINDS[kind].title.toLowerCase())}</b> da obra <b>${esc(m.obra)}</b> na semana <b>${esc(m.semana)}</b>?</p><p class="muted">Os cadastros individuais dessa obra/semana voltam a ser contados. ${kind === 'labor' ? 'Os totais de equipamentos' : 'Os totais de mão de obra'} desse lançamento, se houver, são mantidos.</p>`, 'Remover', true))) return;
  if (Repo.tx(() => dropManualGroup(kind, id)).ok) toast('Total manual removido — cadastros individuais voltaram a ser considerados.', 'ok');
  refreshAll();
}
function openManualForm(existing){
  const m0 = existing || {};
  const num = (k, l) => `<div class="fld"><span>${l}</span><input type="number" min="0" step="1" data-k="${k}" value="${m0[k] ?? ''}" placeholder="vazio"></div>`;
  const body = `
    <div class="form-grid">
      <div class="fld"><span>Obra <span class="req">*</span></span><input data-k="obra" list="dl-obras2" value="${esc(m0.obra || '')}"></div>
      <div class="fld"><span>Semana <span class="req">*</span></span>${weekInputHTML('semana', m0.semana || UI.resumo.week || '')}</div>
    </div><datalist id="dl-obras2">${obraOptions(['labor', 'equip']).map(o => `<option value="${esc(o.l)}">`).join('')}</datalist>
    <h4 style="margin:16px 0 6px">Mão de obra</h4>
    <div class="form-grid">${num('MOD', 'MOD')}${num('MOI', 'MOI')}${num('TER', 'Terceiros')}${num('moTotal', 'Total mão de obra')}</div>
    <h4 style="margin:16px 0 6px">Equipamentos</h4>
    <div class="form-grid">${num('LA', 'Linha amarela')}${num('LB', 'Linha branca')}${num('PP', 'P. porte')}${num('eqTotal', 'Total equipamentos')}</div>
    <div data-ref style="margin-top:12px"></div>
    <div class="alert info" style="margin-top:14px;font-size:12px">
      • Informe obra e semana: as categorias que já têm cadastros vêm <b>preenchidas e bloqueadas</b> (🔒). Preencha somente as que faltam (ex.: Terceiros).<br>
      • Os cadastros nunca são ocultados nem substituídos; o lançamento apenas complementa.<br>
      • Campo vazio = sem complemento · Zero = quantidade zero · O total é calculado automaticamente.<br>
      • Se o grupo não tiver nenhum cadastro, você pode informar só o total (categorias ficam “—”, Sem detalhamento).
    </div><div data-msg></div>`;
  let cur = existing || null;
  const md = openModal({title: existing ? 'Editar complemento de totais' : 'Lançar totais sem detalhamento', body, foot: `<button class="btn" data-x>Cancelar</button><button class="btn primary" data-ok>Salvar</button>`});
  const el = md.el; bindWeekPickers(el);
  const GROUPS = [['labor', ['MOD', 'MOI', 'TER'], 'moTotal'], ['equip', ['LA', 'LB', 'PP'], 'eqTotal']];
  const inp = k => $(`[data-k=${k}]`, el);
  const refInfo = () => {
    const w = parseWeek(inp('semana').value).v; const o = clean(inp('obra').value);
    const out = {}; if (!w || !o) return out;
    for (const kind of ['labor', 'equip']) {
      const recs = Repo.data[kind].filter(r => r.semana === w && obraKey(r.obra) === obraKey(o));
      const c = {}; KINDS[kind].cats.forEach(x => c[x.k] = 0); let t = 0; const seen = new Set();
      for (const r of recs) { if (!activeAtClose(kind, r, weekRange(w).end)) continue; const vk = vincKey(kind, r); if (seen.has(vk)) continue; seen.add(vk); if (r[KINDS[kind].catField] in c) c[r[KINDS[kind].catField]]++; t++; }
      out[kind] = {n: recs.length, c, t};
    }
    return {w, o, ...out};
  };
  const setLock = (i, on, val) => { i.readOnly = on; i.dataset.locked = on ? '1' : ''; i.style.background = on ? '#F3F4F6' : ''; i.style.color = on ? '#6B7280' : ''; i.title = on ? 'Valor dos cadastros individuais (não editável)' : ''; if (val !== undefined) i.value = val; const lab = i.previousElementSibling; if (lab) lab.dataset.lock = on ? '1' : ''; if (lab) lab.textContent = lab.textContent.replace(/ 🔒$/, '') + (on ? ' 🔒' : ''); };
  const sync = () => {
    const ri = refInfo();
    GROUPS.forEach(([kind, ks, tk]) => {
      const x = ri[kind]; const hasRec = !!(x && x.t);
      const t = inp(tk);
      if (hasRec || ks.some(k => inp(k).value !== '')) setLock(t, true, ks.reduce((sm, k) => sm + (Number(inp(k).value) || 0), 0));
      else if (t.dataset.locked) setLock(t, false, cur && !isEmptyVal(cur[tk]) ? cur[tk] : '');
    });
  };
  let lastKey = '';
  const applyLocks = () => {
    const ri = refInfo(); const key = (ri.w || '') + '|' + obraKey(ri.o || '');
    if (key !== lastKey) {
      lastKey = key;
      if (!existing) { cur = ri.w && ri.o ? Repo.data.manual.find(x => x.semana === ri.w && obraKey(x.obra) === obraKey(ri.o)) || null : null; }
      GROUPS.forEach(([kind, ks, tk]) => {
        const x = ri[kind];
        ks.forEach(k => { const i = inp(k); if (x && x.c[k] > 0) setLock(i, true, x.c[k]); else setLock(i, false, cur && !isEmptyVal(cur[k]) ? cur[k] : ''); });
        setLock(inp(tk), false, cur && !isEmptyVal(cur[tk]) ? cur[tk] : '');
      });
    }
    sync();
    const box = $('[data-ref]', el);
    box.innerHTML = (cur && !existing ? `<div class="alert warn" style="font-size:12px;padding:7px 10px">Já existe complemento para esta obra/semana — ele foi carregado para edição.</div>` : '') +
      ['labor', 'equip'].filter(k => ri[k] && ri[k].t).map(k => { const x = ri[k]; const K = KINDS[k];
        const free = K.cats.filter(c => !x.c[c.k]).map(c => c.l);
        return `<div class="alert ok" style="font-size:12px;padding:7px 10px"><b>${K.title}:</b> ${x.t} pelos cadastros (${K.cats.map(c => `${c.l} ${x.c[c.k]}`).join(' · ')}). ${free.length ? `Você pode complementar: <b>${free.join(', ')}</b>.` : 'Todas as categorias já vêm dos cadastros.'}</div>`; }).join('');
  };
  el.addEventListener('input', e => { if (e.target.dataset.k === 'obra' || e.target.dataset.k === 'semana') applyLocks(); else sync(); });
  el.addEventListener('change', e => { if (e.target.dataset.k === 'obra') applyLocks(); });
  applyLocks();
  $('[data-x]', el).onclick = md.close;
  $('[data-ok]', el).onclick = () => {
    const ri = refInfo(); const msg = $('[data-msg]', el); const errs = [];
    if (!ri.o) errs.push('Informe a obra.');
    const w = parseWeek(inp('semana').value); if (!w.v) errs.push(w.err || 'Informe a semana (2026-W40).');
    const rec = {obra: canonObra(inp('obra').value), semana: w.v};
    GROUPS.forEach(([kind, ks, tk]) => {
      const hasRec = !!(ri[kind] && ri[kind].t);
      ks.forEach(k => { const i = inp(k); const v = i.value.trim(); if (i.dataset.locked) { rec[k] = null; return; } if (v !== '' && !(Number.isInteger(Number(v)) && Number(v) >= 0)) errs.push(`Valor inválido em ${k}.`); rec[k] = v === '' ? null : Number(v); });
      const anyCat = ks.some(k => rec[k] != null); const tv = inp(tk).value.trim();
      rec[tk] = (!hasRec && !anyCat && tv !== '') ? Number(tv) : null;
      if (rec[tk] != null && !(Number.isInteger(rec[tk]) && rec[tk] >= 0)) errs.push('Total inválido.');
    });
    const empty = !GROUPS.some(([, ks, tk]) => ks.some(k => rec[k] != null) || rec[tk] != null);
    if (empty && !cur) errs.push('Preencha ao menos uma categoria que não tenha cadastro.');
    if (errs.length) { msg.innerHTML = `<div class="alert err" style="margin-top:10px">${errs.map(esc).join('<br>')}</div>`; return; }
    const res = Repo.tx(() => {
      if (cur) { if (empty) Repo.data.manual = Repo.data.manual.filter(x => x.id !== cur.id); else Object.assign(Repo.data.manual.find(x => x.id === cur.id), rec, {updatedAt: new Date().toISOString()}); }
      else Repo.data.manual.push({id: uid(), ...rec, createdAt: new Date().toISOString()});
    });
    if (res.ok) { toast(empty ? 'Complemento removido (sem valores).' : 'Totais salvos — os cadastros continuam valendo.', 'ok'); md.close(); refreshAll(); }
  };
}

/* =====================================================================
   RESUMO
   ===================================================================== */
function resumoWeekText(){ const R = UI.resumo; return R.week === '__last__' ? 'Semana mais recente de cada obra' : weekLabel(R.week, true); }
function resumoData(){
  const R = UI.resumo; const last = R.week === '__last__';
  const f = {obras: R.obras, weeks: !last && R.week ? [R.week] : [], from: '', to: '', cat: '', sit: R.sit, q: ''};
  const L = computeCells('labor', f), E = computeCells('equip', f);
  const opts = obraOptions(['labor', 'equip']);
  const obras = (R.obras.length ? R.obras : opts.map(o => o.v)).map(k => ({k, l: (opts.find(o => o.v === k) || {}).l || k}));
  // “Mais recente de cada obra”: para cada obra usa a última semana com dados (mão de obra e equipamentos independentes)
  const pick = (C, o) => { if (!last) return {c: C.cells.get(R.week + '||' + o) || null, w: R.week}; const ws = sortedWeeks(C.cells); for (let i = ws.length - 1; i >= 0; i--) { const c = C.cells.get(ws[i] + '||' + o); if (c) return {c, w: ws[i]}; } return {c: null, w: ''}; };
  const rows = obras.map(o => { const l = pick(L, o.k), e = pick(E, o.k); return {obra: o.l, k: o.k, L: l.c, E: e.c, wL: l.w, wE: e.w}; });
  return {rows, f};
}
function buildResumo(){
  const sec = $('#tab-resumo');
  sec.innerHTML = `
  <div class="toolbar"><div><h2>Resumo</h2><div class="sub">Quantitativos por obra na semana de referência</div></div>
    <div class="acts"><button class="btn" data-a="import">⇪ Importar planilha</button><button class="btn" data-a="template">⬇ Modelo Excel</button><button class="btn primary" data-a="manual">＋ Lançar totais sem detalhamento</button><button class="btn" data-a="excel">⊞ Exportar Excel</button><button class="btn dark" data-a="pdf">⎙ PDF</button></div></div>
  <div class="card filters" style="grid-template-columns:repeat(auto-fit,minmax(220px,1fr))">
    <label>Obras<div data-ms="obras"></div></label>
    <label>Semana de referência<select data-week></select></label>
    <label>Situação (mão de obra)<select data-rsit></select></label>
    <div class="seg sub-tabs" style="align-self:end;margin:0"><button data-sub="tabela">Tabela</button><button data-sub="grafico">Gráfico geral</button></div>
  </div>
  <div data-body></div>
  <div class="card"><div class="card-h"><h3>Lançamentos sem detalhamento</h3><span class="hint">Complementos por obra e semana — preenchem só as categorias sem cadastro individual</span><span class="sp"></span><button class="btn sm" data-a="exresumo" title="Gera um Excel com a aba Resumo já preenchida com estes lançamentos, para colar na planilha">⬇ Exportar no formato da aba Resumo</button></div><div data-manual-list></div></div>`;
  UI.ms.resumo = MultiSelect($('[data-ms=obras]', sec), {all: 'Todas as obras', value: [], options: () => obraOptions(['labor', 'equip']), quick: [{l: 'Todas', fn: () => []}], onChange: v => { UI.resumo.obras = v; renderResumo(); }});
  $('[data-week]', sec).onchange = e => { UI.resumo.week = e.target.value; renderResumo(); };
  $('[data-rsit]', sec).onchange = e => { UI.resumo.sit = e.target.value; renderResumo(); };
  sec.addEventListener('click', async e => {
    const s = e.target.closest('[data-sub]'); if (s) { UI.resumo.sub = s.dataset.sub; renderResumo(); return; }
    const j = e.target.closest('[data-jr]'); if (j) { UI.resumo.jr = j.dataset.jr; renderResumo(); return; }
    const b = e.target.closest('[data-a]'); if (!b) return;
    if (b.dataset.a === 'import') openImport('labor');
    else if (b.dataset.a === 'template') downloadTemplate();
    else if (b.dataset.a === 'exresumo') exportResumoSheet();
    else if (b.dataset.a === 'manual') openManualForm();
    else if (b.dataset.a === 'excel') exportResumoExcel();
    else if (b.dataset.a === 'pdf') printResumo();
    else if (b.dataset.a === 'medit') openManualForm(Repo.data.manual.find(m => m.id === b.dataset.id));
    else if (b.dataset.a === 'mdel') { const m = Repo.data.manual.find(x => x.id === b.dataset.id); if (m && await confirmBox('Excluir lançamento', `<p>Excluir os totais sem detalhamento de <b>${esc(m.obra)}</b> em <b>${esc(m.semana)}</b>?</p><p class="muted">Os cadastros individuais voltam a ser considerados.</p>`, 'Excluir', true)) { if (Repo.tx(() => { Repo.data.manual = Repo.data.manual.filter(x => x.id !== m.id); }).ok) toast('Lançamento excluído — cadastros individuais voltaram a ser considerados.', 'ok'); refreshAll(); } }
  });
}
function allWeeks(){ const s = new Set(); Repo.data.labor.forEach(r => s.add(r.semana)); Repo.data.equip.forEach(r => s.add(r.semana)); Repo.data.manual.forEach(r => s.add(r.semana)); return [...s].filter(weekRange).sort().reverse(); }
function cellTxt(c, k){ if (!c) return '0'; if (c[k] == null) return '—'; return fmtN(c[k]); }
function renderResumo(){
  const sec = $('#tab-resumo'); const R = UI.resumo;
  UI.ms.resumo.refresh(); R.obras = UI.ms.resumo.value;
  const weeks = allWeeks(); if (R.week !== '__last__' && !weeks.includes(R.week)) R.week = '__last__';
  $('[data-week]', sec).innerHTML = weeks.length ? `<option value="__last__">Mais recente de cada obra</option>` + weeks.map(w => `<option value="${w}">${esc(weekLabel(w, true))}</option>`).join('') : '<option value="__last__">Sem semanas cadastradas</option>';
  $('[data-week]', sec).value = R.week;
  { const sel = $('[data-rsit]', sec); const vals = new Map();
    for (const r of Repo.data.labor) { const s2 = clean(r.situacao); if (s2 && !vals.has(norm(s2))) vals.set(norm(s2), s2); }
    if (R.sit && R.sit !== '__none__' && !vals.has(R.sit)) R.sit = '';
    sel.innerHTML = `<option value="">Todas as situações</option>${[...vals].sort((a, b) => a[1].localeCompare(b[1], 'pt-BR')).map(([k, l]) => `<option value="${esc(k)}">${esc(l)}</option>`).join('')}<option value="__none__">Não informada</option>`;
    sel.value = R.sit; }
  const sitTxt = () => { const sel = $('[data-rsit]', sec); return R.sit ? sel.options[sel.selectedIndex].text : ''; };
  UI.resumo.sitLabel = sitTxt();
  $$('[data-sub]', sec).forEach(b => b.classList.toggle('on', b.dataset.sub === R.sub));
  const body = $('[data-body]', sec);
  Object.keys(charts).filter(k => k.startsWith('rs-')).forEach(killChart);
  if (journeys['rs-jr']) cancelAnimationFrame(journeys['rs-jr'].raf);
  const {rows} = resumoData(); const rwTxt = resumoWeekText();
  if (R.sub === 'tabela') {
    const L = KINDS.labor.cats, E = KINDS.equip.cats;
    const tot = {}; const unk = {}; [...L, ...E].forEach(c => { tot[c.k] = 0; unk[c.k] = false; }); tot.L = 0; tot.E = 0;
    rows.forEach(r => { L.forEach(c => { if (r.L && r.L[c.k] == null) unk[c.k] = true; else tot[c.k] += r.L ? r.L[c.k] : 0; }); E.forEach(c => { if (r.E && r.E[c.k] == null) unk[c.k] = true; else tot[c.k] += r.E ? r.E[c.k] : 0; }); tot.L += r.L ? r.L.total : 0; tot.E += r.E ? r.E.total : 0; });
    const tag = c => c && c.manual ? ` <span class="chip warn" title="Lançamento manual">${c.unknown ? 'Sem detalhamento' : '+ Complemento'}</span>` : '';
    body.innerHTML = `${R.sit ? `<div class="alert info" style="font-size:12px;padding:7px 10px">Filtro de situação aplicado aos cadastros de mão de obra (<b>${esc(R.sitLabel)}</b>). Complementos manuais não têm situação e continuam somados; equipamentos não são afetados.</div>` : ''}<div class="kpis k4">
        <div class="kpi dark"><div class="l">Total mão de obra</div><div class="v" data-n="${tot.L}">0</div><div class="s">${esc(rwTxt)}</div></div>
        ${L.map(c => `<div class="kpi" style="--c:${c.color}"><div class="l">${esc(c.full)}</div><div class="v" data-n="${tot[c.k]}">0</div><div class="s">${unk[c.k] ? 'há obras sem detalhamento' : ''}</div></div>`).join('')}
        <div class="kpi dark"><div class="l">Total equipamentos</div><div class="v" data-n="${tot.E}">0</div><div class="s">${esc(rwTxt)}</div></div>
        ${E.map(c => `<div class="kpi" style="--c:${c.color}"><div class="l">${esc(c.full)}</div><div class="v" data-n="${tot[c.k]}">0</div><div class="s">${unk[c.k] ? 'há obras sem detalhamento' : ''}</div></div>`).join('')}
      </div>
      <div class="card"><div class="card-h"><h3>Quantitativos por obra</h3><span class="hint">${esc(rwTxt)}</span></div>
      <div class="tbl-wrap"><table><thead><tr><th>Obra</th>${L.map(c => `<th>${c.l}</th>`).join('')}<th>Total MO</th>${E.map(c => `<th>${c.l}</th>`).join('')}<th>Total equip.</th></tr></thead><tbody>
      ${rows.map(r => `<tr><td class="l"><b>${esc(r.obra)}</b>${R.week === '__last__' && (r.wL || r.wE) ? ` <span class="muted" style="font-size:11px">${[...new Set([r.wL, r.wE].filter(Boolean))].map(weekShort).join(' / ')}</span>` : ''}</td>${L.map(c => `<td class="num">${cellTxt(r.L, c.k)}</td>`).join('')}<td class="num" style="background:#FFF4E6">${r.L ? fmtN(r.L.total) : 0}${tag(r.L)}</td>${E.map(c => `<td class="num">${cellTxt(r.E, c.k)}</td>`).join('')}<td class="num" style="background:#FFF4E6">${r.E ? fmtN(r.E.total) : 0}${tag(r.E)}</td></tr>`).join('') || `<tr><td colspan="9" class="muted">Sem obras.</td></tr>`}
      <tr class="tot"><td class="l">TOTAL GERAL (${rows.length} obra${rows.length !== 1 ? 's' : ''})</td>${L.map(c => `<td>${fmtN(tot[c.k])}${unk[c.k] ? '*' : ''}</td>`).join('')}<td class="hl">${fmtN(tot.L)}</td>${E.map(c => `<td>${fmtN(tot[c.k])}${unk[c.k] ? '*' : ''}</td>`).join('')}<td class="hl">${fmtN(tot.E)}</td></tr>
      </tbody></table></div>${Object.values(unk).some(Boolean) ? '<p class="muted" style="font-size:12px">* Categorias desconhecidas (“—”) de lançamentos sem detalhamento não entram na soma da categoria, mas entram no total.</p>' : ''}</div>`;
    $$('.kpi .v', body).forEach(el => countUp(el, Number(el.dataset.n)));
  } else {
    body.innerHTML = `<div class="card journey" data-rsj></div><div class="obra-grid">${rows.map((r, i) => `
      <div class="card obra-card"><h4>${esc(r.obra)} <span class="chip">${esc(R.week === '__last__' ? [...new Set([r.wL, r.wE].filter(Boolean))].map(weekShort).join(' / ') || '—' : R.week)}</span></h4><div class="pair">
        <div><div class="cap">Mão de obra</div><div style="position:relative;height:calc(100% - 16px)"><canvas id="rs-l-${i}"></canvas></div></div>
        <div><div class="cap">Equipamentos</div><div style="position:relative;height:calc(100% - 16px)"><canvas id="rs-e-${i}"></canvas></div></div>
      </div></div>`).join('') || '<div class="empty">Sem obras.</div>'}</div>`;
    rows.forEach((r, i) => { drawDonut(`rs-l-${i}`, 'labor', r.L); drawDonut(`rs-e-${i}`, 'equip', r.E); });
    // Percurso geral das obras selecionadas
    const kind = R.jr; const f = {obras: R.obras, weeks: [], from: '', to: '', cat: '', sit: kind === 'labor' ? R.sit : '', q: ''};
    const model = histogramModel(kind, f, 'semanaobra');
    const jbox = $('[data-rsj]', body);
    renderJourney(jbox, 'rs-jr', obraJourneys(kind, model, f), {unit: KINDS[kind].unit, subtitle: `${KINDS[kind].title} · todas as semanas`,
      extra: `<div class="seg"><button data-jr="labor" class="${kind === 'labor' ? 'on' : ''}">Mão de obra</button><button data-jr="equip" class="${kind === 'equip' ? 'on' : ''}">Equipamentos</button></div>`});
  }
  // lista de lançamentos manuais
  const ml = [...Repo.data.manual].filter(m => !R.obras.length || R.obras.includes(obraKey(m.obra))).sort((a, b) => b.semana.localeCompare(a.semana) || a.obra.localeCompare(b.obra, 'pt-BR'));
  const v = (m, k) => m[k] == null ? '<span class="muted">vazio</span>' : fmtN(m[k]);
  const g = (m, kind) => { const x = manualGroup(m, kind); return x ? `<b>${fmtN(x.total)}</b>${x.unknown ? ' <span class="chip warn">Sem detalhamento</span>' : ''}` : '<span class="muted">cadastros</span>'; };
  $('[data-manual-list]', sec).innerHTML = ml.length ? `<div class="tbl-wrap"><table><thead><tr><th>Ações</th><th>Semana</th><th>Obra</th><th>MOD</th><th>MOI</th><th>Terceiros</th><th>Total MO</th><th>L. amarela</th><th>L. branca</th><th>P. porte</th><th>Total equip.</th></tr></thead><tbody>
    ${ml.map(m => `<tr><td><button class="btn sm" data-a="medit" data-id="${m.id}">✎</button> <button class="btn sm danger" data-a="mdel" data-id="${m.id}">✕</button></td><td>${esc(m.semana)}</td><td class="l">${esc(m.obra)}</td><td>${v(m, 'MOD')}</td><td>${v(m, 'MOI')}</td><td>${v(m, 'TER')}</td><td>${g(m, 'labor')}</td><td>${v(m, 'LA')}</td><td>${v(m, 'LB')}</td><td>${v(m, 'PP')}</td><td>${g(m, 'equip')}</td></tr>`).join('')}
    </tbody></table></div>` : '<div class="muted" style="font-size:13px">Nenhum lançamento.</div>';
}
function drawDonut(id, kind, c){
  const K = KINDS[kind]; const cv = document.getElementById(id); if (!cv) return;
  let labels, data, colors;
  if (!c || !c.total) { labels = ['Sem efetivo']; data = [1]; colors = ['#E5E7EB']; }
  else if (c.unknown) { labels = [SD.l]; data = [c.total]; colors = [stripePattern()]; }
  else { labels = K.cats.map(x => x.l); data = K.cats.map(x => c[x.k]); colors = K.cats.map(x => x.color); }
  charts[id] = new Chart(cv, {type: 'doughnut', data: {labels, datasets: [{data, backgroundColor: colors, borderColor: '#fff', borderWidth: 3, hoverOffset: 8, borderRadius: 4, spacing: 1}]},
    options: {responsive: true, maintainAspectRatio: false, cutout: '66%', animation: reduceMotion ? false : {animateRotate: true, duration: 1000, easing: 'easeOutQuart'},
      plugins: {legend: {position: 'bottom', labels: {usePointStyle: true, pointStyle: 'circle', boxWidth: 8, font: {size: 10, weight: 'bold'}, padding: 8, filter: it => c && c.total}},
        centerText: {text: c ? fmtN(c.total) : '0', sub: c && c.unknown ? 'sem detalhamento' : K.unit},
        tooltip: {callbacks: {label: it => (!c || !c.total) ? ' 0' : ` ${it.label}: ${fmtN(it.raw)} (${Math.round(it.raw / c.total * 100)}%)`}}}}});
}

/* =====================================================================
   EXCEL: MODELOS, EXPORTAÇÃO
   ===================================================================== */
function filterSummary(kind){
  const f = UI.f[kind]; const K = KINDS[kind];
  const ob = obraOptions([kind]); const obras = f.obras.length ? f.obras.map(k => (ob.find(o => o.v === k) || {l: k}).l).join(', ') : 'Todas';
  const out = [['Obra(s)', obras], ['Semanas', f.weeks.length ? [...f.weeks].sort().join(', ') : 'Todas'], ['Período', f.from || f.to ? `${fmtBR(f.from) || '…'} a ${fmtBR(f.to) || '…'}` : 'Todo'], [K.catTitle, f.cat ? catLabel(kind, f.cat) : 'Todas']];
  if (kind === 'labor') { const sel = $('#tab-labor [data-f=sit]'); out.push(['Situação', f.sit ? sel.options[sel.selectedIndex].text : 'Todas as situações']); }
  out.push(['Busca', f.q || '—']);
  return out;
}
function setTextCols(ws, nCols, idCols){
  // mantém identificadores como texto (zeros à esquerda)
  const rg = XLSX.utils.decode_range(ws['!ref']);
  for (let R = rg.s.r + 1; R <= rg.e.r; R++) for (const C of idCols) { const c = ws[XLSX.utils.encode_cell({r: R, c: C})]; if (c) { c.t = 's'; c.z = '@'; } }
  ws['!cols'] = Array.from({length: nCols}, () => ({wch: 16}));
}
function downloadTemplate(){
  if (!libsOk()) return toast('Biblioteca Excel não carregada.', 'err');
  const wb = XLSX.utils.book_new();
  for (const kind of ['labor', 'equip']) {
    const K = KINDS[kind]; const ws = XLSX.utils.aoa_to_sheet([K.fields.map(f => f.l)]); ws['!cols'] = K.fields.map(f => ({wch: Math.max(12, f.l.length + 4)}));
    XLSX.utils.book_append_sheet(wb, ws, IMPORT_SHEET[kind]);
  }
  const wr = XLSX.utils.aoa_to_sheet([RESUMO_COLS.map(c => c.l)]); wr['!cols'] = RESUMO_COLS.map(c => ({wch: Math.max(10, c.l.length + 4)}));
  XLSX.utils.book_append_sheet(wb, wr, 'Resumo');
  const desc = (kind, f) => f.t === 'week' ? 'Semana ISO: 2026-W40 (segunda a domingo). Também aceita 40/2026 ou uma data do período.' : f.t === 'date' ? 'Data dd/mm/aaaa. “-” não é aceito como data (fica como não informado).' : f.t === 'cat' ? (kind === 'labor' ? 'MOD (Direto/Direta/MOD), MOI (Indireto/Indireta/MOI) ou Terceiros/Terceirizado' : 'Linha amarela, Linha branca ou P. porte') : f.k === 'matricula' ? 'Use 0 para PJ — identificado pelo CPF ou, sem CPF, pelo nome na obra.' : f.id ? 'Texto (zeros à esquerda são preservados — formate a coluna como Texto).' : 'Texto';
  const ins = [['Instruções — Controle de Histogramas'], [],
    ['A planilha tem três abas: “MO” (Mão de obra), “Equipamento” (Equipamentos) e “Resumo” (totais por obra e semana). As três são importadas juntas.'],
    ['Não renomeie as abas. Mantenha a linha de cabeçalho; os dados começam na linha 2.'], []];
  for (const kind of ['labor', 'equip']) {
    const K = KINDS[kind];
    ins.push([`Aba ${IMPORT_SHEET[kind]} → ${K.title}`], ['Coluna', 'Obrigatório', 'Formato / valores aceitos'], ...K.fields.map(f => [f.l, f.req ? 'Sim' : 'Não', desc(kind, f)]),
      [kind === 'labor' ? 'Chave: Semana + Obra + Matrícula (PJ: Semana + Obra + CPF ou Nome).' : 'Chave: Semana + Obra + Prefixo.'], []);
  }
  ins.push(['Aba Resumo → totais por obra e semana'], ['Coluna', 'Obrigatório', 'Formato / valores aceitos'],
    ['Obra', 'Sim', 'Texto'], ['Semana', 'Sim', 'Semana ISO: 2026-W40'],
    ['MOD, MOI, Terceiros', 'Não', 'Números inteiros. Vazio = não informado; 0 = zero.'],
    ['Total MO', 'Não', 'Use só quando não houver divisão por categoria (fica “Sem detalhamento”). Se as categorias forem preenchidas, vale a soma delas.'],
    ['Linha amarela, Linha branca, P. porte', 'Não', 'Números inteiros.'],
    ['Total equipamentos', 'Não', 'Use só quando não houver divisão por classificação.'],
    ['Regra', '', 'Uma linha por obra e semana. Os valores complementam os cadastros: categorias que já têm cadastros individuais (abas MO/Equipamento) continuam vindo dos cadastros.'], ['Regra', '', 'Obra/semana que não estiver nesta aba mantém o que já existe no app. Em linha já existente, célula vazia mantém o valor atual; para zerar, informe 0.'], []);
  ins.push(['Regras gerais'],
    ['Cada semana é uma fotografia: importar nova semana cria novos registros e preserva as anteriores.'],
    ['Reimportar a mesma obra/semana atualiza os cadastros (inclusive campos vazios) após confirmação, preservando a data inicial original.'],
    ['Se a planilha não tiver coluna Semana, informe a semana na tela de importação.'],
    ['Efetivo do histograma = ativos no fechamento da semana (domingo), datas inicial e final inclusivas.']);
  ins.push([], ...trkInstr());
  const wi = XLSX.utils.aoa_to_sheet(ins); wi['!cols'] = [{wch: 24}, {wch: 12}, {wch: 95}];
  trkTemplateSheets(wb);
  XLSX.utils.book_append_sheet(wb, wi, 'Instruções');
  XLSX.writeFile(wb, 'modelo_importacao_MO_Equipamento_Resumo.xlsx');
}
function histRows(kind, model){
  const K = KINDS[kind]; const rows = [];
  for (const w of model.weeks) for (const o of model.obras) { const c = model.cells.get(w + '||' + o); if (!c) continue; rows.push([w, weekRange(w).end ? fmtBR(weekRange(w).end) : '', model.labels.get(o), ...K.cats.map(x => c[x.k] == null ? '—' : c[x.k]), c.SD || 0, c.total, c.manual ? (c.unknown ? 'Sem detalhamento' : 'Lançamento manual') : 'Cadastros']); }
  return [['Semana', 'Fechamento', 'Obra', ...K.cats.map(x => x.l), 'Sem detalhamento', 'Total', 'Origem'], ...rows];
}
function exportExcel(kind){
  if (!libsOk()) return toast('Biblioteca Excel não carregada.', 'err');
  const K = KINDS[kind]; const model = UI._last[kind]; const wb = XLSX.utils.book_new();
  const fs = XLSX.utils.aoa_to_sheet([['Controle de Histogramas — ' + K.title], ['Gerado em', new Date().toLocaleString('pt-BR')], [], ['Filtro', 'Valor'], ...filterSummary(kind)]); fs['!cols'] = [{wch: 18}, {wch: 80}];
  XLSX.utils.book_append_sheet(wb, fs, 'Filtros');
  const recs = [...model.rows].sort((a, b) => a.semana.localeCompare(b.semana) || String(a.obra).localeCompare(String(b.obra)));
  const rs = XLSX.utils.aoa_to_sheet([[...K.fields.map(f => f.l), 'No fechamento'], ...recs.map(r => [...K.fields.map(f => displayVal(kind, r, f)), activeAtClose(kind, r, weekRange(r.semana).end) ? 'Sim' : 'Não'])]);
  setTextCols(rs, K.fields.length + 1, K.fields.map((f, i) => f.id ? i : -1).filter(i => i >= 0));
  XLSX.utils.book_append_sheet(wb, rs, 'Registros');
  const hs = XLSX.utils.aoa_to_sheet(histRows(kind, model)); hs['!cols'] = [{wch: 11}, {wch: 12}, {wch: 34}, {wch: 12}, {wch: 12}, {wch: 12}, {wch: 16}, {wch: 8}, {wch: 20}];
  XLSX.utils.book_append_sheet(wb, hs, 'Histograma semanal');
  XLSX.writeFile(wb, `histograma_${kind === 'labor' ? 'mao_de_obra' : 'equipamentos'}_${new Date().toISOString().slice(0, 10)}.xlsx`);
}
/** Exporta os lançamentos manuais no mesmo formato da aba “Resumo” da planilha de importação */
function exportResumoSheet(){
  if (!libsOk()) return toast('Biblioteca Excel não carregada.', 'err');
  const R = UI.resumo; const ml = [...Repo.data.manual].filter(m => !R.obras.length || R.obras.includes(obraKey(m.obra))).sort((a, b) => a.semana.localeCompare(b.semana) || a.obra.localeCompare(b.obra, 'pt-BR'));
  const aoa = [RESUMO_COLS.map(c => c.l), ...ml.map(m => [m.obra, m.semana, ...NUMK.map(k => m[k] == null ? '' : Number(m[k]))])];
  const wb = XLSX.utils.book_new(); const ws = XLSX.utils.aoa_to_sheet(aoa); ws['!cols'] = RESUMO_COLS.map(c => ({wch: Math.max(10, c.l.length + 4)}));
  XLSX.utils.book_append_sheet(wb, ws, 'Resumo'); XLSX.writeFile(wb, 'aba_Resumo_lancamentos.xlsx');
  toast(`${ml.length} lançamento(s) exportado(s) no formato da aba Resumo.`, 'ok');
}
function exportResumoExcel(){
  if (!libsOk()) return toast('Biblioteca Excel não carregada.', 'err');
  const {rows} = resumoData();
  const g = (c, k) => !c ? 0 : c[k] == null ? '—' : c[k];
  const aoa = [['Resumo — ' + resumoWeekText()], ['Situação (mão de obra)', UI.resumo.sitLabel || 'Todas as situações'], [], ['Obra', 'MOD', 'MOI', 'Terceiros', 'Total MO', 'Linha amarela', 'Linha branca', 'P. porte', 'Total equipamentos'],
    ...rows.map(r => [r.obra, g(r.L, 'MOD'), g(r.L, 'MOI'), g(r.L, 'TER'), r.L ? r.L.total : 0, g(r.E, 'LA'), g(r.E, 'LB'), g(r.E, 'PP'), r.E ? r.E.total : 0])];
  const sum = i => aoa.slice(4).reduce((s, r) => s + (typeof r[i] === 'number' ? r[i] : 0), 0);
  aoa.push(['TOTAL GERAL', sum(1), sum(2), sum(3), sum(4), sum(5), sum(6), sum(7), sum(8)]);
  const wb = XLSX.utils.book_new(); const ws = XLSX.utils.aoa_to_sheet(aoa); ws['!cols'] = [{wch: 36}, ...Array(8).fill({wch: 14})];
  XLSX.utils.book_append_sheet(wb, ws, 'Resumo'); XLSX.writeFile(wb, `resumo_${UI.resumo.week === '__last__' ? 'mais_recente' : UI.resumo.week}.xlsx`);
}

/* =====================================================================
   BACKUP / RESTAURAÇÃO
   ===================================================================== */
function backup(){
  const d = {app: 'controle-histogramas-cosampa', version: 1, exportedAt: new Date().toISOString(), counts: Object.fromEntries(ALLK.map(k => [k, Repo.data[k].length])), data: Repo.data};
  download(`backup_histogramas_${new Date().toISOString().slice(0, 19).replace(/[:T]/g, '-')}.json`, new Blob([JSON.stringify(d, null, 1)], {type: 'application/json'}));
  toast('Backup gerado.', 'ok');
}
function openClearData(){
  const G = [['labor', 'Mão de obra (cadastros)', ['labor']], ['equip', 'Equipamentos (cadastros)', ['equip']], ['manual', 'Totais sem detalhamento / aba Resumo', ['manual']], ['ativ', 'Atividades e histórico de avanço', ['ativ', 'avanco', 'lps', 'restr']]];
  const cnt = (ks, o) => ks.reduce((s, k) => s + Repo.data[k].filter(r => !o || obraKey(r.obra) === o).length, 0);
  const body = `<label class="fld" style="display:block;margin-bottom:10px"><span>Obra</span><select data-co style="width:100%;border:1px solid var(--line);border-radius:8px;padding:8px"><option value="">Todas as obras</option>${trkObras().map(o => `<option value="${esc(o.v)}">${esc(o.l)}</option>`).join('')}</select></label>
    <div data-cg>${G.map(([id, l]) => `<label style="display:flex;gap:8px;align-items:center;margin:6px 0"><input type="checkbox" data-cgk="${id}" checked> <span>${l} — <b data-cn="${id}"></b> registro(s)</span></label>`).join('')}</div>
    <div class="alert warn" style="margin-top:10px;font-size:12px">${Repo.mode === 'db' ? '<b>Atenção: os dados serão apagados do BANCO COMPARTILHADO, para todos os usuários.</b> ' : ''}Recomenda-se baixar um backup antes. Um ponto de restauração é criado automaticamente (Configurações → Pontos de restauração).</div>
    <label class="fld" style="display:block;margin-top:10px"><span>Para confirmar, digite <b>APAGAR</b></span><input data-cc autocomplete="off" style="width:100%;border:1px solid var(--line);border-radius:8px;padding:8px"></label>`;
  const md = openModal({title: 'Limpar dados', body, foot: '<button class="btn" data-bk>⬇ Baixar backup</button><span style="flex:1"></span><button class="btn" data-x>Cancelar</button><button class="btn danger" data-ok disabled>Apagar</button>'});
  const el = md.el; $('[data-x]', el).onclick = md.close; $('[data-bk]', el).onclick = backup;
  const sel = () => G.filter(([id]) => $(`[data-cgk=${id}]`, el).checked);
  const upd = () => { const o = $('[data-co]', el).value; G.forEach(([id, , ks]) => $(`[data-cn=${id}]`, el).textContent = fmtN(cnt(ks, o)));
    const n = sel().reduce((s, g) => s + cnt(g[2], o), 0); const b = $('[data-ok]', el); b.disabled = !n || $('[data-cc]', el).value.trim().toUpperCase() !== 'APAGAR'; b.textContent = n ? `Apagar ${fmtN(n)} registro(s)` : 'Nada a apagar'; };
  el.addEventListener('input', upd); el.addEventListener('change', upd); upd();
  $('[data-ok]', el).onclick = () => {
    const o = $('[data-co]', el).value; const ks = sel().flatMap(g => g[2]);
    Snap.take('Antes de limpar dados' + (o ? ' (' + obraLabel(o) + ')' : ''));
    const res = Repo.tx(() => { for (const k of ks) Repo.data[k] = o ? Repo.data[k].filter(r => obraKey(r.obra) !== o) : []; });
    if (res.ok) { toast('Dados apagados. Pode importar novamente.', 'ok'); md.close(); refreshAll(); }
  };
}
async function restore(file){
  let d; try { d = JSON.parse(await file.text()); } catch (e) { return toast('Arquivo inválido: não é um JSON.', 'err'); }
  const data = d && d.data ? d.data : d;
  const errs = [];
  if (!data || !Array.isArray(data.labor) || !Array.isArray(data.equip)) errs.push('Estrutura não reconhecida (esperado labor/equip).');
  const clean2 = (kind, arr) => (arr || []).map((r, i) => { if (!r || typeof r !== 'object' || !r.obra || !weekRange(r.semana)) errs.push(`${kind} #${i + 1}: obra/semana inválida`); return {id: r.id || uid(), ...r}; });
  if (!errs.length) {
    const nd = {labor: clean2('Mão de obra', data.labor), equip: clean2('Equipamentos', data.equip), manual: clean2('Totais', data.manual || [])};
    for (const k of TRK) nd[k] = (Array.isArray(data[k]) ? data[k] : []).filter(r => r && typeof r === 'object' && r.obra).map(r => ({id: r.id || uid(), ...r}));
    if (errs.length) return toast('Backup com problemas: ' + errs.slice(0, 3).join('; '), 'err');
    const ok = await confirmBox('Restaurar backup', `<p>O backup contém <b>${nd.labor.length}</b> cadastro(s) de mão de obra, <b>${nd.equip.length}</b> de equipamentos e <b>${nd.manual.length}</b> lançamento(s) sem detalhamento, além de ${TRK.reduce((a, k) => a + nd[k].length, 0)} registro(s) de acompanhamento (atividades, avanço, compromissos e restrições)${d.exportedAt ? `, gerado em ${new Date(d.exportedAt).toLocaleString('pt-BR')}` : ''}.</p><div class="alert warn"><b>${Repo.mode === 'db' ? 'Todos os dados do BANCO COMPARTILHADO serão substituídos (para todos os usuários).' : 'Todos os dados atuais deste navegador serão substituídos.'}</b> Recomenda-se baixar um backup antes.</div>`, 'Substituir dados', true);
    if (!ok) return;
    Snap.take('Antes de restaurar backup');
    if (Repo.replaceAll(nd)) { toast('Backup restaurado.', 'ok'); refreshAll(); }
  } else toast('Backup inválido: ' + errs.slice(0, 3).join('; '), 'err');
}

/* =====================================================================
   RELATÓRIOS PDF (impressão A4 paisagem)
   ===================================================================== */
const LOGO_COSAMPA = 'data:image/png;base64,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';
function logoHTML(){ return `<img src="${getLogo() || LOGO_COSAMPA}" alt="Grupo COSAMPA">`; }
function getLogo(){ try { return localStorage.getItem(LS_LOGO); } catch (e) { return null; } }
function pHead(title){ return `<div class="p-head">${logoHTML(true)}<span class="tt">Controle de Histogramas — ${esc(title)}</span><span class="dt">Emitido em ${new Date().toLocaleString('pt-BR')}</span></div>`; }
function pFilters(list){ return `<div class="p-filt">${list.map(([k, v]) => `<b>${esc(k)}:</b> ${esc(v)}`).join(' &nbsp;·&nbsp; ')}</div>`; }
function offscreenImage(cfg, w = 1100, h = 460){
  const host = $('#offscreen'); host.style.display = 'block';
  const cv = document.createElement('canvas'); cv.width = w; cv.height = h; cv.style.width = w + 'px'; cv.style.height = h + 'px'; host.appendChild(cv);
  const ch = new Chart(cv, cfg); const url = cv.toDataURL('image/png'); ch.destroy(); cv.remove(); host.style.display = 'none';
  return url;
}
function journeyImage(jm, unit){
  if (!jm) return null;
  const ds = [{label: 'Total', data: jm.totals, borderColor: '#111827', backgroundColor: 'rgba(247,147,30,.18)', fill: true, borderWidth: 3, pointRadius: 3, tension: 0}];
  jm.cats.forEach(c => ds.push({label: c.l, data: c.data, borderColor: c.color, backgroundColor: c.color, borderWidth: 1.5, pointRadius: 2, tension: 0}));
  return offscreenImage({type: 'line', data: {labels: jm.weeks.map(weekShort), datasets: ds}, plugins: [{id: 'pj', beforeInit(c){ c.$j = {n: jm.weeks.length, reveal: jm.weeks.length - 1, cursor: false, miles: jm.miles}; }}],
    options: {responsive: false, animation: false, devicePixelRatio: 2, layout: {padding: {top: 34, right: 20}}, scales: {x: {grid: {display: false}}, y: {beginAtZero: true, grace: '18%', ticks: {precision: 0}}}, plugins: {legend: {position: 'top', align: 'end', labels: {usePointStyle: true}}}}}, 1100, 380);
}
function doPrint(html){
  if (!libsOk()) return toast('Bibliotecas não carregadas.', 'err');
  $('#print-area').innerHTML = html;
  setTimeout(() => { window.print(); }, 120);
}
window.addEventListener('afterprint', () => { $('#print-area').innerHTML = ''; });
function printKind(kind){
  const K = KINDS[kind]; const model = UI._last[kind]; const flt = filterSummary(kind);
  let html = '';
  const CH = 14;
  if (!model.groups.length) html += `<div class="p-page">${pHead(K.title)}${pFilters(flt)}<p>Nenhum dado para os filtros.</p></div>`;
  for (let i = 0; i < model.groups.length; i += CH) {
    const sub = {...model, groups: model.groups.slice(i, i + CH), datasets: model.datasets.map(d => ({...d, data: d.data.slice(i, i + CH), unknown: d.unknown.slice(i, i + CH)}))};
    const img = offscreenImage(barConfig(sub, {print: true, animate: false}));
    const part = model.groups.length > CH ? ` (parte ${i / CH + 1} de ${Math.ceil(model.groups.length / CH)})` : '';
    html += `<div class="p-page">${pHead(K.title)}${i === 0 ? pFilters(flt) : ''}<div class="p-sec">Histograma — efetivo no fechamento da semana${part}</div><img class="p-img" src="${img}">${model.note ? `<div class="p-note">${esc(model.note)}</div>` : ''}${model.datasets.some(d => d.key === 'SD') ? '<div class="p-note">Barras cinza = totais sem detalhamento; “—” = categoria desconhecida.</div>' : ''}</div>`;
  }
  const jl = obraJourneys(kind, model, UI.f[kind]).filter(x => x.jm);
  for (let i = 0; i < jl.length; i += 2) {
    html += `<div class="p-page">${pHead(K.title)}<div class="p-sec">Percurso do efetivo por obra (semana a semana)</div>${jl.slice(i, i + 2).map(x => { const jm = x.jm; return `<div class="p-note"><b>${esc(x.label)}</b> — início ${fmtN(jm.totals[jm.s])} (${jm.weeks[jm.s]}) · pico ${fmtN(jm.totals[jm.p])} (${jm.weeks[jm.p]}) · final ${fmtN(jm.totals[jm.e])} (${jm.weeks[jm.e]})</div><img class="p-img" style="max-height:82mm" src="${journeyImage(jm, K.unit)}">`; }).join('')}</div>`;
  }
  // Totais
  const hr = histRows(kind, model);
  html += `<div class="p-page">${pHead(K.title)}<div class="p-sec">Totais por semana e obra</div><table class="p-tbl"><thead><tr>${hr[0].map(h => `<th>${esc(h)}</th>`).join('')}</tr></thead><tbody>${hr.slice(1).map(r => `<tr>${r.map(v => `<td>${esc(v)}</td>`).join('')}</tr>`).join('')}</tbody></table>
    <div class="p-sec">Relação completa (${model.rows.length} cadastro(s))</div>
    <table class="p-tbl"><thead><tr>${K.fields.map(f => `<th>${esc(f.l)}</th>`).join('')}<th>Fech.</th></tr></thead><tbody>
    ${[...model.rows].sort((a, b) => a.semana.localeCompare(b.semana) || String(a.obra).localeCompare(String(b.obra)) || String(a.nome || a.prefixo).localeCompare(String(b.nome || b.prefixo))).map(r => `<tr>${K.fields.map(f => `<td>${esc(displayVal(kind, r, f))}</td>`).join('')}<td>${activeAtClose(kind, r, weekRange(r.semana).end) ? 'Sim' : 'Não'}</td></tr>`).join('')}
    </tbody></table>${model.hidden.length ? `<div class="p-note">${model.hidden.length} cadastro(s) individuais ocultos por lançamentos sem detalhamento.</div>` : ''}</div>`;
  doPrint(html);
}
function printResumo(){
  const {rows} = resumoData(); const R = UI.resumo;
  const ob = obraOptions(['labor', 'equip']);
  const g = (c, k) => !c ? 0 : c[k] == null ? '—' : fmtN(c[k]);
  const tot = i => rows.reduce((s, r) => s + i(r), 0);
  const sumK = (grp, k) => tot(r => r[grp] && r[grp][k] != null ? r[grp][k] : 0);
  const html = `<div class="p-page">${pHead('Resumo')}${pFilters([['Semana de referência', resumoWeekText()], ['Obras', R.obras.length ? R.obras.map(k => (ob.find(o => o.v === k) || {l: k}).l).join(', ') : 'Todas'], ['Situação (mão de obra)', R.sitLabel || 'Todas as situações']])}
    <table class="p-tbl" style="font-size:9px"><thead><tr><th>Obra</th><th>MOD</th><th>MOI</th><th>Terceiros</th><th>Total MO</th><th>Linha amarela</th><th>Linha branca</th><th>P. porte</th><th>Total equip.</th></tr></thead><tbody>
    ${rows.map(r => `<tr><td style="text-align:left">${esc(r.obra)}${(r.L && r.L.unknown) || (r.E && r.E.unknown) ? ' (sem detalhamento)' : ''}</td><td>${g(r.L, 'MOD')}</td><td>${g(r.L, 'MOI')}</td><td>${g(r.L, 'TER')}</td><td><b>${r.L ? fmtN(r.L.total) : 0}</b></td><td>${g(r.E, 'LA')}</td><td>${g(r.E, 'LB')}</td><td>${g(r.E, 'PP')}</td><td><b>${r.E ? fmtN(r.E.total) : 0}</b></td></tr>`).join('')}
    <tr class="tot"><td style="text-align:left">TOTAL GERAL</td><td>${fmtN(sumK('L', 'MOD'))}</td><td>${fmtN(sumK('L', 'MOI'))}</td><td>${fmtN(sumK('L', 'TER'))}</td><td>${fmtN(tot(r => r.L ? r.L.total : 0))}</td><td>${fmtN(sumK('E', 'LA'))}</td><td>${fmtN(sumK('E', 'LB'))}</td><td>${fmtN(sumK('E', 'PP'))}</td><td>${fmtN(tot(r => r.E ? r.E.total : 0))}</td></tr>
    </tbody></table></div>`;
  doPrint(html);
}

/* =====================================================================
   INICIALIZAÇÃO
   ===================================================================== */
/* ---------- Conexão com a planilha publicada (leitura) ---------- */
const Sync = {
  wb: null, at: null, busy: false, lastWb: null, lastSrc: '', lastAt: null,
  bar(cls, html, acts = ''){ const b = $('#syncBar'); b.className = 'sync ' + cls; b.innerHTML = `<span class="dot"></span><span>${html}</span><span class="acts">${acts}</span>`; },
  summarize(wb){
    const out = {};
    { const sh = guessSheet('manual', wb); const g = sh && sheetGrid(wb.Sheets[sh]); const h = g && mapResumoHeaders(g);
      if (h) { const res = analyzeResumo(g, h, '', false); const v = res.filter(r => !r.errors.length); out.manual = {sheet: sh, novos: v.filter(r => !r.existing).length, alterados: v.filter(r => r.changed).length, pend: res.length - v.length, total: res.length}; } else out.manual = null; }
    for (const k of ['labor', 'equip']) {
      const sh = guessSheet(k, wb); if (!sh) { out[k] = null; continue; }
      const grid = sheetGrid(wb.Sheets[sh]); const hdr = mapHeaders(k, grid); if (!hdr || !DISTINCT[k].some(x => hdr.map[x] != null)) { out[k] = null; continue; }
      const res = analyzeImport(k, grid, hdr, '', false, detectDateOrder(k, grid, hdr) || detectWbOrder(wb) || 'dmy');
      const v = res.filter(r => !r.errors.length);
      out[k] = {sheet: sh, novos: v.filter(r => !r.existing).length, alterados: v.filter(r => r.changed).length, pend: res.length - v.length, total: res.length};
    }
    return out;
  },
  async run(manual){
    if (this.busy || !libsOk()) return; this.busy = true;
    const cfg = sheetCfg();
    this.bar('load', 'Lendo a planilha conectada…');
    try {
      this.wb = await fetchPublishedWorkbook(cfg.url, {labor: cfg.gidL, equip: cfg.gidE}); this.at = new Date(); this.lastWb = this.wb; this.lastSrc = 'Planilha conectada'; this.lastAt = this.at;
      this.paint();
      if (manual) { const sm = this.summarize(this.wb); if (['labor', 'equip', 'manual'].some(k => sm[k] && (sm[k].novos || sm[k].alterados || sm[k].pend))) this.review(); else toast('O app já está igual à planilha.', 'ok'); }
    } catch (e) {
      Log.add('Não foi possível ler a planilha conectada — ' + e.message + ' Sem acesso pelo link, use “Importar planilha” com o arquivo .xlsx baixado.', 'err');
      this.bar('err', `<b>Planilha não conectada</b> · <a href="#" data-sync="avisos" style="color:inherit">ver detalhes em Avisos</a>`, `<button class="btn sm" data-sync="run">Tentar de novo</button><button class="btn sm" data-sync="cfg">Conexão</button>`);
    } finally { this.busy = false; }
  },
  paint(){
    if (!this.wb) return; if (UI.tab === 'config') renderConfig(); if (UI.tab === 'avisos') renderAvisos();
    const sm = this.summarize(this.wb); const hh = this.at.toLocaleTimeString('pt-BR', {hour: '2-digit', minute: '2-digit'});
    const part = k => { const x = sm[k]; const nm = IMPORT_SHEET[k]; if (!x) return `aba ${nm}: não encontrada`; const bits = []; if (x.novos) bits.push(`<b>${x.novos} novo(s)</b>`); if (x.alterados) bits.push(`<b>${x.alterados} alterado(s)</b>`); if (x.pend) bits.push(`${x.pend} com erro`); return `aba ${nm}: ${bits.length ? bits.join(', ') : 'sem novidades'}`; };
    const chg = ['labor', 'equip', 'manual'].some(k => sm[k] && (sm[k].novos || sm[k].alterados));
    this.bar(chg ? 'chg' : 'ok', `<b>Planilha conectada</b> · lida às ${hh} · ${part('labor')} · ${part('equip')}${sm.manual ? ' · ' + part('manual') : ''}`,
      `${chg || ['labor', 'equip', 'manual'].some(k => sm[k] && sm[k].pend) ? '<button class="btn sm primary" data-sync="review">Revisar e importar</button>' : ''}<button class="btn sm" data-sync="run">⟳ Ler de novo</button>`);
  },
  review(){ if (this.wb) openImport(UI.tab === 'equip' ? 'equip' : 'labor', this.wb, `Planilha conectada (lida às ${this.at.toLocaleTimeString('pt-BR', {hour: '2-digit', minute: '2-digit'})})`); },
  markImported(){ setTimeout(() => this.paint(), 0); },
  auto(){ if (sheetCfg().auto) this.run(false); else this.bar('', 'Planilha conectada — leitura automática desativada.', '<button class="btn sm" data-sync="run">⟳ Ler agora</button><button class="btn sm" data-sync="cfg">Conexão</button>'); },
  config(){
    const c = sheetCfg();
    const md = openModal({title: 'Conexão com a planilha', body: `
      <div class="alert info" style="font-size:12px">O app <b>lê</b> a planilha (abas <b>MO</b> e <b>Equipamento</b>) e nunca grava nada no Google Sheets. Para o app conseguir ler sem login, a planilha precisa estar em <b>Compartilhar → Acesso geral → Qualquer pessoa com o link — Leitor</b> (ou publicada na Web). Novidades e alterações são mostradas para você revisar e confirmar.</div>
      <div class="alert warn" style="font-size:12px">Atenção: com o link aberto, qualquer pessoa que tiver o link consegue ver os dados (inclusive CPF e telefone). Se isso não for aceitável, mantenha a planilha privada e use “Importar planilha” com o arquivo baixado.</div>
      <div class="fld"><span>Link da planilha (edição/compartilhamento ou publicado)</span><input data-c="url" value="${esc(c.url)}"></div>
      <div class="form-grid" style="margin-top:10px"><div class="fld"><span>gid da aba MO (opcional)</span><input data-c="gidL" value="${esc(c.gidL)}"></div><div class="fld"><span>gid da aba Equipamento (opcional)</span><input data-c="gidE" value="${esc(c.gidE)}" placeholder="opcional"></div></div>
      <p class="muted" style="font-size:12px">Normalmente não é preciso informar gid: as abas são encontradas pelo nome (MO e Equipamento). Use o gid só se renomear as abas — é o número depois de <i>gid=</i> no endereço da aba.</p>
      <label style="display:flex;gap:8px;align-items:center;margin-top:6px"><input type="checkbox" data-c="auto" ${c.auto ? 'checked' : ''} style="width:auto;min-height:0"> Ler a planilha automaticamente ao abrir o app</label>`,
      foot: `<button class="btn" data-x>Cancelar</button><button class="btn primary" data-ok>Salvar e ler agora</button>`});
    $('[data-x]', md.el).onclick = md.close;
    $('[data-ok]', md.el).onclick = () => { const n = {url: $('[data-c=url]', md.el).value.trim(), gidL: $('[data-c=gidL]', md.el).value.trim(), gidE: $('[data-c=gidE]', md.el).value.trim(), auto: $('[data-c=auto]', md.el).checked}; saveSheetCfg(n); md.close(); this.run(true); };
  }
};
document.addEventListener('click', e => { const b = e.target.closest('[data-sync]'); if (!b) return; const a = b.dataset.sync; e.preventDefault(); if (a === 'run') Sync.run(true); else if (a === 'review') Sync.review(); else if (a === 'cfg') Sync.config(); else if (a === 'avisos') switchTab('avisos'); });
function paintStorage(){ const el = document.getElementById('storeUse'); if (!el) return; const mb = Repo.lastBytes / 1048576; el.innerHTML = `Armazenamento usado: <b>${mb < 0.1 ? Math.max(1, Math.round(Repo.lastBytes / 1024)) + ' KB' : mb.toLocaleString('pt-BR', {maximumFractionDigits: 1}) + ' MB'}</b> de ~5 MB`; el.style.color = mb > 4 ? '#FCA5A5' : ''; }
function paintLogo(){ $('#logoBox').innerHTML = logoHTML(false); $('#logoBoxM').innerHTML = logoHTML(false); }
/* =====================================================================
   ACOMPANHAMENTO DA OBRA
   • Painel da obra — visão única: avanço físico (curva S), IDP, PPC, restrições, efetivo
   • Atividades — cronograma físico, atividades restantes, ritmo atual × necessário, previsão de término
   • Planejamento semanal (Last Planner) — compromissos, PPC e causas de não cumprimento
   • Restrições — lookahead de 6 semanas, vencidas, índice de remoção
   ===================================================================== */
const TK = {obra: '', week: '', ativ: {disc: '', status: 'rest', q: ''}, restr: {status: 'abertas', tipo: ''}};
const DAY = 86400000;
const ymdToUTC = ymd => { const [y, m, d] = ymd.split('-').map(Number); return Date.UTC(y, m - 1, d); };
const daysBetween = (a, b) => Math.round((ymdToUTC(b) - ymdToUTC(a)) / DAY);
const ymdAdd = (ymd, n) => ymdUTC(new Date(ymdToUTC(ymd) + n * DAY));
const weekPlus = (w, n) => isoWeek(ymdAdd(weekRange(w).start, 7 * n));
const todayYMD = () => { const d = new Date(); return `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}`; };
/* % formatado sem arredondar para 100% o que não está completo (99,93% ≠ 100%) nem para 0% o que já começou */
const pctTxt = (v, dec = 1) => {
  if (v == null || isNaN(v)) return '—';
  const f = d => v.toLocaleString('pt-BR', {maximumFractionDigits: d, minimumFractionDigits: d});
  let d = dec; const r = d => Number(v.toFixed(d));
  while (d < 2 && ((v < 100 && r(d) >= 100) || (v > 0 && r(d) <= 0))) d++;
  if (v < 100 && r(d) >= 100) return '99,99%';
  return f(d) + '%';
};
const numTxt = (v, dec = 2) => v == null || isNaN(v) ? '—' : Number(v).toLocaleString('pt-BR', {maximumFractionDigits: dec});

function trkObras(){
  const mp = new Map();
  for (const k of [...TRK, 'labor', 'equip', 'manual']) for (const r of Repo.data[k]) { const o = obraKey(r.obra); if (o && !mp.has(o)) mp.set(o, r.obra); }
  return [...mp].map(([v, l]) => ({v, l})).sort((a, b) => a.l.localeCompare(b.l, 'pt-BR'));
}
function trkWeeks(){
  const s = new Set([todayWeek(), weekPlus(todayWeek(), 1)]);
  for (const k of ['avanco', 'lps', 'labor']) for (const r of Repo.data[k]) if (weekRange(r.semana)) s.add(r.semana);
  return [...s].sort().reverse();
}
function ensureTK(){
  const ob = trkObras();
  if (!ob.some(o => o.v === TK.obra)) TK.obra = (ob.find(o => Repo.data.ativ.some(a => obraKey(a.obra) === o.v)) || ob[0] || {v: ''}).v;
  if (!TK.week || !weekRange(TK.week)) TK.week = todayWeek();
  return ob;
}
function trkPickers(sec, withWeek = true){
  const ob = ensureTK(); const wk = trkWeeks();
  const os = $('[data-tk=obra]', sec), ws = $('[data-tk=week]', sec);
  os.innerHTML = ob.length ? ob.map(o => `<option value="${esc(o.v)}">${esc(o.l)}</option>`).join('') : '<option value="">Nenhuma obra cadastrada</option>';
  os.value = TK.obra;
  if (ws) { ws.innerHTML = wk.map(w => `<option value="${w}">${esc(weekLabel(w, true))}${w === todayWeek() ? ' — atual' : ''}</option>`).join(''); ws.value = TK.week; }
}
const obraLabel = o => (trkObras().find(x => x.v === o) || {l: o}).l;

/* ---------- cálculos: avanço físico ---------- */
function avHist(o){
  const mp = new Map();
  for (const r of Repo.data.avanco) { if (obraKey(r.obra) !== o) continue; const k = norm(r.codigo); if (!mp.has(k)) mp.set(k, []); mp.get(k).push(r); }
  mp.forEach(l => l.sort((a, b) => a.semana.localeCompare(b.semana)));
  return mp;
}
function realAt(a, hist, w){
  const hasV = r => (r.acumulado != null && r.acumulado !== '') || (r.pct != null && r.pct !== '');
  let last = null; for (const r of hist || []) { if (r.semana <= w) { if (hasV(r)) last = r; } else break; }
  if (!last) return {pct: 0, acum: 0, rec: null};
  const q = Number(a.qtd) || 0; const hasAc = last.acumulado != null && last.acumulado !== ''; const hasPct = last.pct != null && last.pct !== '';
  let acum = hasAc ? Number(last.acumulado) : (hasPct && q > 0 ? q * Number(last.pct) / 100 : null);
  let pct = hasAc && q > 0 ? acum / q * 100 : hasPct ? Number(last.pct) : 0;
  return {pct: Math.max(0, Math.min(100, pct)), acum, rec: last};
}
function planAt(a, endYmd){
  if (!a.inicio || !a.fim) return 0;
  if (endYmd < a.inicio) return 0; if (endYmd >= a.fim) return 100;
  return (daysBetween(a.inicio, endYmd) + 1) / (daysBetween(a.inicio, a.fim) + 1) * 100;
}
function obraModel(o, w, hist){
  const ativs = Repo.data.ativ.filter(a => obraKey(a.obra) === o);
  const usePeso = ativs.some(a => Number(a.peso) > 0);
  // serviços sem peso, quando outros têm, recebem o peso médio (não somem do cálculo)
  const pesos = ativs.map(a => Number(a.peso)).filter(v => v > 0); const pesoMed = pesos.length ? pesos.reduce((x, y) => x + y, 0) / pesos.length : 1;
  hist = hist || avHist(o); const end = weekRange(w).end; const w4 = weekPlus(w, -4);
  const rows = ativs.map(a => {
    const h = hist.get(norm(a.codigo)); const r = realAt(a, h, w), r4 = realAt(a, h, w4); const p = planAt(a, end);
    const q = Number(a.qtd) || 0;
    // ritmo: média das últimas 4 semanas; sem histórico de 4 semanas, média desde o início previsto
    let ratePct = Math.max(0, (r.pct - r4.pct) / 4);
    const ini0 = a.inicio_rev || a.inicio;
    if (!r4.rec && r.rec && ini0 && ini0 <= end) ratePct = Math.max(0, r.pct / Math.max(1, daysBetween(ini0, end) / 7));
    let prev = null;
    if (r.pct >= 100) prev = 'Concluída'; else if (ratePct > 0) prev = weekRange(weekPlus(w, Math.ceil((100 - r.pct) / ratePct))).end;
    const alvo = a.revisado || a.fim;
    // atraso pelas datas: término (revisado, senão previsto) vencido sem concluir, ou início previsto passado sem ter começado
    const ref = w === todayWeek() ? todayYMD() : end;
    const ini = a.inicio_rev || a.inicio;
    // iniciada = tem avanço medido OU Início Revisado já chegou (o Início Revisado é a data real de início; futura = reprogramação)
    const iniciada = r.pct > 0 || !!(a.inicio_rev && a.inicio_rev <= ref);
    const termAtr = r.pct < 100 && alvo && ref > alvo, iniAtr = !termAtr && !iniciada && ini && ref > ini;
    const diasAtraso = termAtr ? daysBetween(alvo, ref) : iniAtr ? daysBetween(ini, ref) : 0;
    const status = r.pct >= 100 ? 'Concluída' : termAtr ? 'Término atrasado' : iniAtr ? 'Início atrasado' : !iniciada ? 'Não iniciada' : 'Em andamento';
    const acum = r.acum != null ? r.acum : (q ? q * r.pct / 100 : null);
    const restQtd = q > 0 ? Math.max(0, q - (acum || 0)) : null;
    const semRest = alvo && alvo > end ? Math.max(1, Math.ceil(daysBetween(end, alvo) / 7)) : 0;
    return {a, real: r.pct, plan: p, desvio: r.pct - p, acum, restQtd, restPct: 100 - r.pct, ratePct, ritmoQtd: q > 0 ? ratePct * q / 100 : null,
      ritmoNec: r.pct >= 100 ? 0 : (q > 0 ? (semRest ? restQtd / semRest : null) : (semRest ? (100 - r.pct) / semRest : null)), prev, status, diasAtraso, semMedicao: iniciada && r.pct === 0,
      atrasoPrev: prev && prev !== 'Concluída' && alvo && prev > alvo ? daysBetween(alvo, prev) : 0, alvo,
      reprog: a.revisado && a.fim && a.revisado > a.fim ? daysBetween(a.fim, a.revisado) : 0,
      reprogIni: a.inicio_rev && a.inicio && a.inicio_rev > a.inicio ? daysBetween(a.inicio, a.inicio_rev) : 0,
      rev: revHist(a, h), semMotivo: iniSemMotivo(a) && r.pct < 100,
      w: usePeso ? (Number(a.peso) > 0 ? Number(a.peso) : pesoMed) : 1, lastAv: r.rec};
  });
  const W = rows.reduce((s, x) => s + x.w, 0);
  const real = W ? rows.reduce((s, x) => s + x.w * x.real, 0) / W : 0, plan = W ? rows.reduce((s, x) => s + x.w * x.plan, 0) / W : 0;
  return {rows, real, plan, desvio: real - plan, idp: plan > 0 ? real / plan : null, usePeso, n: rows.length,
    inicio: ativs.map(a => a.inicio).filter(Boolean).sort()[0] || null, fim: ativs.map(a => a.fim).filter(Boolean).sort().pop() || null};
}
function lastAvWeek(o){ let w = ''; for (const r of Repo.data.avanco) if (obraKey(r.obra) === o && r.semana > w) w = r.semana; return w; }
function curvaS(o){
  const m0 = obraModel(o, todayWeek()); if (!m0.inicio || !m0.fim) return null;
  const hist = avHist(o); const lastW = lastAvWeek(o); const tw = todayWeek();
  let a = isoWeek(m0.inicio), b = isoWeek(m0.fim); if (lastW && lastW > b) b = lastW;
  const weeks = weeksBetween(a, b, 260);
  const realLimit = lastW ? (lastW < tw ? lastW : tw) : '';
  const plan = [], real = [];
  for (const w of weeks) { const m = obraModel(o, w, hist); plan.push(+m.plan.toFixed(2)); real.push(realLimit && w <= realLimit ? +m.real.toFixed(2) : null); }
  // efetivo MO no fechamento de cada semana (histograma)
  const C = computeCells('labor', {...blankFilter(), obras: [o]});
  const efet = weeks.map(w => { const c = C.cells.get(w + '||' + o); return c ? c.total : null; });
  return {weeks, plan, real, efet};
}
/* ---------- cálculos: Last Planner ---------- */
function ppcWeek(o, w){
  const l = Repo.data.lps.filter(r => obraKey(r.obra) === o && r.semana === w);
  const concl = l.filter(r => r.status === 'Concluído').length, nao = l.filter(r => r.status === 'Não concluído').length;
  return {total: l.length, concl, nao, pend: l.length - concl - nao, ppc: concl + nao ? concl / (concl + nao) * 100 : null, list: l};
}
function ppcSeries(o, w, n = 12){ const out = []; for (let i = n - 1; i >= 0; i--) { const ww = weekPlus(w, -i); out.push({w: ww, ...ppcWeek(o, ww)}); } return out; }
function causasPareto(o, w, n = 12){
  const from = weekPlus(w, -(n - 1)); const c = {};
  for (const r of Repo.data.lps) if (obraKey(r.obra) === o && r.semana >= from && r.semana <= w && r.status === 'Não concluído') { const k = r.causa || 'Sem causa informada'; c[k] = (c[k] || 0) + 1; }
  return Object.entries(c).sort((a, b) => b[1] - a[1]);
}
/* ---------- cálculos: restrições ---------- */
const restrAberta = r => r.status !== 'Removida';
function restrStats(o){
  const t = todayYMD(); const l = Repo.data.restr.filter(r => obraKey(r.obra) === o);
  const ab = l.filter(restrAberta), rem = l.filter(r => !restrAberta(r));
  const venc = ab.filter(r => r.prazo && r.prazo < t), prox = ab.filter(r => r.prazo && r.prazo >= t && r.prazo <= ymdAdd(t, 7));
  const remNoPrazo = rem.filter(r => !r.prazo || !r.remocao || r.remocao <= r.prazo).length;
  return {list: l, ab, rem, venc, prox, total: l.length, irr: rem.length ? remNoPrazo / rem.length * 100 : null};
}

/* ---------- gráficos ---------- */
function chartCurvaS(id, cv, cs, opts = {}){
  killChart(id); if (!cs) return;
  const ds = [
    {type: 'line', label: 'Planejado', data: cs.plan, borderColor: '#94A3B8', backgroundColor: '#94A3B8', borderDash: [6, 4], borderWidth: 2.5, pointRadius: 0, tension: .25, yAxisID: 'y', order: 1},
    {type: 'line', label: 'Realizado', data: cs.real, borderColor: '#F7931E', backgroundColor: 'rgba(247,147,30,.15)', fill: true, borderWidth: 3.5, pointRadius: 3, pointBackgroundColor: '#fff', tension: .25, yAxisID: 'y', spanGaps: true, order: 0},
  ];
  if (opts.efetivo !== false && cs.efet.some(v => v != null)) ds.push({type: 'bar', label: 'Efetivo MO (histograma)', data: cs.efet, backgroundColor: 'rgba(51,65,85,.22)', borderRadius: 4, yAxisID: 'y2', order: 2, maxBarThickness: 18});
  charts[id] = new Chart(cv, curvaSConfig(cs, ds, false));
}
function curvaSDatasets(cs, efetivo = true){
  const ds = [
    {type: 'line', label: 'Planejado', data: cs.plan, borderColor: '#94A3B8', backgroundColor: '#94A3B8', borderDash: [6, 4], borderWidth: 2.5, pointRadius: 0, tension: .25, yAxisID: 'y', order: 1},
    {type: 'line', label: 'Realizado', data: cs.real, borderColor: '#F7931E', backgroundColor: 'rgba(247,147,30,.15)', fill: true, borderWidth: 3.5, pointRadius: 3, pointBackgroundColor: '#fff', tension: .25, yAxisID: 'y', spanGaps: true, order: 0},
  ];
  if (efetivo && cs.efet.some(v => v != null)) ds.push({type: 'bar', label: 'Efetivo MO (histograma)', data: cs.efet, backgroundColor: 'rgba(51,65,85,.22)', borderRadius: 4, yAxisID: 'y2', order: 2, maxBarThickness: 18});
  return ds;
}
function curvaSConfig(cs, ds, print){
  const tw = cs.weeks.indexOf(TK.week);
  return {type: 'bar', data: {labels: cs.weeks.map(w => weekShort(w)), datasets: ds},
    options: {responsive: !print, maintainAspectRatio: false, animation: print || reduceMotion ? false : {duration: 900}, devicePixelRatio: print ? 2 : undefined,
      interaction: {mode: 'index', intersect: false},
      scales: {x: {grid: {display: false}, ticks: {maxRotation: 0, autoSkip: true, font: {size: 10, weight: 'bold'}}},
        y: {min: 0, max: 100, ticks: {callback: v => v + '%', color: '#9CA3AF'}, grid: {color: '#F0F1F4'}, border: {display: false}},
        y2: {position: 'right', beginAtZero: true, grid: {display: false}, ticks: {precision: 0, color: '#9CA3AF'}, display: ds.length > 2, title: {display: ds.length > 2, text: 'pessoas', color: '#9CA3AF', font: {size: 10}}}},
      plugins: {legend: {position: 'top', align: 'end', labels: {usePointStyle: true, font: {weight: 'bold'}}},
        tooltip: {callbacks: {title: it => weekLabel(cs.weeks[it[0].dataIndex], true), label: it => ` ${it.dataset.label}: ${it.raw == null ? '—' : it.dataset.yAxisID === 'y2' ? fmtN(it.raw) : pctTxt(it.raw)}`}},
        weekMark: {idx: tw}}}, plugins: [weekMarkPlugin]};
}
const weekMarkPlugin = {id: 'weekMark', afterDatasetsDraw(ch, a, o){ if (!o || o.idx == null || o.idx < 0) return; const x = ch.scales.x.getPixelForValue(o.idx); const {top, bottom} = ch.chartArea; const c = ch.ctx; c.save(); c.strokeStyle = 'rgba(220,38,38,.55)'; c.setLineDash([3, 3]); c.beginPath(); c.moveTo(x, top); c.lineTo(x, bottom); c.stroke(); c.fillStyle = '#DC2626'; c.font = 'bold 10px Arial'; c.textAlign = 'center'; c.fillText('semana de referência', x, top + 10); c.restore(); }};
function chartPPC(id, cv, ser){
  killChart(id);
  charts[id] = new Chart(cv, {type: 'bar', data: {labels: ser.map(x => weekShort(x.w)), datasets: [
      {label: 'PPC', data: ser.map(x => x.ppc == null ? null : +x.ppc.toFixed(1)), backgroundColor: ser.map(x => x.ppc == null ? '#E5E7EB' : x.ppc >= 80 ? '#16A34A' : x.ppc >= 60 ? '#F7931E' : '#DC2626'), borderRadius: 6, maxBarThickness: 34},
      {type: 'line', label: 'Meta 80%', data: ser.map(() => 80), borderColor: '#111827', borderDash: [5, 4], borderWidth: 1.5, pointRadius: 0, noLabel: true}]},
    options: {responsive: true, maintainAspectRatio: false, animation: reduceMotion ? false : {duration: 800},
      scales: {y: {min: 0, max: 100, ticks: {callback: v => v + '%'}, grid: {color: '#F0F1F4'}, border: {display: false}}, x: {grid: {display: false}, ticks: {font: {size: 10, weight: 'bold'}}}},
      plugins: {legend: {display: false}, valueLabels: {enabled: true, size: 10},
        tooltip: {callbacks: {title: it => weekLabel(ser[it[0].dataIndex].w, true), afterBody: it => { const x = ser[it[0].dataIndex]; return x.total ? `${x.concl} concluído(s) de ${x.concl + x.nao} avaliado(s)${x.pend ? ` · ${x.pend} sem avaliação` : ''}` : 'Sem compromissos'; }}}}}});
}
function chartHBar(id, cv, labels, data, color){
  killChart(id);
  charts[id] = new Chart(cv, {type: 'bar', data: {labels: labels.map(l => wrapText(l, 26, 2)), datasets: [{data, backgroundColor: color, borderRadius: 5, maxBarThickness: 22}]},
    options: {indexAxis: 'y', responsive: true, maintainAspectRatio: false, animation: reduceMotion ? false : {duration: 700}, layout: {padding: {right: 26}},
      scales: {x: {beginAtZero: true, ticks: {precision: 0}, grid: {color: '#F0F1F4'}, border: {display: false}}, y: {grid: {display: false}, ticks: {font: {size: 11}}}},
      plugins: {legend: {display: false}, tooltip: {callbacks: {title: it => labels[it[0].dataIndex]}}}},
    plugins: [{id: 'hl', afterDatasetsDraw(ch){ const c = ch.ctx; c.save(); c.font = 'bold 11px Arial'; c.fillStyle = '#111'; c.textBaseline = 'middle'; ch.getDatasetMeta(0).data.forEach((el, i) => c.fillText(fmtN(data[i]), el.x + 6, el.y)); c.restore(); }}]});
}

/* ---------- componentes visuais ---------- */
const isAtr = x => x.status === 'Término atrasado' || x.status === 'Início atrasado';
const stChip = (st, dias) => `<span class="chip ${st === 'Concluída' ? 'ok' : /atrasad/.test(st) ? 'err' : st === 'Em andamento' ? 'or' : ''}">${esc(st)}${dias ? ` · ${dias} d` : ''}</span>`;
function progBar(real, plan){
  return `<div class="pbar" title="Realizado ${pctTxt(real)} · Planejado ${pctTxt(plan)}"><i style="width:${Math.max(0, Math.min(100, real))}%;background:${real + 5 < plan ? '#DC2626' : real >= 100 ? '#16A34A' : '#F7931E'}"></i><b style="left:${Math.max(0, Math.min(100, plan))}%"></b></div>`;
}
function kpiBox(l, v, s, c, dark){ return `<div class="kpi ${dark ? 'dark' : ''}" style="--c:${c || '#F7931E'}"><div class="l">${l}</div><div class="v" style="font-size:24px">${v}</div><div class="s">${s || ''}</div></div>`; }
const idpColor = v => v == null ? '#9CA3AF' : v >= .95 ? '#16A34A' : v >= .85 ? '#F7931E' : '#DC2626';
function trkEmpty(msg){ return `<div class="card"><div class="empty"><b>${msg}</b>Cadastre atividades pelo botão “Nova atividade” ou importe a planilha com a aba <b>Atividades</b> (baixe o Modelo Excel).</div></div>`; }
const trkToolbar = (title, sub, acts, withWeek = true) => `<div class="toolbar"><div><h2>${title}</h2><div class="sub">${sub}</div></div><div class="acts">${acts}</div></div>
  <div class="card filters" style="grid-template-columns:repeat(auto-fit,minmax(220px,1fr))"><label>Obra<select data-tk="obra"></select></label>${withWeek ? '<label>Semana de referência<select data-tk="week"></select></label>' : ''}<div data-extra-filters style="display:contents"></div></div>`;
const BTN_IMP = '<button class="btn" data-ta="import">⇪ Importar planilha</button><button class="btn" data-ta="template">⬇ Modelo Excel</button>';

/* ---------- PAINEL DA OBRA ---------- */
function renderPainel(){
  const sec = $('#tab-painel');
  if (!sec.dataset.built) { sec.dataset.built = 1; sec.innerHTML = trkToolbar('Painel da obra', 'Visão única do acompanhamento: avanço físico, prazo, planejamento semanal, restrições e efetivo', BTN_IMP) + '<div data-body></div>'; }
  trkPickers(sec);
  const o = TK.obra, w = TK.week, body = $('[data-body]', sec);
  Object.keys(charts).filter(k => k.startsWith('pn-')).forEach(killChart);
  if (!o) { body.innerHTML = trkEmpty('Nenhuma obra cadastrada'); return; }
  const m = obraModel(o, w); const pc = ppcWeek(o, w); const ser = ppcSeries(o, w, 12); const rs = restrStats(o);
  const ppc4 = ser.slice(-4).filter(x => x.ppc != null); const ppcMed = ppc4.length ? ppc4.reduce((s, x) => s + x.ppc, 0) / ppc4.length : null;
  const C = computeCells('labor', {...blankFilter(), obras: [o]}); const E = computeCells('equip', {...blankFilter(), obras: [o]});
  const pick = X => { const ws = sortedWeeks(X.cells).filter(x => x <= w); for (let i = ws.length - 1; i >= 0; i--) { const c = X.cells.get(ws[i] + '||' + o); if (c) return {c, w: ws[i]}; } return null; };
  const ef = pick(C), eq = pick(E);
  const atras = m.rows.filter(isAtr), rest = m.rows.filter(x => x.status !== 'Concluída');
  // previsão de término da obra (ritmo das últimas 4 semanas)
  const m4 = obraModel(o, weekPlus(w, -4)); const rate = (m.real - m4.real) / 4;
  const prevFim = m.real >= 100 ? 'Concluída' : rate > 0 ? weekRange(weekPlus(w, Math.ceil((100 - m.real) / rate))).end : null;
  const atrasoObra = prevFim && prevFim !== 'Concluída' && m.fim && prevFim > m.fim ? daysBetween(m.fim, prevFim) : 0;
  body.innerHTML = `
    <div class="kpis">
      ${kpiBox('Avanço físico realizado', m.n ? pctTxt(m.real) : '—', m.n ? `planejado ${pctTxt(m.plan)} · desvio <b style="color:${m.desvio < -5 ? '#FCA5A5' : '#86EFAC'}">${m.desvio >= 0 ? '+' : ''}${pctTxt(m.desvio)}</b>` : 'sem atividades cadastradas', '#F7931E', true)}
      ${kpiBox('IDP (índice de prazo)', m.idp == null ? '—' : m.idp.toLocaleString('pt-BR', {maximumFractionDigits: 2, minimumFractionDigits: 2}), m.idp == null ? 'sem planejado na semana' : m.idp >= .95 ? 'no prazo' : m.idp >= .85 ? 'atenção' : 'atrasado', idpColor(m.idp))}
      ${kpiBox('PPC da semana', pc.ppc == null ? '—' : pctTxt(pc.ppc, 0), pc.total ? `${pc.concl} de ${pc.concl + pc.nao} compromisso(s)${pc.pend ? ` · ${pc.pend} sem avaliação` : ''} · média 4 sem.: ${pctTxt(ppcMed, 0)}` : 'sem compromissos na semana', pc.ppc == null ? '#9CA3AF' : pc.ppc >= 80 ? '#16A34A' : '#DC2626')}
      ${kpiBox('Restrições abertas', fmtN(rs.ab.length), `<b style="color:#DC2626">${rs.venc.length} vencida(s)</b> · ${rs.prox.length} vencem em 7 dias`, rs.venc.length ? '#DC2626' : '#16A34A')}
      ${kpiBox('Atividades restantes', fmtN(rest.length), `${atras.length} atrasada(s) · ${m.n - rest.length} concluída(s) de ${m.n}`, atras.length ? '#DC2626' : '#334155')}
      ${kpiBox('Efetivo no fechamento', ef ? fmtN(ef.c.total) : '—', `${ef ? 'mão de obra · ' + weekShort(ef.w) : 'sem histograma'}${eq ? ` · ${fmtN(eq.c.total)} equip.` : ''}`, '#334155')}
      ${kpiBox('Término previsto', m.fim ? fmtBR(m.fim) : '—', prevFim === 'Concluída' ? 'obra concluída' : prevFim ? `no ritmo atual: <b style="color:${atrasoObra ? '#DC2626' : '#16A34A'}">${fmtBR(prevFim)}</b>${atrasoObra ? ` (+${atrasoObra} dias)` : ''}` : 'sem ritmo nas últimas 4 semanas', atrasoObra ? '#DC2626' : '#16A34A')}
    </div>
    <div class="pn-grid">
      <div class="card"><div class="card-h"><h3>Curva S — avanço físico × efetivo</h3><span class="hint">${m.usePeso ? 'ponderado pelo peso das atividades' : 'atividades com peso igual (informe a coluna Peso para ponderar)'}</span></div><div class="chart-frame"><div style="position:relative;height:300px"><canvas id="pnS"></canvas></div></div></div>
      <div class="card"><div class="card-h"><h3>PPC — últimas 12 semanas</h3><span class="hint">meta ≥ 80%</span></div><div class="chart-frame"><div style="position:relative;height:300px"><canvas id="pnP"></canvas></div></div></div>
    </div>
    <div class="pn-grid">
      <div class="card"><div class="card-h"><h3>Atividades críticas</h3><span class="hint">maior atraso em relação ao planejado</span><span class="sp"></span><button class="btn sm" data-goto="ativ">Ver todas</button></div>${critTable(m)}</div>
      <div class="card"><div class="card-h"><h3>Restrições que travam a obra</h3><span class="hint">vencidas e próximas 7 dias</span><span class="sp"></span><button class="btn sm" data-goto="restr">Ver todas</button></div>${restrMini([...rs.venc, ...rs.prox].slice(0, 8))}</div>
    </div>
    <div class="card"><div class="card-h"><h3>Causas de não cumprimento</h3><span class="hint">últimas 12 semanas — atacar as primeiras resolve a maior parte das falhas</span></div><div style="position:relative;height:${Math.max(140, causasPareto(o, w).length * 34 + 40)}px"><canvas id="pnC"></canvas></div></div>`;
  const cs = curvaS(o);
  if (cs) chartCurvaS('pn-s', $('#pnS'), cs); else $('#pnS').parentElement.innerHTML = '<div class="empty">Cadastre atividades com início e término para ver a curva S.</div>';
  chartPPC('pn-p', $('#pnP'), ser);
  const cp = causasPareto(o, w); if (cp.length) chartHBar('pn-c', $('#pnC'), cp.map(x => x[0]), cp.map(x => x[1]), '#F7931E'); else $('#pnC').parentElement.innerHTML = '<div class="muted" style="font-size:13px">Nenhum compromisso não concluído nas últimas 12 semanas.</div>';
}
function critTable(m){
  const l = m.rows.filter(x => x.status !== 'Concluída').sort((a, b) => a.desvio - b.desvio || (b.atrasoPrev - a.atrasoPrev)).slice(0, 8);
  if (!l.length) return '<div class="muted" style="font-size:13px">Nenhuma atividade restante.</div>';
  return `<div class="tbl-wrap" style="max-height:none"><table><thead><tr><th>Código</th><th>Atividade</th><th>Avanço</th><th>Desvio</th><th>Término</th><th>Status</th></tr></thead><tbody>
    ${l.map(x => `<tr><td>${esc(x.a.codigo)}</td><td class="l" style="white-space:normal;min-width:180px">${esc(x.a.nome)}</td><td style="min-width:120px">${progBar(x.real, x.plan)}<small class="muted">${pctTxt(x.real, 0)} / ${pctTxt(x.plan, 0)}</small></td><td class="num" style="color:${x.desvio < -5 ? '#DC2626' : '#16A34A'}">${x.desvio >= 0 ? '+' : ''}${pctTxt(x.desvio, 0)}</td><td>${fmtBR(x.a.fim)}</td><td>${stChip(x.status)}</td></tr>`).join('')}</tbody></table></div>`;
}
function restrMini(l){
  if (!l.length) return '<div class="muted" style="font-size:13px">Nenhuma restrição vencida ou vencendo nos próximos 7 dias. 👍</div>';
  const t = todayYMD();
  return `<div class="tbl-wrap" style="max-height:none"><table><thead><tr><th>Restrição</th><th>Tipo</th><th>Responsável</th><th>Data necessária</th><th></th></tr></thead><tbody>
    ${l.map(r => `<tr><td class="l" style="white-space:normal;min-width:200px"><b>${esc(r.descricao)}</b>${r.atividade || r.codigo ? `<br><small class="muted">${esc([r.codigo, r.atividade].filter(Boolean).join(' · '))}</small>` : ''}</td><td>${esc(r.tipo || '')}</td><td>${esc(r.responsavel || '')}</td><td><span class="chip ${r.prazo < t ? 'err' : 'warn'}">${fmtBR(r.prazo)}</span></td><td><button class="btn sm" data-ta="rmrestr" data-id="${r.id}">✓ Removida</button></td></tr>`).join('')}</tbody></table></div>`;
}

/* ---------- ATIVIDADES ---------- */
/* ---------- Atividades: histograma das 3 últimas semanas (MO e equipamentos) ---------- */
function histo3Data(kind, o, w){
  const K = KINDS[kind]; const weeks = [weekPlus(w, -2), weekPlus(w, -1), w];
  const C = computeCells(kind, {...blankFilter(), obras: [o]});
  const cells = weeks.map(wk => C.cells.get(wk + '||' + o) || null);
  return {kind, K, weeks, cells, totals: cells.map(c => c ? c.total : null), any: cells.some(c => c && c.total)};
}
function histo3Model(d){
  const cats = d.K.cats.map(c => ({k: c.k, l: c.l, color: c.color, data: d.cells.map(x => x ? (x.unknown ? null : (x[c.k] || 0)) : null)}));
  if (d.cells.some(x => x && x.SD)) cats.push({k: 'SD', l: SD.l, color: '#9CA3AF', data: d.cells.map(x => x ? (x.SD || 0) : null)});
  return {cats, jm: journeyModel(d.weeks, d.totals, cats)};
}
function histo3Config(d, print){
  const {cats, jm} = histo3Model(d); const n = d.weeks.length;
  const ds = [{label: 'Total', data: d.totals, borderColor: '#111827', borderWidth: print ? 4 : 3.5, tension: 0, spanGaps: false,
    pointRadius: print ? 5 : 4, pointBackgroundColor: '#fff', pointBorderColor: '#111827', pointBorderWidth: 2, pointHoverRadius: 6,
    fill: true, backgroundColor: c => { const a = c.chart.chartArea; if (!a) return 'rgba(247,147,30,.15)'; const g = c.chart.ctx.createLinearGradient(0, a.top, 0, a.bottom); g.addColorStop(0, 'rgba(247,147,30,.30)'); g.addColorStop(1, 'rgba(247,147,30,0.02)'); return g; }, order: 0}];
  cats.forEach(c => ds.push({label: c.l, data: c.data, borderColor: c.color, backgroundColor: c.color, borderWidth: 2, tension: 0, pointRadius: 3, pointHoverRadius: 5, borderDash: c.k === 'SD' ? [5, 4] : undefined, fill: false, spanGaps: false, order: 1}));
  return {type: 'line',
    plugins: [{id: 'h3jr', beforeInit(c){ if (jm) c.$j = {n, reveal: n - 1, cursor: false, miles: jm.miles}; }}],
    data: {labels: d.weeks.map(w => [weekShort(w), fmtBR(weekRange(w).end).slice(0, 5)]), datasets: ds},
    options: {responsive: !print, maintainAspectRatio: false, animation: print || reduceMotion ? false : {duration: 600}, devicePixelRatio: print ? 2 : undefined,
      layout: {padding: {top: 38, right: 18, left: 8, bottom: 4}},
      interaction: {mode: 'index', intersect: false},
      scales: {x: {offset: true, grid: {display: false}, border: {color: '#E5E7EB'}, ticks: {maxRotation: 0, font: {size: print ? 13 : 11, weight: 'bold'}}},
        y: {beginAtZero: true, grace: '22%', border: {display: false}, grid: {color: '#F0F1F4'}, ticks: {precision: 0, color: '#9CA3AF', maxTicksLimit: 6}}},
      plugins: {legend: {position: 'bottom', labels: {usePointStyle: true, boxHeight: 7, font: {weight: 'bold', size: print ? 12 : 11}}},
        tooltip: {callbacks: {title: it => weekLabel(d.weeks[it[0].dataIndex], true), label: it => ` ${it.dataset.label}: ${it.raw == null ? 'sem registro' : fmtN(it.raw)}`}}}}};
}
/* faixa preta “FINAL · valor · semana · vs. início”, igual ao percurso das abas de histograma */
function histo3Read(d){
  const {jm} = histo3Model(d); const T = d.totals;
  const iE = T.map((v, i) => v == null ? -1 : i).filter(i => i >= 0).pop(); if (iE == null || iE < 0) return '';
  const iS = T.findIndex(v => v != null);
  const m = jm && jm.miles.find(x => x.idx === iE); const k = m ? m.kinds[m.kinds.length - 1] : 'end';
  const tag = m ? m.kinds.map(x => MILE[x].l).join(' · ') : 'FINAL';
  const dv = iS !== iE ? T[iE] - T[iS] : null; const dp = dv != null ? pct(T[iS], T[iE]) : null;
  return `<div class="jr-read h3-read"><span class="rtag" style="background:${MILE[k].c}">${tag}</span><span class="rv">${fmtN(T[iE])}</span><span class="ru">${esc(d.K.unit)}</span><span class="rw">${esc(weekLabel(d.weeks[iE]))}</span>${dv != null ? `<span class="rd">vs. início ${dv >= 0 ? '+' : ''}${fmtN(dv)}${dp != null ? ` (${dp >= 0 ? '+' : ''}${dp.toLocaleString('pt-BR', {maximumFractionDigits: 1})}%)` : ''}</span>` : ''}</div>`;
}
function histo3Table(d){
  const v = (c, k) => c ? (c.unknown && k !== 'SD' ? '—' : fmtN(c[k] || 0)) : '—';
  const rowsC = d.K.cats.map(c => `<tr><td class="l">${esc(c.l)}</td>${d.cells.map(x => `<td class="num">${v(x, c.k)}</td>`).join('')}</tr>`).join('');
  const sd = d.cells.some(x => x && x.SD) ? `<tr><td class="l">Sem detalhamento</td>${d.cells.map(x => `<td class="num">${x && x.SD ? fmtN(x.SD) : '—'}</td>`).join('')}</tr>` : '';
  const dv = d.totals[2] != null && d.totals[1] != null ? d.totals[2] - d.totals[1] : null;
  return `<table class="h3t"><thead><tr><th class="l">${esc(d.K.catTitle)}</th>${d.weeks.map(x => `<th>${esc(weekShort(x))}</th>`).join('')}</tr></thead><tbody>${rowsC}${sd}
    <tr class="tot"><td class="l">Total</td>${d.totals.map(t => `<td class="num">${t == null ? '—' : fmtN(t)}</td>`).join('')}</tr></tbody></table>
    ${dv != null ? `<div class="muted" style="font-size:12px;margin-top:4px">Variação na última semana: <b style="color:${dv > 0 ? '#16A34A' : dv < 0 ? '#DC2626' : 'inherit'}">${dv > 0 ? '+' : ''}${fmtN(dv)}</b> ${esc(d.K.unit)}</div>` : ''}`;
}

/* ---------- Atividades: colunas configuráveis ---------- */
const LS_ATCOLS = 'cosampa_ativ_cols_ocultas';
const dateRevCell = (prev, rev, dias, lst, semanas, print) => {
  if (!rev || rev === prev) return fmtBR(rev || prev) || '—';
  const n = lst.length;
  if (print) return `<b>${fmtBR(rev)}</b><span class="d">${dias ? `+${dias} dias · ` : ''}revisado</span>`;
  return `<span title="${esc((prev ? 'Previsto: ' + fmtBR(prev) : 'Sem data prevista') + (semanas.length ? '\nHistórico:\n' + semanas.join('\n') : ''))}"><b>${fmtBR(rev)}</b><br><small class="chip or" style="padding:1px 6px">revisado</small>${dias ? ` <small style="color:#DC2626;font-weight:700">+${dias} d</small>` : ''}${n > 1 ? ` <small class="chip err" style="padding:1px 6px">${n}ª rev.</small>` : ''}${prev ? `<br><small class="muted">prev. ${fmtBR(prev)}</small>` : ''}</span>`;
};
const AT_COLS = [
  {k: 'trecho', l: 'Trecho', pw: 9.5, td: x => `<td class="l">${esc(x.a.trecho || '')}</td>`, p: (x, c2) => `<td class="l">${c2(x.a.trecho || '')}</td>`},
  {k: 'nome', l: 'Serviço', fixed: 1, pw: 21, td: x => `<td class="l" style="white-space:normal;min-width:240px"><b>${esc(x.a.nome)}</b></td>`, p: (x, c2) => `<td class="l">${c2(x.a.nome)}</td>`},
  {k: 'tipo', l: 'Tipo', pw: 6, td: x => `<td>${esc(x.a.tipo || '')}</td>`, p: (x, c2) => `<td>${c2(x.a.tipo || '')}</td>`},
  {k: 'unidade', l: 'Und.', pw: 3.5, td: x => `<td>${esc(x.a.unidade || '')}</td>`, p: x => `<td>${esc(x.a.unidade || '')}</td>`},
  {k: 'qtd', l: 'Qtd. Prevista', pl: 'Qtd. Prev.', pw: 5, td: x => { const q = Number(x.a.qtd) || 0; return `<td class="num">${q ? numTxt(q) : '—'}</td>`; }, p: x => { const q = Number(x.a.qtd) || 0; return `<td>${q ? numTxt(q) : '—'}</td>`; }},
  {k: 'realizada', l: 'Qtd. Realizada', pl: 'Qtd. Real.', pw: 5, td: x => { const q = Number(x.a.qtd) || 0; return `<td class="num">${q ? numTxt(x.acum || 0) : (x.lastAv ? pctTxt(x.real, 0) : '—')}</td>`; }, p: x => { const q = Number(x.a.qtd) || 0; return `<td>${q ? numTxt(x.acum || 0) : (x.lastAv ? pctTxt(x.real, 0) : '—')}</td>`; }},
  {k: 'restante', l: 'Qtd. Restante', pl: 'Qtd. Rest.', pw: 5, td: x => { const q = Number(x.a.qtd) || 0; return `<td class="num"><b>${q ? numTxt(x.restQtd) : (x.lastAv ? pctTxt(x.restPct, 0) : '—')}</b></td>`; }, p: x => { const q = Number(x.a.qtd) || 0; return `<td><b>${q ? numTxt(x.restQtd) : (x.lastAv ? pctTxt(x.restPct, 0) : '—')}</b></td>`; }},
  {k: 'avanco', l: 'Avanço do serviço (realizado × planejado)', pl: '% Real. × Plan.', pw: 7, td: x => { const q = Number(x.a.qtd) || 0, un = x.a.unidade || '';
      return `<td style="min-width:170px" title="Percentual do próprio serviço. Realizado = Qtd. Realizada ÷ Qtd. Prevista. Planejado = parte do prazo previsto (linha de base) já decorrida no fim da semana${q ? ` — equivale a ${numTxt(q * x.plan / 100)} de ${numTxt(q)} ${esc(un)}` : ''}.">${progBar(x.real, x.plan)}<small class="muted">Realizado <b style="color:var(--tx)">${pctTxt(x.real, 0)}</b> · Planejado ${pctTxt(x.plan, 0)}${q && x.plan > 0 ? `<br>planejado ≈ ${numTxt(q * x.plan / 100)} de ${numTxt(q)} ${esc(un)}` : ''}</small></td>`; },
    p: x => `<td><b>${pctTxt(x.real, 0)}</b> <span style="color:#6B7280">× ${pctTxt(x.plan, 0)}</span></td>`},
  {k: 'inicio', l: 'Início', pw: 7, td: x => `<td>${dateRevCell(x.a.inicio, x.a.inicio_rev, x.reprogIni, x.rev.ini, x.rev.semanas)}</td>`, p: x => `<td>${dateRevCell(x.a.inicio, x.a.inicio_rev, x.reprogIni, x.rev.ini, [], true)}</td>`},
  {k: 'fim', l: 'Término', pw: 7, td: x => `<td>${dateRevCell(x.a.fim, x.a.revisado, x.reprog, x.rev.fim, x.rev.semanas)}</td>`, p: x => `<td>${dateRevCell(x.a.fim, x.a.revisado, x.reprog, x.rev.fim, [], true)}</td>`},
  {k: 'status', l: 'Status', pw: 9, td: x => `<td>${stChip(x.status, x.diasAtraso)}${x.semMedicao ? '<br><small class="muted">sem medição de avanço</small>' : ''}</td>`,
    p: (x, c2) => { const cls = /atrasad/.test(x.status) ? 'st-err' : x.status === 'Concluída' ? 'st-ok' : x.status === 'Em andamento' ? 'st-or' : '';
      return `<td><span class="c2 ${cls}">${esc(x.status)}${x.diasAtraso ? ` · ${x.diasAtraso} d` : ''}</span>${x.semMedicao ? '<span class="d" style="color:#6B7280;font-weight:400">sem medição</span>' : ''}</td>`; }},
  {k: 'impedimentos', l: 'Impedimentos', pl: 'Impedi&shy;mentos', pw: 7, td: x => `<td class="l" style="white-space:normal;min-width:120px">${x.a.impedimentos ? `<span class="chip warn">${esc(x.a.impedimentos)}</span>` : x.semMotivo ? '<span class="chip err" title="Início Revisado diferente do previsto sem impedimento/motivo">⚠ início reprogramado sem motivo</span>' : ''}</td>`,
    p: (x, c2) => `<td>${x.a.impedimentos ? c2(x.a.impedimentos) : x.semMotivo ? '<span class="d">⚠ sem motivo</span>' : ''}</td>`},
];
const LS_ATKPI = 'cosampa_ativ_kpis';
function atShowKpis(){ try { return localStorage.getItem(LS_ATKPI) === '1'; } catch (e) { return !!TK._kpis; } }
function atToggleKpis(){ const v = !atShowKpis(); TK._kpis = v; try { localStorage.setItem(LS_ATKPI, v ? '1' : '0'); } catch (e) {} refreshAll(); }
function atHidden(){ try { const v = JSON.parse(localStorage.getItem(LS_ATCOLS) || '[]'); return new Set(Array.isArray(v) ? v : []); } catch (e) { return new Set(TK._colsHidden || []); } }
function atSetHidden(set){ TK._colsHidden = [...set]; try { localStorage.setItem(LS_ATCOLS, JSON.stringify([...set])); } catch (e) {} }
function atVisibleCols(){ const h = atHidden(); return AT_COLS.filter(c => c.fixed || !h.has(c.k)); }
function openAtCols(){
  const h = atHidden();
  const body = `<p class="muted" style="font-size:13px;margin:0 0 10px">Marque as colunas que quer ver na tabela e no PDF. A escolha fica salva neste navegador.</p>
    <div style="display:grid;grid-template-columns:repeat(auto-fill,minmax(220px,1fr));gap:6px 14px">${AT_COLS.map(c => `<label style="display:flex;gap:8px;align-items:center;font-size:14px${c.fixed ? ';opacity:.6' : ''}"><input type="checkbox" data-col="${c.k}" ${c.fixed || !h.has(c.k) ? 'checked' : ''} ${c.fixed ? 'disabled' : ''}> ${esc(c.l)}${c.fixed ? ' <small class="muted">(sempre)</small>' : ''}</label>`).join('')}</div>`;
  const md = openModal({title: 'Colunas da tabela de atividades', body, foot: '<button class="btn ghost" data-all>Mostrar todas</button><span style="flex:1"></span><button class="btn" data-x>Cancelar</button><button class="btn primary" data-ok>Aplicar</button>'});
  $('[data-x]', md.el).onclick = md.close;
  $('[data-all]', md.el).onclick = () => $$('[data-col]', md.el).forEach(i => i.checked = true);
  $('[data-ok]', md.el).onclick = () => { const s = new Set($$('[data-col]', md.el).filter(i => !i.checked).map(i => i.dataset.col)); atSetHidden(s); md.close(); refreshAll(); };
}

function renderAtivPage(){
  const sec = $('#tab-ativ');
  if (!sec.dataset.built) {
    sec.dataset.built = 1;
    sec.innerHTML = trkToolbar('Atividades', 'Lançamento e acompanhamento dos serviços: atividades restantes, prazos e impedimentos', '<button class="btn primary" data-ta="new" data-kind="ativ">＋ Nova atividade</button><button class="btn dark" data-ta="avanco">✎ Lançar avanço da semana</button>' + BTN_IMP + '<button class="btn" data-ta="xlsativ">⊞ Exportar Excel</button><button class="btn dark" data-ta="pdfativ">⎙ PDF</button>') + '<div data-body></div>';
    $('[data-extra-filters]', sec).innerHTML = `<label>Trecho<select data-af="disc"></select></label><label>Mostrar<select data-af="status"><option value="rest">Atividades restantes</option><option value="atr">Atrasadas</option><option value="imp">Com impedimento</option><option value="rep">Reprogramadas</option><option value="sm">Início reprogramado sem motivo</option><option value="and">Em andamento</option><option value="nao">Não iniciadas</option><option value="conc">Concluídas</option><option value="">Todas</option></select></label><label>Busca<input type="search" data-af="q" placeholder="Serviço, trecho, tipo, impedimento…"></label>`;
  }
  trkPickers(sec);
  const o = TK.obra, w = TK.week, body = $('[data-body]', sec); killChart('at-s'); killChart('at-h-labor'); killChart('at-h-equip');
  const m = obraModel(o, w);
  const discs = [...new Set(m.rows.map(x => x.a.trecho).filter(Boolean))].sort((a, b) => a.localeCompare(b, 'pt-BR'));
  const ds = $('[data-af=disc]', sec); ds.innerHTML = '<option value="">Todas</option>' + discs.map(d => `<option>${esc(d)}</option>`).join(''); ds.value = discs.includes(TK.ativ.disc) ? TK.ativ.disc : ''; TK.ativ.disc = ds.value;
  $('[data-af=status]', sec).value = TK.ativ.status; $('[data-af=q]', sec).value = TK.ativ.q;
  if (!o || !m.n) { body.innerHTML = trkEmpty(o ? 'Nenhuma atividade cadastrada para esta obra' : 'Nenhuma obra cadastrada'); return; }
  const F = TK.ativ; const terms = norm(F.q).split(' ').filter(Boolean);
  const map = {rest: x => x.status !== 'Concluída', atr: isAtr, imp: x => !!(x.a.impedimentos && x.status !== 'Concluída'), rep: x => !!(x.reprog || x.reprogIni), sm: x => x.semMotivo, and: x => x.status === 'Em andamento', nao: x => x.status === 'Não iniciada', conc: x => x.status === 'Concluída', '': () => true};
  const rows = m.rows.filter(x => (!F.disc || x.a.trecho === F.disc) && (map[F.status] || map.rest)(x) && (!terms.length || terms.every(t => norm([x.a.trecho, x.a.nome, x.a.tipo, x.a.impedimentos].join(' ')).includes(t))));
  const cnt = s => m.rows.filter(x => x.status === s).length;
  const cols = atVisibleCols(); const nHid = AT_COLS.length - cols.length;
  TK._view = {m, rows, o, w};
  body.innerHTML = `
    <div class="alert info" style="font-size:12.5px;padding:8px 12px;display:flex;align-items:center;gap:10px">Indicadores, histograma e pontos de atenção da obra estão na <b>Visão gerencial</b>.<span style="flex:1"></span><button class="btn sm" data-vgo>Abrir visão gerencial →</button></div>
    <div class="card"><div class="card-h"><h3>${F.status === 'rest' ? 'Atividades restantes' : 'Atividades'}</h3><span class="hint">${rows.length} atividade(s) · situação no fechamento de ${esc(weekLabel(w, true))}</span><span class="sp"></span><button class="btn sm" data-ta="cols">☰ Colunas${nHid ? ` <span class="chip" style="padding:0 6px">${nHid} oculta(s)</span>` : ''}</button></div>
    <div class="tbl-wrap"><table><thead><tr><th>Ações</th>${cols.map(c => `<th>${esc(c.l)}</th>`).join('')}</tr></thead><tbody>
    ${rows.map(x => `<tr><td style="white-space:nowrap"><button class="btn sm" data-ta="edit" data-kind="ativ" data-id="${x.a.id}">✎</button> <button class="btn sm danger" data-ta="del" data-kind="ativ" data-id="${x.a.id}">✕</button></td>${cols.map(c => c.td(x)).join('')}</tr>`).join('') || `<tr><td colspan="${cols.length + 1}" class="muted" style="padding:18px">Nenhuma atividade neste filtro.</td></tr>`}
    </tbody></table></div>
    <p class="muted" style="font-size:12px;margin:8px 0 0">Início e Término mostram a data <b>revisada</b> quando preenchida (marcada “revisado”, com os dias de deslocamento e a data prevista embaixo); sem revisão, mostram a prevista. Qtd. Restante = prevista − realizada. Planejado = parte do prazo previsto (linha de base) já decorrida. Início Revisado já ocorrido = serviço iniciado, mesmo sem medição. Início atrasado = passou o início sem avanço. Término atrasado = passou o término (revisado, ou o previsto) sem concluir. Início Revisado diferente do previsto exige o impedimento/motivo.</p></div>`;
}
function revCell(d, dias, lst, semanas){
  if (!d) return '<td>—</td>'; const n = lst.length;
  return `<td title="${esc(semanas.length ? 'Histórico:\n' + semanas.join('\n') : '')}"><b>${fmtBR(d)}</b>${dias ? `<br><small style="color:#DC2626">+${dias} dias</small>` : ''}${n > 1 ? `<br><small class="chip err" style="padding:1px 6px">${n}ª revisão</small>` : ''}</td>`;
}
function lastAvWeekA(a){ let w = ''; for (const r of Repo.data.avanco) if (obraKey(r.obra) === obraKey(a.obra) && norm(r.codigo) === norm(a.codigo) && r.semana > w) w = r.semana; return w; }
function openAvancoForm(){
  const o = TK.obra; if (!o) return toast('Selecione uma obra.', 'err');
  const w = TK.week; const m = obraModel(o, w); const rows = m.rows.filter(x => x.status !== 'Concluída' || (x.lastAv && x.lastAv.semana === w));
  if (!rows.length) return toast('Nenhuma atividade restante nesta obra.', 'err');
  const body = `<div class="alert info" style="font-size:12px">Informe o <b>acumulado</b> executado até o fim da semana <b>${esc(weekLabel(w, true))}</b>. Atividades com quantidade prevista recebem quantidade; as demais recebem %. Campos não alterados não são gravados.</div>
    <div class="tbl-wrap" style="max-height:60vh"><table><thead><tr><th>Trecho</th><th>Serviço</th><th>Prevista</th><th>Último lançamento</th><th>Qtd. Realizada acumulada</th></tr></thead><tbody>
    ${rows.map(x => { const q = Number(x.a.qtd) || 0; const cur = x.lastAv && x.lastAv.semana === w ? x.lastAv : null; const val = cur ? (q ? (cur.acumulado ?? '') : (cur.pct ?? '')) : '';
      return `<tr><td class="l">${esc(x.a.trecho || '')}</td><td class="l" style="white-space:normal">${esc(x.a.nome)}</td><td>${q ? numTxt(q) + ' ' + esc(x.a.unidade || '') : '%'}</td><td>${x.lastAv ? `${q ? numTxt(x.acum) + ' ' + esc(x.a.unidade || '') : pctTxt(x.real, 0)} <small class="muted">(${esc(x.lastAv.semana)})</small>` : '—'}</td>
      <td><input type="number" min="0" step="any" data-av="${x.a.id}" data-q="${q}" value="${esc(val)}" placeholder="${q ? 'qtd' : '%'}" style="width:110px;border:1px solid var(--line);border-radius:8px;padding:6px"></td></tr>`; }).join('')}</tbody></table></div><div data-msg></div>`;
  const md = openModal({title: `Lançar avanço — ${obraLabel(o)} · ${w}`, body, wide: true, foot: '<button class="btn" data-x>Cancelar</button><button class="btn primary" data-ok>Salvar avanço</button>'});
  $('[data-x]', md.el).onclick = md.close;
  $('[data-ok]', md.el).onclick = () => {
    const errs = []; const ups = [];
    $$('[data-av]', md.el).forEach(i => { if (i.value === '') return; const v = Number(i.value); const a = Repo.data.ativ.find(x => x.id === i.dataset.av); const q = Number(i.dataset.q);
      if (!(v >= 0)) return errs.push(`${a.nome}: valor inválido`); if (!q && v > 100) return errs.push(`${a.nome}: % acima de 100`); if (q && v > q * 1.5) errs.push(`${a.nome}: acumulado muito acima do previsto (${numTxt(q)})`);
      ups.push({a, v, q}); });
    if (errs.length) { $('[data-msg]', md.el).innerHTML = `<div class="alert err" style="margin-top:10px">${errs.map(esc).join('<br>')}</div>`; return; }
    const now = new Date().toISOString(); let n = 0;
    const res = Repo.tx(() => { for (const {a, v, q} of ups) {
      const ex = Repo.data.avanco.find(r => obraKey(r.obra) === obraKey(a.obra) && norm(r.codigo) === norm(a.codigo) && r.semana === w);
      const patch = q ? {acumulado: v, pct: null} : {pct: v, acumulado: null};
      if (w >= lastAvWeekA(a)) { a.realizada = v; a.updatedAt = now; }
      if (ex) { if ((q ? ex.acumulado : ex.pct) == v) continue; Object.assign(ex, patch, {updatedAt: now}); } else Repo.data.avanco.push({id: uid(), obra: a.obra, codigo: a.codigo, semana: w, ...patch, createdAt: now}); n++; } });
    if (res.ok) { toast(`Avanço salvo: ${n} atividade(s) em ${w}.`, 'ok'); md.close(); refreshAll(); }
  };
}

/* ---------- PLANEJAMENTO SEMANAL (Last Planner) ---------- */
function renderLps(){
  const sec = $('#tab-lps');
  if (!sec.dataset.built) { sec.dataset.built = 1; sec.innerHTML = trkToolbar('Planejamento semanal', 'Last Planner: compromissos da semana, PPC e causas de não cumprimento', '<button class="btn primary" data-ta="new" data-kind="lps">＋ Novo compromisso</button><button class="btn" data-ta="copyprev">↻ Trazer não concluídos da semana anterior</button>' + BTN_IMP) + '<div data-body></div>'; }
  trkPickers(sec);
  const o = TK.obra, w = TK.week, body = $('[data-body]', sec); ['lp-p', 'lp-c'].forEach(killChart);
  if (!o) { body.innerHTML = trkEmpty('Nenhuma obra cadastrada'); return; }
  const pc = ppcWeek(o, w), ser = ppcSeries(o, w, 12), cp = causasPareto(o, w);
  const ok4 = ser.slice(-4).filter(x => x.ppc != null); const med = ok4.length ? ok4.reduce((s, x) => s + x.ppc, 0) / ok4.length : null;
  body.innerHTML = `
    <div class="kpis">
      ${kpiBox('PPC da semana', pc.ppc == null ? '—' : pctTxt(pc.ppc, 0), pc.ppc == null ? 'avalie os compromissos abaixo' : pc.ppc >= 80 ? 'meta atingida (≥ 80%)' : 'abaixo da meta de 80%', pc.ppc == null ? '#9CA3AF' : pc.ppc >= 80 ? '#16A34A' : '#DC2626', true)}
      ${kpiBox('Compromissos', fmtN(pc.total), esc(weekLabel(w)), '#334155')}
      ${kpiBox('Concluídos', fmtN(pc.concl), '', '#16A34A')}
      ${kpiBox('Não concluídos', fmtN(pc.nao), '', '#DC2626')}
      ${kpiBox('Sem avaliação', fmtN(pc.pend), 'marque Concluído ou Não concluído', '#9CA3AF')}
      ${kpiBox('PPC médio (4 semanas)', pctTxt(med, 0), '', '#F7931E')}
    </div>
    <div class="card"><div class="card-h"><h3>Compromissos da semana</h3><span class="hint">${esc(weekLabel(w, true))} — o que a equipe se comprometeu a executar</span></div>
      <div class="tbl-wrap"><table><thead><tr><th>Ações</th><th>Código</th><th>Compromisso</th><th>Responsável</th><th>Meta</th><th>Status</th><th>Causa do não cumprimento</th><th>Observação</th></tr></thead><tbody>
      ${pc.list.sort((a, b) => String(a.codigo || '').localeCompare(String(b.codigo || ''), 'pt-BR', {numeric: true})).map(r => `<tr><td><button class="btn sm" data-ta="edit" data-kind="lps" data-id="${r.id}">✎</button> <button class="btn sm danger" data-ta="del" data-kind="lps" data-id="${r.id}">✕</button></td><td>${esc(r.codigo || '')}</td><td class="l" style="white-space:normal;min-width:220px"><b>${esc(r.atividade)}</b></td><td>${esc(r.responsavel || '')}</td><td>${esc(r.meta || '')}</td>
        <td><select data-lps="status" data-id="${r.id}" class="sel-${r.status === 'Concluído' ? 'ok' : r.status === 'Não concluído' ? 'err' : 'pend'}">${STATUS_LPS.map(s => `<option ${(r.status || 'Planejado') === s ? 'selected' : ''}>${s}</option>`).join('')}</select></td>
        <td><select data-lps="causa" data-id="${r.id}" ${r.status === 'Não concluído' ? '' : 'disabled'}><option value="">${r.status === 'Não concluído' ? 'Informe a causa…' : '—'}</option>${CAUSAS.map(c => `<option ${r.causa === c ? 'selected' : ''}>${esc(c)}</option>`).join('')}</select></td><td class="l" style="white-space:normal">${esc(r.obs || '')}</td></tr>`).join('') || `<tr><td colspan="8" class="muted" style="padding:18px">Nenhum compromisso nesta semana. Use “Novo compromisso” ou importe a aba Compromissos.</td></tr>`}
      </tbody></table></div></div>
    <div class="pn-grid">
      <div class="card"><div class="card-h"><h3>PPC — últimas 12 semanas</h3><span class="hint">PPC = concluídos ÷ avaliados · meta ≥ 80%</span></div><div class="chart-frame"><div style="position:relative;height:260px"><canvas id="lpP"></canvas></div></div></div>
      <div class="card"><div class="card-h"><h3>Causas de não cumprimento</h3><span class="hint">últimas 12 semanas</span></div><div style="position:relative;height:${Math.max(140, cp.length * 34 + 40)}px"><canvas id="lpC"></canvas></div></div>
    </div>`;
  chartPPC('lp-p', $('#lpP'), ser);
  if (cp.length) chartHBar('lp-c', $('#lpC'), cp.map(x => x[0]), cp.map(x => x[1]), '#F7931E'); else $('#lpC').parentElement.innerHTML = '<div class="muted" style="font-size:13px">Nenhum compromisso não concluído no período.</div>';
}
function copyPrevLps(){
  const o = TK.obra, w = TK.week; if (!o) return; const pw = weekPlus(w, -1);
  const prev = Repo.data.lps.filter(r => obraKey(r.obra) === o && r.semana === pw && r.status !== 'Concluído');
  const ex = new Set(Repo.data.lps.filter(r => obraKey(r.obra) === o && r.semana === w).map(r => TRACK.lps.key(r)));
  const add = prev.map(r => ({...r, id: uid(), semana: w, status: 'Planejado', causa: '', obs: r.status === 'Não concluído' ? `Reprogramado de ${pw}` : r.obs || '', createdAt: new Date().toISOString()})).filter(r => !ex.has(TRACK.lps.key(r)));
  if (!add.length) return toast(`Nada a trazer de ${pw}.`);
  if (Repo.tx(() => Repo.data.lps.push(...add)).ok) { toast(`${add.length} compromisso(s) reprogramado(s) de ${pw}.`, 'ok'); refreshAll(); }
}

/* ---------- RESTRIÇÕES ---------- */
function renderRestr(){
  const sec = $('#tab-restr');
  if (!sec.dataset.built) {
    sec.dataset.built = 1;
    sec.innerHTML = trkToolbar('Restrições', 'Lookahead: o que precisa ser resolvido antes de a atividade começar (material, projeto, liberação…)', '<button class="btn primary" data-ta="new" data-kind="restr">＋ Nova restrição</button>' + BTN_IMP, false) + '<div data-body></div>';
    $('[data-extra-filters]', sec).innerHTML = `<label>Mostrar<select data-rf="status"><option value="abertas">Abertas</option><option value="venc">Vencidas</option><option value="rem">Removidas</option><option value="">Todas</option></select></label><label>Tipo<select data-rf="tipo"><option value="">Todos</option>${TIPOS_RESTR.map(t => `<option>${esc(t)}</option>`).join('')}</select></label>`;
  }
  trkPickers(sec, false);
  $('[data-rf=status]', sec).value = TK.restr.status; $('[data-rf=tipo]', sec).value = TK.restr.tipo;
  const o = TK.obra, body = $('[data-body]', sec); killChart('rs-t');
  if (!o) { body.innerHTML = trkEmpty('Nenhuma obra cadastrada'); return; }
  const st = restrStats(o); const t = todayYMD(); const tw = isoWeek(t);
  const F = TK.restr;
  const list = st.list.filter(r => (!F.tipo || r.tipo === F.tipo) && (F.status === 'abertas' ? restrAberta(r) : F.status === 'venc' ? restrAberta(r) && r.prazo && r.prazo < t : F.status === 'rem' ? !restrAberta(r) : true))
    .sort((a, b) => (a.prazo || '9999').localeCompare(b.prazo || '9999'));
  // quadro lookahead: vencidas + próximas 6 semanas + sem data
  const cols = [{k: 'venc', l: 'Vencidas', f: r => r.prazo && r.prazo < t}];
  for (let i = 0; i < 6; i++) { const w = weekPlus(tw, i); const rg = weekRange(w); cols.push({k: w, l: `${weekShort(w)}<small>${fmtBR(rg.start).slice(0, 5)}–${fmtBR(rg.end).slice(0, 5)}</small>`, f: r => r.prazo && r.prazo >= t && isoWeek(r.prazo) === w}); }
  cols.push({k: 'depois', l: 'Depois', f: r => r.prazo && r.prazo >= t && r.prazo > weekRange(weekPlus(tw, 5)).end}, {k: 'sem', l: 'Sem data', f: r => !r.prazo});
  const ab = st.ab.filter(r => !F.tipo || r.tipo === F.tipo);
  const porTipo = TIPOS_RESTR.map(tp => [tp, st.ab.filter(r => (r.tipo || 'Outros') === tp).length]).filter(x => x[1]).sort((a, b) => b[1] - a[1]);
  body.innerHTML = `
    <div class="kpis">
      ${kpiBox('Abertas', fmtN(st.ab.length), 'aguardando remoção', '#F7931E', true)}
      ${kpiBox('Vencidas', fmtN(st.venc.length), 'data necessária já passou', '#DC2626')}
      ${kpiBox('Vencem em 7 dias', fmtN(st.prox.length), '', '#F7931E')}
      ${kpiBox('Removidas', fmtN(st.rem.length), `de ${st.total} cadastradas`, '#16A34A')}
      ${kpiBox('Removidas no prazo', pctTxt(st.irr, 0), 'remoção até a data necessária', '#334155')}
    </div>
    <div class="card"><div class="card-h"><h3>Quadro de restrições abertas (lookahead 6 semanas)</h3><span class="hint">por semana da data necessária</span></div>
      <div class="la-board">${cols.map(c => { const it = ab.filter(c.f); return `<div class="la-col ${c.k === 'venc' ? 'venc' : ''}"><div class="la-h">${c.l}<span class="chip">${it.length}</span></div>${it.map(r => `<div class="la-card"><b>${esc(r.descricao)}</b>${r.atividade || r.codigo ? `<small>${esc([r.codigo, r.atividade].filter(Boolean).join(' · '))}</small>` : ''}<div class="la-meta"><span class="chip">${esc(r.tipo || 'Outros')}</span>${r.responsavel ? `<span>👤 ${esc(r.responsavel)}</span>` : ''}${r.prazo ? `<span>📅 ${fmtBR(r.prazo)}</span>` : ''}</div><div class="la-acts"><button class="btn sm" data-ta="rmrestr" data-id="${r.id}">✓ Removida</button><button class="btn sm" data-ta="edit" data-kind="restr" data-id="${r.id}">✎</button></div></div>`).join('') || '<div class="muted" style="font-size:12px;padding:6px">—</div>'}</div>`; }).join('')}</div></div>
    <div class="pn-grid">
      <div class="card"><div class="card-h"><h3>Lista de restrições</h3><span class="hint">${list.length} no filtro</span></div><div class="tbl-wrap"><table><thead><tr><th>Ações</th><th>Restrição</th><th>Atividade</th><th>Tipo</th><th>Responsável</th><th>Data necessária</th><th>Status</th><th>Removida em</th></tr></thead><tbody>
        ${list.map(r => `<tr><td><button class="btn sm" data-ta="edit" data-kind="restr" data-id="${r.id}">✎</button> <button class="btn sm danger" data-ta="del" data-kind="restr" data-id="${r.id}">✕</button></td><td class="l" style="white-space:normal;min-width:200px"><b>${esc(r.descricao)}</b>${r.obs ? `<br><small class="muted">${esc(r.obs)}</small>` : ''}</td><td class="l">${esc([r.codigo, r.atividade].filter(Boolean).join(' · '))}</td><td>${esc(r.tipo || '')}</td><td>${esc(r.responsavel || '')}</td><td>${r.prazo ? `<span class="chip ${restrAberta(r) && r.prazo < t ? 'err' : ''}">${fmtBR(r.prazo)}</span>` : '—'}</td><td>${restrAberta(r) ? '<span class="chip warn">Aberta</span>' : '<span class="chip ok">Removida</span>'}</td><td>${fmtBR(r.remocao || '')}</td></tr>`).join('') || '<tr><td colspan="8" class="muted" style="padding:18px">Nenhuma restrição neste filtro.</td></tr>'}
      </tbody></table></div></div>
      <div class="card"><div class="card-h"><h3>Abertas por tipo</h3></div><div style="position:relative;height:${Math.max(140, porTipo.length * 34 + 40)}px"><canvas id="rsT"></canvas></div></div>
    </div>`;
  if (porTipo.length) chartHBar('rs-t', $('#rsT'), porTipo.map(x => x[0]), porTipo.map(x => x[1]), '#334155'); else $('#rsT').parentElement.innerHTML = '<div class="muted" style="font-size:13px">Nenhuma restrição aberta.</div>';
}
function removeRestr(id){
  const r = Repo.data.restr.find(x => x.id === id); if (!r) return;
  if (Repo.tx(() => { const x = Repo.data.restr.find(y => y.id === id); x.status = 'Removida'; x.remocao = todayYMD(); x.updatedAt = new Date().toISOString(); }).ok) { toast('Restrição marcada como removida.', 'ok'); refreshAll(); }
}

/* ---------- formulário genérico (atividade, compromisso, restrição) ---------- */
function openTrkForm(kind, existing){
  const S = TRACK[kind]; const r = existing || {obra: obraLabel(TK.obra), semana: TK.week, status: kind === 'lps' ? 'Planejado' : kind === 'restr' ? 'Aberta' : ''};
  const ativOpts = Repo.data.ativ.filter(a => obraKey(a.obra) === TK.obra);
  const body = `<div class="form-grid">${S.fields.filter(f => !f.hidden).map(f => {
    const v = r[f.k] ?? ''; let inp;
    if (f.t === 'date') inp = `<input type="date" data-k="${f.k}" value="${esc(v)}">`;
    else if (f.t === 'week') inp = weekInputHTML(f.k, v);
    else if (f.t === 'num' || f.t === 'pct') inp = `<input type="number" step="any" min="0" data-k="${f.k}" value="${esc(v)}">`;
    else if (f.t === 'opt') inp = `<select data-k="${f.k}"><option value=""></option>${f.opts.map(o => `<option ${v === o ? 'selected' : ''}>${esc(o)}</option>`).join('')}</select>`;
    else if (f.k === 'obra') inp = `<input data-k="obra" list="dl-trk-obras" value="${esc(v)}">`;
    else if (f.k === 'codigo' && kind !== 'ativ') inp = `<input data-k="codigo" list="dl-trk-ativ" value="${esc(v)}" placeholder="opcional">`;
    else inp = `<input data-k="${f.k}" value="${esc(v)}">`;
    return `<div class="fld" data-fd="${f.k}" ${['nome', 'atividade', 'descricao', 'obs', 'impedimentos'].includes(f.k) ? 'style="grid-column:1/-1"' : ''}><span>${esc(f.l)} ${f.req ? '<span class="req">*</span>' : ''}</span>${inp}<span class="err-msg"></span></div>`; }).join('')}</div>
    <datalist id="dl-trk-obras">${trkObras().map(o => `<option value="${esc(o.l)}">`).join('')}</datalist>
    <datalist id="dl-trk-ativ">${ativOpts.map(a => `<option value="${esc(a.codigo)}">${esc(a.nome)}</option>`).join('')}</datalist>
    ${kind === 'ativ' ? '<p class="muted" style="font-size:12px;margin:10px 0 0">Qtd. Realizada = total <b>acumulado</b> executado até hoje; o app guarda o valor de cada semana para calcular o ritmo e a previsão de término. Qtd. Restante é calculada.</p>' : ''}<div data-msg></div>`;
  const md = openModal({title: (existing ? 'Editar ' : 'Nova(o) ') + ({ativ: 'atividade', lps: 'compromisso', restr: 'restrição'}[kind] || S.title), body, foot: '<button class="btn" data-x>Cancelar</button><button class="btn primary" data-ok>Salvar</button>'});
  bindWeekPickers(md.el);
  if (kind !== 'ativ') { const c = $('[data-k=codigo]', md.el), at = $('[data-k=atividade]', md.el); if (c && at) c.addEventListener('change', () => { const a = ativOpts.find(x => norm(x.codigo) === norm(c.value)); if (a && !at.value) at.value = a.nome; }); }
  if (kind === 'lps') { const s = $('[data-k=status]', md.el), cz = $('[data-k=causa]', md.el); const sync = () => { cz.disabled = s.value !== 'Não concluído'; if (cz.disabled) cz.value = ''; }; s.addEventListener('change', sync); sync(); }
  $('[data-x]', md.el).onclick = md.close;
  $('[data-ok]', md.el).onclick = () => {
    const raw = {}; $$('[data-k]', md.el).forEach(i => raw[i.dataset.k] = i.value.trim());
    if (kind === 'ativ' && existing) raw.codigo = existing.codigo || '';
    const {rec, errors} = parseTrkRow(kind, raw, 'dmy');
    if (kind === 'ativ' && iniSemMotivo(rec)) errors.push({field: 'Impedimentos', msg: 'Início Revisado diferente do previsto — informe o impedimento/motivo'});
    $$('.fld', md.el).forEach(x => { x.classList.remove('bad'); $('.err-msg', x).textContent = ''; });
    errors.forEach(e => { const f = S.fields.find(x => x.l === e.field); const b = f && $(`[data-fd="${f.k}"]`, md.el); if (b) { b.classList.add('bad'); $('.err-msg', b).textContent = e.msg; } });
    if (errors.length) return;
    const key = S.key(rec); const clash = Repo.data[kind].find(x => x.id !== (existing && existing.id) && S.key(x) === key);
    if (clash) { $('[data-msg]', md.el).innerHTML = '<div class="alert err" style="margin-top:10px">Já existe um registro igual (mesma obra e identificação). Edite o existente.</div>'; return; }
    const now = new Date().toISOString();
    const res = Repo.tx(() => { let t; if (existing) { t = Repo.data[kind].find(x => x.id === existing.id); const chR = String(existing.realizada ?? '') !== String(rec.realizada ?? ''); const chD = existing.inicio_rev !== rec.inicio_rev || existing.revisado !== rec.revisado; Object.assign(t, rec, {updatedAt: now}); if (kind === 'ativ' && chR) syncRealizada(t, rec.realizada); if (kind === 'ativ' && chD) syncDatas(t); } else { t = {id: uid(), ...rec, createdAt: now}; Repo.data[kind].push(t); if (kind === 'ativ') { syncRealizada(t, rec.realizada); syncDatas(t); } } });
    if (res.ok) { toast('Salvo.', 'ok'); md.close(); TK.obra = obraKey(rec.obra); refreshAll(); }
  };
}
async function deleteTrk(kind, id){
  const r = Repo.data[kind].find(x => x.id === id); if (!r) return;
  const extra = kind === 'ativ' ? Repo.data.avanco.filter(a => obraKey(a.obra) === obraKey(r.obra) && norm(a.codigo) === norm(r.codigo)).length : 0;
  if (!(await confirmBox('Excluir', `<p>Excluir <b>${esc(r.nome || r.atividade || r.descricao)}</b>?</p>${extra ? `<p class="muted">Os ${extra} lançamento(s) de avanço desta atividade também serão excluídos.</p>` : ''}`, 'Excluir', true))) return;
  if (Repo.tx(() => { Repo.data[kind] = Repo.data[kind].filter(x => x.id !== id); if (extra) Repo.data.avanco = Repo.data.avanco.filter(a => !(obraKey(a.obra) === obraKey(r.obra) && norm(a.codigo) === norm(r.codigo))); }).ok) { toast('Excluído.'); refreshAll(); }
}

/* código interno estável da atividade (liga os lançamentos de avanço) */
function autoCode(r){ let h = 0; const t = norm((r.trecho || '') + '|' + (r.nome || '')); for (let i = 0; i < t.length; i++) h = (h * 31 + t.charCodeAt(i)) >>> 0; return 'A' + h.toString(36).toUpperCase(); }
/* grava a Qtd. Realizada (acumulada) como avanço da semana — mantém o histórico semanal */
function syncDatas(a, w){
  if (!a.inicio_rev && !a.revisado && !Repo.data.avanco.some(r => obraKey(r.obra) === obraKey(a.obra) && norm(r.codigo) === norm(a.codigo) && (r.inicio_rev || r.revisado))) return;
  w = w || todayWeek(); const now = new Date().toISOString(); const patch = {inicio_rev: a.inicio_rev || null, revisado: a.revisado || null};
  const ex = Repo.data.avanco.find(r => obraKey(r.obra) === obraKey(a.obra) && norm(r.codigo) === norm(a.codigo) && r.semana === w);
  if (ex) Object.assign(ex, patch, {updatedAt: now}); else Repo.data.avanco.push({id: uid(), obra: a.obra, codigo: a.codigo, semana: w, ...patch, createdAt: now});
}
/* histórico de reprogramações: datas revisadas distintas registradas semana a semana */
function revHist(a, hist){
  const l = hist || []; const ini = [], fim = [];
  for (const r of [...l, a]) { if (r.inicio_rev && ini[ini.length - 1] !== r.inicio_rev) ini.push(r.inicio_rev); if (r.revisado && fim[fim.length - 1] !== r.revisado) fim.push(r.revisado); }
  return {ini: [...new Set(ini)], fim: [...new Set(fim)], semanas: l.filter(r => r.inicio_rev || r.revisado).map(r => `${r.semana}: início ${fmtBR(r.inicio_rev) || '—'} · término ${fmtBR(r.revisado) || '—'}`)};
}
function syncRealizada(a, v, w){
  if (v == null || v === '' || !isFinite(Number(v))) return; w = w || todayWeek(); const q = Number(a.qtd) || 0; const now = new Date().toISOString();
  const patch = q ? {acumulado: Number(v), pct: null} : {pct: Math.min(100, Number(v)), acumulado: null};
  const ex = Repo.data.avanco.find(r => obraKey(r.obra) === obraKey(a.obra) && norm(r.codigo) === norm(a.codigo) && r.semana === w);
  if (ex) Object.assign(ex, patch, {updatedAt: now}); else Repo.data.avanco.push({id: uid(), obra: a.obra, codigo: a.codigo, semana: w, ...patch, createdAt: now});
}
/* ---------- importação das abas de acompanhamento ---------- */
function parsePct(raw){
  if (raw == null || raw === '') return {v: null};
  if (typeof raw === 'number') return {v: raw <= 1 && raw >= 0 ? raw * 100 : raw};
  const t = String(raw).trim(); if (!t || DASHES.has(t)) return {v: null};
  const n = Number(t.replace('%', '').replace(/\s/g, '').replace(',', '.')); if (!isFinite(n) || n < 0) return {v: null, err: `Percentual inválido "${t}"`};
  return {v: /%/.test(t) ? n : (n <= 1 && /[.,]/.test(t) ? n * 100 : n)};
}
function parseNum(raw){
  if (raw == null || raw === '') return {v: null};
  if (typeof raw === 'number') return {v: raw};
  const t = String(raw).trim(); if (!t || DASHES.has(t)) return {v: null};
  const n = Number(/,\d{1,3}$/.test(t) ? t.replace(/\./g, '').replace(',', '.') : t.replace(/,/g, '')); return isFinite(n) && n >= 0 ? {v: n} : {v: null, err: `Número inválido "${t}"`};
}
const SYN = [[/CHUVA|TEMPO|CLIMA|INTEMPER/, 'Clima'], [/MATERIA|SUPRIM|ENTREGA|FORNEC|INSUMO|ACO|CONCRETO/, 'Material / suprimentos'], [/MAO DE OBRA|EQUIPE|FUNCIONAR|PESSOAL|EFETIVO|FALTA DE GENTE|ABSENT/, 'Mão de obra (falta ou deslocamento)'],
  [/EQUIPAMENT|MAQUIN|QUEBRA|MANUTENC/, 'Equipamento'], [/PROJETO|ESCOPO|ALTERAC|MUDANC|ENGENHARIA/, 'Projeto / mudança de escopo'], [/PREDECESS|ANTERIOR|FRENTE|INTERFERENC/, 'Atividade predecessora não concluída'],
  [/PRODUTIV|PLANEJ|SUPERESTIM|PROGRAMAC/, 'Planejamento / produtividade superestimada'], [/SEGURAN|LIBERAC|PERMISS|\bPT\b|ACIDENT/, 'Segurança / liberação'], [/CLIENTE|CONTRAT|FISCAL|MEDICAO|PAGAMENT/, 'Cliente / contratual']];
function matchOpt(opts, v){
  if (opts === CAUSAS) { const n0 = norm(v); const ex = opts.find(o => norm(o) === n0); if (ex) return ex; const s = SYN.find(([re]) => re.test(n0)); if (s) return s[1]; } const n = norm(v); if (!n) return ''; const exact = opts.find(o => norm(o) === n); if (exact) return exact; const part = opts.find(o => norm(o).startsWith(n) || n.startsWith(norm(o).split(' ')[0])); return part || null; }
const iniSemMotivo = x => !!(x.inicio_rev && x.inicio_rev !== x.inicio && !String(x.impedimentos || '').trim());
const UNIT_RE = /^(%|und|unid|un|u|m|ml|m2|m²|m3|m³|km|kg|t|ton|vb|verba|cj|conj|pç|pc|peça|h|hh|l|dia|mês|mes)\.?$/i;
function parseTrkRow(kind, raw, order){
  const S = TRACK[kind]; const rec = {}; const errors = [];
  for (const f of S.fields) {
    const v = raw[f.k];
    if (f.t === 'date') { const d = parseDate(v, order); rec[f.k] = d.v; if (d.err) errors.push({field: f.l, msg: d.err}); }
    else if (f.t === 'week') { const w = parseWeek(v); rec[f.k] = w.v; if (w.err) errors.push({field: f.l, msg: w.err}); }
    else if (f.t === 'num') { const n = parseNum(v); rec[f.k] = n.v; if (n.err) errors.push({field: f.l, msg: UNIT_RE.test(String(v).trim()) ? `contém "${String(v).trim()}", que é uma unidade — os dados parecem estar uma coluna à direita do cabeçalho (confira a coluna Und.)` : n.err, shift: UNIT_RE.test(String(v).trim())}); }
    else if (f.t === 'pct') { const n = parsePct(v); rec[f.k] = n.v; if (n.err) errors.push({field: f.l, msg: n.err}); }
    else if (f.t === 'opt') {
      let x = clean(v);
      if (kind === 'lps' && f.k === 'status' && x) { const n = norm(x); x = /^(SIM|S|OK|X|CONCLUIDO|CONCLUIDA|CUMPRIDO|CUMPRIDA|100%?)$/.test(n) ? 'Concluído' : /^(NAO|N|NAO CONCLUIDO|NAO CONCLUIDA|NAO CUMPRIDO|NAO CUMPRIDA|PENDENTE|0%?)$/.test(n) ? 'Não concluído' : /^(PLANEJADO|PROGRAMADO|A FAZER|EM ANDAMENTO)$/.test(n) ? 'Planejado' : x; }
      if (kind === 'restr' && f.k === 'status' && x) { const n = norm(x); x = /REMOVID|RESOLVID|LIBERAD|CONCLUID|FECHAD|OK/.test(n) ? 'Removida' : 'Aberta'; }
      const m = matchOpt(f.opts, x); rec[f.k] = m === null ? (f.k === 'causa' || f.k === 'tipo' ? 'Outros' : x) : m;
    }
    else rec[f.k] = clean(v);
    if (f.req && (rec[f.k] == null || rec[f.k] === '') && !errors.some(e => e.field === f.l)) errors.push({field: f.l, msg: 'Obrigatório'});
  }
  if (kind === 'ativ' && rec.inicio && rec.fim && rec.fim < rec.inicio) errors.push({field: 'Término Previsto', msg: 'Anterior ao início'});
  if (kind === 'ativ' && rec.inicio_rev && rec.revisado && rec.revisado < rec.inicio_rev) errors.push({field: 'Término Revisado', msg: 'Anterior ao início revisado'});
  if (kind === 'ativ' && !rec.codigo && rec.nome) rec.codigo = autoCode(rec);
  if (kind === 'avanco' && rec.acumulado == null && rec.pct == null) errors.push({field: 'Qtd executada acumulada', msg: 'Informe a quantidade acumulada ou o % acumulado'});
  if (kind === 'restr' && rec.status === 'Removida' && !rec.remocao) rec.remocao = '';
  if (kind === 'restr' && !rec.status) rec.status = 'Aberta';
  if (kind === 'lps' && !rec.status) rec.status = 'Planejado';
  if (rec.obra) rec.obra = canonObra(rec.obra);
  return {rec, errors};
}
function trkSheetMatch(kind, name){ const h = normHeader(name); return kind === 'ativ' ? /^atividades?$|^cronograma/.test(h) : kind === 'avanco' ? /^avanco|^medicao/.test(h) : kind === 'lps' ? /compromisso|lastplanner|planejamentosemanal|^lps/.test(h) : /restric/.test(h); }
function trkMapHeaders(kind, grid){
  const S = TRACK[kind]; let best = null;
  for (const row of grid.slice(0, 15)) {
    const map = {}; let hits = 0;
    row.cells.forEach((c, i) => { if (!c) return; const rawH = String(c.w ?? c.v ?? ''); const h = normHeader(rawH); let f = S.fields.find(x => x.al.includes(h) || x.al.includes(h.replace(/\d+$/, ''))); if (/%/.test(rawH) && S.fields.some(x => x.t === 'pct')) f = S.fields.find(x => x.t === 'pct'); if (f && map[f.k] == null) { map[f.k] = i; hits++; } });
    const ok = kind === 'ativ' ? (map.nome != null && hits >= 3) : (map.obra != null && hits >= 2);
    if (ok && (!best || hits > best.hits)) best = {row: row.n, map, hits};
  }
  return best;
}
function analyzeTrk(kind, grid, hdr, defObra){
  const S = TRACK[kind]; const out = []; const seen = new Map(); const ex = new Map(Repo.data[kind].map(r => [S.key(r), r]));
  // ordem das datas (dd/mm ou mm/dd) pelo próprio arquivo
  let dmy = 0, mdy = 0; const dcols = S.fields.filter(f => f.t === 'date').map(f => hdr.map[f.k]).filter(i => i != null);
  for (const row of grid) { if (row.n <= hdr.row) continue; for (const i of dcols) { const c = row.cells[i]; if (!c || c.t === 'n') continue; const m = String(c.w ?? c.v).match(/^(\d{1,2})[\/.\-](\d{1,2})[\/.\-]\d{2,4}/); if (m) { if (+m[1] > 12) dmy++; if (+m[2] > 12) mdy++; } } }
  const order = mdy && !dmy ? 'mdy' : 'dmy';
  for (const row of grid) {
    if (row.n <= hdr.row) continue; const raw = {};
    for (const f of S.fields) { const i = hdr.map[f.k]; const c = i == null ? null : row.cells[i]; raw[f.k] = !c ? '' : (c.t === 'n' && (f.t === 'date' || f.t === 'num' || f.t === 'pct' || f.t === 'week')) ? c.v : (c.w ?? String(c.v ?? '')); }
    if (S.fields.every(f => isEmptyVal(raw[f.k]))) continue;
    if (hdr.map.obra == null) { if (isEmptyVal(raw.nome)) continue; raw.obra = defObra || ''; }
    const {rec, errors} = parseTrkRow(kind, raw, order);
    const res = {line: row.n, rec, errors, existing: null, changed: false, filled: S.fields.filter(f => (hdr.map[f.k] != null || (f.k === 'obra' && raw.obra)) && !isEmptyVal(raw[f.k])).map(f => f.k)};
    if (!errors.length) { const k = S.key(rec); if (seen.has(k)) errors.push({field: 'Linha', msg: `Duplicada (mesma identificação da linha ${seen.get(k)})`}); else { seen.set(k, row.n); res.existing = ex.get(k) || null; if (res.existing) res.changed = res.filled.some(f => String(res.existing[f] ?? '') !== String(rec[f] ?? '')); }
      if (kind === 'ativ') { const x = {...(res.existing || {}), ...Object.fromEntries(res.filled.map(f => [f, rec[f]]))}; if (iniSemMotivo(x)) errors.push({field: 'Impedimentos', msg: `Início Revisado (${fmtBR(x.inicio_rev)}) diferente do Início Previsto${x.inicio ? ` (${fmtBR(x.inicio)})` : ''} — informe o impedimento/motivo da reprogramação`}); } }
    out.push(res);
  }
  return out;
}
function commitTrk(kind, results){
  const S = TRACK[kind]; const idx = new Map(Repo.data[kind].map(r => [S.key(r), r])); const now = new Date().toISOString(); let c = 0, u = 0;
  for (const r of results) { if (r.errors.length) continue; const k = S.key(r.rec); const ex = idx.get(k);
    if (ex) { if (!r.changed) continue; const chR = r.filled.includes('realizada') && String(ex.realizada ?? '') !== String(r.rec.realizada ?? ''); const d0 = (ex.inicio_rev || '') + '|' + (ex.revisado || ''); r.filled.forEach(f => { if (f !== 'codigo' || !ex.codigo) ex[f] = r.rec[f]; }); ex.updatedAt = now; if (kind === 'ativ' && chR) syncRealizada(ex, ex.realizada); if (kind === 'ativ' && d0 !== (ex.inicio_rev || '') + '|' + (ex.revisado || '')) syncDatas(ex); u++; }   // célula vazia mantém o valor existente
    else { const n = {id: uid(), ...r.rec, createdAt: now}; Repo.data[kind].push(n); idx.set(k, n); if (kind === 'ativ') { syncRealizada(n, n.realizada); syncDatas(n); } c++; } }
  return {created: c, updated: u};
}
function openTrkImport(){
  if (!libsOk()) return toast('Bibliotecas não carregadas.', 'err');
  const st = {wb: null, src: '', parts: {}};
  const body = `<div class="alert info" style="font-size:12px">Lê a aba <b>Atividades</b> (Trecho, Serviço, Tipo, Und., Qtd. Prevista, Qtd. Realizada, Início/Término Previsto, Início/Término Revisado, Impedimentos). Se a aba não tiver coluna <b>Obra</b>, escolha a obra abaixo. A Qtd. Realizada é gravada como avanço da semana atual (histórico para ritmo e previsão). Linhas existentes são atualizadas só nos campos preenchidos; nada é apagado.</div>
    <div class="cols2"><div class="drop" data-drop><b>Arraste o arquivo aqui</b><br>ou clique para escolher (.xlsx)<input type="file" accept=".xlsx,.xls,.xlsm" hidden data-file></div>
      <div class="card" style="margin:0;box-shadow:none"><div style="font-weight:700;margin-bottom:6px">Planilha conectada</div><div class="muted" style="font-size:12px;margin-bottom:8px">${esc(sheetCfg().url.slice(0, 80))}…</div><button class="btn dark" data-gs>Ler planilha conectada</button><div data-gserr></div></div></div>
    <div data-sum style="margin-top:12px"></div>`;
  const m = openModal({title: 'Importar atividades', body, wide: true, foot: '<label data-cw hidden style="margin-right:auto;display:flex;gap:6px;align-items:center;font-size:13px"><input type="checkbox" data-confirm> Confirmo a atualização dos registros existentes</label><button class="btn" data-x>Cancelar</button><button class="btn primary" data-ok disabled>Importar</button>'});
  const el = m.el; $('[data-x]', el).onclick = m.close;
  const drop = $('[data-drop]', el), fi = $('[data-file]', el);
  drop.onclick = () => fi.click(); drop.ondragover = e => { e.preventDefault(); drop.classList.add('over'); }; drop.ondragleave = () => drop.classList.remove('over');
  drop.ondrop = e => { e.preventDefault(); drop.classList.remove('over'); if (e.dataTransfer.files[0]) load(e.dataTransfer.files[0]); };
  fi.onchange = () => fi.files[0] && load(fi.files[0]);
  async function load(f){ try { setWb(XLSX.read(await f.arrayBuffer(), {type: 'array', cellDates: false, cellText: true}), f.name); } catch (e) { toast('Não foi possível ler o arquivo: ' + e.message, 'err'); } }
  $('[data-gs]', el).onclick = async () => { const b = $('[data-gs]', el); b.disabled = true; b.textContent = 'Lendo…'; try { const c = sheetCfg(); setWb(await fetchPublishedWorkbook(c.url, {labor: c.gidL, equip: c.gidE}), 'Planilha conectada'); } catch (e) { $('[data-gserr]', el).innerHTML = `<div class="alert err" style="margin-top:8px;font-size:12px">${esc(e.message)}</div>`; } finally { b.disabled = false; b.textContent = 'Ler planilha conectada'; } };
  function obraPick(){ return `<label style="display:flex;gap:8px;align-items:center;margin:6px 0;font-size:13px"><b>Obra destas atividades</b><input data-obra-imp list="dl-imp-obras" value="${esc(st.obra || '')}" placeholder="Nome da obra" style="border:1px solid var(--line);border-radius:8px;padding:6px;min-width:220px"></label><datalist id="dl-imp-obras">${trkObras().map(o => `<option value="${esc(o.l)}">`).join('')}</datalist>`; }
  function setWb(wb, src){
    st.wb = wb; st.src = src; st.parts = {};
    const html = TRK_ON.map(k => {
      const S = TRACK[k];
      let sh = wb.SheetNames.find(n => trkSheetMatch(k, n));
      if (!sh && k === 'ativ') sh = wb.SheetNames.find(n => !/^(mo|maodeobra|equipamentos?|resumo|instrucoes|listas)$/.test(normHeader(n)) && trkMapHeaders(k, sheetGrid(wb.Sheets[n])));
      if (!sh) return `<div class="card" style="box-shadow:none;margin-bottom:10px"><b>${S.title}</b> <span class="muted" style="font-size:13px">— aba “${S.sheet}” não encontrada (ignorada).</span></div>`;
      const grid = sheetGrid(wb.Sheets[sh]); const hdr = trkMapHeaders(k, grid);
      if (!hdr) return `<div class="card" style="box-shadow:none;margin-bottom:10px"><b>${S.title}</b> <div class="alert err" style="margin-top:6px">Cabeçalhos não reconhecidos na aba “${esc(sh)}”. Esperado: ${S.fields.map(f => f.l).join(', ')}.</div></div>`;
      const needObra = hdr.map.obra == null; if (needObra && !st.obra) st.obra = obraLabel(TK.obra) || '';
      if (needObra && !st.obra) return `<div class="card" style="box-shadow:none;margin-bottom:10px"><b>${S.title}</b> <span class="hint">aba “${esc(sh)}”</span>${obraPick()}<div class="alert warn" style="margin-top:8px">Informe a obra destas atividades.</div></div>`;
      const res = analyzeTrk(k, grid, hdr, st.obra); st.parts[k] = res;
      const v = res.filter(r => !r.errors.length), bad = res.filter(r => r.errors.length), news = v.filter(r => !r.existing), upd = v.filter(r => r.changed);
      const errs = bad.flatMap(r => r.errors.map(e => `<li><b>Linha ${r.line}</b> — ${esc(e.field)}: ${esc(e.msg)}</li>`));
      const nSM = res.filter(r => r.errors.some(e => e.field === 'Impedimentos')).length;
      const warnSM = nSM ? `<div class="alert err" style="font-size:12px;margin-top:6px"><b>${nSM} serviço(s) com Início Revisado diferente do previsto e sem impedimento/motivo</b> — essas linhas não serão importadas. Preencha a coluna Impedimentos (ou corrija o Início Revisado) e importe de novo.</div>` : '';
      const miss = S.fields.filter(f => !f.hidden && hdr.map[f.k] == null && !(f.k === 'obra' && needObra)).map(f => f.l);
      const nShift = res.filter(r => r.errors.some(e => e.shift)).length;
      const shiftMsg = nShift ? `<div class="alert err" style="font-size:12px;margin-top:6px"><b>Colunas deslocadas em ${nShift} linha(s):</b> a coluna <b>Qtd. Prevista</b> está recebendo a unidade (%, und…) e a coluna <b>Und.</b> está vazia. Os dados estão uma coluna à direita dos cabeçalhos — provavelmente sobrou uma coluna vazia ao colar (ex.: a antiga posição de Impedimentos, que agora é a última coluna). Ajuste na planilha e importe de novo. Datas e quantidades vazias são aceitas normalmente.</div>` : '';
      return `<div class="card" style="box-shadow:none;margin-bottom:10px"><div class="card-h" style="margin-bottom:4px"><h3>${S.title}</h3><span class="hint">aba “${esc(sh)}”</span></div>${needObra ? obraPick() : ''}
        <div class="stat-row"><span class="chip">${res.length} linha(s)</span><span class="chip ok">${v.length} válida(s)</span><span class="chip err">${bad.length} pendente(s)</span><span class="chip or">${news.length} nova(s)</span><span class="chip warn">${upd.length} alterada(s)</span><span class="chip">${v.length - news.length - upd.length} sem alteração</span></div>
        ${miss.length ? `<div class="alert info" style="font-size:12px">Colunas não encontradas: ${esc(miss.join(', '))}</div>` : ''}${shiftMsg}${warnSM}${errs.length ? `<ul class="err-list">${errs.slice(0, 200).join('')}</ul>` : ''}</div>`;
    }).join('');
    $('[data-sum]', el).innerHTML = `<div class="muted" style="font-size:12px;margin-bottom:6px">Fonte: ${esc(src)}</div>` + html;
    const op = $('[data-obra-imp]', el); if (op) op.onchange = () => { st.obra = canonObra(op.value.trim()); setWb(st.wb, st.src); };
    upd();
  }
  $('[data-confirm]', el).onchange = upd;
  function upd(){
    const act = TRK_ON.reduce((s, k) => s + (st.parts[k] || []).filter(r => !r.errors.length && (!r.existing || r.changed)).length, 0);
    const ch = TRK_ON.some(k => (st.parts[k] || []).some(r => !r.errors.length && r.changed));
    $('[data-cw]', el).hidden = !ch; const b = $('[data-ok]', el);
    b.disabled = !act || (ch && !$('[data-confirm]', el).checked); b.textContent = act ? `Importar ${act} registro(s)` : (st.wb ? 'Nada a alterar' : 'Importar');
  }
  $('[data-ok]', el).onclick = () => {
    Snap.take('Antes de importar acompanhamento (' + st.src + ')');
    const res = Repo.tx(() => Object.fromEntries(TRK_ON.filter(k => st.parts[k]).map(k => [k, commitTrk(k, st.parts[k])])));
    if (!res.ok) return;
    toast('Acompanhamento importado — ' + Object.entries(res.out).map(([k, r]) => `${TRACK[k].title}: ${r.created} novo(s), ${r.updated} alterado(s)`).join(' · ') + '.', 'ok');
    const first = Repo.data.ativ[0]; if (!TK.obra && first) TK.obra = obraKey(first.obra);
    m.close(); refreshAll();
  };
}
function trkTemplateSheets(wb){
  for (const k of TRK_ON) { const S = TRACK[k]; const h = S.tpl || S.fields.filter(f => !f.hidden).map(f => f.l); const ws = XLSX.utils.aoa_to_sheet([h]); ws['!cols'] = h.map(l => ({wch: l === 'Serviço' ? 50 : l === 'Impedimentos' ? 30 : Math.max(12, l.length + 4)})); XLSX.utils.book_append_sheet(wb, ws, S.sheet); }
}
function trkInstr(){
  return [['ATIVIDADES'],
    ['Aba Atividades', '', 'Uma linha por serviço: Obra, Trecho, Serviço, Tipo, Und., Qtd. Prevista, Qtd. Realizada (ACUMULADA até hoje), Qtd. Restante (calculada pelo app — pode deixar a fórmula), Início Previsto, Término Previsto, Início Revisado (data real de início, ou nova data se ainda não começou), Término Revisado (dd/mm/aaaa) e Impedimentos (última coluna). Use as datas revisadas quando o serviço for reprogramado (ex.: por impedimento).'],
    ['Chave', '', 'Obra + Trecho + Serviço. Sem coluna Obra, o app pergunta a obra na importação. Célula vazia mantém o valor já existente no app.'],
    ['Histórico', '', 'Cada importação grava a Qtd. Realizada na semana atual; com isso o app calcula ritmo, Curva S e previsão de término.'], []];
}
function printAtiv(){
  if (!libsOk()) return toast('Bibliotecas não carregadas.', 'err');
  const v = TK._view; if (!v || !v.o || !v.m.n) return toast('Nenhuma atividade para gerar o PDF.', 'err');
  const {m, rows, o, w} = v; const F = TK.ativ;
  const showTxt = {rest: 'Atividades restantes', atr: 'Atrasadas', imp: 'Com impedimento', rep: 'Reprogramadas', sm: 'Início reprogramado sem motivo', and: 'Em andamento', nao: 'Não iniciadas', conc: 'Concluídas', '': 'Todas'}[F.status] || 'Atividades restantes';
  const flt = [['Obra', obraLabel(o)], ['Semana de referência', weekLabel(w, true)], ['Trecho', F.disc || 'Todos'], ['Mostrar', showTxt]]; if (F.q) flt.push(['Busca', F.q]);
  const cols = atVisibleCols(); const tw = cols.reduce((s, c) => s + c.pw, 0);
  const c2 = t => `<span class="c2">${esc(t)}</span>`;
  const tbl = `<table class="p-at"><colgroup>${cols.map(c => `<col style="width:${(c.pw / tw * 100).toFixed(2)}%">`).join('')}</colgroup><thead><tr>${cols.map(c => `<th>${c.pl || esc(c.l)}</th>`).join('')}</tr></thead><tbody>
    ${rows.map(x => `<tr>${cols.map(c => c.p(x, c2)).join('')}</tr>`).join('') || `<tr><td colspan="${cols.length}">Nenhuma atividade neste filtro.</td></tr>`}
    </tbody></table>`;
  const html = `<div class="p-page">${pHead('Atividades — ' + obraLabel(o))}${pFilters(flt)}<div class="p-sec">${esc(showTxt)} (${rows.length} atividade(s)) — situação no fechamento de ${esc(weekLabel(w, true))}</div>${tbl}
    <div class="p-note">Início e Término = data revisada quando preenchida (“revisado”, com os dias de deslocamento), senão a prevista. Início atrasado = passou o início sem avanço. Término atrasado = passou o término sem concluir. Planejado = parte do prazo previsto já decorrida.</div></div>`;
  doPrint(html);
}
function exportAtivExcel(){
  if (!libsOk()) return; const o = TK.obra; const m = obraModel(o, TK.week);
  const aoa = [['Obra', obraLabel(o)], ['Semana de referência', weekLabel(TK.week, true)], ['Avanço realizado (%)', +m.real.toFixed(1)], ['Avanço planejado (%)', +m.plan.toFixed(1)], [],
    ['Trecho', 'Serviço', 'Tipo', 'Und.', 'Qtd. Prevista', 'Qtd. Realizada', 'Qtd. Restante', '% realizado', '% planejado', 'Início Previsto', 'Término Previsto', 'Início Revisado', 'Término Revisado', 'Status', 'Impedimentos'],
    ...m.rows.map(x => { const q = Number(x.a.qtd) || 0; return [x.a.trecho || '', x.a.nome, x.a.tipo || '', x.a.unidade || '', q || '', q ? (x.acum ?? 0) : '', q ? x.restQtd : '', +x.real.toFixed(1), +x.plan.toFixed(1), fmtBR(x.a.inicio), fmtBR(x.a.fim), fmtBR(x.a.inicio_rev), fmtBR(x.a.revisado), x.status, x.a.impedimentos || '']; })];
  const wb = XLSX.utils.book_new(); const ws = XLSX.utils.aoa_to_sheet(aoa); ws['!cols'] = [{wch: 20}, {wch: 50}, ...Array(12).fill({wch: 14}), {wch: 30}];
  XLSX.utils.book_append_sheet(wb, ws, 'Atividades'); XLSX.writeFile(wb, `atividades_${obraLabel(o).replace(/\W+/g, '_')}_${TK.week}.xlsx`);
}
/* ---------- eventos das páginas de acompanhamento ---------- */
document.addEventListener('change', e => {
  const t = e.target;
  if (t.dataset.tk) { TK[t.dataset.tk] = t.value; refreshAll(); }
  else if (t.dataset.af) { TK.ativ[t.dataset.af] = t.value; if (t.dataset.af !== 'q') refreshAll(); }
  else if (t.dataset.rf) { TK.restr[t.dataset.rf] = t.value; refreshAll(); }
  else if (t.dataset.lps) {
    const id = t.dataset.id, f = t.dataset.lps, v = t.value;
    if (Repo.tx(() => { const r = Repo.data.lps.find(x => x.id === id); r[f] = v; if (f === 'status' && v !== 'Não concluído') r.causa = ''; r.updatedAt = new Date().toISOString(); }).ok) refreshAll();
  }
});
document.addEventListener('input', e => { if (e.target.dataset.af === 'q') { TK.ativ.q = e.target.value; clearTimeout(window._afT); window._afT = setTimeout(refreshAll, 250); } });
document.addEventListener('click', e => {
  const g = e.target.closest('[data-goto]'); if (g) { switchTab(g.dataset.goto); return; }
  const b = e.target.closest('[data-ta]'); if (!b) return;
  const a = b.dataset.ta, k = b.dataset.kind;
  if (a === 'new') openTrkForm(k);
  else if (a === 'edit') openTrkForm(k, Repo.data[k].find(x => x.id === b.dataset.id));
  else if (a === 'del') deleteTrk(k, b.dataset.id);
  else if (a === 'avanco') openAvancoForm();
  else if (a === 'import') openTrkImport();
  else if (a === 'template') downloadTemplate();
  else if (a === 'rmrestr') removeRestr(b.dataset.id);
  else if (a === 'copyprev') copyPrevLps();
  else if (a === 'xlsativ') exportAtivExcel();
  else if (a === 'pdfativ') printAtiv();
  else if (a === 'cols') openAtCols();

});

/* =====================================================================
   VISÃO GERENCIAL — diretor (todas as obras) e coordenador (uma obra)
   Só leitura. Obras nunca são somadas entre si.
   ===================================================================== */
const VG = {obra: '', week: ''};
const SEM = {r: {c: '#DC2626', l: 'Crítico', o: 0}, y: {c: '#D97706', l: 'Atenção', o: 1}, g: {c: '#16A34A', l: 'No prazo', o: 2}, n: {c: '#9CA3AF', l: 'Sem atividades', o: 3}};
const vgRef = w => w === todayWeek() ? todayYMD() : weekRange(w).end;
function vgEfetivo(kind, o, w, n = 6){
  const C = computeCells(kind, {...blankFilter(), obras: [o]});
  const weeks = []; for (let i = n - 1; i >= 0; i--) weeks.push(weekPlus(w, -i));
  const vals = weeks.map(x => { const c = C.cells.get(x + '||' + o); return c ? c.total : null; });
  const cur = vals[n - 1], prev = vals[n - 2];
  return {weeks, vals, cur, prev, d: cur != null && prev != null ? cur - prev : null};
}
/** Semáforo: 🔴 término vencido ou > 15 p.p. abaixo do planejado · 🟡 5–15 p.p. abaixo, impedimento ou início atrasado · 🟢 demais */
function vgSemaforo(m){
  if (!m.n) return {k: 'n', why: 'Nenhuma atividade cadastrada'};
  const pend = m.rows.filter(x => x.status !== 'Concluída');
  const tv = pend.filter(x => x.status === 'Término atrasado').length, ia = pend.filter(x => x.status === 'Início atrasado').length;
  const imp = pend.filter(x => String(x.a.impedimentos || '').trim()).length; const gap = m.plan - m.real;
  const pp = v => v.toLocaleString('pt-BR', {maximumFractionDigits: 1}) + ' p.p.';
  if (tv || gap > 15) return {k: 'r', why: [tv ? `${tv} serviço(s) com término vencido` : '', gap > 15 ? `${pp(gap)} abaixo do planejado` : ''].filter(Boolean).join(' · ')};
  if (gap > 5 || imp || ia) return {k: 'y', why: [gap > 5 ? `${pp(gap)} abaixo do planejado` : '', ia ? `${ia} com início atrasado` : '', imp ? `${imp} com impedimento` : ''].filter(Boolean).join(' · ')};
  return {k: 'g', why: 'Dentro do planejado'};
}
function vgObraData(o, w){
  const m = obraModel(o, w); const ref = vgRef(w);
  const pend = m.rows.filter(x => x.status !== 'Concluída');
  const prox = pend.filter(x => x.alvo && x.alvo >= ref).sort((a, b) => a.alvo.localeCompare(b.alvo))[0] || null;
  return {o, label: obraLabel(o), m, ref, sem: vgSemaforo(m), L: vgEfetivo('labor', o, w), E: vgEfetivo('equip', o, w),
    atr: m.rows.filter(isAtr).length, tv: pend.filter(x => x.status === 'Término atrasado').length, ia: pend.filter(x => x.status === 'Início atrasado').length,
    imp: pend.filter(x => String(x.a.impedimentos || '').trim()).length, rest: pend.length, prox};
}
const vgDelta = (d, unit) => d == null ? '' : `<span class="vg-d ${d > 0 ? 'up' : d < 0 ? 'dn' : ''}">${d > 0 ? '▲ +' : d < 0 ? '▼ ' : '= '}${d ? fmtN(d) : ''}</span>`;
const vgDot = k => `<i class="vg-dot" style="--c:${SEM[k].c}"></i>`;
const vgAvBar = m => m.n ? `<div class="vg-av"><div class="vg-avn"><b>${pctTxt(m.real)}</b><span>plan. ${pctTxt(m.plan)}</span></div>${progBar(m.real, m.plan)}<small class="${m.real - m.plan < -5 ? 'neg' : ''}">${m.real - m.plan >= 0 ? '+' : ''}${(m.real - m.plan).toLocaleString('pt-BR', {maximumFractionDigits: 1})} p.p.</small></div>` : '<span class="muted">—</span>';
const vgEf = (e, unit) => `<div class="vg-ef"><div><b>${e.cur == null ? '—' : fmtN(e.cur)}</b> ${vgDelta(e.d)}</div><div class="vg-sk">${sparkSVG(e.vals)}</div><small>${unit}</small></div>`;

function vgObras(){
  const mp = new Map();
  for (const k of ['ativ', 'labor', 'equip', 'manual']) for (const r of Repo.data[k]) { const o = obraKey(r.obra); if (o && !mp.has(o)) mp.set(o, r.obra); }
  return [...mp.keys()];
}
function vgPickers(sec){
  const wk = trkWeeks(); if (!VG.week || !weekRange(VG.week)) VG.week = todayWeek();
  const ws = $('[data-vgw]', sec); ws.innerHTML = wk.map(w => `<option value="${w}">${esc(weekLabel(w, true))}${w === todayWeek() ? ' — atual' : ''}</option>`).join(''); ws.value = VG.week;
}
function renderVG(){
  const sec = $('#tab-vg');
  if (!sec.dataset.built) {
    sec.dataset.built = 1;
    sec.innerHTML = `<div class="toolbar"><div><h2 data-vgt>Visão gerencial</h2><div class="sub" data-vgs></div></div><div class="acts"><label class="vg-wk">Semana de referência<select data-vgw></select></label><button class="btn dark" data-vg="pdf">⎙ PDF</button></div></div><div data-vgb></div>`;
  }
  vgPickers(sec);
  ['at-h-labor', 'at-h-equip', 'vg-h-labor', 'vg-h-equip'].forEach(killChart);
  if (VG.obra && !vgObras().includes(VG.obra)) VG.obra = '';
  if (VG.obra) return renderVGObra(sec);
  const w = VG.week; const list = vgObras().map(o => vgObraData(o, w)).sort((a, b) => SEM[a.sem.k].o - SEM[b.sem.k].o || (a.m.real - a.m.plan) - (b.m.real - b.m.plan) || a.label.localeCompare(b.label, 'pt-BR'));
  VG._list = list;
  $('[data-vgt]', sec).textContent = 'Visão gerencial — todas as obras';
  $('[data-vgs]', sec).innerHTML = `Situação no fechamento de ${esc(weekLabel(w, true))}. Clique na obra para ver o detalhe.`;
  const cnt = k => list.filter(x => x.sem.k === k).length;
  $('[data-vgb]', sec).innerHTML = !list.length ? `<div class="card"><div class="empty"><b>Nenhuma obra cadastrada</b>Importe a planilha de mão de obra, equipamentos ou atividades.</div></div>` : `
    <div class="vg-sum">${['r', 'y', 'g', 'n'].filter(k => cnt(k) || k !== 'n').map(k => `<div class="vg-pill" style="--c:${SEM[k].c}">${vgDot(k)}<b>${cnt(k)}</b> ${SEM[k].l}</div>`).join('')}
      <span class="muted vg-rule">🔴 término vencido ou &gt; 15 p.p. abaixo do planejado · 🟡 5–15 p.p. abaixo, impedimento ou início atrasado</span></div>
    <div class="card vg-card"><div class="vg-tbl">
      <div class="vg-row vg-h"><div>Obra</div><div>Avanço físico</div><div>Mão de obra</div><div>Equipamentos</div><div>Pendências</div><div>Próximo término</div></div>
      ${list.map(x => `<button type="button" class="vg-row" data-vg="obra" data-o="${esc(x.o)}" style="--c:${SEM[x.sem.k].c}">
        <div class="vg-ob">${vgDot(x.sem.k)}<div><b>${esc(x.label)}</b><small>${esc(x.sem.why)}</small></div></div>
        <div data-l="Avanço físico">${vgAvBar(x.m)}</div>
        <div data-l="Mão de obra">${vgEf(x.L, 'pessoas')}</div>
        <div data-l="Equipamentos">${vgEf(x.E, 'equipamentos')}</div>
        <div data-l="Pendências" class="vg-pd">${x.m.n ? `<span class="${x.atr ? 'bad' : ''}"><b>${x.atr}</b> atrasada(s)</span><span class="${x.imp ? 'warn' : ''}"><b>${x.imp}</b> com impedimento</span><span><b>${x.rest}</b> restante(s)</span>` : '<span class="muted">—</span>'}</div>
        <div data-l="Próximo término" class="vg-nx">${x.prox ? `<b>${fmtBR(x.prox.alvo)}</b><small>${esc(x.prox.a.nome)}</small><small class="muted">em ${daysBetween(x.ref, x.prox.alvo)} dia(s)</small>` : '<span class="muted">—</span>'}</div>
      </button>`).join('')}
    </div></div>`;
}
function vgAttention(d){
  const ref = d.ref, lim = ymdAdd(ref, 7); const pend = d.m.rows.filter(x => x.status !== 'Concluída');
  const term = pend.filter(x => x.alvo && x.alvo <= lim).sort((a, b) => a.alvo.localeCompare(b.alvo));
  const ini = pend.filter(x => x.status === 'Início atrasado').sort((a, b) => b.diasAtraso - a.diasAtraso);
  const imp = pend.filter(x => String(x.a.impedimentos || '').trim());
  return {term, ini, imp};
}
function vgUpcoming(d, nW = 3){
  const out = []; const from = d.ref, to = ymdAdd(d.ref, nW * 7);
  for (const x of d.m.rows) { if (x.status === 'Concluída') continue;
    const ini = x.a.inicio_rev || x.a.inicio; if (ini && ini > from && ini <= to && x.real === 0) out.push({x, t: 'Começa', d: ini});
    if (x.alvo && x.alvo > from && x.alvo <= to) out.push({x, t: 'Termina', d: x.alvo}); }
  return out.sort((a, b) => a.d.localeCompare(b.d));
}
const vgItem = (x, right) => `<li><div><b>${esc(x.a.nome)}</b><small>${esc(x.a.trecho || '')}</small></div><div class="r">${right}</div></li>`;
function renderVGObra(sec){
  const w = VG.week, o = VG.obra; const d = vgObraData(o, w); VG._obra = d;
  const A = vgAttention(d); const U = vgUpcoming(d); const H3 = ['labor', 'equip'].map(k => histo3Data(k, o, w));
  $('[data-vgt]', sec).innerHTML = `<button class="btn sm" data-vg="back">← Todas as obras</button> ${esc(d.label)}`;
  $('[data-vgs]', sec).innerHTML = `${vgDot(d.sem.k)} <b style="color:${SEM[d.sem.k].c}">${SEM[d.sem.k].l}</b> — ${esc(d.sem.why)} · fechamento de ${esc(weekLabel(w, true))}`;
  const box = (t, n, items, empty, cls) => `<div class="card vg-att ${cls}"><div class="card-h"><h3>${t}</h3><span class="chip ${n ? cls : ''}">${n}</span></div>${n ? `<ul class="vg-ul">${items}</ul>` : `<div class="muted" style="padding:6px 2px">${empty}</div>`}</div>`;
  const dAlvo = x => { const n = daysBetween(d.ref, x.alvo); return n < 0 ? `<span class="chip err">vencido há ${-n} d</span>` : n === 0 ? '<span class="chip err">vence hoje</span>' : `<span class="chip warn">em ${n} d</span>`; };
  $('[data-vgb]', sec).innerHTML = `
    <div class="vg-k4">
      <div class="kpi dark" style="--c:#F7931E"><div class="l">Avanço físico</div>${vgAvBar(d.m)}</div>
      <div class="kpi" style="--c:#DC2626"><div class="l">Atrasadas</div><div class="v">${fmtN(d.atr)}</div><div class="s">${d.tv} no término · ${d.ia} no início</div></div>
      <div class="kpi" style="--c:#D97706"><div class="l">Com impedimento</div><div class="v">${fmtN(d.imp)}</div><div class="s">de ${d.rest} serviço(s) restante(s)</div></div>
      <div class="kpi" style="--c:#334155"><div class="l">Efetivo da semana</div><div class="vg-k-ef"><div><b>${d.L.cur == null ? '—' : fmtN(d.L.cur)}</b> pessoas ${vgDelta(d.L.d)}</div><div><b>${d.E.cur == null ? '—' : fmtN(d.E.cur)}</b> equip. ${vgDelta(d.E.d)}</div></div></div>
    </div>
    <h3 class="vg-h3">Precisa de atenção</h3>
    <div class="vg-att3">
      ${box('Término vencido ou nos próximos 7 dias', A.term.length, A.term.slice(0, 10).map(x => vgItem(x, `${fmtBR(x.alvo)}<br>${dAlvo(x)}<small>${pctTxt(x.real, 0)} feito</small>`)).join('') + (A.term.length > 10 ? `<li class="muted">+${A.term.length - 10} outro(s)</li>` : ''), 'Nenhum término vencido ou vencendo.', 'err')}
      ${box('Início atrasado', A.ini.length, A.ini.slice(0, 10).map(x => vgItem(x, `${fmtBR(x.a.inicio_rev || x.a.inicio)}<br><span class="chip err">${x.diasAtraso} d</span>`)).join('') + (A.ini.length > 10 ? `<li class="muted">+${A.ini.length - 10} outro(s)</li>` : ''), 'Nenhum início atrasado.', 'err')}
      ${box('Impedimentos', A.imp.length, A.imp.slice(0, 10).map(x => vgItem(x, `<span class="chip warn">${esc(x.a.impedimentos)}</span>`)).join('') + (A.imp.length > 10 ? `<li class="muted">+${A.imp.length - 10} outro(s)</li>` : ''), 'Nenhum impedimento registrado.', 'warn')}
    </div>
    <div class="card"><div class="card-h"><h3>Próximas 3 semanas</h3><span class="hint">serviços que começam ou terminam até ${fmtBR(ymdAdd(d.ref, 21))}</span></div>
      ${U.length ? `<div class="tbl-wrap"><table><thead><tr><th>Data</th><th>Evento</th><th>Serviço</th><th>Trecho</th><th>Avanço</th><th>Impedimento</th></tr></thead><tbody>${U.map(u => `<tr><td><b>${fmtBR(u.d)}</b></td><td><span class="chip ${u.t === 'Começa' ? 'or' : ''}">${u.t}</span></td><td class="l" style="white-space:normal"><b>${esc(u.x.a.nome)}</b></td><td class="l">${esc(u.x.a.trecho || '')}</td><td>${pctTxt(u.x.real, 0)}</td><td class="l">${u.x.a.impedimentos ? `<span class="chip warn">${esc(u.x.a.impedimentos)}</span>` : ''}</td></tr>`).join('')}</tbody></table></div>` : '<div class="muted" style="padding:6px 2px">Nenhum serviço começa ou termina nas próximas 3 semanas.</div>'}</div>
    <div class="card"><div class="card-h"><h3>Histograma — últimas 3 semanas</h3><span class="hint">efetivo no fechamento de cada semana</span><span class="sp"></span><button class="btn sm" data-vg="ativ">Ver todas as atividades →</button></div>
      <div class="h3-grid">${H3.map(h => `<div class="h3-box"><div class="h3-top"><div class="h3-tt">${h.kind === 'labor' ? 'Mão de obra' : 'Equipamentos'}</div>${h.any ? histo3Read(h) : ''}</div>${h.any ? `<div class="jr-frame"><div style="position:relative;height:240px"><canvas id="vgH-${h.kind}"></canvas></div></div>${histo3Table(h)}` : '<div class="empty" style="padding:22px"><b>Sem histograma nestas 3 semanas</b></div>'}</div>`).join('')}</div></div>`;
  for (const h of H3) { const cv = $(`#vgH-${h.kind}`); if (cv && h.any) charts['vg-h-' + h.kind] = new Chart(cv, histo3Config(h, false)); }
  VG._obra.H3 = H3; VG._obra.A = A; VG._obra.U = U;
}
document.addEventListener('click', e => { if (e.target.closest('[data-vgo]')) { VG.obra = TK.obra; VG.week = TK.week; switchTab('vg'); } });
document.addEventListener('change', e => { if (e.target.matches('[data-vgw]')) { VG.week = e.target.value; refreshAll(); } });
document.addEventListener('click', e => {
  const b = e.target.closest('[data-vg]'); if (!b) return;
  const a = b.dataset.vg;
  if (a === 'obra') { VG.obra = b.dataset.o; refreshAll(); window.scrollTo(0, 0); }
  else if (a === 'back') { VG.obra = ''; refreshAll(); }
  else if (a === 'ativ') { TK.obra = VG.obra; TK.week = VG.week; switchTab('ativ'); }
  else if (a === 'pdf') printVG();
});

/* ---------- PDF da visão gerencial ---------- */
function printVG(){
  if (!libsOk()) return toast('Bibliotecas não carregadas.', 'err');
  const w = VG.week;
  if (!VG.obra) {
    const list = VG._list || []; if (!list.length) return toast('Nenhuma obra para o PDF.', 'err');
    const cnt = k => list.filter(x => x.sem.k === k).length;
    const html = `<div class="p-page">${pHead('Visão gerencial — todas as obras')}${pFilters([['Semana de referência', weekLabel(w, true)], ['Obras', String(list.length)], ['Situação', ['r', 'y', 'g'].map(k => `${cnt(k)} ${SEM[k].l.toLowerCase()}`).join(' · ')]])}
      <table class="p-at p-vg"><colgroup><col style="width:22%"><col style="width:17%"><col style="width:12%"><col style="width:12%"><col style="width:14%"><col style="width:23%"></colgroup>
      <thead><tr><th>Obra</th><th>Avanço físico (real × plan.)</th><th>Mão de obra</th><th>Equipamentos</th><th>Pendências</th><th>Próximo término</th></tr></thead><tbody>
      ${list.map(x => `<tr><td class="l"><span class="pdot" style="background:${SEM[x.sem.k].c}"></span><b>${esc(x.label)}</b><span class="d" style="color:${SEM[x.sem.k].c}">${esc(x.sem.why)}</span></td>
        <td>${x.m.n ? `<b>${pctTxt(x.m.real)}</b> × ${pctTxt(x.m.plan)}<span class="d" style="color:${x.m.real - x.m.plan < -5 ? '#DC2626' : '#16A34A'}">${(x.m.real - x.m.plan).toLocaleString('pt-BR', {maximumFractionDigits: 1})} p.p.</span>` : '—'}</td>
        <td><b>${x.L.cur == null ? '—' : fmtN(x.L.cur)}</b>${x.L.d != null ? `<span class="d" style="color:${x.L.d < 0 ? '#DC2626' : '#16A34A'}">${x.L.d > 0 ? '+' : ''}${fmtN(x.L.d)} vs. sem. ant.</span>` : ''}</td>
        <td><b>${x.E.cur == null ? '—' : fmtN(x.E.cur)}</b>${x.E.d != null ? `<span class="d" style="color:${x.E.d < 0 ? '#DC2626' : '#16A34A'}">${x.E.d > 0 ? '+' : ''}${fmtN(x.E.d)} vs. sem. ant.</span>` : ''}</td>
        <td>${x.m.n ? `${x.atr} atrasada(s)<span class="d" style="color:#B45309">${x.imp} com impedimento</span>` : '—'}</td>
        <td class="l">${x.prox ? `<b>${fmtBR(x.prox.alvo)}</b> <span style="color:#6B7280">(${daysBetween(x.ref, x.prox.alvo)} d)</span><span class="c2" style="font-size:7px;color:#374151">${esc(x.prox.a.nome)}</span>` : '—'}</td></tr>`).join('')}
      </tbody></table><div class="p-note">🔴 Crítico: término vencido ou mais de 15 p.p. abaixo do planejado · 🟡 Atenção: 5–15 p.p. abaixo, impedimento ou início atrasado · 🟢 No prazo. Avanço = média do % realizado dos serviços (ponderada pelo Peso, quando informado). Obras não são somadas entre si.</div></div>`;
    return doPrint(html);
  }
  const d = VG._obra; if (!d) return;
  const lst = (arr, f) => arr.length ? `<ul class="p-ul">${arr.slice(0, 8).map(f).join('')}${arr.length > 8 ? `<li style="color:#6B7280">+${arr.length - 8} outro(s)</li>` : ''}</ul>` : '<div class="p-note">Nenhum.</div>';
  const it = (x, r) => `<li><span>${esc(x.a.nome)}<em>${esc(x.a.trecho || '')}</em></span><b>${r}</b></li>`;
  const kp = `<div class="p-k4">
    <div class="pk"><div class="lb">Avanço físico</div><div class="vl" style="color:#E86F00">${d.m.n ? pctTxt(d.m.real) : '—'}</div><div class="sb">planejado ${d.m.n ? pctTxt(d.m.plan) : '—'}</div></div>
    <div class="pk"><div class="lb">Atrasadas</div><div class="vl" style="color:#DC2626">${d.atr}</div><div class="sb">${d.tv} no término · ${d.ia} no início</div></div>
    <div class="pk"><div class="lb">Com impedimento</div><div class="vl" style="color:#B45309">${d.imp}</div><div class="sb">de ${d.rest} restante(s)</div></div>
    <div class="pk"><div class="lb">Efetivo da semana</div><div class="vl">${d.L.cur == null ? '—' : fmtN(d.L.cur)} <small>pessoas</small> · ${d.E.cur == null ? '—' : fmtN(d.E.cur)} <small>equip.</small></div><div class="sb">${d.L.d != null ? `MO ${d.L.d > 0 ? '+' : ''}${fmtN(d.L.d)}` : ''}${d.E.d != null ? ` · equip. ${d.E.d > 0 ? '+' : ''}${fmtN(d.E.d)}` : ''} vs. semana anterior</div></div></div>`;
  const att = `<div class="p-att3">
    <div><div class="p-h3t">Término vencido ou em 7 dias (${d.A.term.length})</div>${lst(d.A.term, x => it(x, `${fmtBR(x.alvo)} · ${pctTxt(x.real, 0)}`))}</div>
    <div><div class="p-h3t">Início atrasado (${d.A.ini.length})</div>${lst(d.A.ini, x => it(x, `${x.diasAtraso} d`))}</div>
    <div><div class="p-h3t">Impedimentos (${d.A.imp.length})</div>${lst(d.A.imp, x => it(x, esc(x.a.impedimentos)))}</div></div>`;
  const up = d.U.length ? `<table class="p-at"><colgroup><col style="width:10%"><col style="width:9%"><col style="width:44%"><col style="width:20%"><col style="width:17%"></colgroup><thead><tr><th>Data</th><th>Evento</th><th>Serviço</th><th>Trecho</th><th>Impedimento</th></tr></thead><tbody>${d.U.slice(0, 8).map(u => `<tr><td><b>${fmtBR(u.d)}</b></td><td>${u.t}</td><td class="l"><span class="c2">${esc(u.x.a.nome)}</span></td><td class="l"><span class="c2">${esc(u.x.a.trecho || '')}</span></td><td>${esc(u.x.a.impedimentos || '')}</td></tr>`).join('')}</tbody></table>` : '<div class="p-note">Nenhum serviço começa ou termina nas próximas 3 semanas.</div>';
  const hist = `<div class="p-h3">${d.H3.map(h => `<div class="p-h3b"><div class="p-h3top"><div class="p-h3t">${h.kind === 'labor' ? 'Mão de obra' : 'Equipamentos'}</div>${h.any ? histo3Read(h) : ''}</div>${h.any ? `<img src="${offscreenImage(histo3Config(h, true), 640, 300)}" style="border:1px solid #F1F2F4;border-radius:8px">` : '<div class="p-note">Sem histograma nestas 3 semanas.</div>'}</div>`).join('')}</div>`;
  const html = `<div class="p-page">${pHead('Visão gerencial — ' + d.label)}<div class="p-filt"><span class="pdot" style="background:${SEM[d.sem.k].c}"></span><b style="color:${SEM[d.sem.k].c}">${SEM[d.sem.k].l}</b> — ${esc(d.sem.why)} &nbsp;·&nbsp; <b>Semana:</b> ${esc(weekLabel(w, true))}</div>
    ${kp}<div class="p-sec">Precisa de atenção</div>${att}<div class="p-sec">Próximas 3 semanas</div>${up}<div class="p-sec">Histograma — últimas 3 semanas</div>${hist}</div>`;
  doPrint(html);
}

function refreshAll(){
  const T = {resumo: renderResumo, config: renderConfig, avisos: renderAvisos, painel: renderPainel, ativ: renderAtivPage, vg: renderVG, lps: renderLps, restr: renderRestr}[UI.tab];
  if (T) T(); else renderKind(UI.tab);
  updateBadge();
}
function paintSideMini(err){
  const el = document.getElementById('sideMini'); if (!el) return;
  el.innerHTML = Repo.mode === 'db'
    ? `<span style="display:inline-block;width:7px;height:7px;border-radius:50%;background:${err ? '#EF4444' : DB.pending ? '#F7931E' : '#22C55E'};margin-right:5px"></span>Banco compartilhado${DB.pending ? ' · salvando…' : ''}`
    : `<span style="display:inline-block;width:7px;height:7px;border-radius:50%;background:#A1A1AA;margin-right:5px"></span>Modo local (este navegador)`;
}
function renderConfig(){
  const ch = document.getElementById('curHost'); if (ch) ch.textContent = location.host || 'arquivo local';
  const c = sheetCfg();
  const short = c.url.length > 70 ? c.url.slice(0, 67) + '…' : c.url;
  $('#cfgSheet').innerHTML = `Lê as abas <b>MO</b>, <b>Equipamento</b> e <b>Resumo</b> (somente leitura — nada é gravado no Google Sheets).<br>Link: <a href="${esc(c.url)}" target="_blank" rel="noopener">${esc(short)}</a><br>Leitura automática ao abrir: <b>${c.auto ? 'ligada' : 'desligada'}</b>${Sync.at ? `<br>Última leitura: <b>${Sync.at.toLocaleString('pt-BR')}</b>` : ''}`;
  paintDbStatus(); paintStorage();
  const sl = Snap.list();
  $('#snapList').innerHTML = sl.length ? sl.map((x, i) => `<div style="display:flex;gap:8px;align-items:center;padding:6px 0;border-top:1px solid var(--line2)"><div style="flex:1"><b>${new Date(x.t).toLocaleString('pt-BR')}</b><br><span class="muted" style="font-size:12px">${esc(x.label)} · MO ${x.counts.labor} · Equip. ${x.counts.equip} · Manuais ${x.counts.manual}</span></div><button class="btn sm" data-snap="${i}">Restaurar</button></div>`).join('') : '<span class="muted">Nenhum ponto ainda.</span>';
  $$('#snapList [data-snap]').forEach(b => b.onclick = () => Snap.restore(+b.dataset.snap));
}
function avisosAtuais(){
  const out = [];
  out.push({type: Repo.mode === 'db' ? 'ok' : 'info', msg: Repo.mode === 'db' ? 'Os lançamentos são salvos no banco compartilhado (Supabase); todos os usuários autorizados veem as mesmas informações. A planilha conectada é só lida.' : 'Os lançamentos ficam salvos somente neste navegador/dispositivo (armazenamento local). A planilha conectada é só lida (nada é gravado nela). Use “Baixar backup” regularmente.'});
  for (const kind of ['labor', 'equip']) {
    const model = UI._last[kind] || histogramModel(kind, UI.f[kind], UI.mode[kind]);
    const hn = histNotes(kind, model);
    if (hn.extra) out.push({type: 'warn', kind, html: hn.extra.replace(/^<div[^>]*>|<\/div>$/g, '')});
    hn.notes.forEach(n => out.push({type: 'info', kind, msg: n}));
  }
  return out;
}
function updateBadge(){
  const b = document.getElementById('avisosBadge'); if (!b) return;
  const n = (typeof UI !== 'undefined' && UI.tab === 'avisos') ? 0 : Log.unread;
  b.hidden = !n; b.textContent = n > 99 ? '99+' : n;
}
/* ---------- Conferência: planilha × app, por obra e semana ---------- */
function conferencia(kind){
  const wb = Sync.lastWb; if (!wb) return null;
  const sh = guessSheet(kind, wb); if (!sh) return {sheet: null, groups: []};
  const grid = sheetGrid(wb.Sheets[sh]); const hdr = mapHeaders(kind, grid); if (!hdr) return {sheet: sh, groups: []};
  const res = analyzeImport(kind, grid, hdr, '', false, detectDateOrder(kind, grid, hdr) || detectWbOrder(wb) || 'dmy');
  const K = KINDS[kind]; const G = new Map();
  const g = (w, o, label) => { const k = w + '||' + o; if (!G.has(k)) G.set(k, {w, o, label, linhas: 0, cats: {}, pend: 0, dup: 0, motivos: {}, app: 0, ativos: 0, foraAdm: 0, foraDem: 0, appCats: {}, manual: 0, faltando: 0, sobrando: 0}); return G.get(k); };
  const keysPlan = new Map();
  for (const r of res) {
    const w = r.rec.semana || parseWeek(r.rec.semana).v; const o = obraKey(r.rec.obra || ''); if (!w || !o) continue;
    const x = g(w, o, r.rec.obra); x.linhas++; const c = r.rec[K.catField]; if (c) x.cats[c] = (x.cats[c] || 0) + 1;
    if (r.errors.length) { x.pend++; if (r.errors.some(e => /Duplicado/.test(e.msg))) x.dup++; r.errors.forEach(e => x.motivos[e.field + ': ' + e.msg.replace(/".*"/, '…')] = (x.motivos[e.field + ': ' + e.msg.replace(/".*"/, '…')] || 0) + 1); }
    else { if (!keysPlan.has(w + '||' + o)) keysPlan.set(w + '||' + o, new Set()); keysPlan.get(w + '||' + o).add(recKey(kind, r.rec)); }
  }
  const mm = manualMap(kind);
  for (const x of G.values()) {
    const end = weekRange(x.w).end; const seen = new Set(); const ks = keysPlan.get(x.w + '||' + x.o) || new Set();
    for (const r of Repo.data[kind]) {
      if (r.semana !== x.w || obraKey(r.obra) !== x.o) continue; x.app++;
      if (!ks.has(recKey(kind, r))) x.sobrando++;
      if (r[K.start] && r[K.start] > end) { x.foraAdm++; continue; }
      if (r[K.end] && r[K.end] < end) { x.foraDem++; continue; }
      const vk = vincKey(kind, r); if (seen.has(vk)) continue; seen.add(vk); x.ativos++; x.appCats[r[K.catField]] = (x.appCats[r[K.catField]] || 0) + 1;
    }
    x.faltando = Math.max(0, (x.linhas - x.pend) - (x.app - x.sobrando));
    const m = mm.get(x.w + '|' + x.o);
    if (m) { const cell = computeCells(kind, {...blankFilter(), obras: [x.o], weeks: [x.w]}).cells.get(x.w + '||' + x.o); x.manual = cell ? cell.total - x.ativos : 0; }
    x.totalApp = x.ativos + x.manual;
    x.ok = x.totalApp === x.linhas;
  }
  return {sheet: sh, groups: [...G.values()].sort((a, b) => b.w.localeCompare(a.w) || String(a.label).localeCompare(String(b.label), 'pt-BR'))};
}
function conferenciaHTML(){
  if (!Sync.lastWb) return `<div class="muted" style="font-size:13px">Nenhuma planilha lida ainda nesta sessão. <a href="#" data-sync="run">Ler a planilha conectada agora</a> ou use “Importar planilha”.</div>`;
  const fl = UI.f.labor; const filtros = [fl.sit && 'Situação', fl.cat && 'Categoria', fl.q && 'Busca'].filter(Boolean);
  let html = `<div class="muted" style="font-size:12px;margin-bottom:8px">Fonte: ${esc(Sync.lastSrc)} · lida às ${Sync.lastAt.toLocaleTimeString('pt-BR', {hour: '2-digit', minute: '2-digit'})}. A conferência usa <b>todos</b> os registros, sem os filtros de tela${filtros.length ? ` — atenção: na aba Mão de obra há filtro de <b>${filtros.join(', ')}</b> ativo, o que muda os números exibidos lá` : ''}.</div>`;
  for (const kind of ['labor', 'equip']) {
    const c = conferencia(kind); const K = KINDS[kind];
    if (!c || !c.groups.length) { html += `<div class="muted" style="font-size:13px;margin:6px 0">${K.title}: aba não encontrada ou sem linhas.</div>`; continue; }
    html += `<div style="font-weight:900;margin:10px 0 6px">${K.title} — aba “${esc(c.sheet)}”</div><div class="tbl-wrap"><table><thead><tr><th>Semana</th><th>Obra</th><th>Resultado</th><th>Linhas na planilha</th><th>Pendentes (não importadas)</th><th>Faltam importar</th><th>No app</th><th>Fora do fechamento</th><th>Contados no app</th><th>+ Manual / Resumo</th><th>Total no app</th></tr></thead><tbody>
    ${c.groups.map(x => { const why = []; if (x.pend) why.push(`${x.pend} com erro (${Object.entries(x.motivos).map(([m, n]) => `${n}× ${m}`).join('; ')})`); if (x.faltando) why.push(`${x.faltando} ainda não importada(s) — use “Revisar e importar”`); if (x.foraAdm) why.push(`${x.foraAdm} com ${K.startLabel.toLowerCase()} depois do domingo`); if (x.foraDem) why.push(`${x.foraDem} encerrado(s) antes do domingo`); if (x.sobrando) why.push(`${x.sobrando} no app que não estão mais na planilha`); if (x.manual) why.push(`${x.manual} somado(s) por lançamento manual/aba Resumo`);
      const diag = x.ok ? '<span class="chip ok">Igual à planilha</span>' : `<span class="chip warn">Diferença ${x.totalApp - x.linhas > 0 ? '+' : ''}${x.totalApp - x.linhas}</span> <span style="font-size:12px">${esc(why.join(' · ') || 'verificar')}</span>`;
      return `<tr><td>${esc(x.w)}</td><td class="l"><b>${esc(x.label)}</b></td><td class="l" style="white-space:normal;min-width:300px">${diag}</td><td class="num">${x.linhas}</td><td class="num">${x.pend || ''}</td><td class="num">${x.faltando || ''}</td><td class="num">${x.app}</td><td class="num">${(x.foraAdm + x.foraDem) || ''}</td><td class="num">${x.ativos}</td><td class="num">${x.manual || ''}</td><td class="num"><b>${x.totalApp}</b></td></tr>`; }).join('')}
    </tbody></table></div>`;
  }
  return html;
}
function renderAvisos(){
  Log.unread = 0; updateBadge();
  const ic = t => t === 'warn' ? '!' : t === 'err' ? '×' : t === 'ok' ? '✓' : 'i';
  const cur = avisosAtuais();
  const lbl = k => k ? `<span class="chip" style="margin-right:6px">${KINDS[k].title}</span>` : '';
  $('#avisosBody').innerHTML = `
    <div class="card av-sec"><div class="card-h"><h3>Orientações atuais</h3><span class="hint">Atualizadas conforme os dados e os filtros de cada aba</span></div>
      ${cur.map(a => `<div class="av-item ${a.type}"><span class="ic">${ic(a.type)}</span><div>${lbl(a.kind)}${a.html || esc(a.msg)}</div></div>`).join('')}</div>
    <div class="card av-sec"><div class="card-h"><h3>Conferência com a planilha</h3><span class="hint">Linhas da planilha × efetivo contado no app, por obra e semana, com o motivo de cada diferença</span><span class="sp"></span><button class="btn sm" data-sync="run">⟳ Ler planilha</button></div>${conferenciaHTML()}</div>
    <div class="card av-sec"><div class="card-h"><h3>Histórico de mensagens</h3><span class="hint">Últimas ${Log.items.length} mensagens exibidas pelo aplicativo neste navegador</span></div>
      ${Log.items.length ? Log.items.map(a => `<div class="av-item ${a.type || 'info'}"><span class="ic">${ic(a.type)}</span><div>${esc(a.msg)}</div><span class="tm">${new Date(a.t).toLocaleString('pt-BR')}</span></div>`).join('') : '<div class="muted" style="font-size:13px">Nenhuma mensagem registrada.</div>'}</div>`;
  $$('#avisosBody [data-basis]').forEach(b => { const box = b.closest('.av-item'); const kind = box && box.querySelector('.chip') && box.querySelector('.chip').textContent === 'Equipamentos' ? 'equip' : 'labor'; b.onclick = () => { UI.f[kind].basis = 'datas'; const sel = $(`#tab-${kind} [data-f=basis]`); if (sel) sel.value = 'datas'; switchTab(kind); }; });
}
function switchTab(t){
  UI.tab = t; $('#pageTitle').textContent = {labor: 'Mão de obra', equip: 'Equipamentos', resumo: 'Resumo', config: 'Configurações', avisos: 'Avisos', painel: 'Painel da obra', ativ: 'Atividades', vg: 'Visão gerencial', lps: 'Planejamento semanal', restr: 'Restrições'}[t]; window.scrollTo(0, 0); $$('#tabs button').forEach(b => b.classList.toggle('on', b.dataset.tab === t));
  $$('.tab').forEach(s => s.classList.toggle('on', s.id === 'tab-' + t));
  refreshAll();
}
function init(){
  try { Repo.lastBytes = (localStorage.getItem(LS_KEY) || '').length * 2; } catch (e) {} paintLogo(); paintStorage();
  if (!libsOk()) {
    $('main').insertAdjacentHTML('afterbegin', '<div class="alert err">Não foi possível carregar as bibliotecas de gráficos/Excel (Chart.js e SheetJS). Verifique a conexão com a internet e recarregue a página. Seus dados continuam salvos neste navegador.</div>');
    return;
  }
  setupChartDefaults();
  buildKindTab('labor'); buildKindTab('equip'); buildResumo();
  $('#tabs').onclick = e => { const b = e.target.closest('button[data-tab]'); if (b) switchTab(b.dataset.tab); };
  $('#btnBackup').onclick = backup;
  $('#btnClearData').onclick = openClearData;
  $('#btnClearLog').onclick = () => { Log.clear(); renderAvisos(); };
  $('#btnLogo').onclick = () => $('#logoBox').click();
  $('#btnLogoReset').onclick = () => { try { localStorage.removeItem(LS_LOGO); } catch (e) {} paintLogo(); toast('Logo COSAMPA restaurado.', 'ok'); };
  $('#btnRestore').onclick = () => { $('#fileRestore').value = ''; $('#fileRestore').click(); };
  $('#fileRestore').onchange = e => e.target.files[0] && restore(e.target.files[0]);
  $('#logoBoxM').onclick = () => $('#logoBox').click();
  $('#logoBox').onclick = async () => {
    if (getLogo() && await confirmBox('Logo', '<p>Deseja substituir o logo atual? (Cancelar mantém o logo)</p>', 'Escolher novo logo') === false) return;
    $('#fileLogo').value = ''; $('#fileLogo').click();
  };
  $('#fileLogo').onchange = e => { const f = e.target.files[0]; if (!f) return; const rd = new FileReader(); rd.onload = () => { try { localStorage.setItem(LS_LOGO, rd.result); paintLogo(); toast('Logo atualizado (salvo neste navegador).', 'ok'); } catch (er) { toast('Imagem muito grande para salvar.', 'err'); } }; rd.readAsDataURL(f); };
  let rT; window.addEventListener('resize', () => { clearTimeout(rT); rT = setTimeout(() => { if ((UI.tab === 'labor' || UI.tab === 'equip') && UI._last[UI.tab]) renderHist(UI.tab, UI._last[UI.tab]); }, 250); });
  switchTab('labor');
  $('#btnSync').onclick = $('#btnSync2').onclick = $('#btnSyncM').onclick = () => Sync.run(true);
  $('#btnSyncCfg').onclick = () => Sync.config();
  // Dados: banco compartilhado (se configurado) ou armazenamento local
  DB.loadSess();
  if (DB.cfg()) { paintDbStatus(); if (DB.sess) enterDbMode(); else showLogin(); }
  else { Repo.mode = 'local'; Repo.load(); paintDbStatus(); refreshAll(); Sync.auto(); }
  window.addEventListener('beforeunload', e => { if (Repo.mode === 'db' && DB.pending) { e.preventDefault(); e.returnValue = ''; } });
  let lastFocus = Date.now(); window.addEventListener('focus', () => { if (Repo.mode === 'db' && !DB.pending && Date.now() - lastFocus > 5 * 60000 && !document.querySelector('.modal-bg')) { lastFocus = Date.now(); reloadFromDb(true); } });
}
init();
</script>
</body>
</html>
