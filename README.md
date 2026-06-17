<!DOCTYPE html>
<html lang="fr">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width, initial-scale=1.0"/>
<title>Suivi Projet</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<style>
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0}
:root{
  /* Palette Bleu Marine + Blanc + Or */
  --bg:#EEF1F7;
  --surface:#FFFFFF;
  --border:#D4DBE8;
  --border2:#B0BCCE;
  --text:#0D1B2E;
  --text2:#3A4F6A;
  --text3:#7A90A8;

  /* Bleu marine */
  --blue:#0A2E5C;
  --blue-mid:#1A4D8F;
  --blue-light:#2E6BC4;
  --blue-bg:#E8EFF9;
  --blue-b:#9ABAE8;

  /* Or / Ambre */
  --gold:#B8860B;
  --gold-light:#D4A017;
  --gold-bg:#FDF6E3;
  --gold-b:#E8CC7A;

  /* Statuts */
  --green:#1A6B3A;--green-bg:#E6F4EC;--green-b:#85CCA0;
  --amber:#9A6200;--amber-bg:#FFF8E6;--amber-b:#E8C56A;
  --red:#B02020;--red-bg:#FBEAEA;--red-b:#E8A0A0;

  --r:12px;--rs:8px;
  --sh:0 2px 8px rgba(10,46,92,.08),0 1px 3px rgba(10,46,92,.05);
  --sh-hover:0 6px 20px rgba(10,46,92,.14);
  --sh-card:0 1px 4px rgba(10,46,92,.07),0 4px 16px rgba(10,46,92,.05);
}

body{font-family:'Segoe UI',system-ui,-apple-system,sans-serif;background:var(--bg);color:var(--text);min-height:100vh;font-size:14px;-webkit-font-smoothing:antialiased}

/* ── HEADER ── */
.hdr{
  background:linear-gradient(135deg, #0A2E5C 0%, #1A4D8F 100%);
  border-bottom:3px solid var(--gold-light);
  padding:.9rem 1.5rem;
  display:flex;align-items:center;justify-content:space-between;
  flex-wrap:wrap;gap:10px;
  position:sticky;top:0;z-index:100;
  box-shadow:0 2px 12px rgba(10,46,92,.25);
}
.hdr-logo{
  font-size:18px;font-weight:800;letter-spacing:-.3px;
  display:flex;align-items:center;gap:8px;color:#fff;
}
.hdr-logo-icon{
  width:32px;height:32px;
  background:linear-gradient(135deg,var(--gold-light),var(--gold));
  border-radius:8px;
  display:flex;align-items:center;justify-content:center;
  font-size:16px;box-shadow:0 2px 6px rgba(0,0,0,.2);
}
.hdr-logo span{color:var(--gold-light)}
.hdr-proj{display:flex;align-items:center;gap:10px;flex-wrap:wrap}
.proj-name{font-size:13px;font-weight:700;color:#fff}
.proj-meta{font-size:11px;color:rgba(255,255,255,.6);margin-top:1px}
.btn-cfg{
  font-size:11px;font-weight:600;padding:6px 13px;
  border:1.5px solid rgba(255,255,255,.3);border-radius:var(--rs);
  background:rgba(255,255,255,.1);cursor:pointer;color:#fff;
  display:flex;align-items:center;gap:5px;
  transition:all .15s;backdrop-filter:blur(4px);
}
.btn-cfg:hover{background:rgba(255,255,255,.2);border-color:var(--gold-light)}

/* ── TABS ── */
.tabs{
  background:var(--blue);
  border-bottom:1px solid rgba(255,255,255,.1);
  padding:0 1.5rem;
  display:flex;gap:0;
  overflow-x:auto;
  -webkit-overflow-scrolling:touch;
  scrollbar-width:none;
}
.tabs::-webkit-scrollbar{display:none}
.tab{
  font-size:12px;font-weight:600;
  padding:12px 18px;
  border:none;background:none;cursor:pointer;
  color:rgba(255,255,255,.55);
  border-bottom:2.5px solid transparent;
  margin-bottom:-1px;white-space:nowrap;
  transition:color .15s,border-color .15s;
  display:flex;align-items:center;gap:6px;
}
.tab:hover{color:rgba(255,255,255,.85)}
.tab.active{color:#fff;border-bottom-color:var(--gold-light)}
.tab-icon{font-size:14px}

/* ── LAYOUT ── */
.main{padding:1.25rem 1.5rem;max-width:1140px;margin:0 auto}
.panel{display:none}.panel.active{display:block}
.card{
  background:var(--surface);
  border:1px solid var(--border);
  border-radius:var(--r);
  padding:1.25rem;
  margin-bottom:1rem;
  box-shadow:var(--sh-card);
  transition:box-shadow .2s;
}
.card:hover{box-shadow:var(--sh-hover)}
.ct{
  font-size:13px;font-weight:700;color:var(--text);
  margin-bottom:1rem;
  display:flex;align-items:center;justify-content:space-between;
  flex-wrap:wrap;gap:8px;
  padding-bottom:10px;
  border-bottom:2px solid var(--bg);
}
.two-col{display:grid;grid-template-columns:1fr 1fr;gap:1rem}

/* ── KPI STRIP ── */
.kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(130px,1fr));gap:10px;margin-bottom:1rem}
.kpi{
  background:var(--surface);
  border:1px solid var(--border);
  border-radius:var(--r);
  padding:1rem 1.1rem;
  box-shadow:var(--sh);
  transition:box-shadow .2s,transform .2s;
  position:relative;overflow:hidden;
}
.kpi:hover{box-shadow:var(--sh-hover);transform:translateY(-2px)}
.kpi::before{
  content:'';position:absolute;top:0;left:0;right:0;height:3px;
  background:var(--border);border-radius:3px 3px 0 0;
}
.kpi-g::before{background:var(--green)}
.kpi-a::before{background:var(--gold-light)}
.kpi-r::before{background:var(--red)}
.kpi-b::before{background:var(--blue-mid)}
.kl{font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.6px;color:var(--text3);margin-bottom:6px}
.kv{font-size:22px;font-weight:800;line-height:1;letter-spacing:-.5px}
.ks{font-size:11px;color:var(--text3);margin-top:4px}
.kv-g{color:var(--green)}.kv-a{color:var(--gold)}.kv-r{color:var(--red)}.kv-b{color:var(--blue-mid)}

/* ── PILLS ── */
.pill{display:inline-flex;align-items:center;font-size:10px;font-weight:700;padding:3px 9px;border-radius:20px;white-space:nowrap;letter-spacing:.2px}
.pg{background:var(--green-bg);color:var(--green);border:1px solid var(--green-b)}
.pa{background:var(--amber-bg);color:var(--amber);border:1px solid var(--amber-b)}
.pr{background:var(--red-bg);color:var(--red);border:1px solid var(--red-b)}
.pb{background:var(--blue-bg);color:var(--blue-mid);border:1px solid var(--blue-b)}
.px{background:var(--bg);color:var(--text3);border:1px solid var(--border)}

/* ── PROGRESS ── */
.pbar{height:8px;border-radius:8px;background:var(--border);overflow:hidden}
.pbar-sm{height:5px;border-radius:5px;background:var(--border);overflow:hidden;margin-top:4px}
.pfill{height:100%;border-radius:8px;transition:width .4s ease}

/* ── CHARTS ── */
.ch-wrap{position:relative;width:100%;height:240px}
.ch-sm{position:relative;width:100%;height:200px}

/* ── TABLE ── */
.tbl-wrap{overflow-x:auto;-webkit-overflow-scrolling:touch;border-radius:var(--rs);border:1px solid var(--border)}
.tbl{width:100%;border-collapse:collapse;font-size:12px;min-width:600px}
.tbl th{
  font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.5px;
  color:#fff;padding:10px 12px;text-align:left;
  border-bottom:1px solid rgba(255,255,255,.15);
  background:linear-gradient(135deg,var(--blue) 0%,var(--blue-mid) 100%);
  white-space:nowrap;
}
.tbl td{padding:11px 12px;border-bottom:1px solid var(--border);vertical-align:middle}
.tbl tr:last-child td{border-bottom:none}
.tbl tr:hover td{background:var(--blue-bg)}
.btn-ico{
  font-size:12px;padding:5px 9px;
  border:1px solid var(--border);border-radius:var(--rs);
  background:none;cursor:pointer;color:var(--text2);
  transition:all .15s;
}
.btn-ico:hover{background:var(--blue-bg);border-color:var(--blue-b);color:var(--blue-mid)}

/* ── MOBILE TABLE CARDS ── */
.mob-act-card{
  background:var(--surface);border:1px solid var(--border);
  border-radius:var(--r);padding:12px 14px;margin-bottom:8px;
  box-shadow:var(--sh);
  border-left:3px solid var(--blue-mid);
}
.mob-act-name{font-size:13px;font-weight:700;margin-bottom:6px;color:var(--text)}
.mob-act-row{display:flex;justify-content:space-between;align-items:center;font-size:11px;color:var(--text3);margin-bottom:4px}
.mob-act-row strong{color:var(--text2)}

/* ── FORM ── */
.fgrid{display:grid;grid-template-columns:1fr 1fr;gap:12px}
.fg{display:flex;flex-direction:column;gap:5px}
.fg label{font-size:11px;font-weight:700;color:var(--text2);text-transform:uppercase;letter-spacing:.4px}
.fg input,.fg select,.fg textarea{
  font-size:13px;padding:9px 12px;
  border:1.5px solid var(--border2);border-radius:var(--rs);
  background:var(--surface);color:var(--text);
  transition:border-color .15s,box-shadow .15s;
  width:100%;
}
.fg input:focus,.fg select:focus,.fg textarea:focus{
  outline:none;border-color:var(--blue-mid);
  box-shadow:0 0 0 3px rgba(26,77,143,.12);
}
.fg-full{grid-column:1/-1}
.factions{display:flex;gap:8px;justify-content:flex-end;margin-top:6px;grid-column:1/-1;flex-wrap:wrap}

/* ── BUTTONS ── */
.btn{font-size:12px;font-weight:700;padding:8px 16px;border-radius:var(--rs);cursor:pointer;transition:all .15s;letter-spacing:.2px}
.btn-p{border:none;background:var(--blue);color:#fff}.btn-p:hover{background:var(--blue-mid);transform:translateY(-1px);box-shadow:0 4px 12px rgba(10,46,92,.25)}
.btn-s{border:1.5px solid var(--border2);background:none;color:var(--text2)}.btn-s:hover{background:var(--bg)}
.btn-d{border:1.5px solid var(--red-b);background:none;color:var(--red)}.btn-d:hover{background:var(--red-bg)}
.btn-blue{border:none;background:linear-gradient(135deg,var(--gold-light),var(--gold));color:#fff;font-weight:700}
.btn-blue:hover{transform:translateY(-1px);box-shadow:0 4px 12px rgba(184,134,11,.3)}

/* ── MODAL ── */
.mbg{display:none;position:fixed;inset:0;background:rgba(10,30,60,.5);z-index:300;align-items:center;justify-content:center;padding:1rem;backdrop-filter:blur(3px)}
.mbg.open{display:flex}
.modal{
  background:var(--surface);border-radius:14px;
  border:1px solid var(--border);
  border-top:4px solid var(--blue-mid);
  padding:1.75rem;width:560px;max-width:100%;
  max-height:92vh;overflow-y:auto;
  box-shadow:0 20px 60px rgba(10,46,92,.22);
  animation:modalIn .2s ease;
}
@keyframes modalIn{from{opacity:0;transform:translateY(12px) scale(.97)}to{opacity:1;transform:none}}
.modal-title{font-size:16px;font-weight:800;margin-bottom:1.4rem;color:var(--blue)}
.mactions{display:flex;gap:8px;justify-content:flex-end;margin-top:1.4rem;flex-wrap:wrap}

/* ── ALERT ── */
.alert{padding:11px 15px;border-radius:var(--rs);font-size:12px;line-height:1.7;margin-top:8px}
.al-g{background:var(--green-bg);color:var(--green);border:1px solid var(--green-b);border-left:3px solid var(--green)}
.al-a{background:var(--amber-bg);color:var(--amber);border:1px solid var(--amber-b);border-left:3px solid var(--amber)}
.al-r{background:var(--red-bg);color:var(--red);border:1px solid var(--red-b);border-left:3px solid var(--red)}
.al-b{background:var(--blue-bg);color:var(--blue-mid);border:1px solid var(--blue-b);border-left:3px solid var(--blue-mid)}

/* ── EMPTY STATE ── */
.empty{text-align:center;padding:3rem 1rem;color:var(--text3)}
.empty-icon{font-size:42px;margin-bottom:12px;opacity:.7}
.empty p{font-size:13px;line-height:1.6;margin-bottom:16px}

/* ── PHASE DOT ── */
.pdot{display:inline-block;width:9px;height:9px;border-radius:50%;margin-right:6px;flex-shrink:0}

/* ── IMPORT ZONE ── */
.import-zone{
  border:2px dashed var(--border2);border-radius:var(--r);
  padding:2.5rem 2rem;text-align:center;cursor:pointer;
  transition:all .2s;background:var(--bg);
}
.import-zone:hover,.import-zone.drag{border-color:var(--blue-mid);background:var(--blue-bg)}
.import-zone input[type=file]{display:none}
.import-title{font-size:14px;font-weight:700;color:var(--text);margin-bottom:6px}
.import-sub{font-size:12px;color:var(--text3);margin-bottom:14px}
.import-formats{display:flex;gap:8px;justify-content:center;flex-wrap:wrap;margin-top:12px}
.fmt-badge{font-size:11px;font-weight:700;padding:4px 12px;border-radius:20px;background:var(--blue-bg);color:var(--blue-mid);border:1px solid var(--blue-b)}

/* ── GANTT ── */
.gantt-wrap{overflow-x:auto;-webkit-overflow-scrolling:touch}
.gantt-inner{min-width:560px}
.gantt-head{display:flex;border-bottom:2px solid var(--blue);padding-bottom:8px;margin-bottom:4px}
.gantt-lbl-col{width:160px;flex-shrink:0;font-size:10px;font-weight:700;text-transform:uppercase;color:var(--text3)}
.gantt-months-hd{flex:1;display:flex}
.gantt-month-hd{flex:1;text-align:center;font-size:10px;color:var(--text3);font-weight:600}
.gantt-row{display:flex;align-items:center;padding:4px 0;border-bottom:1px solid var(--border)}
.gantt-row:last-child{border-bottom:none}
.gantt-row-lbl{width:160px;flex-shrink:0;font-size:11px;color:var(--text);white-space:nowrap;overflow:hidden;text-overflow:ellipsis;padding-right:10px;font-weight:500}
.gantt-track{flex:1;position:relative;height:20px;background:var(--bg);border-radius:4px}
.gantt-bar{position:absolute;height:100%;border-radius:4px;display:flex;align-items:center;padding:0 6px;font-size:9px;font-weight:700;color:#fff;overflow:hidden;min-width:3px;transition:opacity .2s}
.gantt-bar:hover{opacity:.85}
.gantt-now{position:absolute;top:-4px;bottom:-4px;width:2px;background:var(--gold-light);z-index:2;border-radius:1px}

/* ── IMPORT PREVIEW ── */
.preview-table{width:100%;border-collapse:collapse;font-size:12px;margin-top:10px}
.preview-table th{background:var(--blue);color:#fff;padding:7px 10px;text-align:left;font-size:10px;font-weight:700;text-transform:uppercase;border-bottom:1px solid var(--border)}
.preview-table td{padding:8px 10px;border-bottom:1px solid var(--border)}
.preview-table tr:last-child td{border-bottom:none}
.preview-table tr:hover td{background:var(--blue-bg)}

/* ── PHASE SECTION HEADER ── */
.ph-section{margin-bottom:1.25rem}
.ph-header{
  display:flex;align-items:center;gap:8px;
  padding:9px 14px;
  background:linear-gradient(135deg,var(--blue) 0%,var(--blue-mid) 100%);
  border-radius:var(--rs);
  margin-bottom:8px;
}
.ph-title{font-size:12px;font-weight:700;color:#fff}
.ph-count{font-size:11px;color:rgba(255,255,255,.6)}

/* ── RESPONSIVE ── */
@media(max-width:768px){
  .main{padding:.875rem 1rem}
  .hdr{padding:.75rem 1rem}
  .tabs{padding:0 .5rem}
  .tab{padding:10px 12px;font-size:11px}
  .two-col{grid-template-columns:1fr}
  .kpis{grid-template-columns:repeat(2,1fr)}
  .card{padding:1rem}
  .ct{font-size:12px}
  .tbl-desktop{display:none}
  .mob-acts{display:block}
}
@media(min-width:769px){
  .mob-acts{display:none}
  .tbl-desktop{display:block}
}
@media(max-width:480px){
  .fgrid{grid-template-columns:1fr}
  .kpis{grid-template-columns:repeat(2,1fr)}
  .hdr-logo{font-size:15px}
  .modal{padding:1.25rem}
  .btn{padding:9px 14px;font-size:11px}
  .factions{justify-content:stretch}
  .factions .btn{flex:1;text-align:center}
}
@media(max-width:360px){
  .kpis{grid-template-columns:1fr}
}
</style>
</head>
<body>

<!-- HEADER -->
<div class="hdr">
  <div class="hdr-logo">
    <div class="hdr-logo-icon">🏗</div>
    Suivi<span> Projet</span>
  </div>
  <div class="hdr-proj">
    <div>
      <div class="proj-name" id="hdr-name">Mon Projet</div>
      <div class="proj-meta" id="hdr-meta">Dates non définies</div>
    </div>
    <button class="btn-cfg" onclick="openProjModal()">⚙ Configurer</button>
  </div>
</div>

<!-- TABS -->
<div class="tabs">
  <button class="tab active" onclick="sw('dashboard',this)">📊 Tableau de bord</button>
  <button class="tab" onclick="sw('activites',this)">📋 Activités</button>
  <button class="tab" onclick="sw('cout',this)">💰 Coût</button>
  <button class="tab" onclick="sw('planning',this)">📅 Planning</button>
</div>

<div class="main">

<!-- ═══ DASHBOARD ═══ -->
<div class="panel active" id="panel-dashboard">
  <div class="card"><div class="ct">Taux d'exécution global</div><div id="d-global"></div></div>
  <div class="kpis" id="d-kpis"></div>
  <div class="two-col">
    <div class="card"><div class="ct">Avancement par phase</div><div id="d-phases"></div></div>
    <div class="card"><div class="ct">Répartition des activités</div><div class="ch-sm"><canvas id="donut-ch"></canvas></div></div>
  </div>
  <div class="card"><div class="ct">Activités en alerte</div><div id="d-alerts"></div></div>
</div>

<!-- ═══ ACTIVITÉS ═══ -->
<div class="panel" id="panel-activites">
  <div class="card">
    <div class="ct">Gestion des activités</div>
    <div id="act-wrap"></div>
  </div>
</div>

<!-- ═══ COÛT ═══ -->
<div class="panel" id="panel-cout">
  <div class="kpis" id="c-kpis"></div>
  <div class="card">
    <div class="ct">Budget prévu &amp; Coût réel par phase
      <button class="btn btn-p" style="font-size:11px" onclick="savePhaseCosts()">💾 Enregistrer</button>
    </div>
    <div id="phase-cost-table"></div>
  </div>
  <div class="card"><div class="ct">Comparaison par phase</div><div class="ch-wrap"><canvas id="cost-ch"></canvas></div></div>
  <div class="card"><div class="ct">Détail par phase</div><div id="cv-list"></div></div>
</div>

<!-- ═══ PLANNING ═══ -->
<div class="panel" id="panel-planning">
  <div class="card"><div class="ct">Diagramme de Gantt</div><div class="gantt-wrap"><div class="gantt-inner" id="gantt"></div></div></div>
  <div class="card"><div class="ct">Échéances à venir</div><div id="echeances"></div></div>
</div>

</div>

<!-- ═══ MODAL PROJET ═══ -->
<div class="mbg" id="m-proj">
  <div class="modal">
    <div class="modal-title">⚙ Configuration du projet</div>
    <div class="fgrid">
      <div class="fg fg-full"><label>Nom du projet *</label><input id="p-name" placeholder="Ex: Construction Lycée de Bouaké"/></div>
      <div class="fg"><label>Date de début</label><input type="date" id="p-start"/></div>
      <div class="fg"><label>Date de fin prévue</label><input type="date" id="p-end"/></div>
      <div class="fg fg-full"><label>Maître d'ouvrage</label><input id="p-moa" placeholder="Ex: Ministère de la Santé"/></div>
      <div class="fg fg-full"><label>Budget global du projet (FCFA)</label><input type="number" id="p-budget" placeholder="Ex: 500000000" min="0"/></div>
      <div class="factions">
        <button class="btn btn-s" onclick="closeModals()">Annuler</button>
        <button class="btn btn-p" onclick="saveProj()">Enregistrer</button>
      </div>
    </div>
  </div>
</div>

<!-- ═══ MODAL IMPORT ═══ -->
<div class="mbg" id="m-import">
  <div class="modal" style="width:620px">
    <div class="modal-title">⬆ Importer depuis MS Project</div>

    <div class="al-b alert" style="margin-bottom:1rem">
      <strong>Formats acceptés :</strong><br>
      • <strong>.xlsx</strong> — Export Excel depuis MS Project (Fichier → Enregistrer sous → Excel)<br>
      • <strong>.csv</strong> — Export CSV depuis MS Project ou tout autre outil de planning<br>
      <strong>Colonnes attendues :</strong> Nom de la tâche · Début · Fin · % Achèvement · Coût prévu · Coût réel
    </div>

    <div class="import-zone" id="drop-zone" onclick="document.getElementById('file-inp').click()"
         ondragover="event.preventDefault();this.classList.add('drag')"
         ondragleave="this.classList.remove('drag')"
         ondrop="handleDrop(event)">
      <input type="file" id="file-inp" accept=".xlsx,.xls,.csv" onchange="handleFile(this.files[0])"/>
      <div style="font-size:32px;margin-bottom:8px">📂</div>
      <div class="import-title">Glissez votre fichier ici ou cliquez pour parcourir</div>
      <div class="import-sub">Fichier exporté depuis MS Project</div>
      <div class="import-formats">
        <span class="fmt-badge">.xlsx</span>
        <span class="fmt-badge">.xls</span>
        <span class="fmt-badge">.csv</span>
      </div>
    </div>

    <div id="import-preview" style="display:none;margin-top:1rem">
      <div class="ct" style="margin-bottom:6px">
        Aperçu — <span id="preview-count"></span> activités détectées
        <button class="btn btn-s" style="font-size:11px" onclick="document.getElementById('import-preview').style.display='none';document.getElementById('file-inp').value=''">✕ Réinitialiser</button>
      </div>
      <div style="overflow-x:auto;max-height:240px;overflow-y:auto">
        <table class="preview-table" id="prev-tbl"></table>
      </div>
      <div style="margin-top:8px" id="import-mapping"></div>
    </div>

    <div class="mactions">
      <button class="btn btn-s" onclick="closeModals()">Annuler</button>
      <button class="btn btn-p" id="btn-confirm-import" style="display:none" onclick="confirmImport()">✓ Importer les activités</button>
    </div>
  </div>
</div>

<!-- ═══ MODAL ACTIVITÉ ═══ -->
<div class="mbg" id="m-act">
  <div class="modal">
    <div class="modal-title" id="act-modal-ttl">+ Nouvelle activité</div>
    <div class="fgrid">
      <div class="fg fg-full">
        <label>Nom de l'activité *</label>
        <input id="a-name" placeholder="Ex: Coulage des fondations bloc A"/>
      </div>

      <div class="fg">
        <label>Phase *</label>
        <select id="a-phase">
          <option value="">— Choisir une phase —</option>
          <optgroup label="Phases BTP standards">
            <option>Installation du chantier</option>
            <option>Terrassement</option>
            <option>Fondations</option>
            <option>Gros œuvre</option>
            <option>Second œuvre</option>
            <option>Équipements</option>
            <option>Réception</option>
          </optgroup>
          <optgroup label="Phases personnalisées" id="custom-phases-grp"></optgroup>
          <option value="__new__">➕ Créer une nouvelle phase…</option>
        </select>
      </div>

      <div class="fg">
        <label>Responsable</label>
        <input id="a-resp" placeholder="Ex: Chef de chantier"/>
      </div>

      <div class="fg" id="new-phase-wrap" style="display:none;grid-column:1/-1">
        <label>Nom de la nouvelle phase</label>
        <input id="a-new-phase" placeholder="Ex: Voirie & Réseaux Divers"/>
      </div>

      <div class="fg">
        <label>Date début *</label>
        <input type="date" id="a-start"/>
      </div>
      <div class="fg">
        <label>Date fin *</label>
        <input type="date" id="a-end"/>
      </div>

      <div class="fg">
        <label>Budget prévu (FCFA)</label>
        <input type="number" id="a-budget" placeholder="0" min="0"/>
      </div>
      <div class="fg">
        <label>Coût réel dépensé (FCFA)</label>
        <input type="number" id="a-reel" placeholder="0" min="0"/>
      </div>

      <div class="fg fg-full">
        <label>Avancement : <strong id="a-pct-lbl">0%</strong></label>
        <input type="range" min="0" max="100" id="a-pct" value="0"
               oninput="document.getElementById('a-pct-lbl').textContent=this.value+'%'"/>
      </div>

      <div class="fg fg-full">
        <label>Observations (optionnel)</label>
        <textarea id="a-obs" rows="2" placeholder="Ex: Retard dû aux pluies, matériaux en attente…" style="resize:vertical"></textarea>
      </div>

      <div class="factions">
        <button class="btn btn-s" onclick="closeModals()">Annuler</button>
        <button class="btn btn-d" id="btn-del-act" style="display:none" onclick="delAct()">Supprimer</button>
        <button class="btn btn-p" onclick="saveAct()">Enregistrer</button>
      </div>
    </div>
  </div>
</div>

<script>
// ═══════════ STATE ═══════════
const BTP_PHASES = ["Installation du chantier","Terrassement","Fondations","Gros œuvre","Second œuvre","Équipements","Réception"];
let customPhases = [];
let phaseBudgets = {}; // { phaseName: budget_prevu }
let phaseReel    = {}; // { phaseName: cout_reel }
let project = { name:"Mon Projet", start:"", end:"", moa:"", budget:0 };
let activities = [];
let editId = null;
let importRows = [];
let donutCh = null, costCh = null;

// ═══════════ HELPERS ═══════════
const uid = () => (Date.now()+Math.random()).toString();

// Jours ouvrables lun-ven entre deux dates
function joursOuvrables(start, end){
  if(!start||!end) return null;
  const s=new Date(start+'T12:00:00'), e=new Date(end+'T12:00:00');
  if(isNaN(s)||isNaN(e)||e<s) return null;
  let count=0, cur=new Date(s);
  while(cur<=e){
    const d=cur.getDay();
    if(d!==0&&d!==6) count++;
    cur.setDate(cur.getDate()+1);
  }
  return count;
}
const fM  = n => n===0?"0":(n/1e6).toFixed(2).replace(/\.?0+$/,"")+" M";
const fmtD = d => { if(!d) return "—"; return new Date(d+'T12:00:00').toLocaleDateString('fr-FR',{day:'2-digit',month:'short',year:'numeric'}); };
const ev  = a => a.budget*(a.pct/100);
const tBac = () => activities.reduce((s,a)=>s+a.budget,0);
const tAC  = () => activities.reduce((s,a)=>s+a.reel,0);
const tEV  = () => activities.reduce((s,a)=>s+ev(a),0);
const gCPI = () => { const t=tAC(); return t>0?tEV()/t:null; };
const eac  = () => { const c=gCPI(); return c&&c>0?tBac()/c:null; };
const gPct = () => { if(!activities.length) return 0; const active=activities.filter(a=>a.pct>0); return Math.round(active.reduce((s,a)=>s+a.pct,0)/activities.length); };
const dLeft= () => { if(!project.end) return null; return Math.ceil((new Date(project.end)-new Date())/86400000); };
const pctC = p => p>=80?"var(--green)":p>=50?"var(--amber)":"var(--red)";
const pctPill=p=>p>=80?"pg":p>=50?"pa":"pr";
const cpiC = c => c>=1?"var(--green)":c>=0.85?"var(--amber)":"var(--red)";

const allPhases = () => [...BTP_PHASES,...customPhases];
const phaseIdx  = ph => { const i=allPhases().indexOf(ph); return i>=0?i:allPhases().length; };
const phColors  = ["#0F4FA8","#2A6308","#7A4800","#5230A0","#9C2828","#1A6A6A","#7A5500","#5A5A00","#2A5A6A"];
const phColor   = ph => phColors[phaseIdx(ph)%phColors.length]||"#555";

const statusOf = a => {
  if(a.pct===100) return {l:"Terminé",c:"pg"};
  const now=new Date(), end=new Date(a.end);
  if(a.pct>0&&end>=now) return {l:"En cours",c:"pa"};
  if(end<now&&a.pct<100) return {l:"En retard",c:"pr"};
  return {l:"Non commencée",c:"px"};
};

// ═══════════ TABS ═══════════
function sw(id,btn){
  document.querySelectorAll('.tab').forEach(t=>t.classList.remove('active'));
  document.querySelectorAll('.panel').forEach(p=>p.classList.remove('active'));
  btn.classList.add('active');
  document.getElementById('panel-'+id).classList.add('active');
  if(id==='dashboard') renderDash();
  if(id==='activites') renderActs();
  if(id==='cout'){ renderCout(); setTimeout(buildCostCh,80); }
  if(id==='planning') renderPlanning();
}

// ═══════════ PROJECT ═══════════
function openProjModal(){
  document.getElementById('p-name').value=project.name;
  document.getElementById('p-start').value=project.start;
  document.getElementById('p-end').value=project.end;
  document.getElementById('p-moa').value=project.moa||'';
  document.getElementById('p-budget').value=project.budget||'';
  document.getElementById('m-proj').classList.add('open');
}
function saveProj(){
  const n=document.getElementById('p-name').value.trim();
  if(!n){alert("Saisissez un nom de projet.");return;}
  project={name:n,start:document.getElementById('p-start').value,end:document.getElementById('p-end').value,moa:document.getElementById('p-moa').value.trim(),budget:parseFloat(document.getElementById('p-budget').value)||0};
  document.getElementById('hdr-name').textContent=n;
  document.getElementById('hdr-meta').textContent=project.start&&project.end?fmtD(project.start)+" → "+fmtD(project.end):project.moa||"Dates non définies";
  closeModals(); renderAll();
}

// ═══════════ ACTIVITY MODAL ═══════════
function openActModal(id=null){
  editId=id;
  const a=id?activities.find(x=>x.id===id):null;
  document.getElementById('act-modal-ttl').textContent=a?"✏ Modifier l'activité":"+ Nouvelle activité";
  document.getElementById('btn-del-act').style.display=a?"inline-block":"none";
  refreshPhaseSelect();
  document.getElementById('a-name').value=a?a.name:'';
  document.getElementById('a-phase').value=a?a.phase:'';
  document.getElementById('a-resp').value=a?a.resp:'';
  document.getElementById('a-start').value=a?a.start:'';
  document.getElementById('a-end').value=a?a.end:'';
  document.getElementById('a-budget').value=a?a.budget:'';
  document.getElementById('a-reel').value=a?a.reel:'';
  document.getElementById('a-pct').value=a?a.pct:0;
  document.getElementById('a-pct-lbl').textContent=(a?a.pct:0)+'%';
  document.getElementById('a-obs').value=a?a.obs||'':'';
  document.getElementById('new-phase-wrap').style.display='none';
  document.getElementById('m-act').classList.add('open');
}

function refreshPhaseSelect(){
  const grp=document.getElementById('custom-phases-grp');
  grp.innerHTML=customPhases.map(p=>`<option value="${p}">${p}</option>`).join('');
}

document.addEventListener('change',function(e){
  if(e.target.id==='a-phase'){
    const v=e.target.value;
    document.getElementById('new-phase-wrap').style.display=v==='__new__'?'grid':'none';
  }
});

function saveAct(){
  let phase=document.getElementById('a-phase').value;
  if(phase==='__new__'){
    const np=document.getElementById('a-new-phase').value.trim();
    if(!np){alert("Saisissez le nom de la nouvelle phase.");return;}
    if(!customPhases.includes(np)) customPhases.push(np);
    phase=np;
  }
  const name=document.getElementById('a-name').value.trim();
  const start=document.getElementById('a-start').value;
  const end=document.getElementById('a-end').value;
  if(!name||!phase||!start||!end){alert("Nom, phase, date début et date fin sont obligatoires.");return;}
  const data={
    name,phase,
    resp:document.getElementById('a-resp').value.trim(),
    start,end,
    budget:parseFloat(document.getElementById('a-budget').value)||0,
    reel:parseFloat(document.getElementById('a-reel').value)||0,
    pct:parseInt(document.getElementById('a-pct').value)||0,
    obs:document.getElementById('a-obs').value.trim(),
  };
  if(editId){const i=activities.findIndex(x=>x.id===editId);activities[i]={...activities[i],...data};}
  else activities.push({id:uid(),...data});
  closeModals(); renderAll();
}
function delAct(){
  if(!confirm("Supprimer cette activité ?"))return;
  activities=activities.filter(a=>a.id!==editId);
  closeModals(); renderAll();
}

// ═══════════ IMPORT ═══════════
function openImport(){
  importRows=[];
  document.getElementById('import-preview').style.display='none';
  document.getElementById('btn-confirm-import').style.display='none';
  document.getElementById('file-inp').value='';
  document.getElementById('m-import').classList.add('open');
}

function handleDrop(e){
  e.preventDefault();
  document.getElementById('drop-zone').classList.remove('drag');
  const f=e.dataTransfer.files[0];
  if(f) handleFile(f);
}

function handleFile(file){
  if(!file) return;
  const ext=file.name.split('.').pop().toLowerCase();
  if(!['xlsx','xls','csv'].includes(ext)){alert("Format non supporté. Utilisez .xlsx, .xls ou .csv");return;}
  const reader=new FileReader();
  reader.onload=function(e){
    let rows=[];
    if(ext==='csv'){
      rows=parseCSV(e.target.result);
    } else {
      const wb=XLSX.read(e.target.result,{type:'binary',cellDates:true});
      const ws=wb.Sheets[wb.SheetNames[0]];
      rows=XLSX.utils.sheet_to_json(ws,{header:1,raw:false,dateNF:'YYYY-MM-DD'});
    }
    processImportRows(rows);
  };
  if(ext==='csv') reader.readAsText(file,'UTF-8');
  else reader.readAsBinaryString(file);
}

function parseCSV(text){
  const sep=text.includes(';')?';':',';
  return text.split('\n').map(l=>l.split(sep).map(c=>c.replace(/^"|"$/g,'').trim())).filter(r=>r.some(c=>c));
}

// Mapping flexible des colonnes MS Project
const COL_MAPS={
  name:['nom','name','task name','tâche','activité','tache','libellé','designation'],
  start:['début','debut','start','date début','date de début','start date'],
  end:['fin','end','finish','date fin','date de fin','end date','finish date'],
  pct:['%','achèvement','achevement','% achèvement','avancement','% achèvement','% complete','complete'],
  budget:['coût','cout','budget','coût prévu','cout prevu','baseline cost','planned cost'],
  reel:['coût réel','cout reel','actual cost','coût réel','real cost'],
  resp:['responsable','resource names','resources','assigné','assigne'],
  phase:['phase','wbs','section','étape','etape']
};

function findCol(headers,key){
  const syns=COL_MAPS[key]||[];
  for(let i=0;i<headers.length;i++){
    const h=(headers[i]||'').toString().toLowerCase().trim();
    if(syns.some(s=>h.includes(s))) return i;
  }
  return -1;
}

function cleanDate(v){
  if(!v) return '';
  const s=v.toString().trim();
  if(/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0,10);
  if(/^\d{2}\/\d{2}\/\d{4}/.test(s)){const[d,m,y]=s.split('/');return `${y}-${m.padStart(2,'0')}-${d.padStart(2,'0')}`;}
  if(/^\d{2}\/\d{2}\/\d{2}$/.test(s)){const[d,m,y]=s.split('/');return `20${y}-${m}-${d}`;}
  const dt=new Date(s);
  if(!isNaN(dt)) return dt.toISOString().slice(0,10);
  return '';
}

function cleanNum(v){
  if(!v&&v!==0) return 0;
  const s=v.toString().replace(/[^\d.,]/g,'').replace(',','.');
  return parseFloat(s)||0;
}

function cleanPct(v){
  const n=cleanNum(v);
  return n>1?Math.min(100,Math.round(n)):Math.min(100,Math.round(n*100));
}

function processImportRows(rows){
  if(!rows||rows.length<2){alert("Fichier vide ou non reconnu.");return;}
  const headers=rows[0].map(h=>(h||'').toString());
  const ci={
    name:findCol(headers,'name'),
    start:findCol(headers,'start'),
    end:findCol(headers,'end'),
    pct:findCol(headers,'pct'),
    budget:findCol(headers,'budget'),
    reel:findCol(headers,'reel'),
    resp:findCol(headers,'resp'),
    phase:findCol(headers,'phase'),
  };

  importRows=[];
  for(let i=1;i<rows.length;i++){
    const r=rows[i];
    if(!r||!r.length) continue;
    const name=ci.name>=0?(r[ci.name]||'').toString().trim():'';
    if(!name||name.toLowerCase()==='tâche'||name.toLowerCase()==='task') continue;
    const start=ci.start>=0?cleanDate(r[ci.start]):'';
    const end=ci.end>=0?cleanDate(r[ci.end]):'';
    if(!start&&!end&&!name) continue;
    importRows.push({
      name,
      phase:ci.phase>=0?(r[ci.phase]||'').toString().trim()||'Gros œuvre':'Gros œuvre',
      resp:ci.resp>=0?(r[ci.resp]||'').toString().trim():'',
      start, end,
      pct:ci.pct>=0?cleanPct(r[ci.pct]):0,
      budget:ci.budget>=0?cleanNum(r[ci.budget]):0,
      reel:ci.reel>=0?cleanNum(r[ci.reel]):0,
      obs:'',
    });
  }

  if(!importRows.length){alert("Aucune activité reconnue. Vérifiez que votre fichier contient des colonnes Nom/Tâche, Début, Fin.");return;}

  document.getElementById('preview-count').textContent=importRows.length;
  document.getElementById('prev-tbl').innerHTML=`
    <thead><tr><th>Activité</th><th>Phase</th><th>Début</th><th>Fin</th><th>Avancement</th><th>Budget</th></tr></thead>
    <tbody>${importRows.slice(0,8).map(r=>`<tr>
      <td>${r.name}</td>
      <td><select onchange="importRows[${importRows.indexOf(r)}].phase=this.value" style="font-size:11px;padding:2px 6px;border:1px solid var(--border);border-radius:4px">
        ${allPhases().map(p=>`<option value="${p}" ${p===r.phase?'selected':''}>${p}</option>`).join('')}
      </select></td>
      <td>${fmtD(r.start)}</td><td>${fmtD(r.end)}</td>
      <td>${r.pct}%</td><td>${r.budget>0?fM(r.budget):"—"}</td>
    </tr>`).join('')}
    ${importRows.length>8?`<tr><td colspan="6" style="text-align:center;color:var(--text3);font-style:italic">… et ${importRows.length-8} autres activités</td></tr>`:''}
    </tbody>`;

  document.getElementById('import-preview').style.display='block';
  document.getElementById('btn-confirm-import').style.display='inline-block';

  const found=Object.entries(ci).filter(([k,v])=>v>=0).map(([k])=>k);
  const missing=Object.entries(ci).filter(([k,v])=>v<0).map(([k])=>k);
  document.getElementById('import-mapping').innerHTML=`
    <div class="alert al-g">✅ Colonnes reconnues : ${found.map(k=>`<strong>${k}</strong>`).join(', ')}
    ${missing.length?`<br>⚠ Non trouvées : ${missing.map(k=>`<em>${k}</em>`).join(', ')} (valeurs par défaut appliquées)`:''}</div>`;
}

function confirmImport(){
  importRows.forEach(r=>{
    if(r.phase&&!BTP_PHASES.includes(r.phase)&&!customPhases.includes(r.phase)) customPhases.push(r.phase);
    activities.push({id:uid(),...r});
  });
  closeModals(); renderAll();
  document.querySelectorAll('.tab')[1].click();
}

function closeModals(){ document.querySelectorAll('.mbg').forEach(m=>m.classList.remove('open')); }

// ═══════════ RENDER ALL ═══════════
function renderAll(){ renderDash(); renderActs(); renderCout(); renderPlanning(); }

// ═══════════ DASHBOARD ═══════════
function renderDash(){
  if(!activities.length){
    document.getElementById('d-kpis').innerHTML='';
    document.getElementById('d-global').innerHTML=`<div class="empty"><div class="empty-icon">🏗</div><div>Aucune activité saisie. Allez dans l'onglet <strong>Activités</strong> pour commencer.</div></div>`;
    document.getElementById('d-phases').innerHTML='';
    document.getElementById('d-alerts').innerHTML='';
    if(donutCh){donutCh.destroy();donutCh=null;}
    return;
  }
  const pct=gPct(),done=activities.filter(a=>a.pct===100).length,
        inp=activities.filter(a=>a.pct>0&&a.pct<100).length,
        pln=activities.filter(a=>a.pct===0).length,
        del=activities.filter(a=>statusOf(a).l==="En retard").length,
        dl=dLeft(), bac=tBac(), ac=tAC(), cpi=gCPI();

  document.getElementById('d-kpis').innerHTML=`
    <div class="kpi"><div class="kl">Exécution</div><div class="kv ${pct>=80?"kv-g":pct>=50?"kv-a":"kv-r"}">${pct}%</div><div class="ks">${done}/${activities.length} terminées</div></div>
    <div class="kpi"><div class="kl">En cours</div><div class="kv kv-b">${inp}</div><div class="ks">activités actives</div></div>
    <div class="kpi"><div class="kl">En retard</div><div class="kv ${del===0?"kv-g":"kv-r"}">${del}</div><div class="ks">activités</div></div>
    ${dl!==null?`<div class="kpi"><div class="kl">Jours restants</div><div class="kv ${dl>60?"kv-g":dl>14?"kv-a":"kv-r"}">${dl}</div><div class="ks">avant fin projet</div></div>`:''}
    ${(()=>{ const bp=allPhases().reduce((s,ph)=>s+(phaseBudgets[ph]||0),0); const cr=allPhases().reduce((s,ph)=>s+(phaseReel[ph]||0),0); const ref=project.budget||bp; return ref>0?`<div class="kpi"><div class="kl">Budget engagé</div><div class="kv kv-a">${Math.round(cr/ref*100)}%</div><div class="ks">${fM(cr)} / ${fM(ref)}</div></div>`:''; })()}  `;

  const gc=pctC(pct);
  document.getElementById('d-global').innerHTML=`
    <div style="font-size:34px;font-weight:700;color:${gc};margin-bottom:8px">${pct}%</div>
    <div class="pbar"><div class="pfill" style="width:${pct}%;background:${gc}"></div></div>
    <div style="display:flex;justify-content:space-between;font-size:11px;color:var(--text3);margin-top:6px">
      <span>${done} terminée${done!==1?"s":""}</span><span>${inp} en cours</span><span>${pln} non commencée${pln!==1?"s":""}</span>
    </div>`;

  const phData=allPhases().map(ph=>{
    const a=activities.filter(x=>x.phase===ph);
    if(!a.length) return null;
    const p=Math.round(a.reduce((s,x)=>s+x.pct,0)/a.length);
    return {ph,p,n:a.length,col:phColor(ph)};
  }).filter(Boolean);
  document.getElementById('d-phases').innerHTML=phData.length?phData.map(d=>`
    <div style="margin-bottom:11px">
      <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:4px">
        <span style="font-size:12px;font-weight:600;display:flex;align-items:center">
          <span class="pdot" style="background:${d.col}"></span>${d.ph}
          <span style="font-size:10px;color:var(--text3);font-weight:400;margin-left:5px">${d.n} acti.</span>
        </span>
        <span style="font-size:12px;font-weight:700;color:${d.col}">${d.p}%</span>
      </div>
      <div class="pbar"><div class="pfill" style="width:${d.p}%;background:${d.col}"></div></div>
    </div>`).join(''):'<div style="color:var(--text3);font-size:12px">—</div>';

  if(donutCh){donutCh.destroy();donutCh=null;}
  const donutTotal=done+inp+pln+del||1;
  donutCh=new Chart(document.getElementById('donut-ch'),{
    type:'doughnut',
    data:{labels:["Terminées","En cours","Non commencées","En retard"],
      datasets:[{data:[done,inp,pln,del],backgroundColor:["#2A6308","#7A4800","#9A9890","#9C2828"],borderWidth:2,borderColor:'#fff'}]},
    options:{responsive:true,maintainAspectRatio:false,
      plugins:{
        legend:{position:'bottom',labels:{font:{size:11},boxWidth:12,padding:10}},
        tooltip:{callbacks:{label:c=>{const v=c.parsed;const p=Math.round(v/donutTotal*100);return ` ${c.label} : ${v} (${p}%)`;}}},
      },
      layout:{padding:10}
    },
    plugins:[{
      id:'pctLabels',
      afterDatasetDraw(chart){
        const {ctx,data}=chart;
        const ds=chart.getDatasetMeta(0);
        const total=data.datasets[0].data.reduce((a,b)=>a+b,0)||1;
        ds.data.forEach((arc,i)=>{
          const val=data.datasets[0].data[i];
          if(!val) return;
          const pct=Math.round(val/total*100);
          if(pct<5) return;
          const cx=arc.x, cy=arc.y;
          const mid=(arc.startAngle+arc.endAngle)/2;
          const r=(arc.outerRadius+arc.innerRadius)/2;
          const x=cx+Math.cos(mid)*r;
          const y=cy+Math.sin(mid)*r;
          ctx.save();
          ctx.fillStyle='#fff';
          ctx.font='bold 11px Segoe UI,system-ui,sans-serif';
          ctx.textAlign='center';
          ctx.textBaseline='middle';
          ctx.fillText(pct+'%',x,y);
          ctx.restore();
        });
      }
    }]
  });

  const risk=activities.filter(a=>statusOf(a).l==="En retard"||(a.budget>0&&a.reel>a.budget));
  document.getElementById('d-alerts').innerHTML=risk.length?risk.map(a=>{
    const s=statusOf(a),over=a.budget>0&&a.reel>a.budget;
    return `<div style="display:flex;align-items:center;gap:10px;padding:9px 0;border-bottom:1px solid var(--border)">
      <span class="pill ${s.c}">${s.l}</span>
      <div style="flex:1">
        <div style="font-size:13px;font-weight:500">${a.name}</div>
        <div style="font-size:11px;color:var(--text3)">${a.phase} · Fin : ${fmtD(a.end)}${over?" · ⚠ Dépassement budget":""}</div>
      </div>
      <span style="font-size:12px;font-weight:700;color:${pctC(a.pct)}">${a.pct}%</span>
    </div>`;
  }).join(''):`<div class="alert al-g">✅ Aucune activité en alerte.</div>`;
}

// ═══════════ ACTIVITÉS ═══════════
function renderActs(){
  if(!activities.length){
    document.getElementById('act-wrap').innerHTML=`
      <div class="empty"><div class="empty-icon">🏗</div>
      <p>Aucune activité. Saisissez manuellement ou importez depuis MS Project.</p>
      <div style="display:flex;gap:8px;justify-content:center;flex-wrap:wrap;margin-top:12px">
        <button class="btn btn-blue" onclick="openImport()">⬆ Importer MS Project</button>
        <button class="btn btn-p" onclick="openActModal()">+ Saisie manuelle</button>
      </div></div>`;
    return;
  }
  let html='';
  const usedPhases=allPhases().filter(ph=>activities.some(a=>a.phase===ph));
  usedPhases.forEach(ph=>{
    const acts=activities.filter(a=>a.phase===ph);
    const phPct=Math.round(acts.reduce((s,a)=>s+a.pct,0)/acts.length);
    html+=`<div class="ph-section">
      <div class="ph-header">
        <span class="pdot" style="background:${phColor(ph)}"></span>
        <span class="ph-title">${ph}</span>
        <span class="ph-count">${acts.length} activité${acts.length!==1?"s":""}</span>
        <span class="pill ${pctPill(phPct)}" style="margin-left:auto">${phPct}%</span>
      </div>

      <!-- VERSION DESKTOP -->
      <div class="tbl-desktop">
        <div class="tbl-wrap">
        <table class="tbl"><thead><tr>
          <th>Activité</th><th>Début</th><th>Fin</th>
          <th title="Jours ouvrables">Durée</th>
          <th>Budget prévu</th><th>Coût réel</th>
          <th>Responsable</th><th>Avancement</th><th>Statut</th><th></th>
        </tr></thead><tbody>
        ${acts.map(a=>{
          const s=statusOf(a),c=pctC(a.pct),over=a.budget>0&&a.reel>a.budget;
          const jo=joursOuvrables(a.start,a.end);
          const idx=activities.indexOf(a);
          return `<tr>
            <td><div style="font-weight:600">${a.name}</div>${a.obs?`<div style="font-size:10px;color:var(--text3);margin-top:2px;font-style:italic">${a.obs}</div>`:''}</td>
            <td style="color:var(--text2)">${fmtD(a.start)}</td>
            <td style="color:var(--text2)">${fmtD(a.end)}</td>
            <td style="text-align:center;font-weight:700;color:var(--blue)">${jo!==null?jo+"j":"—"}</td>
            <td>${a.budget>0?fM(a.budget):"—"}</td>
            <td style="color:${over?"var(--red)":"inherit"};font-weight:${over?"700":"400"}">${a.reel>0?fM(a.reel):"—"}${over?" ⚠":""}</td>
            <td style="color:var(--text2)">${a.resp||'—'}</td>
            <td style="font-size:13px;font-weight:800;color:${c}">${a.pct}%</td>
            <td><span class="pill ${s.c}">${s.l}</span></td>
            <td><button class="btn-ico" onclick="openActModal(activities[${idx}].id.toString())">✏</button></td>
          </tr>`;
        }).join('')}
        </tbody></table>
        </div>
      </div>

      <!-- VERSION MOBILE -->
      <div class="mob-acts">
        ${acts.map(a=>{
          const s=statusOf(a),c=pctC(a.pct),over=a.budget>0&&a.reel>a.budget;
          const idx=activities.indexOf(a);
          return `<div class="mob-act-card">
            <div style="display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:8px">
              <div class="mob-act-name">${a.name}</div>
              <button class="btn-ico" onclick="openActModal(activities[${idx}].id.toString())" style="margin-left:8px;flex-shrink:0">✏</button>
            </div>
            <div style="display:flex;gap:6px;flex-wrap:wrap;margin-bottom:8px">
              <span class="pill ${s.c}">${s.l}</span>
              <span style="font-size:11px;font-weight:800;color:${c}">${a.pct}%</span>
            </div>
            <div class="mob-act-row"><span>📅 Fin</span><strong>${fmtD(a.end)}</strong></div>
            ${a.budget>0?`<div class="mob-act-row"><span>💰 Budget prévu</span><strong>${fM(a.budget)}</strong></div>`:''}
            ${a.reel>0?`<div class="mob-act-row"><span>💸 Coût réel</span><strong style="color:${over?"var(--red)":"var(--text2)"}">${fM(a.reel)}${over?" ⚠":""}</strong></div>`:''}
            ${a.resp?`<div class="mob-act-row"><span>👤 Responsable</span><strong>${a.resp}</strong></div>`:''}
            ${a.obs?`<div style="font-size:11px;color:var(--text3);font-style:italic;margin-top:6px">${a.obs}</div>`:''}
          </div>`;
        }).join('')}
      </div>
    </div>`;
  });
  document.getElementById('act-wrap').innerHTML=html+`
    <div style="display:flex;gap:8px;justify-content:flex-end;flex-wrap:wrap;margin-top:8px">
      <button class="btn btn-blue" onclick="openImport()">⬆ Importer MS Project</button>
      <button class="btn btn-p" onclick="openActModal()">+ Saisie manuelle</button>
    </div>`;
}

// ═══════════ COÛT ═══════════
function renderCout(){
  const usedPhases = allPhases().filter(ph => activities.some(a=>a.phase===ph));

  // KPIs globaux
  const totalBudgetPhases = usedPhases.reduce((s,ph)=>s+(phaseBudgets[ph]||0),0);
  const totalReelPhases   = usedPhases.reduce((s,ph)=>s+(phaseReel[ph]||0),0);
  const budgetRef = project.budget || totalBudgetPhases;
  const ecart = totalBudgetPhases - totalReelPhases;

  document.getElementById('c-kpis').innerHTML=`
    ${project.budget>0?`<div class="kpi"><div class="kl">Budget global projet</div><div class="kv kv-b">${fM(project.budget)}</div><div class="ks">FCFA alloués</div></div>`:''}
    <div class="kpi"><div class="kl">Total budget prévu phases</div><div class="kv kv-b">${fM(totalBudgetPhases)}</div><div class="ks">somme phases</div></div>
    <div class="kpi"><div class="kl">Total coût réel phases</div><div class="kv kv-a">${fM(totalReelPhases)}</div><div class="ks">${totalBudgetPhases>0?Math.round(totalReelPhases/totalBudgetPhases*100):0}% consommé</div></div>
    ${project.budget>0?`<div class="kpi"><div class="kl">% Budget utilisé</div><div class="kv ${totalReelPhases/project.budget<=0.85?"kv-g":totalReelPhases/project.budget<=1?"kv-a":"kv-r"}">${Math.round(totalReelPhases/project.budget*100)}%</div><div class="ks">${fM(totalReelPhases)} / ${fM(project.budget)}</div></div>`:''}
    <div class="kpi"><div class="kl">Écart global</div><div class="kv ${ecart>=0?"kv-g":"kv-r"}">${ecart>=0?"+":""}${fM(ecart)}</div><div class="ks">${ecart>=0?"✅ Économie":"❌ Dépassement"}</div></div>
    ${project.budget>0&&totalBudgetPhases>0?`<div class="kpi"><div class="kl">Budget phases / global</div><div class="kv ${totalBudgetPhases<=project.budget?"kv-g":"kv-r"}">${Math.round(totalBudgetPhases/project.budget*100)}%</div><div class="ks">${totalBudgetPhases<=project.budget?"✅ OK":"⚠ Dépasse budget"}</div></div>`:''}
  `;

  // Tableau de saisie par phase
  if(!usedPhases.length){
    document.getElementById('phase-cost-table').innerHTML='<div class="empty"><div class="empty-icon">💰</div><div>Ajoutez des activités avec des phases pour saisir les budgets.</div></div>';
    document.getElementById('cv-list').innerHTML='';
    return;
  }

  document.getElementById('phase-cost-table').innerHTML=`
    <table class="tbl">
      <thead><tr>
        <th>Phase</th>
        <th>Activités</th>
        <th>Budget prévu (FCFA)</th>
        <th>Coût réel (FCFA)</th>
        <th>Écart</th>
        <th>Avancement</th>
      </tr></thead>
      <tbody>
      ${usedPhases.map(ph=>{
        const acts=activities.filter(a=>a.phase===ph);
        const bp=phaseBudgets[ph]||0;
        const cr=phaseReel[ph]||0;
        const ec=bp-cr;
        const phPct=Math.round(acts.reduce((s,a)=>s+a.pct,0)/acts.length);
        const ecC=ec>=0?"var(--green)":ec>-5e6?"var(--amber)":"var(--red)";
        return `<tr>
          <td><span class="pdot" style="background:${phColor(ph)}"></span><strong>${ph}</strong></td>
          <td style="color:var(--text3);font-size:11px">${acts.length} activité${acts.length>1?"s":""}</td>
          <td><input type="number" min="0" placeholder="0"
            value="${bp||''}"
            id="bp-${ph.replace(/\s+/g,'_')}"
            style="width:100%;font-size:12px;padding:5px 8px;border:1px solid var(--border2);border-radius:6px;background:var(--bg)"
            onchange="phaseBudgets['${ph}']=parseFloat(this.value)||0"/></td>
          <td><input type="number" min="0" placeholder="0"
            value="${cr||''}"
            id="cr-${ph.replace(/\s+/g,'_')}"
            style="width:100%;font-size:12px;padding:5px 8px;border:1px solid var(--border2);border-radius:6px;background:var(--bg)"
            onchange="phaseReel['${ph}']=parseFloat(this.value)||0"/></td>
          <td style="font-weight:700;color:${ecC}">${bp>0?(ec>=0?"+":"")+fM(ec):"—"}</td>
          <td>
            <div style="display:flex;align-items:center;gap:6px">
              <div style="flex:1"><div class="pbar-sm"><div class="pfill" style="width:${phPct}%;background:${phColor(ph)}"></div></div></div>
              <span style="font-size:11px;font-weight:700;color:${phColor(ph)}">${phPct}%</span>
            </div>
          </td>
        </tr>`;
      }).join('')}
      </tbody>
      <tfoot>
        <tr style="background:var(--bg)">
          <td colspan="2" style="font-weight:700;font-size:12px;padding:10px">TOTAL</td>
          <td style="font-weight:700;font-size:12px;padding:10px;color:var(--blue)">${fM(totalBudgetPhases)}</td>
          <td style="font-weight:700;font-size:12px;padding:10px;color:var(--amber)">${fM(totalReelPhases)}</td>
          <td style="font-weight:700;font-size:12px;padding:10px;color:${ecart>=0?"var(--green)":"var(--red)"}">${ecart>=0?"+":""}${fM(ecart)}</td>
          <td></td>
        </tr>
      </tfoot>
    </table>`;

  // Détail par phase
  document.getElementById('cv-list').innerHTML=usedPhases.map(ph=>{
    const bp=phaseBudgets[ph]||0, cr=phaseReel[ph]||0;
    const ec=bp-cr, pct=bp>0?Math.min(100,Math.round(cr/bp*100)):0;
    const c=ec>=0?"var(--green)":ec>-5e6?"var(--amber)":"var(--red)";
    return `<div style="padding:10px 0;border-bottom:1px solid var(--border)">
      <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:5px">
        <span style="font-size:12px;font-weight:600;display:flex;align-items:center;gap:5px">
          <span class="pdot" style="background:${phColor(ph)}"></span>${ph}
        </span>
        <span style="font-size:12px;font-weight:700;color:${c}">${bp>0?(ec>=0?"+":"")+fM(ec):"—"}</span>
      </div>
      ${bp>0?`
      <div class="pbar" style="margin-bottom:4px"><div class="pfill" style="width:${pct}%;background:${c}"></div></div>
      <div style="display:flex;justify-content:space-between;font-size:10px;color:var(--text3)">
        <span>Prévu : ${fM(bp)}</span><span>Réel : ${fM(cr)}</span><span>${pct}% consommé</span>
      </div>`:'<div style="font-size:11px;color:var(--text3)">Budget non défini</div>'}
    </div>`;
  }).join('');

  setTimeout(buildCostCh, 80);
}

function savePhaseCosts(){
  // Les valeurs sont déjà mises à jour via onchange, on re-render juste
  renderCout();
}

function buildCostCh(){
  const el=document.getElementById('cost-ch');if(!el)return;
  if(costCh){costCh.destroy();costCh=null;}
  const usedPhases=allPhases().filter(ph=>activities.some(a=>a.phase===ph));
  const bpData=usedPhases.map(ph=>+((phaseBudgets[ph]||0)/1e6).toFixed(2));
  const crData=usedPhases.map(ph=>+((phaseReel[ph]||0)/1e6).toFixed(2));
  if(!bpData.some(v=>v>0)&&!crData.some(v=>v>0)) return;
  costCh=new Chart(el,{type:'bar',
    data:{
      labels:usedPhases.map(ph=>ph.length>14?ph.slice(0,14)+"…":ph),
      datasets:[
        {label:"Budget prévu",data:bpData,backgroundColor:"rgba(15,79,168,.2)",borderColor:"#0F4FA8",borderWidth:1.5,borderRadius:4,barThickness:18},
        {label:"Coût réel",data:crData,backgroundColor:usedPhases.map((ph,i)=>crData[i]>bpData[i]?"rgba(156,40,40,.72)":"rgba(42,99,8,.72)"),borderRadius:4,barThickness:18},
      ]},
    options:{responsive:true,maintainAspectRatio:false,
      plugins:{legend:{position:'top',labels:{font:{size:11},boxWidth:10}},tooltip:{callbacks:{label:c=>c.dataset.label+": "+c.parsed.y+" M FCFA"}}},
      scales:{y:{ticks:{callback:v=>v+" M",font:{size:10}},grid:{color:"rgba(0,0,0,.05)"}},x:{ticks:{font:{size:10}},grid:{display:false}}}}
  });
}

// ═══════════ PLANNING ═══════════
function renderPlanning(){
  if(!activities.length){
    document.getElementById('gantt').innerHTML='<div class="empty"><div class="empty-icon">📅</div><div>Aucune activité à afficher.</div></div>';
    document.getElementById('echeances').innerHTML='';
    return;
  }
  const dts=activities.map(a=>new Date(a.start)).filter(d=>!isNaN(d));
  const dte=activities.map(a=>new Date(a.end)).filter(d=>!isNaN(d));
  if(!dts.length){document.getElementById('gantt').innerHTML='<div style="color:var(--text3);font-size:12px">Dates manquantes sur certaines activités.</div>';return;}
  const minD=new Date(Math.min(...dts));minD.setDate(1);
  const maxD=new Date(Math.max(...dte));maxD.setMonth(maxD.getMonth()+1);maxD.setDate(0);
  const months=[],cur=new Date(minD);
  while(cur<=maxD){months.push(new Date(cur));cur.setMonth(cur.getMonth()+1);}
  const ML=['Jan','Fév','Mar','Avr','Mai','Jun','Jul','Aoû','Sep','Oct','Nov','Déc'];
  const now=new Date();
  const todayPct=Math.max(0,Math.min(100,((now-minD)/(maxD-minD))*100));
  const heads=months.map(m=>`<div class="gantt-month-hd">${ML[m.getMonth()]} ${m.getFullYear().toString().slice(2)}</div>`).join('');
  const rows=activities.map(a=>{
    const s=new Date(a.start),e=new Date(a.end);
    if(isNaN(s)||isNaN(e)) return '';
    const left=Math.max(0,((s-minD)/(maxD-minD))*100);
    const width=Math.max(.5,((e-s)/(maxD-minD))*100);
    const col=a.pct===100?"var(--green)":phColor(a.phase);
    return `<div class="gantt-row">
      <div class="gantt-row-lbl" title="${a.name}">${a.name}</div>
      <div class="gantt-track">
        <div class="gantt-now" style="left:${todayPct}%" title="Aujourd'hui"></div>
        <div class="gantt-bar" style="left:${left}%;width:${width}%;background:${col}" title="${a.name} — ${a.pct}%">${width>7?a.pct+'%':''}</div>
      </div>
    </div>`;
  }).join('');
  document.getElementById('gantt').innerHTML=`
    <div class="gantt-head"><div class="gantt-lbl-col">Activité</div><div class="gantt-months-hd">${heads}</div></div>
    ${rows}
    <div style="font-size:10px;color:var(--text3);margin-top:8px;display:flex;align-items:center;gap:6px"><div style="width:2px;height:12px;background:var(--red);border-radius:1px"></div>Ligne rouge = aujourd'hui</div>`;

  const upcoming=[...activities].filter(a=>a.pct<100).sort((a,b)=>new Date(a.end)-new Date(b.end)).slice(0,6);
  document.getElementById('echeances').innerHTML=upcoming.length?upcoming.map(a=>{
    const d=Math.ceil((new Date(a.end)-now)/86400000),s=statusOf(a);
    const c=d<0?"var(--red)":d<14?"var(--amber)":"var(--green)";
    return `<div style="display:flex;align-items:center;gap:10px;padding:9px 0;border-bottom:1px solid var(--border)">
      <span class="pill ${s.c}">${s.l}</span>
      <div style="flex:1"><div style="font-size:13px;font-weight:500">${a.name}</div>
      <div style="font-size:11px;color:var(--text3)">${a.phase} · Fin : ${fmtD(a.end)}</div></div>
      <div style="text-align:right;font-size:12px;font-weight:700;color:${c}">${d<0?"Dépassé "+Math.abs(d)+"j":d===0?"Aujourd'hui":d+" j restants"}</div>
    </div>`;
  }).join(''):`<div class="alert al-g">✅ Toutes les activités sont à jour.</div>`;
}

// INIT
renderDash(); renderActs();
</script>
</body>
</html>
