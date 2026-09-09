// H_RSM-CRM | worker.js - Portfoy, Talep, Emsal (MHT + Excel Okuyucu)
const HTML_CONTENT = `<!DOCTYPE html>
<html lang="tr">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>H_RSM-CRM | Akıllı Emlak & Emsal Analiz</title>
  <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
  <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
  <style>
    :root { --p: #4f46e5; --bg: #0f172a; --card: #1e293b; --text: #f8fafc; --muted: #94a3b8; --border: #334155; --success: #10b981; }
    * { box-sizing: border-box; margin: 0; padding: 0; }
    body { font-family: system-ui, -apple-system, sans-serif; background: var(--bg); color: var(--text); min-height: 100vh; }
    .app { display: grid; grid-template-columns: 240px 1fr; min-height: 100vh; }
    .side { background: #090d16; border-right: 1px solid var(--border); padding: 20px; position: sticky; top: 0; height: 100vh; }
    .brand { display: flex; align-items: center; gap: 12px; margin-bottom: 28px; }
    .logo { width: 40px; height: 40px; background: var(--p); border-radius: 10px; display: grid; place-items: center; font-weight: 800; font-size: 18px; }
    .nav { display: flex; flex-direction: column; gap: 8px; }
    .nav button { border: 0; background: transparent; color: var(--muted); padding: 12px 14px; border-radius: 8px; text-align: left; font-size: 14px; cursor: pointer; display: flex; align-items: center; gap: 10px; width: 100%; }
    .nav button:hover, .nav button.active { background: rgba(79, 70, 229, 0.15); color: #fff; font-weight: 600; }
    main { padding: 24px; max-width: 1200px; width: 100%; margin: 0 auto; }
    .top-bar { display: flex; justify-content: space-between; align-items: center; margin-bottom: 24px; flex-wrap: wrap; gap: 12px; }
    .grid-stats { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 16px; margin-bottom: 24px; }
    .stat-card { background: var(--card); border: 1px solid var(--border); padding: 18px; border-radius: 12px; }
    .stat-card b { font-size: 26px; display: block; margin-top: 6px; color: #fff; }
    .btn { background: var(--p); color: #fff; border: 0; padding: 10px 16px; border-radius: 8px; font-weight: 600; cursor: pointer; display: inline-flex; align-items: center; gap: 8px; }
    .btn-danger { background: #ef4444; }
    .btn-sec { background: #334155; }
    .btn-green { background: #059669; }
    .table-container { background: var(--card); border: 1px solid var(--border); border-radius: 12px; overflow-x: auto; margin-bottom: 24px; }
    table { width: 100%; border-collapse: collapse; text-align: left; font-size: 14px; min-width: 650px; }
    th, td { padding: 14px 18px; border-bottom: 1px solid var(--border); }
    th { color: var(--muted); font-size: 12px; text-transform: uppercase; }
    .badge { padding: 4px 8px; border-radius: 6px; font-size: 12px; font-weight: 600; background: rgba(79, 70, 229, 0.2); color: #818cf8; }
    .badge-match { background: rgba(16, 185, 129, 0.2); color: #34d399; }
    .badge-excel { background: rgba(5, 150, 105, 0.2); color: #6ee7b7; }
    .modal { position: fixed; inset: 0; background: rgba(0,0,0,0.7); display: none; place-items: center; padding: 16px; z-index: 99; backdrop-filter: blur(4px); }
    .modal.open { display: grid; }
    .modal-box { background: var(--card); border: 1px solid var(--border); border-radius: 14px; width: 100%; max-width: 520px; padding: 22px; max-height: 90vh; overflow-y: auto; }
    .form-group { display: flex; flex-direction: column; gap: 6px; margin-bottom: 14px; }
    input, select, textarea { background: #0f172a; border: 1px solid var(--border); padding: 10px; border-radius: 8px; color: #fff; font: inherit; }
    .upload-zone { border: 2px dashed var(--border); padding: 24px; border-radius: 12px; text-align: center; cursor: pointer; margin-bottom: 14px; background: rgba(15,23,42,0.5); }
    .upload-zone:hover { border-color: var(--p); }
    .tab-content { display: none; }
    .tab-content.active { display: block; }
    @media (max-width: 768px) { .app { grid-template-columns: 1fr; } .side { display: none; } }
  </style>
</head>
<body>
  <div class="app">
    <aside class="side">
      <div class="brand">
        <div class="logo">H</div>
        <div><b>H_RSM-CRM</b><p style="font-size: 11px; color: var(--muted);">Emlak & Emsal Paneli</p></div>
      </div>
      <div class="nav">
        <button id="nav-portfoy" class="active" onclick="switchTab('portfoy')"><i class="fa-solid fa-building"></i> Portföyler</button>
        <button id="nav-talep" onclick="switchTab('talep')"><i class="fa-solid fa-user-tag"></i> Alıcı Talepleri</button>
        <button id="nav-emsal" onclick="switchTab('emsal')"><i class="fa-solid fa-file-invoice"></i> Emsal Havuzu (MHT / Excel)</button>
      </div>
    </aside>

    <main>
      <div class="top-bar">
        <h2 id="viewTitle">Portföyler</h2>
        <div style="display:flex; gap:10px; flex-wrap:wrap;">
          <button class="btn btn-green" onclick="openModal('excelModal')"><i class="fa-solid fa-file-excel"></i> Excel Yükle</button>
          <button class="btn btn-sec" onclick="openModal('mhtModal')"><i class="fa-solid fa-file-arrow-up"></i> Sahibinden MHT</button>
          <button class="btn" id="btnAddAction" onclick="handleNewAction()"><i class="fa-solid fa-plus"></i> Yeni Ekle</button>
        </div>
      </div>

      <div class="grid-stats">
        <div class="stat-card"><span>Aktif Portföy</span><b id="pCount">0</b></div>
        <div class="stat-card"><span>Alıcı Talebi</span><b id="tCount">0</b></div>
        <div class="stat-card"><span>Emsal Havuzu</span><b id="eCount">0</b></div>
      </div>

      <section id="tab-portfoy" class="tab-content active">
        <div class="table-container">
          <table>
            <thead><tr><th>Başlık</th><th>Tür</th><th>Konum</th><th>Fiyat</th><th>Müşteri / İletişim</th><th>İşlem</th></tr></thead>
            <tbody id="rows-portfoy"></tbody>
          </table>
        </div>
      </section>

      <section id="tab-talep" class="tab-content">
        <div class="table-container">
          <table>
            <thead><tr><th>Müşteri</th><th>Aranan Konum</th><th>Tür / Oda</th><th>Bütçe Max</th><th>Eşleşen İlanlar</th><th>İşlem</th></tr></thead>
            <tbody id="rows-talep"></tbody>
          </table>
        </div>
      </section>

      <section id="tab-emsal" class="tab-content">
        <div class="table-container">
          <table>
            <thead><tr><th>İlan Başlığı</th><th>Konum / Mahalle</th><th>Fiyat</th><th>Metrekare</th><th>Oda</th><th>Kaynak</th></tr></thead>
            <tbody id="rows-emsal"></tbody>
          </table>
        </div>
      </section>
    </main>
  </div>

  <div id="itemModal" class="modal">
    <div class="modal-box">
      <h3>Yeni Portföy Ekle</h3>
      <form onsubmit="savePortfoy(event)" style="margin-top:14px;">
        <div class="form-group"><label>Başlık</label><input id="p_title" required placeholder="Örn: 3+1 Cadde Üstü"></div>
        <div class="form-group"><label>Tür</label><select id="p_type"><option>Satılık</option><option>Kiralık</option></select></div>
        <div class="form-group"><label>Konum</label><input id="p_location" required placeholder="Konya / Selçuklu / Yazır"></div>
        <div class="form-group"><label>Fiyat (TL)</label><input type="number" id="p_price" required placeholder="4500000"></div>
        <div class="form-group"><label>Mülk Sahibi & Tel</label><input id="p_client" required placeholder="Ahmet Bey - 0532..."></div>
        <div style="display:flex; justify-content:flex-end; gap:8px; margin-top:16px;">
          <button type="button" class="btn btn-sec" onclick="closeModal('itemModal')">İptal</button>
          <button type="submit" class="btn">Kaydet</button>
        </div>
      </form>
    </div>
  </div>

  <div id="talepModal" class="modal">
    <div class="modal-box">
      <h3>Yeni Alıcı Talebi</h3>
      <form onsubmit="saveTalep(event)" style="margin-top:14px;">
        <div class="form-group"><label>Müşteri Adı & Tel</label><input id="t_client" required placeholder="Mesut Hoca - 0505..."></div>
        <div class="form-group"><label>İstenen Konum</label><input id="t_location" required placeholder="Selçuklu"></div>
        <div class="form-group"><label>İstenen Tip / Oda</label><input id="t_req" required placeholder="3+1 Satılık Daire"></div>
        <div class="form-group"><label>Maksimum Bütçe (TL)</label><input type="number" id="t_budget" required placeholder="5000000"></div>
        <div style="display:flex; justify-content:flex-end; gap:8px; margin-top:16px;">
          <button type="button" class="btn btn-sec" onclick="closeModal('talepModal')">İptal</button>
          <button type="submit" class="btn">Talebi Aç</button>
        </div>
      </form>
    </div>
  </div>

  <div id="mhtModal" class="modal">
    <div class="modal-box">
      <h3>Sahibinden .MHT Dosyası Yükle</h3>
      <p style="font-size:12px; color:var(--muted); margin:8px 0 16px;">Sahibinden ilan dosyasını (.mht veya .html) seçin.</p>
      <div class="upload-zone" onclick="document.getElementById('mhtFileInput').click()">
        <i class="fa-solid fa-file-code" style="font-size:32px; color:var(--p); margin-bottom:8px;"></i>
        <p>MHT Dosyası Seç</p>
        <span id="uploadFileName" style="font-size:12px; color:var(--success);"></span>
      </div>
      <input type="file" id="mhtFileInput" accept=".mht,.mhtml,.html" style="display:none;" onchange="handleFileSelect(event)">
      <div style="display:flex; justify-content:flex-end;">
        <button type="button" class="btn btn-sec" onclick="closeModal('mhtModal')">Kapat</button>
      </div>
    </div>
  </div>

  <div id="excelModal" class="modal">
    <div class="modal-box">
      <h3>Excel Dosyası Yükle (.xlsx / .xls)</h3>
      <p style="font-size:12px; color:var(--muted); margin:8px 0 16px;">Başlık, Fiyat ve Konum içeren Excel tablonuzu emsal havuzuna tek seferde aktarın.</p>
      <div class="upload-zone" onclick="document.getElementById('excelFileInput').click()">
        <i class="fa-solid fa-file-excel" style="font-size:32px; color:#10b981; margin-bottom:8px;"></i>
        <p>Excel Tablosu Seç</p>
        <span id="excelFileName" style="font-size:12px; color:var(--success);"></span>
      </div>
      <input type="file" id="excelFileInput" accept=".xlsx,.xls,.csv" style="display:none;" onchange="handleExcelSelect(event)">
      <div style="display:flex; justify-content:flex-end;">
        <button type="button" class="btn btn-sec" onclick="closeModal('excelModal')">Kapat</button>
      </div>
    </div>
  </div>

  <div id="matchModal" class="modal">
    <div class="modal-box" style="max-width:650px;">
      <h3>Eşleşen Uygun İlanlar & Emsaller</h3>
      <div id="matchResults" style="margin-top:16px;"></div>
      <div style="display:flex; justify-content:flex-end; margin-top:16px;">
        <button class="btn btn-sec" onclick="closeModal('matchModal')">Kapat</button>
      </div>
    </div>
  </div>

  <script>
    let activeTab = 'portfoy';
    let portfoyData = [], talepData = [], emsalData = [];

    function switchTab(tab) {
      activeTab = tab;
      document.querySelectorAll('.tab-content').forEach(el => el.classList.remove('active'));
      document.querySelectorAll('.nav button').forEach(el => el.classList.remove('active'));
      document.getElementById('tab-' + tab).classList.add('active');
      document.getElementById('nav-' + tab).classList.add('active');
      document.getElementById('viewTitle').innerText = tab === 'portfoy' ? 'Portföyler' : (tab === 'talep' ? 'Alıcı Talepleri' : 'Emsal Havuzu');
    }

    function handleNewAction() {
      if (activeTab === 'portfoy') openModal('itemModal');
      else if (activeTab === 'talep') openModal('talepModal');
      else openModal('mhtModal');
    }

    function openModal(id) { document.getElementById(id).classList.add('open'); }
    function closeModal(id) { document.getElementById(id).classList.remove('open'); }

    async function loadAll() {
      const res = await fetch('/api/all');
      const data = await res.json();
      portfoyData = data.items || [];
      talepData = data.talepler || [];
      emsalData = data.emsaller || [];

      document.getElementById('pCount').innerText = portfoyData.length;
      document.getElementById('tCount').innerText = talepData.length;
      document.getElementById('eCount').innerText = emsalData.length;

      renderPortfoy();
      renderTalep();
      renderEmsal();
    }

    function renderPortfoy() {
      document.getElementById('rows-portfoy').innerHTML = portfoyData.map(i => \`
        <tr>
          <td><b>\${i.title}</b></td>
          <td><span class="badge">\${i.type}</span></td>
          <td>\${i.location}</td>
          <td style="color:#34d399; font-weight:600;">\${Number(i.price).toLocaleString('tr-TR')} ₺</td>
          <td>\${i.client}</td>
          <td><button onclick="delRecord('items', \${i.id})" class="btn btn-danger" style="padding:6px 10px; font-size:12px;"><i class="fa-solid fa-trash"></i></button></td>
        </tr>
      \`).join('') || '<tr><td colspan="6" style="text-align:center; padding:18px; color:var(--muted);">Kayıt bulunamadı.</td></tr>';
    }

    function renderTalep() {
      document.getElementById('rows-talep').innerHTML = talepData.map(t => {
        const matches = [...portfoyData, ...emsalData].filter(item => 
          (item.location && item.location.toLowerCase().includes(t.location.toLowerCase())) &&
          (Number(item.price) <= Number(t.budget))
        );
        return \`
          <tr>
            <td><b>\${t.client}</b></td>
            <td>\${t.location}</td>
            <td><span class="badge">\${t.requirement}</span></td>
            <td style="color:#34d399; font-weight:600;">\${Number(t.budget).toLocaleString('tr-TR')} ₺</td>
            <td><button class="btn btn-sec" style="padding:4px 10px; font-size:12px;" onclick='showMatches(\${JSON.stringify(matches)})'>\${matches.length} Eşleşme Gör</button></td>
            <td><button onclick="delRecord('talepler', \${t.id})" class="btn btn-danger" style="padding:6px 10px; font-size:12px;"><i class="fa-solid fa-trash"></i></button></td>
          </tr>
        \`;
      }).join('') || '<tr><td colspan="6" style="text-align:center; padding:18px; color:var(--muted);">Kayıtlı talep yok.</td></tr>';
    }

    function renderEmsal() {
      document.getElementById('rows-emsal').innerHTML = emsalData.map(e => \`
        <tr>
          <td><b>\${e.title}</b></td>
          <td>\${e.location}</td>
          <td style="color:#34d399; font-weight:600;">\${Number(e.price).toLocaleString('tr-TR')} ₺</td>
          <td>\${e.m2 || '-'} m²</td>
          <td>\${e.rooms || '-'}</td>
          <td><span class="badge \${e.source === 'Excel' ? 'badge-excel' : 'badge-match'}">\${e.source || 'Emsal'}</span></td>
        </tr>
      \`).join('') || '<tr><td colspan="6" style="text-align:center; padding:18px; color:var(--muted);">Henüz emsal yüklenmedi.</td></tr>';
    }

    function showMatches(list) {
      const box = document.getElementById('matchResults');
      if(!list.length) {
        box.innerHTML = '<p style="color:var(--muted); text-align:center; padding:12px;">Bu talebe uyan ilan veya emsal bulunamadı.</p>';
      } else {
        box.innerHTML = list.map(m => \`
          <div style="background:#0f172a; border:1px solid var(--border); padding:12px; border-radius:8px; margin-bottom:8px; display:flex; justify-content:space-between; align-items:center;">
            <div>
              <b>\${m.title}</b>
              <p style="font-size:12px; color:var(--muted);">\${m.location} | \${m.client ? 'Müşteri: ' + m.client : (m.source || 'Emsal Kaydı')}</p>
            </div>
            <b style="color:#34d399;">\${Number(m.price).toLocaleString('tr-TR')} ₺</b>
          </div>
        \`).join('');
      }
      openModal('matchModal');
    }

    async function savePortfoy(e) {
      e.preventDefault();
      await fetch('/api/add-portfoy', {
        method: 'POST',
        body: JSON.stringify({
          title: document.getElementById('p_title').value,
          type: document.getElementById('p_type').value,
          location: document.getElementById('p_location').value,
          price: Number(document.getElementById('p_price').value),
          client: document.getElementById('p_client').value
        })
      });
      closeModal('itemModal');
      loadAll();
    }

    async function saveTalep(e) {
      e.preventDefault();
      await fetch('/api/add-talep', {
        method: 'POST',
        body: JSON.stringify({
          client: document.getElementById('t_client').value,
          location: document.getElementById('t_location').value,
          requirement: document.getElementById('t_req').value,
          budget: Number(document.getElementById('t_budget').value)
        })
      });
      closeModal('talepModal');
      loadAll();
    }

    async function handleFileSelect(event) {
      const file = event.target.files[0];
      if (!file) return;
      document.getElementById('uploadFileName').innerText = "Ayrıştırılıyor: " + file.name;
      const reader = new FileReader();
      reader.onload = async function(e) {
        const text = e.target.result;
        const titleMatch = text.match(/<title>([^<]+)<\/title>/i) || text.match(/classifiedDetailTitle[^>]*>([^<]+)/i);
        const priceMatch = text.match(/([0-9]{1,3}(?:\.[0-9]{3})+)\s*(?:TL|tl)/);
        const locMatch = text.match(/classifiedInfo[^>]*>[\s\S]*?([A-Za-zğüşıöçĞÜŞİÖÇ\s\/]+)<\/h2>/i);

        const title = titleMatch ? titleMatch[1].replace(/[\r\n\t]/g, '').trim() : file.name.replace(/\.[^/.]+$/, "");
        const price = priceMatch ? Number(priceMatch[1].replace(/\./g, '')) : 0;
        const location = locMatch ? locMatch[1].trim() : "Konya";

        await fetch('/api/add-emsal', {
          method: 'POST',
          body: JSON.stringify({ title, price, location, m2: 0, rooms: 'Belirtilmedi', source: 'Sahibinden MHT' })
        });

        alert("MHT İlanı Emsal Havuzuna Başarıyla Eklendi!");
        closeModal('mhtModal');
        loadAll();
      };
      reader.readAsText(file);
    }

    async function handleExcelSelect(event) {
      const file = event.target.files[0];
      if (!file) return;
      document.getElementById('excelFileName').innerText = "Yükleniyor: " + file.name;
      const reader = new FileReader();
      reader.onload = async function(e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[firstSheetName];
        const rows = XLSX.utils.sheet_to_json(sheet);

        let count = 0;
        for (const r of rows) {
          const title = r['Başlık'] || r['İlan Başlığı'] || r['Baslik'] || r['Adı'] || 'Excel İlanı';
          const price = Number(r['Fiyat'] || r['Fiyat (TL)'] || r['Tutar'] || 0);
          const location = r['Konum'] || r['İlçe'] || r['Mahalle'] || r['Bölge'] || 'Konya';
          const m2 = Number(r['m2'] || r['Metrekare'] || r['M2'] || 0);
          const rooms = r['Oda'] || r['Oda Sayısı'] || 'Belirtilmedi';

          if (price > 0 || title) {
            await fetch('/api/add-emsal', {
              method: 'POST',
              body: JSON.stringify({ title, price, location, m2, rooms, source: 'Excel' })
            });
            count++;
          }
        }

        alert(count + " adet kayıt Excel'den Emsal Havuzuna başarıyla aktarıldı!");
        closeModal('excelModal');
        loadAll();
      };
      reader.readAsArrayBuffer(file);
    }

    async function delRecord(tbl, id) {
      if(confirm('Silmek istediğinize emin misiniz?')) {
        await fetch(\`/api/del?table=\${tbl}&id=\${id}\`, { method: 'DELETE' });
        loadAll();
      }
    }

    loadAll();
  </script>
</body>
</html>`;

async function initDB(db) {
  await db.prepare("CREATE TABLE IF NOT EXISTS items(id INTEGER PRIMARY KEY AUTOINCREMENT, title TEXT, type TEXT, location TEXT, client TEXT, price REAL, created_at TEXT DEFAULT CURRENT_TIMESTAMP)").run();
  await db.prepare("CREATE TABLE IF NOT EXISTS talepler(id INTEGER PRIMARY KEY AUTOINCREMENT, client TEXT, location TEXT, requirement TEXT, budget REAL, created_at TEXT DEFAULT CURRENT_TIMESTAMP)").run();
  await db.prepare("CREATE TABLE IF NOT EXISTS emsaller(id INTEGER PRIMARY KEY AUTOINCREMENT, title TEXT, price REAL, location TEXT, m2 INTEGER, rooms TEXT, source TEXT, created_at TEXT DEFAULT CURRENT_TIMESTAMP)").run();
}

export default {
  async fetch(req, env) {
    if (!env.DB) return new Response("DB eksik", { status: 500 });
    await initDB(env.DB);
    const url = new URL(req.url);

    if (url.pathname === "/") {
      return new Response(HTML_CONTENT, { headers: { "content-type": "text/html; charset=utf-8" } });
    }

    if (url.pathname === "/api/all") {
      const items = (await env.DB.prepare("SELECT * FROM items ORDER BY id DESC").all()).results || [];
      const talepler = (await env.DB.prepare("SELECT * FROM talepler ORDER BY id DESC").all()).results || [];
      const emsaller = (await env.DB.prepare("SELECT * FROM emsaller ORDER BY id DESC").all()).results || [];
      return Response.json({ items, talepler, emsaller });
    }

    if (url.pathname === "/api/add-portfoy" && req.method === "POST") {
      const b = await req.json();
      await env.DB.prepare("INSERT INTO items(title, type, location, client, price) VALUES(?,?,?,?,?)").bind(b.title, b.type, b.location, b.client, b.price).run();
      return Response.json({ ok: true });
    }

    if (url.pathname === "/api/add-talep" && req.method === "POST") {
      const b = await req.json();
      await env.DB.prepare("INSERT INTO talepler(client, location, requirement, budget) VALUES(?,?,?,?)").bind(b.client, b.location, b.requirement, b.budget).run();
      return Response.json({ ok: true });
    }

    if (url.pathname === "/api/add-emsal" && req.method === "POST") {
      const b = await req.json();
      await env.DB.prepare("INSERT INTO emsaller(title, price, location, m2, rooms, source) VALUES(?,?,?,?,?,?)").bind(b.title, b.price, b.location, b.m2, b.rooms, b.source || 'Emsal').run();
      return Response.json({ ok: true });
    }

    if (url.pathname === "/api/del" && req.method === "DELETE") {
      const id = url.searchParams.get("id");
      const table = url.searchParams.get("table");
      if (['items', 'talepler', 'emsaller'].includes(table)) {
        await env.DB.prepare(`DELETE FROM ${table} WHERE id=?`).bind(id).run();
      }
      return Response.json({ ok: true });
    }

    return new Response("Not found", { status: 404 });
  }
};
