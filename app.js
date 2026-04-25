// ── CONFIG ──────────────────────────────────────────────────────────────────
// After converting your .xlsx to a Google Sheet, paste the Sheet ID here.
// The ID is in the URL: docs.google.com/spreadsheets/d/SHEET_ID/edit
const SHEET_ID = '1txSQJVw7nA42BtS8ilDyrM4H0nYtGWvP3mb4tcZJ034';

// Sheet tab names (must match exactly)
const TABS = {
  orders:       'Orders',
  customers:    'Customers',
  pilots:       'Pilots',
  products:     'Products',
  templates:    'MixTemplates',
  templateProds:'MixTemplateProducts',
  orderProds:   'OrderProducts',
};

// Google Apps Script Web App URL (set up in Step 2 of setup guide)
const GAS_URL = 'https://script.google.com/macros/s/AKfycby-9LUgvhXKSi0Cj-8NRextt-pGjYvodK7YOQHCE35xmrmEKBLEE5FG8QXQOiICCkfUEA/exec';

// ── STATE ───────────────────────────────────────────────────────────────────
let DB = {
  orders: [], customers: [], pilots: [],
  products: [], templates: [], templateProds: [], orderProds: [],
  fields: [], orderFields: []
};

// App settings — persisted to localStorage
let AppSettings = {
  defaultSprayRate: 2,   // gal/ac — used when template has no spray rate
  defaultTankSize:  100, // gallons
};

function loadSettings() {
  try {
    const s = localStorage.getItem('blueraven_settings');
    if (s) AppSettings = { ...AppSettings, ...JSON.parse(s) };
  } catch(e) {}
}

function saveSettings() {
  try { localStorage.setItem('blueraven_settings', JSON.stringify(AppSettings)); } catch(e) {}
}
let currentView = 'dashboard';
let previousView = 'dashboard';
let currentOrderId = null;
let selectedCalcOrders = new Set();

// ── THEME ────────────────────────────────────────────────────────────────────
function toggleTheme() {
  const isLight = document.body.classList.toggle('light');
  localStorage.setItem('blueraven_theme', isLight ? 'light' : 'dark');
  updateThemeButton(isLight);
}

function updateThemeButton(isLight) {
  const label = document.getElementById('themeLabel');
  const icon  = document.getElementById('themeIcon');
  if (!label || !icon) return;
  if (isLight) {
    label.textContent = 'Dark Mode';
    icon.innerHTML = '<path d="M21 12.79A9 9 0 1 1 11.21 3 7 7 0 0 0 21 12.79z"/>';
  } else {
    label.textContent = 'Light Mode';
    icon.innerHTML = '<circle cx="12" cy="12" r="5"/><path d="M12 1v2M12 21v2M4.22 4.22l1.42 1.42M18.36 18.36l1.42 1.42M1 12h2M21 12h2M4.22 19.78l1.42-1.42M18.36 5.64l1.42-1.42"/>';
  }
}

function loadTheme() {
  const saved = localStorage.getItem('blueraven_theme');
  const isLight = saved === 'light';
  if (isLight) document.body.classList.add('light');
  updateThemeButton(isLight);
}

// ── INIT ────────────────────────────────────────────────────────────────────
document.addEventListener('DOMContentLoaded', () => {
  loadTheme();
  document.getElementById('dashDate').textContent =
    new Date().toLocaleDateString('en-US', { weekday:'long', year:'numeric', month:'long', day:'numeric' });

  // Nav clicks
  document.querySelectorAll('.nav-link').forEach(link => {
    link.addEventListener('click', e => {
      e.preventDefault();
      navigateTo(link.dataset.view);
    });
  });

  loadFromLocalStorage();
  syncData();
});

// ── NAVIGATION ──────────────────────────────────────────────────────────────
function navigateTo(view) {
  previousView = currentView;
  currentView = view;
  document.querySelectorAll('.view').forEach(v => v.classList.remove('active'));
  document.querySelectorAll('.nav-link').forEach(l => l.classList.remove('active'));
  const el = document.getElementById('view-' + view);
  if (el) el.classList.add('active');
  const link = document.querySelector(`[data-view="${view}"]`);
  if (link) link.classList.add('active');
  renderView(view);
  toggleSidebar(false);
}

function goBack() {
  navigateTo(previousView === 'order-detail' ? 'orders' : previousView);
}

function renderView(view) {
  switch(view) {
    case 'dashboard':  renderDashboard(); break;
    case 'orders':     renderOrdersList(); break;
    case 'customers':  renderCustomers(); break;
    case 'products':   renderProducts(); break;
    case 'fields':     renderFields(); break;
    case 'calculator': renderCalculator(); break;
    case 'reports':    renderReports(); break;
    case 'gdu':        renderGDU(); break;
  }
}

// ── DATA SYNC ───────────────────────────────────────────────────────────────
async function syncData() {
  const btn = document.getElementById('syncBtn');
  const status = document.getElementById('syncStatus');
  btn.classList.add('syncing');
  status.className = 'sync-status';
  status.textContent = 'Syncing...';

  if (GAS_URL === 'YOUR_APPS_SCRIPT_URL_HERE' || SHEET_ID === 'YOUR_SHEET_ID_HERE') {
    btn.classList.remove('syncing');
    status.className = 'sync-status err';
    status.textContent = 'Setup required';
    showToast('See setup guide to connect Google Sheets', 'error');
    renderView(currentView);
    return;
  }

  try {
    const res = await fetch(`${GAS_URL}?action=read`);
    const data = await res.json();
    if (data.error) throw new Error(data.error);

    DB.orders       = parseSheet(data.orders,       orderHeaders()).map(o => ({
      ...o,
      OrderDate:     toDateStr(o.OrderDate),
      PlantingDate:  toDateStr(o.PlantingDate),
      ScheduledDate: toDateStr(o.ScheduledDate),
      CompletedDate: toDateStr(o.CompletedDate),
    }));
    DB.customers    = parseSheet(data.customers,     customerHeaders());
    DB.pilots       = parseSheet(data.pilots,        pilotHeaders());
    DB.products     = parseSheet(data.products,      productHeaders());
    DB.templates    = parseSheet(data.templates,     templateHeaders());
    DB.templateProds= parseSheet(data.templateProds, templateProdHeaders());
    DB.orderProds   = parseSheet(data.orderProds,    orderProdHeaders());
    DB.fields       = parseSheet(data.fields,        fieldHeaders());
    DB.orderFields  = parseSheet(data.orderFields,   orderFieldHeaders());

    saveToLocalStorage();
    status.className = 'sync-status ok';
    status.textContent = 'Synced ' + new Date().toLocaleTimeString();
    showToast('Data synced', 'success');
  } catch(e) {
    status.className = 'sync-status err';
    status.textContent = 'Sync failed';
    showToast('Sync failed: ' + e.message, 'error');
  }

  btn.classList.remove('syncing');
  renderView(currentView);
}

function parseSheet(rows, expectedHeaders) {
  if (!rows || rows.length < 2) return [];
  // Use ACTUAL sheet header row for column mapping — never assume position
  const actualHeaders = rows[0].map(h => String(h).trim());
  const idCol = actualHeaders[0] || expectedHeaders[0];
  return rows.slice(1).map(row => {
    const obj = {};
    // Map by actual sheet column name
    actualHeaders.forEach((h, i) => { if (h) obj[h] = row[i] !== undefined ? String(row[i]) : ''; });
    // Also fill any expected headers that are missing from sheet (default empty)
    expectedHeaders.forEach(h => { if (!(h in obj)) obj[h] = ''; });
    return obj;
  }).filter(r => r[idCol]);
}

// Header definitions match sheet columns exactly
function orderHeaders()       { return ['OrderID','OrderDate','CustomerID','CustomerName','CropType','PlantingDate','RelativeMaturity','ScheduledDate','CompletedDate','PilotID','PilotName','Status','PricingType','RatePerAcre','TotalAcres','EstimatedTotal','ChemicalCost','TemplateUsed','Invoiced','DJI_FlightFile','Attachments','Notes']; }
function customerHeaders()    { return ['CustomerID','Name','Phone','Email','Address','City','State','Zip','Notes']; }
function pilotHeaders()       { return ['PilotID','Name','Phone','Email','FAA_Part107_Num','Active']; }
function productHeaders()     { return ['ProductID','ProductName','Manufacturer','Unit','CostPerUnit','REI_Hours','PHI_Days','Notes']; }
function templateHeaders()    { return ['TemplateID','TemplateName','CropType','SprayRate','Description','Active']; }
function templateProdHeaders(){ return ['LineID','TemplateID','ProductID','ProductName','RatePerAcre','Unit','SuppliedBy','Notes']; }
function orderProdHeaders()   { return ['LineID','OrderID','ProductID','ProductName','RatePerAcre','Unit','SuppliedBy','CostPerUnit','Acres','TotalUnitsNeeded','TotalProductCost']; }
function fieldHeaders()       { return ['FieldID','CustomerID','CustomerName','FieldName','Acres','CentroidLat','CentroidLng','PolygonKML','CLU_TractID','CLU_FarmNum','Active','Notes']; }
function orderFieldHeaders()  { return ['LineID','OrderID','FieldID','FieldName','CustomerID','Acres','Notes']; }

// ── LOCAL STORAGE ───────────────────────────────────────────────────────────
function saveToLocalStorage() {
  try { localStorage.setItem('blueraven_db', JSON.stringify(DB)); } catch(e) {}
}
function loadFromLocalStorage() {
  loadSettings();
  try {
    const d = localStorage.getItem('blueraven_db');
    if (d) { DB = JSON.parse(d); renderView(currentView); }
  } catch(e) {}
}

// ── DASHBOARD ───────────────────────────────────────────────────────────────
function renderDashboard() {
  const now = new Date();
  const thisMonth = now.getMonth();
  const thisYear  = now.getFullYear();

  const open      = DB.orders.filter(o => o.Status === 'Open').length;
  const scheduled = DB.orders.filter(o => o.Status === 'Scheduled').length;
  const completed = DB.orders.filter(o => {
    if (o.Status !== 'Completed') return false;
    const d = new Date(o.CompletedDate);
    return d.getMonth() === thisMonth && d.getFullYear() === thisYear;
  }).length;
  const uninvoiced = DB.orders.filter(o => o.Status === 'Completed' && o.Invoiced === 'No').length;

  document.getElementById('statOpen').textContent      = open;
  document.getElementById('statScheduled').textContent = scheduled;
  document.getElementById('statCompleted').textContent = completed;
  document.getElementById('statUninvoiced').textContent= uninvoiced;

  // Upcoming: open + scheduled sorted by scheduled date
  const upcoming = DB.orders
    .filter(o => o.Status !== 'Completed')
    .sort((a, b) => new Date(a.ScheduledDate) - new Date(b.ScheduledDate))
    .slice(0, 8);

  document.getElementById('dashOrders').innerHTML =
    upcoming.length ? upcoming.map(orderCardHTML).join('') :
    '<div class="empty-state">No upcoming orders</div>';
}

// ── ORDERS LIST ─────────────────────────────────────────────────────────────
function renderOrdersList() {
  filterOrders();
}

function filterOrders() {
  const search   = (document.getElementById('orderSearch')?.value || '').toLowerCase();
  const status   = document.getElementById('filterStatus')?.value || '';
  const invoiced = document.getElementById('filterInvoiced')?.value || '';

  const filtered = DB.orders.filter(o => {
    const matchSearch = !search ||
      o.CustomerName.toLowerCase().includes(search) ||
      DB.orderFields.filter(f=>f.OrderID===o.OrderID).some(f=>f.FieldName.toLowerCase().includes(search)) ||
      o.OrderID.toLowerCase().includes(search);
    const matchStatus   = !status   || o.Status === status;
    const matchInvoiced = !invoiced || o.Invoiced === invoiced;
    return matchSearch && matchStatus && matchInvoiced;
  }).sort((a, b) => new Date(b.OrderDate) - new Date(a.OrderDate));

  document.getElementById('ordersList').innerHTML =
    filtered.length ? filtered.map(orderCardHTML).join('') :
    '<div class="empty-state">No orders match your filters</div>';
}

function orderCardHTML(o) {
  const badge = statusBadge(o.Status);
  const invoicedBadge = o.Invoiced === 'Yes' ? '<span class="badge badge-invoiced">Invoiced</span>' : '';
  const acres = o.TotalAcres ? `${parseFloat(o.TotalAcres).toLocaleString()} ac` : '';
  const date  = o.ScheduledDate ? fmtDate(o.ScheduledDate) : '';
  const fieldNames = DB.orderFields.filter(f => f.OrderID === o.OrderID).map(f => f.FieldName).join(', ') || '—';
  return `
  <div class="order-card" onclick="viewOrder('${o.OrderID}')">
    <span class="order-id">${o.OrderID}</span>
    <div class="order-main">
      <div class="order-customer">${o.CustomerName}</div>
      <div class="order-field">${fieldNames} · ${o.CropType}</div>
    </div>
    <div class="order-meta">
      <div>${acres}</div>
      <div>${date}</div>
    </div>
    <div style="display:flex;flex-direction:column;gap:0.3rem;align-items:flex-end">
      ${badge}${invoicedBadge}
    </div>
  </div>`;
}

function statusBadge(status) {
  const cls = { Open:'badge-open', Scheduled:'badge-scheduled', Completed:'badge-completed' }[status] || 'badge-open';
  return `<span class="badge ${cls}">${status}</span>`;
}

// ── ORDER DETAIL ─────────────────────────────────────────────────────────────
function viewOrder(orderId) {
  currentOrderId = orderId;
  const o = DB.orders.find(x => x.OrderID === orderId);
  if (!o) return;

  const prods = DB.orderProds.filter(p => p.OrderID === orderId);
  const fields = DB.orderFields.filter(f => f.OrderID === orderId);
  const fieldNames = fields.map(f => f.FieldName).join(', ') || '—';
  const estTotal = o.EstimatedTotal ? '$' + parseFloat(o.EstimatedTotal).toLocaleString('en-US',{minimumFractionDigits:2}) : '—';
  const chemCost = o.ChemicalCost   ? '$' + parseFloat(o.ChemicalCost).toLocaleString('en-US',{minimumFractionDigits:2}) : '—';

  let chemTable = '';
  if (prods.length) {
    chemTable = `<table class="chem-table">
      <thead><tr><th>Product</th><th>Rate/Ac</th><th>Unit</th><th>Supplied By</th><th>Total Qty</th><th>Cost</th></tr></thead>
      <tbody>${prods.map(p => `<tr>
        <td>${p.ProductName}</td><td>${p.RatePerAcre}</td><td>${p.Unit}</td>
        <td>${p.SuppliedBy}</td>
        <td>${parseFloat(p.TotalUnitsNeeded||0).toFixed(1)} ${p.Unit}</td>
        <td>${p.TotalProductCost ? '$'+parseFloat(p.TotalProductCost).toFixed(2) : '—'}</td>
      </tr>`).join('')}</tbody>
    </table>`;
  } else {
    chemTable = '<div style="color:var(--text-sub);font-size:0.85rem">No chemicals on this order</div>';
  }

  const btnComplete  = document.getElementById('btnMarkComplete');
  const btnInvoiced  = document.getElementById('btnMarkInvoiced');
  if (btnComplete) btnComplete.style.display  = o.Status !== 'Completed' ? 'inline-flex' : 'none';
  if (btnInvoiced) btnInvoiced.style.display  = o.Invoiced !== 'Yes'     ? 'inline-flex' : 'none';

  document.getElementById('orderDetailContent').innerHTML = `
    <div class="detail-grid">
      <div class="detail-card">
        <div class="detail-section-title">Job Info</div>
        <div class="detail-row"><span>Order</span><strong>${o.OrderID}</strong></div>
        <div class="detail-row"><span>Customer</span><span>${o.CustomerName}</span></div>
        <div class="detail-row"><span>Status</span>${statusBadge(o.Status)}</div>
        <div class="detail-row"><span>Crop</span><span>${o.CropType || '—'}</span></div>
        ${o.PlantingDate ? `<div class="detail-row"><span>Planted</span><span>${fmtDate(o.PlantingDate)}</span></div>` : ''}
        ${o.RelativeMaturity ? `<div class="detail-row"><span>RM</span><span>${o.RelativeMaturity}</span></div>` : ''}
        ${o.TemplateUsed ? `<div class="detail-row"><span>Mix Template</span><span>${DB.templates.find(t=>t.TemplateID===o.TemplateUsed)?.TemplateName || o.TemplateUsed}</span></div>` : ''}
        <div class="detail-row"><span>Scheduled</span><span>${o.ScheduledDate ? fmtDate(o.ScheduledDate) : '—'}</span></div>
        ${o.CompletedDate ? `<div class="detail-row"><span>Completed</span><span>${fmtDate(o.CompletedDate)}</span></div>` : ''}
        <div class="detail-row"><span>Pilot</span><span>${o.PilotName || '—'}</span></div>
      </div>
      <div class="detail-card">
        <div class="detail-section-title">Fields</div>
        ${fields.length ? fields.map(f => `
          <div class="detail-row"><span>${f.FieldName}</span><span>${f.Acres ? parseFloat(f.Acres).toFixed(1)+' ac' : '—'}</span></div>
        `).join('') : '<div style="color:var(--text-sub);font-size:0.85rem">No fields linked</div>'}
        ${o.TotalAcres ? `<div class="detail-row" style="border-top:1px solid var(--border);margin-top:0.5rem;padding-top:0.5rem"><span><strong>Total</strong></span><strong>${parseFloat(o.TotalAcres).toFixed(1)} ac</strong></div>` : ''}
      </div>
      <div class="detail-card">
        <div class="detail-section-title">Pricing</div>
        <div class="detail-row"><span>Type</span><span>${o.PricingType || '—'}</span></div>
        <div class="detail-row"><span>Rate/Acre</span><span>${o.RatePerAcre ? '$'+parseFloat(o.RatePerAcre).toFixed(2) : '—'}</span></div>
        <div class="detail-row"><span>Chem Cost</span><span>${chemCost}</span></div>
        <div class="detail-row"><span>Est. Total</span><strong>${estTotal}</strong></div>
        <div class="detail-row"><span>Invoiced</span><span>${o.Invoiced || 'No'}</span></div>
      </div>
    </div>
    <div class="detail-card" style="margin-top:1rem">
      <div class="detail-section-title">Chemical Mix</div>
      ${chemTable}
    </div>
    ${o.Notes ? `<div class="detail-card" style="margin-top:1rem"><div class="detail-section-title">Notes</div><p style="color:var(--text-sub);font-size:0.85rem">${o.Notes}</p></div>` : ''}
  `;

  navigateTo('order-detail');
}


function renderCustomers() {
  filterCustomers();
}

function filterCustomers() {
  const search = (document.getElementById('customerSearch')?.value || '').toLowerCase();
  const filtered = DB.customers.filter(c =>
    !search || c.Name.toLowerCase().includes(search) || c.Email.toLowerCase().includes(search)
  );

  document.getElementById('customersList').innerHTML = filtered.length ?
    filtered.map(c => {
      const orders = DB.orders.filter(o => o.CustomerID === c.CustomerID);
      const totalAcres = orders.reduce((s, o) => s + parseFloat(o.TotalAcres||0), 0);
      return `<div class="customer-card" onclick="editCustomer('${c.CustomerID}')">
        <div class="customer-name">${c.Name}</div>
        <div class="customer-info">
          ${c.Phone ? `📞 ${c.Phone}<br>` : ''}
          ${c.Email ? `✉️ ${c.Email}<br>` : ''}
          ${c.City ? `📍 ${c.City}, ${c.State}` : ''}
        </div>
        <div class="customer-stat">${orders.length} orders · ${totalAcres.toLocaleString()} ac total</div>
        <div style="font-size:0.72rem;color:var(--text-sub);margin-top:0.4rem">tap to edit</div>
      </div>`;
    }).join('') :
    '<div class="empty-state">No customers found</div>';
}

// ── CUSTOMER MODAL ────────────────────────────────────────────────────────────
function showNewCustomerModal() {
  document.getElementById('editCustomerId').value = '';
  document.getElementById('customerModalTitle').textContent = 'New Customer';
  ['cName','cPhone','cEmail','cAddress','cCity','cState','cZip','cNotes'].forEach(id => {
    document.getElementById(id).value = '';
  });
  document.getElementById('btnDeleteCustomer').style.display = 'none';
  openModal('customerModal');
}

function editCustomer(customerId) {
  const c = DB.customers.find(x => x.CustomerID === customerId);
  if (!c) return;
  document.getElementById('editCustomerId').value = c.CustomerID;
  document.getElementById('customerModalTitle').textContent = 'Edit Customer';
  document.getElementById('cName').value    = c.Name;
  document.getElementById('cPhone').value   = c.Phone;
  document.getElementById('cEmail').value   = c.Email;
  document.getElementById('cAddress').value = c.Address;
  document.getElementById('cCity').value    = c.City;
  document.getElementById('cState').value   = c.State;
  document.getElementById('cZip').value     = c.Zip;
  document.getElementById('cNotes').value   = c.Notes;
  document.getElementById('btnDeleteCustomer').style.display = '';
  openModal('customerModal');
}

async function saveCustomer() {
  const name = document.getElementById('cName').value.trim();
  if (!name) { showToast('Customer name is required', 'error'); return; }

  const editId = document.getElementById('editCustomerId').value;
  const customer = {
    CustomerID: editId || nextId('CUST', DB.customers.map(c => c.CustomerID)),
    Name:    name,
    Phone:   document.getElementById('cPhone').value.trim(),
    Email:   document.getElementById('cEmail').value.trim(),
    Address: document.getElementById('cAddress').value.trim(),
    City:    document.getElementById('cCity').value.trim(),
    State:   document.getElementById('cState').value.trim().toUpperCase(),
    Zip:     document.getElementById('cZip').value.trim(),
    Notes:   document.getElementById('cNotes').value.trim(),
  };

  if (editId) {
    const idx = DB.customers.findIndex(c => c.CustomerID === editId);
    if (idx > -1) DB.customers[idx] = customer;
  } else {
    DB.customers.push(customer);
  }

  saveToLocalStorage();
  closeModal();
  showToast(editId ? 'Customer updated' : 'Customer added', 'success');
  await writeRow('customers', customer);
  renderCustomers();
}

async function deleteCustomer() {
  const editId = document.getElementById('editCustomerId').value;
  if (!editId) return;
  const inUse = DB.orders.some(o => o.CustomerID === editId);
  if (inUse) { showToast('Cannot delete — customer has existing orders', 'error'); return; }
  if (!confirm('Delete this customer? This cannot be undone.')) return;
  DB.customers = DB.customers.filter(c => c.CustomerID !== editId);
  saveToLocalStorage();
  closeModal();
  showToast('Customer deleted', 'success');
  await writeRow('customers', { CustomerID: editId, _delete: true });
  renderCustomers();
}

// ── PRODUCTS & MIXES ─────────────────────────────────────────────────────────
function renderProducts() {
  // Products table — rows are clickable to edit
  document.getElementById('productsList').innerHTML = DB.products.length ?
    `<table class="product-table">
      <thead><tr><th>Product</th><th>Manufacturer</th><th>Unit</th><th>Cost/Unit</th><th>REI (hrs)</th><th>PHI (days)</th><th>Notes</th></tr></thead>
      <tbody>${DB.products.map(p => `
        <tr onclick="editProduct('${p.ProductID}')" style="cursor:pointer" title="Click to edit">
          <td><strong>${p.ProductName}</strong></td>
          <td>${p.Manufacturer}</td>
          <td class="mono">${p.Unit}</td>
          <td class="mono">$${parseFloat(p.CostPerUnit||0).toFixed(2)}</td>
          <td class="mono">${p.REI_Hours}</td>
          <td class="mono">${p.PHI_Days}</td>
          <td style="color:var(--text-sub)">${p.Notes}</td>
        </tr>`).join('')}
      </tbody>
    </table>` :
    '<div class="empty-state">No products loaded</div>';

  // Mix templates — header is clickable to edit
  document.getElementById('mixesList').innerHTML = DB.templates.length ?
    DB.templates.map(t => {
      const prods = DB.templateProds.filter(p => p.TemplateID === t.TemplateID);
      return `<div class="mix-card">
        <div class="mix-header" onclick="editMix('${t.TemplateID}')">
          <div>
            <div class="mix-name">${t.TemplateName}</div>
            <div class="mix-crop">${t.CropType} · ${prods.length} chemical${prods.length !== 1 ? 's' : ''} · tap to edit</div>
          </div>
          <span class="badge ${t.Active === 'Yes' ? 'badge-completed' : 'badge-open'}">${t.Active === 'Yes' ? 'Active' : 'Inactive'}</span>
        </div>
        <div class="mix-body">
          <table class="chem-table">
            <thead><tr><th>Product</th><th>Rate/Acre</th><th>Unit</th><th>Supplied By</th></tr></thead>
            <tbody>${prods.map(p => `
              <tr><td>${p.ProductName}</td><td>${p.RatePerAcre}</td><td>${p.Unit}</td><td>${p.SuppliedBy}</td></tr>
            `).join('')}</tbody>
          </table>
          ${t.Description ? `<p style="color:var(--text-sub);font-size:0.8rem;margin-top:0.75rem">${t.Description}</p>` : ''}
        </div>
      </div>`;
    }).join('') :
    '<div class="empty-state">No mix templates loaded</div>';
}

// ── PRODUCT MODAL ─────────────────────────────────────────────────────────────
function showNewProductModal() {
  document.getElementById('editProductId').value = '';
  document.getElementById('productModalTitle').textContent = 'New Product';
  ['pName','pManufacturer','pCostPerUnit','pREI','pPHI','pNotes'].forEach(id => {
    document.getElementById(id).value = '';
  });
  document.getElementById('pUnit').value = 'fl oz';
  document.getElementById('btnDeleteProduct').style.display = 'none';
  openModal('productModal');
}

function editProduct(productId) {
  const p = DB.products.find(x => x.ProductID === productId);
  if (!p) return;
  document.getElementById('editProductId').value    = p.ProductID;
  document.getElementById('productModalTitle').textContent = 'Edit Product';
  document.getElementById('pName').value            = p.ProductName;
  document.getElementById('pManufacturer').value    = p.Manufacturer;
  document.getElementById('pUnit').value            = p.Unit;
  document.getElementById('pCostPerUnit').value     = p.CostPerUnit;
  document.getElementById('pREI').value             = p.REI_Hours;
  document.getElementById('pPHI').value             = p.PHI_Days;
  document.getElementById('pNotes').value           = p.Notes;
  document.getElementById('btnDeleteProduct').style.display = '';
  openModal('productModal');
}

async function saveProduct() {
  const name = document.getElementById('pName').value.trim();
  if (!name) { showToast('Product name is required', 'error'); return; }

  const editId = document.getElementById('editProductId').value;
  const product = {
    ProductID:    editId || nextId('PROD', DB.products.map(p => p.ProductID)),
    ProductName:  name,
    Manufacturer: document.getElementById('pManufacturer').value.trim(),
    Unit:         document.getElementById('pUnit').value,
    CostPerUnit:  parseFloat(document.getElementById('pCostPerUnit').value || 0),
    REI_Hours:    parseInt(document.getElementById('pREI').value || 0),
    PHI_Days:     parseInt(document.getElementById('pPHI').value || 0),
    Notes:        document.getElementById('pNotes').value.trim(),
  };

  if (editId) {
    const idx = DB.products.findIndex(p => p.ProductID === editId);
    if (idx > -1) DB.products[idx] = product;
  } else {
    DB.products.push(product);
  }

  saveToLocalStorage();
  closeModal();
  showToast(editId ? 'Product updated' : 'Product added', 'success');
  await writeRow('products', product);
  renderProducts();
}

async function deleteProduct() {
  const editId = document.getElementById('editProductId').value;
  if (!editId) return;
  const inUse = DB.orderProds.some(p => p.ProductID === editId) ||
                DB.templateProds.some(p => p.ProductID === editId);
  if (inUse) { showToast('Cannot delete — product is used in orders or templates', 'error'); return; }
  if (!confirm('Delete this product? This cannot be undone.')) return;
  DB.products = DB.products.filter(p => p.ProductID !== editId);
  saveToLocalStorage();
  closeModal();
  showToast('Product deleted', 'success');
  await writeRow('products', { ProductID: editId, _delete: true });
  renderProducts();
}

// ── MIX TEMPLATE MODAL ────────────────────────────────────────────────────────
function showNewMixModal() {
  document.getElementById('editMixId').value = '';
  document.getElementById('mixModalTitle').textContent = 'New Mix Template';
  ['mName','mDescription'].forEach(id => document.getElementById(id).value = '');
  document.getElementById('mCropType').value = 'Corn';
  document.getElementById('mActive').value = 'Yes';
  document.getElementById('mixChemLines').innerHTML = '';
  document.getElementById('btnDeleteMix').style.display = 'none';
  openModal('mixModal');
}

function editMix(templateId) {
  const t = DB.templates.find(x => x.TemplateID === templateId);
  if (!t) return;
  document.getElementById('editMixId').value          = t.TemplateID;
  document.getElementById('mixModalTitle').textContent = 'Edit Mix Template';
  document.getElementById('mName').value              = t.TemplateName;
  document.getElementById('mCropType').value          = t.CropType;
  document.getElementById('mActive').value            = t.Active;
  document.getElementById('mDescription').value       = t.Description;
  document.getElementById('btnDeleteMix').style.display = '';

  // Load existing chemicals
  const prods = DB.templateProds.filter(p => p.TemplateID === templateId);
  document.getElementById('mixChemLines').innerHTML = '';
  prods.forEach(p => addMixChemLine(p));

  openModal('mixModal');
}

function addMixChemLine(prefill) {
  const container = document.getElementById('mixChemLines');
  const div = document.createElement('div');
  div.className = 'chem-line';
  const prodOptions = DB.products.map(p =>
    `<option value="${p.ProductID}" data-unit="${p.Unit}" ${prefill?.ProductID === p.ProductID ? 'selected' : ''}>${p.ProductName}</option>`
  ).join('');

  div.innerHTML = `
    <div class="form-group" style="margin:0">
      <label class="form-label">Product</label>
      <select class="form-input" onchange="this.closest('.chem-line').querySelector('.unit-field').value = this.options[this.selectedIndex].dataset.unit || 'fl oz'">
        ${prodOptions}
      </select>
    </div>
    <div class="form-group" style="margin:0">
      <label class="form-label">Rate/Ac</label>
      <input class="form-input" type="number" step="0.01" value="${prefill?.RatePerAcre||''}" placeholder="0">
    </div>
    <div class="form-group" style="margin:0">
      <label class="form-label">Unit</label>
      <input class="form-input unit-field" type="text" value="${prefill?.Unit||'fl oz'}" placeholder="fl oz">
    </div>
    <div class="form-group" style="margin:0">
      <label class="form-label">By</label>
      <select class="form-input">
        <option ${prefill?.SuppliedBy==='Me'||!prefill?'selected':''}>Me</option>
        <option ${prefill?.SuppliedBy==='Customer'?'selected':''}>Customer</option>
      </select>
    </div>
    <button class="chem-line-remove" onclick="this.parentElement.remove()">×</button>
  `;
  container.appendChild(div);
}

async function saveMix() {
  const name = document.getElementById('mName').value.trim();
  if (!name) { showToast('Template name is required', 'error'); return; }

  const editId = document.getElementById('editMixId').value;
  const templateId = editId || nextId('TMPL', DB.templates.map(t => t.TemplateID));

  const template = {
    TemplateID:   templateId,
    TemplateName: name,
    CropType:     document.getElementById('mCropType').value,
    SprayRate:    parseFloat(document.getElementById('mSprayRate')?.value) || 15,
    Active:       document.getElementById('mActive').value,
    Description:  document.getElementById('mDescription').value.trim(),
  };

  // Gather chemical lines
  const lines = [...document.getElementById('mixChemLines').querySelectorAll('.chem-line')].map((line, i) => {
    const selects = line.querySelectorAll('select');
    const inputs  = line.querySelectorAll('input');
    const prodId  = selects[0].value;
    const prod    = DB.products.find(p => p.ProductID === prodId);
    return {
      LineID:      nextId('MTP', DB.templateProds.map(p => p.LineID)) + '_' + i,
      TemplateID:  templateId,
      ProductID:   prodId,
      ProductName: prod?.ProductName || '',
      RatePerAcre: parseFloat(inputs[0].value || 0),
      Unit:        inputs[1].value,
      SuppliedBy:  selects[1].value,
      Notes:       '',
    };
  });

  if (editId) {
    const idx = DB.templates.findIndex(t => t.TemplateID === editId);
    if (idx > -1) DB.templates[idx] = template;
    // Remove old template product lines
    DB.templateProds = DB.templateProds.filter(p => p.TemplateID !== editId);
  } else {
    DB.templates.push(template);
  }

  DB.templateProds.push(...lines);
  saveToLocalStorage();
  closeModal();
  showToast(editId ? 'Template updated' : 'Template added', 'success');
  await writeRow('templates', template);
  for (const line of lines) await writeRow('templateProds', line);
  renderProducts();
}

async function deleteMix() {
  const editId = document.getElementById('editMixId').value;
  if (!editId) return;
  if (!confirm('Delete this mix template? This cannot be undone.')) return;
  DB.templates     = DB.templates.filter(t => t.TemplateID !== editId);
  DB.templateProds = DB.templateProds.filter(p => p.TemplateID !== editId);
  saveToLocalStorage();
  closeModal();
  showToast('Template deleted', 'success');
  await writeRow('templates', { TemplateID: editId, _delete: true });
  renderProducts();
}

function switchTab(tab, el) {
  document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
  document.querySelectorAll('.tab-content').forEach(t => t.classList.remove('active'));
  el.classList.add('active');
  document.getElementById('tab-' + tab).classList.add('active');
}


// ── FIELDS ───────────────────────────────────────────────────────────────────
let orderFieldsSelected = []; // fields staged in the order modal

function renderFields() {
  const search  = (document.getElementById('fieldSearch')?.value || '').toLowerCase();
  const custFil = document.getElementById('filterFieldCustomer')?.value || '';
  const actFil  = document.getElementById('filterFieldActive')?.value;

  // Populate customer filter
  const custSel = document.getElementById('filterFieldCustomer');
  if (custSel && custSel.options.length <= 1) {
    DB.customers.forEach(c => {
      const o = document.createElement('option');
      o.value = c.CustomerID; o.textContent = c.Name;
      custSel.appendChild(o);
    });
  }

  const filtered = DB.fields.filter(f => {
    const matchSearch = !search || f.FieldName.toLowerCase().includes(search) || f.CustomerName.toLowerCase().includes(search);
    const matchCust   = !custFil || f.CustomerID === custFil;
    const matchActive = actFil === '' || f.Active === actFil;
    return matchSearch && matchCust && matchActive;
  }).sort((a, b) => a.CustomerName.localeCompare(b.CustomerName) || a.FieldName.localeCompare(b.FieldName));

  const el = document.getElementById('fieldsList');
  if (!el) return;
  if (!filtered.length) { el.innerHTML = '<div class="empty-state">No fields found</div>'; return; }

  // Group by customer
  const grouped = {};
  filtered.forEach(f => {
    if (!grouped[f.CustomerName]) grouped[f.CustomerName] = [];
    grouped[f.CustomerName].push(f);
  });

  el.innerHTML = Object.entries(grouped).map(([custName, fields]) => `
    <div class="field-customer-group">
      <div class="field-customer-header">${custName}</div>
      ${fields.map(f => `
        <div class="field-card" onclick="showEditFieldModal('${f.FieldID}')">
          <div class="field-card-main">
            <span class="field-card-name">${f.FieldName}</span>
            ${f.Active === 'No' ? '<span class="badge badge-inactive">Inactive</span>' : ''}
          </div>
          <div class="field-card-meta">
            ${f.Acres ? parseFloat(f.Acres).toFixed(1) + ' ac' : 'Acres TBD'}
            ${f.CentroidLat ? ' · ' + parseFloat(f.CentroidLat).toFixed(4) + ', ' + parseFloat(f.CentroidLng).toFixed(4) : ''}
          </div>
        </div>
      `).join('')}
    </div>
  `).join('');
}

function filterFields() { renderFields(); }

function showNewFieldModal(presetCustomerId) {
  document.getElementById('fieldModalTitle').textContent = 'New Field';
  document.getElementById('editFieldId').value = '';
  document.getElementById('fFieldName').value = '';
  document.getElementById('fFieldAcres').value = '';
  document.getElementById('fFieldLat').value = '';
  document.getElementById('fFieldLng').value = '';
  document.getElementById('fFieldNotes').value = '';
  document.getElementById('fFieldActive').value = 'Yes';
  document.getElementById('btnDeleteField').style.display = 'none';
  // Ensure KML hidden input exists and is cleared
  let kmlNew = document.getElementById('fFieldKML');
  if (!kmlNew) {
    kmlNew = document.createElement('input');
    kmlNew.type = 'hidden'; kmlNew.id = 'fFieldKML';
    document.getElementById('fieldModal').appendChild(kmlNew);
  }
  kmlNew.value = '';
  const prevNew = document.getElementById('fieldPolygonPreview');
  if (prevNew) prevNew.innerHTML = '';

  // Populate customer dropdown
  const sel = document.getElementById('fFieldCustomer');
  sel.innerHTML = '<option value="">— Select customer —</option>';
  DB.customers.forEach(c => {
    const o = document.createElement('option');
    o.value = c.CustomerID; o.textContent = c.Name;
    if (c.CustomerID === presetCustomerId) o.selected = true;
    sel.appendChild(o);
  });

  openModal('fieldModal');
}

function showEditFieldModal(fieldId) {
  const f = DB.fields.find(x => x.FieldID === fieldId);
  if (!f) return;
  document.getElementById('fieldModalTitle').textContent = 'Edit Field';
  document.getElementById('editFieldId').value = f.FieldID;
  document.getElementById('fFieldName').value = f.FieldName;
  document.getElementById('fFieldAcres').value = f.Acres;
  document.getElementById('fFieldLat').value = f.CentroidLat;
  document.getElementById('fFieldLng').value = f.CentroidLng;
  document.getElementById('fFieldNotes').value = f.Notes;
  document.getElementById('fFieldActive').value = f.Active || 'Yes';
  document.getElementById('btnDeleteField').style.display = 'inline-flex';

  const sel = document.getElementById('fFieldCustomer');
  sel.innerHTML = '';
  DB.customers.forEach(c => {
    const o = document.createElement('option');
    o.value = c.CustomerID; o.textContent = c.Name;
    if (c.CustomerID === f.CustomerID) o.selected = true;
    sel.appendChild(o);
  });

  // Ensure fFieldKML hidden input exists and is populated
  let kmlHidden = document.getElementById('fFieldKML');
  if (!kmlHidden) {
    kmlHidden = document.createElement('input');
    kmlHidden.type = 'hidden';
    kmlHidden.id   = 'fFieldKML';
    document.getElementById('fieldModal').appendChild(kmlHidden);
  }
  kmlHidden.value = f.PolygonKML || '';

  // Reload polygon preview from stored KML
  const preview = document.getElementById('fieldPolygonPreview');
  if (preview) {
    if (f.PolygonKML) {
      const pts = GeoUtils.parsePolygon(f.PolygonKML);
      preview.innerHTML = pts ? GeoUtils.polygonToSVG(pts, { width: 220, height: 120 }) : '';
    } else {
      preview.innerHTML = '';
    }
  }
  openModal('fieldModal');
}

async function saveField() {
  const custId   = document.getElementById('fFieldCustomer').value;
  const custName = document.getElementById('fFieldCustomer').selectedOptions[0]?.text || '';
  const name     = document.getElementById('fFieldName').value.trim();
  const editId   = document.getElementById('editFieldId').value;

  if (!custId) { showToast('Please select a customer', 'error'); return; }
  if (!name)   { showToast('Field name is required', 'error'); return; }

  const fieldId = editId || nextId('FLD', DB.fields.map(f => f.FieldID));
  const field = {
    FieldID: fieldId, CustomerID: custId, CustomerName: custName,
    FieldName: name,
    Acres: document.getElementById('fFieldAcres').value || '',
    CentroidLat: document.getElementById('fFieldLat').value || '',
    CentroidLng: document.getElementById('fFieldLng').value || '',
    PolygonKML: document.getElementById('fFieldKML')?.value || (editId ? (DB.fields.find(f => f.FieldID === editId)?.PolygonKML || '') : ''),
    CLU_TractID: '', CLU_FarmNum: '',
    Active: document.getElementById('fFieldActive').value,
    Notes: document.getElementById('fFieldNotes').value
  };

  const idx = DB.fields.findIndex(f => f.FieldID === fieldId);
  if (idx > -1) DB.fields[idx] = field; else DB.fields.push(field);
  await writeRow('fields', field);
  saveToLocalStorage();
  closeModal();
  renderFields();
  showToast(editId ? 'Field updated' : 'Field created', 'success');

  // If field picker is open, refresh it
  if (document.getElementById('fieldPickerModal')?.classList.contains('active')) {
    const custFilter = document.getElementById('fCustomer')?.value;
    renderFieldPicker(custFilter);
  }
}

async function deleteField() {
  const fieldId = document.getElementById('editFieldId').value;
  const inUse   = DB.orderFields.some(f => f.FieldID === fieldId);
  if (inUse) { showToast('Field is used in orders — set Inactive instead', 'error'); return; }
  if (!confirm('Delete this field permanently?')) return;
  DB.fields = DB.fields.filter(f => f.FieldID !== fieldId);
  await writeRow('fields', { FieldID: fieldId, _delete: true });
  saveToLocalStorage();
  closeModal();
  renderFields();
  showToast('Field deleted', 'success');
}

// ── FIELD PICKER (in order modal) ────────────────────────────────────────────
function loadCustomerFields() {
  // Reset field selection when customer changes
  orderFieldsSelected = [];
  renderOrderFieldsList();
}

function showFieldPicker() {
  const custId = document.getElementById('fCustomer').value;
  if (!custId) { showToast('Select a customer first', 'error'); return; }
  const custName = document.getElementById('fCustomer').selectedOptions[0]?.text || '';
  document.getElementById('fieldPickerCustomerName').textContent = custName + ' — select one or more fields';
  renderFieldPicker(custId);
  document.getElementById('modalOverlay').classList.add('active');
  document.getElementById('fieldPickerModal').classList.add('active');
}

function renderFieldPicker(custId) {
  const fields = DB.fields.filter(f => f.CustomerID === custId && f.Active !== 'No');
  const alreadySelected = new Set(orderFieldsSelected.map(f => f.FieldID));
  const el = document.getElementById('fieldPickerList');
  if (!fields.length) {
    el.innerHTML = '<div class="empty-state">No fields for this customer yet</div>';
    return;
  }
  el.innerHTML = fields.map(f => `
    <label class="field-picker-item ${alreadySelected.has(f.FieldID) ? 'already-added' : ''}">
      <input type="checkbox" value="${f.FieldID}" ${alreadySelected.has(f.FieldID) ? 'checked disabled' : ''}>
      <div class="field-picker-info">
        <span class="field-picker-name">${f.FieldName}</span>
        <span class="field-picker-acres">${f.Acres ? parseFloat(f.Acres).toFixed(1) + ' ac' : 'Acres TBD'}</span>
      </div>
    </label>
  `).join('');
}

function confirmFieldSelection() {
  const custId = document.getElementById('fCustomer').value;
  const checked = document.querySelectorAll('#fieldPickerList input[type=checkbox]:checked:not(:disabled)');
  checked.forEach(cb => {
    const field = DB.fields.find(f => f.FieldID === cb.value);
    if (field && !orderFieldsSelected.find(f => f.FieldID === field.FieldID)) {
      orderFieldsSelected.push(field);
    }
  });
  closeFieldPicker();
  renderOrderFieldsList();
}

function closeFieldPicker() {
  document.getElementById('fieldPickerModal').classList.remove('active');
  document.getElementById('modalOverlay').classList.remove('active');
}

function removeOrderField(fieldId) {
  orderFieldsSelected = orderFieldsSelected.filter(f => f.FieldID !== fieldId);
  renderOrderFieldsList();
}

function renderOrderFieldsList() {
  const el   = document.getElementById('orderFieldsList');
  const sumEl = document.getElementById('orderFieldsSummary');
  const acEl  = document.getElementById('orderFieldsAcres');
  if (!el) return;

  if (!orderFieldsSelected.length) {
    el.innerHTML = '<div class="empty-state" style="padding:0.5rem 0;font-size:0.8rem">No fields selected</div>';
    sumEl.style.display = 'none';
    return;
  }

  const total = orderFieldsSelected.reduce((s, f) => s + parseFloat(f.Acres || 0), 0);
  el.innerHTML = orderFieldsSelected.map(f => `
    <div class="order-field-chip">
      <span class="order-field-chip-name">${f.FieldName}</span>
      <span class="order-field-chip-acres">${f.Acres ? parseFloat(f.Acres).toFixed(1) + ' ac' : '—'}</span>
      <button class="order-field-chip-remove" onclick="removeOrderField('${f.FieldID}')">×</button>
    </div>
  `).join('');
  sumEl.style.display = 'block';
  acEl.textContent = total.toFixed(1) + ' ac total';

  // Update chem line acres
  document.querySelectorAll('.chem-acres').forEach(el => el.textContent = total.toFixed(1));
}


// ── MAP INTEGRATION ───────────────────────────────────────────────────────────

function openMapForOrder() {
  const custId = document.getElementById('fCustomer').value;
  if (!custId) { showToast('Select a customer first', 'error'); return; }

  const cust = DB.customers.find(c => c.CustomerID === custId);
  const addr = cust ? [cust.Address, cust.City, cust.State].filter(Boolean).join(', ') : '';

  // Use centroid of already-selected fields as center if available
  let centerLat, centerLng;
  if (orderFieldsSelected.length) {
    const pts = orderFieldsSelected.flatMap(f => {
      const field = DB.fields.find(x => x.FieldID === f.FieldID);
      return field?.CentroidLat ? [{ lat: parseFloat(field.CentroidLat), lng: parseFloat(field.CentroidLng) }] : [];
    });
    if (pts.length) {
      const ctr = GeoUtils.centroid(pts);
      centerLat = ctr.lat; centerLng = ctr.lng;
    }
  }

  // Build preselected list with points decoded from stored KML
  const preselected = orderFieldsSelected.map(f => {
    const field = DB.fields.find(x => x.FieldID === f.FieldID) || f;
    const points = field.PolygonKML ? GeoUtils.parsePolygon(field.PolygonKML) : null;
    return {
      id:        field.FieldID,
      fieldName: field.FieldName,
      acres:     field.Acres || '0',
      points,
      kml:       field.PolygonKML || '',
      fromDB:    true,
    };
  }).filter(f => f.points && f.points.length >= 3);

  // Load all customer fields with KML as background context
  const customerFields = DB.fields
    .filter(f => f.CustomerID === custId && f.PolygonKML && f.Active !== 'No')
    .map(f => {
      const points = GeoUtils.parsePolygon(f.PolygonKML);
      if (!points || points.length < 3) return null;
      return { id: f.FieldID, fieldName: f.FieldName, acres: f.Acres||'0', points, kml: f.PolygonKML, fromDB: true };
    }).filter(Boolean);

  MapModal.open({
    centerLat,
    centerLng,
    customerAddress: addr,
    preselected,
    customerFields,
    onConfirm: async (mapFields) => {
      // Prompt for name on drawn/pasted fields before saving
      for (const mf of mapFields) {
        if ((mf.id.startsWith('DRAWN') || mf.id.startsWith('PASTE')) && !mf.fieldName) {
          const name = prompt(`Name this drawn field (${mf.acres} ac):`, 'New Field');
          mf.fieldName = name || ('Field-' + mf.id.slice(-4));
        }
      }
      await onMapFieldsConfirmed(mapFields, custId, cust?.Name || '');
    }
  });
}

async function onMapFieldsConfirmed(mapFields, custId, custName) {
  // mapFields: [{id, farmNum, acres, points, kml, centroid, fieldName, fromDB}]
  // Save each as a permanent Field record, then add to order selection.
  // NO prompts — fields get auto-named; user can rename via Edit Field later.

  let added = 0;

  for (const mf of mapFields) {

    // ── Preselected existing DB fields ──────────────────────────────────────
    if (mf.fromDB) {
      const dbField = DB.fields.find(f => f.FieldID === mf.id);
      if (dbField && !orderFieldsSelected.find(f => f.FieldID === dbField.FieldID)) {
        orderFieldsSelected.push(dbField);
        added++;
      }
      continue;
    }

    // ── New field (drawn, pasted, or loaded from shapefile/CLU) ─────────────
    // Determine a stable CLU id for dedup (blank for drawn/pasted/existing)
    const isEphemeral = ['DRAWN-','PASTE-','EXISTING-'].some(p => mf.id.startsWith(p));
    const cluId = isEphemeral ? '' : mf.id;

    // Check if already saved for this customer (same CLU id)
    let field = cluId
      ? DB.fields.find(f => f.CLU_TractID === cluId && f.CustomerID === custId)
      : null;

    if (!field) {
      // Auto-generate a name the user can rename later
      const fieldId = nextId('FLD', DB.fields.map(f => f.FieldID));
      const acres   = parseFloat(mf.acres || 0).toFixed(1);
      const name    = mf.fieldName ||
                      (mf.farmNum ? `Farm ${mf.farmNum}` : `Field ${fieldId}`) +
                      ` (${acres} ac)`;
      const ctr = mf.centroid || GeoUtils.centroid(mf.points || []);

      field = {
        FieldID:      fieldId,
        CustomerID:   custId,
        CustomerName: custName,
        FieldName:    name,
        Acres:        acres,
        CentroidLat:  ctr ? ctr.lat.toFixed(6) : '',
        CentroidLng:  ctr ? ctr.lng.toFixed(6) : '',
        PolygonKML:   (() => {
          // Simplify polygon to reduce URL size (RDP algorithm)
          if (!mf.points || mf.points.length === 0) return '';
          const simplified = GeoUtils.simplifyPolygon(mf.points, 0.00005);
          return GeoUtils.pointsToKML(simplified);
        })(),
        CLU_TractID:  cluId,
        CLU_FarmNum:  mf.farmNum || '',
        Active:       'Yes',
        Notes:        '',
      };
      DB.fields.push(field);
      await writeRow('fields', field);
    }

    // Add to order selection if not already there
    if (!orderFieldsSelected.find(f => f.FieldID === field.FieldID)) {
      orderFieldsSelected.push(field);
      added++;
    }
  }

  saveToLocalStorage();
  renderOrderFieldsList();
  showToast(
    added > 0
      ? `${added} field${added !== 1 ? 's' : ''} added — tap a chip to rename`
      : 'Fields already in order',
    'success'
  );
}

function openMapForField() {
  const custId   = document.getElementById('fFieldCustomer').value;
  const cust     = DB.customers.find(c => c.CustomerID === custId);
  const addr     = cust ? [cust.Address, cust.City, cust.State].filter(Boolean).join(', ') : '';
  const existLat = parseFloat(document.getElementById('fFieldLat').value) || null;
  const existLng = parseFloat(document.getElementById('fFieldLng').value) || null;

  // Build preselected from existing KML in the field form
  const existingKML = document.getElementById('fFieldKML')?.value || '';
  const existingPoints = existingKML ? GeoUtils.parsePolygon(existingKML) : null;
  const preselected = existingPoints && existingPoints.length >= 3 ? [{
    id:        'EXISTING-' + Date.now(),
    fieldName: document.getElementById('fFieldName').value || 'Current boundary',
    acres:     document.getElementById('fFieldAcres').value || '0',
    points:    existingPoints,
    kml:       existingKML,
    fromDB:    true,
  }] : [];

  // Hide field modal while map is open
  document.getElementById('fieldModal').style.display = 'none';

  MapModal.open({
    centerLat:       existLat || (existingPoints ? GeoUtils.centroid(existingPoints)?.lat : null),
    centerLng:       existLng || (existingPoints ? GeoUtils.centroid(existingPoints)?.lng : null),
    customerAddress: addr,
    preselected,
    onConfirm: (mapFields) => {
      // Restore field modal
      document.getElementById('fieldModal').style.display = '';
      if (!mapFields.length) return;
      // Use the first confirmed field (could be existing or newly drawn)
      const mf  = mapFields[0];
      const ctr = mf.centroid || GeoUtils.centroid(mf.points);
      if (ctr) {
        document.getElementById('fFieldLat').value = ctr.lat.toFixed(6);
        document.getElementById('fFieldLng').value = ctr.lng.toFixed(6);
      }
      if (mf.acres && !mf.fromDB) document.getElementById('fFieldAcres').value = mf.acres;
      const preview = document.getElementById('fieldPolygonPreview');
      if (preview && mf.points) preview.innerHTML = GeoUtils.polygonToSVG(mf.points, { width: 240, height: 120 });
      let kmlHidden = document.getElementById('fFieldKML');
      if (!kmlHidden) {
        kmlHidden = document.createElement('input');
        kmlHidden.type = 'hidden'; kmlHidden.id = 'fFieldKML';
        document.getElementById('fieldModal').appendChild(kmlHidden);
      }
      kmlHidden.value = mf.kml || '';
      showToast('Field boundary updated from map', 'success');
    },
    onClose: () => {
      document.getElementById('fieldModal').style.display = '';
    }
  });
}


// ── UNIT CONVERSION ───────────────────────────────────────────────────────────
function smartUnit(amount, unit) {
  // Convert to sensible display unit based on quantity
  const u = (unit || '').toLowerCase().trim();
  if (['fl oz', 'fl. oz', 'floz', 'oz'].includes(u)) {
    if (amount >= 128) return { val: (amount / 128).toFixed(2), unit: 'gal' };
    if (amount >= 32)  return { val: (amount / 32).toFixed(2),  unit: 'qt' };
    if (amount >= 16)  return { val: (amount / 16).toFixed(2),  unit: 'pt' };
    return { val: amount.toFixed(1), unit };
  }
  if (u === 'ml') {
    if (amount >= 3785) return { val: (amount / 3785).toFixed(2), unit: 'gal' };
    if (amount >= 1000) return { val: (amount / 1000).toFixed(2), unit: 'L' };
    return { val: amount.toFixed(1), unit };
  }
  if (u === 'lb' || u === 'lbs') {
    if (amount >= 2000) return { val: (amount / 2000).toFixed(2), unit: 'ton' };
    return { val: amount.toFixed(2), unit: 'lb' };
  }
  return { val: amount % 1 === 0 ? amount : amount.toFixed(2), unit };
}

function fmtAmt(amount, unit) {
  const { val, unit: u } = smartUnit(amount, unit);
  return `${val} ${u}`;
}

// ── BATCH CALCULATOR ─────────────────────────────────────────────────────────
function renderCalculator() {
  switchCalcTab('orders', document.querySelector('.calc-tab'));
}

function switchCalcTab(tab, el) {
  document.querySelectorAll('.calc-tab').forEach(t => t.classList.remove('active'));
  document.querySelectorAll('#view-calculator .tab-content').forEach(t => t.classList.remove('active'));
  if (el) el.classList.add('active');
  document.getElementById('calcTab-' + tab).classList.add('active');
  if (tab === 'orders') renderOrderCalc();
  else                  renderMixCalc();
}

// ── TAB 1: ORDER BATCH CALC ───────────────────────────────────────────────────
function renderOrderCalc() {
  const eligible = DB.orders.filter(o => DB.orderProds.some(p => p.OrderID === o.OrderID));
  document.getElementById('calcOrderList').innerHTML = eligible.length ?
    eligible.map(o => `
      <div class="calc-order-item" onclick="toggleCalcOrder('${o.OrderID}', this)">
        <input type="checkbox" id="calc_${o.OrderID}" ${selectedCalcOrders.has(o.OrderID) ? 'checked' : ''} onclick="event.stopPropagation()">
        <div>
          <div style="font-weight:500">${o.OrderID} — ${o.CustomerName}</div>
          <div style="font-size:0.75rem;color:var(--text-sub)">
            ${DB.orderFields.filter(f=>f.OrderID===o.OrderID).map(f=>f.FieldName).join(', ')||'—'} · ${o.TotalAcres||0} ac · ${o.CropType}
          </div>
        </div>
      </div>`).join('') :
    '<div style="color:var(--text-sub);font-size:0.82rem">No orders with chemicals found</div>';
  runOrderCalc();
}

function toggleCalcOrder(orderId, el) {
  const cb = document.getElementById('calc_' + orderId);
  if (selectedCalcOrders.has(orderId)) { selectedCalcOrders.delete(orderId); if(cb) cb.checked = false; }
  else                                 { selectedCalcOrders.add(orderId);    if(cb) cb.checked = true; }
  runOrderCalc();
}

function runOrderCalc() {
  const resultsEl = document.getElementById('calcResults');
  if (!resultsEl) return;
  if (selectedCalcOrders.size === 0) {
    resultsEl.innerHTML = '<div class="calc-empty">Select one or more orders above</div>';
    return;
  }

  const tankSize  = parseFloat(document.getElementById('tankSize')?.value || AppSettings.defaultTankSize);

  // Derive spray rate per order from its template, fall back to settings default
  let totalAcres   = 0;
  let totalGallons = 0;
  const orderSprayRates = [];
  [...selectedCalcOrders].forEach(id => {
    const o = DB.orders.find(x => x.OrderID === id);
    if (!o) return;
    const ac = parseFloat(o.TotalAcres || 0);
    const tmpl = DB.templates.find(t => t.TemplateID === o.TemplateUsed);
    const rate = parseFloat(tmpl?.SprayRate || AppSettings.defaultSprayRate);
    totalAcres   += ac;
    totalGallons += ac * rate;
    orderSprayRates.push({ id, rate, ac });
  });
  // Show effective rate in UI
  const effectiveRate = totalAcres > 0 ? (totalGallons / totalAcres).toFixed(1) : AppSettings.defaultSprayRate;
  const rateEl = document.getElementById('calcEffectiveRate');
  if (rateEl) rateEl.textContent = effectiveRate + ' gal/ac (weighted avg from templates)';
  const loadsNeeded  = tankSize > 0 ? Math.ceil(totalGallons / tankSize) : 0;
  const acresPerLoad = loadsNeeded > 0 ? (totalAcres / loadsNeeded).toFixed(1) : 0;

  // Aggregate chemicals
  const chemMap = {};
  [...selectedCalcOrders].forEach(orderId => {
    DB.orderProds.filter(p => p.OrderID === orderId).forEach(p => {
      const key = p.ProductID || p.ProductName;
      if (!chemMap[key]) chemMap[key] = { name: p.ProductName, unit: p.Unit, totalUnits: 0, ratePerAc: 0, costPerUnit: parseFloat(p.CostPerUnit||0) };
      chemMap[key].totalUnits += parseFloat(p.TotalUnitsNeeded || 0);
      chemMap[key].ratePerAc  += parseFloat(p.RatePerAcre || 0);
    });
  });

  const chems = Object.values(chemMap);
  if (!chems.length) { resultsEl.innerHTML = '<div class="calc-empty">No chemicals on selected orders</div>'; return; }

  const perLoadRows = chems.map(c => {
    const totalDisp = fmtAmt(c.totalUnits, c.unit);
    const perLoadAmt = loadsNeeded > 0 ? c.totalUnits / loadsNeeded : 0;
    const perLoadDisp = loadsNeeded > 0 ? fmtAmt(perLoadAmt, c.unit) : '—';
    const cost = c.costPerUnit > 0 ? ' · $' + (c.totalUnits * c.costPerUnit).toFixed(2) : '';
    return `<div class="calc-result-item">
      <div class="calc-product-name">${c.name}</div>
      <div class="calc-product-qty">${totalDisp} total</div>
      <div class="calc-product-detail">
        <span>${perLoadDisp}/load${cost}</span>
        <span style="color:var(--text-sub);font-size:0.75rem">${c.ratePerAc.toFixed(1)} ${c.unit}/ac</span>
      </div>
    </div>`;
  }).join('');

  resultsEl.innerHTML = `
    <div class="calc-summary-row">
      <div class="calc-summary-item"><div class="calc-summary-val">${totalAcres.toFixed(1)}</div><div class="calc-summary-lbl">Total Acres</div></div>
      <div class="calc-summary-item"><div class="calc-summary-val">${totalGallons.toFixed(0)}</div><div class="calc-summary-lbl">Total Gallons</div></div>
      <div class="calc-summary-item accent"><div class="calc-summary-val">${loadsNeeded}</div><div class="calc-summary-lbl">Tank Loads</div></div>
      <div class="calc-summary-item"><div class="calc-summary-val">${acresPerLoad}</div><div class="calc-summary-lbl">Ac/Load</div></div>
    </div>
    <div style="font-size:0.75rem;color:var(--text-sub);margin-bottom:0.75rem" id="calcEffectiveRate">${tankSize} gal tank · ${effectiveRate} gal/ac avg spray rate</div>
    ${perLoadRows}`;
}

// ── TAB 2: MIX CALC (no order needed) ────────────────────────────────────────
function renderMixCalc() {
  const tmplSel = document.getElementById('mixCalcTemplate');
  if (tmplSel) {
    tmplSel.innerHTML = '<option value="">— Select template —</option>';
    DB.templates.filter(t => t.Active === 'Yes').forEach(t => {
      const o = document.createElement('option');
      o.value = t.TemplateID;
      o.textContent = t.TemplateName + ' (' + t.CropType + (t.SprayRate ? ' · ' + t.SprayRate + ' gal/ac' : '') + ')';
      tmplSel.appendChild(o);
    });
  }
  runMixCalc();
}

// Track which mix calc input was last changed so they override each other
let mixCalcLastChanged = 'acres'; // 'acres' or 'gallons'

function onMixCalcAcresInput() {
  mixCalcLastChanged = 'acres';
  // Sync gallons from acres
  const tmplId  = document.getElementById('mixCalcTemplate')?.value;
  const tmpl    = DB.templates.find(t => t.TemplateID === tmplId);
  const galPerAc= parseFloat(tmpl?.SprayRate || 15);
  const acres   = parseFloat(document.getElementById('mixCalcAcres')?.value || 0);
  if (acres > 0) {
    const galEl = document.getElementById('mixCalcTotalGal');
    if (galEl) galEl.value = (acres * galPerAc).toFixed(0);
  }
  runMixCalc();
}

function onMixCalcGalInput() {
  mixCalcLastChanged = 'gallons';
  // Sync acres from gallons
  const tmplId  = document.getElementById('mixCalcTemplate')?.value;
  const tmpl    = DB.templates.find(t => t.TemplateID === tmplId);
  const galPerAc= parseFloat(tmpl?.SprayRate || 15);
  const gal     = parseFloat(document.getElementById('mixCalcTotalGal')?.value || 0);
  if (gal > 0 && galPerAc > 0) {
    const acEl = document.getElementById('mixCalcAcres');
    if (acEl) acEl.value = (gal / galPerAc).toFixed(1);
  }
  runMixCalc();
}

function onMixCalcTemplateChange() {
  // When template changes, update spray rate display and recalc
  const tmplId = document.getElementById('mixCalcTemplate')?.value;
  const tmpl   = DB.templates.find(t => t.TemplateID === tmplId);
  const rate   = parseFloat(tmpl?.SprayRate || 15);
  const rateEl = document.getElementById('mixCalcSprayRateDisplay');
  if (rateEl) rateEl.textContent = rate + ' gal/ac';
  // Recalc from whichever was last changed
  if (mixCalcLastChanged === 'acres') onMixCalcAcresInput();
  else onMixCalcGalInput();
}

function runMixCalc() {
  const resultsEl = document.getElementById('mixCalcResults');
  if (!resultsEl) return;
  const tmplId   = document.getElementById('mixCalcTemplate')?.value;
  if (!tmplId) { resultsEl.innerHTML = '<div class="calc-empty">Select a mix template above</div>'; return; }

  const tmpl     = DB.templates.find(t => t.TemplateID === tmplId);
  const galPerAc = parseFloat(tmpl?.SprayRate || 15);
  const totalGal = parseFloat(document.getElementById('mixCalcTotalGal')?.value || 0);
  const totalAc  = parseFloat(document.getElementById('mixCalcAcres')?.value || 0);
  const tankSize = parseFloat(document.getElementById('mixCalcTankSize')?.value || 0);

  const prods = DB.templateProds.filter(p => p.TemplateID === tmplId);
  if (!prods.length) { resultsEl.innerHTML = '<div class="calc-empty">No chemicals in this template</div>'; return; }

  // gallons and acres stay in sync via the input handlers
  // Use whichever is available
  const gallons = totalGal > 0 ? totalGal : (totalAc > 0 ? totalAc * galPerAc : 0);
  const acres   = gallons > 0 && galPerAc > 0 ? gallons / galPerAc : totalAc;
  if (!gallons && !acres) {
    resultsEl.innerHTML = '<div class="calc-empty">Enter acres or total gallons to calculate</div>';
    return;
  }

  // Tank size is informational only — doesn't override gallons or acres
  const loads      = tankSize > 0 ? Math.ceil(gallons / tankSize) : 1;
  const galPerLoad = loads > 0 && gallons > 0 ? gallons / loads : gallons;

  const rows = prods.map(p => {
    const rate       = parseFloat(p.RatePerAcre || 0);
    const totalAmt   = rate * acres;
    const perLoadAmt = loads > 0 ? totalAmt / loads : totalAmt;
    const totalDisp  = fmtAmt(totalAmt, p.Unit);
    const perDisp    = fmtAmt(perLoadAmt, p.Unit);
    return `<div class="calc-result-item">
      <div class="calc-product-name">${p.ProductName}</div>
      <div class="calc-product-qty">${totalDisp} total</div>
      <div class="calc-product-detail">
        <span>${perDisp}/load · ${rate} ${p.Unit}/ac</span>
        <span style="color:var(--text-sub);font-size:0.75rem">${p.SuppliedBy === 'Customer' ? 'By customer' : ''}</span>
      </div>
    </div>`;
  }).join('');

  const loadsDisp    = tankSize > 0 ? loads : '—';
  const galLoadDisp  = tankSize > 0 ? galPerLoad.toFixed(0) : '—';

  resultsEl.innerHTML = `
    <div class="calc-summary-row">
      <div class="calc-summary-item"><div class="calc-summary-val">${acres.toFixed(1)}</div><div class="calc-summary-lbl">Acres</div></div>
      <div class="calc-summary-item"><div class="calc-summary-val">${gallons.toFixed(0)}</div><div class="calc-summary-lbl">Total Gal</div></div>
      <div class="calc-summary-item accent"><div class="calc-summary-val">${loadsDisp}</div><div class="calc-summary-lbl">Loads</div></div>
      <div class="calc-summary-item"><div class="calc-summary-val">${galLoadDisp}</div><div class="calc-summary-lbl">Gal/Load</div></div>
    </div>
    <div style="font-size:0.75rem;color:var(--text-sub);margin-bottom:0.75rem">
      Spray rate: ${galPerAc} gal/ac (from template)${tankSize > 0 ? ' · ' + tankSize + ' gal tank' : ''}
    </div>
    ${rows}`;
}


// ── GDU FUNGICIDE PREDICTOR ──────────────────────────────────────────────────
let gduResults  = [];   // cached results per session
let gduRunning  = false;

function renderGDU() {
  const container = document.getElementById('gduContent');
  if (!container) return;

  // Find all corn orders with planting date + RM set
  const cornOrders = DB.orders.filter(o =>
    o.CropType === 'Corn' && o.PlantingDate && o.RelativeMaturity &&
    o.Status !== 'Completed'
  );

  if (!cornOrders.length) {
    container.innerHTML = `<div class="empty-state">No active corn orders with Planting Date and RM set.<br>
      Add these fields to your corn orders to enable GDU prediction.</div>`;
    return;
  }

  // Show order list with run button if no results yet
  if (!gduResults.length) {
    container.innerHTML = `
      <div class="gdu-intro">
        <p>Found <strong>${cornOrders.length}</strong> active corn order${cornOrders.length !== 1 ? 's' : ''} with planting data.</p>
        <p style="color:var(--text-sub);font-size:0.85rem;margin-top:0.25rem">
          Fetches daily temps from Open-Meteo (free) from each field's planting date through today + 14-day forecast.
        </p>
        <button class="btn-primary" style="margin-top:0.75rem" onclick="runGDUAnalysis()">
          🌽 Run GDU Analysis
        </button>
      </div>
      <div class="gdu-order-preview">
        ${cornOrders.map(o => {
          const fields = DB.orderFields.filter(f => f.OrderID === o.OrderID);
          return `<div class="gdu-preview-row">
            <span>${o.OrderID} — ${o.CustomerName}</span>
            <span style="color:var(--text-sub);font-size:0.8rem">
              RM ${o.RelativeMaturity} · Planted ${GDUCalc.fmtDate(o.PlantingDate)} · 
              ${fields.map(f=>f.FieldName).join(', ') || '—'}
            </span>
          </div>`;
        }).join('')}
      </div>`;
    return;
  }

  // Show results
  container.innerHTML = `
    <div style="display:flex;align-items:center;gap:0.75rem;margin-bottom:1rem;flex-wrap:wrap">
      <button class="btn-secondary" onclick="runGDUAnalysis()">🔄 Refresh Analysis</button>
      <span style="font-size:0.8rem;color:var(--text-sub)" id="gduLastRun"></span>
    </div>
    ${gduResults.map(r => renderGDUCard(r)).join('')}`;
}

function renderGDUCard(r) {
  if (r.error) {
    return `<div class="gdu-card error">
      <div class="gdu-card-header">${r.orderId || 'Unknown'} — ${r.customerName || ''}</div>
      <div style="color:var(--danger);font-size:0.85rem">${r.error}</div>
    </div>`;
  }

  const urgency    = GDUCalc.urgencyClass(r);
  const barWidth   = Math.min(100, r.pctToVT) + '%';
  // Mark window start (VT onset, 80%) and end (R1/R2, 105%) on bar
  const winStartPct = Math.min(100, r.windowStartGDU / r.vtGDU * 100) + '%';
  const winEndPct   = Math.min(100, r.windowEndGDU   / r.vtGDU * 100) + '%';
  const vtPct       = '100%';

  // Find the order for this result to get scheduled date
  const order = DB.orders.find(o => o.OrderID === r.orderId);
  const scheduledDate = order?.ScheduledDate || '';
  const scheduledDisp = scheduledDate ? GDUCalc.fmtDate(scheduledDate) : 'Not set';

  const statusMsg = r.pctToVT >= 105
    ? '⚠ Past R1 — fungicide window closing'
    : r.pctToVT >= 100
    ? '✅ At VT — ideal window open now'
    : r.pctToVT >= 80
    ? '⚡ Approaching VT — apply soon'
    : r.targetProjected
    ? `Projected application: <strong>${GDUCalc.fmtDate(r.targetDate)}</strong>`
    : `Target date: <strong>${GDUCalc.fmtDate(r.targetDate)}</strong>`;

  const schedMatchesTarget = scheduledDate === r.targetDate;

  return `<div class="gdu-card ${urgency}">
    <div class="gdu-card-header">
      <div>
        <div class="gdu-card-title">${r.customerName} — ${r.fieldNames}</div>
        <div class="gdu-card-sub">${r.orderId} · RM ${r.rm} · Planted ${GDUCalc.fmtDate(r.plantDate)} · ${r.stage}</div>
      </div>
      <div class="gdu-card-gdu">
        <div class="gdu-val">${r.currentGDU}</div>
        <div class="gdu-lbl">GDU today</div>
      </div>
    </div>

    <div class="gdu-progress-wrap">
      <div class="gdu-progress-bar">
        <div class="gdu-progress-fill" style="width:${barWidth}"></div>
        <!-- VT-R1 window band -->
        <div class="gdu-window-band" style="left:${winStartPct};width:${Math.min(100,r.windowEndGDU/r.vtGDU*100) - Math.min(100,r.windowStartGDU/r.vtGDU*100)}%" title="VT–R1 fungicide window"></div>
        <div class="gdu-vt-label" style="left:${vtPct}" title="VT — Begin applying">VT</div>
      </div>
      <div class="gdu-progress-labels">
        <span>Planting</span>
        <span style="color:var(--warn);font-size:0.68rem">▓ Fungicide window (VT→R1)</span>
        <span>VT=${r.vtGDU} GDU</span>
      </div>
    </div>

    <div class="gdu-window-row">
      <div class="gdu-window-item">
        <span class="gdu-window-lbl">Window opens (VT)</span>
        <span class="gdu-window-val">${GDUCalc.fmtDate(r.windowStart)}</span>
      </div>
      <div class="gdu-window-item target">
        <span class="gdu-window-lbl">🎯 Peak (VT-R1)</span>
        <span class="gdu-window-val">${GDUCalc.fmtDate(r.targetDate)}${r.targetProjected ? ' *' : ''}</span>
      </div>
      <div class="gdu-window-item">
        <span class="gdu-window-lbl">Window closes (R1)</span>
        <span class="gdu-window-val">${GDUCalc.fmtDate(r.windowEnd)}</span>
      </div>
    </div>

    <div class="gdu-scheduled-row">
      <div>
        <span class="gdu-window-lbl">Scheduled date</span>
        <span class="gdu-scheduled-val ${scheduledDate ? '' : 'unset'}">${scheduledDisp}</span>
      </div>
      ${!schedMatchesTarget && r.targetDate ? `
        <button class="btn-secondary gdu-set-btn" onclick="setGDUScheduledDate('${r.orderId}', '${r.targetDate}')">
          Set to ${GDUCalc.fmtDate(r.targetDate)}
        </button>` : schedMatchesTarget ? '<span style="color:var(--accent2);font-size:0.8rem">✓ Matches target</span>' : ''}
    </div>

    <div class="gdu-status ${urgency}">${statusMsg}${r.targetProjected ? ' <span style="color:var(--text-sub);font-size:0.78rem">(*projected beyond 14-day forecast)</span>' : ''}</div>

    <div class="gdu-data-quality">
      <span>📡 ${r.histDays||0}d ERA5 actual · ${(r.tier2Days||0)+(r.fcstDays||0)}d GFS forecast · ${r.seasDays||0}d SEAS5 seasonal · ${r.coverage||0}% coverage</span>
      <button class="gdu-detail-toggle" onclick="toggleGDUDetail('${r.orderId}')">▼ Daily data</button>
    </div>

    <div class="gdu-daily-table" id="gduDetail-${r.orderId}" style="display:none">
      <table>
        <thead><tr><th>Date</th><th>High</th><th>Low</th><th>GDU/day</th><th>Cum GDU</th><th></th></tr></thead>
        <tbody>
          ${(r.withGDU||[]).filter(d=>(d.cumGDU||0)<=(r.windowEndGDU||9999)*1.05).map(d=>`<tr class="${d.isForecast?'forecast-row':''} ${d.cumGDU>=r.windowStartGDU&&d.cumGDU<=r.windowEndGDU?'window-row':''}">
              <td>${d.date}</td>
              <td>${d.maxF!=null?d.maxF.toFixed(1)+'°':'—'}</td>
              <td>${d.minF!=null?d.minF.toFixed(1)+'°':'—'}</td>
              <td>${d.dailyGDU!=null?d.dailyGDU.toFixed(1):'—'}</td>
              <td>${d.cumGDU!=null?Math.round(d.cumGDU):'—'}</td>
              <td style="color:var(--text-sub);font-size:0.7rem">${d.tier===3?'seasonal':d.isForecast?'GFS':'ERA5'}</td>
            </tr>`).join('')}
        </tbody>
      </table>
      <div style="color:var(--text-sub);font-size:0.72rem;padding:0.3rem 0">Planting → R1 (${r.windowEndGDU} GDU) · ${(r.withGDU||[]).filter(d=>(d.cumGDU||0)<=(r.windowEndGDU||9999)*1.05).length} days</div>
    </div>
  </div>`;
}

function toggleGDUDetail(orderId) {
  const el = document.getElementById('gduDetail-' + orderId);
  const btn = el?.previousElementSibling?.querySelector('.gdu-detail-toggle');
  if (!el) return;
  const show = el.style.display === 'none';
  el.style.display = show ? 'block' : 'none';
  if (btn) btn.textContent = show ? '▲ Daily data' : '▼ Daily data';
}

async function setGDUScheduledDate(orderId, date) {
  const o = DB.orders.find(x => x.OrderID === orderId);
  if (!o) return;
  o.ScheduledDate = date;
  await writeRow('orders', o);
  saveToLocalStorage();
  showToast(`Scheduled date set to ${GDUCalc.fmtDate(date)} for ${orderId}`, 'success');
  // Re-render GDU cards to show updated scheduled date
  renderGDU();
}

async function runGDUAnalysis() {
  if (gduRunning) return;
  gduRunning = true;
  const container = document.getElementById('gduContent');
  if (container) container.innerHTML = '<div class="gdu-loading">🌽 Fetching weather data and calculating GDUs...<br><span style="color:var(--text-sub);font-size:0.82rem">This may take a moment for each field</span></div>';

  const cornOrders = DB.orders.filter(o =>
    o.CropType === 'Corn' && o.PlantingDate && o.RelativeMaturity && o.Status !== 'Completed'
  );

  gduResults = [];
  for (const o of cornOrders) {
    // Get primary field for this order (for lat/lng)
    const orderField = DB.orderFields.find(f => f.OrderID === o.OrderID);
    const field = orderField ? DB.fields.find(f => f.FieldID === orderField.FieldID) : null;
    const result = await GDUCalc.analyzeOrder(o, field);
    gduResults.push(result);
  }

  // Sort by urgency (highest pctToTarget first)
  gduResults.sort((a, b) => (b.pctToVT || 0) - (a.pctToVT || 0));

  gduRunning = false;
  renderGDU();
  const lastRun = document.getElementById('gduLastRun');
  if (lastRun) lastRun.textContent = 'Updated ' + new Date().toLocaleTimeString();
}

// ── REPORTS / OPERATIONS ─────────────────────────────────────────────────────

function onReportDateChange() {
  const val = document.getElementById('reportDateRange')?.value;
  const customEl = document.getElementById('reportCustomDates');
  if (customEl) customEl.style.display = val === 'custom' ? 'flex' : 'none';
  renderReports();
}

function getReportDateRange() {
  const sel = document.getElementById('reportDateRange')?.value || 'week';
  const today = new Date();
  today.setHours(0,0,0,0);

  const fmt = d => d.toISOString().split('T')[0];

  if (sel === 'custom') {
    return {
      from: document.getElementById('reportDateFrom')?.value || fmt(today),
      to:   document.getElementById('reportDateTo')?.value   || fmt(today),
      label: 'Custom range',
    };
  }

  const dow   = today.getDay(); // 0=Sun
  const mon   = new Date(today); mon.setDate(today.getDate() - (dow === 0 ? 6 : dow - 1));
  const sun   = new Date(mon);  sun.setDate(mon.getDate() + 6);
  const nextMon = new Date(mon); nextMon.setDate(mon.getDate() + 7);
  const nextSun = new Date(sun); nextSun.setDate(sun.getDate() + 7);
  const monthStart = new Date(today.getFullYear(), today.getMonth(), 1);
  const monthEnd   = new Date(today.getFullYear(), today.getMonth() + 1, 0);
  const seasonStart= new Date(today.getFullYear(), 3, 1);  // Apr 1
  const seasonEnd  = new Date(today.getFullYear(), 10, 30); // Nov 30

  const tomorrow = new Date(today); tomorrow.setDate(today.getDate() + 1);

  const ranges = {
    today:     { from: fmt(today),      to: fmt(today),      label: 'Today' },
    tomorrow:  { from: fmt(tomorrow),   to: fmt(tomorrow),   label: 'Tomorrow' },
    week:      { from: fmt(mon),        to: fmt(sun),        label: 'This Week' },
    next_week: { from: fmt(nextMon),    to: fmt(nextSun),    label: 'Next Week' },
    month:     { from: fmt(monthStart), to: fmt(monthEnd),   label: 'This Month' },
    season:    { from: fmt(seasonStart),to: fmt(seasonEnd),  label: 'This Season' },
  };
  return ranges[sel] || ranges.week;
}

function renderReports() {
  const { from, to, label } = getReportDateRange();
  const pilotFilter  = document.getElementById('reportPilot')?.value  || '';
  const statusFilter = document.getElementById('reportStatus')?.value || '';

  // Populate pilot dropdown once
  const pilotSel = document.getElementById('reportPilot');
  if (pilotSel && pilotSel.options.length <= 1) {
    DB.pilots.filter(p => p.Active === 'Yes').forEach(p => {
      const o = document.createElement('option');
      o.value = p.PilotID; o.textContent = p.Name;
      pilotSel.appendChild(o);
    });
  }

  // Filter orders by date range, pilot, status
  const filtered = DB.orders.filter(o => {
    const d = o.ScheduledDate || o.OrderDate || '';
    if (!d) return false;
    if (d < from || d > to) return false;
    if (pilotFilter  && o.PilotID  !== pilotFilter)  return false;
    if (statusFilter && o.Status   !== statusFilter)  return false;
    return true;
  }).sort((a, b) => (a.ScheduledDate||a.OrderDate||'').localeCompare(b.ScheduledDate||b.OrderDate||''));

  const el = document.getElementById('reportContent');
  if (!el) return;

  if (!filtered.length) {
    el.innerHTML = `<div class="empty-state">No orders found for ${label}</div>`;
    return;
  }

  // ── SECTION 1: Field summary ──────────────────────────────────────────────
  let totalAcres = 0;
  const fieldRows = filtered.map(o => {
    const fields = DB.orderFields.filter(f => f.OrderID === o.OrderID);
    const ac = parseFloat(o.TotalAcres || 0);
    totalAcres += ac;
    return `<tr>
      <td>${o.ScheduledDate ? fmtDate(o.ScheduledDate) : fmtDate(o.OrderDate)}</td>
      <td>${o.OrderID}</td>
      <td>${o.CustomerName}</td>
      <td>${fields.map(f => f.FieldName).join(', ') || '—'}</td>
      <td>${o.CropType || '—'}</td>
      <td>${o.PilotName || '—'}</td>
      <td>${statusBadge(o.Status)}</td>
      <td style="text-align:right">${ac > 0 ? ac.toFixed(1) : '—'}</td>
    </tr>`;
  }).join('');

  // ── SECTION 2: Product summary ────────────────────────────────────────────
  // Aggregate all products across filtered orders
  // Convert to gallons for water calculation using spray rate from template
  const prodMap = {};
  let totalWater = 0;

  filtered.forEach(o => {
    const tmpl = DB.templates.find(t => t.TemplateID === o.TemplateUsed);
    const sprayRate = parseFloat(tmpl?.SprayRate || AppSettings.defaultSprayRate || 2);
    const ac = parseFloat(o.TotalAcres || 0);
    totalWater += ac * sprayRate;

    DB.orderProds.filter(p => p.OrderID === o.OrderID).forEach(p => {
      const key = p.ProductID || p.ProductName;
      if (!prodMap[key]) {
        prodMap[key] = { name: p.ProductName, unit: p.Unit, total: 0, costPerUnit: parseFloat(p.CostPerUnit||0) };
      }
      prodMap[key].total += parseFloat(p.TotalUnitsNeeded || 0);
    });
  });

  const prodRows = Object.values(prodMap).map(p => {
    const disp = fmtAmt(p.total, p.unit);
    return `<tr>
      <td>${p.name}</td>
      <td style="text-align:right;font-family:var(--font-mono)">${disp}</td>
      <td style="color:var(--text-sub)">${p.unit}/ac rate</td>
    </tr>`;
  }).join('');

  const waterDisp = fmtAmt(totalWater, 'gallon');

  // ── SECTION 3: Map of all fields ─────────────────────────────────────────
  // Collect all field IDs for the filtered orders
  const fieldIds = new Set();
  filtered.forEach(o => DB.orderFields.filter(f => f.OrderID === o.OrderID).forEach(f => fieldIds.add(f.FieldID)));
  const mapFields = [...fieldIds].map(id => DB.fields.find(f => f.FieldID === id)).filter(Boolean);
  const hasMapData = mapFields.some(f => f.PolygonKML || (f.CentroidLat && f.CentroidLng));

  el.innerHTML = `
    <div class="report-summary-header">
      <strong>${label}</strong>
      ${pilotFilter ? ' · ' + (DB.pilots.find(p=>p.PilotID===pilotFilter)?.Name||'') : ''}
      ${statusFilter ? ' · ' + statusFilter : ''}
      · ${filtered.length} order${filtered.length!==1?'s':''} · ${totalAcres.toFixed(1)} ac total
    </div>

    <div class="report-section">
      <div class="report-section-title">Fields to Apply</div>
      <div style="overflow-x:auto">
        <table class="report-table">
          <thead><tr>
            <th>Date</th><th>Order</th><th>Customer</th><th>Fields</th>
            <th>Crop</th><th>Pilot</th><th>Status</th><th style="text-align:right">Acres</th>
          </tr></thead>
          <tbody>${fieldRows}</tbody>
          <tfoot><tr>
            <td colspan="7" style="font-weight:600;padding-top:0.5rem">Total</td>
            <td style="text-align:right;font-weight:600">${totalAcres.toFixed(1)} ac</td>
          </tr></tfoot>
        </table>
      </div>
    </div>

    <div class="report-section">
      <div class="report-section-title">Products Needed</div>
      <div style="overflow-x:auto">
        <table class="report-table">
          <thead><tr><th>Product</th><th style="text-align:right">Total Needed</th><th>Note</th></tr></thead>
          <tbody>${prodRows}</tbody>
          <tfoot><tr>
            <td style="font-weight:600">Water (spray solution)</td>
            <td style="text-align:right;font-weight:600">${waterDisp}</td>
            <td style="color:var(--text-sub)">at ${AppSettings.defaultSprayRate} gal/ac avg</td>
          </tr></tfoot>
        </table>
      </div>
    </div>

    ${hasMapData ? `
    <div class="report-section">
      <div class="report-section-title">Field Map</div>
      <div id="reportMapContainer" class="report-map-container"></div>
    </div>` : ''}
  `;

  // Render the field map if we have data
  if (hasMapData) {
    setTimeout(() => renderReportMap(mapFields, filtered), 100);
  }
}

function renderReportMap(mapFields, orders) {
  const container = document.getElementById('reportMapContainer');
  if (!container) return;

  // Use Leaflet for the report map
  if (typeof L === 'undefined') { container.innerHTML = '<div class="empty-state">Map not available</div>'; return; }

  // Clear any existing map
  if (window._reportMap) { window._reportMap.remove(); window._reportMap = null; }

  window._reportMap = L.map(container, { zoomControl: true });
  L.tileLayer(
    'https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}',
    { attribution: 'Imagery © Esri', maxZoom: 19 }
  ).addTo(window._reportMap);
  L.tileLayer(
    'https://server.arcgisonline.com/ArcGIS/rest/services/Reference/World_Boundaries_and_Places/MapServer/tile/{z}/{y}/{x}',
    { maxZoom: 19, opacity: 0.7 }
  ).addTo(window._reportMap);

  const bounds = [];
  mapFields.forEach(f => {
    const order = orders.find(o => DB.orderFields.some(of => of.OrderID === o.OrderID && of.FieldID === f.FieldID));
    const color = order?.Status === 'Completed' ? '#81C784' : order?.Status === 'Scheduled' ? '#4FC3F7' : '#FFB74D';

    if (f.PolygonKML) {
      const pts = GeoUtils.parsePolygon(f.PolygonKML);
      if (pts && pts.length >= 3) {
        const latlngs = pts.map(p => [p.lat, p.lng]);
        L.polygon(latlngs, { color, weight: 2, fillColor: color, fillOpacity: 0.3 })
          .bindTooltip(`${f.FieldName} · ${f.Acres || '?'} ac`, { permanent: false })
          .addTo(window._reportMap);
        latlngs.forEach(ll => bounds.push(ll));
      }
    } else if (f.CentroidLat && f.CentroidLng) {
      const ll = [parseFloat(f.CentroidLat), parseFloat(f.CentroidLng)];
      L.circleMarker(ll, { color, radius: 8, fillColor: color, fillOpacity: 0.6 })
        .bindTooltip(`${f.FieldName} · ${f.Acres || '?'} ac`)
        .addTo(window._reportMap);
      bounds.push(ll);
    }
  });

  if (bounds.length) {
    window._reportMap.fitBounds(bounds, { padding: [30, 30] });
  }
}


// ── ORDER MODAL ──────────────────────────────────────────────────────────────
function showNewOrderModal() {
  orderFieldsSelected = [];
  document.getElementById('editOrderId').value = '';
  document.getElementById('modalTitle').textContent = 'New Order';

  ['fNotes','fAttachments','fRatePerAcre','fPlantingDate','fRelativeMaturity','fScheduledDate'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.value = '';
  });
  document.getElementById('fStatus').value = 'Open';
  document.getElementById('fPricingType').value = 'Flat Rate';
  document.getElementById('fCropType').value = 'Corn';
  document.getElementById('fInvoiced').value = 'No';
  document.getElementById('chemicalLines').innerHTML = '';
  document.getElementById('fTemplate').value = '';

  renderOrderFieldsList();
  populateModalDropdowns();
  openModal('orderModal');
}

function editCurrentOrder() {
  const o = DB.orders.find(x => x.OrderID === currentOrderId);
  if (!o) return;
  populateModalDropdowns();

  document.getElementById('editOrderId').value = o.OrderID;
  document.getElementById('modalTitle').textContent = 'Edit Order — ' + o.OrderID;
  document.getElementById('fNotes').value = o.Notes || '';
  document.getElementById('fAttachments').value = o.Attachments || '';
  document.getElementById('fRatePerAcre').value = o.RatePerAcre || '';
  document.getElementById('fStatus').value = o.Status || 'Open';
  document.getElementById('fPricingType').value = o.PricingType || 'Flat Rate';
  document.getElementById('fCropType').value = o.CropType || 'Corn';
  document.getElementById('fInvoiced').value = o.Invoiced || 'No';
  document.getElementById('fPlantingDate').value = o.PlantingDate || '';
  document.getElementById('fRelativeMaturity').value = o.RelativeMaturity || '';
  document.getElementById('fScheduledDate').value = o.ScheduledDate || '';

  // Load selected fields from orderFields
  orderFieldsSelected = DB.orderFields
    .filter(f => f.OrderID === o.OrderID)
    .map(f => DB.fields.find(x => x.FieldID === f.FieldID) || {
      FieldID: f.FieldID, FieldName: f.FieldName, Acres: f.Acres
    });
  renderOrderFieldsList();

  // Set customer/pilot selects after DOM settles
  setTimeout(() => {
    document.getElementById('fCustomer').value = o.CustomerID;
    document.getElementById('fPilot').value = o.PilotID;
  }, 50);

  // Load chemical lines
  const prods = DB.orderProds.filter(p => p.OrderID === o.OrderID);
  document.getElementById('chemicalLines').innerHTML = '';
  prods.forEach(p => addChemicalLine(p));

  openModal('orderModal');
}

function populateModalDropdowns() {
  // Customers
  const custSel = document.getElementById('fCustomer');
  custSel.innerHTML = '<option value="">— Select customer —</option>' +
    DB.customers.map(c => `<option value="${c.CustomerID}">${c.Name}</option>`).join('');

  // Pilots
  const pilotSel = document.getElementById('fPilot');
  pilotSel.innerHTML = '<option value="">— Select pilot —</option>' +
    DB.pilots.filter(p => p.Active === 'Yes').map(p => `<option value="${p.PilotID}">${p.Name}</option>`).join('');

  // Templates
  const tmplSel = document.getElementById('fTemplate');
  tmplSel.innerHTML = '<option value="">— Select template —</option>' +
    DB.templates.filter(t => t.Active === 'Yes').map(t => `<option value="${t.TemplateID}">${t.TemplateName} (${t.CropType})</option>`).join('');
}

function addChemicalLine(prefill) {
  const container = document.getElementById('chemicalLines');
  const div = document.createElement('div');
  div.className = 'chem-line';
  const prodOptions = DB.products.map(p =>
    `<option value="${p.ProductID}" data-unit="${p.Unit}" data-cost="${p.CostPerUnit}" ${prefill?.ProductID === p.ProductID ? 'selected' : ''}>${p.ProductName}</option>`
  ).join('');

  div.innerHTML = `
    <div class="form-group" style="margin:0">
      <label class="form-label">Product</label>
      <select class="form-input chem-product">${prodOptions}</select>
    </div>
    <div class="form-group" style="margin:0">
      <label class="form-label">Rate/Ac</label>
      <input class="form-input chem-rate" type="number" step="0.01" value="${prefill?.RatePerAcre||''}" placeholder="0">
    </div>
    <div class="form-group" style="margin:0">
      <label class="form-label">Unit</label>
      <input class="form-input chem-unit" type="text" value="${prefill?.Unit||'fl oz'}" placeholder="fl oz">
    </div>
    <div class="form-group" style="margin:0">
      <label class="form-label">By</label>
      <select class="form-input chem-supplied">
        <option ${prefill?.SuppliedBy==='Me'?'selected':''}>Me</option>
        <option ${prefill?.SuppliedBy==='Customer'?'selected':''}>Customer</option>
      </select>
    </div>
    <button class="chem-line-remove" onclick="this.parentElement.remove()">×</button>
  `;
  container.appendChild(div);
}

function applyTemplate() {
  const tmplId = document.getElementById('fTemplate').value;
  if (!tmplId) return;
  const prods = DB.templateProds.filter(p => p.TemplateID === tmplId);
  document.getElementById('chemicalLines').innerHTML = '';
  prods.forEach(p => addChemicalLine(p));
  showToast(`Applied ${prods.length} chemicals from template`, 'success');
}

async function saveOrder() {
  const custId  = document.getElementById('fCustomer').value;
  const custName = document.getElementById('fCustomer').selectedOptions[0]?.text || '';
  const editId  = document.getElementById('editOrderId').value;
  const orderId = editId || 'ORD-' + String(DB.orders.length + 1).padStart(3, '0');

  if (!custId) { showToast('Please select a customer', 'error'); return; }
  if (orderFieldsSelected.length === 0) { showToast('Please add at least one field', 'error'); return; }

  const totalAcres = orderFieldsSelected.reduce((s, f) => s + parseFloat(f.Acres || 0), 0);
  const ratePerAcre = parseFloat(document.getElementById('fRatePerAcre').value) || 0;
  const pricingType = document.getElementById('fPricingType').value;

  // Collect chemical lines
  const lines = [];
  let chemCost = 0;
  document.querySelectorAll('.chem-line').forEach((row, i) => {
    const prodSel  = row.querySelector('.chem-product');
    const rateFld  = row.querySelector('.chem-rate');
    const unitFld  = row.querySelector('.chem-unit');
    const suppFld  = row.querySelector('.chem-supplied');
    if (!prodSel || !prodSel.value) return;
    const prod = DB.products.find(p => p.ProductID === prodSel.value) || {};
    const rate = parseFloat(rateFld?.value) || 0;
    const cost = parseFloat(prod.CostPerUnit || 0);
    const totalUnits = rate * totalAcres;
    const lineCost   = totalUnits * cost;
    chemCost += lineCost;
    lines.push({
      LineID: (editId ? DB.orderProds.find(l => l.OrderID === editId && l.ProductID === prodSel.value)?.LineID : null) || `OPL-${Date.now()}-${i}`,
      OrderID: orderId, ProductID: prodSel.value,
      ProductName: prod.ProductName || prodSel.value,
      RatePerAcre: rate, Unit: unitFld?.value || prod.Unit || '',
      SuppliedBy: suppFld?.value || 'Me',
      CostPerUnit: cost, Acres: totalAcres,
      TotalUnitsNeeded: totalUnits, TotalProductCost: lineCost
    });
  });

  const estTotal = pricingType === 'Flat Rate + Chemical'
    ? ratePerAcre * totalAcres + chemCost
    : ratePerAcre * totalAcres;

  const order = {
    OrderID: orderId,
    OrderDate: editId ? toDateStr(DB.orders.find(o => o.OrderID === editId)?.OrderDate || new Date().toISOString().split('T')[0]) : new Date().toISOString().split('T')[0],
    CustomerID: custId,
    CustomerName: custName,
    CropType: document.getElementById('fCropType').value,
    PlantingDate: document.getElementById('fPlantingDate').value,
    RelativeMaturity: document.getElementById('fRelativeMaturity').value,
    ScheduledDate: document.getElementById('fScheduledDate').value,
    CompletedDate: editId ? (DB.orders.find(o => o.OrderID === editId)?.CompletedDate || '') : '',
    PilotID: document.getElementById('fPilot').value,
    PilotName: document.getElementById('fPilot').selectedOptions[0]?.text || '',
    Status: document.getElementById('fStatus').value,
    PricingType: pricingType,
    RatePerAcre: ratePerAcre,
    TotalAcres: totalAcres,
    EstimatedTotal: estTotal,
    ChemicalCost: chemCost,
    TemplateUsed: document.getElementById('fTemplate').value,
    Invoiced: document.getElementById('fInvoiced').value,
    DJI_FlightFile: '',
    Attachments: document.getElementById('fAttachments').value,
    Notes: document.getElementById('fNotes').value
  };

  // Upsert order
  const existIdx = DB.orders.findIndex(o => o.OrderID === orderId);
  if (existIdx > -1) DB.orders[existIdx] = order; else DB.orders.push(order);
  await writeRow('orders', order);

  // Save orderFields lines
  if (editId) {
    // Remove old orderFields for this order
    const old = DB.orderFields.filter(f => f.OrderID === orderId);
    for (const f of old) await writeRow('orderFields', { LineID: f.LineID, _delete: true });
    DB.orderFields = DB.orderFields.filter(f => f.OrderID !== orderId);
  }
  for (let i = 0; i < orderFieldsSelected.length; i++) {
    const f = orderFieldsSelected[i];
    const line = { LineID: `OFL-${orderId}-${i}`, OrderID: orderId, FieldID: f.FieldID, FieldName: f.FieldName, CustomerID: custId, Acres: f.Acres, Notes: '' };
    DB.orderFields.push(line);
    await writeRow('orderFields', line);
  }

  // Save/delete chemical lines
  if (editId) {
    const oldProds = DB.orderProds.filter(p => p.OrderID === orderId);
    for (const p of oldProds) await writeRow('orderProds', { LineID: p.LineID, _delete: true });
    DB.orderProds = DB.orderProds.filter(p => p.OrderID !== orderId);
  }
  for (const line of lines) {
    DB.orderProds.push(line);
    await writeRow('orderProds', line);
  }

  saveToLocalStorage();
  closeModal();
  renderView(currentView);
  showToast(editId ? 'Order updated' : 'Order created', 'success');
}

// ── WRITE TO SHEET ───────────────────────────────────────────────────────────
async function writeRow(table, data) {
  // Use GET for small payloads, POST (text/plain) for large ones (e.g. KML polygons).
  // GAS accepts both via doGet and doPost reading e.postData.contents.
  try {
    const body    = JSON.stringify(data);
    const encoded = encodeURIComponent(body);
    const getUrl  = `${GAS_URL}?action=write&table=${table}&data=${encoded}`;

    let res;
    if (getUrl.length > 4000) {
      // Large payload — use POST with text/plain to avoid CORS preflight
      res = await fetch(GAS_URL, {
        method:  'POST',
        mode:    'no-cors',
        headers: { 'Content-Type': 'text/plain' },
        body:    JSON.stringify({ action: 'write', table, data }),
      });
      // no-cors returns opaque response — assume success
      console.log('Write POST (large):', table, data[Object.keys(data)[0]]);
      return;
    }

    res = await fetch(getUrl);
    const result = await res.json();
    if (result.error) console.warn('Write error:', result.error);
    else console.log('Write OK:', result.action, table);
  } catch(e) {
    console.warn('Write failed:', e.message);
  }
}

// ── HELPERS ──────────────────────────────────────────────────────────────────
function nextId(prefix, existing) {
  // Use timestamp suffix to guarantee uniqueness across sessions and devices
  // Fall back to incrementing only if we need a readable ID (e.g. CUST-, PLT-)
  const nums = existing.map(id => parseInt(id.replace(prefix + '-', '') || 0)).filter(Boolean);
  const max  = nums.length ? Math.max(...nums) : 0;
  const next = max + 1;
  // Check if next sequential ID already exists (race condition guard)
  const candidate = `${prefix}-${String(next).padStart(3, '0')}`;
  if (existing.includes(candidate)) {
    // Fallback: use timestamp to guarantee uniqueness
    return `${prefix}-${Date.now().toString(36).toUpperCase().slice(-4)}`;
  }
  return candidate;
}

function fmtDate(dateStr) {
  if (!dateStr || dateStr === '—') return '—';
  try {
    const s = String(dateStr).trim();
    let d;
    // Google Sheets date serial (number of days since Dec 30 1899)
    if (/^\d{4,5}(\.\d+)?$/.test(s)) {
      d = new Date((parseFloat(s) - 25569) * 86400 * 1000);
    }
    // Full ISO timestamp — strip time component
    else if (s.includes('T')) {
      d = new Date(s.split('T')[0] + 'T12:00:00');
    }
    // Plain YYYY-MM-DD
    else if (/^\d{4}-\d{2}-\d{2}$/.test(s)) {
      d = new Date(s + 'T12:00:00');
    }
    // Anything else — try direct parse
    else {
      d = new Date(s);
    }
    if (isNaN(d.getTime())) return dateStr;
    return d.toLocaleDateString('en-US', { month:'short', day:'numeric', year:'numeric' });
  } catch { return String(dateStr); }
}

function toDateStr(val) {
  // Normalize any date value to YYYY-MM-DD string for storage
  if (!val) return '';
  const s = String(val).trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;          // already clean
  if (s.includes('T')) return s.split('T')[0];            // strip timestamp
  if (/^\d{4,5}(\.\d+)?$/.test(s)) {                      // Sheets serial
    const d = new Date((parseFloat(s) - 25569) * 86400 * 1000);
    return d.toISOString().split('T')[0];
  }
  const d = new Date(s);
  if (!isNaN(d.getTime())) return d.toISOString().split('T')[0];
  return s;
}

function openModal(id) {
  document.getElementById('modalOverlay').classList.add('active');
  document.getElementById(id).classList.add('active');
}

function closeModal() {
  document.getElementById('modalOverlay').classList.remove('active');
  document.querySelectorAll('.modal').forEach(m => m.classList.remove('active'));
}

function toggleSidebar(force) {
  const sb = document.getElementById('sidebar');
  if (force === false) { sb.classList.remove('open'); return; }
  sb.classList.toggle('open');
}

function showToast(msg, type) {
  const t = document.getElementById('toast');
  t.textContent = msg;
  t.className = 'toast show ' + (type || '');
  setTimeout(() => { t.className = 'toast'; }, 3000);
}
