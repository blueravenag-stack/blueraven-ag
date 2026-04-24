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

    DB.orders       = parseSheet(data.orders,       orderHeaders());
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

function parseSheet(rows, headers) {
  if (!rows || rows.length < 2) return [];
  return rows.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i] || ''; });
    return obj;
  }).filter(r => r[headers[0]]);
}

// Header definitions match sheet columns exactly
function orderHeaders()       { return ['OrderID','OrderDate','CustomerID','CustomerName','CropType','PlantingDate','RelativeMaturity','ScheduledDate','CompletedDate','PilotID','PilotName','Status','PricingType','RatePerAcre','TotalAcres','EstimatedTotal','ChemicalCost','TemplateUsed','Invoiced','DJI_FlightFile','Attachments','Notes']; }
function customerHeaders()    { return ['CustomerID','Name','Phone','Email','Address','City','State','Zip','Notes']; }
function pilotHeaders()       { return ['PilotID','Name','Phone','Email','FAA_Part107_Num','Active']; }
function productHeaders()     { return ['ProductID','ProductName','Manufacturer','Unit','CostPerUnit','REI_Hours','PHI_Days','Notes']; }
function templateHeaders()    { return ['TemplateID','TemplateName','CropType','Description','Active']; }
function templateProdHeaders(){ return ['LineID','TemplateID','ProductID','ProductName','RatePerAcre','Unit','SuppliedBy','Notes']; }
function orderProdHeaders()   { return ['LineID','OrderID','ProductID','ProductName','RatePerAcre','Unit','SuppliedBy','CostPerUnit','Acres','TotalUnitsNeeded','TotalProductCost']; }
function fieldHeaders()       { return ['FieldID','CustomerID','CustomerName','FieldName','Acres','CentroidLat','CentroidLng','PolygonKML','CLU_TractID','CLU_FarmNum','Active','Notes']; }
function orderFieldHeaders()  { return ['LineID','OrderID','FieldID','FieldName','CustomerID','Acres','Notes']; }

// ── LOCAL STORAGE ───────────────────────────────────────────────────────────
function saveToLocalStorage() {
  try { localStorage.setItem('blueraven_db', JSON.stringify(DB)); } catch(e) {}
}
function loadFromLocalStorage() {
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
      o.FieldName.toLowerCase().includes(search) ||
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
  const acres = o.Acres ? `${parseFloat(o.Acres).toLocaleString()} ac` : '';
  const date  = o.ScheduledDate ? fmtDate(o.ScheduledDate) : '';
  return `
  <div class="order-card" onclick="viewOrder('${o.OrderID}')">
    <span class="order-id">${o.OrderID}</span>
    <div class="order-main">
      <div class="order-customer">${o.CustomerName}</div>
      <div class="order-field">${o.FieldName} · ${o.CropType}</div>
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
  const estTotal = o.EstimatedTotal ? '$' + parseFloat(o.EstimatedTotal).toLocaleString('en-US', {minimumFractionDigits:2}) : '—';
  const chemCost = o.ChemicalCost   ? '$' + parseFloat(o.ChemicalCost).toLocaleString('en-US', {minimumFractionDigits:2}) : '—';

  let chemTable = '';
  if (prods.length) {
    chemTable = `<table class="chem-table">
      <thead><tr><th>Product</th><th>Rate/Ac</th><th>Unit</th><th>Supplied By</th><th>Total Qty</th><th>Cost</th></tr></thead>
      <tbody>${prods.map(p => `
        <tr>
          <td>${p.ProductName}</td>
          <td>${p.RatePerAcre}</td>
          <td>${p.Unit}</td>
          <td>${p.SuppliedBy}</td>
          <td>${parseFloat(p.TotalUnitsNeeded||0).toFixed(1)} ${p.Unit}</td>
          <td>${p.TotalProductCost ? '$'+parseFloat(p.TotalProductCost).toFixed(2) : '—'}</td>
        </tr>`).join('')}
      </tbody></table>`;
  } else {
    chemTable = '<div style="color:var(--text-sub);font-size:0.85rem;padding:0.5rem 0">No chemicals recorded</div>';
  }

  let mapHTML = '';
  if (o.FieldLat && o.FieldLng) {
    const mapsUrl = `https://www.google.com/maps?q=${o.FieldLat},${o.FieldLng}`;
    const staticUrl = `https://maps.googleapis.com/maps/api/staticmap?center=${o.FieldLat},${o.FieldLng}&zoom=13&size=600x200&maptype=satellite&markers=${o.FieldLat},${o.FieldLng}&key=NO_KEY`;
    mapHTML = `<div class="map-container">
      <div style="text-align:center">
        <div style="margin-bottom:0.5rem">📍 ${o.FieldLat}, ${o.FieldLng}</div>
        <a class="map-link" href="${mapsUrl}" target="_blank">Open in Google Maps →</a>
      </div>
    </div>`;
  }

  const attachLink = o.Attachments
    ? `<a class="map-link" href="${o.Attachments}" target="_blank">📎 Open Attachments Folder →</a>`
    : '<span style="color:var(--text-sub);font-size:0.82rem">No attachments linked</span>';

  document.getElementById('orderDetailContent').innerHTML = `
    <div class="detail-header">
      <div class="detail-order-id">${o.OrderID}</div>
      <div class="detail-title">${o.CustomerName} — ${o.FieldName}</div>
      <div class="detail-badges">
        ${statusBadge(o.Status)}
        ${o.Invoiced === 'Yes' ? '<span class="badge badge-invoiced">Invoiced</span>' : ''}
      </div>
    </div>

    ${mapHTML}

    <div class="detail-grid">
      <div class="detail-card">
        <div class="detail-card-title">Job Details</div>
        <div class="detail-row"><span class="detail-key">Crop</span><span class="detail-val">${o.CropType}</span></div>
        <div class="detail-row"><span class="detail-key">Acres</span><span class="detail-val">${parseFloat(o.Acres||0).toLocaleString()} ac</span></div>
        <div class="detail-row"><span class="detail-key">Scheduled</span><span class="detail-val">${fmtDate(o.ScheduledDate)}</span></div>
        <div class="detail-row"><span class="detail-key">Completed</span><span class="detail-val">${fmtDate(o.CompletedDate)}</span></div>
        <div class="detail-row"><span class="detail-key">Pilot</span><span class="detail-val">${o.PilotName}</span></div>
        <div class="detail-row"><span class="detail-key">Notes</span><span class="detail-val" style="max-width:200px;text-align:right">${o.Notes||'—'}</span></div>
      </div>
      <div class="detail-card">
        <div class="detail-card-title">Pricing</div>
        <div class="detail-row"><span class="detail-key">Type</span><span class="detail-val">${o.PricingType}</span></div>
        <div class="detail-row"><span class="detail-key">Rate/Acre</span><span class="detail-val">$${parseFloat(o.RatePerAcre||0).toFixed(2)}</span></div>
        <div class="detail-row"><span class="detail-key">Est. Total</span><span class="detail-val">${estTotal}</span></div>
        <div class="detail-row"><span class="detail-key">Chem Cost</span><span class="detail-val">${chemCost}</span></div>
        <div class="detail-row"><span class="detail-key">Chem Supplied</span><span class="detail-val">${o.ChemicalSuppliedBy||o.SuppliedBy||'—'}</span></div>
        <div class="detail-row"><span class="detail-key">Invoiced</span><span class="detail-val">${o.Invoiced}</span></div>
      </div>
    </div>

    <div class="detail-card" style="margin-bottom:1.25rem">
      <div class="detail-card-title">Chemical Mix</div>
      ${chemTable}
    </div>

    <div class="detail-card">
      <div class="detail-card-title">Attachments</div>
      ${attachLink}
    </div>
  `;

  // Show/hide action buttons
  document.getElementById('btnMarkComplete').style.display = o.Status !== 'Completed' ? '' : 'none';
  document.getElementById('btnMarkInvoiced').style.display = o.Invoiced !== 'Yes' ? '' : 'none';

  navigateTo('order-detail');
}

async function markComplete() {
  const o = DB.orders.find(x => x.OrderID === currentOrderId);
  if (!o) return;
  o.Status = 'Completed';
  o.CompletedDate = new Date().toISOString().split('T')[0];
  await writeRow('orders', o);
  viewOrder(currentOrderId);
  showToast('Order marked complete', 'success');
}

async function markInvoiced() {
  const o = DB.orders.find(x => x.OrderID === currentOrderId);
  if (!o) return;
  o.Invoiced = 'Yes';
  await writeRow('orders', o);
  viewOrder(currentOrderId);
  showToast('Order marked invoiced', 'success');
}

// ── CUSTOMERS ────────────────────────────────────────────────────────────────
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
      const totalAcres = orders.reduce((s, o) => s + parseFloat(o.Acres||0), 0);
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
  // fFieldName moved to Fields tab
  document.getElementById('fFieldAcres').value = '';
  document.getElementById('fFieldLat').value = '';
  document.getElementById('fFieldLng').value = '';
  document.getElementById('fFieldNotes').value = '';
  document.getElementById('fFieldActive').value = 'Yes';
  document.getElementById('btnDeleteField').style.display = 'none';

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

  openModal('fieldModal');
}

async function saveField() {
  const custId   = document.getElementById('fFieldCustomer').value;
  const custName = document.getElementById('fFieldCustomer').selectedOptions[0]?.text || '';
  const name     = document.getElementById('fFieldName').value.trim();
  const editId   = document.getElementById('editFieldId').value;

  if (!custId) { showToast('Please select a customer', 'error'); return; }
  if (!name)   { showToast('Field name is required', 'error'); return; }

  const fieldId = editId || 'FLD-' + String(DB.fields.length + 1).padStart(3, '0');
  const field = {
    FieldID: fieldId, CustomerID: custId, CustomerName: custName,
    FieldName: name,
    Acres: document.getElementById('fFieldAcres').value || '',
    CentroidLat: document.getElementById('fFieldLat').value || '',
    CentroidLng: document.getElementById('fFieldLng').value || '',
    PolygonKML: editId ? (DB.fields.find(f => f.FieldID === editId)?.PolygonKML || '') : '',
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

// ── BATCH CALCULATOR ─────────────────────────────────────────────────────────
function renderCalculator() {
  const eligible = DB.orders.filter(o => {
    const prods = DB.orderProds.filter(p => p.OrderID === o.OrderID);
    return prods.length > 0;
  });

  document.getElementById('calcOrderList').innerHTML = eligible.length ?
    eligible.map(o => `
      <div class="calc-order-item" onclick="toggleCalcOrder('${o.OrderID}', this)">
        <input type="checkbox" id="calc_${o.OrderID}" ${selectedCalcOrders.has(o.OrderID) ? 'checked' : ''}>
        <div>
          <div style="font-weight:500">${o.OrderID} — ${o.CustomerName}</div>
          <div style="font-size:0.75rem;color:var(--text-sub)">${o.FieldName} · ${o.Acres} ac · ${o.CropType}</div>
        </div>
      </div>`).join('') :
    '<div style="color:var(--text-sub);font-size:0.82rem">No orders with chemicals found</div>';

  runCalc();
}

function toggleCalcOrder(orderId, el) {
  const cb = document.getElementById('calc_' + orderId);
  if (selectedCalcOrders.has(orderId)) {
    selectedCalcOrders.delete(orderId);
    cb.checked = false;
  } else {
    selectedCalcOrders.add(orderId);
    cb.checked = true;
  }
  runCalc();
}

function runCalc() {
  if (selectedCalcOrders.size === 0) {
    document.getElementById('calcResults').innerHTML = '<div class="calc-empty">Select orders to see chemical requirements</div>';
    return;
  }

  const tankSize = parseFloat(document.getElementById('tankSize')?.value || 100);
  const totalAcres = [...selectedCalcOrders].reduce((sum, id) => {
    const o = DB.orders.find(x => x.OrderID === id);
    return sum + parseFloat(o?.Acres || 0);
  }, 0);

  // Aggregate chemicals across selected orders
  const chemMap = {};
  [...selectedCalcOrders].forEach(orderId => {
    DB.orderProds.filter(p => p.OrderID === orderId).forEach(p => {
      const key = p.ProductID || p.ProductName;
      if (!chemMap[key]) {
        chemMap[key] = { name: p.ProductName, unit: p.Unit, total: 0, costPerUnit: parseFloat(p.CostPerUnit||0) };
      }
      chemMap[key].total += parseFloat(p.TotalUnitsNeeded || 0);
    });
  });

  const chems = Object.values(chemMap);
  if (!chems.length) {
    document.getElementById('calcResults').innerHTML = '<div class="calc-empty">No chemicals found for selected orders</div>';
    return;
  }

  // Tank loads based on total gallons needed (assume ~15 gal/acre spray rate as default)
  const sprayRate = 15;
  const totalGallons = totalAcres * sprayRate;
  const loadsNeeded = tankSize > 0 ? Math.ceil(totalGallons / tankSize) : '—';

  document.getElementById('calcResults').innerHTML = `
    <div style="margin-bottom:1rem;padding-bottom:1rem;border-bottom:1px solid var(--border)">
      <div style="font-size:0.75rem;color:var(--text-sub);text-transform:uppercase;letter-spacing:0.08em">Total Acres Selected</div>
      <div style="font-family:var(--font-head);font-size:2rem;color:var(--text)">${totalAcres.toLocaleString()} ac</div>
    </div>
    ${chems.map(c => `
      <div class="calc-result-item">
        <div class="calc-product-name">${c.name}</div>
        <div class="calc-product-qty">${c.total.toFixed(1)} ${c.unit}</div>
        <div class="calc-product-detail">
          <span>Total needed</span>
          <span>${c.costPerUnit > 0 ? '$' + (c.total * c.costPerUnit).toFixed(2) + ' est.' : ''}</span>
        </div>
      </div>`).join('')}
    <div class="calc-loads">
      <div class="calc-loads-num">${loadsNeeded}</div>
      <div class="calc-loads-label">tank loads (${tankSize} gal tank · ${sprayRate} gal/ac)</div>
    </div>
  `;
}

// ── REPORTS ──────────────────────────────────────────────────────────────────
function renderReports() {
  // Acres by crop
  const acresByCrop = {};
  DB.orders.filter(o => o.Status === 'Completed').forEach(o => {
    const crop = o.CropType || 'Unknown';
    acresByCrop[crop] = (acresByCrop[crop] || 0) + parseFloat(o.Acres || 0);
  });
  const totalAcres = Object.values(acresByCrop).reduce((s, v) => s + v, 0);

  document.getElementById('reportAcres').innerHTML = Object.keys(acresByCrop).length ?
    Object.entries(acresByCrop).sort((a,b) => b[1]-a[1]).map(([crop, acres]) =>
      `<div class="report-row"><span>${crop}</span><span class="report-val">${acres.toLocaleString()} ac</span></div>`
    ).join('') +
    `<div class="report-total"><span>Total</span><span>${totalAcres.toLocaleString()} ac</span></div>` :
    '<div style="color:var(--text-sub);font-size:0.85rem;padding:1rem 0">No completed orders yet</div>';

  // Product usage
  const prodUsage = {};
  DB.orderProds.forEach(p => {
    const name = p.ProductName;
    if (!prodUsage[name]) prodUsage[name] = { total: 0, unit: p.Unit };
    prodUsage[name].total += parseFloat(p.TotalUnitsNeeded || 0);
  });

  document.getElementById('reportProducts').innerHTML = Object.keys(prodUsage).length ?
    Object.entries(prodUsage).sort((a,b) => b[1].total - a[1].total).map(([name, d]) =>
      `<div class="report-row"><span>${name}</span><span class="report-val">${d.total.toFixed(1)} ${d.unit}</span></div>`
    ).join('') :
    '<div style="color:var(--text-sub);font-size:0.85rem;padding:1rem 0">No product data yet</div>';

  // Revenue summary
  const totalRevenue = DB.orders.reduce((s, o) => s + parseFloat(o.EstimatedTotal || 0), 0);
  const totalChemCost = DB.orders.reduce((s, o) => s + parseFloat(o.ChemicalCost || 0), 0);
  const gross = totalRevenue - totalChemCost;

  document.getElementById('reportRevenue').innerHTML = `
    <div class="report-row"><span>Estimated Revenue</span><span class="report-val">$${totalRevenue.toLocaleString('en-US',{minimumFractionDigits:2})}</span></div>
    <div class="report-row"><span>Chemical Cost</span><span class="report-val">$${totalChemCost.toLocaleString('en-US',{minimumFractionDigits:2})}</span></div>
    <div class="report-total"><span>Gross Margin Est.</span><span>$${gross.toLocaleString('en-US',{minimumFractionDigits:2})}</span></div>
    <div style="color:var(--text-sub);font-size:0.75rem;margin-top:0.75rem">Based on all orders. DJI flight data import coming soon.</div>
  `;
}

// ── ORDER MODAL ──────────────────────────────────────────────────────────────
function showNewOrderModal() {
  document.getElementById('editOrderId').value = '';
  document.getElementById('modalTitle').textContent = 'New Order';

  // Reset fields
  ['fFieldName','fAcres','fLat','fLng','fPolygon','fNotes','fAttachments','fRatePerAcre'].forEach(id => {
    document.getElementById(id).value = '';
  });
  document.getElementById('fStatus').value = 'Open';
  document.getElementById('fPricingType').value = 'Flat Rate';
  document.getElementById('fCropType').value = 'Corn';
  document.getElementById('fInvoiced').value = 'No';
  document.getElementById('fChemSupplied').value = 'Me';
  document.getElementById('chemicalLines').innerHTML = '';
  document.getElementById('fTemplate').value = '';

  populateModalDropdowns();
  openModal('orderModal');
}

function editCurrentOrder() {
  const o = DB.orders.find(x => x.OrderID === currentOrderId);
  if (!o) return;
  populateModalDropdowns();

  document.getElementById('editOrderId').value = o.OrderID;
  document.getElementById('modalTitle').textContent = 'Edit Order — ' + o.OrderID;
  document.getElementById('fFieldName').value = o.FieldName;
  document.getElementById('fAcres').value = o.Acres;
  document.getElementById('fLat').value = o.FieldLat;
  document.getElementById('fLng').value = o.FieldLng;
  document.getElementById('fPolygon').value = o.FieldPolygonKML;
  document.getElementById('fNotes').value = o.Notes;
  document.getElementById('fAttachments').value = o.Attachments;
  document.getElementById('fRatePerAcre').value = o.RatePerAcre;
  document.getElementById('fStatus').value = o.Status;
  document.getElementById('fPricingType').value = o.PricingType;
  document.getElementById('fCropType').value = o.CropType;
  document.getElementById('fInvoiced').value = o.Invoiced;

  // Set customer/pilot selects
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
      <select class="form-input">${prodOptions}</select>
    </div>
    <div class="form-group" style="margin:0">
      <label class="form-label">Rate/Ac</label>
      <input class="form-input" type="number" step="0.01" value="${prefill?.RatePerAcre||''}" placeholder="0">
    </div>
    <div class="form-group" style="margin:0">
      <label class="form-label">Unit</label>
      <input class="form-input" type="text" value="${prefill?.Unit||'fl oz'}" placeholder="fl oz">
    </div>
    <div class="form-group" style="margin:0">
      <label class="form-label">By</label>
      <select class="form-input">
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
    OrderDate: editId ? (DB.orders.find(o => o.OrderID === editId)?.OrderDate || new Date().toISOString().split('T')[0]) : new Date().toISOString().split('T')[0],
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
  // Send writes as GET params to avoid all CORS/redirect issues with GAS POST.
  // GAS doGet handles both reads (?action=read) and writes (?action=write&table=X&data=JSON)
  try {
    const encoded = encodeURIComponent(JSON.stringify(data));
    const url = `${GAS_URL}?action=write&table=${table}&data=${encoded}`;
    const res = await fetch(url);
    const result = await res.json();
    if (result.error) console.warn('Write error from GAS:', result.error);
    else console.log('Write OK:', result.action, table);
  } catch(e) {
    console.warn('Write failed:', e.message);
  }
}

// ── HELPERS ──────────────────────────────────────────────────────────────────
function nextId(prefix, existing) {
  const nums = existing.map(id => parseInt(id.replace(prefix + '-', '') || 0)).filter(Boolean);
  const max  = nums.length ? Math.max(...nums) : 0;
  return `${prefix}-${String(max + 1).padStart(3, '0')}`;
}

function fmtDate(dateStr) {
  if (!dateStr) return '—';
  try {
    const d = new Date(dateStr + 'T12:00:00');
    return d.toLocaleDateString('en-US', { month:'short', day:'numeric', year:'numeric' });
  } catch { return dateStr; }
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
