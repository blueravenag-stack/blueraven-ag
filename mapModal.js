// ── MAPMODAL.JS — Blue Raven Ag ───────────────────────────────────────────────
// Leaflet map modal: satellite imagery, CLU field boundaries, draw tools.
// Depends on: Leaflet (CDN), GeoUtils (geo.js)
// Called by: app.js via MapModal.open(options) / MapModal.close()
// ─────────────────────────────────────────────────────────────────────────────

window.MapModal = (() => {

  let map            = null;
  let cluLayer       = null;
  let drawLayer      = null;
  let selectedFields = [];   // [{id, farmNum, acres, points, polygon (leaflet layer)}]
  let drawPoints     = [];   // points being drawn manually
  let drawPolyline   = null;
  let drawMode       = false;
  let onConfirmCb    = null; // callback(selectedFields)
  let currentBounds  = null;

  // ── OPEN MAP MODAL ─────────────────────────────────────────────────────────
  async function open(opts = {}) {
    // opts: { centerLat, centerLng, customerAddress, onConfirm, preselected }
    onConfirmCb    = opts.onConfirm || null;
    selectedFields = [];
    drawPoints     = [];
    drawMode       = false;

    const modal = document.getElementById('mapModal');
    modal.classList.add('active');
    document.getElementById('modalOverlay').classList.add('active');
    updateSelectedPanel();

    // Init map if first open
    if (!map) {
      initMap();
    } else {
      map.invalidateSize();
    }

    // Restore preselected fields if editing
    if (opts.preselected) {
      // Will highlight after CLU loads — store for later
      map._preselected = opts.preselected;
    }

    // Center map
    const defaultLat = 39.8, defaultLng = -89.6; // Illinois center
    let lat = opts.centerLat || defaultLat;
    let lng = opts.centerLng || defaultLng;

    if (opts.customerAddress && (!opts.centerLat || !opts.centerLng)) {
      setStatus('Locating address...');
      const geo = await GeoUtils.geocodeAddress(opts.customerAddress);
      if (geo) { lat = geo.lat; lng = geo.lng; }
    }

    map.setView([lat, lng], 15);
    setStatus('Zoom to your field area — CLU boundaries load automatically');
  }

  function initMap() {
    map = L.map('mapContainer', { zoomControl: true });

    // ESRI World Imagery (free satellite, no API key)
    L.tileLayer(
      'https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}',
      { attribution: 'Imagery © Esri', maxZoom: 19 }
    ).addTo(map);

    // ESRI World Boundaries overlay (shows roads/labels on top of satellite)
    L.tileLayer(
      'https://server.arcgisonline.com/ArcGIS/rest/services/Reference/World_Boundaries_and_Places/MapServer/tile/{z}/{y}/{x}',
      { attribution: '', maxZoom: 19, opacity: 0.6 }
    ).addTo(map);

    cluLayer  = L.featureGroup().addTo(map);
    drawLayer = L.featureGroup().addTo(map);

    // Load CLU boundaries when map moves
    map.on('moveend zoomend', debounce(loadCLU, 600));
    map.on('click', onMapClick);
  }

  // ── CLU BOUNDARY LOADING ──────────────────────────────────────────────────
  async function loadCLU() {
    if (!map || map.getZoom() < 13) {
      cluLayer.clearLayers();
      setStatus('Zoom in closer to see CLU field boundaries');
      return;
    }

    const b = map.getBounds();
    currentBounds = {
      minLat: b.getSouth(), minLng: b.getWest(),
      maxLat: b.getNorth(), maxLng: b.getEast()
    };

    setStatus('Loading CLU boundaries...');
    const fields = await GeoUtils.fetchCLU(currentBounds);

    cluLayer.clearLayers();

    if (!fields.length) {
      setStatus('No CLU data for this area — draw field manually or paste KML');
      return;
    }

    fields.forEach(f => {
      if (!f.points || f.points.length < 3) return;
      const latlngs = f.points.map(p => [p.lat, p.lng]);
      const isSelected = selectedFields.some(s => s.id === f.id);
      const poly = L.polygon(latlngs, cluStyle(isSelected)).addTo(cluLayer);
      poly._cluData = f;
      poly._cluData.polygon = poly;
      poly.on('click', (e) => { L.DomEvent.stopPropagation(e); toggleField(poly); });

      // Tooltip with acreage
      poly.bindTooltip(`${f.acres} ac${f.farmNum ? ' · Farm ' + f.farmNum : ''}`,
        { permanent: false, direction: 'center', className: 'clu-tooltip' });
    });

    const count = fields.length;
    setStatus(`${count} CLU field${count !== 1 ? 's' : ''} loaded — tap to select`);
  }

  function cluStyle(selected) {
    return selected
      ? { color: '#4FC3F7', weight: 2.5, fillColor: '#4FC3F7', fillOpacity: 0.35 }
      : { color: '#81C784', weight: 1.5, fillColor: '#81C784', fillOpacity: 0.12 };
  }

  function toggleField(poly) {
    if (drawMode) return;
    const f = poly._cluData;
    const idx = selectedFields.findIndex(s => s.id === f.id);
    if (idx > -1) {
      selectedFields.splice(idx, 1);
      poly.setStyle(cluStyle(false));
    } else {
      selectedFields.push(f);
      poly.setStyle(cluStyle(true));
    }
    updateSelectedPanel();
  }

  // ── DRAW MODE ─────────────────────────────────────────────────────────────
  function startDraw() {
    drawMode   = true;
    drawPoints = [];
    drawLayer.clearLayers();
    if (drawPolyline) { drawPolyline = null; }
    setStatus('Click map to place polygon vertices — double-click to finish');
    document.getElementById('mapDrawBtn').textContent  = 'Cancel Draw';
    document.getElementById('mapDrawBtn').classList.add('active');
    map.getContainer().style.cursor = 'crosshair';
  }

  function cancelDraw() {
    drawMode   = false;
    drawPoints = [];
    drawLayer.clearLayers();
    setStatus('Draw cancelled');
    document.getElementById('mapDrawBtn').textContent = '✏ Draw Field';
    document.getElementById('mapDrawBtn').classList.remove('active');
    map.getContainer().style.cursor = '';
  }

  function finishDraw() {
    if (drawPoints.length < 3) {
      showMapToast('Need at least 3 points to close polygon');
      return;
    }
    const points = drawPoints.map(ll => ({ lat: ll.lat, lng: ll.lng }));
    const acres  = GeoUtils.calcAcres(points).toFixed(1);
    const ctr    = GeoUtils.centroid(points);

    // Add as a "drawn" field
    const drawnField = {
      id:      'DRAWN-' + Date.now(),
      farmNum: '',
      acres,
      points,
      isDrawn: true
    };

    const latlngs = drawPoints.map(ll => [ll.lat, ll.lng]);
    const poly = L.polygon(latlngs, cluStyle(true)).addTo(cluLayer);
    poly._cluData = { ...drawnField, polygon: poly };
    poly.on('click', (e) => { L.DomEvent.stopPropagation(e); toggleField(poly); });
    selectedFields.push(drawnField);

    cancelDraw();
    updateSelectedPanel();
    setStatus(`Drawn field added — ${acres} ac`);
  }

  function onMapClick(e) {
    if (!drawMode) return;
    drawPoints.push(e.latlng);

    // Update preview polyline
    if (drawPolyline) map.removeLayer(drawPolyline);
    const previewPts = [...drawPoints, drawPoints[0]].map(ll => [ll.lat, ll.lng]);
    drawPolyline = L.polyline(previewPts, { color: '#FFB74D', weight: 2, dashArray: '6' }).addTo(drawLayer);

    // Place vertex marker
    L.circleMarker(e.latlng, { radius: 5, color: '#FFB74D', fillColor: '#FFB74D', fillOpacity: 1 }).addTo(drawLayer);

    setStatus(`${drawPoints.length} point${drawPoints.length !== 1 ? 's' : ''} placed — double-click or click Finish to close`);
  }

  map && map.on('dblclick', (e) => { if (drawMode) { L.DomEvent.stopPropagation(e); finishDraw(); } });

  // ── PASTE KML ─────────────────────────────────────────────────────────────
  function showPastePanel(show) {
    document.getElementById('mapPastePanel').style.display = show ? 'block' : 'none';
  }

  function applyPastedKML() {
    const text = document.getElementById('mapPasteInput').value.trim();
    if (!text) return;
    const points = GeoUtils.parsePolygon(text);
    if (!points || points.length < 3) {
      showMapToast('Could not parse coordinates — try KML or lat/lng pairs');
      return;
    }
    const acres = GeoUtils.calcAcres(points).toFixed(1);
    const ctr   = GeoUtils.centroid(points);
    const latlngs = points.map(p => [p.lat, p.lng]);

    const pastedField = {
      id:      'PASTE-' + Date.now(),
      farmNum: '',
      acres,
      points,
      isPasted: true
    };
    const poly = L.polygon(latlngs, cluStyle(true)).addTo(cluLayer);
    poly._cluData = { ...pastedField, polygon: poly };
    poly.on('click', (e) => { L.DomEvent.stopPropagation(e); toggleField(poly); });
    selectedFields.push(pastedField);

    map.fitBounds(poly.getBounds(), { padding: [40, 40] });
    showPastePanel(false);
    document.getElementById('mapPasteInput').value = '';
    updateSelectedPanel();
    setStatus(`Pasted field added — ${acres} ac`);
  }

  // ── SELECTED FIELDS PANEL ─────────────────────────────────────────────────
  function updateSelectedPanel() {
    const el    = document.getElementById('mapSelectedFields');
    const total = selectedFields.reduce((s, f) => s + parseFloat(f.acres || 0), 0);
    if (!selectedFields.length) {
      el.innerHTML = '<span style="color:var(--text-sub);font-size:0.8rem">No fields selected</span>';
      document.getElementById('mapConfirmBtn').disabled = true;
    } else {
      el.innerHTML = selectedFields.map(f => `
        <div class="map-sel-chip">
          <span>${f.isDrawn ? '✏ Drawn' : f.isPasted ? '📋 Pasted' : 'CLU ' + (f.id || '?')}</span>
          <span>${f.acres} ac</span>
          <button onclick="MapModal.removeSelected('${f.id}')">×</button>
        </div>`).join('') +
        `<div class="map-sel-total">${total.toFixed(1)} ac total</div>`;
      document.getElementById('mapConfirmBtn').disabled = false;
    }
  }

  function removeSelected(id) {
    const f = selectedFields.find(s => s.id === id);
    if (f?.polygon) f.polygon.setStyle(cluStyle(false));
    selectedFields = selectedFields.filter(s => s.id !== id);
    updateSelectedPanel();
  }

  // ── CONFIRM ───────────────────────────────────────────────────────────────
  function confirm() {
    if (!selectedFields.length) return;
    if (onConfirmCb) onConfirmCb(selectedFields.map(f => ({
      id:      f.id,
      farmNum: f.farmNum,
      acres:   f.acres,
      points:  f.points,
      kml:     GeoUtils.pointsToKML(f.points),
      centroid:GeoUtils.centroid(f.points)
    })));
    close();
  }

  // ── CLOSE ─────────────────────────────────────────────────────────────────
  function close() {
    document.getElementById('mapModal').classList.remove('active');
    document.getElementById('modalOverlay').classList.remove('active');
    if (drawMode) cancelDraw();
  }

  // ── TOGGLE DRAW ───────────────────────────────────────────────────────────
  function toggleDraw() {
    if (drawMode) cancelDraw();
    else startDraw();
  }

  // ── FINISH DRAW (button) ──────────────────────────────────────────────────
  function finishDrawBtn() {
    if (drawMode) finishDraw();
  }

  // ── HELPERS ───────────────────────────────────────────────────────────────
  function setStatus(msg) {
    const el = document.getElementById('mapStatus');
    if (el) el.textContent = msg;
  }

  function showMapToast(msg) {
    const el = document.getElementById('mapStatus');
    if (el) { el.textContent = '⚠ ' + msg; el.style.color = 'var(--warn)'; }
    setTimeout(() => { if (el) el.style.color = ''; }, 2500);
  }

  function debounce(fn, ms) {
    let t;
    return (...args) => { clearTimeout(t); t = setTimeout(() => fn(...args), ms); };
  }

  // ── PUBLIC API ────────────────────────────────────────────────────────────
  return { open, close, confirm, toggleDraw, finishDrawBtn, removeSelected, showPastePanel, applyPastedKML };

})();
