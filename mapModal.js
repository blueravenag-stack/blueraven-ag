// ── MAPMODAL.JS — Blue Raven Ag v1.9 ─────────────────────────────────────────
// Leaflet map modal: satellite imagery, CLU field boundaries, draw tools.
// Depends on: Leaflet (CDN), GeoUtils (geo.js)
// Called by: app.js via MapModal.open(opts) / MapModal.close()
// ─────────────────────────────────────────────────────────────────────────────

window.MapModal = (() => {

  let map            = null;
  let cluLayer       = null;
  let drawLayer      = null;
  let selectedFields = [];
  let drawPoints     = [];
  let drawPolyline   = null;
  let drawMode       = false;
  let onConfirmCb    = null;
  let onCloseCb      = null;

  // ── OPEN ──────────────────────────────────────────────────────────────────
  async function open(opts = {}) {
    onConfirmCb    = opts.onConfirm || null;
    onCloseCb      = opts.onClose   || null;
    selectedFields = [];
    drawPoints     = [];
    drawMode       = false;

    // Show modal first so map container has dimensions
    document.getElementById('mapModal').classList.add('active');
    document.getElementById('mapOverlay').classList.add('active');
    updateSelectedPanel();
    setStatus('Initializing map...');

    // Defer Leaflet init until modal is fully painted and laid out
    setTimeout(async () => {
      if (!map) initMap();
      // Double invalidate — first after paint, second after layout fully settles
      setTimeout(() => { if (map) map.invalidateSize(); }, 200);

      const defaultLat = 39.8, defaultLng = -89.6;
      let lat = opts.centerLat || defaultLat;
      let lng = opts.centerLng || defaultLng;

      if (opts.customerAddress && (!opts.centerLat || !opts.centerLng)) {
        setStatus('Locating address...');
        const geo = await GeoUtils.geocodeAddress(opts.customerAddress);
        if (geo) { lat = geo.lat; lng = geo.lng; }
      }

      map.setView([lat, lng], 15);
      setStatus('Zoom to your field area — CLU boundaries load automatically');
    }, 120);
  }

  // ── INIT MAP ──────────────────────────────────────────────────────────────
  function initMap() {
    map = L.map('mapContainer', { zoomControl: true, doubleClickZoom: false });

    // ESRI World Imagery (free satellite)
    L.tileLayer(
      'https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}',
      { attribution: 'Imagery © Esri', maxZoom: 19 }
    ).addTo(map);

    // Labels overlay
    L.tileLayer(
      'https://server.arcgisonline.com/ArcGIS/rest/services/Reference/World_Boundaries_and_Places/MapServer/tile/{z}/{y}/{x}',
      { maxZoom: 19, opacity: 0.7 }
    ).addTo(map);

    cluLayer  = L.featureGroup().addTo(map);
    drawLayer = L.featureGroup().addTo(map);

    map.on('moveend zoomend', debounce(loadCLU, 700));
    map.on('click',    onMapClick);
    map.on('dblclick', onMapDblClick);
  }

  // ── CLU LOADING ───────────────────────────────────────────────────────────
  async function loadCLU() {
    if (!map) return;
    if (map.getZoom() < 13) {
      cluLayer.clearLayers();
      setStatus('Zoom in closer to see CLU field boundaries');
      return;
    }

    const b = map.getBounds();
    const bounds = {
      minLat: b.getSouth(), minLng: b.getWest(),
      maxLat: b.getNorth(), maxLng: b.getEast()
    };

    setStatus('Loading CLU boundaries...');
    const fields = await GeoUtils.fetchCLU(bounds);
    cluLayer.clearLayers();

    if (!fields.length) {
      setStatus('No CLU data for this area — draw a field manually or paste KML');
      return;
    }

    fields.forEach(f => {
      if (!f.points || f.points.length < 3) return;
      const isSelected = selectedFields.some(s => s.id === f.id);
      const poly = L.polygon(f.points.map(p => [p.lat, p.lng]), fieldStyle(isSelected)).addTo(cluLayer);
      poly._field = { ...f };
      poly.on('click', e => { L.DomEvent.stopPropagation(e); toggleField(poly); });
      poly.bindTooltip(
        `${f.acres} ac${f.farmNum ? ' · Farm ' + f.farmNum : ''}`,
        { permanent: false, direction: 'center', className: 'clu-tooltip' }
      );
    });

    setStatus(`${fields.length} CLU field${fields.length !== 1 ? 's' : ''} loaded — tap to select`);
  }

  function fieldStyle(selected) {
    return selected
      ? { color: '#4FC3F7', weight: 2.5, fillColor: '#4FC3F7', fillOpacity: 0.35 }
      : { color: '#81C784', weight: 1.5, fillColor: '#81C784', fillOpacity: 0.12 };
  }

  function toggleField(poly) {
    if (drawMode) return;
    const f   = poly._field;
    const idx = selectedFields.findIndex(s => s.id === f.id);
    if (idx > -1) {
      selectedFields.splice(idx, 1);
      poly.setStyle(fieldStyle(false));
    } else {
      selectedFields.push({ ...f, polygon: poly });
      poly.setStyle(fieldStyle(true));
    }
    updateSelectedPanel();
  }

  // ── DRAW MODE ─────────────────────────────────────────────────────────────
  function toggleDraw() {
    if (drawMode) cancelDraw();
    else          startDraw();
  }

  function startDraw() {
    drawMode   = true;
    drawPoints = [];
    drawLayer.clearLayers();
    setStatus('Click to place vertices — double-click or press Finish to close');
    const btn = document.getElementById('mapDrawBtn');
    if (btn) { btn.textContent = 'Cancel Draw'; btn.classList.add('active'); }
    const finBtn = document.getElementById('mapFinishBtn');
    if (finBtn) finBtn.style.display = 'inline-flex';
    map.getContainer().style.cursor = 'crosshair';
  }

  function cancelDraw() {
    drawMode   = false;
    drawPoints = [];
    drawLayer.clearLayers();
    if (drawPolyline) { map.removeLayer(drawPolyline); drawPolyline = null; }
    const btn = document.getElementById('mapDrawBtn');
    if (btn) { btn.textContent = '✏ Draw Field'; btn.classList.remove('active'); }
    const finBtn = document.getElementById('mapFinishBtn');
    if (finBtn) finBtn.style.display = 'none';
    map.getContainer().style.cursor = '';
    setStatus('Draw cancelled');
  }

  function finishDrawBtn() {
    if (drawMode) finishDraw();
  }

  function finishDraw() {
    if (drawPoints.length < 3) {
      setStatus('⚠ Need at least 3 points — keep clicking to add more');
      return;
    }
    const points = drawPoints.map(ll => ({ lat: ll.lat, lng: ll.lng }));
    const acres  = GeoUtils.calcAcres(points).toFixed(1);

    const drawnField = {
      id:      'DRAWN-' + Date.now(),
      farmNum: '',
      acres,
      points,
      isDrawn: true
    };

    const poly = L.polygon(points.map(p => [p.lat, p.lng]), fieldStyle(true)).addTo(cluLayer);
    poly._field = { ...drawnField, polygon: poly };
    poly.on('click', e => { L.DomEvent.stopPropagation(e); toggleField(poly); });
    selectedFields.push({ ...drawnField, polygon: poly });

    cancelDraw();
    updateSelectedPanel();
    setStatus(`Drawn field added — ${acres} ac. Tap to deselect or draw another.`);
  }

  function onMapClick(e) {
    if (!drawMode) return;
    drawPoints.push(e.latlng);

    if (drawPolyline) map.removeLayer(drawPolyline);
    const preview = [...drawPoints.map(ll => [ll.lat, ll.lng]), [drawPoints[0].lat, drawPoints[0].lng]];
    drawPolyline = L.polyline(preview, { color: '#FFB74D', weight: 2, dashArray: '6,4' }).addTo(drawLayer);
    L.circleMarker(e.latlng, { radius: 5, color: '#FFB74D', fillColor: '#FFB74D', fillOpacity: 1, weight: 1 }).addTo(drawLayer);

    setStatus(`${drawPoints.length} point${drawPoints.length !== 1 ? 's' : ''} placed — double-click or Finish to close polygon`);
  }

  function onMapDblClick(e) {
    if (drawMode) {
      L.DomEvent.stopPropagation(e);
      finishDraw();
    }
  }

  // ── PASTE KML ─────────────────────────────────────────────────────────────
  function showPastePanel(show) {
    const el = document.getElementById('mapPastePanel');
    if (el) el.style.display = show ? 'block' : 'none';
  }

  function applyPastedKML() {
    const text = (document.getElementById('mapPasteInput')?.value || '').trim();
    if (!text) return;
    const points = GeoUtils.parsePolygon(text);
    if (!points || points.length < 3) {
      setStatus('⚠ Could not parse — try KML, GeoJSON, or lat/lng pairs');
      return;
    }
    const acres = GeoUtils.calcAcres(points).toFixed(1);
    const pastedField = { id: 'PASTE-' + Date.now(), farmNum: '', acres, points, isPasted: true };

    const poly = L.polygon(points.map(p => [p.lat, p.lng]), fieldStyle(true)).addTo(cluLayer);
    poly._field = { ...pastedField, polygon: poly };
    poly.on('click', e => { L.DomEvent.stopPropagation(e); toggleField(poly); });
    selectedFields.push({ ...pastedField, polygon: poly });

    map.fitBounds(poly.getBounds(), { padding: [40, 40] });
    showPastePanel(false);
    document.getElementById('mapPasteInput').value = '';
    updateSelectedPanel();
    setStatus(`Pasted field added — ${acres} ac`);
  }

  // ── SELECTED PANEL ────────────────────────────────────────────────────────
  function updateSelectedPanel() {
    const el    = document.getElementById('mapSelectedFields');
    const btn   = document.getElementById('mapConfirmBtn');
    const total = selectedFields.reduce((s, f) => s + parseFloat(f.acres || 0), 0);

    if (!el) return;
    if (!selectedFields.length) {
      el.innerHTML = '<span style="color:var(--text-sub);font-size:0.8rem">No fields selected — tap a green boundary</span>';
      if (btn) btn.disabled = true;
      return;
    }

    el.innerHTML = selectedFields.map(f => `
      <div class="map-sel-chip">
        <span>${f.isDrawn ? '✏ Drawn' : f.isPasted ? '📋 Pasted' : 'CLU ' + (f.id || '?')}</span>
        <span>${parseFloat(f.acres).toFixed(1)} ac</span>
        <button onclick="MapModal.removeSelected('${f.id}')">×</button>
      </div>`).join('') +
      `<div class="map-sel-total">${total.toFixed(1)} ac total</div>`;

    if (btn) btn.disabled = false;
  }

  function removeSelected(id) {
    const f = selectedFields.find(s => s.id === id);
    if (f?.polygon) f.polygon.setStyle(fieldStyle(false));
    selectedFields = selectedFields.filter(s => s.id !== id);
    updateSelectedPanel();
  }

  // ── CONFIRM ───────────────────────────────────────────────────────────────
  function confirm() {
    if (!selectedFields.length) return;
    if (onConfirmCb) {
      onConfirmCb(selectedFields.map(f => ({
        id:       f.id,
        farmNum:  f.farmNum || '',
        acres:    f.acres,
        points:   f.points,
        kml:      GeoUtils.pointsToKML(f.points),
        centroid: GeoUtils.centroid(f.points)
      })));
    }
    close();
  }

  // ── CLOSE ─────────────────────────────────────────────────────────────────
  function close() {
    document.getElementById('mapModal').classList.remove('active');
    document.getElementById('mapOverlay').classList.remove('active');
    if (drawMode) cancelDraw();
    if (onCloseCb) { onCloseCb(); onCloseCb = null; }
  }

  // ── HELPERS ───────────────────────────────────────────────────────────────
  function setStatus(msg) {
    const el = document.getElementById('mapStatus');
    if (el) el.textContent = msg;
  }

  function debounce(fn, ms) {
    let t;
    return (...args) => { clearTimeout(t); t = setTimeout(() => fn(...args), ms); };
  }

  // ── PUBLIC ────────────────────────────────────────────────────────────────
  return { open, close, confirm, toggleDraw, finishDrawBtn, removeSelected, showPastePanel, applyPastedKML };

})();
