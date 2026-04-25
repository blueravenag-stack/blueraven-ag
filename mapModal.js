// ── MAPMODAL.JS — Blue Raven Ag v1.9 ─────────────────────────────────────────
// Leaflet map modal: satellite imagery, CLU field boundaries, draw tools.
// Depends on: Leaflet (CDN), GeoUtils (geo.js)
// Called by: app.js via MapModal.open(opts) / MapModal.close()
// ─────────────────────────────────────────────────────────────────────────────

window.MapModal = (() => {

  let map            = null;
  let _customerFields = [];  // all customer fields shown as context
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
    // Restore preselected fields if provided (e.g. editing an order)
    if (opts.preselected) {
      selectedFields = opts.preselected.map(f => ({ ...f }));
    } else {
      selectedFields = [];
    }
    // Store customer context fields for drawing (not selected, just visible)
    _customerFields = opts.customerFields || [];
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
      setTimeout(() => { if (map) map.invalidateSize(); }, 200);

      const defaultLat = 39.8, defaultLng = -89.6;
      let lat = opts.centerLat || defaultLat;
      let lng = opts.centerLng || defaultLng;

      if (opts.customerAddress && (!opts.centerLat || !opts.centerLng)) {
        setStatus('Locating address...');
        const geo = await GeoUtils.geocodeAddress(opts.customerAddress);
        if (geo) { lat = geo.lat; lng = geo.lng; }
      }

      // Draw customer context fields first (dim, unselected)
      if (_customerFields.length > 0) {
        const selectedIds = new Set(selectedFields.map(f => f.id));
        _customerFields.forEach(f => {
          if (selectedIds.has(f.id) || !f.points || f.points.length < 3) return;
          const poly = L.polygon(f.points.map(p => [p.lat, p.lng]), {
            color: '#81C784', weight: 1.5, fillColor: '#81C784', fillOpacity: 0.1,
            dashArray: '4,3'
          }).addTo(cluLayer);
          poly._field = { ...f };
          poly.on('click', e => { L.DomEvent.stopPropagation(e); toggleField(poly); });
          poly.bindTooltip(`${f.fieldName} · ${parseFloat(f.acres||0).toFixed(1)} ac`,
            { permanent: false, direction: 'center', className: 'clu-tooltip' });
        });
      }

      // Draw any preselected fields that have stored polygon points
      if (selectedFields.length > 0) {
        const bounds = [];
        selectedFields.forEach(f => {
          if (!f.points || f.points.length < 3) return;
          const latlngs = f.points.map(p => [p.lat, p.lng]);
          const poly = L.polygon(latlngs, fieldStyle(true)).addTo(cluLayer);
          poly._field = { ...f, polygon: poly };
          poly.on('click', e => { L.DomEvent.stopPropagation(e); toggleField(poly); });
          poly.bindTooltip(`${parseFloat(f.acres||0).toFixed(1)} ac — ${f.fieldName||f.id}`,
            { permanent: false, direction: 'center', className: 'clu-tooltip' });
          // Update reference in selectedFields
          const idx = selectedFields.findIndex(s => s.id === f.id);
          if (idx > -1) selectedFields[idx] = { ...selectedFields[idx], polygon: poly };
          latlngs.forEach(ll => bounds.push(ll));
        });
        // Fit map to preselected fields
        if (bounds.length) {
          map.fitBounds(bounds, { padding: [40, 40] });
        } else {
          map.setView([lat, lng], 15);
        }
        updateSelectedPanel();
        setStatus(`${selectedFields.length} field${selectedFields.length!==1?'s':''} loaded — draw more or confirm`);
      } else {
        map.setView([lat, lng], 15);
        setStatus('Zoom to your field area, then draw a boundary or paste KML');
      }
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

    map.on('moveend zoomend', debounce(updateZoomStatus, 400));
    map.on('click',    onMapClick);
    map.on('dblclick', onMapDblClick);
  }

  // ── CLU LAYER MANAGEMENT ─────────────────────────────────────────────────
  // Only clear CLU data layers, keeping drawn/pasted field polys intact
  function clearCLULayer() {
    // Remove layers that are NOT drawn/pasted selected fields
    const selectedIds = new Set(selectedFields.map(f => f.id));
    cluLayer.eachLayer(layer => {
      if (!layer._field || !selectedIds.has(layer._field.id)) {
        cluLayer.removeLayer(layer);
      }
    });
  }

  // ── ZOOM STATUS ───────────────────────────────────────────────────────────
  function updateZoomStatus() {
    if (!map) return;
    const z = map.getZoom();
    if (z < 14) {
      setStatus('Zoom in to field level — then use ✏ Draw or 📋 Paste KML to add a field');
    } else {
      setStatus('✏ Click "Draw Field" to trace the field boundary, or paste KML below');
    }
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

    // Prompt for name immediately while map is still open
    const fieldName = prompt(`Field name for this drawn area (${acres} ac):`, 'New Field') || 'Drawn Field';

    const drawnField = {
      id:        'DRAWN-' + Date.now(),
      farmNum:   '',
      acres,
      points,
      isDrawn:   true,
      fieldName,
    };

    const poly = L.polygon(points.map(p => [p.lat, p.lng]), fieldStyle(true)).addTo(cluLayer);
    poly._field = { ...drawnField, polygon: poly };
    poly.on('click', e => { L.DomEvent.stopPropagation(e); toggleField(poly); });
    selectedFields.push({ ...drawnField, polygon: poly });

    cancelDraw();
    updateSelectedPanel();
    setStatus(`"${fieldName}" added — ${acres} ac. Draw another or confirm.`);
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
    const fieldName = prompt(`Field name for this pasted area (${acres} ac):`, 'Pasted Field') || 'Pasted Field';
    const pastedField = { id: 'PASTE-' + Date.now(), farmNum: '', acres, points, isPasted: true, fieldName };

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

    el.innerHTML = selectedFields.map(f => {
      // Use field name if it was saved as a DB field, otherwise show type label
      const label = f.fieldName || (f.isDrawn ? '✏ Drawn field' : f.isPasted ? '📋 Pasted field' : f.id || 'Field');
      return `<div class="map-sel-chip">
        <span>${label}</span>
        <span>${parseFloat(f.acres).toFixed(1)} ac</span>
        <button onclick="MapModal.removeSelected('${f.id}')">×</button>
      </div>`;
    }).join('') +
      `<div class="map-sel-total">${total.toFixed(1)} ac total</div>`;

    if (btn) btn.disabled = false;
  }

  function removeSelected(id) {
    const f = selectedFields.find(s => s.id === id);
    if (f?.polygon) f.polygon.setStyle(fieldStyle(false));
    selectedFields = selectedFields.filter(s => s.id !== id);
    updateSelectedPanel();
  }

  // ── ADDRESS SEARCH ────────────────────────────────────────────────────────
  async function searchAddress() {
    const input = document.getElementById('mapAddressInput');
    const val   = input?.value?.trim();
    if (!val) return;
    setStatus('Searching...');
    const geo = await GeoUtils.geocodeAddress(val);
    if (geo) {
      map.setView([geo.lat, geo.lng], 16);
      setStatus('Zoomed to location — draw your field boundary or paste KML');
    } else {
      setStatus('⚠ Address not found — try a nearby town or road intersection');
    }
  }

  // ── CONFIRM ───────────────────────────────────────────────────────────────
  function confirm() {
    if (!selectedFields.length) return;
    if (onConfirmCb) {
      onConfirmCb(selectedFields.map(f => ({
        id:        f.id,
        farmNum:   f.farmNum || '',
        acres:     f.acres,
        points:    f.points,
        kml:       GeoUtils.pointsToKML(f.points),
        centroid:  GeoUtils.centroid(f.points),
        fieldName: f.fieldName || ''
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


  // ── SHAPEFILE / CLU LOAD ──────────────────────────────────────────────────
  // Accepts .zip (shapefile package) or .geojson / .json file
  // User downloads CLU zip from USDA Geospatial Data Gateway:
  //   https://gdg.sc.egov.usda.gov/GDGOrder.aspx
  //   Illinois > Common Land Unit (CLU) > Select your county

  async function loadShapefileInput(input) {
    const file = input.files[0];
    if (!file) return;
    setStatus('Reading file...');

    try {
      if (file.name.endsWith('.geojson') || file.name.endsWith('.json')) {
        const text = await file.text();
        loadGeoJSON(JSON.parse(text));
      } else if (file.name.endsWith('.zip')) {
        await loadShapefileZip(file);
      } else {
        setStatus('⚠ Upload a .zip (shapefile) or .geojson file');
      }
    } catch(e) {
      console.error(e);
      setStatus('⚠ Error reading file: ' + e.message);
    }
    input.value = ''; // reset so same file can be reloaded
  }

  async function loadShapefileZip(file) {
    // Requires JSZip + shapefile.js (loaded in index.html)
    if (typeof JSZip === 'undefined' || typeof shapefile === 'undefined') {
      setStatus('⚠ Shapefile libraries not loaded — refresh and try again');
      return;
    }
    setStatus('Unzipping shapefile...');
    const zip = await JSZip.loadAsync(file);
    
    // Find .shp and .dbf files
    const shpFile = Object.values(zip.files).find(f => f.name.toLowerCase().endsWith('.shp'));
    const dbfFile = Object.values(zip.files).find(f => f.name.toLowerCase().endsWith('.dbf'));
    
    if (!shpFile) { setStatus('⚠ No .shp file found in zip'); return; }
    
    setStatus('Parsing fields...');
    const shpBuf = await shpFile.async('arraybuffer');
    const dbfBuf = dbfFile ? await dbfFile.async('arraybuffer') : null;
    
    const source = await shapefile.open(shpBuf, dbfBuf);
    const features = [];
    let result = await source.read();
    while (!result.done) {
      features.push(result.value);
      result = await source.read();
    }
    
    const geojson = { type: 'FeatureCollection', features };
    loadGeoJSON(geojson);
  }

  function loadGeoJSON(geojson) {
    clearCLULayer();
    const features = geojson.features || [];
    let count = 0;
    let fieldIndex = 0;

    features.forEach((feat, featIdx) => {
      const props = feat.properties || {};
      const geom  = feat.geometry;
      if (!geom) return;

      // Extract all rings from the geometry as individual field candidates
      // Handles: Polygon (single or multi-ring), MultiPolygon
      const rings = extractRings(geom);
      if (!rings.length) return;

      // Filter to only outer rings (CLU parts are typically all outer rings)
      // Outer ring determination: positive signed area = CCW (standard GeoJSON)
      // We treat ALL rings as potential fields since CLU exports pack separate
      // fields as multiple parts in a single shape record
      console.log(`Feature ${featIdx}: ${rings.length} rings`);
      rings.forEach((ring, ringIdx) => {
        if (ring.length < 4) return; // need at least 3 unique points + closure

        const points = ring.map(c => ({ lng: c[0], lat: c[1] }));
        const acres  = GeoUtils.calcAcres(points).toFixed(1);
        const farmNum= String(props.FARMNBR || props.farm_number || props.FARM_NUM || '');
        const cluNum = String(props.CLUNBR || props.clu_number || props.CLU_NUM ||
                              props.FID || `${featIdx}-${ringIdx}`);
        const uniqueId = rings.length > 1 ? `${cluNum}-r${ringIdx}` : cluNum;

        // Skip tiny rings (< 0.5 acres) — likely digitizing artifacts
        if (parseFloat(acres) < 0.5) return;

        const isSelected = selectedFields.some(s => s.id === uniqueId);
        const poly = L.polygon(points.map(p => [p.lat, p.lng]), fieldStyle(isSelected)).addTo(cluLayer);
        const field = { id: uniqueId, farmNum, acres, points, fromFile: true };
        poly._field = field;
        if (isSelected) { const sf = selectedFields.find(s => s.id === uniqueId); if(sf) sf.polygon = poly; }

        poly.on('click', e => { L.DomEvent.stopPropagation(e); toggleField(poly); });
        poly.bindTooltip(
          `${parseFloat(acres).toFixed(1)} ac${farmNum ? ' · Farm ' + farmNum : ''}`,
          { permanent: false, direction: 'center', className: 'clu-tooltip' }
        );
        count++;
      });
    });

    if (count > 0) {
      try { map.fitBounds(cluLayer.getBounds(), { padding: [20,20] }); } catch(e) {}
      setStatus(`${count} field${count!==1?'s':''} loaded from ${features.length} shape record${features.length!==1?'s':''} — tap to select`);
    } else {
      setStatus('⚠ No valid field polygons found in file');
    }
  }

  // Extract all coordinate rings from any GeoJSON geometry type
  function extractRings(geom) {
    if (!geom) return [];
    if (geom.type === 'Polygon') {
      // Return all rings (outer + inner) as separate candidates
      // For CLU data, inner rings are sometimes separate fields, not holes
      return (geom.coordinates || []).filter(r => r && r.length >= 4);
    }
    if (geom.type === 'MultiPolygon') {
      // Flatten all polygons and all their rings
      return (geom.coordinates || []).flatMap(poly =>
        (poly || []).filter(r => r && r.length >= 4)
      );
    }
    return [];
  }

  // ── PUBLIC ────────────────────────────────────────────────────────────────
  return { open, close, confirm, toggleDraw, finishDrawBtn, removeSelected, showPastePanel, applyPastedKML, searchAddress, loadShapefileInput };

})();
