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
          // Support multi-polygon KML (multiple rings stored in one field)
          const rings = f.kml ? GeoUtils.parseKMLAllRings(f.kml) :
                        f.points && f.points.length >= 3 ? [f.points] : [];
          if (!rings.length) return;

          // Draw all rings as one combined layer, store ref on first ring's poly
          let primaryPoly = null;
          rings.forEach((pts, ri) => {
            const latlngs = pts.map(p => [p.lat, p.lng]);
            const poly = L.polygon(latlngs, fieldStyle(true)).addTo(cluLayer);
            poly._field = { ...f, polygon: poly };
            poly.on('click', e => { L.DomEvent.stopPropagation(e); toggleField(poly); });
            if (ri === 0) {
              primaryPoly = poly;
              poly.bindTooltip(`${parseFloat(f.acres||0).toFixed(1)} ac — ${f.fieldName||f.id}`,
                { permanent: false, direction: 'center', className: 'clu-tooltip' });
            }
            latlngs.forEach(ll => bounds.push(ll));
          });

          const idx = selectedFields.findIndex(s => s.id === f.id);
          if (idx > -1 && primaryPoly) selectedFields[idx] = { ...selectedFields[idx], polygon: primaryPoly };
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
      'https://server.arcgisonline.com/ArcGIS/rest/services/Reference/World_Transportation/MapServer/tile/{z}/{y}/{x}',
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
    if (_editingField) { exitEditMode(true); return; }
    const f   = poly._field;
    const idx = selectedFields.findIndex(s => s.id === f.id);
    if (idx > -1) {
      // Already selected — enter edit mode on second click
      enterEditMode(poly);
      return;
    }
    selectedFields.push({ ...f, polygon: poly });
    poly.setStyle(fieldStyle(true));
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
      const label = f.fieldName || (f.isDrawn ? '✏ Drawn field' : f.isPasted ? '📋 Pasted field' : f.id || 'Field');
      const safeId = f.id.replace(/'/g, "\'");
      return `<div class="map-sel-chip">
        <input class="map-sel-name" value="${label.replace(/"/g, '&quot;')}"
          onchange="MapModal.renameSelected('${safeId}', this.value)"
          title="Click to rename this field">
        <span class="map-sel-acres">${parseFloat(f.acres).toFixed(1)} ac</span>
        <button onclick="MapModal.removeSelected('${safeId}')">×</button>
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

  function renameSelected(id, newName) {
    const f = selectedFields.find(s => s.id === id);
    if (f) f.fieldName = newName.trim() || f.fieldName;
    // Update tooltip on polygon if it exists
    if (f?.polygon) {
      f.polygon.getTooltip()?.setContent(`${parseFloat(f.acres||0).toFixed(1)} ac — ${f.fieldName}`);
    }
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
    console.log('[MapModal.confirm] selectedFields count:', selectedFields.length,
      selectedFields.map(f => ({id:f.id, acres:f.acres, hasPoints:!!(f.points?.length)})));
    if (onConfirmCb) {
      const payload = selectedFields.map(f => ({
        id:        f.id,
        farmNum:   f.farmNum || '',
        acres:     f.acres,
        points:    f.points,
        kml:       GeoUtils.pointsToKML(f.points),
        centroid:  GeoUtils.centroid(f.points),
        fieldName: f.fieldName || ''
      }));
      console.log('[MapModal.confirm] payload count:', payload.length);
      onConfirmCb(payload);
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

  // ── CLU LAYER — stored parsed data, viewport-filtered rendering ──────────────
  let _cluData       = null;   // [{id, acres, farmNum, bbox, rings}] parsed once, kept in memory
  let _cluLayerGroup = null;   // Leaflet layer group for CLU polygons
  let _cluVisible    = false;  // whether CLU is currently shown

  async function loadShapefileZip(file) {
    if (typeof JSZip === 'undefined' || typeof shapefile === 'undefined') {
      setStatus('⚠ Shapefile libraries not loaded — refresh and try again');
      return;
    }
    setStatus('Loading shapefile — this may take a few seconds for large files...');

    try {
      const zip = await JSZip.loadAsync(file);
      const shpFile = Object.values(zip.files).find(f => f.name.toLowerCase().endsWith('.shp'));
      const dbfFile = Object.values(zip.files).find(f => f.name.toLowerCase().endsWith('.dbf'));
      if (!shpFile) { setStatus('⚠ No .shp file found in zip'); return; }

      setStatus('Parsing field boundaries...');
      const shpBuf = await shpFile.async('arraybuffer');
      const dbfBuf = dbfFile ? await dbfFile.async('arraybuffer') : null;

      const source = await shapefile.open(shpBuf, dbfBuf);
      _cluData = [];
      let result = await source.read();
      let count = 0;

      while (!result.done) {
        const feat = result.value;
        const props = feat.properties || {};
        const rings = extractRings(feat.geometry);
        const acres_prop = parseFloat(props.Acres || props.ACRES || props.CALCACRES || 0);

        rings.forEach((ring, ringIdx) => {
          if (ring.length < 4) return;
          // Reproject to WGS84 if coordinates are in Albers meters
          const wgsRing = GeoUtils.normalizeRing(ring);
          // Compute bbox for viewport filtering (in WGS84 degrees)
          const lngs = wgsRing.map(c => c[0]);
          const lats = wgsRing.map(c => c[1]);
          const bbox = {
            xmin: Math.min(...lngs), xmax: Math.max(...lngs),
            ymin: Math.min(...lats), ymax: Math.max(...lats),
          };
          const points = wgsRing.map(c => ({ lng: c[0], lat: c[1] }));
          const acres  = acres_prop > 0 && rings.length === 1
            ? acres_prop.toFixed(1)
            : GeoUtils.calcAcres(points).toFixed(1);

          if (parseFloat(acres) < 0.5) return; // skip slivers

          _cluData.push({
            id:      rings.length > 1 ? `${count}-r${ringIdx}` : `${count}`,
            farmNum: String(props.FARMNBR || props.farm_number || ''),
            acres,
            bbox,
            points,
          });
        });
        count++;
        result = await source.read();
      }

      // Attach viewport renderer to map move events
      if (!_cluLayerGroup) {
        _cluLayerGroup = L.featureGroup().addTo(map);
        map.on('moveend zoomend', renderCLUViewport);
      }
      _cluVisible = true;

      if (_cluData.length <= 200) {
        // Small file: show all polygons directly via loadGeoJSON (no viewport filter)
        const features = _cluData.map(f => ({
          type: 'Feature',
          geometry: { type: 'Polygon', coordinates: [f.points.map(p => [p.lng, p.lat])] },
          properties: { FARMNBR: f.farmNum, CALCACRES: f.acres }
        }));
        _cluData = [];   // clear so we don't confuse viewport renderer
        _cluLayerGroup?.clearLayers();
        loadGeoJSON({ type: 'FeatureCollection', features });
      } else {
        // Large file: viewport-filtered rendering — do NOT zoom out
        setStatus(`✓ ${_cluData.length.toLocaleString()} fields loaded — CLU boundaries will show at zoom 12+`);
        renderCLUViewport();
        // Only fit bounds if map is at default/uninitialized position
        // (i.e. user hasn't navigated yet — check if at default IL center)
        try {
          const center = map.getCenter();
          const isDefaultPos = Math.abs(center.lat - 39.8) < 0.5 && Math.abs(center.lng + 89.6) < 0.5;
          if (isDefaultPos) {
            const allLats = _cluData.flatMap(f => [f.bbox.ymin, f.bbox.ymax]);
            const allLngs = _cluData.flatMap(f => [f.bbox.xmin, f.bbox.xmax]);
            map.fitBounds([
              [Math.min(...allLats), Math.min(...allLngs)],
              [Math.max(...allLats), Math.max(...allLngs)]
            ], { padding: [30, 30] });
          }
        } catch(e) {}
      }

    } catch(e) {
      console.error('Shapefile load error:', e);
      setStatus('⚠ Error loading shapefile: ' + e.message);
    }
  }

  function renderCLUViewport() {
    if (!_cluData || !_cluVisible || !map || !_cluLayerGroup) return;
    if (map.getZoom() < 12) {
      _cluLayerGroup.clearLayers();
      setStatus(`CLU: ${_cluData.length.toLocaleString()} fields loaded — zoom in to see boundaries`);
      return;
    }

    const b      = map.getBounds();
    const vxmin  = b.getWest(), vxmax  = b.getEast();
    const vymin  = b.getSouth(), vymax = b.getNorth();

    const visible = _cluData.filter(f =>
      f.bbox.xmax >= vxmin && f.bbox.xmin <= vxmax &&
      f.bbox.ymax >= vymin && f.bbox.ymin <= vymax
    );

    _cluLayerGroup.clearLayers();

    visible.forEach(f => {
      const isSelected = selectedFields.some(s => s.id === f.id);
      const poly = L.polygon(f.points.map(p => [p.lat, p.lng]),
        isSelected ? fieldStyle(true) : {
          color: '#4DB6AC', weight: 1, fillColor: '#4DB6AC', fillOpacity: 0.1
        }
      ).addTo(_cluLayerGroup);

      poly._field = { ...f, fromFile: true };
      poly.on('click', e => { L.DomEvent.stopPropagation(e); toggleField(poly); });
      poly.bindTooltip(
        `${parseFloat(f.acres).toFixed(1)} ac${f.farmNum ? ' · Farm ' + f.farmNum : ''}`,
        { permanent: false, direction: 'center', className: 'clu-tooltip' }
      );
    });

    setStatus(visible.length > 0
      ? `${visible.length} CLU field${visible.length!==1?'s':''} visible — tap to select`
      : `No CLU fields in this view — pan to your area`
    );
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


  // ── GITHUB CLU TOGGLE ──────────────────────────────────────────────────────
  // Fetches the CLU zip from GitHub repo, parses once, shows viewport-filtered
  let _cluLoaded     = false;
  let _cluLoadingNow = false;

  async function toggleCLU() {
    const btn = document.getElementById('mapCLUBtn');

    if (_cluVisible && _cluData) {
      // Turn OFF — preserve selected field polygons, only remove unselected CLU
      _cluVisible = false;
      const selectedIds = new Set(selectedFields.map(f => f.id));
      if (_cluLayerGroup) {
        _cluLayerGroup.eachLayer(layer => {
          if (layer._field && selectedIds.has(layer._field.id)) {
            // Move selected polygon to cluLayer so it persists
            _cluLayerGroup.removeLayer(layer);
            cluLayer.addLayer(layer);
          }
        });
        _cluLayerGroup.clearLayers();
      }
      if (btn) { btn.textContent = '🌾 Show CLU Fields'; btn.classList.remove('active'); }
      setStatus('CLU boundaries hidden — selected fields remain visible');
      return;
    }

    if (_cluData && _cluData.length > 0) {
      // Already parsed — just re-enable in current view
      _cluVisible = true;
      renderCLUViewport();
      if (btn) { btn.textContent = '🌾 Hide CLU Fields'; btn.classList.add('active'); }
      return;
    }

    if (_cluLoadingNow) return;
    _cluLoadingNow = true;
    if (btn) { btn.textContent = '⏳ Loading CLU...'; btn.disabled = true; }
    setStatus('Fetching CLU boundaries from GitHub...');

    // URL to the CLU zip stored in the GitHub repo
    const CLU_URL = 'IL_ACPF_Greene_Jersey_Macoupin_Madison.zip';

    try {
      const res = await fetch(CLU_URL);
      if (!res.ok) throw new Error(`HTTP ${res.status} — make sure IL_ACPF_Greene_Jersey_Macoupin_Madison.zip is in your GitHub repo root`);
      const arrayBuf = await res.arrayBuffer();
      const blob = new Blob([arrayBuf], { type: 'application/zip' });
      const file = new File([blob], 'clu.zip');
      await loadShapefileZip(file);
      _cluVisible = true;
      if (btn) { btn.textContent = '🌾 Hide CLU Fields'; btn.classList.remove('active'); btn.disabled = false; btn.classList.add('active'); }
    } catch(e) {
      setStatus('⚠ CLU load failed: ' + e.message);
      if (btn) { btn.textContent = '🌾 Show CLU Fields'; btn.disabled = false; }
    }
    _cluLoadingNow = false;
  }


  // ── POLYGON EDITING ────────────────────────────────────────────────────────
  // Click a SELECTED field again to enter vertex-edit mode.
  // Drag vertices to adjust. Click elsewhere to finish.

  let _editingField  = null;  // {fieldId, poly, originalPoints}
  let _editMarkers   = [];    // vertex drag markers

  function vertexIcon() {
    return L.divIcon({
      className: '',
      html: '<div style="width:14px;height:14px;background:#FFB74D;border:2px solid #fff;border-radius:50%;cursor:grab;box-shadow:0 1px 3px rgba(0,0,0,0.4)"></div>',
      iconSize: [14, 14],
      iconAnchor: [7, 7],
    });
  }

  function makeVertexMarker(ll, poly) {
    const marker = L.marker(ll, {
      icon:      vertexIcon(),
      draggable: true,
      autoPan:   false,
      bubblingMouseEvents: false,
      zIndexOffset: 1000,
    }).addTo(map);

    marker.on('drag', () => {
      poly.setLatLngs([_editMarkers.map(m => m.getLatLng())]);
    });

    // Double-click vertex to delete it
    marker.on('dblclick', (e) => {
      L.DomEvent.stopPropagation(e);
      if (_editMarkers.length <= 4) { setStatus('⚠ Need at least 3 points'); return; }
      const i = _editMarkers.indexOf(marker);
      _editMarkers.splice(i, 1);
      map.removeLayer(marker);
      poly.setLatLngs([_editMarkers.map(m => m.getLatLng())]);
      setStatus('Vertex deleted — double-click another to remove, click edge to add');
    });

    // Right-click also deletes (desktop)
    marker.on('contextmenu', (e) => {
      L.DomEvent.stopPropagation(e);
      if (_editMarkers.length <= 4) { setStatus('⚠ Need at least 3 points'); return; }
      const i = _editMarkers.indexOf(marker);
      _editMarkers.splice(i, 1);
      map.removeLayer(marker);
      poly.setLatLngs([_editMarkers.map(m => m.getLatLng())]);
    });

    return marker;
  }

  function enterEditMode(poly) {
    if (_editingField) exitEditMode(true);

    const f = poly._field;
    const latlngs = poly.getLatLngs()[0] || [];
    _editingField = {
      id: f.id,
      poly,
      originalPoints: latlngs.map(ll => ({ lat: ll.lat, lng: ll.lng }))
    };
    poly.setStyle({ color: '#FFB74D', weight: 3, fillColor: '#FFB74D', fillOpacity: 0.2 });

    _editMarkers = latlngs.map(ll => makeVertexMarker(ll, poly));

    document.getElementById('mapEditFinishBtn')?.style.setProperty('display', 'inline-flex');
    setStatus('Edit mode: drag vertices · right-click to delete · tap edge to add · Finish Edit when done');

    map.off('click', onMapClick);
    map.on('click', onEditMapClick);
    setStatus('Edit: drag vertices · dbl-click vertex to delete · click near edge to add point · Finish Edit when done');
  }

  function onEditMapClick(e) {
    if (!_editingField) return;
    const ll   = e.latlng;
    const poly = _editingField.poly;
    const pts  = _editMarkers.map(m => m.getLatLng());

    // Find closest edge and insertion point
    let minDist = Infinity, insertAt = pts.length;
    for (let i = 0; i < pts.length; i++) {
      const a = pts[i], b = pts[(i+1) % pts.length];
      const d = pointToSegmentDist(ll, a, b);
      if (d < minDist) { minDist = d; insertAt = i+1; }
    }

    // Only insert if within ~20 pixels of the edge at current zoom
    const metersPerPixel = 40075016.686 * Math.cos(ll.lat * Math.PI/180) / Math.pow(2, map.getZoom() + 8);
    if (minDist > metersPerPixel * 40) return; // within 40px of edge

    const newMarker = makeVertexMarker(ll, poly);
    _editMarkers.splice(insertAt, 0, newMarker);
    poly.setLatLngs([_editMarkers.map(m => m.getLatLng())]);
    setStatus('Vertex added — drag to adjust · dbl-click to remove');
  }

  function pointToSegmentDist(p, a, b) {
    const dx = b.lng - a.lng, dy = b.lat - a.lat;
    if (dx === 0 && dy === 0) return p.distanceTo(a);
    const t = Math.max(0, Math.min(1, ((p.lng-a.lng)*dx + (p.lat-a.lat)*dy) / (dx*dx+dy*dy)));
    return p.distanceTo(L.latLng(a.lat+t*dy, a.lng+t*dx));
  }

  function exitEditMode(save) {
    if (!_editingField) return;
    const { id, poly } = _editingField;

    if (save && _editMarkers.length >= 3) {
      // Update the field's points from current marker positions
      const newLatlngs = _editMarkers.map(m => m.getLatLng());
      const newPoints  = newLatlngs.map(ll => ({ lat: ll.lat, lng: ll.lng }));
      poly._field.points = newPoints;
      // Recalculate acres from updated points
      const editedAcres = GeoUtils.calcAcres(newPoints);
      // Re-style as selected
      poly.setStyle(fieldStyle(true));
      // Update selectedFields entry — replace first ring KML, preserve any extra rings
      const sf = selectedFields.find(s => s.id === id);
      if (sf) {
        sf.points = newPoints;
        const newRingKML = GeoUtils.pointsToKML(newPoints);
        if (sf.kml && sf.kml.includes('<Polygon>') ) {
          // Replace the first <Polygon>...</Polygon> block with the edited ring
          sf.kml = sf.kml.replace(/<Polygon>[\s\S]*?<\/Polygon>/, newRingKML);
          // Recalculate total acres across all rings in updated KML
          const allRings = GeoUtils.parseKMLAllRings(sf.kml);
          sf.acres = allRings.reduce((sum, r) => sum + GeoUtils.calcAcres(r), 0).toFixed(1);
        } else {
          sf.kml   = newRingKML;
          sf.acres = editedAcres.toFixed(1);
        }
        poly._field.acres = sf.acres;
      }
      setStatus('Polygon updated (' + (sf?.acres || editedAcres.toFixed(1)) + ' ac) — confirm to save changes');
    } else {
      // Restore original
      poly.setLatLngs([_editingField.originalPoints.map(p => [p.lat, p.lng])]);
      poly.setStyle(fieldStyle(true));
    }

    _editMarkers.forEach(m => map.removeLayer(m));
    _editMarkers = [];
    map.off('click', onEditMapClick);
    map.on('click', onMapClick);
    _editingField = null;
    document.getElementById('mapEditFinishBtn')?.style.setProperty('display', 'none');
  }

  function toggleEditMode() { exitEditMode(true); }

  // ── PUBLIC ────────────────────────────────────────────────────────────────
  return { open, close, confirm, toggleDraw, finishDrawBtn, removeSelected, renameSelected, showPastePanel, applyPastedKML, searchAddress, loadShapefileInput, toggleCLU, toggleEditMode };

})();
