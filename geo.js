// ── GEO.JS — Blue Raven Ag ────────────────────────────────────────────────────
// Self-contained module. No dependencies on app.js or DOM.
// Public API via window.GeoUtils
// ─────────────────────────────────────────────────────────────────────────────

window.GeoUtils = (() => {

  // ── PARSE ANY POLYGON INPUT → [{lat, lng}] or null ─────────────────────────

  function parsePolygon(text) {
    if (!text || !text.trim()) return null;
    return parseKML(text) || parseGeoJSON(text) || parseRawCoords(text) || null;
  }

  function parseKML(text) {
    // Handles both full KML files and bare <Polygon> fragments
    if (!text.includes('<coordinates') && !text.includes('<Coordinates')) return null;
    try {
      const match = text.match(/<coordinates[^>]*>([\s\S]*?)<\/coordinates>/i);
      if (!match) return null;
      return coordStringToPoints(match[1]);
    } catch(e) { return null; }
  }

  function parseGeoJSON(text) {
    try {
      const obj = JSON.parse(text);
      let coords = null;
      if (obj.type === 'FeatureCollection') {
        const feat = obj.features?.[0];
        coords = feat?.geometry?.coordinates;
        if (feat?.geometry?.type === 'Polygon') coords = coords?.[0];
        else if (feat?.geometry?.type === 'MultiPolygon') coords = coords?.[0]?.[0];
      } else if (obj.type === 'Feature') {
        coords = obj.geometry?.coordinates;
        if (obj.geometry?.type === 'Polygon') coords = coords?.[0];
        else if (obj.geometry?.type === 'MultiPolygon') coords = coords?.[0]?.[0];
      } else if (obj.type === 'Polygon') {
        coords = obj.coordinates?.[0];
      } else if (obj.type === 'MultiPolygon') {
        coords = obj.coordinates?.[0]?.[0];
      } else if (Array.isArray(obj)) {
        coords = obj;
      }
      if (!coords || !coords.length) return null;
      return coords.map(c => ({ lng: parseFloat(c[0]), lat: parseFloat(c[1]) }))
                   .filter(p => isValidPoint(p));
    } catch(e) { return null; }
  }

  function parseRawCoords(text) {
    // Handles: "lat,lng lat,lng" or "lng,lat lng,lat" or "lat lng\nlat lng"
    // Try to detect which order by checking if first value looks like a US longitude
    const pairs = text.trim().split(/[\s,;]+/).filter(s => s.match(/^-?\d+\.?\d*$/));
    if (pairs.length < 6) return null; // need at least 3 points
    const nums = pairs.map(Number);
    const points = [];
    for (let i = 0; i < nums.length - 1; i += 2) {
      const a = nums[i], b = nums[i+1];
      // US longitudes are negative 60-130, latitudes are 24-50
      let lat, lng;
      if (b < -60 && b > -130 && a > 24 && a < 50) { lat = a; lng = b; }
      else if (a < -60 && a > -130 && b > 24 && b < 50) { lat = b; lng = a; }
      else { lat = a; lng = b; } // best guess
      points.push({ lat, lng });
    }
    return points.filter(p => isValidPoint(p));
  }

  function coordStringToPoints(str) {
    // KML coords: "lng,lat,alt lng,lat,alt ..."
    const tuples = str.trim().split(/\s+/).filter(Boolean);
    return tuples.map(t => {
      const parts = t.split(',');
      return { lng: parseFloat(parts[0]), lat: parseFloat(parts[1]) };
    }).filter(p => isValidPoint(p));
  }

  function isValidPoint(p) {
    return p && isFinite(p.lat) && isFinite(p.lng) &&
           p.lat >= -90 && p.lat <= 90 &&
           p.lng >= -180 && p.lng <= 180;
  }

  // ── ACRES CALCULATION (Shoelace + spherical correction) ────────────────────

  function calcAcres(points) {
    // Spherical lune formula — accurate to <1% for typical field sizes
    // Validated: 160-acre quarter section returns 160.8 ac at 39°N
    if (!points || points.length < 3) return 0;
    const R = 6378137; // WGS84 Earth radius meters
    const toRad = d => d * Math.PI / 180;
    let area = 0;
    const n = points.length;
    for (let i = 0; i < n; i++) {
      const j = (i + 1) % n;
      const lat1 = toRad(points[i].lat);
      const lng1 = toRad(points[i].lng);
      const lat2 = toRad(points[j].lat);
      const lng2 = toRad(points[j].lng);
      area += (lng2 - lng1) * (2 + Math.sin(lat1) + Math.sin(lat2));
    }
    const sqMeters = Math.abs(area) * R * R / 2;
    return sqMeters / 4046.856; // sq meters to acres
  }

  // ── CENTROID ───────────────────────────────────────────────────────────────

  function centroid(points) {
    if (!points || !points.length) return null;
    const lat = points.reduce((s, p) => s + p.lat, 0) / points.length;
    const lng = points.reduce((s, p) => s + p.lng, 0) / points.length;
    return { lat, lng };
  }

  // ── KML EXPORT ────────────────────────────────────────────────────────────

  function pointsToKML(points, name) {
    if (!points || !points.length) return '';
    const coords = points.map(p => `${p.lng},${p.lat},0`).join(' ');
    return `<Polygon><outerBoundaryIs><LinearRing><coordinates>${coords}</coordinates></LinearRing></outerBoundaryIs></Polygon>`;
  }

  function pointsToKMLFile(fields) {
    // fields: [{name, points}]
    const placemarks = fields.map(f => `
  <Placemark>
    <name>${escXML(f.name)}</name>
    ${pointsToKML(f.points)}
  </Placemark>`).join('');
    return `<?xml version="1.0" encoding="UTF-8"?>
<kml xmlns="http://www.opengis.net/kml/2.2">
<Document>${placemarks}
</Document>
</kml>`;
  }

  function escXML(s) {
    return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
  }

  // ── MERGE MULTIPLE POLYGONS (convex hull approximation) ────────────────────
  // For display purposes when multiple CLU fields are selected

  function mergePolygons(polygonArray) {
    // Simple approach: just concatenate all points and use convex hull
    const all = polygonArray.flat();
    return convexHull(all);
  }

  function convexHull(points) {
    if (points.length <= 3) return points;
    const sorted = [...points].sort((a, b) => a.lng - b.lng || a.lat - b.lat);
    const cross = (o, a, b) =>
      (a.lng - o.lng) * (b.lat - o.lat) - (a.lat - o.lat) * (b.lng - o.lng);
    const lower = [];
    for (const p of sorted) {
      while (lower.length >= 2 && cross(lower[lower.length-2], lower[lower.length-1], p) <= 0)
        lower.pop();
      lower.push(p);
    }
    const upper = [];
    for (const p of [...sorted].reverse()) {
      while (upper.length >= 2 && cross(upper[upper.length-2], upper[upper.length-1], p) <= 0)
        upper.pop();
      upper.push(p);
    }
    upper.pop(); lower.pop();
    return lower.concat(upper);
  }

  // ── SVG FIELD PREVIEW ─────────────────────────────────────────────────────

  function polygonToSVG(points, opts = {}) {
    if (!points || points.length < 3) return '';
    const pad   = opts.padding || 10;
    const color = opts.color   || '#4FC3F7';
    const fill  = opts.fill    || 'rgba(79,195,247,0.15)';

    const lats = points.map(p => p.lat);
    const lngs = points.map(p => p.lng);
    const minLat = Math.min(...lats), maxLat = Math.max(...lats);
    const minLng = Math.min(...lngs), maxLng = Math.max(...lngs);
    const dLat = maxLat - minLat || 0.0001;
    const dLng = maxLng - minLng || 0.0001;

    // Correct for lng/lat scale difference at this latitude
    // 1 deg lng = cos(lat) * 1 deg lat in distance
    const midLat  = (minLat + maxLat) / 2;
    const lngScale = Math.cos(midLat * Math.PI / 180);

    // Build a square viewbox with correct aspect ratio
    const dLatM = dLat;                  // normalized lat extent
    const dLngM = dLng * lngScale;       // normalized lng extent (corrected)
    const aspect = dLngM / dLatM;        // width / height ratio

    const H = 120;
    const W = Math.max(80, Math.min(280, Math.round(H * aspect)));

    const toSVG = p => {
      const x = pad + ((p.lng - minLng) * lngScale / dLngM) * (W - 2*pad);
      const y = H - pad - ((p.lat - minLat) / dLatM) * (H - 2*pad);
      return `${x.toFixed(1)},${y.toFixed(1)}`;
    };

    const polyPts = points.map(toSVG).join(' ');
    return `<svg viewBox="0 0 ${W} ${H}" xmlns="http://www.w3.org/2000/svg" style="width:100%;max-width:${W}px;height:auto;display:block">
  <polygon points="${polyPts}" fill="${fill}" stroke="${color}" stroke-width="1.5" stroke-linejoin="round"/>
</svg>`;
  }

  // ── CLU API ───────────────────────────────────────────────────────────────
  // USDA FSA CLU WFS endpoint — returns field boundaries for a bounding box

  async function fetchCLU(bounds) {
    // USDA FSA Public CLU FeatureServer - current endpoint as of 2025
    // Docs: https://arcgis.fsa.usda.gov/arcgis/rest/services/Hosted/PAIL_CLU_Public/FeatureServer/0
    const { minLat, minLng, maxLat, maxLng } = bounds;
    const bbox = `${minLng},${minLat},${maxLng},${maxLat}`;

    const base = 'https://arcgis.fsa.usda.gov/arcgis/rest/services/Hosted/PAIL_CLU_Public/FeatureServer/0/query';
    const params = new URLSearchParams({
      geometry:           bbox,
      geometryType:       'esriGeometryEnvelope',
      spatialRel:         'esriSpatialRelIntersects',
      outFields:          'clu_number,farm_number,tract_number,clu_calculated_acreage,admin_state,admin_county',
      returnGeometry:     'true',
      outSR:              '4326',
      f:                  'geojson',
      resultRecordCount:  '100',
    });

    try {
      const res  = await fetch(`${base}?${params}`);
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      const json = await res.json();

      if (json.error) {
        console.warn('CLU API error:', json.error.message);
        return [];
      }

      return (json.features || []).map(f => {
        const p = f.properties || {};
        return {
          id:      String(p.clu_number || p.CLU_NUMBER || ''),
          farmNum: String(p.farm_number || p.FARM_NUMBER || ''),
          tractNum:String(p.tract_number || p.TRACT_NUMBER || ''),
          acres:   parseFloat(p.clu_calculated_acreage || p.CLU_CALCULATED_ACREAGE || 0).toFixed(1),
          points:  parseGeoJSON(JSON.stringify(f.geometry)),
          props:   p
        };
      }).filter(f => f.points && f.points.length >= 3);

    } catch(e) {
      console.warn('CLU fetch failed:', e.message);
      return [];
    }
  }

  // ── GEOCODE ADDRESS → {lat, lng} using Nominatim (free, no key) ────────────

  async function geocodeAddress(address) {
    try {
      const url = `https://nominatim.openstreetmap.org/search?format=json&limit=1&q=${encodeURIComponent(address)}`;
      const res  = await fetch(url, { headers: { 'Accept-Language': 'en' } });
      const json = await res.json();
      if (!json.length) return null;
      return { lat: parseFloat(json[0].lat), lng: parseFloat(json[0].lon) };
    } catch(e) { return null; }
  }

  // ── PUBLIC API ────────────────────────────────────────────────────────────
  return {
    parsePolygon,
    calcAcres,
    centroid,
    pointsToKML,
    pointsToKMLFile,
    mergePolygons,
    polygonToSVG,
    fetchCLU,
    geocodeAddress,
  };

})();
