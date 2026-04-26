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
    // Handles single <Polygon>, multi-polygon (concatenated), and full KML files
    if (!text.includes('<coordinates') && !text.includes('<Coordinates')) return null;
    try {
      const matches = [...text.matchAll(/<coordinates[^>]*>([\s\S]*?)<\/coordinates>/gi)];
      if (!matches.length) return null;
      if (matches.length === 1) return coordStringToPoints(matches[0][1]);
      // Multi-polygon: return points from the LARGEST ring (for acres/centroid calcs)
      // All rings are stored; display uses polygonToSVGMulti
      const rings = matches.map(m => coordStringToPoints(m[1])).filter(Boolean);
      if (!rings.length) return null;
      // Return largest ring as primary (for calcAcres, centroid)
      return rings.reduce((best, r) => r.length > best.length ? r : best, rings[0]);
    } catch(e) { return null; }
  }

  function parseKMLAllRings(text) {
    // Returns ALL rings from a multi-polygon KML as array of point arrays
    if (!text || !text.includes('<coordinates')) return [];
    const matches = [...text.matchAll(/<coordinates[^>]*>([\s\S]*?)<\/coordinates>/gi)];
    return matches.map(m => coordStringToPoints(m[1])).filter(r => r && r.length >= 3);
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


  // Render SVG from a KML string that may contain multiple polygons
  function polygonToSVGFromKML(kml, opts = {}) {
    const rings = parseKMLAllRings(kml);
    if (!rings.length) return '';
    if (rings.length === 1) return polygonToSVG(rings[0], opts);

    // Multi-polygon: compute combined bounds
    const allPts = rings.flat();
    const lats = allPts.map(p => p.lat), lngs = allPts.map(p => p.lng);
    const minLat = Math.min(...lats), maxLat = Math.max(...lats);
    const minLng = Math.min(...lngs), maxLng = Math.max(...lngs);
    const dLat = maxLat - minLat || 0.0001;
    const dLng = maxLng - minLng || 0.0001;
    const midLat = (minLat + maxLat) / 2;
    const lngScale = Math.cos(midLat * Math.PI / 180);
    const dLngM = dLng * lngScale;
    const aspect = dLngM / dLat;
    const H = opts.height || 120;
    const W = Math.max(80, Math.min(280, Math.round(H * aspect)));
    const pad = opts.padding || 10;
    const color = opts.color || '#4FC3F7';
    const fill  = opts.fill  || 'rgba(79,195,247,0.15)';

    const toSVG = p => {
      const x = pad + ((p.lng - minLng) * lngScale / dLngM) * (W - 2*pad);
      const y = H - pad - ((p.lat - minLat) / dLat) * (H - 2*pad);
      return `${x.toFixed(1)},${y.toFixed(1)}`;
    };

    const polys = rings.map(ring =>
      `<polygon points="${ring.map(toSVG).join(' ')}" fill="${fill}" stroke="${color}" stroke-width="1.5" stroke-linejoin="round"/>`
    ).join(' ');

    return `<svg viewBox="0 0 ${W} ${H}" xmlns="http://www.w3.org/2000/svg" style="width:100%;max-width:${W}px;height:auto;display:block">
  ${polys}
</svg>`;
  }

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


  // ── COORDINATE REPROJECTION ───────────────────────────────────────────────
  // Inverse Albers Equal Area Conic → WGS84 (for USDA ACPF / FSA CLU files)
  // PRJ: Origin=23N, CM=-96, SP1=29.5, SP2=45.5, NAD83, FE=0, FN=0

  function albersToWGS84(x, y) {
    const toRad = d => d * Math.PI / 180;
    const phi1 = toRad(29.5), phi2 = toRad(45.5), phi0 = toRad(23.0);
    const lam0 = toRad(-96.0);
    const a = 6378137.0, f = 1/298.257222101;
    const e2 = 2*f - f*f, e = Math.sqrt(e2);

    function alpha(phi) {
      const s = Math.sin(phi);
      return (1-e2)*(s/(1-e2*s*s) - Math.log((1-e*s)/(1+e*s))/(2*e));
    }
    function mfn(phi) { const s=Math.sin(phi); return Math.cos(phi)/Math.sqrt(1-e2*s*s); }

    const m1=mfn(phi1), m2=mfn(phi2);
    const a1=alpha(phi1), a2=alpha(phi2), a0=alpha(phi0);
    const n = (m1*m1 - m2*m2) / (a2 - a1);
    const C = m1*m1 + n*a1;
    const rho0 = a * Math.sqrt(C - n*a0) / n;

    let rho = Math.sqrt(x*x + (rho0-y)*(rho0-y));
    if (n < 0) rho = -rho;
    const theta = Math.atan2(x, rho0 - y);
    const alp = (C - (rho*n/a)*(rho*n/a)) / n;

    let phi = Math.asin(alp / 2);
    for (let i = 0; i < 12; i++) {
      const s = Math.sin(phi);
      const d = ((1-e2*s*s)*(1-e2*s*s) / (2*Math.cos(phi))) *
                (alp/(1-e2) - s/(1-e2*s*s) + Math.log((1-e*s)/(1+e*s))/(2*e));
      phi += d;
      if (Math.abs(d) < 1e-12) break;
    }
    const lam = lam0 + theta / n;
    return { lat: phi * 180/Math.PI, lng: lam * 180/Math.PI };
  }

  // Detect if a coordinate array is in Albers meters (vs WGS84 degrees)
  // Albers coords for CONUS are ~300000–900000 (x) and ~1500000–2500000 (y)
  function isAlbersCoords(ring) {
    if (!ring || !ring.length) return false;
    const x = Math.abs(ring[0][0]), y = Math.abs(ring[0][1]);
    return x > 1000 || y > 1000;  // WGS84 is always < 180/90
  }

  // Convert a ring of coordinates to WGS84 if in Albers
  function normalizeRing(ring) {
    if (!ring || !ring.length) return ring;
    if (!isAlbersCoords(ring)) return ring;
    return ring.map(c => {
      const p = albersToWGS84(c[0], c[1]);
      return [p.lng, p.lat];
    });
  }

  // ── POLYGON SIMPLIFICATION (Ramer-Douglas-Peucker) ───────────────────────
  // Reduces point count while preserving shape. epsilon in degrees (~0.00001 = ~1m)

  function simplifyPolygon(points, epsilon) {
    if (!points || points.length < 3) return points;
    epsilon = epsilon || 0.00005; // ~5 meters — good for field boundaries
    return rdp(points, epsilon);
  }

  function rdp(pts, eps) {
    if (pts.length <= 2) return pts;
    let maxDist = 0, maxIdx = 0;
    const start = pts[0], end = pts[pts.length-1];
    for (let i = 1; i < pts.length-1; i++) {
      const d = pointLineDistance(pts[i], start, end);
      if (d > maxDist) { maxDist = d; maxIdx = i; }
    }
    if (maxDist > eps) {
      const left  = rdp(pts.slice(0, maxIdx+1), eps);
      const right = rdp(pts.slice(maxIdx), eps);
      return left.slice(0, -1).concat(right);
    }
    return [start, end];
  }

  function pointLineDistance(p, a, b) {
    const dx = b.lng - a.lng, dy = b.lat - a.lat;
    if (dx === 0 && dy === 0) {
      return Math.sqrt((p.lng-a.lng)**2 + (p.lat-a.lat)**2);
    }
    const t = ((p.lng-a.lng)*dx + (p.lat-a.lat)*dy) / (dx*dx + dy*dy);
    const tc = Math.max(0, Math.min(1, t));
    return Math.sqrt((p.lng - (a.lng+tc*dx))**2 + (p.lat - (a.lat+tc*dy))**2);
  }

  // ── PUBLIC API ────────────────────────────────────────────────────────────
  return {
    parsePolygon,
    parseKMLAllRings,
    calcAcres,
    centroid,
    pointsToKML,
    pointsToKMLFile,
    mergePolygons,
    polygonToSVG,
    polygonToSVGFromKML,
    fetchCLU,
    geocodeAddress,
    albersToWGS84,
    normalizeRing,
    isAlbersCoords,
    simplifyPolygon,
  };

})();
