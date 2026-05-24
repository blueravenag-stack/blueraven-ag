// ── GDU.JS — Blue Raven Ag ────────────────────────────────────────────────────
// Corn Growing Degree Unit calculator for fungicide timing prediction
// Uses Open-Meteo API (free, no API key required)
// ─────────────────────────────────────────────────────────────────────────────

window.GDUCalc = (() => {

  // ── CONSTANTS ─────────────────────────────────────────────────────────────
  const BASE_TEMP   = 50;  // °F — base temperature for corn GDU
  const MAX_TEMP    = 86;  // °F — ceiling temperature for corn GDU
  const GDU_PER_RM  = 13;  // GDU to reach VT per RM unit (IL empirical)
  // GDU offsets from VT (fixed, not percentage-based)
  const GDU_VT_TO_R1 = 60;   // VT → R1 silking (~2-3 days in mid-summer)
  const GDU_VT_TO_R2 = 310;  // VT → R2 blister (60 + 250)
  // Alert window starts at 80% of GDU to VT so pilot can schedule ahead
  const FUNG_ALERT  = 0.80;

  // ── CORE GDU MATH ─────────────────────────────────────────────────────────

  function calcDailyGDU(maxF, minF) {
    const hi  = Math.min(maxF, MAX_TEMP);
    const lo  = Math.max(minF, BASE_TEMP);
    const avg = (hi + lo) / 2;
    return Math.max(0, avg - BASE_TEMP);
  }

  function calcCumulativeGDU(dailyTemps) {
    // dailyTemps: [{date, maxF, minF}]
    let cum = 0;
    return dailyTemps.map(d => {
      cum += calcDailyGDU(d.maxF, d.minF);
      return { ...d, dailyGDU: calcDailyGDU(d.maxF, d.minF), cumGDU: cum };
    });
  }

  function gduToVT(rm) {
    return rm * GDU_PER_RM;
  }

  function fungicideWindow(rm) {
    const vtGDU = gduToVT(rm);
    const r1GDU = vtGDU + GDU_VT_TO_R1;          // R1 silking (~60 GDU after VT)
    const r2GDU = vtGDU + GDU_VT_TO_R2;          // R2 blister (~310 GDU after VT) — hard cutoff
    // Ideal application: midpoint of VT→R1 (best disease protection, canopy access)
    const targetGDU = Math.round(vtGDU + GDU_VT_TO_R1 / 2);  // ~30 GDU after VT
    return {
      start:  Math.round(vtGDU * FUNG_ALERT),    // 80% of VT = early scheduling alert
      target: targetGDU,                          // ideal = midpoint VT→R1
      end:    r2GDU,                              // R2 = hard cutoff, last effective timing
      vtGDU,
      r1GDU,
      r2GDU,
    };
  }

  // ── OPEN-METEO API ────────────────────────────────────────────────────────
  // Free weather API — no key needed
  // Historical data available from 1940, forecast 7-16 days ahead

  // ── THREE-TIER WEATHER STRATEGY ─────────────────────────────────────────
  // Tier 1: ERA5 archive          planting date → 7 days ago   (historical actual)
  // Tier 2: GFS/HRRR forecast     7 days ago → +14 days        (high-res forecast)
  // Tier 3: ECMWF SEAS5 seasonal  +14 days → VT projection     (6-month ensemble)

  async function fetchWeather(lat, lng, startDate, endDate) {
    const today    = new Date().toISOString().split('T')[0];
    const t1End    = addDays(today, -7);   // archive reliable up to 7 days ago
    const t2Start  = addDays(today, -2);   // forecast API covers from 2 days ago
    const t2End    = addDays(today, 14);   // GFS/HRRR goes 14 days ahead
    const t3Start  = addDays(today, 15);   // seasonal picks up after day 14
    const results  = [];

    // TIER 1 — ERA5 archive (oldest reliable historical)
    if (startDate <= t1End) {
      const end = endDate < t1End ? endDate : t1End;
      try {
        const hist = await fetchHistorical(lat, lng, startDate, end);
        results.push(...hist.map(d => ({ ...d, tier: 1, isForecast: false })));
      } catch(e) { console.warn('Tier 1 (archive) failed:', e.message); }
    }

    // TIER 2 — GFS/HRRR forecast + past_days=92 for recent history
    // Use startDate directly so recent plantings (< 92 days) get full coverage
    const t2S = startDate; // past_days=92 covers up to 3 months back
    const t2E = endDate < t2End ? endDate : t2End;
    if (t2S <= t2E) {
      try {
        const fcst = await fetchForecast(lat, lng, t2S, t2E);
        results.push(...fcst.map(d => ({ ...d, tier: 2 })));
      } catch(e) { console.warn('Tier 2 (forecast) failed:', e.message); }
    }

    // TIER 3 — ECMWF SEAS5 seasonal (beyond 14 days, up to 6 months)
    const t3S = startDate > t3Start ? startDate : t3Start;
    if (endDate >= t3Start && t3S <= endDate) {
      try {
        const seas = await fetchSeasonal(lat, lng, t3S, endDate);
        results.push(...seas.map(d => ({ ...d, tier: 3, isForecast: true })));
      } catch(e) { console.warn('Tier 3 (seasonal) failed:', e.message); }
    }

    // Merge: deduplicate by date, prefer lower tier number (more accurate)
    const byDate = {};
    // Add in reverse priority so lower tier (more accurate) wins on overwrite
    [...results].sort((a, b) => b.tier - a.tier).forEach(d => {
      byDate[d.date] = d;
    });

    return Object.values(byDate)
      .filter(d => d.date >= startDate && d.date <= endDate)
      .sort((a, b) => a.date.localeCompare(b.date));
  }

  async function fetchHistorical(lat, lng, start, end) {
    const url = `https://archive-api.open-meteo.com/v1/archive?` +
      `latitude=${lat}&longitude=${lng}` +
      `&start_date=${start}&end_date=${end}` +
      `&daily=temperature_2m_max,temperature_2m_min` +
      `&temperature_unit=fahrenheit&timezone=America%2FChicago`;
    const res  = await fetch(url);
    if (!res.ok) throw new Error(`Archive API HTTP ${res.status}`);
    const data = await res.json();
    if (data.error) throw new Error(data.reason || 'Archive API error');
    return zipTemps(data);
  }

  async function fetchForecast(lat, lng, start, end) {
    // past_days=92 lets forecast API serve ~3 months of recent history
    // This is the primary fallback when archive fails for recent plantings
    const url = `https://api.open-meteo.com/v1/forecast?` +
      `latitude=${lat}&longitude=${lng}` +
      `&daily=temperature_2m_max,temperature_2m_min` +
      `&temperature_unit=fahrenheit&timezone=America%2FChicago` +
      `&forecast_days=16&past_days=92`;
    const res  = await fetch(url);
    if (!res.ok) throw new Error(`Forecast API HTTP ${res.status}`);
    const data = await res.json();
    if (data.error) throw new Error(data.reason || 'Forecast API error');
    const all   = zipTemps(data);
    const today2= new Date().toISOString().split('T')[0];
    all.forEach(d => { d.isForecast = d.date > today2; });
    return all.filter(d => d.date >= start && d.date <= end);
  }

  async function fetchSeasonal(lat, lng, start, end) {
    // ECMWF SEAS5 seasonal forecast — 51 ensemble members averaged
    // Returns daily max/min temps for up to 6 months ahead
    const url = `https://seasonal-api.open-meteo.com/v1/seasonal?` +
      `latitude=${lat}&longitude=${lng}` +
      `&daily=temperature_2m_max,temperature_2m_min` +
      `&forecast_months=6` +
      `&temperature_unit=fahrenheit` +
      `&timezone=America%2FChicago`;
    const res  = await fetch(url);
    if (!res.ok) throw new Error(`Seasonal API HTTP ${res.status}`);
    const data = await res.json();
    if (data.error) throw new Error(data.reason || 'Seasonal API error');

    const daily = data.daily || {};
    const dates = daily.time || [];

    // Seasonal returns ensemble members: temperature_2m_max_member01, _member02, etc.
    // OR temperature_2m_max_mean if means are available. Try mean first, fall back to avg.
    const maxKeys = Object.keys(daily).filter(k => k.startsWith('temperature_2m_max'));
    const minKeys = Object.keys(daily).filter(k => k.startsWith('temperature_2m_min'));

    return dates.map((date, i) => {
      // Average all ensemble members for this day
      const maxVals = maxKeys.map(k => daily[k]?.[i]).filter(v => v != null);
      const minVals = minKeys.map(k => daily[k]?.[i]).filter(v => v != null);
      const maxF = maxVals.length ? maxVals.reduce((a, b) => a + b, 0) / maxVals.length : null;
      const minF = minVals.length ? minVals.reduce((a, b) => a + b, 0) / minVals.length : null;
      return { date, maxF, minF, isForecast: true, memberCount: maxVals.length };
    })
    .filter(d => d.maxF !== null && d.date >= start && d.date <= end);
  }

  function zipTemps(data) {
    const dates = data?.daily?.time || [];
    const maxes = data?.daily?.temperature_2m_max || [];
    const mins  = data?.daily?.temperature_2m_min || [];
    return dates.map((date, i) => ({
      date,
      maxF: maxes[i] ?? null,
      minF: mins[i]  ?? null,
      isForecast: false,
    })).filter(d => d.maxF !== null);
  }

  // ── PROJECT FORWARD ───────────────────────────────────────────────────────
  // After forecast data runs out, use 10-year historical average for that date

  function projectToTarget(withGDU, targetGDU, historicalAvg) {
    // withGDU: cumulative GDU array through today + forecast
    // If target already reached, return the date it was reached
    const reached = withGDU.find(d => d.cumGDU >= targetGDU);
    if (reached) return { date: reached.date, projected: false, gduAtDate: reached.cumGDU };

    // Project forward using historical avg daily GDU
    const lastEntry = withGDU[withGDU.length - 1];
    if (!lastEntry) return null;
    let cumGDU = lastEntry.cumGDU;
    let date   = lastEntry.date;
    let days   = 0;

    while (cumGDU < targetGDU && days < 120) {
      date    = addDays(date, 1);
      const avgGDU = historicalAvg || 12; // ~12 GDU/day mid-season IL
      cumGDU += avgGDU;
      days++;
    }

    return { date, projected: true, gduAtDate: Math.round(cumGDU), daysProjected: days };
  }

  // ── HISTORICAL AVERAGE DAILY GDU ─────────────────────────────────────────
  // Simple lookup: avg daily GDU by month for central IL (Medora area)
  const IL_AVG_DAILY_GDU = {
    '01':  0, '02':  0, '03':  2, '04':  5,
    '05': 10, '06': 16, '07': 18, '08': 16,
    '09': 11, '10':  5, '11':  1, '12':  0,
  };

  function avgGDUForDate(dateStr) {
    const month = dateStr.slice(5, 7);
    return IL_AVG_DAILY_GDU[month] || 10;
  }

  // ── MAIN ENTRY: analyzeOrder ───────────────────────────────────────────────
  // Returns full GDU analysis for one corn order

  async function analyzeOrder(order, field, fetchFn) {
    const rm          = parseFloat(order.RelativeMaturity);
    const plantDate   = order.PlantingDate;

    if (!plantDate || !rm) {
      return { error: 'Order needs Planting Date and RM set' };
    }

    // Use field centroid for weather, or Medora IL as fallback
    const lat = parseFloat(field?.CentroidLat || 39.17);
    const lng = parseFloat(field?.CentroidLng || -90.14);

    const today     = new Date().toISOString().split('T')[0];
    // Extend end date far enough to cover VT for any RM (max ~210 days from early April)
    const fcstEnd   = addDays(today, 180);

    try {
      // Fetch weather from planting through seasonal forecast (6 months out)
      const temps = fetchFn
        ? await fetchFn(lat, lng, plantDate, fcstEnd)
        : await fetchWeather(lat, lng, plantDate, fcstEnd);

      // Mark forecast days
      temps.forEach(d => { d.isForecast = d.date > today; });

      // Calculate cumulative GDU
      const withGDU = calcCumulativeGDU(temps);
      const currentGDU = withGDU.filter(d => !d.isForecast).slice(-1)[0]?.cumGDU || 0;

      // Fungicide window
      const window = fungicideWindow(rm);

      // Project target date using month-aware avg GDU (falls back automatically)
      const lastDate   = withGDU.length ? withGDU[withGDU.length-1].date : new Date().toISOString().split('T')[0];
      const avgGDU     = avgGDUForDate(lastDate);

      const targetResult   = projectToTarget(withGDU, window.target, avgGDU);
      const windowStart    = projectToTarget(withGDU, window.start,  avgGDU);
      const vtResult       = projectToTarget(withGDU, window.vtGDU,  avgGDU);
      const r1Result       = projectToTarget(withGDU, window.r1GDU,  avgGDU);
      const windowEnd      = projectToTarget(withGDU, window.r2GDU,  avgGDU);

      // GDU progress as percentage toward VT
      const pctToVT    = Math.min(100, Math.round(currentGDU / window.vtGDU * 100));
      const pctToTarget= Math.min(100, Math.round(currentGDU / window.target * 100));

      // Stage estimate based on GDU
      const stage = estimateStage(currentGDU, rm);

      // Data quality summary
      const histDays   = withGDU.filter(d => d.tier === 1).length;
      const fcstDays   = withGDU.filter(d => d.tier === 2 && d.isForecast).length;
      const seasDays   = withGDU.filter(d => d.tier === 3).length;
      const tier2Days  = withGDU.filter(d => d.tier === 2 && !d.isForecast).length;
      const totalDays = withGDU.length;
      const daysSincePlant = Math.round((new Date() - new Date(plantDate + 'T12:00:00')) / 86400000);
      const coverage  = totalDays > 0 ? Math.round(totalDays / Math.max(daysSincePlant, 1) * 100) : 0;

      return {
        orderId:       order.OrderID,
        customerName:  order.CustomerName,
        fieldNames:    field?.FieldName || '—',
        plantDate,
        rm,
        currentGDU:    Math.round(currentGDU),
        vtGDU:         window.vtGDU,
        r1GDU:         window.r1GDU,
        r2GDU:         window.r2GDU,
        targetGDU:     window.target,
        windowStartGDU:window.start,
        windowEndGDU:  window.r2GDU,
        pctToVT,
        pctToTarget,
        stage,
        windowStart:   windowStart?.date,
        vtDate:        vtResult?.date,
        targetDate:    targetResult?.date,
        r1Date:        r1Result?.date,
        windowEnd:     windowEnd?.date,
        targetProjected: targetResult?.projected,
        withGDU,
        lat, lng,
        // Data quality
        histDays, fcstDays, seasDays, tier2Days, totalDays, daysSincePlant, coverage,
      };
    } catch(e) {
      return { error: 'Weather fetch failed: ' + e.message, orderId: order.OrderID };
    }
  }

  function estimateStage(cumGDU, rm) {
    const vtGDU = gduToVT(rm);
    const r1GDU = vtGDU + GDU_VT_TO_R1;
    const r2GDU = vtGDU + GDU_VT_TO_R2;
    const pct   = cumGDU / vtGDU;
    if (pct < 0.15) return 'Emergence / V2';
    if (pct < 0.35) return 'V3–V5';
    if (pct < 0.55) return 'V6–V8';
    if (pct < 0.75) return 'V9–V12';
    if (pct < 0.88) return 'V13–V15 (approaching tassel)';
    if (cumGDU < vtGDU) return '⚡ Pre-VT (schedule now)';
    if (cumGDU < r1GDU) return '✅ VT — Ideal fungicide window';
    if (cumGDU < r2GDU) return '⚠️ R1 Silking — apply soon';
    return '🚫 R2+ Blister — window closed';
  }

  // ── HELPERS ───────────────────────────────────────────────────────────────

  function addDays(dateStr, n) {
    const d = new Date(dateStr + 'T12:00:00');
    d.setDate(d.getDate() + n);
    return d.toISOString().split('T')[0];
  }

  function fmtDate(str) {
    if (!str) return '—';
    const d = new Date(str + 'T12:00:00');
    return d.toLocaleDateString('en-US', { month: 'short', day: 'numeric' });
  }

  function urgencyClass(order) {
    const pct = order.pctToTarget;
    if (!pct) return '';
    if (pct >= 95) return 'urgent';    // window closing
    if (pct >= 80) return 'soon';      // coming up
    if (pct >= 60) return 'tracking';  // getting close
    return 'early';
  }

  // ── PUBLIC ────────────────────────────────────────────────────────────────
  return { analyzeOrder, fetchWeatherPublic: fetchWeather, calcDailyGDU, calcCumulativeGDU, fungicideWindow, fmtDate, urgencyClass, addDays };

})();
