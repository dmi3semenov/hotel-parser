'use strict';

const ExcelJS = require('exceljs');
const https   = require('https');
const fs      = require('fs');
const path    = require('path');

// ── Geocoding (Nominatim) ─────────────────────────────────────────────
function geocodeRequest(query) {
  return new Promise(resolve => {
    const url = `https://nominatim.openstreetmap.org/search?q=${encodeURIComponent(query)}&format=json&limit=1&countrycodes=by`;
    https.get(url, { headers: { 'User-Agent': 'hotel-parser/1.0 (educational)' } }, res => {
      let data = '';
      res.on('data', c => data += c);
      res.on('end', () => {
        try {
          const r = JSON.parse(data);
          resolve(r.length ? { lat: parseFloat(r[0].lat), lng: parseFloat(r[0].lon) } : null);
        } catch { resolve(null); }
      });
    }).on('error', () => resolve(null));
  });
}

function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }

// HEAD-проверка URL: возвращает true если 2xx-3xx статус. Используется для отсева битых фото.
function isUrlAlive(url) {
  return new Promise(resolve => {
    try {
      const u = new URL(url);
      const opts = {
        method: 'HEAD',
        host: u.hostname,
        path: u.pathname + u.search,
        headers: { 'User-Agent': 'Mozilla/5.0 hotel-parser', 'Accept': 'image/*,*/*' },
        timeout: 6000,
      };
      const req = https.request(opts, res => {
        const ok = res.statusCode >= 200 && res.statusCode < 400;
        res.resume();
        resolve(ok);
      });
      req.on('error', () => resolve(false));
      req.on('timeout', () => { req.destroy(); resolve(false); });
      req.end();
    } catch { resolve(false); }
  });
}

async function validateAttractionPhotos(attractions) {
  const cacheFile = path.join('output', 'photo-cache.json');
  const cache = fs.existsSync(cacheFile) ? JSON.parse(fs.readFileSync(cacheFile, 'utf8')) : {};
  let updated = false;

  for (const a of attractions) {
    const photos = a.фото || [];
    const valid = [];
    for (const url of photos) {
      // Кешируем только успехи. Неудачи перепроверяем — Wikimedia любит
      // временно срезать соединения при пачке запросов.
      if (cache[url] === true) { valid.push(url); continue; }

      let ok = false;
      for (let attempt = 0; attempt < 2; attempt++) {
        ok = await isUrlAlive(url);
        if (ok) break;
        await sleep(800);
      }
      if (ok) cache[url] = true;
      updated = true;
      if (ok) {
        valid.push(url);
      } else {
        console.log(`  ⚠ битое фото у «${a.название}»: ${url.slice(0, 80)}…`);
      }
      await sleep(250); // не давим на сервер
    }
    a.фото = valid;
  }

  if (updated) fs.writeFileSync(cacheFile, JSON.stringify(cache, null, 2));
}

// Bounding box check — results outside city bounds are rejected
function inCityBounds(lat, lng) {
  return lat >= CITY_LAT - 0.22 && lat <= CITY_LAT + 0.22 &&
         lng >= CITY_LNG - 0.35 && lng <= CITY_LNG + 0.35;
}

function haversineMeters(lat1, lng1, lat2, lng2) {
  const R = 6371000;
  const toRad = d => d * Math.PI / 180;
  const dLat = toRad(lat2 - lat1);
  const dLng = toRad(lng2 - lng1);
  const a = Math.sin(dLat/2)**2 + Math.cos(toRad(lat1))*Math.cos(toRad(lat2))*Math.sin(dLng/2)**2;
  return 2 * R * Math.asin(Math.sqrt(a));
}

function findMetroStation(name) {
  if (!name) return null;
  const norm = name.toLowerCase().replace(/ё/g, 'е').trim();
  return METRO.find(s => {
    const sn = s.название.toLowerCase().replace(/ё/g, 'е');
    return sn.includes(norm) || norm.includes(sn);
  }) || null;
}

function findManualCoords(name) {
  if (!name) return null;
  const lower = name.toLowerCase();
  const overrides = config.ручные_координаты || {};
  for (const [key, coords] of Object.entries(overrides)) {
    if (lower.includes(key.toLowerCase())) return coords;
  }
  return null;
}

async function geocodeHotels(hotels, cityName, country) {
  const cacheFile = path.join('output', 'geocode-cache.json');
  const cache = fs.existsSync(cacheFile) ? JSON.parse(fs.readFileSync(cacheFile, 'utf8')) : {};
  let updated = false;

  for (const h of hotels) {
    if (h.lat && h.lng) continue;

    // 1. Manual override by hotel name (highest priority, авто Nominatim часто врёт)
    const manual = findManualCoords(h.name);
    if (manual) {
      h.lat = manual.lat; h.lng = manual.lng;
      console.log(`  📌 ${h.name.slice(0, 45).padEnd(45)} → ручные координаты`);
      continue;
    }

    const addr = parseAddress(h.address);
    // Extra safety strip in case parseAddress still leaves distance suffixes
    const rawStreet = (addr.street || h.address || '')
      .replace(/,?\s*[А-ЯЁа-яё]+\s*[\d,]+\s*(км|м)\s+от\s+(центра|метро).*/i, '')
      .replace(/^(Amazing|Great|Good|Outstanding|Wonderful|Very good|Superb|Excellent)$/i, '')
      .trim().replace(/,$/, '');

    // If no street, fall back to hotel name (works well for named landmarks)
    const searchTerm = rawStreet || h.name;
    if (!searchTerm) continue;

    const query = `${searchTerm}, ${cityName}`;
    const key   = query.toLowerCase();

    let coords = null;
    let fromCache = false;
    if (cache[key]) {
      coords = cache[key];
      fromCache = true;
    } else {
      process.stdout.write(`  🔍 ${query} … `);
      coords = await geocodeRequest(query);
      await sleep(1200); // Nominatim: max 1 req/sec
    }

    // 2. Validate: if we know claimed metro distance, geocode must roughly agree
    let valid = !!(coords && inCityBounds(coords.lat, coords.lng));
    if (valid && addr.metro_m && addr.metro_station) {
      const station = findMetroStation(addr.metro_station);
      if (station) {
        const actual = haversineMeters(coords.lat, coords.lng, station.lat, station.lng);
        // Allow 2.5x claimed distance + 400m slack (адрес может быть приблизителен)
        const tolerance = Math.max(addr.metro_m * 2.5, addr.metro_m + 400);
        if (actual > tolerance) {
          console.log(`  ⚠ ${h.name.slice(0, 40)}: геокод в ${(actual/1000).toFixed(1)} км от ${station.название}, заявлено ${addr.metro_m} м — отклонено`);
          valid = false;
        }
      }
    }

    if (valid) {
      h.lat = coords.lat; h.lng = coords.lng;
      if (!fromCache) {
        cache[key] = coords; updated = true;
        console.log(`✓ ${coords.lat.toFixed(4)}, ${coords.lng.toFixed(4)}`);
      }
      continue;
    }

    // 3. Fallback: place hotel on its metro station (грубо, но лучше чем мимо города)
    if (addr.metro_station) {
      const station = findMetroStation(addr.metro_station);
      if (station) {
        h.lat = station.lat; h.lng = station.lng;
        h._approx_from_metro = station.название;
        console.log(`  📍 ${h.name.slice(0, 40)} → fallback на станцию ${station.название}`);
        continue;
      }
    }

    if (!fromCache) {
      if (coords) console.log(`✗ вне города (${coords.lat.toFixed(3)}, ${coords.lng.toFixed(3)})`);
      else        console.log('✗ не найдено');
    }
  }

  if (updated) fs.writeFileSync(cacheFile, JSON.stringify(cache, null, 2));
}

// ── Load data ─────────────────────────────────────────────────────────
const config = JSON.parse(fs.readFileSync('hotel-config.json', 'utf8'));

const latestFile = path.join('output', 'latest.json');
if (!fs.existsSync(latestFile)) {
  console.error('❌ Нет данных. Сначала запустите: node hotel-parser.js');
  process.exit(1);
}

const allHotels = JSON.parse(fs.readFileSync(latestFile, 'utf8'));
console.log(`📊 Загружено: ${allHotels.length} отелей`);

const hotels = allHotels
  .filter(h => !/хостел|hostel|дорм|dorm/i.test(h.name || ''))
  .sort((a, b) => (a.price_per_night_rub || 999999) - (b.price_per_night_rub || 999999));

// ── Config ────────────────────────────────────────────────────────────
const CITY_LAT      = config.город.lat;
const CITY_LNG      = config.город.lng;
const CITY_NAME     = config.город.название;
const CITY_SLUG     = config.город.slug;
const ATTRACTIONS   = config.достопримечательности || [];
const PLACES        = config.интересные_места || [];
const METRO         = config.метро || [];

const PLACE_CATEGORIES = {
  'еда':      { цвет: '#FB8C00', фон: '#FFF3E0', label: 'Поесть/выпить' },
  'тусовка':  { цвет: '#8E24AA', фон: '#F3E5F5', label: 'Бары/тусовка' },
  'шопинг':   { цвет: '#1E88E5', фон: '#E3F2FD', label: 'Шопинг/ТЦ' },
  'природа':  { цвет: '#43A047', фон: '#E8F5E9', label: 'Парки/природа' },
  'загород':  { цвет: '#6D4C41', фон: '#EFEBE9', label: 'За городом' },
};
const checkin   = config.даты.заезд;
const checkout  = config.даты.выезд;
const nights    = hotels.find(h => h.nights)?.nights ?? 2;

// ── Address parser ────────────────────────────────────────────────────
function parseAddress(addr) {
  if (!addr) return { street: null, center_km: null, metro_m: null, metro_station: null };

  // Metro in meters: "581 м от метро Первомайская"
  const metroMMatch  = addr.match(/(\d[\d\s]*)\s*м\s+от\s+метро\s+([А-ЯЁа-яёA-Za-z][^,\d]+)/i);
  // Metro in km: "1,8 км от метро Борисовский тракт"
  const metroKmMatch = addr.match(/([\d,]+)\s*км\s+от\s+метро\s+([А-ЯЁа-яёA-Za-z][^,\d]+)/i);

  let metro_m = null, metro_station = null;
  if (metroMMatch) {
    metro_m       = parseInt(metroMMatch[1].replace(/\s/g, ''));
    metro_station = metroMMatch[2].trim();
  } else if (metroKmMatch) {
    metro_m       = Math.round(parseFloat(metroKmMatch[1].replace(',', '.')) * 1000);
    metro_station = metroKmMatch[2].trim();
  }

  // Distance from center in km: "1,7 км от центра"
  const centerKmMatch = addr.match(/([\d,]+)\s*км\s+от\s+центра/i);
  // Distance from center in m: "936 м от центра"
  const centerMMatch  = addr.match(/(\d[\d\s]*)\s*м\s+от\s+центра/i);
  let center_km = null;
  if (centerKmMatch) {
    center_km = parseFloat(centerKmMatch[1].replace(',', '.'));
  } else if (centerMMatch) {
    center_km = Math.round(parseInt(centerMMatch[1].replace(/\s/g, '')) / 100) / 10;
  }

  // Street: strip distance suffixes. Handles both "Минск936 м" and "Минск1,7 км" concatenated forms.
  const street = addr
    .replace(/,?\s*[А-ЯЁа-яё]+\s*[\d,]+\s*(км|м)\s+от\s+(центра|метро).*/i, '')
    .replace(/,?\s*\b[А-ЯЁ][а-яё]+\s*$/, '')
    .trim()
    .replace(/,$/, '');

  return { street: street || null, center_km, metro_m, metro_station };
}

function metroStr(metro_m) {
  if (!metro_m) return null;
  const min = Math.round(metro_m / 80);
  return `${metro_m} м (${min} мин)`;
}

function centerStr(center_km) {
  if (center_km == null) return null;
  return center_km < 1.5 ? `${center_km} км ✓` : `${center_km} км`;
}

// ── Price / category helpers ──────────────────────────────────────────
function getPriceCategory(ppn) {
  if (!ppn) return 'Нет цены';
  if (ppn <  3000) return 'Бюджет (< 3 000 ₽/н)';
  if (ppn <  7000) return 'Средний (3–7 тыс ₽/н)';
  if (ppn < 15000) return 'Выше среднего (7–15 тыс ₽/н)';
  return 'Премиум (15 000+ ₽/н)';
}

const CAT_ORDER = [
  'Бюджет (< 3 000 ₽/н)',
  'Средний (3–7 тыс ₽/н)',
  'Выше среднего (7–15 тыс ₽/н)',
  'Премиум (15 000+ ₽/н)',
  'Нет цены',
];

const CAT_COLORS = {
  'Бюджет (< 3 000 ₽/н)':         { bg: 'C6EFCE', fg: '276221' },
  'Средний (3–7 тыс ₽/н)':        { bg: 'FFEB9C', fg: '9C5700' },
  'Выше среднего (7–15 тыс ₽/н)': { bg: 'FCE4D6', fg: '843C0C' },
  'Премиум (15 000+ ₽/н)':        { bg: 'F4CCCC', fg: '660000' },
  'Нет цены':                       { bg: 'F2F2F2', fg: '888888' },
};

const SOURCE_BG = { booking: 'DDEEFF', trip: 'DDFFEE', ostrovok: 'FFF0DD' };

function srcLabel(s) {
  if (s === 'booking')  return 'Booking';
  if (s === 'trip')     return 'Trip.com';
  if (s === 'ostrovok') return 'Ostrovok';
  return s;
}

// ── Excel style helpers ───────────────────────────────────────────────
function styleHeader(row, bgHex = '1F3864', fgHex = 'FFFFFF') {
  row.height = 22;
  row.eachCell(cell => {
    cell.fill      = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + bgHex } };
    cell.font      = { bold: true, color: { argb: 'FF' + fgHex }, size: 10 };
    cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: false };
    cell.border    = { bottom: { style: 'thin', color: { argb: 'FFAAAAAA' } } };
  });
}

function styleRow(row, bgHex) {
  row.eachCell(cell => {
    cell.fill      = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + bgHex } };
    cell.alignment = { vertical: 'middle', wrapText: false };
    cell.border    = { bottom: { style: 'hair', color: { argb: 'FFDDDDDD' } } };
  });
}

function linkCell(row, colNum, name, url) {
  if (url) {
    row.getCell(colNum).value = { text: name, hyperlink: url };
    row.getCell(colNum).font  = { color: { argb: 'FF0563C1' }, underline: true };
  } else {
    row.getCell(colNum).value = name;
  }
}

function priceFmt(row, ...cols) {
  for (const c of cols) row.getCell(c).numFmt = '#,##0 "₽"';
}

// ── Sheet: Итог ───────────────────────────────────────────────────────
function buildSummarySheet(wb) {
  const ws = wb.addWorksheet('Итог', { properties: { tabColor: { argb: 'FFFF9800' } } });
  ws.columns = [
    { width: 36 }, { width: 7 }, { width: 13 }, { width: 13 }, { width: 13 },
    { width: 11 }, { width: 18 }, { width: 14 }, { width: 22 },
  ];

  const dateRange = `${new Date(checkin).toLocaleDateString('ru-RU')} – ${new Date(checkout).toLocaleDateString('ru-RU')}`;

  ws.mergeCells('A1:I1');
  const title = ws.getCell('A1');
  title.value = `Отели в ${CITY_NAME}  |  ${nights} ночи  |  ${dateRange}`;
  title.font  = { bold: true, size: 13 };
  title.alignment = { horizontal: 'center', vertical: 'middle' };
  ws.getRow(1).height = 28;

  ws.mergeCells('A2:I2');
  ws.getCell('A2').value = [
    `Всего: ${allHotels.length}`,
    `Ostrovok: ${allHotels.filter(h => h.source === 'ostrovok').length}`,
    `Trip.com: ${allHotels.filter(h => h.source === 'trip').length}`,
    `Без хостелов: ${hotels.length}`,
    `С ценой: ${hotels.filter(h => h.price_per_night_rub).length}`,
  ].join('   |   ');
  ws.getCell('A2').font = { italic: true, color: { argb: 'FF555555' }, size: 10 };
  ws.getRow(2).height = 16;
  ws.addRow([]);

  for (const cat of CAT_ORDER) {
    const catHotels = hotels.filter(h => getPriceCategory(h.price_per_night_rub) === cat);
    if (catHotels.length === 0) continue;

    const { bg, fg } = CAT_COLORS[cat] ?? { bg: 'EEEEEE', fg: '000000' };
    const catRow = ws.addRow([`${cat}  (${catHotels.length} вариантов)`]);
    ws.mergeCells(`A${catRow.number}:I${catRow.number}`);
    catRow.getCell(1).fill      = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + bg } };
    catRow.getCell(1).font      = { bold: true, size: 11, color: { argb: 'FF' + fg } };
    catRow.getCell(1).alignment = { horizontal: 'left', vertical: 'middle', indent: 1 };
    catRow.height = 24;

    const hdr = ws.addRow(['Отель', '★', 'Рейтинг', 'Цена/ночь', 'Итого', 'Источник', 'До метро', 'До центра', 'Улица']);
    styleHeader(hdr, '444444', 'FFFFFF');

    for (const h of catHotels.slice(0, 15)) {
      const a   = parseAddress(h.address);
      const row = ws.addRow([
        h.name,
        h.stars ? '★'.repeat(Math.min(h.stars, 5)) : '?',
        h.rating ? `${h.rating}${h.rating_label ? '  ' + h.rating_label : ''}` : '—',
        h.price_per_night_rub ?? null,
        h.price_total_rub ?? null,
        srcLabel(h.source),
        metroStr(a.metro_m) ?? h.distance_text ?? '—',
        centerStr(a.center_km) ?? '—',
        a.street ?? h.address ?? '—',
      ]);
      styleRow(row, SOURCE_BG[h.source] ?? 'FFFFFF');
      linkCell(row, 1, h.name, h.url);
      priceFmt(row, 4, 5);
      row.height = 18;
    }

    if (catHotels.length > 15) {
      const more = ws.addRow([`    … ещё ${catHotels.length - 15} вариантов — смотри лист "Все отели"`]);
      ws.mergeCells(`A${more.number}:I${more.number}`);
      more.getCell(1).font = { italic: true, color: { argb: 'FF999999' }, size: 9 };
    }
    ws.addRow([]);
  }
}

// ── Sheet: Все отели ──────────────────────────────────────────────────
function buildAllSheet(wb) {
  const ws = wb.addWorksheet('Все отели', { properties: { tabColor: { argb: 'FF4472C4' } } });
  ws.columns = [
    { header: 'Отель',         width: 36 },
    { header: '★',             width: 7  },
    { header: 'Рейтинг',      width: 9  },
    { header: 'Отзывов',      width: 9  },
    { header: 'Цена/ночь',    width: 12 },
    { header: 'Итого',         width: 12 },
    { header: 'Источник',     width: 11 },
    { header: 'Категория',    width: 22 },
    { header: 'До метро',     width: 16 },
    { header: 'Станция',      width: 22 },
    { header: 'До центра',    width: 11 },
    { header: 'В центре?',    width: 10 },
    { header: 'Улица',         width: 32 },
  ];
  styleHeader(ws.getRow(1));

  for (const h of hotels) {
    const a   = parseAddress(h.address);
    const row = ws.addRow([
      h.name,
      h.stars ?? '?',
      h.rating ?? null,
      h.review_count ?? null,
      h.price_per_night_rub ?? null,
      h.price_total_rub ?? null,
      srcLabel(h.source),
      getPriceCategory(h.price_per_night_rub),
      metroStr(a.metro_m) ?? h.distance_text ?? '—',
      a.metro_station ?? '—',
      a.center_km != null ? a.center_km : '—',
      a.center_km != null ? (a.center_km < 1.5 ? '✓ да' : 'нет') : '—',
      a.street ?? h.address ?? '—',
    ]);
    styleRow(row, SOURCE_BG[h.source] ?? 'FFFFFF');
    linkCell(row, 1, h.name, h.url);
    priceFmt(row, 5, 6);
    row.height = 17;
  }
  ws.autoFilter = { from: 'A1', to: 'M1' };
}

// ── Sheet: По рейтингу ────────────────────────────────────────────────
function buildRatingSheet(wb) {
  const ws = wb.addWorksheet('По рейтингу', { properties: { tabColor: { argb: 'FF70AD47' } } });
  ws.columns = [
    { header: 'Отель',      width: 36 },
    { header: 'Рейтинг',   width: 9  },
    { header: 'Отзывов',   width: 9  },
    { header: '★',          width: 7  },
    { header: 'Цена/ночь', width: 12 },
    { header: 'Итого',      width: 12 },
    { header: 'Источник',  width: 11 },
    { header: 'До метро',  width: 16 },
    { header: 'До центра', width: 11 },
    { header: 'Улица',      width: 32 },
  ];
  styleHeader(ws.getRow(1), '2E7D32', 'FFFFFF');

  const sorted = [...hotels].sort((a, b) => (b.rating ?? 0) - (a.rating ?? 0));
  for (const h of sorted) {
    const a   = parseAddress(h.address);
    const row = ws.addRow([
      h.name,
      h.rating ?? null,
      h.review_count ?? null,
      h.stars ?? '?',
      h.price_per_night_rub ?? null,
      h.price_total_rub ?? null,
      srcLabel(h.source),
      metroStr(a.metro_m) ?? h.distance_text ?? '—',
      a.center_km != null ? a.center_km : '—',
      a.street ?? h.address ?? '—',
    ]);
    styleRow(row, SOURCE_BG[h.source] ?? 'FFFFFF');
    linkCell(row, 1, h.name, h.url);
    priceFmt(row, 5, 6);
    row.height = 17;
  }
  ws.autoFilter = { from: 'A1', to: 'J1' };
}

// ── Sheet: Рекомендации ───────────────────────────────────────────────
function buildRecommendationsSheet(wb) {
  const ws = wb.addWorksheet('🏆 Топ', { properties: { tabColor: { argb: 'FF7030A0' } } });
  ws.columns = [
    { width: 4 }, { width: 34 }, { width: 13 }, { width: 9 }, { width: 16 }, { width: 11 }, { width: 28 },
  ];

  const withPrice   = hotels.filter(h => h.price_per_night_rub);
  const withRating  = hotels.filter(h => h.rating && h.price_per_night_rub);

  // Value score: rating per 1000 rub
  const byValue    = [...withRating].sort((a, b) =>
    (b.rating / b.price_per_night_rub) - (a.rating / a.price_per_night_rub)
  );

  // Location score: closest center + metro
  const byLocation = [...hotels].sort((a, b) => {
    const aAddr = parseAddress(a.address);
    const bAddr = parseAddress(b.address);
    const aScore = (aAddr.center_km ?? 99) + (aAddr.metro_m ?? 9999) / 1000;
    const bScore = (bAddr.center_km ?? 99) + (bAddr.metro_m ?? 9999) / 1000;
    return aScore - bScore;
  });

  const cheapWithRating = [...withRating]
    .filter(h => h.rating >= 8)
    .sort((a, b) => a.price_per_night_rub - b.price_per_night_rub);

  const topRating = [...withRating].sort((a, b) => b.rating - a.rating);

  const sections = [
    { emoji: '💎', title: 'Лучшая цена/качество', desc: 'Максимальный рейтинг за минимальные деньги', picks: byValue.slice(0, 4) },
    { emoji: '📍', title: 'Лучшее расположение', desc: 'Ближайшие к центру с удобным метро', picks: byLocation.slice(0, 4) },
    { emoji: '💰', title: 'Бюджетный выбор', desc: 'Дешевле всего при рейтинге ≥ 8.0', picks: cheapWithRating.slice(0, 3) },
    { emoji: '⭐', title: 'Наивысший рейтинг', desc: 'Топ по оценкам гостей', picks: topRating.slice(0, 4) },
  ];

  const dateRange = `${new Date(checkin).toLocaleDateString('ru-RU')} – ${new Date(checkout).toLocaleDateString('ru-RU')}`;
  ws.mergeCells('A1:G1');
  ws.getCell('A1').value = `Рекомендации — ${CITY_NAME}, ${dateRange}`;
  ws.getCell('A1').font  = { bold: true, size: 14 };
  ws.getCell('A1').alignment = { horizontal: 'center', vertical: 'middle' };
  ws.getRow(1).height = 30;
  ws.addRow([]);

  for (const sec of sections) {
    if (sec.picks.length === 0) continue;

    // Section header
    ws.mergeCells(`A${ws.rowCount + 1}:G${ws.rowCount + 1}`);
    const secRow = ws.addRow([`${sec.emoji}  ${sec.title}`]);
    ws.mergeCells(`A${secRow.number}:G${secRow.number}`);
    secRow.getCell(1).font      = { bold: true, size: 12, color: { argb: 'FF333333' } };
    secRow.getCell(1).fill      = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF3E5FF' } };
    secRow.getCell(1).alignment = { indent: 1, vertical: 'middle' };
    secRow.height = 26;

    const descRow = ws.addRow([null, sec.desc]);
    ws.mergeCells(`B${descRow.number}:G${descRow.number}`);
    descRow.getCell(2).font = { italic: true, color: { argb: 'FF888888' }, size: 9 };

    const hdr = ws.addRow(['#', 'Отель', 'Цена/ночь', 'Рейтинг', 'До метро', 'До центра', 'Улица']);
    styleHeader(hdr, '5B2C8D', 'FFFFFF');

    sec.picks.forEach((h, idx) => {
      const a   = parseAddress(h.address);
      const row = ws.addRow([
        idx + 1,
        h.name,
        h.price_per_night_rub ?? null,
        h.rating ?? null,
        metroStr(a.metro_m) ?? '—',
        a.center_km != null ? a.center_km + ' км' : '—',
        a.street ?? '—',
      ]);
      styleRow(row, SOURCE_BG[h.source] ?? 'FFFFFF');
      linkCell(row, 2, h.name, h.url);
      row.getCell(3).numFmt = '#,##0 "₽"';
      row.getCell(1).alignment = { horizontal: 'center', vertical: 'middle' };
      row.getCell(1).font = { bold: true, color: { argb: 'FF5B2C8D' } };
      row.height = 18;
    });

    ws.addRow([]);
  }
}

// ── Sheet: Источники ──────────────────────────────────────────────────
function buildSourcesSheet(wb) {
  const ws = wb.addWorksheet('Источники', { properties: { tabColor: { argb: 'FFED7D31' } } });

  function srcStats(list) {
    const wp = list.filter(h => h.price_per_night_rub);
    const pp = wp.map(h => h.price_per_night_rub);
    return [
      list.length,
      wp.length,
      pp.length ? Math.min(...pp) : '—',
      pp.length ? Math.max(...pp) : '—',
      pp.length ? Math.round(pp.reduce((a, b) => a + b, 0) / pp.length) : '—',
      list.filter(h => h.rating).length
        ? +(list.reduce((s, h) => s + (h.rating || 0), 0) / list.filter(h => h.rating).length).toFixed(1)
        : '—',
    ];
  }

  const bk = srcStats(hotels.filter(h => h.source === 'booking'));
  const tr = srcStats(hotels.filter(h => h.source === 'trip'));
  const os = srcStats(hotels.filter(h => h.source === 'ostrovok'));

  const stats = [
    ['',                  'Booking.com', 'Trip.com', 'Ostrovok.ru'],
    ['Всего отелей',      bk[0], tr[0], os[0]],
    ['С ценой',           bk[1], tr[1], os[1]],
    ['Мин цена/ночь',     bk[2], tr[2], os[2]],
    ['Макс цена/ночь',    bk[3], tr[3], os[3]],
    ['Средняя цена/ночь', bk[4], tr[4], os[4]],
    ['Средний рейтинг',   bk[5], tr[5], os[5]],
  ];

  ws.columns = [{ width: 22 }, { width: 16 }, { width: 16 }, { width: 16 }];
  ws.mergeCells('A1:D1');
  ws.getCell('A1').value = 'Сравнение источников';
  ws.getCell('A1').font  = { bold: true, size: 13 };
  ws.getCell('A1').alignment = { horizontal: 'center' };
  ws.getRow(1).height = 26;

  for (const [i, rowData] of stats.entries()) {
    const row = ws.addRow(rowData);
    if (i === 0) { styleHeader(row); continue; }
    row.height = 18;
    for (let c = 2; c <= 4; c++) {
      if (typeof row.getCell(c).value === 'number' && i >= 3 && i <= 5)
        row.getCell(c).numFmt = '#,##0 "₽"';
    }
  }
}

// ── HTML Map ──────────────────────────────────────────────────────────
function buildHtmlMap() {
  function pinColor(ppn) {
    if (!ppn) return '#9E9E9E';
    if (ppn <  3000) return '#2E7D32';  // тёмно-зелёный
    if (ppn <  7000) return '#F9A825';  // жёлтый
    if (ppn < 15000) return '#E65100';  // оранжевый
    return '#B71C1C';                    // красный
  }

  function pinLabel(ppn) {
    if (!ppn) return '?';
    if (ppn >= 10000) return Math.round(ppn / 1000) + 'к';
    if (ppn >=  1000) return (ppn / 1000).toFixed(1).replace('.0', '') + 'к';
    return ppn + '₽';
  }

  const safe = s => (s || '').replace(/`/g, "'").replace(/\\/g, '\\\\').replace(/\n/g, ' ');

  function priorityBadge(p) {
    if (!p) return '';
    const stars = '★'.repeat(Math.min(Math.max(p, 1), 3));
    const color = p === 3 ? '#D32F2F' : p === 2 ? '#F57C00' : '#777';
    return ` <span style="color:${color};font-size:11px;font-weight:700;letter-spacing:1px">${stars}</span>`;
  }

  const markers = allHotels.map((h, i) => {
    const lat  = h.lat || (CITY_LAT + (((i * 7919) % 200) - 100) * 0.0002);
    const lng  = h.lng || (CITY_LNG + (((i * 6271) % 200) - 100) * 0.0003);
    const ppn  = h.price_per_night_rub;
    const col  = pinColor(ppn);
    const lbl  = pinLabel(ppn);
    const a    = parseAddress(h.address);

    const metro_min   = a.metro_m ? Math.round(a.metro_m / 80) : null;
    const metroSub    = metro_min ? `<div class="hpin-sub">🚇 ${metro_min} мин</div>` : '';

    let approx = '';
    if (h._approx_from_metro) {
      approx = `<br><small style="color:#E65100">⚠ координаты приблизительные (метка ставится на станцию ${safe(h._approx_from_metro)})</small>`;
    } else if (!h.lat) {
      approx = '<br><small style="color:#aaa">⚠ координаты приблизительные</small>';
    }
    const priceStr  = ppn ? ppn.toLocaleString('ru-RU') + ' ₽/н' : h.price_display || 'нет цены';
    const totalStr  = h.price_total_rub ? `<br>Итого: <b>${h.price_total_rub.toLocaleString('ru-RU')} ₽</b>` : '';
    const ratingStr = h.rating ? `<br>★ ${h.rating}${h.review_count ? ` (${h.review_count} отз.)` : ''}` : '';
    const metroInfo = a.metro_m ? `<br>🚇 ${a.metro_station} — ${metroStr(a.metro_m)}` : '';
    const ctrInfo   = a.center_km != null ? `<br>📍 ${a.center_km} км от центра` : '';
    const streetInfo = a.street ? `<br><span style="color:#666">${safe(a.street)}</span>` : '';
    const linkHtml  = h.url ? `<br><a href="${h.url}" target="_blank" style="color:#1565C0;font-weight:bold">Открыть →</a>` : '';
    const srcBadge  = h.source === 'booking' ? '📘' : h.source === 'ostrovok' ? '🏝' : '✈️';

    return `L.marker([${lat.toFixed(6)}, ${lng.toFixed(6)}], {
  icon: L.divIcon({
    className: '',
    html: '<div class="hpin" style="background:${col}">${lbl}${metroSub}</div>',
    iconSize: [null, null],
    iconAnchor: [22, ${metro_min ? 20 : 14}],
    popupAnchor: [0, -22]
  })
}).bindPopup(\`<div class="hpop">
<b style="font-size:14px">${safe(h.name)}</b>
<div style="color:${col};font-size:15px;font-weight:bold;margin:4px 0">${priceStr}</div>
${totalStr}${ratingStr}
${metroInfo}${ctrInfo}${streetInfo}
<div style="margin-top:6px;color:#888;font-size:11px">${srcBadge} ${srcLabel(h.source)}</div>
${approx}${linkHtml}
</div>\`).addTo(map);`;
  }).join('\n\n');

  const attractionMarkers = ATTRACTIONS.map(a => {
    const safeName = safe(a.название);
    const safeDesc = safe(a.описание || '');
    const photos   = (a.фото || []);
    const slideId  = 'sl' + Math.abs(safeName.split('').reduce((h, c) => (h << 5) - h + c.charCodeAt(0), 0));

    let photoHtml = '';
    if (photos.length > 0) {
      // onerror: вместо display:none — заменяем на серый плейсхолдер,
      // чтобы высота слайда сохранялась и попап не «прыгал».
      const placeholder = "data:image/gif;base64,R0lGODlhAQABAIAAAOXl5QAAACH5BAAAAAAALAAAAAABAAEAAAICRAEAOw==";
      const imgs = photos.map((src, i) =>
        `<img src="${src}" class="aslide${i===0?' active':''}" data-idx="${i}" loading="eager" onerror="this.onerror=null;this.src='${placeholder}'">`
      ).join('');
      const arrows = photos.length > 1 ? `
<button class="anav aprev" onclick="apMove('${slideId}',-1)" aria-label="prev">‹</button>
<button class="anav anext" onclick="apMove('${slideId}',1)" aria-label="next">›</button>
<div class="acounter"><span class="acur">1</span>/${photos.length}</div>` : '';
      const dots = photos.length > 1 ? `<div class="adots">${
        photos.map((_, i) => `<span class="adot${i===0?' active':''}" onclick="apGoto('${slideId}',${i})"></span>`).join('')
      }</div>` : '';
      photoHtml = `<div class="aslider" id="${slideId}" data-total="${photos.length}">${imgs}${arrows}</div>${dots}`;
    }

    const badge   = priorityBadge(a.приоритет);
    const popupContent = `<div class="hpop apop">
<b style="font-size:14px">${safeName}</b>${badge}
${photoHtml}${safeDesc ? `<div style="color:#444;font-size:12px;line-height:1.5;margin-top:4px">${safeDesc}</div>` : ''}
</div>`;

    return `markerByName[${JSON.stringify(a.название)}] = L.marker([${a.lat}, ${a.lng}], {
  icon: L.divIcon({
    className: '',
    html: '<div class="apin">${a.тип}</div>',
    iconSize: [30, 30], iconAnchor: [15, 15], popupAnchor: [0, -18]
  })
}).bindPopup(\`${popupContent}\`, {maxWidth: 300, minWidth: 260}).addTo(map);`;
  }).join('\n\n');

  // ── «Интересные места»: кафе/бары/шопинг/природа/загород ──────────────
  const placeholder = "data:image/gif;base64,R0lGODlhAQABAIAAAOXl5QAAACH5BAAAAAAALAAAAAABAAEAAAICRAEAOw==";
  const placesMarkers = PLACES.map(p => {
    const cat = PLACE_CATEGORIES[p.категория] || { цвет: '#666', фон: '#eee', label: p.категория };
    const safeName = safe(p.название);
    const safeDesc = safe(p.описание || '');
    const safeAddr = safe(p.адрес || '');
    const photos   = p.фото || [];
    const trendBadge = p.тренд
      ? '<span style="background:#FFEB3B;color:#5D4037;font-size:10px;font-weight:700;padding:2px 6px;border-radius:8px;margin-left:6px">🔥 тренд</span>'
      : '';

    // Слайдер фото (та же логика что у достопримечательностей)
    const slideId = 'pl' + Math.abs(safeName.split('').reduce((h, c) => (h << 5) - h + c.charCodeAt(0), 0));
    let photoHtml = '';
    if (photos.length > 0) {
      const imgs = photos.map((src, i) =>
        `<img src="${src}" class="aslide${i===0?' active':''}" data-idx="${i}" loading="eager" onerror="this.onerror=null;this.src='${placeholder}'">`
      ).join('');
      const arrows = photos.length > 1 ? `
<button class="anav aprev" onclick="apMove('${slideId}',-1)" aria-label="prev">‹</button>
<button class="anav anext" onclick="apMove('${slideId}',1)" aria-label="next">›</button>
<div class="acounter"><span class="acur">1</span>/${photos.length}</div>` : '';
      const dots = photos.length > 1 ? `<div class="adots">${
        photos.map((_, i) => `<span class="adot${i===0?' active':''}" onclick="apGoto('${slideId}',${i})"></span>`).join('')
      }</div>` : '';
      photoHtml = `<div class="aslider" id="${slideId}" data-total="${photos.length}">${imgs}${arrows}</div>${dots}`;
    }

    const popup = `<div class="hpop ppop apop">
<b style="font-size:14px">${p.тип} ${safeName}</b>${trendBadge}
${photoHtml}${safeDesc ? `<div style="color:#444;font-size:12px;line-height:1.5;margin-top:6px">${safeDesc}</div>` : ''}
${safeAddr ? `<div style="color:#777;font-size:11px;margin-top:4px">📍 ${safeAddr}</div>` : ''}
<div style="margin-top:6px"><span style="background:${cat.фон};color:${cat.цвет};font-size:10px;font-weight:600;padding:2px 8px;border-radius:8px">${cat.label}</span></div>
</div>`;
    return `markerByName[${JSON.stringify(p.название)}] = L.marker([${p.lat}, ${p.lng}], {
  icon: L.divIcon({
    className: '',
    html: '<div class="ppin" style="background:${cat.фон};border-color:${cat.цвет}">${p.тип}</div>',
    iconSize: [26, 26], iconAnchor: [13, 13], popupAnchor: [0, -16]
  })
}).bindPopup(\`${popup}\`, {maxWidth: 300, minWidth: 260}).addTo(placesLayer);`;
  }).join('\n\n');

  const bCount = allHotels.filter(h => h.source === 'booking').length;
  const tCount = allHotels.filter(h => h.source === 'trip').length;
  const oCount = allHotels.filter(h => h.source === 'ostrovok').length;
  const dateRange = `${new Date(checkin).toLocaleDateString('ru-RU')} – ${new Date(checkout).toLocaleDateString('ru-RU')}`;

  const html = `<!DOCTYPE html>
<html lang="ru">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Отели ${CITY_NAME} — карта</title>
<link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css"/>
<script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"></script>
<style>
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif; }
  #map { height: 100vh; }

  .hpin {
    color: #fff;
    font-size: 11px;
    font-weight: 700;
    padding: 4px 9px;
    border-radius: 14px;
    border: 2px solid rgba(255,255,255,0.9);
    box-shadow: 0 2px 8px rgba(0,0,0,.35);
    white-space: nowrap;
    cursor: pointer;
    text-align: center;
    line-height: 1.3;
    transition: transform .15s;
  }
  .hpin:hover { transform: scale(1.15); }
  .hpin-sub { font-size: 9px; opacity: 0.88; margin-top: 1px; }

  .apin {
    font-size: 17px;
    background: white;
    border-radius: 50%;
    width: 30px;
    height: 30px;
    display: flex;
    align-items: center;
    justify-content: center;
    box-shadow: 0 2px 10px rgba(0,0,0,.40);
    cursor: pointer;
    border: 2px solid rgba(0,0,0,0.12);
    transition: transform .15s;
  }
  .apin:hover { transform: scale(1.2); }

  .ppin {
    font-size: 14px;
    border-radius: 50%;
    width: 26px;
    height: 26px;
    display: flex;
    align-items: center;
    justify-content: center;
    box-shadow: 0 2px 6px rgba(0,0,0,.30);
    cursor: pointer;
    border: 2px solid;
    transition: transform .15s;
  }
  .ppin:hover { transform: scale(1.25); z-index: 1000; }
  .ppop { font-family: -apple-system, sans-serif; font-size: 13px; min-width: 240px; line-height: 1.5; }

  .hpop { font-family: -apple-system, sans-serif; font-size: 13px; min-width: 220px; line-height: 1.5; }
  .apop { min-width: 260px; }

  .aslider { position: relative; margin: 8px 0 4px; border-radius: 8px; overflow: hidden; background: #f0f0f0; }
  .aslide { display: none; width: 100%; height: 160px; object-fit: cover; }
  .aslide.active { display: block; }
  .anav {
    position: absolute; top: 50%; transform: translateY(-50%);
    background: rgba(0,0,0,0.42); color: #fff;
    border: none; cursor: pointer;
    width: 30px; height: 38px; font-size: 22px; line-height: 1;
    display: flex; align-items: center; justify-content: center;
    transition: background .15s;
  }
  .anav:hover { background: rgba(0,0,0,0.65); }
  .aprev { left: 0; border-radius: 0 4px 4px 0; }
  .anext { right: 0; border-radius: 4px 0 0 4px; }
  .acounter {
    position: absolute; bottom: 6px; right: 8px;
    background: rgba(0,0,0,0.55); color: #fff;
    font-size: 11px; padding: 2px 7px; border-radius: 10px;
    pointer-events: none;
  }
  .adots { text-align: center; margin: 4px 0 6px; }
  .adot {
    display: inline-block; width: 7px; height: 7px;
    border-radius: 50%; background: #ccc; margin: 0 3px;
    cursor: pointer; transition: background .15s;
  }
  .adot.active { background: #555; }

  #legend {
    position: absolute; top: 12px; right: 12px; z-index: 1000;
    background: white; padding: 14px 18px; border-radius: 12px;
    box-shadow: 0 4px 16px rgba(0,0,0,.15); font-size: 13px; min-width: 190px;
  }
  #legend h3 { margin-bottom: 10px; font-size: 15px; font-weight: 700; }
  .leg-row { display: flex; align-items: center; gap: 10px; margin: 6px 0; }
  .leg-pin {
    min-width: 38px; text-align: center;
    color: #fff; font-size: 10px; font-weight: 700;
    padding: 3px 7px; border-radius: 11px;
    border: 2px solid rgba(255,255,255,0.8);
    box-shadow: 0 1px 4px rgba(0,0,0,.3);
  }
  .leg-label { color: #444; }
  hr { margin: 10px 0; border: none; border-top: 1px solid #eee; }
  .leg-src { color: #666; font-size: 12px; margin: 3px 0; }

  .mpin {
    background: white;
    border-radius: 50%;
    width: 20px; height: 20px;
    display: flex; align-items: center; justify-content: center;
    font-size: 9px; font-weight: 900; letter-spacing: -0.5px;
    box-shadow: 0 1px 5px rgba(0,0,0,.35);
    cursor: pointer; border: 2.5px solid;
    transition: transform .15s;
  }
  .mpin:hover { transform: scale(1.3); }

  #toggle-metro, #toggle-places {
    display: block; width: 100%; margin-top: 6px;
    background: #f5f5f5; border: 1px solid #ddd; border-radius: 7px;
    padding: 5px 8px; font-size: 12px; cursor: pointer; text-align: left;
  }
  #toggle-metro:hover, #toggle-places:hover { background: #eee; }
  .leg-cat-head { font-weight: 600; font-size: 11px; color: #555; margin-top: 8px; margin-bottom: 3px; text-transform: uppercase; letter-spacing: 0.5px; }
  .leg-place { color: #555; font-size: 11px; margin: 2px 0 2px 6px; line-height: 1.35; }
  .leg-place .trend { font-size: 9px; color: #B71C1C; margin-left: 3px; }

  .clickable { cursor: pointer; padding: 1px 5px; border-radius: 4px; margin-left: -5px; transition: background .12s; }
  .clickable:hover { background: #eaeaea; color: #000; }
  .clickable:active { background: #d8d8d8; }
</style>
</head>
<body>
<div id="map"></div>
<div id="legend">
  <h3>🏨 Отели ${CITY_NAME}</h3>
  <div style="color:#888;font-size:11px;margin-bottom:8px">${dateRange} · ${allHotels.length} объектов</div>
  <div class="leg-row"><div class="leg-pin" style="background:#2E7D32">2к</div><span class="leg-label">до 3 000 ₽/н</span></div>
  <div class="leg-row"><div class="leg-pin" style="background:#F9A825">5к</div><span class="leg-label">3 000 – 7 000 ₽/н</span></div>
  <div class="leg-row"><div class="leg-pin" style="background:#E65100">10к</div><span class="leg-label">7 000 – 15 000 ₽/н</span></div>
  <div class="leg-row"><div class="leg-pin" style="background:#B71C1C">20к</div><span class="leg-label">от 15 000 ₽/н</span></div>
  <div class="leg-row"><div class="leg-pin" style="background:#9E9E9E">?</div><span class="leg-label">цена неизвестна</span></div>
  <hr>
  <div class="leg-src">📘 Booking.com: ${bCount}</div>
  <div class="leg-src">✈️ Trip.com: ${tCount}</div>
  <div class="leg-src">🏝 Ostrovok.ru: ${oCount}</div>
  <hr>
  <div style="font-weight:600;font-size:12px;margin-bottom:6px">Достопримечательности</div>
  ${[...ATTRACTIONS].sort((x, y) => (y.приоритет || 0) - (x.приоритет || 0)).map(a => `<div class="leg-src clickable" data-name="${safe(a.название)}" onclick="focusItem(this.dataset.name)">${a.тип} ${a.название}${priorityBadge(a.приоритет)}</div>`).join('\n  ')}
  <div style="color:#aaa;font-size:10px;margin-top:4px">★★★ топ · ★★ интересно · ★ если останется время</div>
  ${PLACES.length ? `<hr>
  <div style="font-weight:600;font-size:12px">Интересные места <span style="color:#aaa;font-weight:400">${PLACES.length} шт.</span></div>
  ${Object.entries(PLACE_CATEGORIES).map(([key, cat]) => {
    const list = PLACES.filter(p => p.категория === key);
    if (!list.length) return '';
    const head = `<div class="leg-cat-head" style="color:${cat.цвет}"><span style="display:inline-block;width:8px;height:8px;border-radius:50%;background:${cat.цвет};margin-right:5px;vertical-align:middle"></span>${cat.label} (${list.length})</div>`;
    const items = list.map(p => `<div class="leg-place clickable" data-name="${safe(p.название)}" onclick="focusItem(this.dataset.name)">${p.тип} ${p.название}${p.тренд ? '<span class="trend">🔥</span>' : ''}</div>`).join('\n  ');
    return head + '\n  ' + items;
  }).filter(Boolean).join('\n  ')}` : ''}
  <hr>
  <button id="toggle-places" onclick="togglePlaces()">${PLACES.length ? '🍔 Интересные места: вкл' : ''}</button>
  <button id="toggle-metro" onclick="toggleMetro()">🚇 Метро: вкл</button>
  <div style="color:#aaa;font-size:11px;margin-top:8px">Клик по метке — подробности</div>
</div>
<script>
const map = L.map('map').setView([${CITY_LAT}, ${CITY_LNG}], 13);
L.tileLayer('https://{s}.basemaps.cartocdn.com/rastertiles/voyager/{z}/{x}/{y}{r}.png', {
  attribution: '© <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a> © <a href="https://carto.com/">CARTO</a>',
  maxZoom: 19, subdomains: 'abcd'
}).addTo(map);

${markers}

const markerByName = {};

${attractionMarkers}

const placesLayer = L.layerGroup().addTo(map);
${placesMarkers}

const metroLayer = L.layerGroup().addTo(map);

// Polyline линий метро (как на Яндекс.Картах: полупрозрачные цветные линии под станциями)
${(() => {
    const lines = {};
    for (const s of METRO) (lines[s.линия] = lines[s.линия] || []).push(s);
    const colors = { 1: '#CC0000', 2: '#0055AA', 3: '#168E44' };
    return Object.entries(lines).map(([n, stations]) => {
      const color = colors[n] || '#888';
      const coords = stations.map(s => `[${s.lat},${s.lng}]`).join(',');
      return `L.polyline([${coords}], { color: '${color}', weight: 5, opacity: 0.35, lineCap: 'round', lineJoin: 'round' }).addTo(metroLayer);`;
    }).join('\n');
  })()}

${METRO.map(s => {
    const colors = { 1: '#CC0000', 2: '#0055AA', 3: '#168E44' };
    const labels = { 1: 'М₁', 2: 'М₂', 3: 'М₃' };
    const color = colors[s.линия] || '#888';
    const label = s.пересадка ? 'М' : (labels[s.линия] || 'М');
    const safeName = safe(s.название);
    return `L.marker([${s.lat}, ${s.lng}], {
  icon: L.divIcon({
    className: '',
    html: '<div class="mpin" style="color:${color};border-color:${color}">${label}</div>',
    iconSize: [20, 20], iconAnchor: [10, 10], popupAnchor: [0, -14]
  })
}).bindPopup('<div class="hpop"><b>🚇 ${safeName}</b><br><span style="color:${color};font-size:11px">Линия ${s.линия}${s.пересадка ? ' · пересадка' : ''}</span></div>').addTo(metroLayer);`;
  }).join('\n')}

function apMove(id, delta) {
  const sl = document.getElementById(id);
  if (!sl) return;
  const slides = sl.querySelectorAll('.aslide');
  const total  = slides.length;
  let cur = 0;
  for (let i = 0; i < total; i++) if (slides[i].classList.contains('active')) cur = i;
  const next = (cur + delta + total) % total;
  apGoto(id, next);
}
function apGoto(id, idx) {
  const sl = document.getElementById(id);
  if (!sl) return;
  const slides = sl.querySelectorAll('.aslide');
  const total  = slides.length;
  for (let i = 0; i < total; i++) slides[i].classList.toggle('active', i === idx);
  const cur = sl.querySelector('.acur'); if (cur) cur.textContent = String(idx + 1);
  // dots живут в соседнем .adots — найдём через ближайший .apop
  const root = sl.closest('.apop');
  if (root) {
    const dots = root.querySelectorAll('.adot');
    dots.forEach((d, i) => d.classList.toggle('active', i === idx));
  }
}

document.addEventListener('keydown', e => {
  if (e.key !== 'ArrowLeft' && e.key !== 'ArrowRight') return;
  const sl = document.querySelector('.leaflet-popup .aslider');
  if (!sl) return;
  apMove(sl.id, e.key === 'ArrowLeft' ? -1 : 1);
});

let metroVisible = true;
function toggleMetro() {
  if (metroVisible) { map.removeLayer(metroLayer); document.getElementById('toggle-metro').textContent = '🚇 Метро: выкл'; }
  else { metroLayer.addTo(map); document.getElementById('toggle-metro').textContent = '🚇 Метро: вкл'; }
  metroVisible = !metroVisible;
}

let placesVisible = true;
function togglePlaces() {
  const btn = document.getElementById('toggle-places');
  if (placesVisible) { map.removeLayer(placesLayer); btn.textContent = '🍔 Интересные места: выкл'; }
  else { placesLayer.addTo(map); btn.textContent = '🍔 Интересные места: вкл'; }
  placesVisible = !placesVisible;
}

// Клик по пункту легенды → перелёт + открытие попапа маркера.
function focusItem(name) {
  const m = markerByName[name];
  if (!m) return;
  // Если маркер живёт в placesLayer и слой выключен — включаем обратно
  if (!placesVisible && placesLayer.hasLayer(m)) togglePlaces();
  const ll = m.getLatLng();
  // Загородные точки (Несвиж/Мир/Дудутки) — меньший zoom чтобы виден контекст
  const cityLat = ${CITY_LAT}, cityLng = ${CITY_LNG};
  const isFar = Math.abs(ll.lat - cityLat) > 0.15 || Math.abs(ll.lng - cityLng) > 0.25;
  map.flyTo(ll, isFar ? 11 : 16, { duration: 0.8 });
  setTimeout(() => m.openPopup(), 850);
}
</script>
</body>
</html>`;

  const dateStr = new Date().toISOString().slice(0, 16).replace('T', '_').replace(':', '-');
  const outFile = path.join('output', `hotels_${CITY_SLUG}_${dateStr}_map.html`);
  fs.writeFileSync(outFile, html);
  return outFile;
}

// ── MAIN ──────────────────────────────────────────────────────────────
async function main() {
  const needGeo = allHotels.filter(h => !h.lat || !h.lng);
  if (needGeo.length > 0) {
    console.log(`🌍 Геокодирование ${needGeo.length} отелей без координат...`);
    await geocodeHotels(allHotels, CITY_NAME, config.город.страна);
    console.log('');
  }

  // Валидация фото отключена: Wikimedia режет batch-HEAD случайным образом,
  // выкидывая живые URL'ы как битые. На клиенте onerror подменяет фото на
  // серый плейсхолдер (высота сохраняется), этого достаточно.
  // Чтобы включить валидацию обратно: await validateAttractionPhotos(ATTRACTIONS);

  console.log('📊 Генерация отчёта...\n');

  if (hotels.length === 0) {
    console.log('⚠️  Нет отелей после фильтрации хостелов. Всего записей:', allHotels.length);
    return;
  }

  const wb = new ExcelJS.Workbook();
  buildRecommendationsSheet(wb);
  buildSummarySheet(wb);
  buildAllSheet(wb);
  buildRatingSheet(wb);
  buildSourcesSheet(wb);

  const dateStr  = new Date().toISOString().slice(0, 16).replace('T', '_').replace(':', '-');
  const xlsxFile = path.join('output', `hotels_${CITY_SLUG}_${dateStr}.xlsx`);
  await wb.xlsx.writeFile(xlsxFile);

  const mapFile = buildHtmlMap();

  console.log('✅ Готово!\n');
  console.log(`   Excel: ${xlsxFile}`);
  console.log(`   Карта: ${mapFile}`);
  console.log('\nОткрыть:');
  console.log(`   open "${xlsxFile}"`);
  console.log(`   open "${mapFile}"`);
}

main().catch(err => {
  console.error('❌', err);
  process.exit(1);
});
