'use strict';

const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');

const config = JSON.parse(fs.readFileSync('hotel-config.json', 'utf8'));
const args = process.argv.slice(2);
const ONLY_BOOKING  = args.includes('--booking');
const ONLY_TRIP     = args.includes('--trip');
const ONLY_OSTROVOK = args.includes('--ostrovok');
const runAll        = !ONLY_BOOKING && !ONLY_TRIP && !ONLY_OSTROVOK;
const runBooking    = ONLY_BOOKING  || runAll;
const runTrip       = ONLY_TRIP     || runAll;
const runOstrovok   = ONLY_OSTROVOK || runAll;

const { заезд: checkin, выезд: checkout } = config.даты;
const nights = Math.round((new Date(checkout) - new Date(checkin)) / 86400000);

// ── Output setup ──────────────────────────────────────────────────────
const ts = new Date().toISOString().slice(0, 16).replace('T', '_').replace(':', '-');
const outDir = path.join('output', `run_${ts}`);
fs.mkdirSync(outDir, { recursive: true });
const checkpointFile = path.join(outDir, 'hotels.jsonl');

function saveHotel(hotel) {
  fs.appendFileSync(checkpointFile, JSON.stringify(hotel) + '\n');
  const priceStr = hotel.price_per_night_rub
    ? hotel.price_per_night_rub.toLocaleString('ru-RU') + ' ₽/н'
    : hotel.price_display || 'нет цены';
  const rating = hotel.rating ? ` | ★${hotel.rating}` : '';
  console.log(`  [${hotel.source.padEnd(7)}] ${hotel.name.slice(0, 45).padEnd(45)} ${priceStr}${rating}`);
}

const USD_TO_RUB = 90;

function parsePriceRub(str) {
  if (!str) return null;
  const norm = str.replace(/[  ]/g, ' ');

  // RUB / ₽
  const rubMatch = norm.match(/([\d][\d\s]{2,}[\d])\s*[₽Р]/) ||
                   norm.match(/([₽Р])\s*([\d][\d\s]{2,}[\d])/) ||
                   norm.match(/([\d]{4,})\s*[₽Р]/);
  if (rubMatch) {
    const digits = (rubMatch[1] === '₽' || rubMatch[1] === 'Р') ? rubMatch[2] : rubMatch[1];
    const num = parseInt(digits.replace(/\s/g, ''), 10);
    return isNaN(num) || num < 500 || num > 2000000 ? null : num;
  }

  // USD / $
  const usdMatch = norm.match(/\$\s*([\d][\d,.]*)/) ||
                   norm.match(/([\d][\d,.]*(?:\.\d+)?)\s*USD/i);
  if (usdMatch) {
    const num = parseFloat(usdMatch[1].replace(/,/g, ''));
    if (!isNaN(num) && num >= 5 && num <= 20000) return Math.round(num * USD_TO_RUB);
  }

  return null;
}

function wait(ms) { return new Promise(r => setTimeout(r, ms)); }

async function scrollPage(page, times = 10, delayMs = 800) {
  for (let i = 0; i < times; i++) {
    await page.evaluate(() => window.scrollBy(0, window.innerHeight * 1.5));
    await wait(delayMs);
  }
  await wait(2000);
}

// ── BOOKING.COM ───────────────────────────────────────────────────────

async function scrapeBooking(page) {
  const { взрослых: adults, номеров: rooms } = config.гости;
  const url = [
    'https://www.booking.com/searchresults.ru.html',
    `?ss=${encodeURIComponent(config.город.query_booking)}`,
    `&checkin=${checkin}&checkout=${checkout}`,
    `&group_adults=${adults}&no_rooms=${rooms}&group_children=0`,
    `&lang=ru&selected_currency=RUB&order=price`,
  ].join('');

  console.log('\n╔══════════════════════════════════════════╗');
  console.log('║  BOOKING.COM                             ║');
  console.log('╚══════════════════════════════════════════╝');
  console.log(`URL: ${url}\n`);
  console.log('⏳ Открываю страницу. Если появится капча — решите её.');
  console.log('   Жду 60 секунд перед началом парсинга...\n');

  await page.goto(url, { waitUntil: 'domcontentloaded', timeout: 90000 });
  await wait(60000);

  // Dismiss cookie consent
  for (const sel of [
    '#onetrust-accept-btn-handler',
    'button[data-gdpr-consent="accept"]',
    '[aria-label="Принять"]',
    'button:has-text("Принять")',
    'button:has-text("Accept")',
  ]) {
    const btn = page.locator(sel).first();
    if (await btn.isVisible({ timeout: 1000 }).catch(() => false)) {
      await btn.click().catch(() => {});
      await wait(800);
      break;
    }
  }

  console.log('📜 Прокручиваю страницу для подгрузки всех результатов...');
  await scrollPage(page, 10, 800);
  await page.screenshot({ path: path.join(outDir, 'booking_screenshot.png') });

  const raw = await page.evaluate(() => {
    const cards = [...document.querySelectorAll('[data-testid="property-card"]')];

    return cards.map(card => {
      const t   = (s) => card.querySelector(s)?.textContent?.trim() ?? null;

      const name = t('[data-testid="title"]');
      if (!name) return null;

      // Rating
      const scoreEl = card.querySelector('[data-testid="review-score"]');
      const scoreBlocks = scoreEl ? [...scoreEl.querySelectorAll('div')].map(d => d.textContent.trim()) : [];
      const ratingStr = scoreBlocks.find(s => /^\d[\d,.]?\d?$/.test(s));
      const rating = ratingStr ? parseFloat(ratingStr.replace(',', '.')) : null;
      const ratingLabel = scoreBlocks.find(s => s.length > 2 && /[А-яa-z]/i.test(s)) ?? null;
      const reviewMatch = (scoreEl?.parentElement?.textContent ?? '').match(/(\d[\d\s]*)\s*отзыв/i);
      const reviewCount = reviewMatch ? parseInt(reviewMatch[1].replace(/\s/g, '')) : null;

      // Price — find ₽ in card text
      const cardText = card.textContent;
      const priceMatch = cardText.match(/([\d][\d\s]{2,}[\d])\s*₽/) ||
                         cardText.match(/([\d]{4,})\s*₽/);
      const priceDisplay = priceMatch ? priceMatch[0].trim() : null;

      // Stars
      const starsMatch = (card.querySelector('[aria-label*="звёзд"]')?.getAttribute('aria-label') ??
                          card.querySelector('[aria-label*="star"]')?.getAttribute('aria-label') ?? '')
                         .match(/(\d)/);
      const stars = starsMatch ? parseInt(starsMatch[1]) : null;

      // Location
      const address     = t('[data-testid="address"]');
      const distEl      = card.querySelector('[data-testid="distance"]') ??
                          card.querySelector('[class*="distance"]');
      const distanceText = distEl?.textContent?.trim() ?? null;

      // Coordinates (sometimes present as data attributes on the card or parent)
      const lat = parseFloat(card.getAttribute('data-latitude') ??
                             card.closest('[data-latitude]')?.getAttribute('data-latitude') ?? '') || null;
      const lng = parseFloat(card.getAttribute('data-longitude') ??
                             card.closest('[data-longitude]')?.getAttribute('data-longitude') ?? '') || null;

      const url = card.querySelector('a[href*="/hotel/"]')?.href ?? null;
      const img = card.querySelector('img[data-testid="image"]')?.src ??
                  card.querySelector('img')?.src ?? null;

      return { name, rating, rating_label: ratingLabel, review_count: reviewCount,
               price_display: priceDisplay, stars, address, distance_text: distanceText,
               lat, lng, url, thumbnail: img };
    }).filter(Boolean);
  });

  console.log(`\n✅ Booking.com: найдено ${raw.length} карточек`);
  if (raw.length === 0) {
    console.log('⚠️  Нет результатов. Скриншот: ' + path.join(outDir, 'booking_screenshot.png'));
  }

  for (const h of raw) {
    const totalRub = parsePriceRub(h.price_display);
    saveHotel({
      source: 'booking',
      name: h.name,
      stars: h.stars,
      rating: h.rating,
      rating_label: h.rating_label,
      review_count: h.review_count,
      price_display: h.price_display,
      price_total_rub: totalRub,
      price_per_night_rub: totalRub ? Math.round(totalRub / nights) : null,
      nights,
      address: h.address,
      distance_text: h.distance_text,
      lat: h.lat,
      lng: h.lng,
      url: h.url,
      thumbnail: h.thumbnail,
    });
  }
}

// ── TRIP.COM ──────────────────────────────────────────────────────────

async function scrapeTrip(page) {
  const { взрослых: adults, номеров: rooms } = config.гости;

  // Trip.com accepts YYYYMMDD
  const cin  = checkin.replace(/-/g, '');
  const cout = checkout.replace(/-/g, '');

  const cityParam = config.город.cityid_trip
    ? config.город.cityid_trip
    : encodeURIComponent(config.город.query_trip);

  const url = [
    'https://www.trip.com/hotels/list',
    `?city=${cityParam}`,
    `&checkin=${cin}&checkout=${cout}`,
    `&adult=${adults}&children=0&rooms=${rooms}`,
    `&curr=USD&sortType=4`,
  ].join('');

  console.log('\n╔══════════════════════════════════════════╗');
  console.log('║  TRIP.COM                                ║');
  console.log('╚══════════════════════════════════════════╝');
  console.log(`URL: ${url}\n`);
  console.log('⏳ Открываю страницу... Жду 35 секунд.\n');

  await page.goto(url, { waitUntil: 'domcontentloaded', timeout: 90000 });
  await wait(5000);

  // If redirected to login page — wait up to 120 sec for the user to log in
  if (page.url().includes('/account/signin')) {
    console.log('\n🔑 Trip.com просит войти в аккаунт.');
    console.log('   Войдите в браузере (Google/Apple/Email), потом вернитесь сюда.');
    console.log('   Жду 120 секунд...\n');
    await wait(120000);
  } else {
    await wait(30000);
  }

  // Dismiss consent/popup
  for (const sel of [
    'button:has-text("Accept")',
    'button:has-text("Принять")',
    'button:has-text("OK")',
    '[class*="close"]',
    '[aria-label="Close"]',
  ]) {
    const btn = page.locator(sel).first();
    if (await btn.isVisible({ timeout: 1000 }).catch(() => false)) {
      await btn.click().catch(() => {});
      await wait(500);
      break;
    }
  }

  console.log('📜 Прокручиваю страницу (ищу все отели)...');
  await scrollPage(page, 20, 1200);
  await page.screenshot({ path: path.join(outDir, 'trip_screenshot.png') });

  const currentUrl = page.url();
  console.log(`Текущий URL: ${currentUrl}`);

  const raw = await page.evaluate(() => {
    // Try known Trip.com selectors (they change over time — we try several)
    const CARD_SELECTORS = [
      '[data-testid="hotel-card"]',
      '[class*="hotel-item"]',
      '[class*="hotelItem"]',
      '[class*="HotelItem"]',
      '[class*="property-item"]',
      '.hotel_card',
      '.hotel-list-item',
    ];

    let cards = [];
    for (const sel of CARD_SELECTORS) {
      const found = [...document.querySelectorAll(sel)];
      if (found.length > 2) { cards = found; break; }
    }

    if (cards.length === 0) {
      // Generic fallback: find list container and grab its direct children
      const listEl = document.querySelector(
        '[class*="hotel-list"], [class*="hotelList"], [class*="list-container"], [class*="ListWrapper"]'
      );
      if (listEl) {
        cards = [...listEl.querySelectorAll('li, article, [class*="item"]')]
          .filter(el => el.textContent.length > 80);
      }
    }

    return cards.map(card => {
      const allText = card.textContent;

      // Name
      const nameEl = card.querySelector('h2, h3, [class*="name"], [class*="Name"], [class*="title"], [class*="Title"]');
      const name = nameEl?.textContent?.trim() ?? null;
      if (!name || name.length < 3) return null;

      // Price — RUB (₽), USD ($), EUR (€), BYN
      const priceMatch = allText.match(/([\d][\d\s,]*[\d])\s*[₽Р]/) ||
                         allText.match(/([\d]{4,})\s*[₽Р]/) ||
                         allText.match(/\$\s*([\d][\d,]*[\d])/) ||
                         allText.match(/([\d][\d,]*)\s*USD/) ||
                         allText.match(/([\d][\d,]*)\s*BYN/) ||
                         allText.match(/€\s*([\d][\d,]*[\d])/);
      const priceDisplay = priceMatch ? priceMatch[0].trim() : null;

      // Rating — number like 8.5 or 4.2
      const ratingMatch = allText.match(/\b([5-9]\.\d|[1-9]\d?\.\d)\b/) ||
                          allText.match(/\b(10\.0|10)\b/);
      const rating = ratingMatch ? parseFloat(ratingMatch[0]) : null;

      // Reviews
      const reviewMatch = allText.match(/(\d+)\s*(отзыв|review|оценк)/i);
      const reviewCount = reviewMatch ? parseInt(reviewMatch[1]) : null;

      // Stars — count star SVG icons or look for Xзв / Xstar
      const starIcons = card.querySelectorAll('[class*="star"], [class*="Star"]');
      const starsFromText = allText.match(/(\d)\s*звезд/i);
      const stars = starsFromText ? parseInt(starsFromText[1]) :
                    (starIcons.length >= 2 && starIcons.length <= 5 ? starIcons.length : null);

      // Address / distance — avoid [class*="desc"] which matches rating labels
      const addrEl = card.querySelector('[class*="address"], [class*="location"], [class*="dist"]');
      const address = addrEl?.textContent?.trim() ?? null;

      const url = card.querySelector('a')?.href ?? null;
      const img = card.querySelector('img')?.src ?? null;

      return { name, price_display: priceDisplay, rating, review_count: reviewCount,
               stars, address, url, thumbnail: img };
    }).filter(Boolean);
  });

  console.log(`\n✅ Trip.com: найдено ${raw.length} карточек`);
  if (raw.length === 0) {
    console.log('⚠️  Нет результатов. Скриншот: ' + path.join(outDir, 'trip_screenshot.png'));
    console.log('   Возможно, Trip.com перенаправил на другую страницу или изменил DOM.');
  }

  for (const h of raw) {
    const totalRub = parsePriceRub(h.price_display);
    saveHotel({
      source: 'trip',
      name: h.name,
      stars: h.stars,
      rating: h.rating,
      rating_label: null,
      review_count: h.review_count,
      price_display: h.price_display,
      price_total_rub: totalRub,
      price_per_night_rub: totalRub ? Math.round(totalRub / nights) : null,
      nights,
      address: h.address,
      distance_text: null,
      lat: null,
      lng: null,
      url: h.url,
      thumbnail: h.thumbnail,
    });
  }
}

// ── OSTROVOK.RU ───────────────────────────────────────────────────────

async function scrapeOstrovok(page) {
  const { взрослых: adults, номеров: rooms } = config.гости;
  const citySlug = config.город.slug_ostrovok || config.город.slug;
  const country  = config.город.country_ostrovok || 'russia';

  const url = [
    `https://ostrovok.ru/hotel/${country}/${citySlug}/`,
    `?date_from=${checkin}&date_to=${checkout}`,
    `&adults=${adults}&rooms=${rooms}`,
    `&sort=price`,
  ].join('');

  console.log('\n╔══════════════════════════════════════════╗');
  console.log('║  OSTROVOK.RU                             ║');
  console.log('╚══════════════════════════════════════════╝');
  console.log(`URL: ${url}\n`);
  console.log('⏳ Открываю страницу... Жду 25 секунд.\n');

  await page.goto(url, { waitUntil: 'domcontentloaded', timeout: 90000 });
  await wait(25000);

  // Dismiss cookie banner
  const cookieBtn = page.locator('button:has-text("Хорошо"), button:has-text("Принять")').first();
  if (await cookieBtn.isVisible({ timeout: 3000 }).catch(() => false)) {
    await cookieBtn.click().catch(() => {});
    await wait(500);
  }

  // Handle the date-picker modal Ostrovok shows when URL dates aren't recognized
  const modalVisible = await page.locator('[role="dialog"], [class*="ModalHeader"]').first()
    .isVisible({ timeout: 3000 }).catch(() => false);

  if (modalVisible) {
    console.log('📅 Вижу модалку — нажимаю «Найти» чтобы понять формат URL...');

    // Step 1: Click "Найти" with whatever dates are pre-filled
    const submitBtn = page.locator('[role="dialog"] button:has-text("Найти")').first();
    if (await submitBtn.isVisible({ timeout: 2000 }).catch(() => false)) {
      await submitBtn.click();
      await wait(8000);
    }

    const urlAfterSubmit = page.url();
    console.log(`   URL после Найти: ${urlAfterSubmit}`);

    // Step 2: If the URL now contains date info, replace with our target dates
    if (urlAfterSubmit.includes('date_from') || urlAfterSubmit.includes('checkin') || urlAfterSubmit.includes('from=')) {
      const correctedUrl = urlAfterSubmit
        .replace(/date_from=[^&]+/, `date_from=${checkin}`)
        .replace(/date_to=[^&]+/, `date_to=${checkout}`)
        .replace(/checkin=[^&]+/, `checkin=${checkin}`)
        .replace(/checkout=[^&]+/, `checkout=${checkout}`)
        .replace(/from=[^&]+/, `from=${checkin}`)
        .replace(/to=[^&]+/, `to=${checkout}`);
      console.log(`   Переходим на наши даты: ${correctedUrl}`);
      await page.goto(correctedUrl, { waitUntil: 'domcontentloaded', timeout: 60000 });
      await wait(10000);
    }
  }

  console.log('📜 Прокручиваю страницу...');
  await scrollPage(page, 18, 1000);
  await page.screenshot({ path: path.join(outDir, 'ostrovok_screenshot.png') });

  const currentUrl = page.url();
  console.log(`Текущий URL: ${currentUrl}`);

  const raw = await page.evaluate(() => {
    // Ostrovok uses CSS module classes like HotelCard_container__xxx
    let cards = [...document.querySelectorAll('[class*="HotelCard_container"]')];

    // Fallback for DOM changes
    if (cards.length === 0) {
      cards = [...document.querySelectorAll('[class*="HotelCard"], [data-hotel-id]')]
        .filter(el => el.textContent.length > 80);
    }

    // Deduplicate by name (avoid nested matches)
    const seen = new Set();

    return cards.map(card => {
      const allText = card.textContent;

      // Name — prefer h2/h3
      const name = card.querySelector('h2')?.textContent?.trim()
               ?? card.querySelector('h3')?.textContent?.trim()
               ?? card.querySelector('[class*="name" i]')?.textContent?.trim()
               ?? null;
      if (!name || name.length < 3 || seen.has(name)) return null;
      seen.add(name);

      // Price — RUB ₽
      const priceMatch = allText.match(/([\d][\d\s]*[\d])\s*[₽Р]/) ||
                         allText.match(/([\d]{4,})\s*[₽Р]/);
      const priceDisplay = priceMatch ? priceMatch[0].trim() : null;

      // Rating
      const ratingMatch = allText.match(/\b([5-9]\.\d|[1-9]\d?\.\d|10(?:\.0)?)\b/);
      const rating = ratingMatch ? parseFloat(ratingMatch[0]) : null;

      // Reviews
      const reviewMatch = allText.match(/(\d[\d\s]*)\s*(отзыв|оценк)/i);
      const reviewCount = reviewMatch ? parseInt(reviewMatch[1].replace(/\s/g, '')) : null;

      // Stars
      const starsEl = card.querySelector('[class*="stars"], [class*="Stars"], [aria-label*="звезд"]');
      const starsText = starsEl?.getAttribute('aria-label') || starsEl?.textContent || '';
      const starsMatch = starsText.match(/(\d)/);
      const stars = starsMatch ? parseInt(starsMatch[1]) : null;

      // Address / distance
      const addrEl = card.querySelector('[class*="address"], [class*="location"], [class*="distance"]');
      const address = addrEl?.textContent?.trim() ?? null;

      const url = card.querySelector('a[href*="/hotel/"]')?.href ??
                  card.querySelector('a')?.href ?? null;
      const img = card.querySelector('img')?.src ?? null;

      return { name, price_display: priceDisplay, rating, review_count: reviewCount,
               stars, address, url, thumbnail: img };
    }).filter(Boolean);
  });

  console.log(`\n✅ Ostrovok.ru: найдено ${raw.length} карточек`);
  if (raw.length === 0) {
    console.log('⚠️  Нет результатов. Скриншот: ' + path.join(outDir, 'ostrovok_screenshot.png'));
  }

  for (const h of raw) {
    const totalRub = parsePriceRub(h.price_display);
    saveHotel({
      source: 'ostrovok',
      name: h.name,
      stars: h.stars,
      rating: h.rating,
      rating_label: null,
      review_count: h.review_count,
      price_display: h.price_display,
      price_total_rub: totalRub,
      price_per_night_rub: totalRub ? Math.round(totalRub / nights) : null,
      nights,
      address: h.address,
      distance_text: null,
      lat: null,
      lng: null,
      url: h.url,
      thumbnail: h.thumbnail,
    });
  }
}

// ── MAIN ──────────────────────────────────────────────────────────────

async function main() {
  console.log('🏨  Hotel Parser');
  console.log(`📍  ${config.город.название}, ${config.город.страна}`);
  console.log(`📅  ${checkin} → ${checkout} (${nights} ${nights === 1 ? 'ночь' : nights < 5 ? 'ночи' : 'ночей'})`);
  console.log(`👤  ${config.гости.взрослых} взрослых, ${config.гости.номеров} номер`);
  console.log(`📁  ${outDir}\n`);

  // Persistent context: cookies are saved between runs so login is only needed once
  const sessionDir = path.join(__dirname, 'session');
  const ctx = await chromium.launchPersistentContext(sessionDir, {
    headless: false,
    viewport: null,
    userAgent: 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36',
    locale: 'ru-RU',
  });
  const page = await ctx.newPage();

  try {
    if (runBooking)   await scrapeBooking(page);
    if (runTrip)      await scrapeTrip(page);
    if (runOstrovok)  await scrapeOstrovok(page);
  } catch (err) {
    if (err.message.includes('closed')) {
      console.error('\n❌ Браузер был закрыт вручную. Запустите снова.');
    } else {
      console.error('\n❌ Ошибка парсинга:', err.message);
    }
  } finally {
    await ctx.close();
  }

  if (!fs.existsSync(checkpointFile)) {
    console.log('\n⚠️  Нет данных для сохранения.');
    return;
  }

  const lines    = fs.readFileSync(checkpointFile, 'utf8').trim().split('\n').filter(Boolean);
  let   hotels   = lines.map(l => JSON.parse(l));

  // When running a subset of sources, preserve data from other sources in latest.json
  if (!runAll) {
    const sourcesRun = new Set([
      ...(runBooking  ? ['booking']  : []),
      ...(runTrip     ? ['trip']     : []),
      ...(runOstrovok ? ['ostrovok'] : []),
    ]);
    const latestFile = path.join('output', 'latest.json');
    if (fs.existsSync(latestFile)) {
      const prev = JSON.parse(fs.readFileSync(latestFile, 'utf8'));
      const kept = prev.filter(h => !sourcesRun.has(h.source));
      if (kept.length > 0) {
        console.log(`\n🔀 Сохраняю ${kept.length} отелей из предыдущих источников`);
        hotels = [...kept, ...hotels];
      }
    }
  }

  fs.writeFileSync(path.join(outDir, 'hotels.json'), JSON.stringify(hotels, null, 2));
  fs.mkdirSync('output', { recursive: true });
  fs.writeFileSync(path.join('output', 'latest.json'), JSON.stringify(hotels, null, 2));
  fs.writeFileSync(path.join('output', 'latest_dir.txt'), outDir);

  console.log(`\n🎉 Готово!`);
  console.log(`   Всего:        ${hotels.length}`);
  console.log(`   Booking.com:  ${hotels.filter(h => h.source === 'booking').length}`);
  console.log(`   Trip.com:     ${hotels.filter(h => h.source === 'trip').length}`);
  console.log(`   Ostrovok.ru:  ${hotels.filter(h => h.source === 'ostrovok').length}`);
  console.log(`   С ценой:      ${hotels.filter(h => h.price_per_night_rub).length}`);
  console.log(`\n▶  Следующий шаг: node build-report.js`);
}

main().catch(err => {
  console.error('Fatal:', err);
  process.exit(1);
});
