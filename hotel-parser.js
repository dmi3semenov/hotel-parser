'use strict';

const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');

const config = JSON.parse(fs.readFileSync('hotel-config.json', 'utf8'));
const args = process.argv.slice(2);
const ONLY_BOOKING  = args.includes('--booking');
const ONLY_TRIP     = args.includes('--trip');
const ONLY_OSTROVOK = args.includes('--ostrovok');
const ONLY_AGODA    = args.includes('--agoda');
const ONLY_GOOGLE   = args.includes('--google');
const runAll        = !ONLY_BOOKING && !ONLY_TRIP && !ONLY_OSTROVOK && !ONLY_AGODA
                      && !ONLY_GOOGLE;
const runBooking    = ONLY_BOOKING  || runAll;
const runTrip       = ONLY_TRIP     || runAll;
const runOstrovok   = ONLY_OSTROVOK || runAll;
const runAgoda      = ONLY_AGODA    || runAll;
const runGoogle     = ONLY_GOOGLE   || runAll;

// Google отдаёт список по 18-20 карточек и листается кнопкой «Далее» внизу.
// По умолчанию идём до конца - цикл сам остановится, когда кнопка пропадёт.
// 200 здесь не потолок выгрузки, а предохранитель от бесконечного цикла.
const gPagesArg = args.find(a => a.startsWith('--google-pages='));
const GOOGLE_MAX_PAGES = gPagesArg
  ? (parseInt(gPagesArg.split('=')[1], 10) || 200) : 200;

// Второй проход: по странице каждого отеля за блоком «Все варианты».
// Он и есть смысл источника - там лежат Traveloka, klook, Prestigia и прочие,
// которых не парсит никто из остальных четырёх. Флаг для быстрой проверки
// списка без него, лимит - для отладки на десятке отелей.
const GOOGLE_OFFERS = !args.includes('--google-no-offers');
const gLimitArg = args.find(a => a.startsWith('--google-offers-limit='));
const GOOGLE_OFFERS_LIMIT = gLimitArg ? parseInt(gLimitArg.split('=')[1], 10) || 0 : 0;
// Страница отеля берётся обычным GET без рендера, поэтому её можно тянуть
// в несколько потоков. Пять - осторожная величина: одна страница приезжает
// примерно за полторы секунды, пяти хватает на тысячу отелей за пять минут,
// и это далеко от той нагрузки, на которой Google начинает отдавать 429.
const gConcArg = args.find(a => a.startsWith('--google-concurrency='));
const GOOGLE_CONCURRENCY = gConcArg ? parseInt(gConcArg.split('=')[1], 10) || 5 : 5;

// Третий проход: отели, которые нашли остальные четыре источника, а в списке
// Google их нет. Их ищем поимённо. Файл - выгрузка с другими источниками.
//
// При полном прогоне это чекпоинт ТЕКУЩЕГО прогона: Google идёт последним,
// остальные четыре источника к тому моменту в него уже дописались. Раньше
// по умолчанию брался output/latest.json, а в нём при смене города лежит
// ПРОШЛЫЙ город - парсер честно шёл искать в Дананге нячангские отели.
// При запуске одного источника latest.json как раз то, что нужно: там
// соседи с прошлого прогона по этому же городу.
// Значение считается ниже, после того как определён checkpointFile.
const gMatchArg = args.find(a => a.startsWith('--google-match='));

// Сколько страниц выдачи проходить. Без флага число берётся со страницы
// («Стр. 1 из 10»), а не зашивается: для другого города оно другое.
const pagesArg = args.find(a => a.startsWith('--agoda-pages='));
const AGODA_MAX_PAGES = pagesArg ? (parseInt(pagesArg.split('=')[1], 10) || 0) : 0;

// Прокрутка выдачи Agoda. Доля экрана — единственная величина здесь, от
// которой зависит полнота сбора: она обязана быть МЕНЬШЕ единицы, иначе
// между кадрами остаётся слепая полоса. Почему это ломает выгрузку —
// подробно у agodaHarvestVisible. Остальные три числа — предохранители от
// бесконечного цикла, а не потолок выгрузки: страница останавливает сбор
// сама, когда перестаёт расти.
const AGODA_STEP_RATIO = 0.8;
const AGODA_SCROLL_PAUSE = 1200;
const AGODA_STALE_STOP = 4;
const AGODA_MAX_STEPS = 250;
const AGODA_PAGE_BUDGET_MS = 180000;

// Островок: по 20 карточек на страницу, листается кнопкой «Вперед».
// По умолчанию идём до конца — цикл сам остановится, когда кнопка «Вперед»
// пропадёт. 200 здесь не потолок выгрузки, а предохранитель от бесконечного
// цикла: с прежним значением 15 выгрузка молча обрывалась на 297 отелях.
const oPagesArg = args.find(a => a.startsWith('--ostrovok-pages='));
const OSTROVOK_MAX_PAGES = oPagesArg
  ? (parseInt(oPagesArg.split('=')[1], 10) || 200) : 200;

const { заезд: checkin, выезд: checkout } = config.даты;
const nights = Math.round((new Date(checkout) - new Date(checkin)) / 86400000);

// ── Output setup ──────────────────────────────────────────────────────
// Секунды в метке обязательны. С точностью до минуты два прогона, запущенных
// подряд, получают ОДНУ папку и дописываются в один hotels.jsonl: 15.08 так
// сложились догоняющие прогоны Agoda по Гонконгу и по Гуанчжоу, и выгрузка
// двух городов оказалась в одном файле. Повезло, что второй собрал ноль,
// иначе города перемешались бы молча.
const ts = new Date().toISOString().slice(0, 19).replace('T', '_').replace(/:/g, '-');
const outDir = path.join('output', `run_${ts}`);
fs.mkdirSync(outDir, { recursive: true });
const checkpointFile = path.join(outDir, 'hotels.jsonl');

const GOOGLE_MATCH_FILE = args.includes('--google-no-match') ? null
  : (gMatchArg ? gMatchArg.split('=')[1]
    : (runAll ? checkpointFile : path.join('output', 'latest.json')));

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

  // Trip.com печатает КОД валюты перед числом и запятую в разрядах:
  // «Total price: RUB 14,027». Ни один шаблон ниже такого не ловил —
  // из-за этого у Trip.com цена не снималась вообще (12 отелей, 0 цен).
  const rubCode = norm.match(/RUB\s*([\d][\d,\s]*\d)/i);
  if (rubCode) {
    const num = parseInt(rubCode[1].replace(/[,\s]/g, ''), 10);
    if (!isNaN(num) && num >= 100 && num <= 2000000) return num;
  }

  // RUB / ₽ / «руб.» (Booking.ru показывает цены текстом «4 469 руб.», а не символом)
  const rubMatch = norm.match(/([\d][\d\s]{2,}[\d])\s*(?:[₽Р]|руб)/i) ||
                   norm.match(/([₽Р])\s*([\d][\d\s]{2,}[\d])/) ||
                   norm.match(/([\d]{4,})\s*(?:[₽Р]|руб)/i);
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

// Залогинен ли пользователь. Цены с уровнем лояльности (Genius, GURU,
// Trip Rewards, Agoda VIP) отличаются от анонимных, поэтому состояние входа
// надо не предполагать, а печатать в лог каждого прогона.
// Вход засчитываем ТОЛЬКО по положительному признаку: отсутствие кнопки
// «Войти» ещё ничего не значит — её может просто не быть в этой вёрстке.
async function logLoginState(page, label) {
  try {
    const r = await page.evaluate(() => {
      const acc = ['[data-selenium="user-avatar"]', '[data-element-name="user-avatar"]',
        '[data-testid="header-avatar"]', '[data-testid="avatar"]',
        '[class*="Avatar" i]', '[class*="userMenu" i]', '[class*="ProfileButton" i]']
        .filter(s => { try { return document.querySelector(s); } catch { return false; } });
      const txt = (document.body.innerText || '');
      const signIn = [...document.querySelectorAll('button,a,span')]
        .map(e => (e.innerText || '').trim())
        .filter(t => t && t.length < 28
          && /^(войти|вход|зарегистрироваться|создать аккаунт|sign in|register|log in)$/i.test(t));
      const member = /genius|мои бронирования|мой аккаунт|мои поездки|личный кабинет|уровень|бонусны|agodacash|trip coins/i
        .test(txt.slice(0, 6000));
      return { acc, signIn: [...new Set(signIn)].slice(0, 3), member };
    });
    // Кнопка «Войти» на странице — надёжный признак, что НЕ вошёл.
    // Её отсутствие — слабый признак обратного: у Booking и Островка
    // узел аватара называется не так, как в списке выше, и acc там 0
    // даже у залогиненного пользователя.
    const verdict = r.signIn.length ? 'НЕ вошёл'
      : (r.acc.length || r.member) ? 'вошёл'
      : 'скорее вошёл (кнопок входа нет)';
    console.log(`🔑 ${label}: ${verdict}`
      + `  [аккаунт-узлы: ${r.acc.length || '—'}, кнопки входа: ${r.signIn.join('/') || '—'}`
      + `, признаки кабинета: ${r.member ? 'есть' : 'нет'}]`);
  } catch (e) {
    console.log(`🔑 ${label}: проверить не удалось (${e.message.split('\n')[0].slice(0, 50)})`);
  }
}

// Booking и Островок отдают первую порцию карточек, а остальное прячут за
// кнопкой «Показать ещё». Фиксированного числа прокруток не хватает: раньше
// из-за этого выходило 73 отеля у Booking и 20 у Островка вместо всей выдачи.
// Крутим вниз и жмём кнопку, пока карточки прибавляются.
async function loadAllResults(page, cardSelector, buttonLabels, maxRounds = 45) {
  const sel = buttonLabels.map(l => `button:has-text("${l}"), a:has-text("${l}")`).join(', ');
  let last = 0, stale = 0;
  for (let i = 0; i < maxRounds && stale < 3; i++) {
    await page.evaluate(() => window.scrollTo(0, document.body.scrollHeight));
    await wait(1600);

    const btn = page.locator(sel).first();
    if (await btn.isVisible({ timeout: 2000 }).catch(() => false)) {
      await btn.click().catch(() => {});
      await wait(3000);
    }

    const n = await page.locator(cardSelector).count().catch(() => 0);
    if (n <= last) stale++; else { stale = 0; last = n; }
    if (i % 4 === 0) console.log(`   ...карточек на странице: ${n}`);
  }
  console.log(`   итого карточек в DOM: ${last}`);
  return last;
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
    `&nflt=ht_id%3D204`,   // только тип «Отели» (без апартаментов/хостелов/гестхаусов)
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

  await logLoginState(page, 'Booking.com');

  console.log('📜 Прокручиваю страницу для подгрузки всех результатов...');
  // Кнопка называется «Загрузить больше результатов» — именно она, а не
  // «Показать ещё». В фильтрах слева Booking сам пишет «Отели 258», так что
  // 75 в выгрузке означало, что до кнопки просто не дожали.
  await loadAllResults(page, '[data-testid="property-card"]',
    ['Загрузить больше результатов', 'Загрузить ещё результаты',
     'Показать ещё результаты', 'Load more results']);
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

      // ── Цена ──────────────────────────────────────────────────────────
      // Итог за все ночи лежит в отдельном узле:
      //   <span data-testid="price-and-discounted-price">2 513 руб.</span>
      // Строкой выше стоит «3 ночи, 2 взрослых» (`price-for-x-nights`), ниже —
      // «Включая налоги и сборы» (`taxes-and-charges`). То есть узел — это
      // цена за весь период на искомое число гостей, с налогами.
      //
      // Когда у отеля скидка, РЯДОМ лежит цена без скидки — зачёркнутая
      // (`text-decoration: line-through`), и в тексте карточки она идёт ПЕРВОЙ:
      //   «3 ночи, 2 взрослых5 026 руб.2 513 руб.Прежняя цена составляла…»
      // Отсюда и брались завышенные цены: разбор по тексту карточки не
      // различает зачёркнутое, поэтому берём узел, а не текст.
      const cardText = card.textContent;
      const priceNum = (s, min) => {
        const m = (s || '').match(/([\d][\d\s]*\d)\s*(?:руб|₽)/i);
        if (!m) return null;
        const n = parseInt(m[1].replace(/\s/g, ''), 10);
        return isNaN(n) || n < min ? null : n;
      };

      // Листовые узлы с ценой и признаком «зачёркнута». getComputedStyle
      // считается один раз на узел с «руб» — по всем узлам карточки это сотни
      // пересчётов стиля на карточку и заметная задержка на всей выдаче.
      const priceLeaves = [...card.querySelectorAll('*')]
        .filter(el => !el.children.length && /руб|₽/.test(el.textContent || ''))
        .map(el => ({ el, text: el.textContent,
                      struck: getComputedStyle(el).textDecorationLine === 'line-through' }));

      let priceRub = null, priceFrom = null;
      // 1. Узел итога. В карточке он один, но при нескольких вариантах
      //    размещения берём минимальный — как и остальные источники, которые
      //    отдают самое дешёвое предложение на эти даты.
      const totals = [...card.querySelectorAll('[data-testid="price-and-discounted-price"]')]
        .map(el => priceNum(el.textContent, 100)).filter(Boolean);
      if (totals.length) { priceRub = Math.min(...totals); priceFrom = 'узел итога'; }
      // 2. Резерв на переименование testid: строка для скринридера
      //    «Прежняя цена составляла X руб.. Текущая цена составляет Y руб.».
      if (priceRub === null) {
        const cur = [...cardText.matchAll(/Текущая цена составляет\s*([\d][\d\s]*\d)\s*руб/gi)]
          .map(m => parseInt(m[1].replace(/\s/g, ''), 10)).filter(n => !isNaN(n));
        if (cur.length) { priceRub = Math.min(...cur); priceFrom = 'строка скринридера'; }
      }
      // 3. Последний резерв — цены из НЕзачёркнутых узлов, минимальная из них.
      //    Максимум брать нельзя: это и есть цена до скидки.
      if (priceRub === null) {
        const alive = priceLeaves
          .filter(l => !l.struck && !/налог|сбор|%/i.test(l.text))
          .map(l => priceNum(l.text, 500)).filter(Boolean);
        if (alive.length) { priceRub = Math.min(...alive); priceFrom = 'незачёркнутые узлы'; }
      }
      const priceDisplay = priceRub === null ? null
        : priceRub.toString().replace(/\B(?=(\d{3})+(?!\d))/g, ' ') + ' руб.';
      // Зачёркнутая цена идёт не в выгрузку, а в лог прогона: по ней видно,
      // сколько накручивал прежний разбор, и заметно, если Booking снова
      // поменяет вёрстку.
      const struckPrices = priceLeaves.filter(l => l.struck)
        .map(l => priceNum(l.text, 100)).filter(Boolean);
      const strikeRub = struckPrices.length ? Math.max(...struckPrices) : null;

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
               price_display: priceDisplay, price_rub: priceRub, price_from: priceFrom,
               price_strike: strikeRub, stars, address, distance_text: distanceText,
               lat, lng, url, thumbnail: img };
    }).filter(Boolean);
  });

  console.log(`\n✅ Booking.com: найдено ${raw.length} карточек`);
  if (raw.length === 0) {
    console.log('⚠️  Нет результатов. Скриншот: ' + path.join(outDir, 'booking_screenshot.png'));
  }

  // Чем снята цена — в лог каждого прогона. Шаблон «N руб за 3 ночи» писался
  // 03.06.2026 под тогдашнюю вёрстку; когда он перестал совпадать, разбор молча
  // съехал на фолбэк и стал отдавать цену до скидки — ни один прогон об этом
  // не сообщил. Печатаем стратегию, чтобы такое было видно сразу, а не через
  // сверку с Google.
  const byFrom = {};
  for (const h of raw) byFrom[h.price_from ?? 'цены нет'] = (byFrom[h.price_from ?? 'цены нет'] ?? 0) + 1;
  console.log(`   цена снята: ${Object.entries(byFrom).map(([k, v]) => `${k} — ${v}`).join(', ')}`);
  if (raw.length && !byFrom['узел итога']) {
    console.log('⚠️  Ни одной цены из [data-testid="price-and-discounted-price"] — '
              + 'похоже, Booking переименовал узел. Проверьте разбор, прежде чем верить цифрам.');
  }
  const disc = raw.filter(h => h.price_strike && h.price_rub && h.price_strike > h.price_rub)
                  .map(h => h.price_strike / h.price_rub).sort((a, b) => a - b);
  if (disc.length) {
    console.log(`   скидка показана у ${disc.length} из ${raw.length} карточек, `
              + `медиана «цена до скидки / настоящая» ${disc[Math.floor(disc.length / 2)].toFixed(2)}x`);
  }

  for (const h of raw) {
    const totalRub = h.price_rub ?? parsePriceRub(h.price_display);
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

// Селекторы сняты с живой выдачи 13.08.2026, не угаданы.
//
// Старый код перебирал предполагаемые классы ([class*="hotel-item"] и т.п.),
// ни один не совпадал, и срабатывал generic-фолбэк по [class*="item"] —
// отсюда склеенные имена вроде «Autumn Sea Boutique HotelNew to Trip.com».
// Настоящая карточка — div.hotel-card-versionb, её id это id отеля.
//
// Цену не брал ни один шаблон: Trip печатает код валюты ПЕРЕД числом и с
// запятой в разрядах — «RUB 4,717», а в коде были только «N ₽» и «N USD».
const TRIP_CARD = 'div.hotel-card-versionb';

// Trip подгружает партиями и плохо реагирует на рывки — крутим мягко
async function tripScrollAll(page) {
  let last = 0, stale = 0;
  for (let i = 0; i < 40 && stale < 5; i++) {
    await page.evaluate(() => window.scrollBy(0, window.innerHeight * 0.8));
    await wait(1600);
    const n = await page.locator(TRIP_CARD).count().catch(() => 0);
    if (n <= last) stale++; else { stale = 0; last = n; }
  }
  await wait(3000);
  console.log(`   карточек в DOM: ${last}`);
  return last;
}

// Исполняется в странице. Не должна ссылаться на внешние переменные.
function tripExtract() {
  const money = s => {
    const m = String(s || '').match(/RUB\s*([\d][\d,\s]*\d)/i);
    if (!m) return null;
    const v = parseInt(m[1].replace(/[,\s]/g, ''), 10);
    return isNaN(v) ? null : v;
  };
  const int = s => {
    const m = String(s || '').match(/(\d[\d,\s]*)/);
    if (!m) return null;
    const v = parseInt(m[1].replace(/[,\s]/g, ''), 10);
    return isNaN(v) ? null : v;
  };

  return [...document.querySelectorAll('div.hotel-card-versionb')].map(card => {
    const q = s => card.querySelector(s);
    const t = s => q(s)?.textContent?.trim() || null;

    const name = t('a.hotelName');
    if (!name) return null;

    // «Total price: RUB 16,048 / 1 room × 3 nights incl. taxes & fees»
    const explain = t('div.price-explain-versionb') || '';
    const nightsM = explain.match(/(\d+)\s*night/i);

    // В price-line-versionb два числа: до скидки и со скидкой.
    // Последнее — то, что показано крупно.
    const line = [...card.querySelectorAll('div.price-line-versionb span')]
      .map(s => money(s.textContent)).filter(v => v != null);

    const starN = card.querySelectorAll('[class*="star-item"], [class*="starItem"]').length;

    return {
      hotel_id: card.id || null,
      name,
      total_rub: money(t('span.price-highlight')) || money(explain),
      per_night_rub: line.length ? line[line.length - 1] : null,
      covers_nights: nightsM ? parseInt(nightsM[1], 10) : null,
      price_line: t('div.price-line-versionb'),
      rating: parseFloat((t('span.score') || '').replace(',', '.')) || null,
      rating_label: t('span.comment-desc-versionb'),
      review_count: int(t('span.comment-num')),
      stars: starN >= 1 && starN <= 5 ? starN : null,
      address: [...card.querySelectorAll('span.position-desc')]
        .map(e => e.textContent.trim()).filter(Boolean).join(', ') || null,
      room: t('div.room-name'),
      url: q('a.hotelName')?.href || q('a')?.href || null,
      thumbnail: q('img')?.src || null,
    };
  }).filter(Boolean);
}

async function scrapeTrip(page) {
  const { взрослых: adults, номеров: rooms } = config.гости;

  // Trip.com accepts YYYYMMDD
  const cin  = checkin.replace(/-/g, '');
  const cout = checkout.replace(/-/g, '');

  console.log('\n╔══════════════════════════════════════════╗');
  console.log('║  TRIP.COM                                ║');
  console.log('╚══════════════════════════════════════════╝');

  if (config.город.cityid_trip) {
    // Числовой id есть — идём прямым URL, это быстрее и надёжнее
    // Параметр называется cityId, а не city: со старым city=<id> Trip.com
    // молча отдаёт пустую выдачу. Id смотрится в URL после ручного поиска.
    const url = [
      'https://www.trip.com/hotels/list',
      `?cityId=${config.город.cityid_trip}`,
      `&checkin=${cin}&checkout=${cout}`,
      `&adult=${adults}&children=0&rooms=${rooms}`,
      `&curr=RUB&locale=ru-RU&sortType=4`,
    ].join('');
    console.log(`URL: ${url}\n⏳ Открываю страницу...\n`);
    await page.goto(url, { waitUntil: 'domcontentloaded', timeout: 90000 });
    await wait(5000);
  } else {
    // Без id параметр city=<текст> Trip.com не понимает и отдаёт пустую выдачу,
    // а их API поиска городов закрыт (403/405). Поэтому ищем как человек:
    // печатаем город в поле, берём первую подсказку, кликаем даты в календаре.
    console.log('Ищу через интерфейс: cityid_trip в конфиге не задан.\n');
    await page.goto('https://www.trip.com/hotels/?locale=ru-RU&curr=RUB', {
      waitUntil: 'domcontentloaded', timeout: 90000,
    });
    await wait(8000);

    // Реальные id на trip.com/hotels: destinationInput, checkInInput, checkOutInput
    const box = page.locator(
      '#destinationInput, #hotels-destination, [data-testid="destination-input"], '
      + 'input[placeholder*="Where"], input[placeholder*="аправлен"]').first();
    await box.click({ timeout: 25000 });
    await box.fill(config.город.query_trip || config.город.название);
    await wait(3500);
    await page.keyboard.press('ArrowDown').catch(() => {});
    await page.keyboard.press('Enter').catch(() => {});
    await wait(3000);

    for (const [inputSel, d] of [['#checkInInput', checkin], ['#checkOutInput', checkout]]) {
      await page.locator(inputSel).first().click({ timeout: 8000 }).catch(() => {});
      await wait(1500);
      const cell = page.locator(
        `[data-date="${d}"], td[data-date="${d}"], [aria-label*="${d}"]`).first();
      if (await cell.isVisible({ timeout: 6000 }).catch(() => false)) {
        await cell.click().catch(() => {});
        await wait(1200);
      } else {
        console.log(`   ⚠️  день ${d} в календаре не найден`);
      }
    }

    const search = page.locator(
      '[data-testid="search-button"], button:has-text("Найти"), '
      + 'button:has-text("Search")').first();
    await search.click({ timeout: 15000 }).catch(() => {});
    console.log('⏳ Жду выдачу...\n');
    await wait(20000);
  }

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

  await logLoginState(page, 'Trip.com');

  console.log('📜 Прокручиваю страницу (ищу все отели)...');
  await tripScrollAll(page);
  await page.screenshot({ path: path.join(outDir, 'trip_screenshot.png') });

  console.log(`Текущий URL: ${page.url()}`);

  const raw = await page.evaluate(tripExtract);

  console.log(`\n✅ Trip.com: найдено ${raw.length} карточек`);
  if (raw.length === 0) {
    console.log('⚠️  Нет результатов. Скриншот: ' + path.join(outDir, 'trip_screenshot.png'));
    console.log('   Возможно, Trip.com перенаправил на другую страницу или изменил DOM.');
  }

  let saved = 0, noPrice = 0;
  for (const h of raw) {
    // «Total price: RUB 16 048» при «1 room × 3 nights incl. taxes & fees» —
    // это цена за ВЕСЬ период с налогами и сборами. Остальные источники тоже
    // хранят период, поэтому берём именно её, а не ставку×nights:
    // 4 717 × 3 = 14 151, а реальный итог 16 048 — разница в налогах.
    let total = h.total_rub;
    if (total && h.covers_nights && h.covers_nights !== nights) {
      // Подпись обещает другое число ночей — приводим к нашему
      total = Math.round(total / h.covers_nights * nights);
    }
    if (!total && h.per_night_rub) total = h.per_night_rub * nights;
    if (!total) { noPrice++; continue; }

    saveHotel({
      source: 'trip',
      name: h.name,
      stars: h.stars,
      rating: h.rating,
      rating_label: h.rating_label,
      review_count: h.review_count,
      price_display: h.price_line || `RUB ${total.toLocaleString('ru-RU')}`,
      price_total_rub: total,
      price_per_night_rub: Math.round(total / nights),
      nights,
      address: h.address,
      distance_text: null,
      room: h.room,
      lat: null,
      lng: null,
      url: h.url,
      thumbnail: h.thumbnail,
    });
    saved++;
  }
  console.log(`   сохранено ${saved}, без цены ${noPrice}`);
}

// ── OSTROVOK.RU ───────────────────────────────────────────────────────

// Перелистывание кнопкой «Вперед». Ждём смены первой карточки, иначе
// соберём ту же страницу второй раз.
async function ostrovokNextPage(page) {
  const btn = page.locator('button[data-testid="next-pagination-button"]').first();
  if (!await btn.isVisible({ timeout: 5000 }).catch(() => false)) return false;
  if (await btn.isDisabled().catch(() => false)) return false;

  // Смену страницы ловим по параметру page= в URL, а не по заголовку первой
  // карточки: заголовок не всегда лежит в h2/h3, из-за чего проверка молча
  // не срабатывала и парсер решал, что страницы кончились, уже стоя на page=2.
  const pageNo = () => {
    const m = page.url().match(/[?&]page=(\d+)/);
    return m ? parseInt(m[1], 10) : 1;
  };

  const before = pageNo();
  await btn.click().catch(() => {});

  for (let i = 0; i < 25; i++) {
    await wait(1000);
    if (pageNo() !== before) { await wait(4000); return true; }
  }
  return false;
}

async function scrapeOstrovok(page) {
  const { взрослых: adults, номеров: rooms } = config.гости;
  const citySlug = config.город.slug_ostrovok || config.город.slug;
  const country  = config.город.country_ostrovok || 'russia';

  // Ostrovok принимает даты в формате dates=DD.MM.YYYY-DD.MM.YYYY&guests=N.
  // Старый date_from/date_to НЕ распознавался → вылезала модалка с дефолтными
  // датами (6-7 июня), и парсились цены за 1 ночь не на те даты.
  const toDM = s => s.split('-').reverse().join('.');  // 2026-06-10 → 10.06.2026
  const url = [
    `https://ostrovok.ru/hotel/${country}/${citySlug}/`,
    `?dates=${toDM(checkin)}-${toDM(checkout)}`,
    `&guests=${adults}`,
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

  await logLoginState(page, 'Ostrovok.ru');

  console.log('📜 Собираю выдачу постранично...');

  // Островок отдаёт по 20 карточек на страницу и перелистывается кнопкой
  // «Вперед» (button[data-testid="next-pagination-button"]), а не подгрузкой
  // по скроллу — из-за этого в выгрузке стабильно было ровно 20 отелей.
  const bag = new Map();
  for (let p = 1; p <= OSTROVOK_MAX_PAGES; p++) {
    await scrollPage(page, 8, 900);
    const batch = await page.evaluate(ostrovokExtract).catch(() => []);
    for (const c of batch) if (!bag.has(c.name)) bag.set(c.name, c);
    console.log(`   стр. ${p}: +${batch.length}, всего ${bag.size}`);
    if (p >= OSTROVOK_MAX_PAGES) {
      console.log(`   ⏹  упёрлись в предохранитель --ostrovok-pages=${OSTROVOK_MAX_PAGES},`
        + ' выдача может быть неполной');
      break;
    }
    if (!await ostrovokNextPage(page)) { console.log('   ⏹  страницы кончились'); break; }
  }

  await page.screenshot({ path: path.join(outDir, 'ostrovok_screenshot.png') });
  console.log(`Текущий URL: ${page.url()}`);
  const raw = [...bag.values()];

  function ostrovokExtract() {
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

      // Price — RUB ₽. Приоритет: итоговая цена «N ₽ за M ночей» (а не GURU-цена
      // по логину и не зачёркнутая «до скидки»). Иначе — первое ₽-число.
      const totalMatch = allText.match(/([\d][\d\s]{2,}[\d])\s*₽\s*за\s+\d+\s+ноч/i);
      const priceMatch = totalMatch ||
                         allText.match(/([\d][\d\s]*[\d])\s*[₽Р]/) ||
                         allText.match(/([\d]{4,})\s*[₽Р]/);
      const priceDisplay = priceMatch
        ? (totalMatch ? totalMatch[1].trim() + ' ₽' : priceMatch[0].trim())
        : null;

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
  }

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

// ── AGODA.COM ─────────────────────────────────────────────────────────

// Селекторы сняты с живой выдачи 13.08.2026, не угаданы.
//
// ГЛАВНОЕ: цену берём ТОЛЬКО из отдельного числового узла display-price.
// Регэксп по тексту карточки брал первое ₽-число, а первым в блоке
// property-card-price идёт бейдж «С учетом 2 039 ₽» — это РАЗМЕР СКИДКИ
// (consolidated-applied-discount-badge), а не цена. Из-за этого в выгрузку
// попадали 2 039 ₽ за ultra-all-inclusive вместо настоящих 22 748 ₽.
// Второй кандидат из старого кода, fpc-cor-price, — это зачёркнутая цена
// ДО скидки, тоже не то.
const AGODA_CARD = '[data-selenium="hotel-item"]';

// Исполняется в странице. Не должна ссылаться на внешние переменные.
function agodaExtract() {
  const num = s => {
    if (!s) return null;
    const m = String(s).replace(/[    ]/g, ' ').match(/(\d[\d\s]*)/);
    if (!m) return null;
    const v = parseInt(m[1].replace(/\s/g, ''), 10);
    return isNaN(v) ? null : v;
  };

  return [...document.querySelectorAll('[data-selenium="hotel-item"]')].map(n => {
    const q = s => n.querySelector(s);
    const t = s => q(s)?.textContent?.trim() || null;

    const name = t('[data-selenium="hotel-name"]');
    if (!name) return null;

    // Agoda метит узлы то data-selenium, то data-element-name, то data-testid —
    // причём один и тот же по смыслу блок в разных карточках по-разному, а
    // у одного и того же узла семейство атрибута со временем меняется.
    // Ищем по значению во всех трёх: иначе половина полей молча пустая
    // (так рейтинг, звёзды и отзывы приходили null у всех 575 отелей).
    const byName = v => n.querySelector(`[data-selenium="${v}"]`)
                     || n.querySelector(`[data-element-name="${v}"]`)
                     || n.querySelector(`[data-testid="${v}"]`);
    const tn = v => byName(v)?.textContent?.trim() || null;

    // Цена — только именованный узел; регэксп по тексту карточки брать
    // НЕЛЬЗЯ (первым в блоке идёт бейдж «С учетом 2 751 ₽» — это размер
    // скидки, а fpc-cor-price — зачёркнутая цена до скидки).
    // 15.08.2026 Agoda перенесла цену из data-element-name в data-selenium:
    // display-price, hotel-currency и per-night-begins-with под старым
    // атрибутом не находятся уже ни в одной карточке. Цена держалась на
    // запасном final-price, а валюта и подпись периода молча стали null —
    // ровно тот тихий отказ, ради которого byName и заведён. Число берём
    // как есть: узел отдаёт то «19 850», то «₽21 189».
    let price = num(tn('display-price'));
    if (price === null) price = num(tn('final-price'));
    if (price === null) {
      price = num(byName('fpc-room-price')?.getAttribute('data-fpc-value'));
    }

    // Подпись сама говорит, за что цена («За ночь после налогов и сборов»).
    // Читаем её, а не предполагаем период.
    const priceLabel = tn('per-night-begins-with');

    // «Количество звезд: 5 из 5»
    const starM = (tn('ssr-property-card-star-rating') || '')
      .match(/(\d[,.]?\d?)\s*из\s*5/);

    // «Средняя оценка Отлично 8,7 из 10 по 3 555 отзывам»
    const revTxt = tn('property-card-review') || tn('ReviewWithDemographic') || '';
    const rateM = revTxt.match(/(\d[,.]\d)\s*из\s*10/);
    const cntM  = revTxt.match(/по\s*([\d\s  ]+?)\s*отзыв/);

    // «Восточный Кам Хай, Нячанг - 20,8 км от центра»
    const area = tn('area-city-text') || tn('area-city');
    const distM = (area || '').match(/(\d+[,.]?\d*)\s*км от центра/);

    return {
      hotel_id: n.getAttribute('data-hotelid') || null,
      name,
      price,
      currency: tn('hotel-currency'),
      price_label: priceLabel,
      stars:   starM ? parseFloat(starM[1].replace(',', '.')) : null,
      rating:  rateM ? parseFloat(rateM[1].replace(',', '.')) : null,
      reviews: cntM ? num(cntM[1]) : null,
      area,
      dist_km: distM ? parseFloat(distM[1].replace(',', '.')) : null,
      is_ad: !!byName('sponsored-badge'),
      url: q('a')?.href || null,
      thumbnail: q('img')?.src || null,
    };
  }).filter(Boolean);
}

async function scrapeAgoda(page) {
  const { взрослых: adults, номеров: rooms } = config.гости;

  // Agoda не принимает поисковый URL с текстовым городом: без внутреннего
  // cityId она сбрасывает на главную. Поэтому ведём себя как человек —
  // печатаем город в поле, берём первую подсказку и жмём поиск.
  console.log('\n╔══════════════════════════════════════════╗');
  console.log('║  AGODA.COM                               ║');
  console.log('╚══════════════════════════════════════════╝\n');

  // Прямой URL с внутренним id города: календарь Agoda кликать не нужно,
  // даты она принимает параметрами. Id берётся из URL после ручного поиска.
  if (config.город.cityid_agoda) {
    const url = [
      'https://www.agoda.com/ru-ru/search',
      `?city=${config.город.cityid_agoda}`,
      `&checkIn=${checkin}&checkOut=${checkout}&los=${nights}`,
      `&adults=${adults}&rooms=${rooms}`,
      '&currency=RUB&locale=ru-ru',
    ].join('');
    console.log(`URL: ${url}\n⏳ Открываю, жду 30 секунд...\n`);
    await page.goto(url, { waitUntil: 'domcontentloaded', timeout: 90000 });
    await wait(30000);
    await logLoginState(page, 'Agoda.com');
    return await harvestAgoda(page);
  }

  await page.goto('https://www.agoda.com/ru-ru/', {
    waitUntil: 'domcontentloaded', timeout: 90000,
  });
  await wait(6000);

  for (const label of ['Разрешить все', 'Принять все', 'ОК', 'Accept All']) {
    const b = page.locator(`button:has-text("${label}")`).first();
    if (await b.isVisible({ timeout: 1500 }).catch(() => false)) {
      await b.click().catch(() => {});
      await wait(600);
      break;
    }
  }

  const box = page.locator('#textInput, input[data-selenium="textInput"]').first();
  await box.click({ timeout: 20000 });
  await box.fill(config.город.query_trip || config.город.название);
  await wait(3000);
  await page.keyboard.press('ArrowDown').catch(() => {});
  await page.keyboard.press('Enter').catch(() => {});
  await wait(2500);

  // Календарь: кликаем нужные дни по data-selenium-date="YYYY-MM-DD"
  for (const d of [checkin, checkout]) {
    const cell = page.locator(`[data-selenium-date="${d}"]`).first();
    if (await cell.isVisible({ timeout: 8000 }).catch(() => false)) {
      await cell.click().catch(() => {});
      await wait(1200);
    } else {
      console.log(`   ⚠️  день ${d} в календаре не найден`);
    }
  }

  const search = page.locator('button[data-selenium="searchButton"], '
    + 'button:has-text("Найти отели"), button:has-text("Поиск")').first();
  await search.click({ timeout: 15000 }).catch(() => {});
  console.log('⏳ Жду выдачу, 30 секунд...\n');
  await wait(30000);

  return await harvestAgoda(page);
}

// Собираем ПО ХОДУ прокрутки, а не один раз в конце: Agoda дорисовывает и
// карточки, и цены только по факту попадания в зону видимости.
//
// Списком правит react-lazyload. Очередная порция карточек висит на
// «сентинеле» — пустом div[data-testid="lazy-load-component"] НУЛЕВОЙ
// высоты. Порция монтируется, только когда сентинел побывал в зоне
// видимости, и вешает следующий сентинел ниже себя. Отсюда три правила,
// каждое снято с живой выдачи Нячанга 15.08.2026, а не выведено умозрительно.
//
// 1. Шаг прокрутки ОБЯЗАН быть меньше экрана. Прежние 1,5 экрана оставляли
//    между двумя кадрами слепую полосу в пол-экрана. Сентинел нулевой
//    высоты, попавший в неё, не пересекается с вьюпортом никогда — цепочка
//    не стартует, и страница навсегда остаётся с теми 11 карточками, что
//    пришли с сервера. Замер: шаг 1,5 экрана — 11 карточек и lazy 0/8 на
//    каждом шаге; шаг 0,8 экрана — первый сентинел срабатывает на первом же
//    шаге, дальше 45 карточек, дальше 90. Отсюда и брались 92 отеля вместо
//    368: 11 страниц по 8-9 карточек.
// 2. Стоянием на месте не догрузить ничего. 45 секунд внизу страницы не
//    добавили ни одной карточки: список растёт только под прокруткой.
// 3. Заканчивать прыжком scrollTo(scrollHeight) нельзя — это ровно тот
//    слепой прыжок, который сбор и ломает.
//
// Останавливаемся не по числу шагов, а по факту: доехали до низа И ни
// карточек, ни высоты страницы не прибавилось несколько шагов подряд.
// Высоту смотрим наравне с карточками, потому что порция сначала занимает
// место и только потом наполняется ценами.
async function agodaHarvestVisible(page, bag) {
  const absorb = list => {
    for (const c of list) {
      const key = c.hotel_id || c.name;
      const prev = bag.get(key);
      // Карточка с ценой всегда лучше той же карточки без цены
      if (!prev || (prev.price == null && c.price != null)) bag.set(key, c);
    }
  };

  // Страница после клика по пагинации и так приходит наверх, но начинать
  // проход из известной точки дешевле, чем полагаться на это.
  await page.evaluate(() => window.scrollTo(0, 0)).catch(() => {});
  await wait(1200);

  const t0 = Date.now();
  let stale = 0, prevSize = -1, prevHeight = -1, steps = 0, onPage = 0;
  let stop = 'страница перестала расти';

  for (; steps < AGODA_MAX_STEPS; steps++) {
    const batch = await page.evaluate(agodaExtract).catch(() => []);
    onPage = Math.max(onPage, batch.length);
    absorb(batch);

    const geo = await page.evaluate(() => ({
      height: document.body.scrollHeight,
      y: Math.round(window.scrollY),
      view: window.innerHeight,
    })).catch(() => null);
    if (!geo) { stop = 'страница не ответила'; break; }

    const atBottom = geo.y + geo.view >= geo.height - 8;
    if (bag.size !== prevSize || geo.height !== prevHeight) {
      stale = 0; prevSize = bag.size; prevHeight = geo.height;
    } else if (atBottom) {
      stale++;                    // и не растём, и ехать дальше некуда
    }
    if (stale >= AGODA_STALE_STOP) break;

    if (Date.now() - t0 > AGODA_PAGE_BUDGET_MS) {
      stop = `бюджет ${Math.round(AGODA_PAGE_BUDGET_MS / 1000)} с исчерпан`;
      break;
    }

    await page.evaluate(r => window.scrollBy(0, Math.round(window.innerHeight * r)),
      AGODA_STEP_RATIO);
    await wait(AGODA_SCROLL_PAUSE);
  }
  if (steps >= AGODA_MAX_STEPS) stop = `предохранитель ${AGODA_MAX_STEPS} шагов`;

  // Последние карточки монтируются раньше, чем в них приезжает цена, а
  // цена уже не двигает ни счётчик, ни высоту и потому не сбрасывает
  // счётчик простоя. Один добор после выхода из цикла их и подбирает.
  const tail = await page.evaluate(agodaExtract).catch(() => []);
  onPage = Math.max(onPage, tail.length);
  absorb(tail);

  return { steps, sec: Math.round((Date.now() - t0) / 1000), height: prevHeight, stop, onPage };
}

// Выдача разбита на страницы («Стр. 1 из 10») — скроллом они не добираются
async function agodaNextPage(page) {
  const btn = page.locator('[data-selenium="pagination-next-btn"]').first();
  if (!await btn.isVisible({ timeout: 5000 }).catch(() => false)) return false;
  if (await btn.isDisabled().catch(() => false)) return false;

  const before = await page.locator(AGODA_CARD).first()
    .getAttribute('data-hotelid').catch(() => null);
  await btn.click().catch(() => {});

  for (let i = 0; i < 30; i++) {
    await wait(1000);
    const now = await page.locator(AGODA_CARD).first()
      .getAttribute('data-hotelid').catch(() => null);
    if (now && now !== before) { await wait(4000); return true; }
  }
  return false;
}

async function harvestAgoda(page) {
  const bag = new Map();

  // Сколько всего страниц — спрашиваем у самой Agoda («Стр. 1 из 10»),
  // а не зашиваем числом: для другого города оно другое.
  let limit = AGODA_MAX_PAGES || 40;

  for (let p = 1; p <= limit; p++) {
    const label = await page.locator('[data-selenium="pagination-text"]').first()
      .textContent({ timeout: 5000 }).catch(() => null);

    if (p === 1 && !AGODA_MAX_PAGES) {
      const m = (label || '').match(/из\s*(\d+)/i);
      if (m) {
        limit = parseInt(m[1], 10);
        console.log(`📄 Agoda сообщает: всего страниц ${limit}`);
      }
    }
    console.log(`📄 Agoda, страница ${p}${label ? ` (${label.trim()})` : ''}...`);

    const before = bag.size;
    const r = await agodaHarvestVisible(page, bag);
    const got = bag.size - before;
    const withPrice = [...bag.values()].filter(c => c.price != null).length;
    console.log(`   на странице ${r.onPage}, из них новых ${got}`
      + ` — ${r.steps} шагов / ${r.sec} с, высота ${r.height} px, стоп: ${r.stop}`);
    console.log(`   собрано ${bag.size}, из них с ценой ${withPrice}`);

    // Тихий недобор — та самая поломка, из-за которой выгрузка съехала с 368
    // до 92 и никто этого не заметил: страниц столько же, а карточек втрое
    // меньше. Пусть кричит в лог, а не выясняется через месяц по отчёту.
    // Считаем именно карточки НА СТРАНИЦЕ, а не новые: на последних страницах
    // Agoda повторяет уже показанное, и по приросту это неотличимо от поломки
    // (страница 11 Нячанга: 6 новых при полусотне карточек в DOM).
    if (r.onPage < 15) {
      console.log(`   ⚠️  на странице всего ${r.onPage} карточек — похоже, ленивая`
        + ` подгрузка не стартовала (ждём десятки, а не единицы)`);
    }

    if (p >= limit) {
      console.log(AGODA_MAX_PAGES
        ? `   ⏹  стоп: лимит --agoda-pages=${AGODA_MAX_PAGES}`
        : '   ⏹  прошли все страницы, что заявила Agoda');
      break;
    }
    if (!await agodaNextPage(page)) { console.log('   ⏹  страницы кончились'); break; }
  }

  console.log(`\n✅ Agoda: карточек ${bag.size}`);

  let saved = 0, noPrice = 0, badCurrency = 0;
  for (const c of bag.values()) {
    if (c.price == null) { noPrice++; continue; }
    // Просим currency=RUB. Если пришло иное — лучше пропустить, чем
    // записать донги в поле с рублями.
    if (c.currency && !/[₽Р]|RUB|руб/i.test(c.currency)) { badCurrency++; continue; }
    if (c.price < 100 || c.price > 5000000) { noPrice++; continue; }

    // «За ночь после налогов и сборов» → covers = 1. Если Agoda когда-нибудь
    // отдаст «за 3 ночи» — поделим, а не умножим вслепую на nights.
    const m = (c.price_label || '').match(/за\s+(\d+)\s+ноч/i);
    const covers = m ? parseInt(m[1], 10) : 1;
    const perNight = Math.round(c.price / covers);

    saveHotel({
      source: 'agoda',
      name: c.name,
      stars: c.stars,
      rating: c.rating,
      rating_label: null,
      review_count: c.reviews,
      price_display: `${c.price.toLocaleString('ru-RU')} ${c.currency || '₽'}`,
      price_total_rub: perNight * nights,
      price_per_night_rub: perNight,
      nights,
      address: c.area,
      distance_text: c.dist_km != null ? `${c.dist_km} км от центра` : null,
      distance_center_km: c.dist_km,
      is_ad: c.is_ad,
      lat: null,
      lng: null,
      url: c.url,
      thumbnail: c.thumbnail,
    });
    saved++;
  }

  console.log(`   сохранено ${saved}, без цены ${noPrice}`
    + (badCurrency ? `, чужая валюта ${badCurrency}` : ''));
  if (!saved) console.log('   ⚠️  цен не распознано — вёрстка могла поменяться');
}

// ── Google Hotels ─────────────────────────────────────────────────────
//
// Пятый источник, и самый полезный: Google - метапоиск, он показывает цены
// тех агрегаторов, которых у нас нет. Скидка на Potique в 20% нашлась именно
// через него, у Traveloka, а Traveloka не парсит никто из остальных четырёх.
//
// Три особенности, из-за которых модуль не похож на соседние.
//
// 1. Даты. Google игнорирует checkin/checkout в ссылке - он держит их
//    в параметре ts, это protobuf в base64url. Собирать его с нуля нельзя,
//    сервер отвечает 500. Рабочий приём: взять живой шаблон и подменить
//    в нём байты дней и число ночей. Длины при этом не меняются, поэтому
//    внутренние префиксы длин остаются валидными.
// 2. Постраничность есть, но невидимая. Внизу списка написано «Результаты
//    1-18 из 1091» и рядом стоит кнопка «Далее» БЕЗ подписи: текст лежит
//    внутри span, а сама кнопка неотличима от стрелок каруселей внутри
//    карточек - у тех такой же класс и aria-label «Далее». Отличается она
//    тем, что у неё непустой innerText. Прокрутка колесом список не растит,
//    поэтому раньше источник и упирался в 17 карточек.
// 3. «Все варианты» кликом не раскрываются. Кнопка «Показать больше
//    вариантов» лежит в контейнере нулевой высоты, Playwright до неё не
//    доскроллит, а el.click() из evaluate не даёт ни одного запроса. Клик
//    и не нужен: весь список партнёров приезжает с первым же GET внутри
//    блока AF_initDataCallback({key: 'ds:1', ...}) - это тот же массив,
//    из которого страница рисуется. Разбираем его напрямую, JS исполнять
//    не требуется, поэтому карточку отеля берём обычным ctx.request.get
//    за полторы секунды вместо пяти на рендер.

const GOOGLE_TS_TEMPLATE =
  'CAEaOAoaEhYKCS9tLzA0NGNqdjoJTmhhIFRyYW5nGgASGhIUCgcI6g8QCBgREgcI6g8QCBgVGAQyAggBKgkKBToDUlVCGgA';
// в шаблоне зашиты заезд 17 августа, выезд 21 августа и 4 ночи
const GOOGLE_DATE_MARK = Buffer.from('08ea0f100818', 'hex');   // {2026, 8, день}

function googleTs(ci, co) {
  const raw = Buffer.from(GOOGLE_TS_TEMPLATE.replace(/-/g, '+').replace(/_/g, '/'), 'base64');
  const spots = [];
  let from = 0;
  for (;;) {
    const i = raw.indexOf(GOOGLE_DATE_MARK, from);
    if (i === -1) break;
    spots.push(i + GOOGLE_DATE_MARK.length);
    from = i + 1;
  }
  if (spots.length !== 2) throw new Error(`в шаблоне ts найдено дат: ${spots.length}`);
  const d1 = new Date(ci), d2 = new Date(co);
  raw[spots[0]] = d1.getUTCDate();
  raw[spots[1]] = d2.getUTCDate();
  if (raw[spots[1] + 1] === 0x18) {
    raw[spots[1] + 2] = Math.round((d2 - d1) / 86400000);
  }
  return raw.toString('base64').replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
}

function googleExtract() {
  const num = s => {
    if (!s) return null;
    const m = String(s).replace(/[    ]/g, ' ').match(/(\d[\d\s]*)/);
    if (!m) return null;
    const v = parseInt(m[1].replace(/\s/g, ''), 10);
    return isNaN(v) ? null : v;
  };

  // .Zvwhrc - контейнер карточки. Класс генерируется сборкой Google и может
  // смениться, поэтому есть запасной путь: карточкой считается узел, где
  // одновременно есть заголовок, цена и ссылка на выдачу.
  let cards = [...document.querySelectorAll('.Zvwhrc')];
  if (!cards.length) {
    cards = [...document.querySelectorAll('div')].filter(d =>
      d.querySelector('h2') && /₽/.test(d.innerText || '') &&
      d.querySelector('a[href*="/travel/"]') && (d.innerText || '').length < 900);
  }

  return cards.map(n => {
    const text = (n.innerText || '').replace(/[    ]/g, ' ');
    const name = n.querySelector('h2, [role="heading"]')?.textContent?.trim() || null;
    if (!name) return null;

    // Цена за ночь и итог за все ночи лежат в подписи ссылки:
    // «2 426 ₽Итоговая цена: 7 278 ₽За 3 ночи»
    const priceLink = [...n.querySelectorAll('a[href], button')]
      .map(a => (a.innerText || '') + ' ' + (a.getAttribute('aria-label') || ''))
      .find(s => /₽/.test(s)) || text;
    const perNight = num((priceLink.match(/([\d\s]{3,})\s*₽/) || [])[1]);
    const total = num((priceLink.match(/Итоговая цена:\s*([\d\s]+)\s*₽/) || [])[1]);

    // Рейтинг у Google по пятибалльной шкале: «4,6 (843)»
    const rm = text.match(/\b([1-5],\d)\b\s*\n?\((\d[\d\s]*)\)/);
    const stars = num((text.match(/(\d)\s*звезд/) || [])[1]);

    const link = n.querySelector('a[href*="/travel/"]');
    const img = n.querySelector('img');

    // Идентификатор отеля у Google лежит в data-href ссылки карточки:
    // data-href="/entity/ChkIm5GK1o2cvfVGGg0vZy8xMWZocXI0cXk4EAE". По нему
    // открывается страница отеля со всеми предложениями партнёров, и это
    // единственная ниточка от списка к ценам - в самой карточке их нет.
    const ent = [...n.querySelectorAll('[data-href]')]
      .map(a => (a.getAttribute('data-href') || '').match(/\/entity\/([A-Za-z0-9_-]+)/))
      .find(Boolean);

    return {
      name,
      entity_id: ent ? ent[1] : null,
      per_night: perNight,
      total,
      rating5: rm ? parseFloat(rm[1].replace(',', '.')) : null,
      reviews: rm ? num(rm[2]) : null,
      stars,
      deal: /СКИДКА/i.test(text),
      // «Отель», «Апартаменты», «Хостел» - Google подписывает тип
      kind: (text.match(/\b(Отель|Апартаменты|Хостел|Мотель|Курорт|Вилла)\b/) || [])[1] || null,
      url: link ? link.href : null,
      thumbnail: img ? img.src : null,
    };
  }).filter(Boolean);
}

async function googleHarvest(page, bag) {
  const batch = await page.evaluate(googleExtract).catch(() => []);
  let added = 0;
  for (const c of batch) {
    // Ключ - идентификатор, а не название: у Google в одном городе полно
    // одноимённых апартаментов, и по имени они схлопывались в одну запись.
    const key = c.entity_id || `имя:${c.name}`;
    if (!bag.has(key)) { bag.set(key, c); added++; }
  }
  return { seen: batch.length, added };
}

// Счётчик «Результаты 19-38 из 2 760» - единственный надёжный признак того,
// что страница действительно перелистнулась: DOM Google перерисовывает
// асинхронно, и карточки какое-то время остаются от прошлой страницы.
async function googleCounter(page) {
  return page.evaluate(() => {
    const n = [...document.querySelectorAll('*')]
      .find(e => !e.children.length && /Результаты\s*[\d\s]+[–-][\d\s]+\s*из/.test(e.textContent || ''));
    return n ? n.textContent.replace(/\s+/g, ' ').trim() : null;
  }).catch(() => null);
}

// Порядок сортировки. Нужен не ради порядка: постраничность у Google
// обрывается примерно на 690 карточках, хотя счётчик обещает 2750. Каждая
// сортировка показывает свои 690 из той же тысячи, и объединение проходов
// достаёт то, до чего один проход не доходит. Первый проход - как открылось.
const GOOGLE_SORTS = [
  ['по релевантности', null],
  ['по цене',          'По цене (в порядке возрастания)'],
  ['по оценке',        'По оценке пользователей (в порядке убывания)'],
  ['по отзывам',       'Наибольшее количество отзывов'],
];
const gSortsArg = args.find(a => a.startsWith('--google-sorts='));
const GOOGLE_SORT_COUNT = gSortsArg
  ? Math.max(1, Math.min(GOOGLE_SORTS.length, parseInt(gSortsArg.split('=')[1], 10) || 1))
  : GOOGLE_SORTS.length;

async function googleSortBy(page, label) {
  const opener = page.locator('button', { hasText: 'Сортировка результатов' }).first();
  if (!await opener.count()) return false;
  await opener.click({ timeout: 15000 });
  await wait(700);
  const item = page.locator(`text="${label}"`).last();
  if (!await item.count()) return false;
  await item.click({ timeout: 15000 });
  await wait(3000);
  return true;
}

// Кнопка постраничности. Отличаем её от стрелок каруселей внутри карточек
// (у тех тот же класс и тот же aria-label) по непустому тексту.
async function googleClickNext(page) {
  const btn = page.locator('button', { hasText: /^Далее$/ }).last();
  if (!await btn.count()) return false;
  if (await btn.isDisabled().catch(() => false)) return false;
  await btn.click({ timeout: 15000 });
  return true;
}

// ── Страница отеля: блок «Все варианты» ───────────────────────────────

// Блоки AF_initDataCallback({key: 'ds:N', ..., data: [...], sideChannel: ...})
// - это состояние страницы, отданное сервером вместе с HTML.
function googleDsBlocks(html) {
  const out = {};
  const re = /AF_initDataCallback\(\{key: '(ds:\d+)'.*?data:(\[.*?\]), sideChannel/gs;
  let m;
  while ((m = re.exec(html))) {
    try { out[m[1]] = JSON.parse(m[2]); } catch { /* блок не наш, пропускаем */ }
  }
  return out;
}

// Ищем узлы по форме, а не по фиксированному пути [0,6,2,21]: индексы
// у Google меняются от сборки к сборке, а форма - нет. Предложение партнёра
// узнаётся по тройке [название, id, "/travel/lodging/clk?..."] в голове.
function googleFindOffers(root) {
  let best = null;
  const isOffer = n => Array.isArray(n) && Array.isArray(n[0])
    && typeof n[0][0] === 'string'
    && typeof n[0][2] === 'string' && n[0][2].includes('/travel/lodging/clk');
  const walk = node => {
    if (!Array.isArray(node)) return;
    if (node.length && node.every(isOffer)) {
      if (!best || node.length > best.length) best = node;
      return;                       // внутрь найденного списка не лезем
    }
    for (const v of node) walk(v);
  };
  walk(root);
  return best || [];
}

// Цена внутри предложения лежит пятёркой ["7 833 ₽", null, 7833.25, null, 7833].
// Их две: за ночь и за всё проживание. Меньшая - за ночь.
function googlePriceTuples(node, acc = []) {
  if (!Array.isArray(node)) return acc;
  if (typeof node[0] === 'string' && /\d\s*₽/.test(node[0]) && typeof node[2] === 'number') {
    acc.push(Math.round(node[2]));
    return acc;
  }
  for (const v of node) googlePriceTuples(v, acc);
  return acc;
}

function googleParseEntity(html) {
  const ds = googleDsBlocks(html);
  // Блок отеля узнаём по форме: [0][1] - название, [0][2][0] - координаты.
  // Предложений партнёров в нём может не быть совсем: на эти даты отель
  // никто не продаёт. Это нормальный ответ, а не сбой разбора, поэтому
  // выбираем блок по данным отеля, а не по наличию предложений - иначе
  // теряются координаты и рейтинг, а с ними и отсев по расстоянию.
  let data = null, offers = [];
  for (const v of Object.values(ds)) {
    if (typeof v?.[0]?.[1] !== 'string' || !Array.isArray(v?.[0]?.[2]?.[0])) continue;
    const o = googleFindOffers(v);
    if (!data || o.length > offers.length) { offers = o; data = v; }
  }
  if (!data) return null;

  const flat = JSON.stringify(data);
  const at = p => { let n = data; for (const i of p) { if (n == null) return null; n = n[i]; } return n; };

  const parsed = offers.map(o => {
    const prices = googlePriceTuples(o).filter(v => v >= 50 && v <= 5000000);
    if (!prices.length) return null;
    const perNight = Math.min(...prices);
    const total = Math.max(...prices);
    const clk = o[0][2] || '';
    // В ссылке-редиректе настоящий адрес партнёра лежит в pcurl=
    const pc = clk.match(/[?&]pcurl=([^&]+)/);
    let partnerUrl = null;
    if (pc) { try { partnerUrl = decodeURIComponent(pc[1]); } catch { partnerUrl = pc[1]; } }
    return {
      // Первой строкой Google ставит бронирование напрямую, и подписывает
      // его названием отеля, а не словами. В таблице это выглядит как ещё
      // один агрегатор, поэтому переименовываем.
      partner: o[0][0] === at([0, 1]) ? 'Официальный сайт' : o[0][0],
      per_night_rub: perNight,
      total_rub: total > perNight ? total : perNight * nights,
      url: partnerUrl || ('https://www.google.com' + clk),
    };
  }).filter(Boolean);

  const rating = at([0, 7, 0]);
  const coords = at([0, 2, 0]);
  const stars = (flat.match(/"(\d) звезд[аы]?",\s*(\d)/) || [])[1];
  return {
    name: at([0, 1]),
    lat: Array.isArray(coords) ? coords[0] : null,
    lng: Array.isArray(coords) ? coords[1] : null,
    address: at([0, 2, 1, 0, 0, 0]),
    rating5: Array.isArray(rating) && typeof rating[0] === 'number' ? rating[0] : null,
    reviews: Array.isArray(rating) && typeof rating[1] === 'number' ? rating[1] : null,
    stars: stars ? parseInt(stars, 10) : null,
    offers: parsed,
  };
}

// Расстояние до центра города по прямой: у карточек Google его нет,
// а планка «не дальше 8 км» без него не работает - половина «нячангских»
// отелей стоит в Камрани за 20-35 км.
function kmFromCenter(lat, lng) {
  if (lat == null || lng == null || config.город.lat == null) return null;
  const R = 6371, rad = d => d * Math.PI / 180;
  const dLat = rad(lat - config.город.lat), dLng = rad(lng - config.город.lng);
  const a = Math.sin(dLat / 2) ** 2
    + Math.cos(rad(config.город.lat)) * Math.cos(rad(lat)) * Math.sin(dLng / 2) ** 2;
  return Math.round(2 * R * Math.asin(Math.sqrt(a)) * 10) / 10;
}

// Название для сравнения между источниками. Слова, которые ничего не значат
// («отель», «nha trang», «by», «the»), выбрасываем; «resort» и «condotel»
// оставляем - они различают объекты: «Diamond Bay Hotel» в городе и
// «Diamond Bay Resort & Spa» в Камрани это разные гостиницы.
function normName(s) {
  return (s || '').toLowerCase()
    .replace(/\b(отель|hotel|spa|nha|trang|by|the|and|апарт|apartment)\b/g, ' ')
    .replace(/[^a-zа-я0-9]+/g, ' ').trim().replace(/\s+/g, ' ');
}

// Поиск отеля по названию. Даты тут не нужны и мешают: с параметром ts,
// в котором зашит город, Google на запрос с именем отеля отвечает «Ничего
// не найдено». Идентификатор отеля от дат не зависит, поэтому ищем без ts,
// а цены потом берём страницей отеля уже с нашими датами.
async function googleFindEntity(ctx, name) {
  const q = `${name} ${config.город.query_trip}`;
  const res = await ctx.request.get(
    `https://www.google.com/travel/search?q=${encodeURIComponent(q)}&hl=ru&gl=ru&curr=RUB`,
    { headers: { 'accept-language': 'ru-RU,ru;q=0.9' }, timeout: 45000 });
  if (!res.ok()) throw new Error(`HTTP ${res.status()}`);
  const html = await res.text();
  // Первая ссылка не всегда верхний результат: рядом с карточкой Google
  // рисует «похожие отели» с такими же data-href. Отдаём первые несколько,
  // а выбирает вызывающий - по совпадению названия.
  return [...new Set([...html.matchAll(/data-href="\/entity\/([A-Za-z0-9_-]+)"/g)]
    .map(m => m[1]))].slice(0, 3);
}

async function googleFetchEntity(ctx, entityId, ts) {
  const url = `https://www.google.com/travel/hotels/entity/${entityId}/prices`
    + `?hl=ru&gl=ru&curr=RUB&ts=${ts}`;
  const res = await ctx.request.get(url, {
    headers: { 'accept-language': 'ru-RU,ru;q=0.9' }, timeout: 45000,
  });
  if (!res.ok()) throw new Error(`HTTP ${res.status()}`);
  return googleParseEntity(await res.text());
}

async function scrapeGoogle(page) {
  const ts = googleTs(checkin, checkout);
  const url = 'https://www.google.com/travel/search?'
    + `q=${encodeURIComponent('hotels ' + config.город.query_trip)}`
    + `&hl=ru&gl=ru&curr=RUB&ts=${ts}`;
  console.log('\n🌐  Google Hotels');
  console.log(`   ${url.slice(0, 96)}...`);
  await page.goto(url, { waitUntil: 'domcontentloaded', timeout: 60000 });
  await wait(3500);

  // Проверяем, что даты доехали: иначе цены будут не за наши ночи
  const shown = await page.evaluate(() => {
    const v = a => {
      const el = [...document.querySelectorAll('input')]
        .find(i => i.getAttribute('aria-label') === a);
      return el ? el.value : null;
    };
    return { in: v('Заезд'), out: v('Выезд') };
  }).catch(() => ({}));
  console.log(`   даты на странице: ${shown.in} → ${shown.out}`);

  // ── Проход 1: список, постранично, по каждой сортировке ──
  const bag = new Map();
  const t0 = Date.now();
  for (const [sortName, sortLabel] of GOOGLE_SORTS.slice(0, GOOGLE_SORT_COUNT)) {
    const wasSize = bag.size;
    console.log(`\n   ── сортировка ${sortName} ──`);
    if (sortLabel) {
      await page.goto(url, { waitUntil: 'domcontentloaded', timeout: 60000 });
      await wait(3500);
      let switched = false;
      try { switched = await googleSortBy(page, sortLabel); }
      catch (e) { console.log(`   ⚠️  ${e.message.split('\n')[0]}`); }
      if (!switched) { console.log('   ⚠️  переключить сортировку не удалось, проход пропущен'); continue; }
    }

    let stale = 0, prevCounter = null, page_ = 0;
    for (page_ = 1; page_ <= GOOGLE_MAX_PAGES; page_++) {
      const { seen, added } = await googleHarvest(page, bag);
      const counter = await googleCounter(page);
      console.log(`   стр. ${String(page_).padStart(3)}: на странице ${seen}, новых ${added},`
        + ` всего ${bag.size}${counter ? ` (${counter})` : ''}`);

      // Предохранитель: счётчик стоит и новых карточек нет - дальше некуда
      if (!added && counter === prevCounter) {
        if (++stale >= 2) { console.log('   ⏹  список перестал расти'); break; }
      } else stale = 0;
      prevCounter = counter;

      let clicked = false;
      try { clicked = await googleClickNext(page); }
      catch (e) { console.log(`   ⏹  «Далее» не нажалась: ${e.message.split('\n')[0]}`); }
      if (!clicked) { console.log('   ⏹  кнопки «Далее» больше нет — это последняя страница'); break; }

      // Ждём, пока счётчик сменится: карточки перерисовываются с задержкой
      // и без этой паузы следующий проход собрал бы ту же страницу.
      for (let w = 0; w < 20; w++) {
        await wait(400);
        if (await googleCounter(page) !== counter) break;
      }
    }
    if (page_ > GOOGLE_MAX_PAGES) {
      console.log(`   ⚠️  упёрлись в предохранитель --google-pages=${GOOGLE_MAX_PAGES}`);
    }
    console.log(`   сортировка ${sortName} добавила ${bag.size - wasSize}, всего ${bag.size}`);
  }
  console.log(`\n✅ Google, список: карточек ${bag.size} за ${Math.round((Date.now() - t0) / 1000)} с`);

  const cards = [...bag.values()];
  const withId = cards.filter(c => c.entity_id);
  console.log(`   с идентификатором отеля: ${withId.length}, без него: ${cards.length - withId.length}`);

  // ── Проход 2: страница каждого отеля, блок «Все варианты» ──
  const details = new Map();
  if (GOOGLE_OFFERS && withId.length) {
    const targets = GOOGLE_OFFERS_LIMIT ? withId.slice(0, GOOGLE_OFFERS_LIMIT) : withId;
    console.log(`\n🔎  Google, варианты партнёров: ${targets.length} отелей,`
      + ` по ${GOOGLE_CONCURRENCY} за раз`);
    const t1 = Date.now();
    let done = 0, failed = 0, offersTotal = 0;
    for (let i = 0; i < targets.length; i += GOOGLE_CONCURRENCY) {
      const chunk = targets.slice(i, i + GOOGLE_CONCURRENCY);
      await Promise.all(chunk.map(async c => {
        try {
          const d = await googleFetchEntity(page.context(), c.entity_id, ts);
          if (d) { details.set(c.entity_id, d); offersTotal += d.offers.length; }
        } catch (e) {
          failed++;
          if (failed <= 5) console.log(`   ⚠️  ${c.name.slice(0, 40)}: ${e.message.split('\n')[0]}`);
        }
        done++;
      }));
      if (done % 50 < GOOGLE_CONCURRENCY || done === targets.length) {
        const sec = (Date.now() - t1) / 1000;
        console.log(`   ${String(done).padStart(4)}/${targets.length}`
          + ` | предложений ${offersTotal} | ошибок ${failed}`
          + ` | ${Math.round(sec)} с, ${(done / Math.max(sec, 1)).toFixed(1)}/с`);
      }
    }
    console.log(`   готово: ${details.size} страниц, ${offersTotal} предложений,`
      + ` ошибок ${failed}, ${Math.round((Date.now() - t1) / 1000)} с`);
  }

  // ── Проход 3: отели соседних источников, которых нет в списке Google ──
  if (GOOGLE_MATCH_FILE && GOOGLE_OFFERS && fs.existsSync(GOOGLE_MATCH_FILE)) {
    const txt = fs.readFileSync(GOOGLE_MATCH_FILE, 'utf8').trim();
    const prev = txt.startsWith('[') ? JSON.parse(txt)
      : txt.split('\n').filter(Boolean).map(l => JSON.parse(l));
    const known = new Set(cards.map(c => normName(c.name)));
    for (const d of details.values()) known.add(normName(d.name));
    const wanted = new Map();
    for (const h of prev) {
      if (h.source === 'google' || !h.name || !h.price_total_rub) continue;
      const k = normName(h.name);
      if (k && !known.has(k) && !wanted.has(k)) wanted.set(k, h.name);
    }
    console.log(`\n🔗  Google, поимённый добор: в ${GOOGLE_MATCH_FILE}`
      + ` отелей других источников без пары — ${wanted.size}`);

    const names = [...wanted.values()];
    const t2 = Date.now();
    let found = 0, mism = 0, missed = 0, failed2 = 0, done2 = 0;
    for (let i = 0; i < names.length; i += GOOGLE_CONCURRENCY) {
      await Promise.all(names.slice(i, i + GOOGLE_CONCURRENCY).map(async nm => {
        // Считаем в начале, а не в конце: ниже три ранних return, и после
        // них счётчик не доезжал. Прогресс печатался по первым удачным
        // находкам, а дальше проход шёл молча - со стороны неотличимо
        // от зависшего процесса.
        done2++;
        try {
          const ids = await googleFindEntity(page.context(), nm);
          if (!ids.length) { missed++; return; }
          const a = normName(nm);
          let hit = null;
          for (const id of ids) {
            if (details.has(id)) continue;          // уже взят под другим именем
            const d = await googleFetchEntity(page.context(), id, ts);
            if (!d) continue;
            // Google охотно отвечает «похожим» отелем, поэтому сверяем имя
            // из его же карточки: без этого в таблицу поедут чужие цены.
            const b = normName(d.name);
            if (a === b || (a.length >= 4 && b.includes(a)) || (b.length >= 4 && a.includes(b))) {
              hit = { id, d };
              break;
            }
          }
          if (!hit) { mism++; return; }
          details.set(hit.id, hit.d);
          cards.push({
            name: hit.d.name, entity_id: hit.id, per_night: null, total: null,
            rating5: null, reviews: null, stars: null, deal: false, kind: null,
            url: `https://www.google.com/travel/hotels/entity/${hit.id}/prices`
              + `?hl=ru&gl=ru&curr=RUB&ts=${ts}`,
            thumbnail: null,
          });
          found++;
        } catch (e) {
          failed2++;
          if (failed2 <= 5) console.log(`   ⚠️  ${nm.slice(0, 40)}: ${e.message.split('\n')[0]}`);
        }
      }));
      if (done2 % 50 < GOOGLE_CONCURRENCY || done2 === names.length) {
        console.log(`   ${String(done2).padStart(4)}/${names.length}`
          + ` | добавлено ${found} | чужой отель ${mism} | не нашлось ${missed}`
          + ` | ошибок ${failed2} | ${Math.round((Date.now() - t2) / 1000)} с`);
      }
    }
    console.log(`   добор закончен: +${found} отелей, отклонено по несовпадению имени ${mism},`
      + ` не нашлось ${missed}, ошибок ${failed2}`);
  }

  // ── Сохранение ──
  let saved = 0, noPrice = 0;
  for (const c of cards) {
    const d = c.entity_id ? details.get(c.entity_id) : null;
    const offers = d && d.offers.length ? d.offers : [];
    const bestOffer = offers.length
      ? offers.reduce((a, b) => (b.per_night_rub < a.per_night_rub ? b : a)) : null;

    let perNight = bestOffer ? bestOffer.per_night_rub
      : (c.per_night != null ? c.per_night
        : (c.total != null ? Math.round(c.total / nights) : null));
    if (perNight == null) { noPrice++; continue; }
    if (perNight < 100 || perNight > 5000000) { noPrice++; continue; }
    const total = bestOffer ? bestOffer.total_rub
      : (c.total != null ? c.total : perNight * nights);

    const rating5 = d && d.rating5 != null ? d.rating5 : c.rating5;
    const km = d ? kmFromCenter(d.lat, d.lng) : null;

    saveHotel({
      source: 'google',
      name: (d && d.name) || c.name,
      entity_id: c.entity_id,
      stars: (d && d.stars) || c.stars,
      // Приводим пятибалльную шкалу Google к десятибалльной, как у соседей,
      // иначе 4,6 встанет в общий рейтинг ниже любой тройки с Booking.
      rating: rating5 != null ? Math.round(rating5 * 2 * 100) / 100 : null,
      rating_label: rating5 != null ? `${rating5} из 5` : null,
      review_count: (d && d.reviews != null) ? d.reviews : c.reviews,
      price_display: `${perNight.toLocaleString('ru-RU')} ₽/ночь`,
      price_total_rub: total,
      price_per_night_rub: perNight,
      nights,
      address: (d && d.address) || c.kind,
      distance_text: km != null ? `${km} км от центра` : null,
      distance_center_km: km,
      is_deal: c.deal,
      lat: d ? d.lat : null,
      lng: d ? d.lng : null,
      url: c.url,
      thumbnail: c.thumbnail,
      // Ради этого всё и затевалось: у кого именно дешевле. Traveloka,
      // klook, Prestigia и прочих не парсит ни один из остальных источников.
      best_partner: bestOffer ? bestOffer.partner : null,
      offers_count: offers.length,
      offers,
    });
    saved++;
  }
  const withOffers = cards.filter(c => details.get(c.entity_id)?.offers.length).length;
  console.log(`   сохранено ${saved}, без цены ${noPrice}, с вариантами партнёров ${withOffers}`);
  if (!saved) console.log('   ⚠️  цен не распознано — вёрстка могла поменяться');
}

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
    // Каждый источник в своём try: раньше падение одного обрывало весь прогон,
    // и Agoda не запускалась из-за таймаута селектора на Trip.com.
    const sources = [
      [runBooking,  'Booking.com', scrapeBooking],
      [runTrip,     'Trip.com',    scrapeTrip],
      [runOstrovok, 'Ostrovok.ru', scrapeOstrovok],
      [runAgoda,    'Agoda.com',   scrapeAgoda],
      [runGoogle,   'Google',      scrapeGoogle],
    ];
    for (const [enabled, label, fn] of sources) {
      if (!enabled) continue;
      try {
        await fn(page);
      } catch (e) {
        if (e.message.includes('closed')) throw e;
        console.error(`\n⚠️  ${label} не отработал: ${e.message.split('\n')[0]}`);
        console.error('   Иду к следующему источнику.\n');
      }
    }
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
      ...(runAgoda    ? ['agoda']    : []),
      ...(runGoogle   ? ['google']   : []),
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
  console.log(`   Agoda.com:    ${hotels.filter(h => h.source === 'agoda').length}`);
  console.log(`   Google:       ${hotels.filter(h => h.source === 'google').length}`);
  console.log(`   С ценой:      ${hotels.filter(h => h.price_per_night_rub).length}`);
  console.log(`\n▶  Следующий шаг: node build-report.js`);
}

main().catch(err => {
  console.error('Fatal:', err);
  process.exit(1);
});
