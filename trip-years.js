'use strict';

/**
 * Год постройки, год ремонта и звёзды с Trip.com.
 *
 * Ни один из пяти источников самого парсера года не отдаёт: в выгрузке есть
 * цена, рейтинг и расстояние, а «поновее или постарее» - главный критерий
 * после цены, и по таблице его было не проверить. До 17.08.2026 справочник
 * `vacation/data/hotel_years.json` заполнялся руками по одному отелю, и по
 * Дананту в нём стояло 22 записи на 2 620 отелей.
 *
 * Работает обычными GET-запросами, без браузера и без рендера:
 *
 *   1. Внутренние id отелей берутся из SEO-списков вида
 *      `/hotels/da-nang-hotels-list-1356/`. Обычная выдача (`/hotels/list`)
 *      для этого не годится: она отдаёт двенадцать карточек и подгружает
 *      остальное скроллом, а в headless-браузере не подгружает вовсе -
 *      проверено 17.08.2026, двенадцать и на пятидесяти прокрутках.
 *      SEO-страница отдаёт сорок отелей сразу и ссылается на такие же
 *      списки по районам (`/zone14494/`), а страница отеля - на соседние
 *      отели. Обход идёт по этим ссылкам вширь, пока находятся новые.
 *
 *   2. Годы лежат в SSR-JSON прямо в HTML страницы отеля: `openYear`,
 *      `fitmentYear`, `starInfo.level`. JSON там экранирован, поэтому
 *      кавычка в регулярках - `\\?"`.
 *
 * Число номеров Trip.com в SSR не отдаёт; оно достаётся с Островка на
 * стороне vacation (scripts/fill_hotel_years.py).
 *
 *     node trip-years.js                 # обход до упора
 *     node trip-years.js --max=300       # не больше 300 отелей
 *     node trip-years.js --concurrency=8
 */

const fs = require('fs');
const path = require('path');

const config = JSON.parse(fs.readFileSync('hotel-config.json', 'utf8'));
const args = process.argv.slice(2);
const numArg = (name, def) => {
  const a = args.find((x) => x.startsWith(`--${name}=`));
  return a ? parseInt(a.split('=')[1], 10) || def : def;
};

const CITY = config.город.название;
const CITY_ID = config.город.cityid_trip;
const CITY_SLUG = config.город.slug_trip || config.город.slug;
const { заезд: checkin, выезд: checkout } = config.даты;
// Страница приезжает за секунду-полторы и рендера не требует, поэтому её
// можно тянуть в несколько потоков. Пять - та же осторожная величина, что
// у поимённого добора Google в самом парсере.
const CONCURRENCY = numArg('concurrency', 5);
const MAX_HOTELS = numArg('max', 0);
const UA = 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36'
  + ' (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36';

const log = (m) => console.log(`${new Date().toISOString().slice(11, 19)}  ${m}`);

async function get(url) {
  const res = await fetch(url, {
    headers: { 'User-Agent': UA, 'Accept-Language': 'ru-RU,ru;q=0.9' },
  });
  return res.ok ? res.text() : '';
}

/** Значение поля из экранированного SSR-JSON: openYear, fitmentYear, level. */
function grab(html, key) {
  const m = html.match(new RegExp(`\\\\?"${key}\\\\?":\\\\?"?([^",\\\\]+)`));
  return m ? m[1] : null;
}

/** Все пары (id, slug) со страницы: и в списках, и в блоке «похожие отели». */
function hotelLinks(html) {
  return [...html.matchAll(/hotel-detail-(\d+)\/([a-z0-9-]+)\//g)]
    .map((m) => ({ id: m[1], slug: m[2] }));
}

/**
 * Соседние SEO-списки того же города. Их два семейства, и второе даёт
 * заметно больше: кроме районов (`.../zone14494/`) Trip.com держит списки
 * по ориентирам - пляж Ми Кхе, Драконий мост, аэропорт, вокзал
 * (`da-nang-my-khe-beach/hotels-c1356m123/`). Без них обход упирался
 * в 283 отеля и терял, например, Monarque.
 */
function listLinks(html) {
  const zones = new RegExp(
    `https://ru\\.trip\\.com/hotels/${CITY_SLUG}-hotels-list-${CITY_ID}/[a-z0-9-]+/`, 'g');
  const spots = new RegExp(
    `https://ru\\.trip\\.com/hotels/${CITY_SLUG}-[a-z0-9-]+/hotels-c${CITY_ID}[a-z]\\d+/`, 'g');
  return [...new Set([...(html.match(zones) || []), ...(html.match(spots) || [])])];
}

async function crawlIds() {
  const seed = `https://ru.trip.com/hotels/${CITY_SLUG}-hotels-list-${CITY_ID}/`;
  const pages = [seed];
  const visited = new Set();
  const hotels = new Map();
  while (pages.length) {
    const batch = pages.splice(0, CONCURRENCY).filter((u) => !visited.has(u));
    batch.forEach((u) => visited.add(u));
    const htmls = await Promise.all(batch.map((u) => get(u).catch(() => '')));
    for (const html of htmls) {
      hotelLinks(html).forEach((h) => hotels.set(h.id, h.slug));
      listLinks(html).forEach((u) => { if (!visited.has(u)) pages.push(u); });
    }
    log(`   списки: обойдено ${visited.size}, в очереди ${pages.length}, `
      + `отелей ${hotels.size}`);
  }
  return [...hotels.entries()].map(([id, slug]) => ({ id, slug }));
}

async function fetchYears(hotel) {
  // Каноническая SEO-ссылка, а не `/hotels/detail/?hotelId=`: вторая форма
  // закрыта в robots.txt Trip.com (`Disallow: /hotels/detail/?hotelId=*`),
  // первая открыта. Отдают они одну и ту же страницу с тем же SSR-JSON.
  const url = `https://ru.trip.com/hotels/${CITY_SLUG}-hotel-detail-${hotel.id}/`
    + `${hotel.slug}/`;
  const html = await get(url);
  if (!html) return { ...hotel, ошибка: 'пустой ответ' };
  const year = (v) => (/^\d{4}$/.test(v || '') ? parseInt(v, 10) : null);
  return {
    ...hotel,
    имя: grab(html, 'nameLocale') || grab(html, 'name') || hotel.slug,
    построен: year(grab(html, 'openYear')),
    ремонт: year(grab(html, 'fitmentYear')),
    звёзд: parseInt(grab(html, 'level'), 10) || null,
    // Страница отеля ссылается на соседние - это и есть топливо обхода.
    соседи: hotelLinks(html),
  };
}

async function main() {
  if (!CITY_ID) throw new Error('cityid_trip в hotel-config.json не задан');
  log(`Trip.com: годы по городу ${CITY} (cityId ${CITY_ID}, slug ${CITY_SLUG})`);

  const queue = await crawlIds();
  log(`из SEO-списков: ${queue.length} отелей`);

  const done = new Map();
  const stamp = new Date().toISOString().slice(0, 19).replace('T', '_').replace(/:/g, '-');
  const file = path.join('output', `trip_years_${CITY}_${stamp}.json`);
  // Чекпоинт после каждой пачки: обход на несколько сотен отелей идёт минуты,
  // и терять его из-за обрыва сети нельзя - правило про long-running задачи.
  const save = () => fs.writeFileSync(file, JSON.stringify(
    [...done.values()].map(({ соседи, ...rest }) => rest), null, 1));

  while (queue.length && (!MAX_HOTELS || done.size < MAX_HOTELS)) {
    const chunk = [];
    while (chunk.length < CONCURRENCY && queue.length) {
      const h = queue.shift();
      if (!done.has(h.id)) chunk.push(h);
    }
    if (!chunk.length) break;
    const got = await Promise.all(chunk.map((h) => fetchYears(h).catch(
      (e) => ({ ...h, ошибка: String(e).slice(0, 80) }),
    )));
    for (const h of got) {
      done.set(h.id, h);
      for (const n of h.соседи || []) {
        if (!done.has(n.id)) queue.push(n);
      }
    }
    save();
    const known = [...done.values()].filter((h) => h.построен).length;
    log(`   отелей ${done.size}, год известен у ${known}, в очереди ${queue.length}`);
  }
  save();
  const all = [...done.values()];
  log(`готово: ${file}`);
  log(`отелей ${all.length}, год у ${all.filter((h) => h.построен).length}, `
    + `ремонт у ${all.filter((h) => h.ремонт).length}, `
    + `звёзды у ${all.filter((h) => h.звёзд).length}`);
}

main().catch((e) => {
  console.error('упало:', e);
  process.exit(1);
});
