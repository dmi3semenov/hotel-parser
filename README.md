# Hotel Parser

Парсер отелей с Booking.com и Trip.com. Собирает цены, рейтинги, адреса — сохраняет в Excel + интерактивную HTML-карту.

## Стек

- **Node.js** + Playwright (headless: false — нужен живой браузер для обхода капч)
- **ExcelJS** — генерация xlsx-отчёта
- **Leaflet** — HTML-карта с маркерами

## Быстрый старт

```bash
npm install
npx playwright install chromium
```

### 1. Настроить параметры поиска

Отредактировать `hotel-config.json`:

```json
{
  "город": {
    "название": "Минск",
    "страна": "Беларусь",
    "query_booking": "Минск, Беларусь",
    "query_trip": "Minsk",
    "lat": 53.9024726,
    "lng": 27.5618244,
    "slug": "minsk"
  },
  "даты": {
    "заезд": "2026-04-29",
    "выезд": "2026-05-01"
  },
  "гости": {
    "взрослых": 2,
    "номеров": 1
  }
}
```

### 2. Запустить парсер

```bash
# Оба источника
npm run parse

# Только Booking.com
npm run parse:booking

# Только Trip.com
npm run parse:trip
```

Откроется браузер. На Booking.com скрипт ждёт **60 секунд** — если появится капча, нужно решить её вручную.

### 3. Сгенерировать отчёт

```bash
npm run report
```

Результат в `output/`:
- `hotels_<slug>_<date>.xlsx` — Excel с 4 листами (Итог, Все отели, По рейтингу, Booking vs Trip)
- `hotels_<slug>_<date>_map.html` — карта с цветными маркерами по ценовым категориям
- `latest.json` — последний полный массив отелей
- `run_<timestamp>/hotels.jsonl` — чекпоинт (append-log, не теряется при краше)

## Структура проекта

```
hotel-parser/
├── hotel-config.json    # параметры поиска (город, даты, гости)
├── hotel-parser.js      # парсер Booking.com + Trip.com
├── build-report.js      # генерация Excel + HTML-карты
├── output/              # результаты (в .gitignore)
└── package.json
```

## Ценовые категории

| Категория       | ₽/ночь       | Цвет на карте |
|-----------------|--------------|---------------|
| Бюджет          | до 3 000     | зелёный       |
| Средний         | 3 000–7 000  | синий         |
| Выше среднего   | 7 000–15 000 | оранжевый     |
| Премиум         | от 15 000    | красный       |
