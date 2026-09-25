# STATE — точка продовження для агента

Оновлено: 2026-09-25 (сесія 2). Гілка: `claude/peaceful-gates-emn8g8`.
Наступне автоматичне продовження: див. кінець файлу.

## Місія
Довести хоча б одну бізнес-модель до реальної оплати, потім оптимізувати прибуток. Бюджет 0 DKK без дозволу.
Правила: README.md. Дохід — лише підтверджений (ops/LEDGER.md).

## Поточний експеримент
EXP-001 — цифрові інструменти для клінінгових бізнесів на Etsy (experiments/EXP-001-etsy-cleaning-templates.md).
- Продукт №1 "Cleaning Price & Quote Calculator" — ГОТОВИЙ (4 xlsx + PDF, 8 зображень, опис, теги).
- Продукт №2 "Commercial Cleaning Bid Calculator" — ГОТОВИЙ (products/commercial-bid-calculator, 7 зображень).
- Спільні модулі: products/_lib (xlsxkit, qakit, mockkit, guidekit) — використовувати для нових продуктів.
- Статус: чекаємо рішень/налаштувань власника (ops/OWNER_ACTIONS.md: A1–A6, B1–B5).

## Відновлення середовища (свіжий контейнер)
```bash
source ops/setup.sh                       # pip-пакети, LibreOffice Calc, профіль з перерахунком формул
cd products/pricing-calculator
python3 build.py dist && python3 qa.py dist            # збірка + QA (має бути без PROBLEM)
python3 mockups.py dist listing/images                  # зображення лістингу (+ listing/raw/)
python3 guide.py listing/raw/quote-calculator.png dist/Quick-Start-Guide.pdf
cd ../commercial-bid-calculator
python3 build.py dist && python3 qa.py dist && python3 mockups.py dist listing/images
python3 guide.py listing/raw/bid-calculator-1.png dist/Quick-Start-Guide.pdf
```
Файли `dist/` і `listing/raw/` не комітяться (репозиторій був публічним) — відтворюються з коду.

## Наступні дії (по черзі)
1. Перевірити, чи власник відповів (чат; Gmail — лише листи від власника/Etsy). Оновити OWNER_ACTIONS.md.
2. Якщо є `ETSY_KEYSTRING` + `ETSY_REFRESH_TOKEN` і мережа до openapi.etsy.com:
   `python3 tools/etsy_api.py me` → `taxonomy templates` → `publish products/pricing-calculator/listing/listing.json --made-by "<варіант A3>"`.
   Записати listing_id і дату в EXP-001; витрату $0.20 — у LEDGER (Expenses) лише після підтвердження дозволу A1.
3. Поки чекаємо: продукт №3 (Short-let / Airbnb turnover kit: чекліст + журнал + трекер витратних матеріалів, ~79 DKK)
   на модулях products/_lib (шаблон — products/commercial-bid-calculator). Після №3 більше продуктів не будувати,
   доки немає даних про попит з Etsy (правило: не створювати багато до перевірки попиту).
4. Після публікації: щотижня `tools/etsy_api.py listings` (перегляди/обране) і `receipts` (продажі) → EXP-001 метрики.
   Критерії: ≥100 переглядів за 14 днів; ≥1 продаж за 30 днів; стоп-правила — в EXP-001.
5. Другий трек (після мережі Full + акаунта Apify): research/apify-store.md.

## Відомі обмеження
- Мережа "trusted": лише GitHub/npm/PyPI; WebFetch заблокований; WebSearch працює.
- Не просити власника вставляти токени в чат; секрети — лише як змінні середовища.
- Холодні email/DM у Данії заборонені (Markedsføringsloven §10) — лише вхідні канали.

## Розклад
- Наступне автоматичне продовження: 2026-09-25 06:06 UTC (send_later, trig_01Wa4MzKcE5289D97UYUj5Yt) — продукт №3 або публікація.
- Якщо все блокується власником: перевірки не частіше ніж раз на добу.
