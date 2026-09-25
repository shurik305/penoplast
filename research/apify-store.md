# Дослідження: Apify Store (платні Actors) (2026-09-24)

## Висновок: ДРУГИЙ ТРЕК (відкладено до отримання доступів)
Реальний попит і виплати, але довгий шлях до першої виплати та потрібен повний доступ до мережі.

## Ключові факти
- Модель pay-per-event (PPE) — стандарт; rental припиняється 01.10.2026 (SOURCED, blog.apify.com).
- Автору 80% доходу мінус витрати платформи (SOURCED, docs.apify.com).
- Виплати: мін. $20 через PayPal, $100 банком; щомісячно (SOURCED, help.apify.com).
- Фізособи можуть отримувати виплати (18+, KYC) (SOURCED).
- Apify заявляє ~$1M/міс виплат авторам (самозвіт, не аудит).
- ~99% акторів мають 0–5 користувачів (UNVERIFIED, dev.to); автор 98 акторів: перші 2 міс. — одиниці запусків, тяга з 3–4 міс. (Apify blog).
- Оцінка: перший платний запуск 1–2 міс.; $100+/міс через 2–4 міс.; пасивний дохід — 6–12 міс., портфель акторів.

## Юридичні ризики (Данія/ЄС)
- BoligPortal v. ReData (Sø- og Handelsretten, 22.10.2025): систематичний скрейпінг бази оголошень порушує
  sui generis право на базу даних і закон про маркетинг → заборона. Данські портали (житло, вакансії) — ризиковано.
- Персональні дані = GDPR. Уникати логінів, профілів людей, великих каталогів.

## Ідеї з низьким ризиком
1. Збагачення відкритих державних даних (реєстри, тендери, гранти з відкритою ліцензією).
2. Витяг відгуків про SaaS/застосунки + категоризація болів (публічні, псевдонімні).
3. Нішеві борди вакансій поза Данією (після юридичної перевірки ToS).

## Потрібно для старту
- Акаунт Apify (власник), API-токен, мережа "Full" (зараз api.apify.com заблоковано).

## Джерела
- https://docs.apify.com/platform/actors/publishing/monetize/pay-per-event
- https://blog.apify.com/standardizing-actor-pricing/
- https://docs.apify.com/platform/actors/publishing/monetize/pricing-and-costs
- https://help.apify.com/en/articles/10057167-how-developer-payouts-work
- https://docs.apify.com/legal/store-publishing-terms-and-conditions
- https://blog.apify.com/building-98-actors-on-apify-store/
- https://dev.to/agenthustler/the-apify-actor-survival-guide-why-99-of-scrapers-get-zero-users-and-how-to-fix-it-5eoh
- https://docs.apify.com/actors/publishing/quality-score
- https://apify.com/ideas
- https://bdkadvokati.com/boligportal-v-redata-danish-court-finds-data-scraping-infringes-database-rights
- https://techcrunch.com/2024/11/13/nokia-acquires-rapid-the-api-company-once-valued-at-1b/
