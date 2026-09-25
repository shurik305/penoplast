# Дослідження: open-source bounties (2026-09-24)

Метод: лише WebSearch. Більшість цифр — зі зведень пошуку; багато джерел 2026 р. самі є AI-агентами.

## Висновок: ВІДХИЛЕНО як перший експеримент
Ринок перенасичений AI-агентами; легітимних відкритих bounty мало; кілька незалежних спроб 2026 р. дали $0.

## Ключові факти
- Algora: `/attempt` → PR з `/claim #N` → виплата через Stripe Express за 1–3 дні після merge (SOURCED, docs.algora.io).
  Комісія для виконавця — дані суперечливі (4%–23%) (UNVERIFIED).
- ~561 відкритих bounty на $1.14M, але 99.3% суми в 3 репозиторіях; лише ~5 "реально доступних" (незалежний перепис, серп. 2026).
- Конкуренція: 8–158 PR на свіжий bounty за кілька годин; один bounty $341 — 236 спроб за 21 міс. без merge.
- "I Sent an AI Agent to Hunt Open-Source Bounties for Three Days: It Earned $0" (HackerNoon); "21 bounties — won zero" (dev.to).
- Сміттєві/honeypot "bounty"-репозиторії для AI-агентів: 35 із 60 вибраних результатів — боти/ферми.
- curl закрив bug bounty (січ. 2026) через AI-спам; GitHub урізав виплати bug bounty на 50%+ (лип. 2026).
- Проєкти, що забороняють AI-внески: Zig, NetBSD, GIMP, QEMU, GCC, Godot та ін. Мейнстрим вимагає розкриття + тести.

## Можливі винятки на майбутнє (низький пріоритет)
- huntr.com (безпека AI/ML бібліотек, Stripe, $150–4000) — потрібні справжні верифіковані вразливості.
- Expensify ($250 за баг через Upwork, спершу треба отримати призначення).
- Tenstorrent tt-metal ($1.5k–12.5k, C++/ML-ядра, забороняє автоматизоване "захоплення" задач).

## Джерела
- https://docs.algora.io/bounties/workflow
- https://github.com/AsherKasper/bounty-census
- https://github.com/ztc00/algora-scout/blob/main/POST.md
- https://dev.to/aion_autonomous_org/i-measured-the-open-source-bounty-market-before-entering-it-then-i-didnt-enter-93n
- https://hackernoon.com/i-sent-an-ai-agent-to-hunt-open-source-bounties-for-three-days-it-earned-$0
- https://dev.to/pogo-tb/we-entered-21-agent-marketplace-bounties-and-won-zero-the-math-says-that-was-the-expected-result-4fak
- https://dev.to/listwright/githubs-search-api-says-327-open-bounties-i-read-all-60-recent-ones-35-were-bots-and-bounty-2p49
- https://github.com/melissawm/open-source-ai-contribution-policies
- https://www.theregister.com/devops/2026/07/23/github-slashes-public-bug-bounty-payouts-as-ai-report-flood-buries-its-security-team/5277046
- https://thenewstack.io/curls-daniel-stenberg-ai-is-ddosing-open-source-and-fixing-its-bugs/
- https://huntr.com/guidelines
- https://docs.tenstorrent.com/bounty_terms.html
