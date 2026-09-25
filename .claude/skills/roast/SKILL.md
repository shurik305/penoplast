---
name: roast
description: Run a critical "roast" of the autonomous-business project with independent critic subagents, then triage the findings into decisions and actions. Use when the owner asks for a roast, critique, audit or "прожарка" of the project, before big decisions (new product, new channel, any spending), or about every 7 days during autonomous operation.
---

# Roast: independent critical review of the project

The goal is to move the project forward, not to produce a report. A roast ends with changed decisions, fixed
problems and updated state files.

## 1. Snapshot (2 min)
Read ops/STATE.md, ops/LEDGER.md, the last 3 entries of ops/JOURNAL.md, experiments/*.md and ops/ROASTS.md.
Note which findings from the previous roast are still open.

## 2. Launch critics in parallel (background)
Use the `project-critic` agent (subagent_type "project-critic"; if that type is not available in this session,
use general-purpose and paste the body of .claude/agents/project-critic.md into the prompt). Launch two lenses:
- **Lens A — Buyer, market and channel:** who buys, why, at what price, how they find it; test market claims with
  WebSearch; competitors; pricing; listing SEO; trust (reviews); alternative channels and business models.
- **Lens B — Execution, quality and risk:** owner friction and blockers, process waste, unverified claims in
  listings, adversarial testing of products (edge-case inputs in a temp copy, recalculated with LibreOffice),
  legal/platform/privacy risks, cost of agent work vs expected revenue, stop rules.
Give each critic today's date, the one-paragraph project context and the specific questions for this roast.
Add a third lens only for a big decision (e.g. "Devil's advocate for the proposed pivot").

## 3. Self-roast while critics run
Before reading their reports, write your own top 5 criticisms (avoid anchoring on theirs).

## 4. Triage every finding
Accept / Reject (with a reason) / Needs data (with the cheapest way to get it). Verify surprising claims yourself;
critics can be wrong. Merge duplicates. Order accepted items by impact on the first real payment.

## 5. Act
- Fix accepted small items immediately (copy, claims, bugs, docs), with QA.
- Turn larger items into numbered next actions in ops/STATE.md.
- If the owner is needed, rewrite ops/OWNER_ACTIONS.md so the ask is as small as possible.
- Never publish, spend or message third parties just because a critic suggested it — normal permission rules apply.

## 6. Record
Append to ops/ROASTS.md: date, lenses used, each critic's verdict, the triage table (finding → decision → action),
what was fixed in this session, and the date of the next roast.

## 7. Report to the owner (Ukrainian, concise)
Strengths, weaknesses, decisions taken, what changed, the development path, and the (minimal) ask.
