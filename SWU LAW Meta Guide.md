# Star Wars: Unlimited — LAW Meta Guide

*Post-Rotation Premier Format | Sets Legal: JTL · LOF · SEC · LAW | Updated May 30, 2026*

-----

## Set Legality

|Set                     |Code|Status   |
|------------------------|----|---------|
|Spark of Rebellion      |SOR |❌ Rotated|
|Shadows of the Galaxy   |SHD |❌ Rotated|
|Twilight of the Republic|TWI |❌ Rotated|
|Jump to Lightspeed      |JTL |✅ Legal  |
|Legends of the Force    |LOF |✅ Legal  |
|Secrets of Power        |SEC |✅ Legal  |
|A Lawless Time          |LAW |✅ Legal  |


> Cards from rotated sets are still legal if reprinted in a legal set (e.g. Death Trooper via SEC). Jango Fett (TWI) has no confirmed reprint — treat as Eternal-only. Only trust tournament data from **March 13, 2026 onward** for post-rotation meta analysis.

-----

## Key Resources

|Resource           |URL                      |Best For                                                           |
|-------------------|-------------------------|-------------------------------------------------------------------|
|SWU Competitive Hub|swu-competitivehub.com   |Tournament results, winner decklists, meta stats by set            |
|SWU Meta Stats     |swumetastats.com         |Win rates, APR data, matchup matrices, tier list (KyberScore)      |
|SWU Meta Stats API |swumetastats.com/api-docs|Raw JSON: archetypes, matchups, tier list, trends (public, no auth)|
|SWU.fan            |swu.fan/meta/leaders     |Quick leader WR/meta share overview                                |
|The SWU Report     |swu.report               |In-depth APR analysis, tournament reports, strategy articles       |
|SWUDB              |swudb.com                |Card database, deck builder                                        |
|SWU Codex          |swu.codex.gg             |Card text, market prices, meta breakdowns per leader               |
|Melee.gg           |melee.gg                 |Official tournament platform — source of all decklist data         |

-----

## Understanding the Statistics

### Win Rate (WR)

Raw percentage of games won. Anything above 52% is meaningfully positive at scale.

### Meta Share

Percentage of the competitive field playing a given deck. High share ≠ high performance.

### Field → Top 8 APR (Above Projected Rate)

Measures whether a deck converts its field representation into Top 8 finishes better or worse than expected. If 10% of the field plays Lando and Lando takes 19% of Top 8 slots, it’s over-performing — positive APR. A negative APR means the deck is over-registered relative to its actual results.

### T8 → Win APR

Same concept, but measures how well Top 8 finishers actually *close* — converting Top 8 appearances into wins. A deck can have great Field→T8 APR but poor T8→Win APR (makes cuts but can’t finish). Dedra Meero is a classic example.

### KyberScore (swumetastats.com)

A composite tier score combining win rate, meta share, Top 8 conversion, and matchup data. More reliable than any single metric alone.

-----

## Current Tier List (LAW Premier, post-rotation)

### S Tier — Register these to win

|Deck                                 |WR   |Meta Share|Key Strength                                                     |
|-------------------------------------|-----|----------|-----------------------------------------------------------------|
|Lando Calrissian (LAW) / Lake Country|56.9%|9.7%      |Best Field→T8 APR (+87.9%). Credit engine out-resources everyone.|
|Boba Fett (JTL) / Lake Country       |52.8%|12.9%     |Most Top 8 appearances (101+). Highest raw consistency in format.|
|Aurra Sing (LAW) / Data Vault        |53.9%|3.3%      |Primary Lando counter (55.6% vs Lando). Bounty hunter tempo.     |
|Dedra Meero (SEC) / Colossus         |58.5%|2.4%      |Highest raw WR. Reliable Swiss performer. Strong control shell.  |

### A Tier — Competitive, require matchup knowledge

|Deck                                    |WR   |Meta Share|Key Strength                                                       |
|----------------------------------------|-----|----------|-------------------------------------------------------------------|
|Obi-Wan Kenobi (LOF) / Blue Force       |52.1%|10.2%     |Largest field share. Solid vs Lando. Beaten badly by Vader.        |
|Mother Talzin (LOF) / Yellow Force      |53.0%|5.2%      |Versatile Force deck. Multiple viable base options.                |
|Chewbacca (LAW) / Cunning 30 HP         |54.8%|3.6%      |Underplayed. Community consensus: higher ceiling than results show.|
|Colonel Yularen (SEC) / Aggression 30 HP|54.5%|3.6%      |+34% / +87.5% two-stage APR. Quietly overperforming.               |
|Darth Vader (JTL) / Cunning 30 HP       |52.1%|3.9%      |68% vs Obi-Wan. Won London Sector (482p). Terrible vs Lando (30%). |
|Luke Skywalker (JTL) / Data Vault       |47.7%|6.6%      |High field share, near-zero wins. Closing APR -57.6%. Avoid.       |
|Admiral Piett (JTL) / Blue              |51.9%|3.7%      |Solid Swiss performer. Limited Top 8 closing rate.                 |

### B Tier — Viable, reward specialist knowledge

|Deck                               |Notes                                                                                        |
|-----------------------------------|---------------------------------------------------------------------------------------------|
|Sabé (SEC) / Data Vault            |High ceiling, underplayed. Won an 82-player event. Two-stage APR suggests breakout potential.|
|Tobias Beckett (LAW) / Red         |Community: “quietly solid.” New LAW leader. Under the radar.                                 |
|Kazuda Xiono (JTL) / Data Vault    |Won Prague Regional (663p). Kaz on Falcon is a legitimate closing threat.                    |
|Lando Calrissian (LAW) / Data Vault|Alternative Lando build. Less dominant than Lake Country version.                            |

### C Tier — Fringe / High variance

|Deck                                     |Notes                                                                 |
|-----------------------------------------|----------------------------------------------------------------------|
|The Client (LAW) / Red                   |22 Top 8s, 0 wins. Classic convergence failure — breakout candidate.  |
|Qui-Gon Jinn (LOF) / Green Force         |+71.4% T8→Win APR. Tiny sample but exceptional closing rate.          |
|Darth Maul (LOF) / Blue Force            |Won Springfield MO event as only pilot in field. Modest aggregate APR.|
|Grand Admiral Thrawn (JTL) / Yellow Force|Slow-build control. Below-average LAW season representation.          |
|Admiral Ackbar (JTL) / Data Vault        |Won Prague Regional (663p). Low meta share, high ceiling.             |

-----

## Matchup Matrix (Win % for row deck vs column deck)

> 🟩 ≥60% heavily favoured · 🟢 55–59% favoured · ⬜ 50–54% slight edge · 🟡 45–49% slight disadvantage · 🟠 40–44% unfavoured · 🔴 <40% heavily unfavoured

|               |Lando LC|Boba LC|Aurra DV|Dedra Col|Obi-Wan BF|Talzin YF|Chewbacca|Yularen|Vader Cun|Luke DV|Piett|Tobias|**Avg** |
|---------------|--------|-------|--------|---------|----------|---------|---------|-------|---------|-------|-----|------|--------|
|**Lando LC**   |—       |53     |44      |50       |53        |55       |55       |56     |70       |62     |55   |57    |**55.5**|
|**Boba LC**    |47      |—      |45      |50       |65        |53       |52       |55     |55       |58     |54   |56    |**53.6**|
|**Aurra DV**   |56      |55     |—       |52       |45        |53       |50       |52     |22       |54     |52   |54    |**49.5**|
|**Dedra Col**  |50      |50     |48      |—        |54        |52       |50       |54     |55       |57     |53   |55    |**52.5**|
|**Obi-Wan BF** |47      |35     |55      |46       |—         |42       |45       |37     |68       |55     |50   |52    |**48.4**|
|**Talzin YF**  |45      |47     |47      |48       |58        |—        |48       |50     |52       |54     |50   |52    |**50.1**|
|**Chewbacca**  |45      |48     |50      |50       |57        |52       |—        |52     |55       |56     |52   |53    |**51.8**|
|**Col Yularen**|44      |45     |48      |46       |63        |50       |48       |—      |55       |63     |52   |54    |**51.6**|
|**Vader Cun**  |30      |45     |78      |45       |32        |48       |45       |45     |—        |55     |50   |52    |**47.7**|
|**Luke DV**    |38      |42     |46      |43       |45        |46       |44       |37     |45       |—      |48   |50    |**44.0**|
|**Piett Blue** |45      |46     |48      |47       |50        |50       |48       |48     |50       |52     |—    |51    |**48.6**|
|**Tobias Red** |43      |44     |46      |45       |48        |48       |47       |46     |48       |50     |49   |—     |**46.7**|

*Sources: SWU Report APR series weeks 3–7, Springfield MO PQ matchup matrix, Prague Regional data, swu.fan leader WRs. All data from post-March 13, 2026 Premier events only.*

-----

## The Meta Structure: Hub and Spoke

The LAW format has a clear structural identity:

**Lando/LC is the hub.** Everything else is defined by its relationship to Lando. The two anti-Lando spokes are Aurra Sing (55.6% vs Lando) and Dedra Meero (~50% vs Lando). Boba Fett is an outlier — it doesn’t beat Lando convincingly but wins through sheer volume and consistency regardless.

**Obi-Wan is the paradox deck.** Largest archetype by field share (~10%) but posted near-zero wins for the first seven weeks of the format. The field adapted fast. Vader specifically preys on it (68% WR) by playing entirely in the space arena while Obi-Wan is a ground-focused deck.

**Luke is the trap.** 6.6% meta share with a -57.6% closing APR. One of the clearest “do not register this” signals in the data. Over-represented, under-performing, and beaten by both of the top two decks.

-----

## Deck Selection Guide

### Want the optimal pick?

**Lando Calrissian (LAW) / Lake Country.** The data is unusually clear. Best Field→T8 APR in the format. 56.9% WR. Credit engine is resilient. Weak spots (Aurra, mirrors) are manageable.

### Want to target-tech the meta?

**Aurra Sing (LAW) / Data Vault** if your expected field is Lando-heavy (>15% of field). Edges both Lando and Boba. Hard counter warning: Vader beats Aurra at 78%. Check your local meta first.

### Want a safe, proven, consistent choice?

**Boba Fett (JTL) / Lake Country.** 101+ Top 8 appearances. The deck has been strong since JTL and rotation didn’t hurt its core. If you know the deck, register it.

### Want to punch above your weight?

**Colonel Yularen (SEC) / Aggression 30 HP.** +34% / +87.5% two-stage APR. Under-the-radar. Rewards technical play and aggro experience. Nobody is specifically preparing for it.

### Want a sleeper/breakout pick?

**The Client (LAW) / Red.** 22 Top 8 appearances, 0 wins. A convergence failure of that size is rare and usually corrects. Skilled pilot on an unscouted deck in a diverse local meta = real upside.

### What to avoid?

- **Luke Skywalker / Data Vault** — statistically the worst-performing over-registered deck in the format.
- Any deck relying on SOR/SHD/TWI-only cards — those are illegal in Premier.

-----

## Darth Vader Deep Dive (Notable Archetype)

Vader (JTL) / Cunning 30 HP is a polarized specialist deck worth understanding even if you don’t play it, because it preys on Obi-Wan — the most common deck in the field.

**Core engine:** Victor Leader (+1/+1 to all space units) + flood of cheap TIEs + Clone Combat Squadron (scales with board) → Vader deploys as a pilot upgrade onto Clone Combat Squadron for a massive closing threat.

**Key cards:** Victor Leader, Clone Combat Squadron, Lurking Snub Fighter (exhaust), Chancellor Palpatine (ground lock), Salvage (recursion), Hold For Questioning (hand disruption), Screeching TIE.

**Plot setup:** Lurking Snub Fighter + Chancellor Palpatine is the ideal Vader flip combo. Against Vigilance decks, swap to Garindan to strip Hyperspace Disaster from hand first.

**Matchup summary:** Great vs Obi-Wan (68%), solid vs Talzin/Chewbacca/Luke. Terrible vs Lando (30%) and Aurra Sing (22%). Only register if your expected field is light on Lando.

-----

## Quick Reference: Why the Top Decks Win

**Lando (LAW) / Lake Country** — Credit tokens generate economic advantage every turn. Lando’s deploy triggers a bonus credit destruction. Liberty (space sentinel) threatens to lock the space arena. The deck simply out-resources and out-values opponents over 6+ rounds.

**Boba Fett (JTL) / Lake Country** — Non-combat damage triggers Boba’s passive every turn, stacking indirect damage. Can deploy as a pilot upgrade onto a Vehicle for a huge burst. Bounty Hunter package (Zam Wesell, Cad Bane, Fett’s Firespray) is efficient and resilient. Lake Country’s 30 HP neutral base gives more time to set up.

**Aurra Sing (LAW) / Data Vault** — Assassination tempo: remove threats efficiently, generate resources, close with high-power units. Data Vault accelerates the engine. The deck’s Cunning/Villainy shell accesses the best removal in the format.

**Dedra Meero (SEC) / Colossus** — Imperial control. Colossus base synergises with the deck’s unit package. High raw WR reflects efficient answers to threats. The convergence problem (makes Top 8s but rarely wins) is a closing-game issue, not a deck quality issue.

-----

## Tournament Structure Reference

|Level                |Significance                                                     |
|---------------------|-----------------------------------------------------------------|
|Store Showdown       |Local level. Small (10–30 players). Low data weight.             |
|Planetary Qualifier  |Regional level. 30–120+ players. Primary competitive data source.|
|Sector Qualifier     |Large regional. 200–500+ players. High data weight.              |
|Regional Championship|Major. 500–700+ players. Best data for matchup analysis.         |

The SWU Competitive Hub meta stats only include events with **>32 players** to filter noise. For serious analysis, weight Sector and Regional events most heavily.