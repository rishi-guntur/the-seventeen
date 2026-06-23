# The Seventeen — Critical Minerals Primer & Investment Thesis

*A primer on the rare-earth / critical-minerals supply chain, with a directional investment view on where the asymmetry sits.*

**Last updated:** 2026-06-21
**Author:** Rishi Guntur
**Status:** Working draft — investment views and single-name thesis to be filled in by author. Placeholders are marked in **bold double-brackets** (`[[ ... ]]`); see the checklist at the end of this file.

> **A note on placeholders.** This document is a scaffold. Industry-context and primer facts are written and cited inline with a date. Anything that is my own judgment, a live price, or a deposit-specific input is left as a marked placeholder of one of three types:
> - **`[[MY VIEW: ... ]]`** — my analysis / conclusion to write.
> - **`[[VERIFY: ... — as of <date>]]`** — a fact or live data point to confirm before circulation.
> - **`[[INPUT: ... ]]`** — a hard number I must supply (e.g. a deposit's element split).

---

## Executive Summary

**[[MY VIEW: 3–4 conclusions, written last. State the single sharpest claim first (the thing I believe that the market does not yet price). Then: (i) where in the value chain the asymmetry sits, (ii) the one catalyst that forces a re-rate and roughly when, (iii) the single name (or pair) I would express it through and the one risk that would make me wrong. Keep to ~150 words; this is the only thing a busy reader is guaranteed to read.]]**

---

## Part 1 — Primer

*Table stakes. Compressed deliberately. If you know the chain, skip to Part 2.*

### 1.1 The 17 elements, at group level

"Rare earths" are the 15 lanthanides plus scandium and yttrium — 17 elements that sit together because their chemistry is nearly interchangeable, which is exactly why separating them is hard and value-additive. They are not geologically rare; they are *dispersed*, rarely concentrated enough to mine economically, and almost never found apart from one another (USGS, *Mineral Commodity Summaries 2025*, Rare Earths chapter, Jan 2025).

The economically meaningful split is **light (LREE)** vs **heavy (HREE)**:

- **Light** — lanthanum, cerium, praseodymium, neodymium, samarium (the "low" atomic numbers). Abundant; cerium and lanthanum are effectively byproducts in oversupply. The value here is concentrated in **Nd and Pr**.
- **Heavy** — europium through lutetium, plus yttrium. Scarcer, more geologically concentrated (notably in ionic-adsorption clays in southern China and Myanmar), and structurally harder to source outside China. The value here is concentrated in **dysprosium (Dy) and terbium (Tb)**.

The practical takeaway for an investor: you are almost never buying "rare earths." You are buying exposure to **four or five elements that matter** (Nd, Pr, Dy, Tb, and — for a specific magnet class — Sm) carried along by ~12 co-products that range from modestly useful to actively unsellable.

### 1.2 Where the value actually concentrates

The defining feature of this market is that **value is concentrated in a handful of elements but cost is spread across all of them.** A typical hard-rock deposit (bastnäsite/monazite) is dominated by cerium and lanthanum by mass, while the revenue is dominated by the praseodymium-neodymium ("NdPr") fraction and, where present, the heavies.

- **NdPr** is the workhorse: the feedstock for NdFeB permanent magnets, which are the demand engine (EVs, wind, robotics, defense, consumer electronics). This is where the bulk of basket revenue sits for most Western projects.
- **Dy and Tb** are the *heavy* magnet additives — added in small percentages to raise the magnet's coercivity so it holds its field at high operating temperatures (EV traction motors, defense). Tiny mass fraction, outsized value and outsized supply risk, because heavies are the most China- and Myanmar-concentrated part of the chain.
- **Sm** matters for one specific, underappreciated reason: **samarium-cobalt (SmCo) magnets**, which tolerate much higher temperatures and are the default in defense/aerospace (missiles, jet actuators) where NdFeB would demagnetize. Sm is otherwise a low-value light. China's inclusion of samarium in its April 2025 export-control list (below) is a tell that the defense angle is understood by Beijing.
- **The rest** — La, Ce, and most of the other heavies (Gd, Ho, Er, Tm, Yb, Lu, Eu, Y) — matter *less* to the equity story. They have real industrial uses (catalysts, polishing, phosphors, batteries) but trade cheaply, are frequently in surplus, and in a hard-rock basket are closer to a disposal cost than a revenue line. **The "balance problem" (§1.4) is precisely the problem of being forced to produce these to get at the NdPr.**

This is the single most important primer fact for valuation work: **a deposit's grade headline is close to meaningless; the element *split* and the realized price on the NdPr/Dy/Tb fraction is what determines basket economics.** The model in `/model/basket-economics.xlsx` is built around exactly this.

### 1.3 The value chain — and where margin and the bottleneck sit

```
  MINE  →  SEPARATE  →  METAL / ALLOY  →  MAGNET  →  (motor / OEM)
 (ore,    (mixed REO   (rare-earth      (NdFeB or
  conc.)   → individual  metal, then     SmCo magnet,
           oxides)       NdFeB strip-     incl. Dy/Tb
                         cast alloy)      grain-boundary)
```

- **Mine.** Many viable orebodies exist outside China (Mountain Pass in the US, Mount Weld in Australia, plus a long pipeline of development-stage deposits). Mining is *not* the binding constraint. This is the most commoditized, lowest-margin link and the one Western capital has over-funded.
- **Separate.** Solvent-extraction separation of mixed oxides into individual high-purity oxides is the **first real bottleneck** and the step where China's dominance is most entrenched — not because the chemistry is secret, but because of accumulated process know-how, tolerated effluent/permitting, scale, and cost. China controls the large majority of *separation* capacity even for ore mined elsewhere.
- **Metal / alloy.** Converting oxide → rare-earth metal → NdFeB strip-cast alloy is the **second, deeper bottleneck**, and ex-China capacity is thinner still. This is the step most Western "mine-to-magnet" stories quietly under-resource.
- **Magnet.** Sintered NdFeB magnet-making (including the Dy/Tb grain-boundary diffusion that economizes on heavy usage) is overwhelmingly Chinese and Japanese. Greenfield Western magnet capacity is being built but is early and subsidy-dependent.

**Where the margin is:** broadly, *downstream of separation* — in separation, alloy, and magnets — not at the mine. The investable insight is that the market keeps funding the abundant link (mining) and under-funding the scarce links (separation/metal/magnet). [[MY VIEW: state whether you agree the margin pool is genuinely downstream and durable, or whether downstream margin is itself a policy artifact that compresses once capacity catches up.]]

See the interactive **[supply-chain map](../viz/supply-chain-map.html)** for where each named asset sits on this chain and which carry a policy tag — it makes the China-midstream concentration and the policy-vs-geology point visually immediate.

### 1.4 Demand drivers and the "balance problem"

**Demand drivers.** NdFeB magnet demand is the spine: EV traction motors, wind-turbine generators (especially direct-drive offshore), industrial automation / humanoid robotics, and defense (precision-guided munitions, radar, aircraft actuation). Adamas Intelligence and IEA both project NdPr-oxide demand growth running well ahead of GDP through the early 2030s, with magnet rare earths the tightest sub-segment (IEA, *Global Critical Minerals Outlook 2024/2025*; Adamas Intelligence magnet-market updates, 2024–2025). [[VERIFY: pull the latest specific NdPr demand-CAGR and any 2030 balance estimate — as of <date>; figures here are directional, not quoted.]]

Robotics/humanoids are the newest and least-priced demand vector: each humanoid actuator stack is magnet-intensive, and if unit forecasts are even partially right, it is a step-change in magnet demand that is not yet in most supply-demand balances. [[MY VIEW: do I believe the humanoid demand case enough to underwrite it, or treat it as optionality?]]

**The balance problem.** Because the 17 elements come out of the ground *together* in roughly fixed natural ratios, you cannot produce more Nd/Pr/Dy/Tb without also producing more La/Ce/Y that the market may not want. Optimizing supply for the elements in demand floods the market with the ones that aren't — capping the byproduct credit and, in a downturn, turning co-products into a cost. Any project's economics, and any price forecast, lives or dies on this. It is also why "we'll just build more mines" is not a clean answer to a magnet shortage: it oversupplies the cheap elements long before it relieves the heavies.

### 1.5 Two-market pricing

There is no single rare-earth price. There are at least two reference regimes, and they diverge:

1. **China domestic** — RMB-denominated prices for oxides/metals inside China (e.g. as reported by the Asian Metal / Shanghai Metals Market complex). This is the deepest, most liquid reference and the one that sets the global marginal cost, but it is influenced by domestic quota and stockpiling policy.
2. **Ex-China / CIF** — prices for material transacted outside China (often quoted FOB China or CIF Rotterdam, e.g. *Argus*, *Fastmarkets*). Thinner, and increasingly carrying a **security-of-supply premium** as Western buyers pay up for non-Chinese units.

Two consequences for the work: (i) **all pricing in this memo and model is opaque and goes stale fast** — every number must be dated and sourced; (ii) the *gap* between the two regimes is itself a thesis variable. Western projects are, in effect, a bet that an ex-China price (or a contracted floor) decouples upward from the China domestic price. [[VERIFY: current NdPr-oxide price in both regimes, with dates and sources — as of <date>. Do not circulate with stale prices.]]

### 1.6 Policy / catalyst landscape — framed as investable catalysts and risks

This is the part of the primer that is actually *investment* content, because in this market policy is the price.

- **China export controls (Apr 2025).** On 4 April 2025 China placed seven medium-and-heavy rare earths and related magnet items — **samarium, gadolinium, terbium, dysprosium, lutetium, scandium, and yttrium** — under export licensing (Chinese MOFCOM/Customs announcement, 4 Apr 2025; widely reported, e.g. Reuters, Apr 2025). The licensing regime throttled heavy-RE and magnet flows to Western buyers through mid-2025 and was the proximate cause of the year's supply scare. **This is the single most important catalyst in the complex** and the reason the West's downstream-capacity push has urgency. [[VERIFY: exact element list and current licensing status / any general-license carve-outs — as of <date>.]]
- **The ~Nov 2026 horizon.** Subsequent US–China de-escalation reportedly produced a general-license / suspension arrangement on rare-earth export controls with a roughly **one-year** life, putting an expiry/renewal decision point in **late 2026**. Whether controls snap back, are renewed, or are made permanent around that window is a discrete, dateable catalyst. [[VERIFY: the precise terms and expiry date of the US–China rare-earth licensing arrangement, and what specifically lapses in ~Nov 2026 — as of <date>. Frame as the lead catalyst but confirm the mechanics before relying on them.]] [[MY VIEW: how I'd position into and through this window — is it a "buy the insurance" event or a "fade the panic" event?]]
- **US DoD price floor / MP Materials (Jul 2025).** In July 2025 the US Department of Defense struck a deal with **MP Materials** that included an equity stake and, critically, a **price floor on NdPr** (reported at ~US$110/kg) plus offtake support for a new magnet facility — effectively converting a commodity producer into something closer to a regulated, floor-protected utility for a strategic input (reported widely, Jul 2025; e.g. Reuters / DoD statements). This is the template for "industrial policy as a put option" and reframes how to underwrite the Western names. [[VERIFY: the exact floor price, stake size, and tenor — as of <date>.]] [[MY VIEW: does a government price floor make MP a credit-like, floor-protected long, or does it cap the upside and socialize the downside in a way that's bad for equity? This is a core view to take.]]
- **Offtake SPVs / strategic-reserve vehicles.** Both the US and allies are standing up offtake and stockpiling vehicles (DoD, the US strategic stockpile, allied equivalents) that contract for non-Chinese units above market. Each new SPV/offtake is a de-risking catalyst for whichever project it backs. [[VERIFY: list current US/allied offtake or stockpile vehicles and what they've contracted — as of <date>.]]
- **EU CRMA.** The EU **Critical Raw Materials Act** (in force 2024) sets 2030 benchmarks — domestic capacity for extraction (10%), processing (40%), recycling (25%), and a cap of 65% of any strategic raw material from a single third country (European Commission, CRMA, 2024). It is a demand-pull/permitting-acceleration backdrop more than a hard subsidy, and it underwrites European processing/recycling plays (e.g. Carester, magnet recyclers). [[VERIFY: latest CRMA strategic-project designations relevant to names in this memo — as of <date>.]]

The throughline: **in rare earths, the catalyst calendar is a policy calendar.** Demand growth is slow-moving and largely known; the re-rates come from export controls, floors, and offtakes. That is unusually good for an event-driven or catalyst-led book and unusually bad for a pure DCF.

### 1.7 Company landscape

*One line each: value-chain position + stage. Public valuations/financials deliberately omitted — see `[[VERIFY]]`/`[[INPUT]]` flags; do not quote multiples from memory.*

| Company | Value-chain position | Stage / note |
|---|---|---|
| **MP Materials** (MP) | Mine (Mountain Pass) + separation + building magnets (Texas) | Producing miner/separator; DoD floor + magnet build-out is the US flagship. |
| **USA Rare Earth** (USAR) | Magnet-maker + Round Top (heavy-RE) deposit | Downstream-led (Stillwater, OK magnet plant); pre-scale. |
| **Lynas Rare Earths** (LYC.AX) | Mine (Mt Weld) + separation (Malaysia, now US/Texas) | Largest ex-China separator; the most established Western mine-to-oxide. |
| **Energy Fuels** (UUUU) | Uranium + monazite-fed RE separation (White Mesa, UT) | Diversified; building RE separation off monazite byproduct. |
| **Iluka Resources** (ILU.AX) | Mineral sands + Eneabba RE refinery (Australia) | Refinery under construction with Australian government debt; mineral-sands cash cow funds it. |
| **Arafura Rare Earths** (ARU.AX) | Nolans mine + integrated NdPr (Australia) | Development-stage, integrated NdPr, offtake/financing dependent. |
| **Chinese incumbents** — China Northern Rare Earth, China Rare Earth Group, JL MAG, Zhenghai, Ningbo Yunsheng | Full chain, dominant in separation + magnets | The incumbent cost and capacity benchmark; set the marginal price. |
| **Carester** (private, France) | Separation + recycling (Caremag project) | EU/recycling downstream; CRMA-aligned, partly state/financier-backed. |
| **Serra Verde** (private, Brazil) | Ionic-adsorption-clay mine (heavies ex-China) | One of the few non-China/Myanmar ionic-clay sources — heavy-RE supply. |
| **Magnet / alloy makers** — e.g. Vacuumschmelze (DE), Neo Performance Materials (Estonia plant), Less Common Metals (UK), Noveon/USA | Alloy + sintered NdFeB magnets ex-China | The scarce downstream link; mostly private/subsidized, early scale. |

[[VERIFY: confirm each company's current stage/asset status and add any material names I'm missing (e.g. Vulcan, Ucore, Aclara, Hastings) — as of <date>.]]
[[MY VIEW: which 2–3 of these are actually investable for my mandate, and why the rest are noise.]]

---

## Part 2 — Investment Views

*The core of the document. This is where I earn the read. ~60% of the memo lives here.*

### 2.1 Views

> Format for each: **Thesis** (one sentence) → **What would have to be true** (the underwriting conditions) → **What the market misses** (the variant view) → **What would make me wrong** (the kill criteria). Keep each thesis falsifiable and dated.

**View 1.** **[[MY VIEW: Thesis #1 — the structural call. e.g. "the bottleneck and therefore the durable margin is in separation/metal/alloy, not the mine, and the market keeps mispricing mine-stage assets as if they were the scarce link." State it in one sentence.]]**
- *What would have to be true:* **[[MY VIEW: the 2–3 conditions that have to hold for View 1 — e.g. ex-China separation stays capacity-short through 20XX; floors/offtakes hold; China doesn't flood.]]**
- *What the market misses:* **[[MY VIEW: the specific thing consensus has wrong / isn't pricing.]]**
- *What would make me wrong:* **[[MY VIEW: kill criteria for View 1 — the observable that, if it happens, ends the thesis.]]**

**View 2.** **[[MY VIEW: Thesis #2 — the policy-as-price call. e.g. how to treat DoD floors / offtake SPVs: are floor-protected Western producers a credit-like long (limited downside, asymmetric to a control snap-back) or a capped, subsidy-dependent equity?]]**
- *What would have to be true:* **[[MY VIEW]]**
- *What the market misses:* **[[MY VIEW]]**
- *What would make me wrong:* **[[MY VIEW]]**

**View 3.** **[[MY VIEW: Thesis #3 — the catalyst/timing call around the ~Nov 2026 export-control window. Long-vol into the window, or fade the consensus panic? What's the trade expression and the time horizon?]]**
- *What would have to be true:* **[[MY VIEW]]**
- *What the market misses:* **[[MY VIEW]]**
- *What would make me wrong:* **[[MY VIEW]]**

**View 4 (optional).** **[[MY VIEW: Thesis #4 — the demand-side / wildcard call, e.g. robotics/humanoid magnet demand as unpriced optionality, or recycling as the real Western edge. Only include if I genuinely hold it.]]**
- *What would have to be true:* **[[MY VIEW]]**
- *What the market misses:* **[[MY VIEW]]**
- *What would make me wrong:* **[[MY VIEW]]**

### 2.2 Single-name thesis

*Pick one name and go deep. This is the section a PM actually grades. Scaffold only — I fill in every line. Do not quote any financials from memory; tie each number to a `[[VERIFY]]`/`[[INPUT]]`.*

**Name:** **[[MY VIEW: the single name. State it and the position (long/short, sizing intuition) in one line.]]**

**One-paragraph thesis.** **[[MY VIEW: why this name, why now, what's the edge, what's the expected payoff and over what horizon.]]**

**Cap-structure map.** **[[MY VIEW: walk the capital structure — debt (instruments, maturities, covenants), converts/warrants, government instruments (DoD equity/floor, gov't debt), preferreds, share count and dilution overhang. Note where in the stack I'd rather be.]]**
- Inputs: **[[VERIFY: current debt, cash, maturities, share count, government instruments and their terms — as of <date>.]]**

**Liquidity runway.** **[[MY VIEW: months of runway at current burn; next financing need; whether the next raise is dilutive equity, gov't debt, or offtake prepayment; what that does to the equity.]]**
- Inputs: **[[VERIFY: latest cash balance, quarterly burn, capex schedule, undrawn facilities — as of <date>.]]**

**The specific catalyst.** **[[MY VIEW: the one dateable event that forces a re-rate (offtake signing, refinery first-oxide, DoD-style floor award, the Nov-2026 control decision) and why the market hasn't priced it.]]**

**Downside & recovery scenario.** **[[MY VIEW: the bear case and the floor — what protects me (gov't floor? asset value? offtake?), recovery on the debt/equity if the project stalls, and the realistic downside price. This is the part that earns trust.]]**
- Inputs: **[[INPUT: my downside deck — basket value at floor price, liquidation/asset value, recovery assumptions.]]**

**What the market is getting wrong.** **[[MY VIEW: the single variant-perception sentence. If I can't write this, I don't have a thesis.]]**

---

## Part 3 — Tailored framings

*One page total. Same underlying view, two audiences. Honest about each lens.*

### 3.1 Credit / catalyst lens (Readystate cut)

**[[MY VIEW: ~half a page. Frame the opportunity as a credit + catalyst trade: where in the capital structure is the asymmetry (gov't-floor-protected debt or structured paper with limited downside and a control-snap-back kicker), what's the catalyst lead into the ~Nov-2026 export-control decision, and how I'd size/hedge it. Lead with the catalyst calendar and the downside protection, not the growth story.]]**

### 3.2 Owner's lens (BDT & MSD cut)

**[[MY VIEW: ~half a page. Answer one question honestly: is any part of this chain a durable, cash-generative, founder-led business an owner would want to hold for a decade — or is it a portfolio of subsidy-dependent projects that only clear because of government floors and offtakes? If the private exception exists (e.g. a founder-led separation/recycling or magnet business, a Carester-type or a Serra Verde-type asset with real cost advantage and repeat cash flow), name it and say why it's different. If it doesn't, say that plainly — that candor is the point of this cut.]]**

---

## Sources & caveats

**Caveats.**
- **Pricing is opaque and goes stale.** There is no single rare-earth price (§1.5); quotes are thin, regime-dependent (China-domestic vs ex-China/CIF), and move on policy. Every price in this memo and in `/model/basket-economics.xlsx` is a dated `[[VERIFY]]`/`[[INPUT]]` placeholder until I refresh it. **Do not circulate with stale prices.**
- **Policy is the price.** The biggest moves are export-control and subsidy events, which are discrete and hard to forecast; treat the catalyst calendar (§1.6) as the real risk model.
- **The model is illustrative.** Element splits and prices in the spreadsheet are example placeholders until I input real deposit and market data (see the `Notes` tab and the checklist below).

**Primary sources cited (date them on refresh).**
- USGS, *Mineral Commodity Summaries 2025* — Rare Earths (Jan 2025). [[VERIFY: pull exact production/reserve figures if cited in final — as of <date>.]]
- IEA, *Global Critical Minerals Outlook* (2024/2025). [[VERIFY: latest edition + specific demand figures — as of <date>.]]
- Adamas Intelligence — NdPr / magnet-market updates (2024–2025). [[VERIFY.]]
- Chinese MOFCOM/Customs export-control announcement, 4 Apr 2025 (seven medium/heavy REEs + magnets). [[VERIFY: element list + current status — as of <date>.]]
- US DoD–MP Materials agreement, Jul 2025 (equity stake + ~US$110/kg NdPr floor + magnet offtake). [[VERIFY: terms — as of <date>.]]
- European Commission, *Critical Raw Materials Act* (in force 2024) — 2030 benchmarks (10% extraction / 40% processing / 25% recycling / 65% single-country cap). [[VERIFY: latest strategic-project list — as of <date>.]]
- Price references: *Asian Metal*, *Shanghai Metals Market*, *Argus*, *Fastmarkets* (China-domestic and ex-China/CIF). [[VERIFY: live quotes both regimes — as of <date>.]]

---

## Placeholder checklist (my to-do list)

*Every placeholder in this file, grouped. Fill these in and delete this note. The same checklist for all deliverables is reprinted in `/README.md`.*

**Executive Summary**
- [ ] `[[MY VIEW]]` — 3–4 conclusions, written last.

**Part 1 — Primer**
- [ ] §1.3 `[[MY VIEW]]` — is downstream margin durable or a policy artifact?
- [ ] §1.4 `[[VERIFY]]` — latest NdPr demand CAGR / 2030 balance.
- [ ] §1.4 `[[MY VIEW]]` — underwrite the humanoid demand case or treat as optionality?
- [ ] §1.5 `[[VERIFY]]` — current NdPr price, both regimes, dated.
- [ ] §1.6 `[[VERIFY]]` — Apr-2025 element list + current licensing status.
- [ ] §1.6 `[[VERIFY]]` — terms/expiry of US–China licensing arrangement (~Nov 2026).
- [ ] §1.6 `[[MY VIEW]]` — how to position into/through the Nov-2026 window.
- [ ] §1.6 `[[VERIFY]]` — MP/DoD floor price, stake, tenor.
- [ ] §1.6 `[[MY VIEW]]` — does the DoD floor make MP a credit-like long or cap the equity?
- [ ] §1.6 `[[VERIFY]]` — current US/allied offtake & stockpile vehicles.
- [ ] §1.6 `[[VERIFY]]` — latest CRMA strategic-project designations.
- [ ] §1.7 `[[VERIFY]]` — confirm each company's stage; add missing names.
- [ ] §1.7 `[[MY VIEW]]` — which 2–3 names are actually investable for my mandate.

**Part 2 — Investment Views**
- [ ] §2.1 View 1 — thesis / WWHTBT / what-market-misses / what-makes-me-wrong.
- [ ] §2.1 View 2 — (same four).
- [ ] §2.1 View 3 — (same four).
- [ ] §2.1 View 4 — optional; (same four).
- [ ] §2.2 Single name — name & position.
- [ ] §2.2 — one-paragraph thesis.
- [ ] §2.2 — cap-structure map (+ `[[VERIFY]]` inputs).
- [ ] §2.2 — liquidity runway (+ `[[VERIFY]]` inputs).
- [ ] §2.2 — the specific catalyst.
- [ ] §2.2 — downside & recovery (+ `[[INPUT]]` downside deck).
- [ ] §2.2 — what the market is getting wrong (one sentence).

**Part 3 — Tailored framings**
- [ ] §3.1 `[[MY VIEW]]` — Readystate / credit + catalyst cut.
- [ ] §3.2 `[[MY VIEW]]` — BDT & MSD / owner's-lens cut.

**Sources**
- [ ] Date every cited source on final refresh (USGS, IEA, Adamas, MOFCOM, DoD/MP, CRMA, price refs).
