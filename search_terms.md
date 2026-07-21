# Market Intel — Search Term Reference

Reference vocabulary for the **middle group** of the Google News query in
`update_market_intel.py` (the `QUERY` constant). **Not yet wired into the program** —
this is the agreed term set for when we expand the search.

The query is an AND of three groups:

```
(fund term)  AND  (SECTOR TERM — this list)  AND  (action word)
```

Today the middle group is just `(infrastructure)`. The lists below expand it.

## Query mechanics (read before editing the query)

- **Quote every multi-word term.** Google reads a space as AND, so an unquoted
  `energy transition` breaks the OR-group logic. All multi-word entries below are shown
  quoted for this reason. Single words don't need quotes.
- **Google News caps results at ~100 per query.** A broad single OR-list does not return
  *more* than ~100 — it changes *which* ~100. High-volume renewables/energy coverage can
  crowd out a niche fund deal. Preferred architecture at build time: split into themed
  queries (renewables / digital / general) and merge + dedupe, rather than one mega-query.
- **Spelling variants are doubled on purpose** (`decarbonisation`/`decarbonization`,
  `fibre`/`fiber`, `"data centre"`/`"data center"`) — Google doesn't reliably treat these
  as equivalent inside quotes.

---

## In scope — sector vocabulary

### General infrastructure / real assets
`infrastructure` · `infra` · `"core infrastructure"` · `"real assets"` · `"essential infrastructure"` · `"infrastructure fund"` · `"infrastructure debt"`

### Renewables & energy transition
`renewable` · `renewables` · `"renewable energy"` · `"energy transition"` · `"clean energy"` · `"new energy"` · `solar` · `"solar farm"` · `wind` · `"offshore wind"` · `"onshore wind"` · `"wind farm"` · `hydro` · `hydropower` · `geothermal` · `hydrogen` · `"green hydrogen"` · `"battery storage"` · `"energy storage"` · `biomass` · `"carbon capture"` · `decarbonisation` · `decarbonization` · `"net zero"`

### Energy (conventional & midstream)
`energy` · `"power generation"` · `"power plant"` · `pipeline` · `pipelines` · `LNG` · `"natural gas"` · `nuclear` · `"small modular reactor"` · `midstream`

### Digital infrastructure
`"data centre"` · `"data center"` · `"digital infrastructure"` · `fibre` · `fiber` · `"fibre optic"` · `broadband` · `"telecom towers"` · `"cell towers"` · `5G` · `"subsea cable"` · `"AI infrastructure"`

### Utilities
`utility` · `utilities` · `"water utility"` · `wastewater` · `desalination` · `"district heating"` · `"electricity distribution"` · `"power grid"`

---

## Held in reserve (NOT in the query)

### Deal-structure bucket
`"project finance"` · `greenfield` · `brownfield` · `concession` · `PPP` · `"public-private partnership"` · `"co-investment"` · `"platform investment"` · `"take-private"` · `"equity stake"`

> Note: several of these are *deal-type* terms that overlap the query's third (action)
> group rather than the sector group — sort out placement at build time.

### Noisy singletons
bare `power` · bare `grid` · `port`

### Deleted categories
Transport and Social infrastructure buckets were removed by request.

---

## Notes

- `energy` (bare) is the broadest in-scope term and the real volume driver — it overlaps
  heavily with the renewables results. Kept intentionally.
- `hydro` carries mild regional noise (in Canada "hydro" often means the electricity bill).
- `solar` / `wind` (bare) were kept despite being noisy singletons, because they were
  explicitly requested; the AND-with-fund-term-and-action-word structure suppresses most
  of the astronomy/weather false positives.
