# Launch Plan (Phases 13–25)

Rule from the master plan: **do not market before the workbook is reliable.** Gates G1–G3 must be green before any public listing.

## Release gates

| Gate | Requirement | Status |
|---|---|---|
| G1 Accounting QA | Automated suite green (reference model, 35 scenarios, validation, edge cases, randomized books) | ✅ automated, every build |
| G2 Compatibility | Manual pass in Excel 365 Windows + Excel for Mac (+ Excel Online). Google Sheets tested before any claim | ⏳ needs a machine with Excel |
| G3 UAT | 3 non-developers (owner, bookkeeper, Excel user) complete the 10 UAT tasks without help (`qa-plan.md §5`) | ⏳ |
| G4 Brand | Name trademark/domain check, logo, cover image | ⏳ |
| G5 Commerce | Store page live, payment + download tested end-to-end, refund policy, support inbox | ⏳ |

## Tiers & pricing hypothesis

| Tier | Contents | Price hypothesis | Role |
|---|---|---|---|
| Free | Income & expense tracker + mini dashboard | $0 (email) | Lead magnet, trust |
| Standard | This workbook + sample file + guides | **$24** (list $39 · launch $19) | Core offer |
| Premium (v1.2+) | + budget vs actual, forecast, ratios, inventory | $49 | Upsell |
| Bundle | Standard + Premium add-ons + future family products | $69 | AOV |

Market evidence: paid templates cluster at $10–$30; software costs about $456/yr (`market-research.md`). Track views → add-to-cart → purchase, conversion %, refunds and support tickets per 100 sales, and change one variable at a time.

## Channels (in order)

1. **Etsy.** Search-driven demand, small-business buyers. Listing copy: `marketing/listing-copy/etsy.md`.
2. **Payhip or Lemon Squeezy** for direct sales and email capture. (Gumroad is an option, but note recent creator payout complaints in public reviews. Evaluate before committing.)
3. **Own site** later, for SEO, bundles and support.

## Launch sequence

| Week | Activities |
|---|---|
| −2 | Close G2/G3, fix findings, screenshots (`tools/` can render PDFs/PNGs of every sheet from the sample file), 60-s demo video |
| −1 | Soft launch to 10–20 owners from your network and small-business groups; collect bugs, testimonials, wording |
| 0 | Public listing, email announcement, 5–7 social posts, 3–5 short videos, 1 long tutorial |
| +1..+4 | Weekly: sales, conversion, questions → docs or UX fixes; monthly review per master plan §56 |

## Package contents (Phase 15)

```text
SimpleLedger-Small-Business-Accounting/
├── Excel/SimpleLedger-Small-Business-Accounting.xlsx   (blank, ready to use)
├── Examples/SimpleLedger-Sample-Business-Data.xlsx      (Sunrise Bakery demo)
├── Guides/Quick-Start-Guide.pdf, User-Guide.pdf         (from docs/user-guide.md)
├── License/License-and-Terms.txt
└── Read-Me.txt
```

`python tools/package.py` assembles this folder and a zip from `dist/` (PDF guides are added once the user guide is final).

## Marketing assets (drafts in `marketing/`)

- `listing-copy/etsy.md`: title, tags, description, FAQ
- `email/sequence.md`: 6-email lead-magnet funnel
- `social/calendar.md`: weekly content system + 20 post ideas
- `video/scripts.md`: 3 short-form scripts (hook → problem → workbook → result → CTA)
