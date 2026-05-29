# Source Audit

Date: 2026-05-25

## Original Workbook Source Audit

This audit was run against the original multi-sheet workbook before the public repo was slimmed down. The public default workbook now keeps only the `Свод` sheet, because the dashboard and Excel importer do not need company-detail sheets. The old embedded source links were mostly aggregator/portal links, not official filings.

| Host | Cells | Unique URLs | Verdict |
| --- | ---: | ---: | --- |
| `uk.investing.com` | 855 | 55 | Poor automation source. Direct financial-statement pages returned Cloudflare `403`, and many cells point to `pro/pricing` rather than data. |
| `finance.yahoo.com` | 10 | 4 | Weak automation source. Direct Michelin financial URLs returned `404` in a machine fetch; Yahoo endpoints also tend to require cookies/crumbs and are not stable public APIs. |
| `marketscreener.com` | 9 | 9 | Weak automation source. Direct finance pages returned `Access Denied` in a machine fetch. |
| `bridgestone.com` | 1 | 1 | Good source family. Official IR pages are readable and publish PDFs/Excel files. |

## Better Source Hierarchy

1. Official structured regulatory APIs / XBRL where available.
2. Official investor-relations PDFs, Excel files, and HTML tables.
3. Paid or low-cost financial data APIs if the user accepts dependency/cost.
4. LLM extraction from official PDFs/HTML only for cases without structured data.
5. Aggregator web pages only as manual QA or fallback evidence, not as the quarter-end automation backbone.

## Company Notes

| Company group | Current source pattern | Better source pattern |
| --- | --- | --- |
| Goodyear | Investing.com in workbook | SEC EDGAR / SEC Company Facts JSON. Proven in `scripts/goodyear_sec_dry_run.py`. |
| Michelin, Continental, Pirelli, Nokian | Yahoo / Investing / MarketScreener | Official IR reports and ESEF/iXBRL where available. `filings.xbrl.org` can expose xBRL-JSON for many EU/UK ESEF filings, but the index is incomplete and Germany is called out as a missing-data country. |
| Bridgestone | Mostly Investing plus one official page | Official Bridgestone IR. The earnings library publishes quarterly PDFs and a Financial and Sales Data Excel file. |
| Yokohama, Sumitomo Rubber, Toyo | Investing / MarketScreener | Official IR results pages, TDnet/EDINET-style filings, official PDFs, and in some cases official HTML financial tables. |
| Apollo Tyres | Investing.com | Official Apollo investor page. It exposes financial results and Excel/PDF artifacts. NSE/BSE filings can be secondary official sources. |
| Hankook, Nexen, Kumho | Investing / MarketScreener | OpenDART is the best structured route for Korean listed companies; official IR pages/press releases are good LLM-readable fallback. |
| Zhongce, Sailun, Linglong | Investing / MarketScreener | Official company IR pages plus CNINFO / exchange disclosure PDFs. StockAnalysis is more readable than Investing for some China financial tables, but it is still an aggregator. |

## Practical Conclusion

Do not build the automation on the old workbook links as-is. They were useful clues for metric names and historical values, but the source registry should replace most of them with official sources and structured APIs. Gemini/DeepSeek/Qwen should consume downloaded official artifacts and return strict JSON; they should not browse Investing.com/MarketScreener as the primary data source.

## Latest-Period Selector Feasibility

Date: 2026-05-29

The current source package does not yet deterministically pick a named quarter/year for every issuer. It points AI Studio at official index pages, structured feeds, and fallback tables, then asks for the latest period. A more reliable next step is to parse each official index page first, select the exact period artifact, and pass only that artifact to the extractor.

| Company / group | Current observed pattern | Selector feasibility |
| --- | --- | --- |
| Goodyear | SEC Company Facts is structured. SEC submissions/filing indexes identify 10-K and 10-Q accessions by form, filing date, and period. | High. Use SEC submissions to select latest `10-K` or `10-Q`, then use Company Facts or the selected accession as evidence. |
| Michelin | Official `Results and sales` page is chronological and has clear sections such as `2026 first quarter sales`, `2025 annual results`, `2025 third quarter and 9 month sales`, with direct PDFs/XLSX links. | High for sales/annual result packs. Selector should parse headings and prefer `Financial information` / key-figures files for the requested period. |
| Continental | IR press page is date-sorted and exposes filenames like `20260506-pr-continental-q1-2026-en.pdf`; the site also has quarterly-publications/report surfaces. | High/medium. Use title/date/filename matching, but prefer a dedicated results/publications page over the broad press-release feed. |
| Pirelli | `Presentations and Webcasts` page lists year tabs and result blocks: `1Q 2026 Results`, `Full Year 2025 Preliminary Results`, `9M 2025`, `1H 2025`. | High. Parse blocks by period label and pick the `Report` / `Press Release` / `Presentation` links beside the block. |
| Nokian Tyres | Reports page is a chronological table with `Date`, `Subject`, `Category`, and direct materials: `Interim Report for January-March 2026`, `Annual Report 2025`, etc. | High. Select rows by category plus normalized period label. |
| Bridgestone | Earnings page has anchors by year and sections like `First Quarter`, `First Half`, `Third Quarter`, `Full Year`, with direct `Consolidated Financial Statements`, `Supplementary Information`, and Excel links. | High. This is one of the clearest candidates for deterministic year/quarter selection. |
| Sumitomo Rubber | Financial report page is grouped by year and currently shows links such as `Three Months Ended March 31, 2026`, `1Q Financial Report`. | High for latest quarter; annual/full-year selection should be confirmed against older year blocks. |
| Toyo Tires | Official pages are stable, but the fetched release page did not expose obvious 2026 result labels in the simple text snapshot. | Medium. Likely scriptable, but needs a deeper page-specific parser or alternate official financial tables. |
| Apollo Tyres | Financial reporting page exposes `Choose Year` controls and lists `Annual Report FY2025`, `Audited Financial Results - March 31, 2026`, and quarterly/half-year result rows by date. | High. The selector can map fiscal year ranges and quarter-end dates. |
| Hankook / Nexen / Kumho | Official IR pages exist, and OpenDART is the best structured source but needs an API key/corp-code workflow. | Medium/high with OpenDART; medium without it. Add corp codes and use OpenDART for deterministic annual/quarterly reports. |
| Zhongce / Sailun / Linglong | Generic CNINFO homepage is not enough; company-specific stock codes/org IDs and disclosure category filters are needed. | Medium after registry enrichment. Use CNINFO/search APIs or company-specific disclosure URLs, not the generic homepage. |

Conclusion: yes, deterministic period selection is the better architecture. It should be implemented as a source-discovery layer that outputs `selected_period`, `selected_source_url`, `source_document_type`, and confidence before AI Studio runs. Do not rely on hard-coded URL formulas unless a site proves a stable pattern; parsing the official index page is safer because file slugs and CDN paths change.
