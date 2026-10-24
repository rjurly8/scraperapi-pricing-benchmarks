# scrape website to excel: Which Route Fits Your Target Site — Power Query, No-Code Tools, or the ScraperAPI Route, with Real Pricing, Credit Math, and Success-Rate Data

Getting website data into an Excel spreadsheet comes down to three practical routes: Excel's built-in Power Query for static public tables, a no-code scraper you configure by clicking, or an API like ScraperAPI that renders pages and handles anti-bot measures server-side before you write the results to `.xlsx` with a few lines of Python. The right choice depends on whether your target site is static or JavaScript-heavy, whether it sits behind protection like Cloudflare, and whether you need a one-off pull or a recurring pipeline. This article walks through all three, then goes deep on the part most tutorials skip: what ScraperAPI actually charges, how its credit system drains faster on certain sites, and which targets it handles well according to independent benchmark data.

> **Quick verdict:** static public table, one-off job → Excel Power Query, costs nothing. Logged-in pages, heavy pagination, bot protection, or scheduled runs → ScraperAPI's free plan gives you 1,000 credits per month with 5 concurrent connections, plus 5,000 free requests during your first 7 days, no credit card required, and a 7-day no-questions-asked refund if you pay and change your mind. 👉 [Grab the free credits and test your target site first](https://www.scraperapi.com/signup?fp_ref=coupons)

## Three routes, and where each one breaks

Power Query, built into Excel 2016 and later, is the most underrated scraping tool. Under Data → Get Data → From Web, paste a URL, and Excel parses the HTML tables on the page. The import is a live query: right-click to refresh, or set an automatic refresh interval under Data → Queries & Connections → Properties. For a public, static table — a government statistics page, a Wikipedia list — this takes five minutes and costs nothing.

The limitation is equally clear. Power Query reads the raw HTML the server sends on first load. No browser rendering, no login handling, no anti-bot logic. On JavaScript-rendered pages (most modern e-commerce, real estate, and job sites), it returns an empty result. Google Sheets' IMPORTHTML and IMPORTXML formulas hit the same wall, and Google's shared import IPs are blocked by many sites by default.

No-code tools — Octoparse, Browse AI, and browser extensions — solve the "I can't code" problem. You click the fields you want, the tool generates extraction rules, and you export to Excel or CSV. The trade-off: expect roughly 30–60 minutes of setup per site, retraining after every redesign, and scheduling or proxy features locked behind paid cloud tiers.

| Method | Setup | JS-rendered pages | Anti-bot sites | Scheduled runs | Cost |
| --- | --- | --- | --- | --- | --- |
| Excel Power Query | Minutes | No | No | Only while the workbook is open | Included with Excel |
| Google Sheets IMPORTHTML/XML | Minutes | No | No | Rough hourly recalc | Free |
| No-code tools | 30–60 min per site | Yes | Partial (paid proxies) | On paid cloud plans | Subscription |
| Python + scraping API | ~15 min once | Yes | Yes | cron / cloud jobs | Free tier + usage |

ScraperAPI sits in the last row. You send an API request with your target URL; it rotates proxies, renders JavaScript, solves CAPTCHAs, and retries failed requests, then returns the final HTML or a normalized JSON. Everything after that is just writing a file.

## The API route in practice

Three steps: get your API key, make the request, write the spreadsheet.

python
import requests, pandas as pd
from bs4 import BeautifulSoup

resp = requests.get(
    "https://api.scraperapi.com/",
    params={"api_key": "YOUR_KEY", "url": "https://example.com/pricing"}
)
soup = BeautifulSoup(resp.text, "html.parser")
rows = [[c.get_text(strip=True) for c in tr.find_all(["th", "td"])]
        for tr in soup.select("table tbody tr")]
pd.DataFrame(rows).to_excel("output.xlsx", index=False)


For Amazon, Google, and Walmart, there's a shortcut: ScraperAPI includes structured data APIs on every plan, returning ready-parsed JSON fields like title, price, and rating. No selectors to write or maintain — `pd.DataFrame(data).to_excel(...)` finishes the job. When the target site redesigns its layout, the normalized fields keep coming, which is the main durability advantage over selector-based methods.

## ScraperAPI's full plan lineup

The official pricing page currently shows seven paid tiers plus a free plan, with 10% off for annual billing. Every paid plan includes JS rendering, premium proxies, JSON auto parsing, rotating proxy pools, custom headers, CAPTCHA and anti-bot detection, custom sessions, automatic retries, unlimited bandwidth, and a 99.9% uptime guarantee.

| Plan | Monthly price | Annual (per month) | Credits/month | Concurrent threads | Get started |
| --- | --- | --- | --- | --- | --- |
| Free | $0 | — | 1,000 | 5 | [Sign up free](https://www.scraperapi.com/signup?fp_ref=coupons) |
| Hobby | $49 | $44.10 | 100,000 | 20 | [View Hobby plan](https://www.scraperapi.com/pricing?fp_ref=coupons) |
| Startup | $149 | $134.10 | 1,000,000 | 50 | [View Startup plan](https://www.scraperapi.com/pricing?fp_ref=coupons) |
| Business | $299 | $269.10 | 3,000,000 | 100 | [View Business plan](https://www.scraperapi.com/pricing?fp_ref=coupons) |
| Scaling (most popular) | $475 | $427.50 | 5,000,000 | 200 | [View Scaling plan](https://www.scraperapi.com/pricing?fp_ref=coupons) |
| Professional | $975 | $877.50 | 10,500,000 | 300 | [View Professional plan](https://www.scraperapi.com/pricing?fp_ref=coupons) |
| Advanced | $1,975 | $1,777.50 | 21,500,000 | 500 | [View Advanced plan](https://www.scraperapi.com/pricing?fp_ref=coupons) |
| Enterprise | Custom | — | 22,000,000+ | 500+ | [Contact sales](https://www.scraperapi.com/contact-sales/?fp_ref=coupons) |

Two limits worth noting before you pick a tier: geotargeting on Hobby and Startup covers US and EU only, with country-level targeting starting at Business (ScraperAPI's own comparison page puts it at 30+ countries, included in the price). The analytics dashboard keeps only the last 30 days of data on the three lower tiers.

## Credit math: the number that actually determines your bill

A standard page costs 1 credit, but several common targets cost far more — and this is where budget estimates fall apart. Per ScraperAPI's pricing page: Amazon costs 5 credits per request, Google and Bing (including subdomains) cost 25, LinkedIn costs 30, and sites behind Cloudflare, Datadome, or PerimeterX protection add 10 credits per request when bypassed.

Run the numbers on a $49 Hobby plan with 100,000 credits:

- Generic pages: about 100,000 requests
- Amazon pages: about 20,000
- Google result pages: about 4,000
- LinkedIn pages: about 3,300

So "how many rows can I pull into Excel this month" has four very different answers depending on your source. Before committing to a plan, use the Domain Cost Estimator in the dashboard to check the exact cost of any URL, and set a `max_cost` per scrape to cap what any single request can burn. Billing is request-based only — there is no bandwidth-based pricing option.

👉 [Check your target site's exact credit cost on the pricing page](https://www.scraperapi.com/pricing?fp_ref=coupons)

## What independent benchmarks actually show

ScraperAPI's marketing pages don't publish success rates, so third-party benchmarks are the honest yardstick. Scrapeway, which runs a recurring panel of scraping APIs against fixed real-world targets, put ScraperAPI at a 52% overall success rate across 11 target sites in its September window, ranking 4th of 7 services tested, at an average of 7.5 seconds per request and $4.05 per 1,000 requests. For context, Scrapfly led that same panel at 98% success, while ScrapingBee and ZenRows scored 33% and 22% respectively.

The per-target breakdown matters more for Excel work than the average:

- Etsy 96%, Zillow 94%, Amazon 90%, Walmart 84% — reliable sources
- LinkedIn 58%, Indeed 28% — expect failures and retries
- Booking.com 1%, Twitter and Instagram 0% — effectively non-starters

One caveat: these numbers move between test windows. An earlier Scrapeway window the same year had Amazon at 96.4% and LinkedIn at 93%, so month-to-month variance is real. The practical takeaway is unchanged — if your data source is Amazon, Walmart, Etsy, or Zillow, the success rate justifies the cost; if it's Booking or Instagram, don't spend credits there. Verify your specific target with the free tier before paying for anything. On Capterra, ScraperAPI holds a 4.6 rating from 62 reviews.

## Rules hidden in the billing docs that affect real projects

Most reviews stop at the price table. These details from ScraperAPI's own documentation change how you'd use it:

- **Credits don't roll over.** Unused credits expire at each billing cycle.
- **Running out pauses you by default.** Scraping stops until your credits renew — no surprise charges.
- **Auto-upgrade is opt-in** on Hobby through Professional: when credits run out, your account moves to the next tier automatically, usually at a better per-credit rate.
- **Pay-as-you-go exists on Scaling, Professional, Advanced, and Enterprise**, letting you keep scraping past your limit at a fixed per-credit rate up to a monthly spending cap you set. The cap is a limit, not a pre-charge.
- **Concurrency is hard-capped.** Exceeding your thread limit returns a 429 and the request is rejected rather than queued — but there's no extra charge for going over.
- **Cancellation is free anytime** from the dashboard, and card details can be changed at any point.

The $49 Hobby tier is a real annual commitment if you bill yearly: $529.20 for twelve months versus $588 billed monthly. The $58.80 difference is small; the real reason to prefer monthly first is that your first months are when you find out whether your target site behaves like the benchmark numbers.

## When you don't need any of this

If your source is a static public table and you need the data once, Power Query already does the job for free, and no API subscription makes it better. Crawlora's comparison of the no-code route is a useful benchmark here: its own free tier offers 2,000 credits per month with pay-on-success billing, which undercuts ScraperAPI's 1,000 free credits on volume — though ScraperAPI's included structured-data endpoints for Amazon, Google, and Walmart mean you often spend fewer credits per useful row. ScraperAPI's value concentrates exactly where built-in tools fail: JavaScript rendering, dozens or hundreds of paginated pages, bot protection, and unattended scheduled runs feeding a spreadsheet.

## Trial, refund, and what happens after you subscribe

The risk-reduction terms are straightforward and come from official pages: 5,000 free requests during your first 7 days with no credit card, a 7-day no-questions-asked refund policy after you pay, free cancellation at any time, and the default pause-when-empty behavior described above. Between the free week and the refund window, you can realistically validate a full pipeline — render, parse, write to Excel — before any money is at stake. 👉 [Start the 7-day trial with 5,000 free requests](https://www.scraperapi.com/signup?fp_ref=coupons)

## FAQ

**Does ScraperAPI output .xlsx files directly?**
No. It returns HTML or JSON; you write the spreadsheet yourself with `pandas.to_excel()` (real `.xlsx`, with proper number and text types) or paste JSON into Power Query and let it expand into rows.

**Will I get charged unexpectedly if I exceed my credits?**
No. Scraping pauses by default until renewal. Extra charges only occur if you enable auto-upgrade or pay-as-you-go yourself, and the latter is bounded by a spending cap you configure.

**Can it run on a schedule and fill Excel automatically?**
Yes, two ways: DataPipeline handles scheduled scraping and delivery (it consumes extra credits), or run your own cron job calling the API and writing the `.xlsx` — no added cost.

**What's the row limit in Excel?**
A single worksheet holds 1,048,576 rows. That's Excel's own ceiling, not the scraper's; larger datasets should be split across sheets or kept in CSV.

## Final picks by situation

Static public table → Power Query, and stop reading. One or two dynamic pages, zero code → a no-code tool, and budget for retraining after redesigns. Amazon, Walmart, Etsy, or Zillow as your source, or a scheduled pipeline → validate with the free credits, then start at Hobby ($49, 100,000 credits) and scale only if your credit math says so. Booking.com, Twitter, or Instagram as your source → this API won't get you there; look for a specialized approach. If the numbers above match your project, 👉 [sign up for ScraperAPI and put your free 5,000 requests to work](https://www.scraperapi.com/signup?fp_ref=coupons).
