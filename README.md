# Frontier AI News Extractor
My work at OpenAI CoE constants requires me to look for current updates regarding models, feature updates, competitor news etc. and i found it difficult to track each and every website daily so i created a tool that can help me track it.

Tracks news from the frontier AI labs and keeps it in one Excel file. [Exa](https://exa.ai) reads each lab's official newsroom. **OpenAI GPT-6 Luna** filters, summarises and categorises each post. New stories are added on every run, and your own notes are kept.

## What you get

`frontier_news.xlsx` has these sheets:

**News**: one row per story, newest first, with filters and frozen headers.

| Column | What it holds |
|---|---|
| Date | When the news was published or announced |
| Company | OpenAI, Anthropic, Google DeepMind, DeepSeek, Mistral AI, … |
| Category | Model Launch · Improvement / Feature · Finances · Company · Partnership / M&A · Research · Policy / Legal · Infrastructure / Compute · Safety & Security |
| Headline | Short neutral headline |
| News Summary | 2–3 sentences: what happened, the key specifics, why it matters |
| Link | Original article (tracking parameters removed) |
| Source | e.g. `Anthropic (official)`, or the outlet name |
| Importance | High / Medium / Low (High is highlighted) |
| Models / Products | Names that are central to the story |
| Coverage, Other Sources | How many sources reported it, plus their links (when duplicates were merged) |
| First Seen | The run that found it |
| Read, Starred, Notes | **Yours.** These are kept on every run |

Categories are colour-coded and have dropdowns. Any extra column you add, such as "Priority", is kept too. If you delete a story, it doesn't come back.

**Dashboard**: a live count of stories by company × category, plus totals, High importance and last 7 days.
**Run Log**: one line per run with what was checked, added and merged.
**Seen URLs** (hidden): every URL already processed, so no article is paid for twice.

## Sources

| Company | Newsroom |
|---|---|
| OpenAI | https://openai.com/news/ |
| Anthropic | https://www.anthropic.com/news |
| Google DeepMind | https://deepmind.google/blog/ |
| DeepSeek | https://www.deepseek.com/en/news/ |
| Mistral AI | https://mistral.ai/news/ |

For each newsroom, the script:
1. reads the page through Exa and keeps the article links, matched against each site's URL pattern so navigation links are skipped
2. runs an Exa search limited to that company's domains, to catch posts that aren't on the first page

Add `--media` to also search news outlets such as TechCrunch and Reuters. That brings in funding and deal coverage that labs don't always post themselves. Duplicate coverage is merged into one row.

## Setup

```bash
# 1. install (uv or pip)
uv sync
# or: pip install -r requirements.txt

# 2. add your keys
cp .env.example .env      # then fill in EXA_API_KEY and OPENAI_API_KEY
```

## Usage

```bash
python news_extractor.py                          # last 7 days, all 5 newsrooms
python news_extractor.py --days 30                # longer window (good for a first run)
python news_extractor.py --sources OpenAI Anthropic
python news_extractor.py --media                  # also news-outlet coverage
python news_extractor.py --output ai_news.xlsx
```

With uv, run `uv run python news_extractor.py`.

Run it daily or weekly. Each run only fetches and analyses links it hasn't seen before.

**Schedule it**
- macOS/Linux (cron, every day at 08:00): `0 8 * * * cd /path/to/frontier-news-extractor && .venv/bin/python news_extractor.py`
- Windows: create a Task Scheduler task that runs `python news_extractor.py` in this folder.

Close the Excel file before a run. If it is open, results go to a timestamped copy instead.

## Customise

Everything is in the **CONFIGURATION** block at the top of `news_extractor.py`.

**Add a newsroom.** Append an entry to `SOURCES`:
```python
{
    "company": "xAI",
    "listing_url": "https://x.ai/news",
    "domains": ["x.ai"],
    "article_patterns": [r"x\.ai/news/[^/]+"],   # matched against "host/path" (no www, no trailing /)
    "aliases": ["Grok"],
},
```
Open the newsroom page and look at where its article links point. Put that pattern in `article_patterns`.

**Other settings**

| Setting | Default | Meaning |
|---|---|---|
| `CATEGORIES` | 9 categories | Names and the definitions the model follows |
| `NEWS_CRITERIA` | | What counts as news and what gets skipped |
| `IMPORTANCE_GUIDE` | | What High / Medium / Low mean |
| `OPENAI_MODEL` | `gpt-6-luna` | Can also be set in `.env` |
| `REASONING_EFFORT` | `low` | `none` is cheapest. Set it to an empty value for models without reasoning |
| `DAYS_BACK`, `MAX_LINKS_PER_SOURCE`, `RESULTS_PER_SEARCH` | 7, 25, 10 | How far back and how wide to look |
| `INCLUDE_MEDIA` | False | Turn news-outlet search on by default |

## How it works

1. **Collect.** Exa reads the newsroom pages, searches each company's domains and fetches the text of new article links.
2. **Pre-filter.** URLs already processed, articles outside the date window and empty pages are dropped before any model call.
3. **Analyse.** GPT-6 Luna reads articles in batches of 6 and returns structured JSON (strict schema): relevant?, company, category, headline, summary, date, importance, models. If a batch fails, it is split in half and retried.
4. **Merge duplicates.** A second short Luna call groups entries that describe the same event. One row is kept and the others are listed under *Other Sources*. If that call fails, headline similarity is used instead.
5. **Save.** New rows are merged with the existing file and sorted by date, and the dashboard and run log are rebuilt. The previous file is backed up to `backups/` (the last 5 are kept).

## Cost

- **Exa:** each run makes 5 newsroom page reads and 5 domain searches, plus one content fetch per *new* article link. The first run fetches up to ~125 links; later runs fetch only new ones. `--media` adds about 8 searches. See exa.ai/pricing.
- **OpenAI:** Luna is the low-cost GPT-6 tier. A run sends only new, in-window articles (about 1.5k tokens each), so after the first run that's usually a handful per day. See OpenAI's pricing page for current rates.

## Notes and troubleshooting

- **Dates.** Exa's published date is used, or the date the model reads from the article. On the very first run, articles with no date anywhere are skipped. After that, a link that newly appears on a newsroom page is treated as new and dated to the run.
- **`0 links on page` for a source.** The site may have blocked the page read or changed its layout. The domain search still covers that company. If it persists, check the source's `article_patterns`.
- **`ERROR from the OpenAI API`.** Check `OPENAI_API_KEY` and that your account can use `OPENAI_MODEL`. Nothing is written when this happens.
- **Re-checking skipped articles.** Delete the hidden *Seen URLs* sheet (right-click a tab → Unhide). The next run will look again at everything that isn't already a row in *News*.

## License

MIT
