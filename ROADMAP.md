# Email Archive Tool - Roadmap & Task Tracking

Development repository (source of truth, private):
https://github.com/nassertaher/Email-Archive-Tool-Dev

Public releases repository (stable only):
https://github.com/nassertaher/Email-Archive-Tool-1.0.0.0

## Working from two PCs (home and work)

The GitHub dev repo is the single source of truth. Both PCs use the same
account. Always `git pull` before starting and `git push` when done.

Remotes on every PC:

| Remote   | Repo                                   | Use                     |
|----------|----------------------------------------|-------------------------|
| origin   | Email-Archive-Tool-Dev (private)       | daily work + testing    |
| public   | Email-Archive-Tool-1.0.0.0 (public)    | publish releases only   |

First time on a second PC:

```
git clone https://github.com/nassertaher/Email-Archive-Tool-Dev.git
cd Email-Archive-Tool-Dev
git remote add public https://github.com/nassertaher/Email-Archive-Tool-1.0.0.0.git
python -m pip install -r requirements.txt
gh auth login          # only needed to publish releases
```

Daily loop:

```
git pull                # get the latest from the other PC
# ... work ...
git add -A
git commit -m "type: short description"
git push                # push to origin (dev)
```

Publish a tested build: `git push public main` then create a release on the
public repo (see README). Never commit personal files - `config.json`,
`token_cache.json`, logs and archives are git-ignored.

## Legend

- [ ] todo
- [~] in progress
- [x] done

## Phase 1 - Archive browser (MVP)

- [x] 1.1 Archive tab reads log.csv directly in the UI + Refresh
- [x] 1.2 Sortable columns (click a heading)
- [x] 1.3 Filters: free text, date range, agency, action, has-attachments
- [ ] 1.4 Detail pane: all fields + message body (from EML) + open PDF/EML/attachments

## Phase 2 - Register management

- [ ] 2.1 Edit Action / Response Date / Notes / Agency inline and save to log.csv + log.xlsx
- [ ] 2.2 Import any CSV into the Archive tab; merge rows into log.csv
- [ ] 2.3 Richer filters on the live Emails tab (date range, unread, has attachment, archived)

## Phase 3 - Insights

- [ ] 3.1 Dashboard: totals, by Agency, by month, outstanding actions
- [ ] 3.2 Export the filtered view to CSV / Excel

## Changelog

- 2026-10-10: Project opened as open source; v1.0.0.0 released publicly.
