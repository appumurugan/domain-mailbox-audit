# Domain & Mailbox Audit (GoDaddy × Apollo × Warmup Inbox)

A Google Apps Script that tells you **which of your cold-email domains still have mailboxes on them and which are safe to cancel**.

Outbound teams pile up secondary sending domains over time. Mailboxes get disconnected, domains get rotated out after deliverability problems, and nobody turns off auto-renew. This script matches every domain you own against every mailbox in your sending and warm-up tools. It refreshes a Google Sheet every hour.

## What it does

1. Pulls every domain from **GoDaddy**: status, expiry date and auto-renew setting.
2. Pulls every linked sending mailbox from **Apollo**.
3. Pulls every inbox from **Warmup Inbox**, with its warm-up status and score.
4. Checks each domain's **MX record**, to see whether email is actually set up on it (Google Workspace, Microsoft 365 or none).
5. Matches mailboxes to domains (subdomains included) and gives each domain a verdict:

| Verdict | Meaning |
|---|---|
| In use | At least one active Apollo mailbox |
| Review – only inactive Apollo mailboxes | Linked in Apollo but switched off |
| Review – warming only, not in Apollo | Only in Warmup Inbox so far |
| Review – email set up, no tracked mailbox | MX points to Google or Microsoft, but neither tool knows about a mailbox. Check your admin console. |
| **Terminate candidate** | No mailboxes and no email set up |
| Excluded (keep) | Domains you list in the Exclusions tab |
| Not active at GoDaddy | Expired, cancelled, etc. |

6. Keeps a running **mailbox registry**. New mailboxes are added automatically, and removed ones are marked "Removed" rather than deleted. Every change goes to a **Change Log** tab, with optional email alerts.

## Sheet tabs

- **Summary**: totals for each verdict, plus mailboxes on domains that aren't in GoDaddy
- **Domain Audit**: one colour-coded row per domain
- **Mailboxes**: one row per mailbox, with source, status, warm-up score, first seen and last seen
- **Exclusions**: domains that should never be flagged
- **Warmup Import**: optional backup; any email addresses pasted here are merged in
- **Change Log**: new or removed mailboxes and domains, verdict changes, errors

## Setup

1. Create a Google Sheet and open **Extensions → Apps Script**. Paste in `Code.gs` and save.
2. Under **Project Settings → Script Properties**, add:

   | Property | Value |
   |---|---|
   | `GODADDY_PAT` | GoDaddy Personal Access Token with domains read access. Or use `GODADDY_KEY` + `GODADDY_SECRET` for the legacy key pair. |
   | `APOLLO_API_KEY` | Apollo **master** API key, which the `email_accounts` endpoint requires |
   | `WARMUP_API_KEY` | Warmup Inbox API key |
   | `NOTIFY_EMAIL` | *(optional)* where to send change alerts |

3. Run these functions in order:
   - `setupSheets`: creates the tabs and asks you to authorise the script
   - add your keep-forever domains to **Exclusions**
   - `runAudit`: first full run
   - `installTrigger`: refreshes every hour
4. Reload the sheet. A **Domain Audit** menu appears for future runs.

No API keys are stored in the code. Everything lives in Script Properties.

## APIs used

- GoDaddy: `GET /v1/domains` (paginated with `marker`)
- Apollo: `GET /api/v1/email_accounts`
- Warmup Inbox: `GET /v1/inboxes` (`x-api-key` header)
- Google Public DNS: `https://dns.google/resolve?type=MX` (no key needed)

## Notes

- The first run doesn't log changes, because everything would count as "new". Logging starts from the second run.
- To cancel a domain at GoDaddy, you normally switch off auto-renew. The **Expires** column shows how long the domain stays live after that.
- The matching logic (`buildAudit_`) is a pure function, so it's easy to test outside Apps Script.

## License

MIT
