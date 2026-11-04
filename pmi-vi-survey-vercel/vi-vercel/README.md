# VI Rollout · Market Inventory survey (Vercel version)

Home with six category cards, a question menu per category, green "answered" states, % completion per category,
a confirmation pop-up on submit, and the filled Excel template delivered by email.

    public/            pages (market.html, admin.html, index.html) and static/app.js + style.css
    api/index.py       serverless backend (storage, validation, Excel export, email)
    template/          NEW_PMI_Market_Quantitative_Workbook.xlsx (the exact workbook that gets filled)
    vercel.json        routing (/m/<market>, /admin) and function settings
    dev.py             local preview with in-memory data (python3 dev.py)

## Deploy on Vercel (about 10 minutes)
1. **Upload the folder** to the existing project (assetsinventory): via GitHub (push this folder, connect the repo) or with the Vercel CLI (`vercel --prod` inside this folder).
   Framework preset: "Other". No build command is needed.
2. **Add the database**: Project → Storage → Marketplace → *Upstash Redis* → create (free plan is sufficient) → connect to the project.
   This injects `KV_REST_API_URL` and `KV_REST_API_TOKEN` automatically.
3. **Add environment variables** (Project → Settings → Environment Variables):

   | Variable | Value |
   |---|---|
   | `ADMIN_KEY` | a long random string; opens `/admin?a=<ADMIN_KEY>` |
   | `MAIL_TO` | recipient of submissions (comma-separated for several) |
   | `RESEND_API_KEY` | API key from resend.com (free plan); alternative below |
   | `MAIL_FROM` | optional, e.g. `VI Survey <survey@yourdomain.com>` (default: `onboarding@resend.dev`) |
   | `REQUIRE_COMPLETE` | optional; `true` = submit only when all touchpoints are answered (default `false`) |

   Resend's test sender (`onboarding@resend.dev`) delivers only to the email address the Resend account was created with.
   A verified domain in Resend removes that restriction. As an alternative to Resend, an SMTP mailbox can be used by setting
   `SMTP_HOST`, `SMTP_PORT`, `SMTP_USER`, `SMTP_PASS` (and `MAIL_FROM`) instead of `RESEND_API_KEY`.
4. **Redeploy** (Deployments → ⋯ → Redeploy) so the variables take effect.
5. **Switzerland exists from the first request**: `https://<project>.vercel.app/m/switzerland?k=swiss-preview-2026` (the key can be changed with the `SWITZERLAND_KEY` variable). Further markets: open `/admin?a=<ADMIN_KEY>`, create one per country (name + currency) and copy each market link.

## Notes
- Every submission is stored (answers snapshot) and downloadable from the admin page, even if email delivery is not set up or fails; the admin table shows the email status per submission.
- Contributors poll every 15 seconds; this keeps Redis usage low (a few commands per poll).
- Vercel's Hobby plan is meant for non-commercial use; a client project normally calls for the Pro plan.
- Data is held in an external service (Vercel + Upstash). PMI IT may prefer an approved environment; the same code structure can move to Azure.
- The Recap tab of the template still holds broken `#REF!` links from the previous structure; the export leaves it untouched.
