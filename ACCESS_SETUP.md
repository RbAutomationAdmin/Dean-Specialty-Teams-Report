# Locking the dashboard to @rbacentralnj.com — Cloudflare Access runbook

Goal: only people who can receive email at an `@rbacentralnj.com` address can open
the dashboard. Enforced by Cloudflare Access (free for up to 50 users) sitting in
front of the Netlify site. Users verify with a one-time PIN sent to their work
email — no passwords to manage, nothing to install.

Why this needs a custom domain: Cloudflare Access can only protect hostnames whose
DNS runs through Cloudflare. The `specialtyteamsreport.netlify.app` URL can never be
protected directly.

## Step 0 — Pick the domain (one-time decision)

**Ask IT first:** is `rbacentralnj.com` already on Cloudflare? If yes, this whole
setup is a 10-minute favor from whoever manages it (they do steps 1–4 on a
subdomain like `specialtyreport.rbacentralnj.com`, you write the Access policy in
step 5 or hand them this doc).

If not (or you don't want to involve IT): buy a standalone domain in Cloudflare
(dash.cloudflare.com → Domain Registration → Register domain, ~$10/yr, e.g.
`rbaspecialtyreport.com`). Do NOT try to move `rbacentralnj.com` itself onto
Cloudflare — its DNS carries the company's Microsoft 365 mail records.

## Step 1 — DNS record (Cloudflare → your domain → DNS)

| Field | Value |
|---|---|
| Type | CNAME |
| Name | `dashboard` (gives you dashboard.yourdomain.com) or `@` for the bare domain |
| Target | `specialtyteamsreport.netlify.app` |
| Proxy status | **DNS only (grey cloud)** for now |

## Step 2 — Tell Netlify about the domain

Netlify → the specialtyteamsreport site → Domain management → **Add custom
domain** → enter `dashboard.<yourdomain>` → **set it as primary domain**.
Wait until the SSL/TLS certificate section shows the cert as issued (usually a
few minutes; requires the grey cloud from step 1).

Setting it primary makes `specialtyteamsreport.netlify.app` permanently redirect
to the custom domain, which closes the "just use the netlify.app link" bypass.

## Step 3 — Turn on the Cloudflare proxy

Back in Cloudflare DNS: flip the record from step 1 to **Proxied (orange cloud)**.
Then under SSL/TLS → Overview, set encryption mode to **Full**.

## Step 4 — Create the Access application

Cloudflare dashboard → **Zero Trust** (first visit asks you to pick a free team
name — anything works) → Access → Applications → **Add an application** →
**Self-hosted**:

- Application name: `Specialty Teams Dashboard`
- Session duration: `1 week` (re-verify weekly; pick shorter if preferred)
- Application domain: `dashboard.<yourdomain>` (path blank = protect everything)

## Step 5 — The policy (this is the actual lock)

Add a policy on that application:

- Policy name: `RBA staff only`
- Action: **Allow**
- Include → selector **Emails ending in** → value `@rbacentralnj.com`

Leave login methods at the default **One-time PIN**. Save.

## Step 6 — Verify (2 minutes, do all three)

1. Incognito window → `https://dashboard.<yourdomain>` → should show a Cloudflare
   login page, NOT the dashboard. Enter your work email → PIN arrives → dashboard.
2. Same flow with a personal email (gmail etc.) → must be rejected before any
   PIN is sent.
3. `https://specialtyteamsreport.netlify.app` → must redirect to the custom
   domain (which then demands login).

## Ongoing

- Nothing changes about the refresh workflow — push to main, Netlify deploys,
  Cloudflare passes traffic through Access.
- New hires need nothing: any @rbacentralnj.com address passes the rule.
- To later swap PIN login for full Microsoft SSO: Zero Trust → Settings →
  Authentication → add Azure AD / Entra ID as a login method (optional polish,
  the PIN gate is already domain-enforced).
