---
title: Deploying the Example Service
excerpt: End-to-end deployment procedure for a sample service, including rollback and the emergency path.
author: gareth-cheyne
tags: [deployment, runbook, example]
category: operations
date: 2026-09-03
draft: false
restricted: true
restrictedRoles:
  - Developers
  - Platform-Ops
---

A realistic NextDocs page, copied out of a documentation repository and converted
as-is. Nothing here was written for the converter — this is what an ordinary
NextDocs document looks like, and it is the case worth testing.

Convert it two ways to see the difference:

```powershell
Convert-Markdown .\nextdocs-page.md -Format Pdf
Convert-Markdown .\nextdocs-page.md -Format Pdf -ExcludeRestricted -Force
```

## Before you start

> [!IMPORTANT]
> Deployments outside the change window need an approved emergency CR. The
> platform on-call must be told before the drain begins, not after.

> [!CAUTION]
> Rolling back after the schema migration has run needs the database restore
> procedure, not the rollback endpoint below.

## Prerequisites

| Requirement | Check |
| --- | --- |
| Pipeline green | Build and integration stages both passed |
| Change record | Approved, and inside the window |
| Artefact version | Matches the version named in the CR |
| Secret store access | You can read `example-service/production` |

## Procedure

:::steps
### Verify the build
Confirm the pipeline is green and the artefact version matches the change record.
A mismatch here is the single most common cause of a failed deployment.

### Announce the drain
Post in the team channel. The on-call engineer acknowledges before you
continue.

### Drain the node pool
Run the drain script and wait for active connections to reach zero. This
typically takes two to three minutes.

### Deploy
Roll the new artefact forward one availability zone at a time. Wait for the
health endpoint to return :badge[stable] in each zone before starting the next.

### Verify
Check the dashboard for error rate, latency and queue depth. Watch for a full
ten minutes before closing the change.
:::

## Rollback

:::api POST /admin/rollback
Roll the service back to the previously deployed artefact.
---
**Body:**

- `version` (string) — the artefact version to restore
- `reason` (string) — free text, written to the audit log

**Response:** `202 Accepted`, with a job id to poll.
:::

:::api GET /admin/rollback/:jobId
Poll the status of a rollback job.
:::

## Configuration change

The health check timeout was raised in this release. If you are rolling back to
an artefact older than `2026.08.01`, revert the config too:

:::diff yaml
filename: example-service/production/values.yaml
---
  livenessProbe:
    httpGet:
      path: /healthz
-     timeoutSeconds: 2
+     timeoutSeconds: 5
    periodSeconds: 10
:::

## Where things live

:::filetree
- example-service/
  - production/
    - values.yaml
    - secrets.sops.yaml
  - staging/
    - values.yaml
- scripts/
  - drain.ps1
  - rollback.ps1
:::

## Environment differences

:::tabs
@tab Production
Three availability zones, blue/green, manual approval gate. Drain takes two to
three minutes per zone.

@tab Staging
Single zone, rolling update, no approval gate. Safe to redeploy freely.

@tab Local
`docker compose up`. No drain step and no health gate — the compose file starts
the dependencies for you.
:::

## Credentials

!variant!# Platform-Ops
Production credentials live in the operations secret store under `example-service/production`.
Rotation is quarterly and tracked in the operations calendar. The rotation runbook
is linked from the vault entry itself.

Break-glass access needs two approvers and is logged to the security channel
automatically.
!endvariant!

Everyone else should raise a request through the service desk, not in the
team channel.

## Troubleshooting

:::details Health check fails immediately after deploy
Almost always the timeout, not the service. Check `timeoutSeconds` in the values
file for the environment you deployed to — see the diff above.
:::

:::details Drain never reaches zero connections
A long-lived websocket client is holding the connection open. The drain script
has a `-Force` switch that closes them after the grace period; use it only once
you have confirmed no batch job is running.
:::

## Rollout history

:::release
teams: Platform, Documentation
version: 2026.09.03.1
---
### Added
- Zone-by-zone health gating during rollout.

### Changed
- Health check timeout raised from 2s to 5s.

### Fixed
- Drain script no longer exits before the last connection closes.
:::
