2026-08-18 15:24:24 | === Macro started (v3.5) ===
2026-08-18 15:24:24 | RECOVER: stale OSINT state detected on entry - force-clearing (calc/events/alerts/cursor/screen/print/statusbar)
2026-08-18 15:24:24 | STEP 2: before search-mode prompt
2026-08-18 15:24:26 | MODE: search mode set to Fast (USE_HEADLESS=True, jitter=2-3s, protections=False)
2026-08-18 15:24:26 | STEP 2: search-mode prompt returned mode=Fast (USE_HEADLESS=True)
2026-08-18 15:24:26 | STEP 2: before AutoPrewarmDNS
2026-08-18 15:24:27 | STEP 2: after AutoPrewarmDNS
2026-08-18 15:24:27 | STEP 2: before PrewarmEdgeProfiles
2026-08-18 15:24:27 | PREWARM: fired 2 Edge profile warmer(s) - running in background while user is prompted
2026-08-18 15:24:27 | STEP 2: after PrewarmEdgeProfiles
2026-08-18 15:24:27 | STEP 3: before customer negnews InputBox
2026-08-18 15:24:30 | STEP 3: customer negnews returned, len=15
2026-08-18 15:24:30 |   NormalizeSpaces entry, inLen=15
2026-08-18 15:24:30 |   NormalizeSpaces exit, outLen=15 iter=0
2026-08-18 15:24:30 | STEP 3: after NormalizeSpaces, len=15
2026-08-18 15:24:30 | STEP 3: before customer legal-extension InputBox
2026-08-18 15:24:35 |   NormalizeSpaces entry, inLen=11
2026-08-18 15:24:35 |   NormalizeSpaces exit, outLen=11 iter=0
2026-08-18 15:24:35 | STEP 3: before Sigma address MsgBox
2026-08-18 15:24:38 | STEP 3: Sigma TopMostMsgBox returned 7
2026-08-18 15:24:38 | STEP 3: customer prompts complete
2026-08-18 15:24:41 |   NormalizeSpaces entry, inLen=11
2026-08-18 15:24:41 |   NormalizeSpaces exit, outLen=11 iter=0
2026-08-18 15:24:43 |   NormalizeSpaces entry, inLen=7
2026-08-18 15:24:43 |   NormalizeSpaces exit, outLen=7 iter=0
2026-08-18 15:24:45 |   NormalizeSpaces entry, inLen=26
2026-08-18 15:24:45 |   NormalizeSpaces exit, outLen=26 iter=0
2026-08-18 15:24:52 |   NormalizeSpaces entry, inLen=0
2026-08-18 15:24:52 |   NormalizeSpaces exit, outLen=0 iter=0
2026-08-18 15:24:52 | STEP 3: before Starting Searches MsgBox
2026-08-18 15:24:55 | STEP 3: Starting Searches MsgBox dismissed -> proceeding to searches
2026-08-18 15:24:59 | BG TUNE: lowered 22 background process(es) to BelowNormal
2026-08-18 15:24:59 | SEARCH PHASE start - total entities to process: 3
2026-08-18 15:24:59 | SORT: reordered 19 tasks (heavy first) - heaviest weight=10, lightest=1
2026-08-18 15:24:59 | SEARCH PHASE - queued 19 tasks across all entities (heaviest first)
2026-08-18 15:24:59 | BATCH ALL start, n=19, maxParallel=2
2026-08-18 15:24:59 | PREWARM: drain done in 0.60s (allDone=True)
2026-08-18 15:25:00 | COOKIE WARM: fired 2 HEADLESS (hidden) warm-up(s) to pick up Google session cookies
2026-08-18 15:25:08 | COOKIE WARM: done after 8s (headless)
2026-08-18 15:25:08 | BATCH ALL: cookie warm-up RAN before main pass
2026-08-18 15:25:08 | PASS [main] start, n=19, maxRetries=7
2026-08-18 15:25:08 |   TASK launch [1/19] w0 running=1 [Customer (Currentware INC)] Negative News page 1
2026-08-18 15:25:08 |   TASK launch [2/19] w1 running=2 [Customer (Currentware INC)] Negative News page 2
2026-08-18 15:25:13 |   WORKER w0 [Fast] cooldown 2.1s before next task
2026-08-18 15:25:13 |   TASK retry [1/19] CAPTCHA-size 52410B - requeued (attempt 2/8), backing off 0.0s before relaunch
2026-08-18 15:25:13 |   WORKER w1 [Fast] cooldown 3.0s before next task
2026-08-18 15:25:13 |   TASK done  [2/19] w1 wait=0s edgeStart=5s render=0s dur=5s size=201134B running=0
2026-08-18 15:25:15 |   TASK launch [1/19] w0 running=1 [Customer (Currentware INC)] Negative News page 1
2026-08-18 15:25:15 |   TASK launch [3/19] w1 running=2 [Customer (Currentware INC)] Negative News w/o legal ext, page 1
2026-08-18 15:25:21 |   WORKER w0 [Fast] cooldown 2.5s before next task
2026-08-18 15:25:21 |   TASK retry [1/19] CAPTCHA-size 52410B - requeued (attempt 3/8), backing off 0.0s before relaunch
2026-08-18 15:25:22 |   WORKER w1 [Fast] cooldown 2.6s before next task
2026-08-18 15:25:22 |   TASK done  [3/19] w1 wait=7s edgeStart=7s render=0s dur=7s size=212909B running=0
2026-08-18 15:25:23 |   TASK launch [1/19] w0 running=1 [Customer (Currentware INC)] Negative News page 1
2026-08-18 15:25:24 |   TASK launch [4/19] w1 running=2 [Customer (Currentware INC)] Negative News w/o legal ext, page 2
2026-08-18 15:25:30 |   WORKER w1 [Fast] cooldown 2.0s before next task
2026-08-18 15:25:33 |   KILL w1 - terminated 16 orphan msedge.exe process(es)
2026-08-18 15:25:33 |   TASK done  [4/19] w1 wait=16s edgeStart=6s render=3s dur=9s size=223811B running=1
2026-08-18 15:25:34 |   TASK launch [5/19] w1 running=2 [CP 1 of 2 (Cadable LLC)] Negative News page 1
2026-08-18 15:25:40 |   WORKER w0 [Fast] cooldown 2.4s before next task
2026-08-18 15:25:40 |   TASK done  [1/19] w0 wait=15s edgeStart=17s render=0s dur=17s size=275369B running=1
2026-08-18 15:25:42 |   TASK launch [6/19] w0 running=2 [CP 1 of 2 (Cadable LLC)] Negative News page 2
2026-08-18 15:25:51 |   WORKER w0 [Fast] cooldown 2.9s before next task
2026-08-18 15:25:51 |   TASK done  [6/19] w0 wait=34s edgeStart=9s render=0s dur=9s size=151450B running=1
2026-08-18 15:25:53 |   TASK launch [7/19] w0 running=2 [CP 1 of 2 (Cadable LLC)] Negative News w/o legal ext, page 1
2026-08-18 15:25:59 |   WORKER w0 [Fast] cooldown 2.4s before next task
2026-08-18 15:26:01 |   KILL w0 - terminated 16 orphan msedge.exe process(es)
2026-08-18 15:26:01 |   TASK done  [7/19] w0 wait=45s edgeStart=5s render=3s dur=8s size=247186B running=1
2026-08-18 15:26:01 |   TASK launch [8/19] w0 running=2 [CP 1 of 2 (Cadable LLC)] Negative News w/o legal ext, page 2
2026-08-18 15:26:10 |   WORKER w1 [Fast] cooldown 2.8s before next task
2026-08-18 15:26:10 |   TASK retry [5/19] TIMEOUT after 35s (NO FILE EVER) - requeued (attempt 2)
2026-08-18 15:26:12 |   TASK launch [5/19] w1 running=2 [CP 1 of 2 (Cadable LLC)] Negative News page 1
2026-08-18 15:26:18 |   WORKER w1 [Fast] cooldown 2.8s before next task
2026-08-18 15:26:18 |   TASK done  [5/19] w1 wait=64s edgeStart=6s render=0s dur=6s size=309079B running=1
2026-08-18 15:26:20 |   TASK launch [9/19] w1 running=2 [CP 2 of 2 (First National Bank Cortez)] Negative News page 1
2026-08-18 15:26:26 |   WORKER w1 [Fast] cooldown 2.5s before next task
2026-08-18 15:26:26 |   TASK done  [9/19] w1 wait=72s edgeStart=6s render=0s dur=6s size=198531B running=1
2026-08-18 15:26:28 |   TASK launch [10/19] w1 running=2 [CP 2 of 2 (First National Bank Cortez)] Negative News page 2
2026-08-18 15:26:36 |   WORKER w1 [Fast] cooldown 2.2s before next task
2026-08-18 15:26:39 |   KILL w1 - terminated 15 orphan msedge.exe process(es)
2026-08-18 15:26:39 |   TASK done  [10/19] w1 wait=80s edgeStart=8s render=3s dur=11s size=212916B running=1
2026-08-18 15:26:39 |   TASK launch [11/19] w1 running=2 [Customer (Currentware INC)] Name + Address
2026-08-18 15:26:39 |   WORKER w0 [Fast] cooldown 2.9s before next task
2026-08-18 15:26:39 |   TASK retry [8/19] TIMEOUT after 35s (NO FILE EVER) - requeued (attempt 2)
2026-08-18 15:26:41 |   TASK launch [8/19] w0 running=2 [CP 1 of 2 (Cadable LLC)] Negative News w/o legal ext, page 2
2026-08-18 15:26:48 |   WORKER w0 [Fast] cooldown 2.1s before next task
2026-08-18 15:26:48 |   TASK done  [8/19] w0 wait=93s edgeStart=6s render=1s dur=7s size=151559B running=1
2026-08-18 15:26:50 |   TASK launch [12/19] w0 running=2 [CP 1 of 2 (Cadable LLC)] Name + Address
2026-08-18 15:27:09 |   WORKER w0 [Fast] cooldown 2.6s before next task
2026-08-18 15:27:09 |   TASK done  [12/19] w0 wait=102s edgeStart=19s render=0s dur=19s size=194562B running=1
2026-08-18 15:27:11 |   TASK launch [13/19] w0 running=2 [CP 2 of 2 (First National Bank Cortez)] Name + Address
2026-08-18 15:27:15 |   WORKER w1 [Fast] cooldown 2.5s before next task
2026-08-18 15:27:15 |   TASK retry [11/19] TIMEOUT after 35s (NO FILE EVER) - requeued (attempt 2)
2026-08-18 15:27:17 |   TASK launch [11/19] w1 running=2 [Customer (Currentware INC)] Name + Address
2026-08-18 15:27:22 |   WORKER w1 [Fast] cooldown 2.4s before next task
2026-08-18 15:27:24 |   KILL w1 - terminated 14 orphan msedge.exe process(es)
2026-08-18 15:27:24 |   TASK done  [11/19] w1 wait=129s edgeStart=5s render=2s dur=7s size=206932B running=1
2026-08-18 15:27:24 |   TASK launch [14/19] w1 running=2 [Customer (Currentware INC)] Google name
2026-08-18 15:27:32 |   WORKER w0 [Fast] cooldown 2.9s before next task
2026-08-18 15:27:32 |   TASK done  [13/19] w0 wait=123s edgeStart=20s render=1s dur=21s size=215490B running=1
2026-08-18 15:27:32 |   WORKER w1 [Fast] cooldown 2.8s before next task
2026-08-18 15:27:32 |   TASK done  [14/19] w1 wait=136s edgeStart=8s render=0s dur=8s size=181308B running=0
2026-08-18 15:27:34 |   TASK launch [15/19] w0 running=1 [Customer (Currentware INC)] Address
2026-08-18 15:27:34 |   TASK launch [16/19] w1 running=2 [CP 1 of 2 (Cadable LLC)] Google name
2026-08-18 15:27:52 |   WORKER w1 [Fast] cooldown 3.0s before next task
2026-08-18 15:27:55 |   KILL w1 - terminated 19 orphan msedge.exe process(es)
2026-08-18 15:27:55 |   TASK done  [16/19] w1 wait=146s edgeStart=17s render=4s dur=21s size=169551B running=1
2026-08-18 15:27:55 |   TASK launch [17/19] w1 running=2 [CP 1 of 2 (Cadable LLC)] Address
2026-08-18 15:27:55 |   WORKER w0 [Fast] cooldown 2.4s before next task
2026-08-18 15:27:55 |   TASK done  [15/19] w0 wait=146s edgeStart=21s render=0s dur=21s size=511405B running=1
2026-08-18 15:27:57 |   TASK launch [18/19] w0 running=2 [CP 2 of 2 (First National Bank Cortez)] Google name
2026-08-18 15:28:07 |   WORKER w0 [Fast] cooldown 2.7s before next task
2026-08-18 15:28:07 |   TASK done  [18/19] w0 wait=169s edgeStart=10s render=0s dur=10s size=195492B running=1
2026-08-18 15:28:09 |   TASK launch [19/19] w0 running=2 [CP 2 of 2 (First National Bank Cortez)] Address
2026-08-18 15:28:19 |   WORKER w0 [Fast] cooldown 2.6s before next task
2026-08-18 15:28:22 |   KILL w0 - terminated 22 orphan msedge.exe process(es)
2026-08-18 15:28:22 |   TASK done  [19/19] w0 wait=181s edgeStart=10s render=3s dur=13s size=384838B running=1
2026-08-18 15:28:31 |   WORKER w1 [Fast] cooldown 2.1s before next task
2026-08-18 15:28:31 |   TASK retry [17/19] TIMEOUT after 35s (NO FILE EVER) - requeued (attempt 2)
2026-08-18 15:28:31 |   TASK launch [17/19] w0 running=1 [CP 1 of 2 (Cadable LLC)] Address
2026-08-18 15:28:41 |   WORKER w0 [Fast] cooldown 3.0s before next task
2026-08-18 15:28:41 |   TASK done  [17/19] w0 wait=203s edgeStart=9s render=1s dur=10s size=366369B running=0
2026-08-18 15:28:41 | PASS [main] end, success=19/19 wall=213s
2026-08-18 15:28:41 |   DIST [main] min=5s avg=11s max=21s | <10s=10 <20s=6 <30s=3 <45s=0 <60s=0 60+s=0
2026-08-18 15:28:41 |   UTIL w0 busy=184s/213s = 86%
2026-08-18 15:28:41 |   KILL w0 - terminated 5 orphan msedge.exe process(es)
2026-08-18 15:28:41 |   UTIL w1 busy=188s/213s = 88%
2026-08-18 15:28:42 |   DISP [main] launches=25 inlineRetries=6 dispatcher-sleep=180.3s
2026-08-18 15:28:42 | BATCH ALL end, success=19/19 wall=223s rescue=0 captchaFlagged=0
2026-08-18 15:28:42 | SEARCH PHASE end - elapsed=223s, searches=19
2026-08-18 15:29:07 | BG TUNE: restored 22 background process(es) to Normal
