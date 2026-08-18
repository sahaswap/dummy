2026-08-18 15:50:36 | === Macro started (v3.5) ===
2026-08-18 15:50:36 | RECOVER: stale OSINT state detected on entry - force-clearing (calc/events/alerts/cursor/screen/print/statusbar)
2026-08-18 15:50:36 | STEP 2: before search-mode prompt
2026-08-18 15:50:41 | MODE: search mode set to Fast (USE_HEADLESS=True, jitter=2-3s, protections=False)
2026-08-18 15:50:41 | STEP 2: search-mode prompt returned mode=Fast (USE_HEADLESS=True)
2026-08-18 15:50:41 | STEP 2: before AutoPrewarmDNS
2026-08-18 15:50:41 | STEP 2: after AutoPrewarmDNS
2026-08-18 15:50:41 | STEP 2: before PrewarmEdgeProfiles
2026-08-18 15:50:41 | PREWARM: fired 2 Edge profile warmer(s) - running in background while user is prompted
2026-08-18 15:50:41 | STEP 2: after PrewarmEdgeProfiles
2026-08-18 15:50:41 | STEP 3: before customer negnews InputBox
2026-08-18 15:50:50 | STEP 3: customer negnews returned, len=15
2026-08-18 15:50:50 |   NormalizeSpaces entry, inLen=15
2026-08-18 15:50:50 |   NormalizeSpaces exit, outLen=15 iter=0
2026-08-18 15:50:50 | STEP 3: after NormalizeSpaces, len=15
2026-08-18 15:50:50 | STEP 3: before customer legal-extension InputBox
2026-08-18 15:50:53 |   NormalizeSpaces entry, inLen=11
2026-08-18 15:50:53 |   NormalizeSpaces exit, outLen=11 iter=0
2026-08-18 15:50:53 | STEP 3: before Sigma address MsgBox
2026-08-18 15:50:56 | STEP 3: Sigma TopMostMsgBox returned 7
2026-08-18 15:50:56 | STEP 3: customer prompts complete
2026-08-18 15:50:58 |   NormalizeSpaces entry, inLen=11
2026-08-18 15:50:58 |   NormalizeSpaces exit, outLen=11 iter=0
2026-08-18 15:51:03 |   NormalizeSpaces entry, inLen=7
2026-08-18 15:51:03 |   NormalizeSpaces exit, outLen=7 iter=0
2026-08-18 15:51:06 |   NormalizeSpaces entry, inLen=26
2026-08-18 15:51:06 |   NormalizeSpaces exit, outLen=26 iter=0
2026-08-18 15:51:08 |   NormalizeSpaces entry, inLen=0
2026-08-18 15:51:08 |   NormalizeSpaces exit, outLen=0 iter=0
2026-08-18 15:51:08 | STEP 3: before Starting Searches MsgBox
2026-08-18 15:51:09 | STEP 3: Starting Searches MsgBox dismissed -> proceeding to searches
2026-08-18 15:51:12 | BG TUNE: lowered 22 background process(es) to BelowNormal
2026-08-18 15:51:12 | SEARCH PHASE start - total entities to process: 3
2026-08-18 15:51:21 | GAP-FILL: analyst chose FULL re-run of all 19 task(s)
2026-08-18 15:51:21 | SORT: reordered 19 tasks (heavy first) - heaviest weight=10, lightest=1
2026-08-18 15:51:21 | SEARCH PHASE - queued 19 tasks across all entities (heaviest first)
2026-08-18 15:51:21 | BATCH ALL start, n=19, maxParallel=2
2026-08-18 15:51:22 | PREWARM: drain done in 1.08s (allDone=True)
2026-08-18 15:51:22 | COOKIE WARM: fired 2 HEADLESS (hidden) warm-up(s) to pick up Google session cookies
2026-08-18 15:51:31 | COOKIE WARM: done after 8s (headless)
2026-08-18 15:51:31 | BATCH ALL: cookie warm-up RAN before main pass
2026-08-18 15:51:31 | PASS [main] start, n=19, maxRetries=7
2026-08-18 15:51:31 |   TASK launch [1/19] w0 running=1 [Customer (Currentware INC)] Negative News page 1
2026-08-18 15:51:31 |   TASK launch [2/19] w1 running=2 [Customer (Currentware INC)] Negative News page 2
2026-08-18 15:51:37 |   WORKER w1 [Fast] cooldown 2.0s before next task
2026-08-18 15:51:37 |   TASK retry [2/19] CAPTCHA-size 69124B - requeued (attempt 2/8), backing off 0.0s before relaunch
2026-08-18 15:51:37 |   WORKER w0 [Fast] cooldown 2.3s before next task
2026-08-18 15:51:37 |   TASK done  [1/19] w0 wait=0s edgeStart=6s render=0s dur=6s size=198832B running=0
2026-08-18 15:51:39 |   TASK launch [2/19] w0 running=1 [Customer (Currentware INC)] Negative News page 2
2026-08-18 15:51:39 |   TASK launch [3/19] w1 running=2 [Customer (Currentware INC)] Negative News w/o legal ext, page 1
2026-08-18 15:51:44 |   WORKER w0 [Fast] cooldown 2.4s before next task
2026-08-18 15:51:44 |   TASK retry [2/19] CAPTCHA-size 62130B - requeued (attempt 3/8), backing off 0.0s before relaunch
2026-08-18 15:51:46 |   TASK launch [2/19] w0 running=2 [Customer (Currentware INC)] Negative News page 2
2026-08-18 15:51:48 |   WORKER w1 [Fast] cooldown 3.0s before next task
2026-08-18 15:51:48 |   TASK done  [3/19] w1 wait=8s edgeStart=9s render=0s dur=9s size=287985B running=1
2026-08-18 15:51:50 |   TASK launch [4/19] w1 running=2 [Customer (Currentware INC)] Negative News w/o legal ext, page 2
2026-08-18 15:51:52 |   WORKER w0 [Fast] cooldown 2.9s before next task
2026-08-18 15:51:54 |   KILL w0 - terminated 10 orphan msedge.exe process(es)
2026-08-18 15:51:54 |   TASK done  [2/19] w0 wait=15s edgeStart=6s render=2s dur=8s size=203980B running=1
2026-08-18 15:51:54 |   TASK launch [5/19] w0 running=2 [CP 1 of 2 (Cadable LLC)] Negative News page 1
2026-08-18 15:51:59 |   WORKER w1 [Fast] cooldown 2.8s before next task
2026-08-18 15:51:59 |   TASK done  [4/19] w1 wait=19s edgeStart=9s render=0s dur=9s size=223816B running=1
2026-08-18 15:52:01 |   TASK launch [6/19] w1 running=2 [CP 1 of 2 (Cadable LLC)] Negative News page 2
2026-08-18 15:52:03 |   WORKER w0 [Fast] cooldown 2.8s before next task
2026-08-18 15:52:03 |   TASK done  [5/19] w0 wait=23s edgeStart=8s render=1s dur=9s size=229325B running=1
2026-08-18 15:52:06 |   TASK launch [7/19] w0 running=2 [CP 1 of 2 (Cadable LLC)] Negative News w/o legal ext, page 1
2026-08-18 15:52:11 |   WORKER w0 [Fast] cooldown 2.6s before next task
2026-08-18 15:52:13 |   KILL w0 - terminated 15 orphan msedge.exe process(es)
2026-08-18 15:52:13 |   TASK done  [7/19] w0 wait=35s edgeStart=5s render=2s dur=7s size=317144B running=1
2026-08-18 15:52:14 |   TASK launch [8/19] w0 running=2 [CP 1 of 2 (Cadable LLC)] Negative News w/o legal ext, page 2
2026-08-18 15:52:20 |   WORKER w0 [Fast] cooldown 2.2s before next task
2026-08-18 15:52:20 |   TASK done  [8/19] w0 wait=43s edgeStart=5s render=1s dur=6s size=307975B running=1
2026-08-18 15:52:22 |   TASK launch [9/19] w0 running=2 [CP 2 of 2 (First National Bank Cortez)] Negative News page 1
2026-08-18 15:52:28 |   WORKER w0 [Fast] cooldown 2.1s before next task
2026-08-18 15:52:28 |   TASK done  [9/19] w0 wait=51s edgeStart=6s render=0s dur=6s size=202599B running=1
2026-08-18 15:52:30 |   TASK launch [10/19] w0 running=2 [CP 2 of 2 (First National Bank Cortez)] Negative News page 2
2026-08-18 15:52:37 |   WORKER w1 [Fast] cooldown 2.6s before next task
2026-08-18 15:52:37 |   TASK retry [6/19] TIMEOUT after 35s (NO FILE EVER) - requeued (attempt 2)
2026-08-18 15:52:39 |   TASK launch [6/19] w1 running=2 [CP 1 of 2 (Cadable LLC)] Negative News page 2
2026-08-18 15:52:39 |   WORKER w0 [Fast] cooldown 2.2s before next task
2026-08-18 15:52:41 |   KILL w0 - terminated 10 orphan msedge.exe process(es)
2026-08-18 15:52:41 |   TASK done  [10/19] w0 wait=59s edgeStart=9s render=2s dur=11s size=214079B running=1
2026-08-18 15:52:41 |   TASK launch [11/19] w0 running=2 [Customer (Currentware INC)] Name + Address
2026-08-18 15:52:45 |   WORKER w1 [Fast] cooldown 2.2s before next task
2026-08-18 15:52:45 |   TASK done  [6/19] w1 wait=68s edgeStart=5s render=1s dur=6s size=150778B running=1
2026-08-18 15:52:47 |   TASK launch [12/19] w1 running=2 [CP 1 of 2 (Cadable LLC)] Name + Address
2026-08-18 15:53:01 |   WORKER w1 [Fast] cooldown 2.1s before next task
2026-08-18 15:53:01 |   TASK done  [12/19] w1 wait=76s edgeStart=14s render=0s dur=14s size=244494B running=1
2026-08-18 15:53:03 |   TASK launch [13/19] w1 running=2 [CP 2 of 2 (First National Bank Cortez)] Name + Address
2026-08-18 15:53:16 |   WORKER w1 [Fast] cooldown 2.9s before next task
2026-08-18 15:53:17 |   KILL w1 - terminated 3 orphan msedge.exe process(es)
2026-08-18 15:53:17 |   TASK done  [13/19] w1 wait=92s edgeStart=13s render=1s dur=14s size=294460B running=1
2026-08-18 15:53:18 |   WORKER w0 [Fast] cooldown 2.1s before next task
2026-08-18 15:53:18 |   TASK retry [11/19] TIMEOUT after 35s (NO FILE EVER) - requeued (attempt 2)
2026-08-18 15:53:18 |   TASK launch [11/19] w1 running=1 [Customer (Currentware INC)] Name + Address
2026-08-18 15:53:20 |   TASK launch [14/19] w0 running=2 [Customer (Currentware INC)] Google name
2026-08-18 15:53:24 |   WORKER w1 [Fast] cooldown 2.9s before next task
2026-08-18 15:53:24 |   TASK done  [11/19] w1 wait=107s edgeStart=6s render=0s dur=6s size=251549B running=1
2026-08-18 15:53:26 |   TASK launch [15/19] w1 running=2 [Customer (Currentware INC)] Address
2026-08-18 15:53:27 |   WORKER w0 [Fast] cooldown 2.6s before next task
2026-08-18 15:53:27 |   TASK done  [14/19] w0 wait=109s edgeStart=6s render=1s dur=7s size=255016B running=1
2026-08-18 15:53:29 |   TASK launch [16/19] w0 running=2 [CP 1 of 2 (Cadable LLC)] Google name
2026-08-18 15:53:39 |   WORKER w0 [Fast] cooldown 2.5s before next task
2026-08-18 15:53:42 |   KILL w0 - terminated 12 orphan msedge.exe process(es)
2026-08-18 15:53:42 |   TASK done  [16/19] w0 wait=118s edgeStart=9s render=4s dur=13s size=244862B running=1
2026-08-18 15:53:42 |   TASK launch [17/19] w0 running=2 [CP 1 of 2 (Cadable LLC)] Address
2026-08-18 15:53:42 |   WORKER w1 [Fast] cooldown 2.7s before next task
2026-08-18 15:53:42 |   TASK done  [15/19] w1 wait=115s edgeStart=16s render=0s dur=16s size=500997B running=1
2026-08-18 15:53:44 |   TASK launch [18/19] w1 running=2 [CP 2 of 2 (First National Bank Cortez)] Google name
2026-08-18 15:53:51 |   WORKER w1 [Fast] cooldown 2.3s before next task
2026-08-18 15:53:51 |   TASK done  [18/19] w1 wait=133s edgeStart=7s render=0s dur=7s size=274275B running=1
2026-08-18 15:53:53 |   TASK launch [19/19] w1 running=2 [CP 2 of 2 (First National Bank Cortez)] Address
2026-08-18 15:54:03 |   WORKER w1 [Fast] cooldown 2.7s before next task
2026-08-18 15:54:04 |   KILL w1 - terminated 10 orphan msedge.exe process(es)
2026-08-18 15:54:04 |   TASK done  [19/19] w1 wait=142s edgeStart=10s render=1s dur=11s size=366697B running=1
2026-08-18 15:54:18 |   WORKER w0 [Fast] cooldown 2.2s before next task
2026-08-18 15:54:18 |   TASK retry [17/19] TIMEOUT after 35s (NO FILE EVER) - requeued (attempt 2)
2026-08-18 15:54:18 |   TASK launch [17/19] w1 running=1 [CP 1 of 2 (Cadable LLC)] Address
2026-08-18 15:54:27 |   WORKER w1 [Fast] cooldown 2.8s before next task
2026-08-18 15:54:27 |   TASK done  [17/19] w1 wait=167s edgeStart=9s render=0s dur=9s size=442939B running=0
2026-08-18 15:54:27 | PASS [main] end, success=19/19 wall=177s
2026-08-18 15:54:27 |   DIST [main] min=6s avg=9s max=16s | <10s=13 <20s=6 <30s=0 <45s=0 <60s=0 60+s=0
2026-08-18 15:54:27 |   UTIL w0 busy=151s/177s = 85%
2026-08-18 15:54:28 |   UTIL w1 busy=143s/177s = 80%
2026-08-18 15:54:28 |   DISP [main] launches=24 inlineRetries=5 dispatcher-sleep=150.9s
2026-08-18 15:54:28 | BATCH ALL end, success=19/19 wall=187s rescue=0 captchaFlagged=0
2026-08-18 15:54:28 | SEARCH PHASE end - elapsed=196s, searches=19
2026-08-18 15:54:54 | BG TUNE: restored 22 background process(es) to Normal
