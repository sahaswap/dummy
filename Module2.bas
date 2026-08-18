2026-08-18 14:38:01 | === Macro started (v3.5) ===
2026-08-18 14:38:01 | RECOVER: stale OSINT state detected on entry - force-clearing (calc/events/alerts/cursor/screen/print/statusbar)
2026-08-18 14:38:02 | STEP 2: before search-mode prompt
2026-08-18 14:38:04 | MODE: search mode set to Fast (USE_HEADLESS=True, jitter=2-3s, protections=False)
2026-08-18 14:38:04 | STEP 2: search-mode prompt returned mode=Fast (USE_HEADLESS=True)
2026-08-18 14:38:04 | STEP 2: before AutoPrewarmDNS
2026-08-18 14:38:04 | STEP 2: after AutoPrewarmDNS
2026-08-18 14:38:04 | STEP 2: before PrewarmEdgeProfiles
2026-08-18 14:38:04 | PREWARM: fired 2 Edge profile warmer(s) - running in background while user is prompted
2026-08-18 14:38:04 | STEP 2: after PrewarmEdgeProfiles
2026-08-18 14:38:04 | STEP 3: before customer negnews InputBox
2026-08-18 14:38:08 | STEP 3: customer negnews returned, len=18
2026-08-18 14:38:08 |   NormalizeSpaces entry, inLen=18
2026-08-18 14:38:08 |   NormalizeSpaces exit, outLen=18 iter=0
2026-08-18 14:38:08 | STEP 3: after NormalizeSpaces, len=18
2026-08-18 14:38:08 | STEP 3: before customer legal-extension InputBox
2026-08-18 14:38:10 |   NormalizeSpaces entry, inLen=14
2026-08-18 14:38:10 |   NormalizeSpaces exit, outLen=14 iter=0
2026-08-18 14:38:10 | STEP 3: before Sigma address MsgBox
2026-08-18 14:38:12 | STEP 3: Sigma TopMostMsgBox returned 6
2026-08-18 14:38:12 | STEP 3: before additional-address InputBox
2026-08-18 14:38:30 | STEP 3: additional-address returned
2026-08-18 14:38:30 |   NormalizeSpaces entry, inLen=51
2026-08-18 14:38:30 |   NormalizeSpaces exit, outLen=51 iter=0
2026-08-18 14:38:30 | STEP 3: customer prompts complete
2026-08-18 14:38:33 |   NormalizeSpaces entry, inLen=24
2026-08-18 14:38:33 |   NormalizeSpaces exit, outLen=24 iter=0
2026-08-18 14:38:36 |   NormalizeSpaces entry, inLen=0
2026-08-18 14:38:36 |   NormalizeSpaces exit, outLen=0 iter=0
2026-08-18 14:38:37 |   NormalizeSpaces entry, inLen=19
2026-08-18 14:38:37 |   NormalizeSpaces exit, outLen=19 iter=0
2026-08-18 14:38:40 |   NormalizeSpaces entry, inLen=9
2026-08-18 14:38:40 |   NormalizeSpaces exit, outLen=9 iter=0
2026-08-18 14:38:41 |   NormalizeSpaces entry, inLen=25
2026-08-18 14:38:41 |   NormalizeSpaces exit, outLen=25 iter=0
2026-08-18 14:38:43 |   NormalizeSpaces entry, inLen=21
2026-08-18 14:38:43 |   NormalizeSpaces exit, outLen=21 iter=0
2026-08-18 14:38:43 | STEP 3: before Starting Searches MsgBox
2026-08-18 14:38:44 | STEP 3: Starting Searches MsgBox dismissed -> proceeding to searches
2026-08-18 14:38:47 | BG TUNE: lowered 22 background process(es) to BelowNormal
2026-08-18 14:38:47 | SEARCH PHASE start - total entities to process: 4
2026-08-18 14:38:51 | GAP-FILL: analyst chose FULL re-run of all 27 task(s)
2026-08-18 14:38:51 | SORT: reordered 27 tasks (heavy first) - heaviest weight=10, lightest=1
2026-08-18 14:38:51 | SEARCH PHASE - queued 27 tasks across all entities (heaviest first)
2026-08-18 14:38:51 | BATCH ALL start, n=27, maxParallel=2
2026-08-18 14:38:51 | PREWARM: drain done in 0.62s (allDone=True)
2026-08-18 14:38:51 | COOKIE WARM: fired 2 HEADLESS (hidden) warm-up(s) to pick up Google session cookies
2026-08-18 14:39:00 | COOKIE WARM: done after 8s (headless)
2026-08-18 14:39:00 | BATCH ALL: cookie warm-up RAN before main pass
2026-08-18 14:39:00 | PASS [main] start, n=27, maxRetries=7
2026-08-18 14:39:00 |   TASK launch [1/27] w0 running=1 [Customer (Ritzy Charters LLC)] Negative News page 1
2026-08-18 14:39:00 |   TASK launch [2/27] w1 running=2 [Customer (Ritzy Charters LLC)] Negative News page 2
2026-08-18 14:39:05 |   WORKER w0 [Fast] cooldown 2.5s before next task
2026-08-18 14:39:05 |   TASK retry [1/27] CAPTCHA-size 52279B - requeued (attempt 2/8), backing off 0.0s before relaunch
2026-08-18 14:39:06 |   WORKER w1 [Fast] cooldown 2.3s before next task
2026-08-18 14:39:06 |   TASK done  [2/27] w1 wait=0s edgeStart=6s render=0s dur=6s size=150828B running=0
2026-08-18 14:39:07 |   TASK launch [1/27] w0 running=1 [Customer (Ritzy Charters LLC)] Negative News page 1
2026-08-18 14:39:08 |   TASK launch [3/27] w1 running=2 [Customer (Ritzy Charters LLC)] Negative News w/o legal ext, page 1
2026-08-18 14:39:12 |   WORKER w0 [Fast] cooldown 2.9s before next task
2026-08-18 14:39:12 |   TASK retry [1/27] CAPTCHA-size 52279B - requeued (attempt 3/8), backing off 0.0s before relaunch
2026-08-18 14:39:13 |   WORKER w1 [Fast] cooldown 2.8s before next task
2026-08-18 14:39:13 |   TASK done  [3/27] w1 wait=8s edgeStart=5s render=0s dur=5s size=176805B running=0
2026-08-18 14:39:14 |   TASK launch [1/27] w0 running=1 [Customer (Ritzy Charters LLC)] Negative News page 1
2026-08-18 14:39:15 |   TASK launch [4/27] w1 running=2 [Customer (Ritzy Charters LLC)] Negative News w/o legal ext, page 2
2026-08-18 14:39:21 |   WORKER w0 [Fast] cooldown 2.8s before next task
2026-08-18 14:39:21 |   TASK done  [1/27] w0 wait=14s edgeStart=6s render=1s dur=7s size=267589B running=1
2026-08-18 14:39:23 |   TASK launch [5/27] w0 running=2 [CP 1 of 3 (A Chernacov And P Tekiel)] Negative News page 1
2026-08-18 14:39:23 |   WORKER w1 [Fast] cooldown 2.7s before next task
2026-08-18 14:39:23 |   TASK done  [4/27] w1 wait=15s edgeStart=8s render=0s dur=8s size=151309B running=1
2026-08-18 14:39:25 |   TASK launch [6/27] w1 running=2 [CP 1 of 3 (A Chernacov And P Tekiel)] Negative News page 2
2026-08-18 14:39:33 |   WORKER w0 [Fast] cooldown 2.4s before next task
2026-08-18 14:39:33 |   TASK done  [5/27] w0 wait=23s edgeStart=9s render=1s dur=10s size=240452B running=1
2026-08-18 14:39:35 |   TASK launch [7/27] w0 running=2 [CP 2 of 3 (Eric G Stucchi Arce)] Negative News page 1
2026-08-18 14:39:42 |   WORKER w1 [Fast] cooldown 3.0s before next task
2026-08-18 14:39:42 |   TASK done  [6/27] w1 wait=25s edgeStart=17s render=0s dur=17s size=165862B running=1
2026-08-18 14:39:44 |   TASK launch [8/27] w1 running=2 [CP 2 of 3 (Eric G Stucchi Arce)] Negative News page 2
2026-08-18 14:39:44 |   WORKER w0 [Fast] cooldown 2.5s before next task
2026-08-18 14:39:44 |   TASK done  [7/27] w0 wait=35s edgeStart=9s render=0s dur=9s size=221054B running=1
2026-08-18 14:39:46 |   TASK launch [9/27] w0 running=2 [CP 2 of 3 (Eric G Stucchi Arce)] Negative News w/o middle, page 1
2026-08-18 14:40:02 |   WORKER w0 [Fast] cooldown 2.2s before next task
2026-08-18 14:40:08 |   KILL w0 - terminated 28 orphan msedge.exe process(es)
2026-08-18 14:40:08 |   TASK done  [9/27] w0 wait=46s edgeStart=16s render=6s dur=22s size=210974B running=1
2026-08-18 14:40:08 |   TASK launch [10/27] w0 running=2 [CP 2 of 3 (Eric G Stucchi Arce)] Negative News w/o middle, page 2
2026-08-18 14:40:08 |   WORKER w1 [Fast] cooldown 2.8s before next task
2026-08-18 14:40:08 |   TASK done  [8/27] w1 wait=44s edgeStart=24s render=0s dur=24s size=151703B running=1
2026-08-18 14:40:10 |   TASK launch [11/27] w1 running=2 [CP 3 of 3 (Speak Breath Retreats LLC)] Negative News page 1
2026-08-18 14:40:14 |   WORKER w0 [Fast] cooldown 2.8s before next task
2026-08-18 14:40:14 |   TASK done  [10/27] w0 wait=68s edgeStart=6s render=0s dur=6s size=219376B running=1
2026-08-18 14:40:16 |   TASK launch [12/27] w0 running=2 [CP 3 of 3 (Speak Breath Retreats LLC)] Negative News page 2
2026-08-18 14:40:17 |   WORKER w1 [Fast] cooldown 2.8s before next task
2026-08-18 14:40:17 |   TASK done  [11/27] w1 wait=70s edgeStart=6s render=1s dur=7s size=145303B running=1
2026-08-18 14:40:19 |   TASK launch [13/27] w1 running=2 [CP 3 of 3 (Speak Breath Retreats LLC)] Negative News w/o legal ext, page 1
2026-08-18 14:40:30 |   WORKER w1 [Fast] cooldown 2.8s before next task
2026-08-18 14:40:30 |   TASK done  [13/27] w1 wait=79s edgeStart=11s render=0s dur=11s size=165566B running=1
2026-08-18 14:40:31 |   WORKER w0 [Fast] cooldown 2.2s before next task
2026-08-18 14:40:31 |   TASK done  [12/27] w0 wait=76s edgeStart=14s render=1s dur=15s size=153175B running=0
2026-08-18 14:40:32 |   TASK launch [14/27] w1 running=1 [CP 3 of 3 (Speak Breath Retreats LLC)] Negative News w/o legal ext, page 2
2026-08-18 14:40:33 |   TASK launch [15/27] w0 running=2 [Customer (Ritzy Charters LLC)] Name + Address
2026-08-18 14:40:49 |   WORKER w1 [Fast] cooldown 2.9s before next task
2026-08-18 14:40:49 |   TASK done  [14/27] w1 wait=92s edgeStart=17s render=0s dur=17s size=151473B running=1
2026-08-18 14:40:51 |   TASK launch [16/27] w1 running=2 [CP 1 of 3 (A Chernacov And P Tekiel)] Name + Address
2026-08-18 14:40:54 |   WORKER w0 [Fast] cooldown 2.5s before next task
2026-08-18 14:40:54 |   TASK done  [15/27] w0 wait=93s edgeStart=21s render=0s dur=21s size=281245B running=1
2026-08-18 14:40:56 |   TASK launch [17/27] w0 running=2 [CP 2 of 3 (Eric G Stucchi Arce)] Name + Address
2026-08-18 14:41:09 |   WORKER w0 [Fast] cooldown 2.7s before next task
2026-08-18 14:41:12 |   KILL w0 - terminated 19 orphan msedge.exe process(es)
2026-08-18 14:41:12 |   TASK done  [17/27] w0 wait=116s edgeStart=13s render=3s dur=16s size=107482B running=1
2026-08-18 14:41:12 |   TASK launch [18/27] w0 running=2 [CP 3 of 3 (Speak Breath Retreats LLC)] Name + Address
2026-08-18 14:41:13 |   WORKER w1 [Fast] cooldown 2.3s before next task
2026-08-18 14:41:13 |   TASK done  [16/27] w1 wait=111s edgeStart=21s render=1s dur=22s size=155004B running=1
2026-08-18 14:41:15 |   TASK launch [19/27] w1 running=2 [Customer (Ritzy Charters LLC)] Google name
2026-08-18 14:41:18 |   WORKER w0 [Fast] cooldown 2.2s before next task
2026-08-18 14:41:18 |   TASK done  [18/27] w0 wait=132s edgeStart=6s render=0s dur=6s size=147333B running=1
2026-08-18 14:41:20 |   TASK launch [20/27] w0 running=2 [Customer (Ritzy Charters LLC)] Address
2026-08-18 14:41:26 |   WORKER w1 [Fast] cooldown 2.7s before next task
2026-08-18 14:41:26 |   TASK done  [19/27] w1 wait=135s edgeStart=11s render=0s dur=11s size=192491B running=1
2026-08-18 14:41:28 |   TASK launch [21/27] w1 running=2 [Customer (Ritzy Charters LLC)] Name + Additional Address
2026-08-18 14:41:30 |   WORKER w0 [Fast] cooldown 2.1s before next task
2026-08-18 14:41:30 |   TASK done  [20/27] w0 wait=140s edgeStart=10s render=0s dur=10s size=377191B running=1
2026-08-18 14:41:32 |   TASK launch [22/27] w0 running=2 [CP 1 of 3 (A Chernacov And P Tekiel)] Google name
2026-08-18 14:41:42 |   WORKER w0 [Fast] cooldown 2.2s before next task
2026-08-18 14:41:42 |   TASK done  [22/27] w0 wait=152s edgeStart=10s render=0s dur=10s size=181455B running=1
2026-08-18 14:41:44 |   TASK launch [23/27] w0 running=2 [CP 1 of 3 (A Chernacov And P Tekiel)] Address
2026-08-18 14:41:45 |   WORKER w1 [Fast] cooldown 2.4s before next task
2026-08-18 14:41:45 |   TASK done  [21/27] w1 wait=148s edgeStart=17s render=0s dur=17s size=227085B running=1
2026-08-18 14:41:47 |   TASK launch [24/27] w1 running=2 [CP 2 of 3 (Eric G Stucchi Arce)] Google name
2026-08-18 14:42:15 |   WORKER w1 [Fast] cooldown 2.7s before next task
2026-08-18 14:42:15 |   TASK retry [24/27] CAPTCHA-size 71520B - requeued (attempt 2/8), backing off 0.0s before relaunch
2026-08-18 14:42:17 |   TASK launch [24/27] w1 running=2 [CP 2 of 3 (Eric G Stucchi Arce)] Google name
2026-08-18 14:42:21 |   KILL w0 - terminated 4 orphan msedge.exe process(es)
2026-08-18 14:42:21 |   WORKER w0 [Fast] cooldown 2.5s before next task
2026-08-18 14:42:21 |   TASK retry [23/27] TIMEOUT after 35s (NO FILE EVER) - requeued (attempt 2)
2026-08-18 14:42:23 |   TASK launch [23/27] w0 running=2 [CP 1 of 3 (A Chernacov And P Tekiel)] Address
2026-08-18 14:42:28 |   WORKER w1 [Fast] cooldown 2.4s before next task
2026-08-18 14:42:28 |   TASK retry [24/27] CAPTCHA-size 71520B - requeued (attempt 3/8), backing off 0.0s before relaunch
2026-08-18 14:42:30 |   TASK launch [24/27] w1 running=2 [CP 2 of 3 (Eric G Stucchi Arce)] Google name
2026-08-18 14:42:30 |   WORKER w0 [Fast] cooldown 2.4s before next task
2026-08-18 14:42:30 |   TASK done  [23/27] w0 wait=203s edgeStart=7s render=0s dur=7s size=558359B running=1
2026-08-18 14:42:32 |   TASK launch [25/27] w0 running=2 [CP 2 of 3 (Eric G Stucchi Arce)] Address
2026-08-18 14:42:51 |   WORKER w0 [Fast] cooldown 2.3s before next task
2026-08-18 14:42:54 |   KILL w0 - terminated 14 orphan msedge.exe process(es)
2026-08-18 14:42:54 |   TASK done  [25/27] w0 wait=212s edgeStart=19s render=3s dur=22s size=374394B running=1
2026-08-18 14:42:54 |   TASK launch [26/27] w0 running=2 [CP 3 of 3 (Speak Breath Retreats LLC)] Google name
2026-08-18 14:42:54 |   HUMAN-SOLVE armed after 2 retr(ies) (size-based CAPTCHA, PDF 71520B)
2026-08-18 14:42:54 |   WORKER w1 [Fast] cooldown 2.4s before next task
2026-08-18 14:42:54 |   TASK retry [24/27] CAPTCHA-size 71520B - requeued (attempt 4/8), backing off 0.0s before relaunch
2026-08-18 14:42:54 | HUMAN-SOLVE: CAPTCHA hit - pausing for analyst to solve
2026-08-18 14:42:56 |   KILL w0 - terminated 7 orphan msedge.exe process(es)
2026-08-18 14:42:56 |   KILL w1 - terminated 2 orphan msedge.exe process(es)
2026-08-18 14:43:34 |   KILL w0 - terminated 13 orphan msedge.exe process(es)
2026-08-18 14:43:49 |   KILL w1 - terminated 11 orphan msedge.exe process(es)
2026-08-18 14:43:50 | HUMAN-SOLVE: analyst confirmed - resuming (in-flight tasks requeued)
2026-08-18 14:43:50 |   TASK launch [24/27] w0 running=1 [CP 2 of 3 (Eric G Stucchi Arce)] Google name
2026-08-18 14:43:50 |   TASK launch [26/27] w1 running=2 [CP 3 of 3 (Speak Breath Retreats LLC)] Google name
2026-08-18 14:43:55 |   WORKER w1 [Fast] cooldown 2.9s before next task
2026-08-18 14:43:55 |   TASK retry [26/27] CAPTCHA-size 70996B - requeued (attempt 2/8), backing off 0.0s before relaunch
2026-08-18 14:43:56 |   WORKER w0 [Fast] cooldown 2.8s before next task
2026-08-18 14:43:56 |   TASK done  [24/27] w0 wait=290s edgeStart=6s render=0s dur=6s size=206910B running=0
2026-08-18 14:43:57 |   TASK launch [26/27] w1 running=1 [CP 3 of 3 (Speak Breath Retreats LLC)] Google name
2026-08-18 14:43:58 |   TASK launch [27/27] w0 running=2 [CP 3 of 3 (Speak Breath Retreats LLC)] Address
2026-08-18 14:44:04 |   WORKER w1 [Fast] cooldown 2.8s before next task
2026-08-18 14:44:04 |   TASK retry [26/27] CAPTCHA-size 70991B - requeued (attempt 3/8), backing off 0.0s before relaunch
2026-08-18 14:44:06 |   TASK launch [26/27] w1 running=2 [CP 3 of 3 (Speak Breath Retreats LLC)] Google name
2026-08-18 14:44:08 |   WORKER w0 [Fast] cooldown 2.9s before next task
2026-08-18 14:44:08 |   TASK done  [27/27] w0 wait=298s edgeStart=10s render=0s dur=10s size=260461B running=1
2026-08-18 14:44:14 |   WORKER w1 [Fast] cooldown 2.0s before next task
2026-08-18 14:44:14 |   TASK retry [26/27] CAPTCHA-size 70990B - requeued (attempt 4/8), backing off 0.0s before relaunch
2026-08-18 14:44:14 |   TASK launch [26/27] w0 running=1 [CP 3 of 3 (Speak Breath Retreats LLC)] Google name
2026-08-18 14:44:20 |   WORKER w0 [Fast] cooldown 2.0s before next task
2026-08-18 14:44:20 |   TASK done  [26/27] w0 wait=314s edgeStart=6s render=0s dur=6s size=251106B running=0
2026-08-18 14:44:20 | PASS [main] end, success=27/27 wall=320s
2026-08-18 14:44:20 |   DIST [main] min=5s avg=12s max=24s | <10s=11 <20s=11 <30s=5 <45s=0 <60s=0 60+s=0
2026-08-18 14:44:20 |   UTIL w0 busy=230s/320s = 71%
2026-08-18 14:44:24 |   KILL w0 - terminated 15 orphan msedge.exe process(es)
2026-08-18 14:44:24 |   UTIL w1 busy=228s/320s = 71%
2026-08-18 14:44:26 |   KILL w1 - terminated 1 orphan msedge.exe process(es)
2026-08-18 14:44:26 |   DISP [main] launches=37 inlineRetries=9 dispatcher-sleep=240.2s
2026-08-18 14:44:26 | BATCH ALL end, success=27/27 wall=336s rescue=0 captchaFlagged=0
2026-08-18 14:44:26 | SEARCH PHASE end - elapsed=339s, searches=27
2026-08-18 14:44:56 | BG TUNE: restored 22 background process(es) to Normal
