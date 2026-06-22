# Silent Report Parallelism Implementation Plan

1. Add failing architecture tests for log-only timings and research/analysis overlap, then implement pipeline orchestration.
2. Add failing analytics tests for single preparation/reuse, then implement the request-local prepared-analysis packet and bounded helper memoization.
3. Add failing writing-stage tests for global concurrency bound, output order, and fallback preservation, then add the shared two-worker executor.
4. Add failing chart tests for cache hits, independent streams, and locked rendering, then add the bounded PNG-byte LRU.
5. Run targeted tests, the relevant report suite, compilation, deterministic document equivalence, and one/two-job timing checks.
6. Review the complete diff, deploy changed runtime files to the VPS with neutral timestamped backups, restart the service, verify loopback/public health and readiness plus preservation counts, and remove temporary artifacts.
