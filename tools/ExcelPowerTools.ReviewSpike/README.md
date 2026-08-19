# Review.xlsx retention and scale spike

This .NET 8 headless harness exercises the production `ReviewWorkbookFile` boundary. It creates a deterministic
visible A:K change projection plus deterministic synthetic mini-grid snapshot payloads in embedded state, validates
the closed package, and emits one compact JSON metrics object. These payloads are not drawings or proof of future
image-snapshot retention; that remains a separate rendering spike. This harness never starts or automates Excel.

## Generate and verify one artifact

```powershell
dotnet run --project tools\ExcelPowerTools.ReviewSpike -- generate `
  --output C:\Temp\review-10000.xlsx --rows 10000 `
  --state-bytes 2097152 `
  --baseline-output C:\Temp\review-10000.before.json

# Run this after a deliberate Excel open/save/close cycle.
dotnet run --project tools\ExcelPowerTools.ReviewSpike -- verify `
  --input C:\Temp\review-10000.xlsx `
  --baseline C:\Temp\review-10000.before.json
```

`--state-bytes` is a bounded deterministic minimum/target from 0 through 32 MiB. It uses deterministic,
non-highly-compressible padding and reports both requested and actual UTF-8 bytes. `--chunk-size` ranges from 1 byte
through 4 MiB. `verify --baseline-output <new.json>` can capture a baseline independently of generation.

Verification holds one deny-write/delete file handle while it runs the strict production reader, validates the exact
spike schema/generator/version/project ID and deterministic snapshots, and independently streams every Changes row,
A:K reference, and expected value into `visibleProjectionSemanticSha256`. A baseline comparison requires the logical
project ID, rows, embedded-state and projection hashes, independent visible checksum, chunk count/item IDs, and
decision-sidecar identity to match. The raw ZIP/package hash and byte size are reported but are explicitly
nondeterministic and may change after a harmless Excel save.

## Canonical scale matrix

```powershell
dotnet run --project tools\ExcelPowerTools.ReviewSpike -- matrix `
  --output-directory C:\Temp\review-matrix
```

The existing output directory must be empty for the planned names. The fixed matrix is:

| Visible rows | Requested embedded state |
| ---: | ---: |
| 10,000 | 2 MiB |
| 100,000 | 8 MiB |
| 250,000 | 24 MiB |

Each tier gets a create-new `.xlsx` and baseline. Tiers are independent: a failed tier is recorded as structured JSON,
its own partial output is removed, and later tiers still run. Both the production reader and the independent Changes
inspector retain the production 512 MiB streaming worksheet-part bound; the harness does not weaken that guard.

Measured headless checkpoint on 2026-07-17:

| Visible rows | Package bytes | `Changes` XML bytes | Generate + strict verify |
| ---: | ---: | ---: | ---: |
| 10,000 | 1,863,775 | 8,450,734 | 6.25 s |
| 100,000 | 9,928,110 | 85,797,300 | 30.14 s |
| 250,000 | 27,679,172 | 216,658,200 | 82.98 s |

The full one-process matrix took 119.50 seconds and reported a 1,023,442,944-byte peak (about 976 MiB) after
generating all three tiers. A fresh process verifying only the 250k artifact took 46.11 seconds—25.85 seconds in
the production validator and 20.07 seconds in the independent inspector—and peaked at 232,353,792 bytes. Every semantic baseline
matched. These numbers prove the headless package path, not x86 Excel behavior. The real-Excel runner still needs a
deliberately isolated run with every pre-existing Excel process closed.

Outputs are create-new publications and are never overwritten. The harness accepts only an existing directory on a
local fixed-drive path whose current ancestors are not reparse points. That directory is a trusted publication
boundary and must not be replaced while a command is running. Success exits 0; command or strict validation failure
exits 1; a baseline semantic mismatch exits 3; an incomplete matrix emits all tier results and exits 4. Metrics include
phase timings, compressed/uncompressed sizes, hashes and retention booleans, allocation and process working-set
observations, runtime/architecture, and disk capacity metadata.

For the deliberately isolated real-Excel open/save/reopen workflow, see
[RETENTION-RUNNER.md](./RETENTION-RUNNER.md). Close all Excel instances before using that runner.
