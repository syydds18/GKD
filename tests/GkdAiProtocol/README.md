# GKD single-turn AI contract tests

Run from the repository root (replace `dotnet` with the local SDK executable when needed):

```sh
dotnet run --project tests/GkdAiProtocol/GkdAiProtocol.Tests.csproj -c Release
dotnet build tests/GkdAiProtocol/Compatibility/GkdAiCompatibility.csproj -c Release
```

Default execution makes HTTP requests only to an ephemeral local HTTP fixture. It covers both providers' request bodies, authorization, custom Dify names and optional route, scalar/table output, full request latency, provider/HTTP failure sanitization, redirect refusal, normal/chunked oversized responses, timeout, and cancellation. Separate checks cover profile validation, typed input matrices, strict JSON and rectangular tables, final-cell limits, and settings persistence/corruption.

Optional live execution requires an existing isolated Dify app key file containing a `token` property. Keys are read directly into memory and never copied into this repository or printed:

```sh
dotnet run --project tests/GkdAiProtocol/GkdAiProtocol.Tests.csproj -c Release -- --live /absolute/path/to/key.json
# Or set GKD_DIFY_KEY_FILE and use --live without a file argument.
```

Live mode uses the existing `https://aibuilder.hdec.com/v1/workflows/run` endpoint, route `hdec235`, inputs `question`/`route`, output `answer`. It has a shared 120-second deadline and checks scalar 42 and an exact synthetic titled table. It never edits workflows. The default suite does not run it.

## Integration API and caveats

- `GkdAiProfile` and `GkdAiSettings` are internal sealed classes with public properties. Profile `Clone()` snapshots configuration; `Validate()` requires a usable API key. Save permits an empty key for unfinished placeholder profiles. The default factory profile is a Dify endpoint placeholder without a key.
- `GkdAiStore.Load()`, `Save(settings)`, and `ResolveDefault()` use `%LOCALAPPDATA%/GKD/ai-settings.json`. `ResolveDefault(settings)` returns a cloned profile. Explicit `Load(path, Func<string,string> decrypt)` and `Save(settings,path,Func<string,string> encrypt)` overloads exist for tests; production uses Windows DPAPI CurrentUser under `NETFRAMEWORK` and never falls back to plaintext. The net8 portable suite injects its own reversible codec and does not prove Windows DPAPI runtime behavior.
- Store files have version 1 and a SHA-256 checksum of their serialized encrypted payload. The checksum detects corruption; it is not a signature against a malicious local editor. DPAPI protects key ciphertext. Writes use a same-directory temporary file, flush, and atomic `File.Replace`/`File.Move`. Duplicate profile IDs (case-insensitive), invalid/missing defaults and unsupported versions are rejected. No automatic legacy `ai-workflow.json` migration occurs.
- `NormalizeMode(null/empty)` returns `单值`; explicit modes are `单值` and `表格`. `BuildPrompt(prompt,object[,],mode)` accepts only explicit matrix data, never infers workbook references, and includes a Chinese instruction/data boundary. The input JSON contains `prompt`, `mode`, and typed `data`; the requested model table answer is strictly `{"title":"表名","data":[["表头"],["数据"]]}`. Prompt injection resistance is a model instruction, not a guarantee.
- Input bounds: 1–10000 cells, prompt at most 16000 UTF-16 characters, complete prompt at most 262144 characters. Null/DBNull become JSON null; strings/numbers/bools retain type; DateTime/DateTimeOffset become ISO strings. Other types and nonfinite numbers are rejected. Excel-specific empty/error objects must be normalized by the UDF bridge before passing them here. Excel Value2 numeric dates remain numeric because number formats are not part of this API.
- `ParseAnswer(text,mode)` removes closed `<think>...</think>` blocks. Scalars are trimmed strings. Table responses require a nonempty string `title` and a nonempty rectangular `data` matrix, with no extra top-level properties. Optional JSON code fences are accepted; bare arrays and untitled legacy objects are rejected. Strict JSON rejects comments, trailing commas, duplicates, nested cell objects/arrays and nonfinite numbers. Titles are trimmed; a supplied header row is kept unchanged, never invented.
- Table return values are `object[,]`: row 0 has title only in column 0 and empty strings elsewhere; row 1 is entirely empty strings; row 2 contains source headers; following rows contain data. Null table cells become empty strings. At most **10000 final cells, including both layout rows**, are allowed. Formula-looking strings remain literal strings; these functions never evaluate formulas or call Excel COM. Styling/merging is the parent's UI-thread responsibility after spill completion.
- `GkdAiClient.ExecuteAsync(profile,prompt,data,mode,CancellationToken)` returns `Task<GkdAiReply>` with `Value` and total `ElapsedMilliseconds`. OpenAI uses the full configured chat/completions URL, one user message, `stream:false`, model and bearer key; it reads string `choices[0].message.content`. If present, finish_reason must be `stop`; structured content arrays, streaming, tools, conversation IDs and multi-turn state are unsupported.
- Dify uses blocking workflow run, `inputs[PromptInput]`, optional `inputs[RouteInput]` only when both route field and Model are nonempty, and fixed non-personal `user:"gkd-excel"`. It requires `data.status:"succeeded"` and a string at `data.outputs[AnswerOutput]`. Failed/incomplete workflows are rejected even on HTTP 200. Dify structured-object output is not accepted; configure answer as text containing JSON for tables.
- HTTPS is required outside loopback HTTP development. URL userinfo/query/fragment/control characters are rejected; endpoint is used as supplied without appending paths. Redirects are disabled and never followed. HTTP response size is limited to 1 MiB including chunked bodies; answer text to 262144 characters. No retry is performed. Timeout is configurable 1–600 seconds; external cancellation throws OperationCanceledException, deadline throws TimeoutException. Other client errors are bounded Chinese InvalidDataException messages without original exceptions, keys, prompts, URLs or response bodies.
- The parent add-in project needs Newtonsoft.Json 13.0.3 and System.Security. This task changes no add-in project file, Ribbon, ExcelBridge, old AiWorkflow files, native formatting or installation flow. The compatibility project compiles the four production sources as net48/C#7.3 separately; macOS cannot validate a Windows Excel host.

## Observed TDD/verification evidence

New profile/protocol/store/client slices initially failed compilation because the requested APIs were absent, then passed after their implementations. Subsequent executable RED tests exposed the titled-table contract change (7 passing groups, 2 failing), permissive JSON syntax (8/1), encrypted-key length validation (8/1), and endpoint control normalization (8/1); each was fixed and rerun green. The final default suite reports 9 groups passed, 0 failed.

Authorized live Dify execution returned scalar 42 in 1139 ms and a verified 5x2 title/blank/header/data result in 1202 ms. Those are observed timings from one run, not performance guarantees.
