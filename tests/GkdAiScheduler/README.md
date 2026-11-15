# GkdAiScheduler tests and integration contract

Run behavioral tests (no network and no production backend dependencies):

```sh
dotnet run --project tests/GkdAiScheduler/GkdAiScheduler.Tests.csproj -c Release
```

Compile the production source alone against .NET Framework 4.8 / C# 7.3:

```sh
dotnet build tests/GkdAiScheduler/Compatibility/GkdAiScheduler.Net48.csproj -c Release
```

The internal sealed `Outline.Addin.GkdAiScheduler` implements `IDisposable`:

```csharp
public GkdAiScheduler(int concurrency = 3, int cacheLimit = 256);
public Task<object> RunAsync(string contentKey,
    Func<CancellationToken, Task<object>> work, CancellationToken consumerToken);
public void CancelAll();
public void ClearCache();
public void Invalidate(string key);
public void Dispose();
```

- Keys use ordinal equality. The caller must include profile/settings identity, prompt, input data, output mode, and caller-managed revision. No key construction or Excel/backend dependencies live here.
- Each content key has one deduplicated operation, with independent subscriber tasks. Cancellation of the last subscriber removes the entry and requests backend cancellation. Explicit future calls may retry; no canceled operation retries automatically.
- `CancelAll` cancels the current running/waiting subscribers, including entries detached by invalidation. It preserves completed cache successes. Caller-managed revision changes prevent old formula generations from being submitted again.
- Backend cancellation callbacks run on the thread pool, outside scheduler locks. Consumer tasks are canceled before those callbacks run, so callback latency/failure cannot hold up `CancelAll` or consumer cancellation. Backend token cancellation is consequently asynchronous.
- Semaphore waits are admitted under the scheduler lock; all waits use `WaitAsync`, keeping admission ordered. A delegate that ignores cancellation retains its slot until it exits, preserving the physical concurrency bound.
- `Invalidate` evicts one success and detaches inflight work for that key. Existing subscribers may still receive its result; later callers start fresh, and the old result cannot refill or overwrite the cache.
- `ClearCache` empties successes and detaches all existing inflight cache writes without canceling their subscribers.
- Successful results use a bounded LRU cache; `cacheLimit = 0` disables retention. Faults/cancellations are never cached. Results returned normally by a backend are successes; represent backend failures as exceptions rather than error-valued successes if they must be retryable.
- Every `object[,]` returned to a subscriber/cache consumer is a separate array, and the cache owns a separate snapshot. Cells are expected to be immutable Excel scalar values; nested mutable objects inside cells are not deep-cloned.
- `Dispose` is idempotent, rejects new live requests, cancels current subscribers, clears successes, and defers resource disposal until runners/cancellation callbacks finish. Cancellation-ignoring work cannot be forcibly terminated.
- Storage is in memory only. No user input/result persistence or logging is performed.

The tests use real async delegates and task gates for concurrency, duplicate subscriptions, subscriber cancellation, all-consumer cancellation, queued cancellation, global cancellation, stale-result suppression, exception retry, array isolation, LRU capacity, disposal, slow cancellation callbacks, parallel admission, and repeated cancellation/completion/invalidation/disposal races.
