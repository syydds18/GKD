using System;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using Outline.Addin;

namespace Outline.Addin { } // The first RED reports a missing scheduler explicitly.

static class Program
{
    static async Task Main()
    {
        await Run("success is reused by content key", SuccessIsCached);
        await Run("same-key concurrent callers share one operation", InflightIsDeduplicated);
        await Run("distinct work is bounded and queued in order", ConcurrencyIsBounded);
        await Run("canceling one subscriber preserves another", SubscriberCancellation);
        await Run("last subscriber cancels backend and permits fresh retry", LastSubscriberCancellation);
        await Run("CancelAll cancels running and waiting without replay", CancelAllStopsCurrentWork);
        await Run("cached and shared spill arrays are isolated", ArraysAreCloned);
        await Run("success cache evicts least recently used at capacity", CacheIsBounded);
        await Run("Invalidate detaches stale inflight results", InvalidatePreventsStaleCache);
        await Run("ClearCache also suppresses older inflight cache writes", ClearCacheSuppressesStaleWork);
        await Run("faults and backend cancellation are retryable", FailuresAreNotCached);
        await Run("pre-canceled consumer cannot receive cached success", PreCanceledConsumer);
        await Run("Dispose cancels detached and queued work and rejects reuse", DisposeStopsWork);
        await Run("slow backend cancellation callbacks cannot block CancelAll", SlowCancellationCallback);
        await Run("invalid configuration is rejected and zero cache disables reuse", ValidateConfiguration);
        await Run("parallel subscribers dedupe at default concurrency", ParallelAdmission);
        await Run("cancellation completion invalidation disposal race stress", CancellationRaceStress);
        await Run("queued cancellation never invokes abandoned delegate", QueuedCancellation);
        await Run("CancelAll preserves completed successes", CancelAllPreservesCache);
    }

    static async Task Run(string name, Func<Task> test)
    {
        await test().WaitAsync(TimeSpan.FromSeconds(10));
        Console.WriteLine("PASS " + name);
    }

    static async Task SuccessIsCached()
    {
        var type = Assembly.GetExecutingAssembly().GetType("Outline.Addin.GkdAiScheduler");
        Check(type != null, "GkdAiScheduler must exist");
        using (var scheduler = (IDisposable)Activator.CreateInstance(type, new object[] { 3, 256 }))
        {
            int calls = 0;
            Func<CancellationToken, Task<object>> work = async ct => { Interlocked.Increment(ref calls); await Task.Yield(); return "answer"; };
            var method = type.GetMethod("RunAsync");
            var first = (Task<object>)method.Invoke(scheduler, new object[] { "key", work, CancellationToken.None });
            Check((string)await first == "answer", "first result");
            var second = (Task<object>)method.Invoke(scheduler, new object[] { "key", work, CancellationToken.None });
            Check((string)await second == "answer", "cached result");
            Check(calls == 1, "successful work must be cached");
        }
    }

    static async Task InflightIsDeduplicated()
    {
        using (var scheduler = new GkdAiScheduler())
        {
            int calls = 0;
            var release = Signal();
            Func<CancellationToken, Task<object>> work = async ct => { Interlocked.Increment(ref calls); await release.Task; return "shared"; };
            var first = scheduler.RunAsync("same", work, CancellationToken.None);
            var second = scheduler.RunAsync("same", work, CancellationToken.None);
            release.TrySetResult(true);
            await Task.WhenAll(first, second);
            Check(calls == 1, "inflight duplicate invoked work " + calls + " times");
        }
    }

    static async Task ConcurrencyIsBounded()
    {
        using (var scheduler = new GkdAiScheduler(1))
        {
            int active = 0, maximum = 0;
            var order = new System.Collections.Generic.List<int>();
            var entered = Signal();
            var release = Signal();
            Func<int, Func<CancellationToken, Task<object>>> work = index => async ct =>
            {
                var count = Interlocked.Increment(ref active);
                lock (order) { maximum = Math.Max(maximum, count); order.Add(index); }
                entered.TrySetResult(true);
                await release.Task;
                Interlocked.Decrement(ref active);
                return index;
            };
            var first = scheduler.RunAsync("1", work(1), CancellationToken.None);
            await entered.Task;
            var second = scheduler.RunAsync("2", work(2), CancellationToken.None);
            var third = scheduler.RunAsync("3", work(3), CancellationToken.None);
            await Task.Delay(40);
            lock (order) Check(order.Count == 1, "queued delegates must not start before permit");
            release.TrySetResult(true);
            await Task.WhenAll(first, second, third);
            Check(maximum == 1, "concurrency exceeded configured maximum");
            Check(string.Join(",", order) == "1,2,3", "FIFO admission");
        }
    }

    static async Task SubscriberCancellation()
    {
        using (var scheduler = new GkdAiScheduler())
        using (var consumer = new CancellationTokenSource())
        {
            var entered = Signal(); var release = Signal();
            CancellationToken backendToken = default(CancellationToken);
            int calls = 0;
            Func<CancellationToken, Task<object>> work = async ct =>
            {
                Interlocked.Increment(ref calls); backendToken = ct; entered.TrySetResult(true);
                await release.Task; return "survivor";
            };
            var first = scheduler.RunAsync("shared", work, consumer.Token);
            await entered.Task;
            var second = scheduler.RunAsync("shared", work, CancellationToken.None);
            consumer.Cancel();
            await ExpectCanceled(first);
            Check(!backendToken.IsCancellationRequested, "one canceled subscriber canceled shared backend");
            release.TrySetResult(true);
            Check((string)await second == "survivor", "remaining subscriber result");
            Check(calls == 1, "shared work invoked twice");
        }
    }

    static async Task LastSubscriberCancellation()
    {
        using (var scheduler = new GkdAiScheduler(1))
        using (var consumer = new CancellationTokenSource())
        {
            var entered = Signal(); var canceled = Signal();
            var task = scheduler.RunAsync("last", async ct =>
            {
                entered.TrySetResult(true);
                using (ct.Register(() => canceled.TrySetResult(true))) await Task.Delay(Timeout.Infinite, ct);
                return "unreachable";
            }, consumer.Token);
            await entered.Task; consumer.Cancel();
            await ExpectCanceled(task); await canceled.Task;
            Check((string)await scheduler.RunAsync("last", ct => Task.FromResult<object>("retry"), CancellationToken.None) == "retry", "canceled key must be retryable");
        }
    }

    static async Task CancelAllStopsCurrentWork()
    {
        using (var scheduler = new GkdAiScheduler(1))
        {
            var entered = Signal(); var release = Signal(); var backendCanceled = Signal();
            int starts = 0; CancellationToken backendToken = default(CancellationToken);
            var running = scheduler.RunAsync("running", async ct =>
            {
                backendToken = ct; ct.Register(() => backendCanceled.TrySetResult(true));
                Interlocked.Increment(ref starts); entered.TrySetResult(true);
                await release.Task; return "stale"; // Deliberately ignores cancellation.
            }, CancellationToken.None);
            await entered.Task;
            var waiting = scheduler.RunAsync("waiting", ct => { Interlocked.Increment(ref starts); return Task.FromResult<object>("queued"); }, CancellationToken.None);
            var cancelAll = typeof(GkdAiScheduler).GetMethod("CancelAll");
            Check(cancelAll != null, "CancelAll must exist");
            cancelAll.Invoke(scheduler, null);
            await ExpectCanceled(running); await ExpectCanceled(waiting);
            await backendCanceled.Task;
            Check(backendToken.IsCancellationRequested, "backend token canceled");
            release.TrySetResult(true);
            Check((string)await scheduler.RunAsync("running", ct => Task.FromResult<object>("fresh"), CancellationToken.None) == "fresh", "late result must not repopulate cache");
            Check(starts == 1, "canceled queued delegate restarted");
        }
    }

    static async Task ArraysAreCloned()
    {
        using (var scheduler = new GkdAiScheduler())
        {
            var source = new object[,] { { "original", 42.0 } };
            var release = Signal();
            Func<CancellationToken, Task<object>> work = async ct => { await release.Task; return source; };
            var a = scheduler.RunAsync("array", work, CancellationToken.None);
            var b = scheduler.RunAsync("array", work, CancellationToken.None);
            release.TrySetResult(true);
            var first = (object[,])await a;
            var second = (object[,])await b;
            first[0, 0] = "mutated"; source[0, 1] = -1;
            Check((string)second[0, 0] == "original", "subscriber arrays share mutable cells");
            var cached = (object[,])await scheduler.RunAsync("array", work, CancellationToken.None);
            Check((string)cached[0, 0] == "original" && (double)cached[0, 1] == 42.0, "backend or consumer mutated cached array");
            Check(!ReferenceEquals(first, second) && !ReferenceEquals(second, cached), "each consumer needs its own array");
        }
    }

    static async Task CacheIsBounded()
    {
        using (var scheduler = new GkdAiScheduler(3, 2))
        {
            int calls = 0;
            Func<CancellationToken, Task<object>> work = ct => Task.FromResult<object>(Interlocked.Increment(ref calls));
            await scheduler.RunAsync("a", work, CancellationToken.None);
            await scheduler.RunAsync("b", work, CancellationToken.None);
            await scheduler.RunAsync("a", work, CancellationToken.None); // Touch a; evict b next.
            await scheduler.RunAsync("c", work, CancellationToken.None);
            await scheduler.RunAsync("a", work, CancellationToken.None);
            await scheduler.RunAsync("b", work, CancellationToken.None);
            Check(calls == 4, "bounded LRU should rerun evicted b; calls=" + calls);
        }
    }

    static async Task InvalidatePreventsStaleCache()
    {
        using (var scheduler = new GkdAiScheduler(2))
        {
            var entered = Signal(); var release = Signal();
            var old = scheduler.RunAsync("refresh", async ct => { entered.TrySetResult(true); await release.Task; return "old"; }, CancellationToken.None);
            await entered.Task;
            var invalidate = typeof(GkdAiScheduler).GetMethod("Invalidate");
            Check(invalidate != null, "Invalidate must exist");
            invalidate.Invoke(scheduler, new object[] { "refresh" });
            var fresh = await scheduler.RunAsync("refresh", ct => Task.FromResult<object>("fresh"), CancellationToken.None);
            Check((string)fresh == "fresh", "refresh joined invalidated work");
            release.TrySetResult(true);
            Check((string)await old == "old", "invalidation must preserve existing subscribers");
            Check((string)await scheduler.RunAsync("refresh", ct => Task.FromResult<object>("unexpected"), CancellationToken.None) == "fresh", "stale completion overwrote refreshed cache");
            invalidate.Invoke(scheduler, new object[] { "refresh" });
            Check((string)await scheduler.RunAsync("refresh", ct => Task.FromResult<object>("evicted"), CancellationToken.None) == "evicted", "invalidate did not evict success");
        }
    }

    static async Task ClearCacheSuppressesStaleWork()
    {
        using (var scheduler = new GkdAiScheduler(2))
        {
            var entered = Signal(); var release = Signal();
            await scheduler.RunAsync("cached", ct => Task.FromResult<object>("cached"), CancellationToken.None);
            var pending = scheduler.RunAsync("pending", async ct => { entered.TrySetResult(true); await release.Task; return "old"; }, CancellationToken.None);
            await entered.Task;
            var clear = typeof(GkdAiScheduler).GetMethod("ClearCache");
            Check(clear != null, "ClearCache must exist");
            clear.Invoke(scheduler, null);
            Check((string)await scheduler.RunAsync("cached", ct => Task.FromResult<object>("new"), CancellationToken.None) == "new", "success cache was not cleared");
            release.TrySetResult(true); await pending;
            Check((string)await scheduler.RunAsync("pending", ct => Task.FromResult<object>("new"), CancellationToken.None) == "new", "old inflight repopulated cleared cache");
        }
    }

    static async Task FailuresAreNotCached()
    {
        using (var scheduler = new GkdAiScheduler(1))
        {
            int calls = 0;
            Func<CancellationToken, Task<object>> work = async ct =>
            {
                await Task.Yield();
                if (Interlocked.Increment(ref calls) == 1) throw new InvalidOperationException("backend failure");
                return "recovered";
            };
            try { await scheduler.RunAsync("fault", work, CancellationToken.None); throw new Exception("failure expected"); }
            catch (InvalidOperationException e) { Check(e.Message == "backend failure", "original failure preserved"); }
            Check((string)await scheduler.RunAsync("fault", work, CancellationToken.None) == "recovered", "fault cached");
            var canceled = scheduler.RunAsync("canceled", ct => Task.FromCanceled<object>(new CancellationToken(true)), CancellationToken.None);
            await ExpectCanceled(canceled);
            Check((string)await scheduler.RunAsync("canceled", ct => Task.FromResult<object>("retry"), CancellationToken.None) == "retry", "backend cancellation cached");
        }
    }

    static async Task PreCanceledConsumer()
    {
        using (var scheduler = new GkdAiScheduler())
        {
            int calls = 0;
            Func<CancellationToken, Task<object>> work = ct => Task.FromResult<object>(Interlocked.Increment(ref calls));
            await scheduler.RunAsync("cached", work, CancellationToken.None);
            await ExpectCanceled(scheduler.RunAsync("cached", work, new CancellationToken(true)));
            await ExpectCanceled(scheduler.RunAsync("never", work, new CancellationToken(true)));
            Check(calls == 1, "pre-canceled consumer started work");
        }
    }

    static async Task DisposeStopsWork()
    {
        var scheduler = new GkdAiScheduler(1);
        var entered = Signal(); var release = Signal();
        var running = scheduler.RunAsync("disposed", async ct => { entered.TrySetResult(true); await release.Task; return "old"; }, CancellationToken.None);
        await entered.Task;
        scheduler.Invalidate("disposed"); // Still must be canceled after it leaves dedupe map.
        var queued = scheduler.RunAsync("queued", ct => Task.FromResult<object>("unexpected"), CancellationToken.None);
        scheduler.Dispose(); scheduler.Dispose();
        await ExpectCanceled(running); await ExpectCanceled(queued);
        bool rejected = false;
        try { await scheduler.RunAsync("new", ct => Task.FromResult<object>("bad"), CancellationToken.None); }
        catch (ObjectDisposedException) { rejected = true; }
        Check(rejected, "disposed scheduler accepted new work");
        release.TrySetResult(true);
    }

    static async Task SlowCancellationCallback()
    {
        using (var scheduler = new GkdAiScheduler(1))
        using (var callbackRelease = new ManualResetEventSlim())
        {
            var entered = Signal(); var callbackEntered = Signal(); var release = Signal();
            var running = scheduler.RunAsync("slow-cancel", async ct =>
            {
                using (ct.Register(() => { callbackEntered.TrySetResult(true); callbackRelease.Wait(); }))
                {
                    entered.TrySetResult(true); await release.Task; return "ignored";
                }
            }, CancellationToken.None);
            await entered.Task;
            var cancellation = Task.Run(() => scheduler.CancelAll());
            try
            {
                await callbackEntered.Task;
                await ExpectCanceled(running);
                await cancellation.WaitAsync(TimeSpan.FromMilliseconds(500));
            }
            finally { callbackRelease.Set(); release.TrySetResult(true); }
        }
    }

    static async Task ValidateConfiguration()
    {
        Throws<ArgumentOutOfRangeException>(() => new GkdAiScheduler(0));
        Throws<ArgumentOutOfRangeException>(() => new GkdAiScheduler(-1));
        Throws<ArgumentOutOfRangeException>(() => new GkdAiScheduler(1, -1));
        using (var scheduler = new GkdAiScheduler(1, 0))
        {
            int calls = 0;
            Func<CancellationToken, Task<object>> work = ct => Task.FromResult<object>(Interlocked.Increment(ref calls));
            Throws<ArgumentNullException>(() => scheduler.RunAsync(null, work, CancellationToken.None));
            Throws<ArgumentNullException>(() => scheduler.RunAsync("key", null, CancellationToken.None));
            await scheduler.RunAsync("key", work, CancellationToken.None);
            await scheduler.RunAsync("key", work, CancellationToken.None);
            Check(calls == 2, "zero cache still caches success");
        }
    }

    static async Task ParallelAdmission()
    {
        using (var scheduler = new GkdAiScheduler())
        {
            var release = Signal(); var allStarted = Signal();
            int starts = 0, active = 0, maximum = 0;
            var tasks = new Task<object>[60];
            Func<CancellationToken, Task<object>> work = async ct =>
            {
                var count = Interlocked.Increment(ref active);
                int observed;
                do { observed = Volatile.Read(ref maximum); }
                while (count > observed && Interlocked.CompareExchange(ref maximum, count, observed) != observed);
                if (Interlocked.Increment(ref starts) == 3) allStarted.TrySetResult(true);
                await release.Task;
                Interlocked.Decrement(ref active);
                return "shared";
            };
            var ready = new CountdownEvent(tasks.Length);
            var admission = new Task[tasks.Length];
            for (int i = 0; i < tasks.Length; i++)
            {
                int index = i;
                admission[i] = Task.Run(() =>
                {
                    tasks[index] = scheduler.RunAsync("key-" + (index % 12), work, CancellationToken.None);
                    ready.Signal();
                });
            }
            await Task.WhenAll(admission); await allStarted.Task;
            Check(ready.IsSet && starts == 3, "default scheduler must run exactly three delegates before release");
            release.TrySetResult(true);
            await Task.WhenAll(tasks);
            Check(starts == 12 && maximum == 3, "parallel admission broke dedupe or default concurrency");
            ready.Dispose();
        }
    }

    static async Task CancellationRaceStress()
    {
        for (int round = 0; round < 100; round++)
        {
            using (var scheduler = new GkdAiScheduler(2, 2))
            using (var consumer = new CancellationTokenSource())
            {
                var release = Signal(); var entered = Signal();
                var first = scheduler.RunAsync("race", async ct =>
                {
                    using (ct.Register(() => { if ((round % 3) == 0) throw new InvalidOperationException("callback failure"); }))
                    {
                        entered.TrySetResult(true); await release.Task; return "done";
                    }
                }, consumer.Token);
                await entered.Task;
                var second = scheduler.RunAsync("race", ct => Task.FromResult<object>("duplicate"), CancellationToken.None);
                var queued = scheduler.RunAsync("other", async ct => { await Task.Yield(); return "other"; }, consumer.Token);
                await Task.WhenAll(
                    Task.Run(() => consumer.Cancel()),
                    Task.Run(() => { scheduler.Invalidate("race"); scheduler.ClearCache(); }),
                    Task.Run(() => scheduler.CancelAll()),
                    Task.Run(() => { release.TrySetResult(true); scheduler.Dispose(); }));
                foreach (var task in new[] { first, second, queued })
                {
                    try { await task; }
                    catch (OperationCanceledException) { Check(task.IsCanceled, "cancellation faulted during race"); }
                }
            }
        }
    }

    static async Task QueuedCancellation()
    {
        using (var scheduler = new GkdAiScheduler(1))
        using (var consumer = new CancellationTokenSource())
        {
            var entered = Signal(); var release = Signal();
            int calls = 0;
            var running = scheduler.RunAsync("blocker", async ct => { entered.TrySetResult(true); await release.Task; return "blocker"; }, CancellationToken.None);
            await entered.Task;
            var queued = scheduler.RunAsync("queued", ct => { Interlocked.Increment(ref calls); return Task.FromResult<object>("bad"); }, consumer.Token);
            consumer.Cancel(); await ExpectCanceled(queued);
            var fresh = scheduler.RunAsync("queued", ct => Task.FromResult<object>("fresh"), CancellationToken.None);
            release.TrySetResult(true); await running;
            Check((string)await fresh == "fresh" && calls == 0, "abandoned queued delegate executed");
        }
    }

    static async Task CancelAllPreservesCache()
    {
        using (var scheduler = new GkdAiScheduler())
        {
            int calls = 0;
            Func<CancellationToken, Task<object>> work = ct => Task.FromResult<object>(Interlocked.Increment(ref calls));
            await scheduler.RunAsync("success", work, CancellationToken.None);
            scheduler.CancelAll();
            await scheduler.RunAsync("success", work, CancellationToken.None);
            Check(calls == 1, "CancelAll unexpectedly cleared completed cache");
        }
    }

    static void Throws<T>(Action action) where T : Exception
    {
        try { action(); } catch (T) { return; }
        throw new Exception("expected " + typeof(T).Name);
    }

    static async Task ExpectCanceled(Task<object> task)
    {
        try { await task.WaitAsync(TimeSpan.FromMilliseconds(500)); }
        catch (OperationCanceledException) { Check(task.IsCanceled, "task must be canceled, not faulted"); return; }
        throw new Exception("expected canceled task");
    }

    static TaskCompletionSource<bool> Signal() { return new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously); }
    static void Check(bool condition, string message) { if (!condition) throw new Exception(message); }
}
