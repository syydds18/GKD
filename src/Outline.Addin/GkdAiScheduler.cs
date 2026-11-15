using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;

namespace Outline.Addin
{
    /// <summary>Process-local AI results. Keys must include every input and profile revision.</summary>
    internal sealed class GkdAiScheduler : IDisposable
    {
        readonly object gate = new object();
        readonly Dictionary<string, CacheItem> cache = new Dictionary<string, CacheItem>(StringComparer.Ordinal);
        readonly LinkedList<string> recency = new LinkedList<string>();
        readonly int cacheLimit;
        readonly Dictionary<string, Entry> inflight = new Dictionary<string, Entry>(StringComparer.Ordinal);
        readonly HashSet<Entry> active = new HashSet<Entry>();
        readonly SemaphoreSlim slots;
        bool disposed, slotsDisposed;

        sealed class Entry
        {
            internal readonly string Key;
            internal readonly CancellationTokenSource Cancellation = new CancellationTokenSource();
            internal readonly HashSet<Subscriber> Subscribers = new HashSet<Subscriber>();
            internal bool Finished, WorkDone, CancelInProgress;
            internal bool CacheAllowed = true;
            internal Entry(string key) { Key = key; }
        }

        sealed class Subscriber
        {
            internal readonly TaskCompletionSource<object> Result = new TaskCompletionSource<object>(TaskCreationOptions.RunContinuationsAsynchronously);
        }

        sealed class CacheItem
        {
            internal object Value;
            internal LinkedListNode<string> Node;
        }

        public GkdAiScheduler(int concurrency = 3, int cacheLimit = 256)
        {
            if (concurrency < 1) throw new ArgumentOutOfRangeException(nameof(concurrency));
            if (cacheLimit < 0) throw new ArgumentOutOfRangeException(nameof(cacheLimit));
            slots = new SemaphoreSlim(concurrency, concurrency);
            this.cacheLimit = cacheLimit;
        }

        public Task<object> RunAsync(string contentKey, Func<CancellationToken, Task<object>> work, CancellationToken consumerToken)
        {
            if (contentKey == null) throw new ArgumentNullException(nameof(contentKey));
            if (work == null) throw new ArgumentNullException(nameof(work));
            if (consumerToken.IsCancellationRequested) return Task.FromCanceled<object>(consumerToken);
            Entry entry;
            Task permit = null;
            var subscriber = new Subscriber();
            lock (gate)
            {
                if (disposed) throw new ObjectDisposedException(nameof(GkdAiScheduler));
                CacheItem cached;
                if (cache.TryGetValue(contentKey, out cached))
                {
                    recency.Remove(cached.Node);
                    recency.AddLast(cached.Node);
                    return Task.FromResult(CloneResult(cached.Value));
                }
                if (!inflight.TryGetValue(contentKey, out entry))
                {
                    entry = new Entry(contentKey);
                    inflight.Add(contentKey, entry);
                    active.Add(entry);
                    // Enqueue under the same lock as dictionary admission, so worker scheduling
                    // cannot reorder semaphore waiters. All waiters use the async semaphore path.
                    permit = slots.WaitAsync(entry.Cancellation.Token);
                }
                entry.Subscribers.Add(subscriber);
            }
            var result = AwaitSubscriber(entry, subscriber, consumerToken);
            if (permit != null)
            {
                var admitted = permit;
                var ignored = Task.Run(() => Execute(entry, admitted, work));
            }
            return result;
        }

        async Task<object> AwaitSubscriber(Entry entry, Subscriber subscriber, CancellationToken token)
        {
            using (token.Register(() => CancelSubscriber(entry, subscriber)))
                return await subscriber.Result.Task.ConfigureAwait(false);
        }

        void CancelSubscriber(Entry entry, Subscriber subscriber)
        {
            bool cancel = false;
            lock (gate)
            {
                if (entry.Subscribers.Remove(subscriber))
                {
                    subscriber.Result.TrySetCanceled();
                    if (!entry.Finished && entry.Subscribers.Count == 0)
                    {
                        entry.Finished = true;
                        RemoveEntry(entry);
                        entry.CancelInProgress = true;
                        cancel = true;
                    }
                }
            }
            if (cancel) QueueCancellation(entry);
        }

        void QueueCancellation(Entry entry)
        {
            // Cancellation callbacks belong to the backend and may block or throw.
            // Consumer completion is already published before dispatching them.
            ThreadPool.QueueUserWorkItem(_ => CancelBackend(entry));
        }

        // Never invoke backend callbacks while holding gate. WorkDone/CancelInProgress
        // jointly own the CTS, so completion cannot dispose it during Cancel().
        void CancelBackend(Entry entry)
        {
            try { entry.Cancellation.Cancel(); }
            catch (AggregateException) { /* A backend cancellation callback must not break cleanup. */ }
            finally
            {
                lock (gate)
                {
                    entry.CancelInProgress = false;
                    if (entry.WorkDone)
                    {
                        active.Remove(entry);
                        entry.Cancellation.Dispose();
                    }
                    DisposeSlotsIfIdle();
                }
            }
        }

        async Task Execute(Entry entry, Task permit, Func<CancellationToken, Task<object>> work)
        {
            bool acquired = false;
            try
            {
                await permit.ConfigureAwait(false);
                acquired = true;
                lock (gate) if (entry.Finished) return;
                entry.Cancellation.Token.ThrowIfCancellationRequested();
                var value = await work(entry.Cancellation.Token).ConfigureAwait(false);
                lock (gate)
                {
                    if (!entry.Finished)
                    {
                        if (entry.CacheAllowed) CacheSuccess(entry.Key, value);
                        Complete(entry, value, null, false);
                    }
                }
            }
            catch (OperationCanceledException) { lock (gate) Complete(entry, null, null, true); }
            catch (Exception error) { lock (gate) Complete(entry, null, error, false); }
            finally
            {
                if (acquired) slots.Release();
                lock (gate)
                {
                    entry.WorkDone = true;
                    if (!entry.CancelInProgress)
                    {
                        active.Remove(entry);
                        entry.Cancellation.Dispose();
                    }
                    DisposeSlotsIfIdle();
                }
            }
        }

        void Complete(Entry entry, object value, Exception error, bool canceled)
        {
            if (entry.Finished) return;
            entry.Finished = true;
            RemoveEntry(entry);
            foreach (var subscriber in entry.Subscribers)
            {
                if (canceled) subscriber.Result.TrySetCanceled();
                else if (error != null) subscriber.Result.TrySetException(error);
                else subscriber.Result.TrySetResult(CloneResult(value));
            }
            entry.Subscribers.Clear();
        }

        void CacheSuccess(string key, object value)
        {
            if (cacheLimit == 0) return;
            CacheItem previous;
            if (cache.TryGetValue(key, out previous)) recency.Remove(previous.Node);
            cache[key] = new CacheItem { Value = CloneResult(value), Node = recency.AddLast(key) };
            while (cache.Count > cacheLimit)
            {
                cache.Remove(recency.First.Value);
                recency.RemoveFirst();
            }
        }

        static object CloneResult(object value)
        {
            var grid = value as object[,];
            return grid == null ? value : grid.Clone();
        }

        void RemoveEntry(Entry entry)
        {
            Entry current;
            if (inflight.TryGetValue(entry.Key, out current) && ReferenceEquals(current, entry)) inflight.Remove(entry.Key);
        }

        public void CancelAll()
        {
            List<Entry> canceled;
            lock (gate)
            {
                canceled = CancelActive();
            }
            foreach (var entry in canceled) QueueCancellation(entry);
        }

        List<Entry> CancelActive()
        {
            var canceled = new List<Entry>();
            foreach (var entry in active) if (!entry.Finished) canceled.Add(entry);
            foreach (var entry in canceled)
            {
                Complete(entry, null, null, true);
                entry.CancelInProgress = true;
            }
            return canceled;
        }

        public void Invalidate(string key)
        {
            if (key == null) throw new ArgumentNullException(nameof(key));
            lock (gate)
            {
                CacheItem cached;
                if (cache.TryGetValue(key, out cached))
                {
                    cache.Remove(key);
                    recency.Remove(cached.Node);
                }
                Entry entry;
                if (inflight.TryGetValue(key, out entry))
                {
                    entry.CacheAllowed = false;
                    inflight.Remove(key);
                }
            }
        }

        public void ClearCache()
        {
            lock (gate)
            {
                cache.Clear();
                recency.Clear();
                foreach (var entry in active) entry.CacheAllowed = false;
                inflight.Clear();
            }
        }

        public void Dispose()
        {
            List<Entry> canceled;
            lock (gate)
            {
                if (disposed) return;
                disposed = true;
                cache.Clear();
                recency.Clear();
                canceled = CancelActive();
                DisposeSlotsIfIdle();
            }
            foreach (var entry in canceled) QueueCancellation(entry);
        }

        // A cancellation-ignoring delegate retains its permit until it really exits.
        // Defer semaphore disposal until every runner has released its permit.
        void DisposeSlotsIfIdle()
        {
            if (disposed && !slotsDisposed && active.Count == 0)
            {
                slotsDisposed = true;
                slots.Dispose();
            }
        }
    }
}
