using System;
using System.Collections.Concurrent;
using System.Diagnostics;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace OneFinder
{
    /// <summary>
    /// 在专用 STA 线程上序列化所有 OneNote COM 调用
    /// </summary>
    internal sealed class OneNoteScheduler : IDisposable
    {
        private readonly BlockingCollection<Action> _queue = new();
        private readonly Thread _thread;
        private OneNoteService? _service;
        private bool _disposed;

        public OneNoteScheduler()
        {
            _thread = new Thread(ThreadProc)
            {
                Name = "OneNote-STA",
                IsBackground = true,
            };
            _thread.SetApartmentState(ApartmentState.STA);
            _thread.Start();
        }

        public Task<T> Run<T>(Func<OneNoteService, T> func)
        {
            var tcs = new TaskCompletionSource<T>(TaskCreationOptions.RunContinuationsAsynchronously);
            _queue.Add(() =>
            {
                try
                {
                    EnsureService();
                    tcs.SetResult(func(_service!));
                }
                catch (OperationCanceledException ex)
                {
                    tcs.SetException(ex);
                }
                catch (Exception ex)
                {
                    InvalidateService();
                    tcs.SetException(ex);
                }
            });
            return tcs.Task;
        }

        public Task Run(Action<OneNoteService> action)
            => Run<bool>(svc => { action(svc); return true; });

        private void EnsureService()
        {
            if (_service != null && !IsOneNoteRunning())
                InvalidateService();

            _service ??= new OneNoteService();
        }

        private static bool IsOneNoteRunning()
            => Process.GetProcessesByName("ONENOTE").Any()
            || Process.GetProcessesByName("OneNote").Any();

        private void InvalidateService()
        {
            try { _service?.Dispose(); } catch { }
            _service = null;
        }

        private void ThreadProc()
        {
            foreach (var action in _queue.GetConsumingEnumerable())
                action();
        }

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;
            _queue.CompleteAdding();
            _thread.Join(millisecondsTimeout: 2000);
            _service?.Dispose();
            _queue.Dispose();
        }
    }
}
