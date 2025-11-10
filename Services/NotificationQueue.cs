using System.Collections.Concurrent;

namespace ClipDiscordApp.Services
{
    // シンプルな送信キュー：非同期ワーカー、重複抑止、簡易再試行
    public class NotificationQueue : IDisposable
    {
        private readonly Func<NotificationPayload, CancellationToken, Task<bool>> _sendFunc;
        private readonly ConcurrentQueue<NotificationPayload> _queue = new();
        private readonly CancellationTokenSource _cts = new();
        private readonly ConcurrentDictionary<string, DateTime> _recentKeys = new();
        private readonly Task _worker;

        // duplicateWindow: 同一キーを再送しない時間ウィンドウ
        public NotificationQueue(Func<NotificationPayload, CancellationToken, Task<bool>> sendFunc, TimeSpan? duplicateWindow = null)
        {
            _sendFunc = sendFunc ?? throw new ArgumentNullException(nameof(sendFunc));
            DuplicateWindow = duplicateWindow ?? TimeSpan.FromSeconds(10);
            _worker = Task.Run(WorkerLoopAsync);
        }

        public TimeSpan DuplicateWindow { get; }

        public void Enqueue(NotificationPayload p)
        {
            if (p == null) return;
            var key = MakeKey(p);
            var now = DateTime.UtcNow;

            // 重複チェック（簡易）
            if (_recentKeys.TryGetValue(key, out var t))
            {
                if (now - t < DuplicateWindow) return;
            }

            _recentKeys[key] = now;
            _queue.Enqueue(p);
        }

        private static string MakeKey(NotificationPayload p)
        {
            return (p.Title ?? "") + "|" + (p.Body ?? "");
        }

        private async Task WorkerLoopAsync()
        {
            var token = _cts.Token;
            while (!token.IsCancellationRequested)
            {
                try
                {
                    if (_queue.TryDequeue(out var item))
                    {
                        bool ok = false;
                        try
                        {
                            ok = await _sendFunc(item, token).ConfigureAwait(false);
                        }
                        catch (OperationCanceledException) { throw; }
                        catch
                        {
                            ok = false;
                        }

                        if (!ok)
                        {
                            // 再試行（単純な遅延を入れて再キュー）
                            await Task.Delay(1000, token).ConfigureAwait(false);
                            _queue.Enqueue(item);
                        }
                        // レート調整
                        await Task.Delay(300, token).ConfigureAwait(false);
                    }
                    else
                    {
                        await Task.Delay(200, token).ConfigureAwait(false);
                    }
                }
                catch (OperationCanceledException) { break; }
                catch { await Task.Delay(500, token).ConfigureAwait(false); }
            }
        }

        public void Dispose()
        {
            _cts.Cancel();
            try { _worker.Wait(2000); } catch { }
            _cts.Dispose();
        }
    }

    // 簡易 DTO（既にプロジェクトにあれば重複しないよう注意）
    public class NotificationPayload
    {
        public string Title { get; set; } = "";
        public string Body { get; set; } = "";
    }
}