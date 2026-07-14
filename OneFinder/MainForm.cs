using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Runtime.InteropServices;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OneFinder
{
    public partial class MainForm : Form
    {
        // DWM API
        [DllImport("dwmapi.dll")]
        private static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int attrValue, int attrSize);

        private const int DWMWA_USE_IMMERSIVE_DARK_MODE = 20;
        private const int DWMWA_CAPTION_COLOR = 35;
        private const int DWMWA_BORDER_COLOR = 34;

        // Foreground activation helpers
        [DllImport("user32.dll")] private static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
        [DllImport("kernel32.dll")] private static extern uint GetCurrentThreadId();
        [DllImport("user32.dll")] private static extern bool AttachThreadInput(uint idAttach, uint idAttachTo, bool fAttach);
        [DllImport("user32.dll")] private static extern bool SetForegroundWindow(IntPtr hWnd);
        [DllImport("user32.dll")] private static extern bool BringWindowToTop(IntPtr hWnd);
        [DllImport("user32.dll")] private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
        private const int SW_RESTORE = 9;

        // Color Scheme
        public static class ModernColors
        {
            public static readonly Color Primary = Color.FromArgb(128, 57, 123);
            public static readonly Color PrimaryDark = Color.FromArgb(102, 45, 98);
            public static readonly Color Accent = Color.FromArgb(185, 85, 211);
            public static readonly Color Background = Color.FromArgb(246, 246, 246);
            public static readonly Color CardBackground = Color.White;
            public static readonly Color TextPrimary = Color.FromArgb(33, 33, 33);
            public static readonly Color TextSecondary = Color.FromArgb(97, 97, 97);
            public static readonly Color TextHint = Color.FromArgb(158, 158, 158);
            public static readonly Color Divider = Color.FromArgb(224, 224, 224);
            public static readonly Color Highlight = Color.FromArgb(0, 0, 0);
            public static readonly Color HighlightBg = Color.FromArgb(255, 242, 0);
            public static readonly Color SelectionBg = Color.FromArgb(240, 230, 250);
            public static readonly Color StatusBorder = Color.FromArgb(230, 230, 230);
        }

        private ModernTextBox   _searchBox    = null!;
        private ModernButton    _searchButton = null!;
        private Label           _pinButton    = null!;
        private CheckBox        _currentNotebookOnly = null!;
        private ListBox         _resultList   = null!;
        private Label           _statusLabel  = null!;
        private ProgressBar     _progress     = null!;

        private List<MatchResult> _currentResults = new();
        private CancellationTokenSource? _cts;
        private int _searchVersion;
        private readonly OneNoteScheduler _scheduler = new();
        private readonly CancellationTokenSource _shutdownCts = new();

        // 预览加载刷新限流
        private readonly System.Windows.Forms.Timer _previewRefreshTimer = new() { Interval = 1000 };
        private bool _previewDirty;

        public MainForm()
        {
            InitializeComponent();
            BuildModernUI();

            // 恢复窗口置顶状态
            var saved = WindowSizeStore.Load();
            if (saved != null)
            {
                this.TopMost = saved.Value.TopMost;
                UpdatePinButtonState();
            }

            this.Paint += MainForm_Paint;

            this.HandleCreated += (s, e) =>
            {
                ApplyPurpleTitleBar();
                ListenForOneNoteShutdown();
            };

            this.Shown += (s, e) => { ForceActivate(); LoadRecentPages(); };

            // 预览加载刷新限流定时器
            _previewRefreshTimer.Tick += (s, e) =>
            {
                if (_previewDirty)
                {
                    _previewDirty = false;
                    _resultList.Invalidate();
                }
            };
            _previewRefreshTimer.Start();

            // 关闭时释放 STA 线程和 COM 连接
            this.FormClosed += (s, e) =>
            {
                WindowSizeStore.Save(this.Width, this.Height, this.TopMost);
                _cts?.Cancel();
                _shutdownCts.Cancel();
                _scheduler.Dispose();
            };
        }

        /// <summary>
        /// 后台等待 OneNote 退出信号并关闭 OneFinder
        /// </summary>
        private void ListenForOneNoteShutdown()
        {
            // Signal 0: OneFinder-Activate      — 已运行时再次点击按钮，置顶窗口（AutoReset，可重复触发）
            // Signal 1: OneFinder-ReleaseCOM    — OneNote 即将退出，立即释放 COM 对象（仅处理一次）
            // Signal 2: OneFinder-OneNoteShutdown — OneNote 已开始关闭，关闭 OneFinder 窗口
            var activateEvent = new EventWaitHandle(
                initialState: false,
                mode: EventResetMode.AutoReset,
                name: "Local\\OneFinder-Activate");

            var releaseEvent = new EventWaitHandle(
                initialState: false,
                mode: EventResetMode.ManualReset,
                name: "Local\\OneFinder-ReleaseCOM");

            var shutdownEvent = new EventWaitHandle(
                initialState: false,
                mode: EventResetMode.ManualReset,
                name: "Local\\OneFinder-OneNoteShutdown");

            var token = _shutdownCts.Token;
            System.Threading.Thread listener = new(() =>
            {
                try
                {
                    OneNoteService.Log("[MainForm] Shutdown listener started");

                    WaitHandle[] phase1Handles = { activateEvent, releaseEvent, shutdownEvent, token.WaitHandle };

                    // Phase 1: handle Activate (repeatable) and wait for ReleaseCOM/Shutdown/Cancel
                    while (true)
                    {
                        int idx = WaitHandle.WaitAny(phase1Handles);
                        OneNoteService.Log($"[MainForm] Phase1 WaitAny idx={idx}");

                        if (token.IsCancellationRequested)
                        {
                            OneNoteService.Log("[MainForm] Cancelled in Phase1, listener exiting");
                            return;
                        }

                        if (idx == 0) // Activate: bring window to front
                        {
                            OneNoteService.Log("[MainForm] Activate signal received");
                            BeginInvoke((Action)ForceActivate);
                            continue; // AutoReset — safe to loop
                        }

                        if (idx == 1) // ReleaseCOM: release COM then move to Phase 2
                        {
                            OneNoteService.Log("[MainForm] ReleaseCOM signal received, releasing COM...");
                            _ = _scheduler.ReleaseCom().ContinueWith(t =>
                                OneNoteService.Log($"[MainForm] ReleaseCom task completed, faulted={t.IsFaulted}"));
                            break; // proceed to phase 2
                        }

                        if (idx == 2) // Shutdown arrived without ReleaseCOM
                        {
                            OneNoteService.Log("[MainForm] Shutdown signal received in Phase1, closing window");
                            BeginInvoke(Close);
                            return;
                        }
                    }

                    // Phase 2: ReleaseCOM handled — wait only for shutdown or cancel
                    OneNoteService.Log("[MainForm] Phase2: waiting for shutdown signal...");
                    int idx2 = WaitHandle.WaitAny(new WaitHandle[] { shutdownEvent, token.WaitHandle });
                    OneNoteService.Log($"[MainForm] Phase2 WaitAny idx={idx2}");

                    if (idx2 == 0 && !token.IsCancellationRequested)
                    {
                        OneNoteService.Log("[MainForm] Shutdown signal received in Phase2, closing window");
                        BeginInvoke(Close);
                    }
                    else
                    {
                        OneNoteService.Log("[MainForm] Cancelled in Phase2, listener exiting");
                    }
                }
                finally
                {
                    activateEvent.Dispose();
                    releaseEvent.Dispose();
                    shutdownEvent.Dispose();
                    OneNoteService.Log("[MainForm] Shutdown listener exited");
                }
            })
            {
                IsBackground = true,
                Name = "OneNote-Shutdown-Listener",
            };
            listener.Start();
        }

        private const int DWMWA_TEXT_COLOR = 36;

        private void ApplyPurpleTitleBar()
        {
            if (Environment.OSVersion.Version.Major >= 10)
            {
                int r = ModernColors.Primary.R;
                int g = ModernColors.Primary.G;
                int b = ModernColors.Primary.B;
                int bgrColor = b << 16 | g << 8 | r;

                DwmSetWindowAttribute(this.Handle, DWMWA_CAPTION_COLOR, ref bgrColor, sizeof(int));
                DwmSetWindowAttribute(this.Handle, DWMWA_BORDER_COLOR, ref bgrColor, sizeof(int));

                // Fix caption text always showing as translucent when window hasn't been
                // the foreground owner for long. Explicitly pin it to white so DWM
                // never falls back to its inactive-state grey regardless of focus.
                int white = 0x00FFFFFF;
                DwmSetWindowAttribute(this.Handle, DWMWA_TEXT_COLOR, ref white, sizeof(int));
            }
        }

        /// <summary>
        /// 强制将窗口置于前台。
        /// 通过 AttachThreadInput 将本线程的输入队列临时挂接到当前前台线程，
        /// 绕过 Windows 对 SetForegroundWindow 的限制。
        /// </summary>
        private void ForceActivate()
        {
            var hwnd = this.Handle;
            ShowWindow(hwnd, SW_RESTORE);

            IntPtr fgHwnd = GetForegroundWindow();
            uint fgTid = GetWindowThreadProcessId(fgHwnd, out _);
            uint myTid = GetCurrentThreadId();

            if (fgTid != myTid)
                AttachThreadInput(myTid, fgTid, true);

            SetForegroundWindow(hwnd);
            BringWindowToTop(hwnd);

            if (fgTid != myTid)
                AttachThreadInput(myTid, fgTid, false);
        }

        private void MainForm_Paint(object? sender, PaintEventArgs e)
        {
            if (FormBorderStyle == FormBorderStyle.Sizable)
            {
                using (var pen = new Pen(ModernColors.Primary, 1))
                {
                    e.Graphics.DrawRectangle(pen, 0, 0, Width - 1, Height - 1);
                }
            }
        }

        private void BuildModernUI()
        {
            Text = Loc.Get("WindowTitle");
            var saved = WindowSizeStore.Load();
            Size = saved is (int w, int h, _) && w >= 700 && h >= 500
                ? new Size(w, h)
                : new Size(950, 990);
            MinimumSize = new Size(700, 500);
            StartPosition = FormStartPosition.CenterScreen;
            BackColor = ModernColors.Background;
            Font = new Font(Loc.GetFontPrimary(), 9.5f);
            FormBorderStyle = FormBorderStyle.Sizable;

            var mainPanel = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(20, 16, 20, 16),
                BackColor = Color.Transparent
            };

            // Title Bar
            var titlePanel = new Panel
            {
                Dock = DockStyle.Top,
                Height = 60,
                BackColor = Color.Transparent
            };

            var titleLabel = new Label
            {
                Text = Loc.Get("TitleLabel"),
                Font = new Font(Loc.GetFontPrimary(), 18f, FontStyle.Bold),
                ForeColor = ModernColors.Primary,
                AutoSize = true,
                Location = new Point(0, 1)
            };


            titlePanel.Controls.Add(titleLabel);

            // Pin / Always-on-Top button
            _pinButton = new Label
            {
                Text = Loc.Get("PinEmoji"),
                AutoSize = false,
                Size = new Size(48, 50),
                TextAlign = ContentAlignment.MiddleCenter,
                BackColor = Color.Transparent,
                ForeColor = ModernColors.TextSecondary,
                Font = new Font("Segoe UI Emoji", 12f),
                Cursor = Cursors.Hand,
                Margin = new Padding(0),
                Padding = new Padding(0),
            };
            _pinButton.Click += (s, e) =>
            {
                this.TopMost = !this.TopMost;
                UpdatePinButtonState();
                WindowSizeStore.Save(this.Width, this.Height, this.TopMost);
            };
            _pinButton.MouseEnter += (s, e) =>
            {
                if (!this.TopMost)
                    _pinButton.ForeColor = ModernColors.Primary;
            };
            _pinButton.MouseLeave += (s, e) =>
            {
                if (!this.TopMost)
                    _pinButton.ForeColor = ModernColors.TextSecondary;
            };

            var pinToolTip = new ToolTip();
            pinToolTip.SetToolTip(_pinButton, Loc.Get("PinTooltip"));

            // 在 Layout 时正确定位按钮到 titlePanel 右侧
            titlePanel.Layout += (s, e) =>
            {
                _pinButton.Location = new Point(
                    titlePanel.Width - _pinButton.Width - 16,
                    (titlePanel.Height - _pinButton.Height) / 2);
            };

            titlePanel.Controls.Add(_pinButton);

            // Search Container
            var searchContainer = new SearchBoxContainer
            {
                Dock = DockStyle.Top,
                Height = 54,
                Margin = new Padding(0, 16, 0, 0),
            };

            _searchBox = new ModernTextBox
            {
                Dock = DockStyle.None,
                Font = new Font(Loc.GetFontPrimary(), 11f),
                PlaceholderText = Loc.Get("SearchPlaceholder"),
                BorderStyle = BorderStyle.None,
                BackColor = Color.White,
            };
            _searchBox.KeyDown += (s, e) =>
            {
                if (e.KeyCode == Keys.Enter) StartSearch();
            };
            _searchBox.GotFocus += (s, e) => searchContainer.SetFocused(true);
            _searchBox.LostFocus += (s, e) => searchContainer.SetFocused(false);

            var separator = new Panel
            {
                Dock = DockStyle.Right,
                Width = 1,
                BackColor = ModernColors.Divider,
            };

            _searchButton = new ModernButton
            {
                Text = Loc.Get("SearchButton"),
                Dock = DockStyle.Right,
                Width = 110,
                CornerRadius = 3,
                RoundLeftCorners = false,
                RoundRightCorners = true,
                BackColor = ModernColors.Primary,
                ForeColor = Color.White,
                Font = new Font(Loc.GetFontPrimary(), 10.5f, FontStyle.Bold),
            };
            _searchButton.Click += (s, e) => StartSearch();

            searchContainer.Controls.Add(_searchBox);
            searchContainer.Controls.Add(separator);
            searchContainer.Controls.Add(_searchButton);
            searchContainer.SetSearchBox(_searchBox);

            // Options Panel
            var optionsPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = 60,
                Padding = new Padding(4, 1, 0, 4),
                BackColor = Color.Transparent,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
            };

            _currentNotebookOnly = new CheckBox
            {
                Text = Loc.Get("CurrentNotebookOnly"),
                AutoSize = true,
                Font = new Font(Loc.GetFontPrimary(), 9f),
                ForeColor = ModernColors.TextSecondary,
                Checked = false,
            };

            optionsPanel.Controls.Add(_currentNotebookOnly);

            // Results Card
            var resultsCard = new ModernCard
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(0),
                Margin = new Padding(0, 16, 0, 0),
            };

            _resultList = new ListBox
            {
                Dock = DockStyle.Fill,
                IntegralHeight = false,
                ItemHeight = 88,
                Font = new Font(Loc.GetFontPrimary(), 9.5f),
                DrawMode = DrawMode.OwnerDrawFixed,
                BorderStyle = BorderStyle.None,
                BackColor = ModernColors.CardBackground,
            };
            _resultList.DrawItem += ResultList_DrawItem;
            _resultList.DoubleClick += ResultList_DoubleClick;
            _resultList.KeyDown += (s, e) =>
            {
                if (e.KeyCode == Keys.Enter) NavigateToSelected();
            };

            resultsCard.Controls.Add(_resultList);

            // Status Bar
            var statusPanel = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 40,
                BackColor = ModernColors.Background,
                Padding = new Padding(10, 8, 0, 1),
            };
            statusPanel.Paint += (s, e) =>
            {
                using (var pen = new Pen(ModernColors.Divider, 1))
                {
                    e.Graphics.DrawLine(pen, 0, 0, statusPanel.Width, 0);
                }
            };

            _progress = new ProgressBar
            {
                Dock = DockStyle.Right,
                Width = 150,
                Height = 20,
                Style = ProgressBarStyle.Marquee,
                Visible = false,
                MarqueeAnimationSpeed = 30,
                Margin = new Padding(10, 0, 0, 0),
            };

            _statusLabel = new Label
            {
                Dock = DockStyle.Fill,
                Text = Loc.Get("StatusReady"),
                AutoSize = false,
                TextAlign = ContentAlignment.MiddleLeft,
                ForeColor = ModernColors.TextSecondary,
                Font = new Font(Loc.GetFontPrimary(), 8.5f),
                UseCompatibleTextRendering = true,
            };

            statusPanel.Controls.Add(_progress);
            statusPanel.Controls.Add(_statusLabel);

            // Assemble UI
            mainPanel.Controls.Add(resultsCard);
            mainPanel.Controls.Add(statusPanel);
            mainPanel.Controls.Add(optionsPanel);
            mainPanel.Controls.Add(searchContainer);
            mainPanel.Controls.Add(titlePanel);

            Controls.Add(mainPanel);
        }

        // Search Logic
        private void StartSearch()
        {
            string query = _searchBox.Text.Trim();
            if (string.IsNullOrEmpty(query))
            {
                LoadRecentPages();
                return;
            }

            _cts?.Cancel();
            _cts?.Dispose();
            _cts = new CancellationTokenSource();
            var token = _cts.Token;
            int searchVersion = Interlocked.Increment(ref _searchVersion);
            bool currentNotebookOnly = _currentNotebookOnly.Checked;

            _resultList.Items.Clear();
            _currentResults.Clear();
            _progress.Visible = true;
            SetStatus(Loc.Get("StatusSearching"));

            Task.Run(async () =>
            {
                try
                {
                    var results = await _scheduler.Run(svc => svc.Search(query,
                        currentNotebookOnly: currentNotebookOnly,
                        fastSearch: true,
                        progress: msg =>
                        {
                            if (!token.IsCancellationRequested)
                                BeginInvoke(() => SetStatus(msg));
                        }, token));

                    if (token.IsCancellationRequested || searchVersion != _searchVersion) return;

                    BeginInvoke(() =>
                    {
                        if (token.IsCancellationRequested || searchVersion != _searchVersion) return;
                        ShowResults(results, query);
                    });
                }
                catch (OperationCanceledException)
                {
                    if (searchVersion != _searchVersion) return;

                    BeginInvoke(() =>
                    {
                        if (searchVersion != _searchVersion) return;
                        SetStatus(Loc.Get("StatusSearchCancelled"));
                        _progress.Visible = false;
                    });
                }
                catch (Exception ex)
                {
                    if (!token.IsCancellationRequested && searchVersion == _searchVersion)
                        BeginInvoke(() =>
                        {
                            if (searchVersion != _searchVersion) return;
                            string msg = ex is System.Runtime.InteropServices.COMException || ex is InvalidOperationException
                                ? Loc.Get("ErrorCannotConnect")
                                : Loc.Fmt("ErrorFormat", ex.Message);
                            SetStatus(msg);
                            _progress.Visible = false;
                        });
                }
            }, token);
        }

        private void LoadRecentPages()
        {
            _cts?.Cancel();
            _cts?.Dispose();
            _cts = new CancellationTokenSource();
            var token = _cts.Token;
            int version = Interlocked.Increment(ref _searchVersion);
            bool currentNotebookOnly = _currentNotebookOnly.Checked;

            _resultList.Items.Clear();
            _currentResults.Clear();
            _progress.Visible = true;
            SetStatus(Loc.Get("StatusLoadingRecent"));

            Task.Run(async () =>
            {
                try
                {
                    // 阶段 1：快速获取元数据（单次 COM 调用）
                    var results = await _scheduler.Run(svc => svc.GetRecentPages(
                        maxCount: 10,
                        currentNotebookOnly: currentNotebookOnly));

                    if (token.IsCancellationRequested || version != _searchVersion) return;

                    BeginInvoke(() =>
                    {
                        if (token.IsCancellationRequested || version != _searchVersion) return;
                        ShowRecentResults(results);
                        SetStatus(Loc.Fmt("StatusRecentLoading", results.Count));
                    });

                    // 阶段 2：逐页获取内容预览（渐进式）
                    for (int i = 0; i < results.Count; i++)
                    {
                        if (token.IsCancellationRequested || version != _searchVersion) return;

                        string pageId = results[i].PageId;
                        try
                        {
                            string? preview = await _scheduler.Run(svc =>
                                svc.ExtractPagePreview(pageId));

                            if (token.IsCancellationRequested || version != _searchVersion) return;

                            if (!string.IsNullOrEmpty(preview))
                            {
                                BeginInvoke(() =>
                                {
                                    if (token.IsCancellationRequested || version != _searchVersion) return;
                                    UpdatePagePreview(pageId, preview);
                                });
                            }
                        }
                        catch
                        {
                            // 单个页面预览获取失败则跳过
                        }
                    }

                    // 全部预览加载完成
                    if (!token.IsCancellationRequested && version == _searchVersion)
                    {
                        BeginInvoke(() =>
                        {
                            if (token.IsCancellationRequested || version != _searchVersion) return;
                            SetStatus(Loc.Fmt("StatusRecentDone", results.Count));
                        });
                    }
                }
                catch (OperationCanceledException)
                {
                    if (version != _searchVersion) return;

                    BeginInvoke(() =>
                    {
                        if (version != _searchVersion) return;
                        SetStatus(Loc.Get("StatusCancelled"));
                        _progress.Visible = false;
                    });
                }
                catch (Exception ex)
                {
                    if (!token.IsCancellationRequested && version == _searchVersion)
                        BeginInvoke(() =>
                        {
                            if (version != _searchVersion) return;
                            string msg = ex is System.Runtime.InteropServices.COMException || ex is InvalidOperationException
                                ? Loc.Get("ErrorCannotConnect")
                                : Loc.Fmt("ErrorFormat", ex.Message);
                            SetStatus(msg);
                            _progress.Visible = false;
                        });
                }
            }, token);
        }

        private void UpdatePagePreview(string pageId, string preview)
        {
            for (int i = 0; i < _currentResults.Count; i++)
            {
                if (_currentResults[i].PageId == pageId)
                {
                    _currentResults[i].Snippet = preview;
                    _previewDirty = true;
                    return;
                }
            }
        }

        private void ShowRecentResults(List<PageResult> results)
        {
            _currentResults.Clear();
            _resultList.Items.Clear();

            foreach (var pageResult in results)
            {
                var matchResult = new MatchResult
                {
                    NotebookName     = pageResult.NotebookName,
                    SectionName      = pageResult.SectionName,
                    PageName         = pageResult.PageName,
                    PageId           = pageResult.PageId,
                    Snippet          = string.Empty,
                    ObjectId         = null,
                    MatchIndex       = 0,
                    TotalMatches     = 0,
                    LastModifiedTime = pageResult.LastModifiedTime,
                };

                _currentResults.Add(matchResult);
                _resultList.Items.Add(matchResult);
            }

            SetStatus(results.Count == 0
                ? Loc.Get("StatusNoPagesFound")
                : Loc.Fmt("StatusRecentDone", results.Count));

            _progress.Visible = false;
        }

        private void ShowResults(List<PageResult> results, string query)
        {
            _currentResults.Clear();
            _resultList.Items.Clear();

            int totalMatches = 0;
            foreach (var pageResult in results)
            {
                int matchCount = pageResult.Snippets.Count;
                for (int i = 0; i < matchCount; i++)
                {
                    var matchResult = new MatchResult
                    {
                        NotebookName = pageResult.NotebookName,
                        SectionName = pageResult.SectionName,
                        PageName = pageResult.PageName,
                        PageId = pageResult.PageId,
                        Snippet = pageResult.Snippets[i],
                        ObjectId = i < pageResult.HitObjectIds.Count
                            ? pageResult.HitObjectIds[i]
                            : null,
                        MatchIndex = i + 1,
                        TotalMatches = matchCount,
                        LastModifiedTime = pageResult.LastModifiedTime,
                    };

                    _currentResults.Add(matchResult);
                    _resultList.Items.Add(matchResult);
                    totalMatches++;
                }
            }

            SetStatus(results.Count == 0
                ? Loc.Fmt("StatusNoResultsForQuery", query)
                : Loc.Fmt("StatusResultsFound", results.Count, totalMatches));

            _progress.Visible = false;
        }

        // Custom Drawing
        private void ResultList_DrawItem(object? sender, DrawItemEventArgs e)
        {
            if (e.Index < 0 || e.Index >= _currentResults.Count) return;

            var match = _currentResults[e.Index];
            bool isSelected = (e.State & DrawItemState.Selected) == DrawItemState.Selected;

            Color bgColor = isSelected
                ? ModernColors.SelectionBg
                : (e.Index % 2 == 0 ? ModernColors.CardBackground : Color.FromArgb(252, 252, 252));

            using (var bgBrush = new SolidBrush(bgColor))
            {
                e.Graphics.FillRectangle(bgBrush, e.Bounds);
            }

            if (isSelected)
            {
                using (var accentBrush = new SolidBrush(ModernColors.Primary))
                {
                    e.Graphics.FillRectangle(accentBrush,
                        new Rectangle(e.Bounds.Left, e.Bounds.Top, 5, e.Bounds.Height));
                }
            }

            using var pageNameBrush = new SolidBrush(ModernColors.TextPrimary);
            using var pathBrush = new SolidBrush(ModernColors.TextHint);
            using var snippetBrush = new SolidBrush(ModernColors.TextSecondary);
            using var highlightBrush = new SolidBrush(ModernColors.Highlight);
            using var matchInfoBrush = new SolidBrush(ModernColors.Primary);
            using var iconBrush = new SolidBrush(ModernColors.TextHint);

            var pageNameFont = new Font(Loc.GetFontPrimary(), 10.5f, FontStyle.Bold);
            var pathFont = new Font(Loc.GetFontPrimary(), 9f, FontStyle.Regular);
            var snippetFont = new Font(Loc.Get("FontConsole"), 9.5f, FontStyle.Regular);
            var matchInfoFont = new Font(Loc.GetFontPrimary(), 8.5f, FontStyle.Bold);
            var iconFont = new Font(Loc.Get("FontEmoji"), 12f);

            float leftMargin = e.Bounds.Left + (isSelected ? 16 : 12);
            float topMargin = e.Bounds.Top + 14;

            e.Graphics.DrawString(Loc.Get("PageIcon"), iconFont, iconBrush,
                new PointF(leftMargin, topMargin - 1));

            float contentX = leftMargin + 44;
            e.Graphics.DrawString(match.PageName, pageNameFont, pageNameBrush,
                new PointF(contentX, topMargin));

            var pageNameSize = e.Graphics.MeasureString(match.PageName, pageNameFont);

            string matchInfo = match.GetSecondaryInfo();
            float matchInfoX = contentX + pageNameSize.Width + 8;
            if (!string.IsNullOrEmpty(matchInfo))
            {
                e.Graphics.DrawString(matchInfo, matchInfoFont, matchInfoBrush,
                    new PointF(matchInfoX, topMargin + 2));
                matchInfoX += e.Graphics.MeasureString(matchInfo, matchInfoFont).Width + 8;
            }

            string path = $"{match.NotebookName} ▸ {match.SectionName}";

            // 路径始终在行 1 末尾
            e.Graphics.DrawString(path, pathFont, pathBrush,
                new PointF(matchInfoX, topMargin + 3));

            float snippetY = topMargin + 38;
            float snippetX = contentX;

            if (!string.IsNullOrEmpty(match.Snippet))
            {
                // 有片段内容：搜索匹配（高亮）或最近页面的预览文本
                DrawHighlightedSnippet(e.Graphics, match.Snippet, snippetFont,
                    snippetBrush, highlightBrush, snippetX, snippetY, e.Bounds.Width - (int)snippetX - 12);
            }
            else if (match.LastModifiedTime != DateTime.MinValue)
            {
                // 最近页面预览尚未加载
                string placeholder = Loc.Get("LoadingPreview");
                using var placeholderBrush = new SolidBrush(ModernColors.TextHint);
                var placeholderFont = new Font(Loc.GetFontPrimary(), 9f, FontStyle.Italic);
                e.Graphics.DrawString(placeholder, placeholderFont, placeholderBrush,
                    new PointF(snippetX, snippetY));
            }

            if (!isSelected)
            {
                using var separatorPen = new Pen(ModernColors.Divider);
                e.Graphics.DrawLine(separatorPen,
                    e.Bounds.Left + 12, e.Bounds.Bottom - 1,
                    e.Bounds.Right - 12, e.Bounds.Bottom - 1);
            }
        }

        private void DrawHighlightedSnippet(Graphics g, string snippet, Font font,
            Brush normalBrush, Brush highlightBrush, float x, float y, int maxWidth)
        {
            float currentX = x;
            int currentIndex = 0;

            while (currentIndex < snippet.Length)
            {
                int startBracket = snippet.IndexOf('[', currentIndex);
                if (startBracket == -1)
                {
                    string remaining = snippet.Substring(currentIndex);

                    if (g.MeasureString(remaining, font).Width + currentX - x > maxWidth)
                    {
                        while (remaining.Length > 0 &&
                               g.MeasureString(remaining + "...", font).Width + currentX - x > maxWidth)
                        {
                            remaining = remaining.Substring(0, remaining.Length - 1);
                        }
                        remaining += "...";
                    }

                    g.DrawString(remaining, font, normalBrush, new PointF(currentX, y));
                    break;
                }

                if (startBracket > currentIndex)
                {
                    string before = snippet.Substring(currentIndex, startBracket - currentIndex);
                    g.DrawString(before, font, normalBrush, new PointF(currentX, y));
                    currentX += g.MeasureString(before, font).Width;
                }

                int endBracket = snippet.IndexOf(']', startBracket);
                if (endBracket == -1) break;

                string highlighted = snippet.Substring(startBracket + 1, endBracket - startBracket - 1);

                var highlightSize = g.MeasureString(highlighted, font);
                using (var highlightBg = new SolidBrush(ModernColors.HighlightBg))
                {
                    g.FillRectangle(highlightBg, currentX - 2, y, highlightSize.Width + 4, highlightSize.Height);
                }

                g.DrawString(highlighted, font, highlightBrush, new PointF(currentX, y));
                currentX += highlightSize.Width;

                currentIndex = endBracket + 1;
            }
        }

        // Navigation
        private void ResultList_DoubleClick(object? sender, EventArgs e) =>
            NavigateToSelected();

        private void NavigateToSelected()
        {
            int idx = _resultList.SelectedIndex;
            if (idx < 0 || idx >= _currentResults.Count) return;

            var match = _currentResults[idx];
            _ = _scheduler.Run(svc => svc.NavigateToPage(match.PageId, match.ObjectId))
                .ContinueWith(t =>
                {
                    if (t.IsFaulted)
                    {
                        var ex = t.Exception!.InnerException ?? t.Exception;
                        string msg = ex is System.Runtime.InteropServices.COMException
                            ? Loc.Get("ErrorCannotConnect") + $"\n\n({ex.Message})"
                            : Loc.Fmt("ErrorCannotOpenPage", ex.Message);
                        BeginInvoke(() => MessageBox.Show(msg, "OneFinder",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning));
                    }
                }, TaskScheduler.Default);
        }

        private void SetStatus(string text) => _statusLabel.Text = text;

        private void UpdatePinButtonState()
        {
            if (_pinButton == null) return;

            if (this.TopMost)
            {
                _pinButton.ForeColor = ModernColors.Primary;
                _pinButton.Font = new Font("Segoe UI Emoji", 13f, FontStyle.Bold);
            }
            else
            {
                _pinButton.ForeColor = ModernColors.TextSecondary;
                _pinButton.Font = new Font("Segoe UI Emoji", 12f, FontStyle.Regular);
            }
        }
    }

    // Custom Controls

    public class ModernCard : Panel
    {
        public ModernCard()
        {
            BackColor = Color.White;
            Padding = new Padding(16);
            DoubleBuffered = true;
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);

            using (var shadowBrush = new SolidBrush(Color.FromArgb(12, 0, 0, 0)))
            {
                e.Graphics.FillRectangle(shadowBrush,
                    new Rectangle(2, 2, Width, Height));
            }

            using (var cardBrush = new SolidBrush(BackColor))
            {
                e.Graphics.FillRectangle(cardBrush,
                    new Rectangle(0, 0, Width - 2, Height - 2));
            }

            using (var borderPen = new Pen(Color.FromArgb(224, 224, 224), 1))
            {
                e.Graphics.DrawRectangle(borderPen,
                    new Rectangle(0, 0, Width - 3, Height - 3));
            }
        }
    }

    public class ModernTextBox : TextBox
    {
        public ModernTextBox()
        {
            BorderStyle = BorderStyle.None;
            Padding = new Padding(12, 0, 12, 0);
            Font = new Font(Loc.GetFontPrimary(), 11f);
        }
    }

    /// <summary>
    /// 搜索框容器
    /// </summary>
    public class SearchBoxContainer : Panel
    {
        private bool _isFocused = false;
        private TextBox? _searchBox;

        public SearchBoxContainer()
        {
            BackColor = Color.White;
            Padding = new Padding(12, 1, 1, 1);
            DoubleBuffered = true;
        }

        protected override void OnLayout(LayoutEventArgs levent)
        {
            base.OnLayout(levent);
            if (_searchBox != null)
            {
                int topOffset = (this.Height - _searchBox.Height) / 2;
                _searchBox.Top = topOffset;
            }
        }

        public void SetSearchBox(TextBox searchBox)
        {
            _searchBox = searchBox;
            _searchBox.Dock = DockStyle.None;
            _searchBox.Anchor = AnchorStyles.Left | AnchorStyles.Right;
            _searchBox.Left = this.Padding.Left;
            _searchBox.Width = this.Width - 110 - 1 - this.Padding.Left - this.Padding.Right;
            int topOffset = (this.Height - _searchBox.Height) / 2;
            _searchBox.Top = topOffset;

            this.Resize += (s, e) =>
            {
                _searchBox.Width = this.Width - 110 - 1 - this.Padding.Left - this.Padding.Right;
                _searchBox.Top = (this.Height - _searchBox.Height) / 2;
            };
        }

        public void SetFocused(bool focused)
        {
            _isFocused = focused;
            Invalidate();
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;

            var borderColor = _isFocused
                ? MainForm.ModernColors.Primary
                : MainForm.ModernColors.Divider;
            var borderWidth = _isFocused ? 2 : 1;

            using (var borderPen = new Pen(borderColor, borderWidth))
            {
                var rect = new Rectangle(
                    borderWidth / 2,
                    borderWidth / 2,
                    Width - borderWidth,
                    Height - borderWidth);

                int radius = 6;
                using (var path = GetRoundedRectPath(rect, radius))
                {
                    e.Graphics.DrawPath(borderPen, path);
                }
            }
        }

        private System.Drawing.Drawing2D.GraphicsPath GetRoundedRectPath(Rectangle rect, int radius)
        {
            var path = new System.Drawing.Drawing2D.GraphicsPath();
            int diameter = radius * 2;

            path.AddArc(rect.X, rect.Y, diameter, diameter, 180, 90);
            path.AddArc(rect.Right - diameter, rect.Y, diameter, diameter, 270, 90);
            path.AddArc(rect.Right - diameter, rect.Bottom - diameter, diameter, diameter, 0, 90);
            path.AddArc(rect.X, rect.Bottom - diameter, diameter, diameter, 90, 90);
            path.CloseFigure();

            return path;
        }
    }

    public class ModernButton : Button
    {
        private Color _hoverBackColor;
        private bool _isHovering = false;

        public int CornerRadius { get; set; }
        public bool RoundLeftCorners { get; set; } = true;
        public bool RoundRightCorners { get; set; } = true;

        public ModernButton()
        {
            FlatStyle = FlatStyle.Flat;
            FlatAppearance.BorderSize = 0;
            Cursor = Cursors.Hand;
            Font = new Font(Loc.GetFontPrimary(), 10f, FontStyle.Bold);
        }

        protected override void OnBackColorChanged(EventArgs e)
        {
            base.OnBackColorChanged(e);

            int r = Math.Min(255, (int)(BackColor.R * 1.15));
            int g = Math.Min(255, (int)(BackColor.G * 1.15));
            int b = Math.Min(255, (int)(BackColor.B * 1.15));
            _hoverBackColor = Color.FromArgb(BackColor.A, r, g, b);
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;

            Color currentDrawColor = _isHovering ? _hoverBackColor : BackColor;

            using (var bgBrush = new SolidBrush(currentDrawColor))
            {
                var rect = new Rectangle(0, 0, Width, Height);
                if (CornerRadius > 0 && (RoundLeftCorners || RoundRightCorners))
                {
                    using var path = GetButtonPath(rect, CornerRadius, RoundLeftCorners, RoundRightCorners);
                    e.Graphics.FillPath(bgBrush, path);
                }
                else
                {
                    e.Graphics.FillRectangle(bgBrush, rect);
                }
            }

            var textSize = e.Graphics.MeasureString(Text, Font);
            var textX = (Width - textSize.Width) / 2;
            var textY = (Height - textSize.Height) / 2;

            using (var textBrush = new SolidBrush(ForeColor))
            {
                e.Graphics.DrawString(Text, Font, textBrush, new PointF(textX, textY));
            }
        }

        private GraphicsPath GetButtonPath(Rectangle rect, int radius, bool roundLeftCorners, bool roundRightCorners)
        {
            var path = new GraphicsPath();
            if (radius <= 0 || (!roundLeftCorners && !roundRightCorners))
            {
                path.AddRectangle(rect);
                return path;
            }

            int diameter = radius * 2;
            int leftInset = roundLeftCorners ? radius : 0;
            int rightInset = roundRightCorners ? radius : 0;

            path.StartFigure();
            path.AddLine(rect.Left + leftInset, rect.Top, rect.Right - rightInset, rect.Top);

            if (roundRightCorners)
                path.AddArc(rect.Right - diameter, rect.Top, diameter, diameter, 270, 90);

            path.AddLine(rect.Right, rect.Top + rightInset, rect.Right, rect.Bottom - rightInset);

            if (roundRightCorners)
                path.AddArc(rect.Right - diameter, rect.Bottom - diameter, diameter, diameter, 0, 90);

            path.AddLine(rect.Right - rightInset, rect.Bottom, rect.Left + leftInset, rect.Bottom);

            if (roundLeftCorners)
                path.AddArc(rect.Left, rect.Bottom - diameter, diameter, diameter, 90, 90);

            path.AddLine(rect.Left, rect.Bottom - leftInset, rect.Left, rect.Top + leftInset);

            if (roundLeftCorners)
                path.AddArc(rect.Left, rect.Top, diameter, diameter, 180, 90);

            path.CloseFigure();
            return path;
        }

        protected override void OnMouseEnter(EventArgs e)
        {
            base.OnMouseEnter(e);
            _isHovering = true;
            Invalidate();
        }

        protected override void OnMouseLeave(EventArgs e)
        {
            base.OnMouseLeave(e);
            _isHovering = false;
            Invalidate();
        }
    }

    /// <summary>
    /// 窗口尺寸持久化 — 保存/恢复到 %LocalAppData%\OneFinder\window.json
    /// </summary>
    internal static class WindowSizeStore
    {
        private static string FilePath =>
            Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "OneFinder", "window.json");

        public static (int Width, int Height, bool TopMost)? Load()
        {
            try
            {
                if (File.Exists(FilePath))
                {
                    var json = File.ReadAllText(FilePath);
                    using var doc = JsonDocument.Parse(json);
                    int w = doc.RootElement.GetProperty("Width").GetInt32();
                    int h = doc.RootElement.GetProperty("Height").GetInt32();
                    bool topMost = doc.RootElement.TryGetProperty("TopMost", out var tm)
                        && tm.GetBoolean();
                    return (w, h, topMost);
                }
            }
            catch { }
            return null;
        }

        public static void Save(int width, int height, bool topMost = false)
        {
            try
            {
                var dir = Path.GetDirectoryName(FilePath)!;
                if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);
                var topMostJson = topMost ? "true" : "false";
                File.WriteAllText(FilePath,
                    $"{{\"Width\":{width},\"Height\":{height},\"TopMost\":{topMostJson}}}");
            }
            catch { }
        }
    }
}
