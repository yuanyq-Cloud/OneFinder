// OneFinder.AddIn — OneNote COM 插件（.NET Framework 4.8）
// 架构参照 OneMore: https://github.com/stevencohn/OneMore/blob/main/OneMore/AddIn.cs
//
// 加载链：OneNote.exe → mscoree.dll → CLR v4 → 本 DLL
// 注册表键：HKCU\SOFTWARE\Microsoft\Office\OneNote\AddIns\OneFinder.AddIn

using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using Extensibility;
using Microsoft.Office.Core;

namespace OneFinder.AddIn
{
    /// <summary>
    /// OneNote COM 插件主类。
    /// CLSID / ProgID 必须与安装包写入注册表的值完全一致。
    /// </summary>
    [ComVisible(true)]
    [Guid(AddIn.AddinClsid)]
    [ProgId(AddIn.AddinProgId)]
    [ClassInterface(ClassInterfaceType.None)]
    public class AddIn : IDTExtensibility2, IRibbonExtensibility
    {
        // 修改这两个常量时，安装包的 WXS 也要同步更改
        public const string AddinClsid  = "6B29FC40-CA47-1067-B31D-00DD010662DA";
        public const string AddinProgId = "OneFinder.AddIn";

        private IRibbonUI? _ribbon;

        // ── IDTExtensibility2 ────────────────────────────────────────────────

        public void OnConnection(object Application, ext_ConnectMode ConnectMode,
            object AddInInst, ref Array custom) { }

        public void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom) { }

        public void OnAddInsUpdate(ref Array custom) { }

        public void OnStartupComplete(ref Array custom) { }

        public void OnBeginShutdown(ref Array custom) { }

        // ── IRibbonExtensibility ─────────────────────────────────────────────

        /// <summary>
        /// OneNote 启动时调用，返回 Ribbon 自定义 XML。
        /// </summary>
        public string GetCustomUI(string RibbonID)
        {
            var asm = Assembly.GetExecutingAssembly();
            // 嵌入资源名格式：<RootNamespace>.<文件名>
            using (var stream = asm.GetManifestResourceStream("OneFinder.AddIn.Ribbon.xml"))
            {
                if (stream == null) return string.Empty;
                using (var reader = new StreamReader(stream))
                    return reader.ReadToEnd();
            }
        }

        // ── Ribbon 回调 ──────────────────────────────────────────────────────

        public void RibbonLoaded(IRibbonUI ribbon)
        {
            _ribbon = ribbon;
        }

        /// <summary>
        /// 按下"全文搜索"按钮：启动或激活 OneFinder.exe。
        /// </summary>
        public void OnSearchClick(IRibbonControl control)
        {
            try
            {
                var dir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                var exe = Path.Combine(dir, "OneFinder.exe");

                if (!File.Exists(exe))
                {
                    System.Windows.Forms.MessageBox.Show(
                        string.Format("未找到 OneFinder.exe：\n{0}", exe),
                        "OneFinder",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Warning);
                    return;
                }

                // 如果已运行则前置窗口，否则启动
                var procName = Path.GetFileNameWithoutExtension(exe);
                var procs = Process.GetProcessesByName(procName);
                if (procs.Length > 0 && procs[0].MainWindowHandle != IntPtr.Zero)
                {
                    NativeMethods.ShowWindow(procs[0].MainWindowHandle, 9);  // SW_RESTORE
                    NativeMethods.SetForegroundWindow(procs[0].MainWindowHandle);
                }
                else
                {
                    var psi = new ProcessStartInfo
                    {
                        FileName = exe,
                        UseShellExecute = true,
                        WorkingDirectory = dir
                    };
                    Process.Start(psi);
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(
                    string.Format("启动 OneFinder 失败：\n{0}", ex.Message),
                    "OneFinder",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error);
            }
        }
    }

    internal static class NativeMethods
    {
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        internal static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
    }
}
