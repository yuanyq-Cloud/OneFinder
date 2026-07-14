// OneFinder.AddIn - OneNote COM addin (.NET Framework 4.8)
// Load chain: OneNote.exe -> mscoree.dll -> CLR v4 -> this DLL
// Registry: HKCU\SOFTWARE\Microsoft\Office\OneNote\AddIns\OneFinder.AddIn

using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using Extensibility;
using Microsoft.Office.Core;

namespace OneFinder.AddIn
{
    [ComVisible(true)]
    [Guid(AddinClsid)]
    [ProgId(AddinProgId)]
    public class AddIn : IDTExtensibility2, IRibbonExtensibility
    {
        public const string AddinClsid  = "6B29FC40-CA47-1067-B31D-00DD010662DA";
        public const string AddinProgId = "OneFinder.AddIn";

        private IRibbonUI? _ribbon;

        private static string GetInstallDir()
        {
            try
            {
                var codeBase = Assembly.GetExecutingAssembly().CodeBase;
                if (!string.IsNullOrEmpty(codeBase))
                {
                    var uri = new Uri(codeBase);
                    if (uri.IsFile)
                        return Path.GetDirectoryName(uri.LocalPath) ?? AppDomain.CurrentDomain.BaseDirectory;
                }
            }
            catch (Exception ex)
            {
                Log($"GetInstallDir (CodeBase): {ex.Message}");
            }

            try
            {
                var loc = Assembly.GetExecutingAssembly().Location;
                if (!string.IsNullOrEmpty(loc))
                    return Path.GetDirectoryName(loc) ?? AppDomain.CurrentDomain.BaseDirectory;
            }
            catch (Exception ex)
            {
                Log($"GetInstallDir (Location): {ex.Message}");
            }

            return AppDomain.CurrentDomain.BaseDirectory;
        }

        private static void Log(string msg)
        {
            try
            {
                var path = Path.Combine(Path.GetTempPath(), "OneFinder.AddIn.log");
                File.AppendAllText(path, $"[{DateTime.Now:HH:mm:ss}] {msg}{Environment.NewLine}");
            }
            catch (Exception ex)
            {
                // Silent fallback - can't log if logging fails
                System.Diagnostics.Debug.WriteLine($"Log write failed: {ex.Message}");
            }
        }

        public void OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
        {
            Log($"OnConnection ConnectMode={ConnectMode}");
        }

        public void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        {
            Log($"OnDisconnection RemoveMode={RemoveMode}");
            try
            {
                if (_ribbon != null)
                {
                    Log("OnDisconnection: releasing IRibbonUI COM reference");
                    Marshal.ReleaseComObject(_ribbon);
                    _ribbon = null;
                    Log("OnDisconnection: IRibbonUI released");
                }
            }
            catch (Exception ex)
            {
                Log($"OnDisconnection: release ribbon error: {ex.Message}");
            }

            // Force GC to collect any other lingering COM RCWs (e.g. Application object
            // passed to OnConnection) so ONENOTE.EXE reference count can reach zero.
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            Log("OnDisconnection: GC collect done");
        }

        public void OnAddInsUpdate(ref Array custom) { }

        public void OnStartupComplete(ref Array custom)
        {
            Log("OnStartupComplete");
        }

        public void OnBeginShutdown(ref Array custom)
        {
            Log("OnBeginShutdown: start");

            // Step 1: Signal OneFinder to release COM references immediately.
            // This must happen before OneNote starts tearing down its COM infrastructure.
            SignalEvent("Local\\OneFinder-ReleaseCOM", "ReleaseCOM");

            // Small delay to allow OneFinder's STA thread to process the ReleaseCOM action
            // before OneNote proceeds with shutdown and invalidates COM objects.
            System.Threading.Thread.Sleep(300);
            Log("OnBeginShutdown: post-ReleaseCOM delay elapsed");

            // Step 2: Signal OneFinder to close its window.
            SignalEvent("Local\\OneFinder-OneNoteShutdown", "OneNoteShutdown");

            Log("OnBeginShutdown: done");
        }

        private static void SignalEvent(string name, string label)
        {
            try
            {
                if (EventWaitHandle.TryOpenExisting(name, out var handle))
                {
                    using (handle)
                    {
                        handle.Set();
                        Log($"SignalEvent: '{label}' signaled OK");
                    }
                }
                else
                {
                    Log($"SignalEvent: '{label}' handle not found (OneFinder not running?)");
                }
            }
            catch (Exception ex)
            {
                Log($"SignalEvent '{label}': {ex.Message}");
            }
        }

        public string GetCustomUI(string RibbonID)
        {
            Log($"GetCustomUI RibbonID={RibbonID}");
            try
            {
                // Read the language preference from the registry.
                // Checks HKCU first (set by bootstrapper), then HKLM (MSI default).
                string language = "zh-CN";
                try
                {
                    var langValue = Microsoft.Win32.Registry.CurrentUser
                        .OpenSubKey(@"Software\OneFinder")
                        ?.GetValue("Language") as string;
                    if (string.IsNullOrEmpty(langValue))
                    {
                        langValue = Microsoft.Win32.Registry.LocalMachine
                            .OpenSubKey(@"Software\OneFinder")
                            ?.GetValue("Language") as string;
                    }
                    if (!string.IsNullOrEmpty(langValue))
                        language = langValue;
                }
                catch (Exception ex)
                {
                    Log($"GetCustomUI: registry read error: {ex.Message}");
                }

                // Select the appropriate embedded Ribbon XML resource.
                // Try the language-specific resource first; fall back to Ribbon.xml
                // if it's missing (e.g. addin DLL wasn't rebuilt with the new resource).
                var asm = Assembly.GetExecutingAssembly();

                string primaryResource = language == "en-US"
                    ? "OneFinder.AddIn.Ribbon_en.xml"
                    : "OneFinder.AddIn.Ribbon.xml";
                string fallbackResource = "OneFinder.AddIn.Ribbon.xml";

                Log($"GetCustomUI: language={language}, primary={primaryResource}");

                // Try primary (language-specific) resource
                var stream = asm.GetManifestResourceStream(primaryResource);

                // If not found, try the fallback (default Chinese)
                if (stream == null && primaryResource != fallbackResource)
                {
                    Log($"GetCustomUI: {primaryResource} not found, falling back to {fallbackResource}");
                    stream = asm.GetManifestResourceStream(fallbackResource);
                }

                if (stream == null)
                {
                    Log($"ERROR: no Ribbon resource found (tried {primaryResource}, {fallbackResource})");
                    return string.Empty;
                }

                using (stream)
                using (var reader = new StreamReader(stream))
                {
                    var xml = reader.ReadToEnd();
                    Log($"GetCustomUI OK, {xml.Length} chars");
                    return xml;
                }
            }
            catch (Exception ex)
            {
                Log($"GetCustomUI: {ex.Message}");
                return string.Empty;
            }
        }

        public void RibbonLoaded(IRibbonUI ribbon)
        {
            _ribbon = ribbon;
            Log("RibbonLoaded OK");
        }

        public void OnSearchClick(IRibbonControl control)
        {
            Log("OnSearchClick fired");
            try
            {
                var dir = GetInstallDir();
                var exe = Path.Combine(dir, "OneFinder.exe");
                Log($"exe path={exe} exists={File.Exists(exe)}");

                if (!File.Exists(exe))
                {
                    Log($"ERROR: exe not found at {exe}");
                    return;
                }

                var procName = Path.GetFileNameWithoutExtension(exe);
                var procs = Process.GetProcessesByName(procName);
                var hwnd = procs.Length > 0 ? procs[0].MainWindowHandle : IntPtr.Zero;
                
                if (hwnd != IntPtr.Zero)
                {
                    Log("already running - signaling activate");
                    SignalEvent("Local\\OneFinder-Activate", "Activate");
                }
                else
                {
                    // Kill stale processes before starting new one
                    foreach (var stale in procs)
                    {
                        try
                        {
                            Log($"killing stale pid={stale.Id}");
                            stale.Kill();
                            stale.WaitForExit(1000); // Wait up to 1s for process to exit
                        }
                        catch (Exception ex)
                        {
                            Log($"kill stale process: {ex.Message}");
                        }
                        finally
                        {
                            stale.Dispose();
                        }
                    }

                    Log($"starting {exe}");
                    using (var p = Process.Start(new ProcessStartInfo
                    {
                        FileName = exe,
                        UseShellExecute = true,
                        WorkingDirectory = dir
                    }))
                    {
                        Log($"started pid={p?.Id ?? -1}");
                        if (p != null)
                        {
                            // Grant the new process permission to call SetForegroundWindow.
                            // Without this, Windows blocks foreground activation from a
                            // process that wasn't started by the current foreground owner.
                            NativeMethods.AllowSetForegroundWindow(p.Id);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Log($"OnSearchClick: {ex.Message}\n{ex.StackTrace}");
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

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool AllowSetForegroundWindow(int dwProcessId);
    }
}