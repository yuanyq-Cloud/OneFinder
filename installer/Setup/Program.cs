using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Windows.Forms;

namespace OneFinder.Setup
{
    internal static class Program
    {
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            // Show language selection dialog.
            using var langDialog = new LanguageDialog();
            if (langDialog.ShowDialog() != DialogResult.OK)
                return;

            string language = langDialog.SelectedLanguage;

            // Write language preference to HKCU (no admin required).
            // The app and AddIn read HKCU first, so this takes effect
            // even when msiexec enters maintenance mode and skips HKLM.
            SaveLanguageToRegistry(language);

            // Extract the embedded MSI to a temp file.
            string? msiPath = null;
            try
            {
                msiPath = ExtractMsi();
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Failed to extract the installer.\n\n{ex.Message}",
                    "OneFinder Setup Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // Launch msiexec.
            try
            {
                var startInfo = new ProcessStartInfo
                {
                    FileName = "msiexec.exe",
                    Arguments = $"/i \"{msiPath}\"",
                    UseShellExecute = true,
                    Verb = "runas",
                };

                var process = Process.Start(startInfo);
                if (process != null)
                {
                    process.WaitForExit();
                    try { File.Delete(msiPath); } catch { }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Failed to start the installer.\n\n{ex.Message}",
                    "OneFinder Setup Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Write the language preference to HKCU so it survives MSI maintenance mode.
        /// </summary>
        private static void SaveLanguageToRegistry(string language)
        {
            try
            {
                using var key = Microsoft.Win32.Registry.CurrentUser
                    .CreateSubKey(@"Software\OneFinder");
                key?.SetValue("Language", language);
            }
            catch
            {
                // Non-fatal: the app will fall back to zh-CN.
            }
        }

        /// <summary>
        /// Extract the embedded MSI resource to a temporary file.
        /// </summary>
        private static string ExtractMsi()
        {
            var asm = Assembly.GetExecutingAssembly();
            var tempDir = Path.GetTempPath();
            var msiPath = Path.Combine(tempDir, "OneFinderSetup.msi");

            using (var stream = asm.GetManifestResourceStream("OneFinderSetup.msi"))
            {
                if (stream == null)
                    throw new InvalidOperationException(
                        "Embedded MSI not found in the setup executable.");

                using (var fileStream = File.Create(msiPath))
                {
                    stream.CopyTo(fileStream);
                }
            }

            return msiPath;
        }
    }
}
