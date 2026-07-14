using System;
using System.Globalization;
using System.Resources;
using System.Threading;

namespace OneFinder
{
    /// <summary>
    /// Localization helper — reads the display language from
    /// HKCU\Software\OneFinder\Language (set by the bootstrapper),
    /// falling back to HKLM, then zh-CN.
    /// </summary>
    public static class Loc
    {
        private static ResourceManager? _rm;
        private static CultureInfo _culture = new("zh-CN", useUserOverride: false);

        /// <summary>
        /// The current UI culture (e.g. "zh-CN" or "en-US").
        /// </summary>
        public static string CurrentLanguage { get; private set; } = "zh-CN";

        /// <summary>
        /// Initialize localization. Reads the language preference from the registry.
        /// Checks HKCU first (set by the bootstrapper, runs without admin),
        /// then HKLM (set by MSI as default). Falls back to zh-CN if neither is set.
        /// </summary>
        public static void Initialize()
        {
            string? lang = null;
            try
            {
                // HKCU first — bootstrapper writes here on every run, no admin needed.
                lang = Microsoft.Win32.Registry.CurrentUser
                    .OpenSubKey(@"Software\OneFinder")
                    ?.GetValue("Language") as string;
            }
            catch { }

            if (string.IsNullOrEmpty(lang))
            {
                try
                {
                    lang = Microsoft.Win32.Registry.LocalMachine
                        .OpenSubKey(@"Software\OneFinder")
                        ?.GetValue("Language") as string;
                }
                catch { }
            }

            Initialize(lang);
        }

        /// <summary>
        /// Initialize with an explicit language code.
        /// </summary>
        public static void Initialize(string? language)
        {
            CurrentLanguage = (!string.IsNullOrEmpty(language) && language == "en-US")
                ? "en-US"
                : "zh-CN";

            _culture = new CultureInfo(CurrentLanguage, useUserOverride: false);

            // Set for the current (main) thread
            Thread.CurrentThread.CurrentUICulture = _culture;
            Thread.CurrentThread.CurrentCulture = _culture;

            // Set defaults so background threads (Task.Run) inherit the correct culture
            CultureInfo.DefaultThreadCurrentUICulture = _culture;
            CultureInfo.DefaultThreadCurrentCulture = _culture;

            _rm = new ResourceManager("OneFinder.Strings", typeof(Loc).Assembly);

            System.Diagnostics.Debug.WriteLine(
                $"[Loc] Initialized: language={CurrentLanguage}, culture={_culture.Name}");
        }

        /// <summary>
        /// Get a localized string by key. Returns the key itself if the resource is missing.
        /// Uses the stored culture so background threads always get the correct language.
        /// </summary>
        public static string Get(string key)
        {
            if (_rm == null)
                Initialize();

            return _rm!.GetString(key, _culture) ?? key;
        }

        /// <summary>
        /// Get a localized format string by key and apply arguments.
        /// </summary>
        public static string Fmt(string key, params object[] args)
        {
            var format = Get(key);
            return args.Length > 0 ? string.Format(format, args) : format;
        }

        /// <summary>
        /// Returns the primary UI font name for the current language.
        /// Chinese → Microsoft YaHei; English → Segoe UI.
        /// </summary>
        public static string GetFontPrimary()
        {
            return CurrentLanguage == "en-US" ? "Segoe UI" : "Microsoft YaHei";
        }
    }
}
