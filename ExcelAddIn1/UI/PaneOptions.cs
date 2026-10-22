using System;
using System.Globalization;
using System.IO;

namespace ExcelAddIn1.UI
{
    /// <summary>面板选项，持久化到 %LOCALAPPDATA%\ExcelAddIn1\settings.ini。</summary>
    public static class PaneOptions
    {
        private static readonly object Sync = new object();
        private static bool _loaded;
        private static bool _showHiddenSheets = true;
        private static double _paneWidth = 380;
        private static bool _lastPaneVisible;
        private static string _filePath;

        public static bool ShowHiddenSheets
        {
            get { EnsureLoaded(); return _showHiddenSheets; }
            set { EnsureLoaded(); _showHiddenSheets = value; Save(); }
        }

        public static double PaneWidth
        {
            get { EnsureLoaded(); return _paneWidth; }
            set
            {
                EnsureLoaded();
                if (value < 220) value = 220;
                if (value > 700) value = 700;
                _paneWidth = value;
                Save();
            }
        }

        /// <summary>上次退出时面板是否开着，用来决定下次启动是否自动显示。</summary>
        public static bool LastPaneVisible
        {
            get { EnsureLoaded(); return _lastPaneVisible; }
            set
            {
                EnsureLoaded();
                if (_lastPaneVisible == value) return;
                _lastPaneVisible = value;
                Save();
            }
        }

        private static string FilePath
        {
            get
            {
                if (_filePath == null)
                {
                    var dir = Path.Combine(
                        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                        "ExcelAddIn1");
                    _filePath = Path.Combine(dir, "settings.ini");
                }
                return _filePath;
            }
        }

        private static void EnsureLoaded()
        {
            if (_loaded) return;
            lock (Sync)
            {
                if (_loaded) return;
                try
                {
                    if (File.Exists(FilePath))
                    {
                        foreach (var line in File.ReadAllLines(FilePath))
                        {
                            var idx = line.IndexOf('=');
                            if (idx <= 0) continue;
                            var key = line.Substring(0, idx).Trim();
                            var value = line.Substring(idx + 1).Trim();
                            if (string.Equals(key, "ShowHiddenSheets", StringComparison.OrdinalIgnoreCase))
                            {
                                bool parsed;
                                if (bool.TryParse(value, out parsed)) _showHiddenSheets = parsed;
                            }
                            else if (string.Equals(key, "LastPaneVisible", StringComparison.OrdinalIgnoreCase))
                            {
                                bool visible;
                                if (bool.TryParse(value, out visible)) _lastPaneVisible = visible;
                            }
                            else if (string.Equals(key, "PaneWidth", StringComparison.OrdinalIgnoreCase))
                            {
                                double parsed;
                                if (double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out parsed)
                                    && parsed >= 220 && parsed <= 700)
                                {
                                    _paneWidth = parsed;
                                }
                            }
                        }
                    }
                }
                catch
                {
                    // 选项文件损坏不影响使用，用默认值
                }
                _loaded = true;
            }
        }

        private static void Save()
        {
            try
            {
                var dir = Path.GetDirectoryName(FilePath);
                if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);
                File.WriteAllLines(FilePath, new[]
                {
                    "ShowHiddenSheets=" + _showHiddenSheets.ToString(),
                    "LastPaneVisible=" + _lastPaneVisible.ToString(),
                    "PaneWidth=" + _paneWidth.ToString(CultureInfo.InvariantCulture)
                });
            }
            catch
            {
                // 写不进去就算了
            }
        }
    }
}
