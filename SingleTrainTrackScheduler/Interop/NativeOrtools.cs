using System;
using System.ComponentModel;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;

namespace SingleTrainTrackScheduler.Interop
{
    internal static class NativeOrtools
    {
        [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        static extern IntPtr LoadLibrary(string lpFileName);

        [DllImport("kernel32.dll", SetLastError = true)]
        static extern bool SetDefaultDllDirectories(uint DirectoryFlags);

        [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        static extern IntPtr AddDllDirectory(string NewDirectory);

        [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        static extern bool SetDllDirectory(string lpPathName);

        const uint LOAD_LIBRARY_SEARCH_DEFAULT_DIRS = 0x00001000;

        public static void EnsureLoaded()
        {
            string baseDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)!;
            string runtimeDir = Path.Combine(baseDir, "runtimes", "win-x64", "native");

            // Önce arama yolunu güvenli şekilde genişlet
            try
            {
                SetDefaultDllDirectories(LOAD_LIBRARY_SEARCH_DEFAULT_DIRS);
                AddDllDirectory(baseDir);
                if (Directory.Exists(runtimeDir)) AddDllDirectory(runtimeDir);
            }
            catch
            {
                // Eski Windows ise
                SetDllDirectory(baseDir);
                if (Directory.Exists(runtimeDir)) SetDllDirectory(runtimeDir);
            }

            // Doğru yer önce (runtimes), sonra kök denenir
            string[] candidates =
            {
                Path.Combine(runtimeDir, "google-ortools-native.dll"),
                Path.Combine(baseDir,   "google-ortools-native.dll")
            };

            var report = new StringBuilder();
            foreach (var path in candidates)
            {
                if (!File.Exists(path))
                {
                    report.AppendLine($"- {path} (yok)");
                    continue;
                }

                var h = LoadLibrary(path);
                if (h != IntPtr.Zero)
                    return; // yüklendi

                int err = Marshal.GetLastWin32Error();
                report.AppendLine($"- {path} (Win32Error={err})");
            }

            throw new InvalidOperationException(
                "OR-Tools yerel DLL yüklenemedi.\n" +
                $"Process x64: {Environment.Is64BitProcess}, OS x64: {Environment.Is64BitOperatingSystem}\n" +
                "Denenen yollar:\n" + report.ToString() +
                "\nÇözüm: DLL yalnızca 'runtimes\\win-x64\\native' altında bulunmalı ve " +
                "hedef makinede 'Microsoft Visual C++ 2015–2022 x64 Redistributable' kurulu olmalı.");
        }
    }
}
