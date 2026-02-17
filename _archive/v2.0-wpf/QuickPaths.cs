// QuickPaths.cs - Desktop floating path quick-copy tool
// Build: build.cmd (csc.exe from .NET Framework 4.x)

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Text;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Effects;
using System.Windows.Threading;
using Microsoft.Win32;
using Polyline = System.Windows.Shapes.Polyline;
using Ellipse = System.Windows.Shapes.Ellipse;

// --- JSON data contracts (keys match legacy paths.json / config.json) ---

[DataContract]
class PathItem
{
    [DataMember(Name = "name")] public string Name { get; set; }
    [DataMember(Name = "path")] public string ItemPath { get; set; }
}

[DataContract]
class AppConfig
{
    [DataMember(Name = "left")] public int Left { get; set; }
    [DataMember(Name = "top")] public int Top { get; set; }
    [DataMember(Name = "claudeMode")] public bool ClaudeMode { get; set; }
    [DataMember(Name = "scale")] public double Scale { get; set; }
}

// --- COM IFileOpenDialog multi-folder picker (Ctrl+click multi-select) ---

class FolderMultiPicker
{
    [ComImport, Guid("DC1C5A9C-E88A-4DDE-A5A1-60F82A20AEF7")]
    class FileOpenDialogCOM { }

    [ComImport, Guid("D57C7288-D4AD-4768-BE02-9D969532D960")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IFileOpenDialog
    {
        [PreserveSig] int Show(IntPtr hwndOwner);
        void SetFileTypes(uint cFileTypes, IntPtr rgFilterSpec);
        void SetFileTypeIndex(uint iFileType);
        void GetFileTypeIndex(out uint piFileType);
        void Advise(IntPtr pfde, out uint pdwCookie);
        void Unadvise(uint dwCookie);
        void SetOptions(uint fos);
        void GetOptions(out uint pfos);
        void SetDefaultFolder(IShellItem psi);
        void SetFolder(IShellItem psi);
        void GetFolder(out IShellItem ppsi);
        void GetCurrentSelection(out IShellItem ppsi);
        void SetFileName([MarshalAs(UnmanagedType.LPWStr)] string pszName);
        void GetFileName([MarshalAs(UnmanagedType.LPWStr)] out string pszName);
        void SetTitle([MarshalAs(UnmanagedType.LPWStr)] string pszTitle);
        void SetOkButtonLabel([MarshalAs(UnmanagedType.LPWStr)] string pszText);
        void SetFileNameLabel([MarshalAs(UnmanagedType.LPWStr)] string pszLabel);
        void GetResult(out IShellItem ppsi);
        void AddPlace(IShellItem psi, int fdap);
        void SetDefaultExtension([MarshalAs(UnmanagedType.LPWStr)] string pszDefaultExtension);
        void Close(int hr);
        void SetClientGuid(ref Guid guid);
        void ClearClientData();
        void SetFilter(IntPtr pFilter);
        void GetResults(out IShellItemArray ppenum);
        void GetSelectedItems(out IntPtr ppsai);
    }

    [ComImport, Guid("43826D1E-E718-42EE-BC55-A1E261C37BFE")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IShellItem
    {
        void BindToHandler(IntPtr pbc, ref Guid bhid, ref Guid riid, out IntPtr ppv);
        void GetParent(out IShellItem ppsi);
        void GetDisplayName(uint sigdnName, [MarshalAs(UnmanagedType.LPWStr)] out string ppszName);
        void GetAttributes(uint sfgaoMask, out uint psfgaoAttribs);
        void Compare(IShellItem psi, uint hint, out int piOrder);
    }

    [ComImport, Guid("B63EA76D-1F85-456F-A19C-48159EFA858B")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IShellItemArray
    {
        void BindToHandler(IntPtr pbc, ref Guid bhid, ref Guid riid, out IntPtr ppvOut);
        void GetPropertyStore(int flags, ref Guid riid, out IntPtr ppv);
        void GetPropertyDescriptionList(IntPtr keyType, ref Guid riid, out IntPtr ppv);
        void GetAttributes(int attribFlags, uint sfgaoMask, out uint psfgaoAttribs);
        void GetCount(out uint pdwNumItems);
        void GetItemAt(uint dwIndex, out IShellItem ppsi);
        void EnumItems(out IntPtr ppenumShellItems);
    }

    public static string[] ShowDialog(string title)
    {
        IFileOpenDialog dialog = (IFileOpenDialog)new FileOpenDialogCOM();
        try
        {
            uint options;
            dialog.GetOptions(out options);
            dialog.SetOptions(options | 0x20u | 0x200u | 0x40u); // PICKFOLDERS | ALLOWMULTISELECT | FORCEFILESYSTEM
            if (!string.IsNullOrEmpty(title))
                dialog.SetTitle(title);
            int hr = dialog.Show(IntPtr.Zero);
            if (hr != 0) return new string[0];
            IShellItemArray results;
            dialog.GetResults(out results);
            uint count;
            results.GetCount(out count);
            var paths = new List<string>();
            for (uint i = 0; i < count; i++)
            {
                IShellItem item;
                results.GetItemAt(i, out item);
                string path;
                item.GetDisplayName(0x80058000u, out path); // SIGDN_FILESYSPATH
                if (!string.IsNullOrEmpty(path))
                    paths.Add(path);
            }
            return paths.ToArray();
        }
        finally
        {
            Marshal.ReleaseComObject(dialog);
        }
    }
}

// --- Main application ---

class QuickPaths
{
    // P/Invoke for cursor position (replaces System.Windows.Forms dependency)
    [DllImport("user32.dll")]
    static extern bool GetCursorPos(out POINT lpPoint);
    [StructLayout(LayoutKind.Sequential)]
    struct POINT { public int X, Y; }

    // --- Constants ---
    const int ECG_PAD_X = 12, ECG_PAD_Y = 8;
    const int ECG_WAVE_W = 60, ECG_WAVE_H = 22;
    const int ECG_COUNT = 360;
    const string MUTEX_NAME = @"Global\QuickPaths_Singleton";
    const string REG_RUN = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Run";

    // --- State ---
    string dir, dataFile, configFile, logFile;
    List<PathItem> paths = new List<PathItem>();
    bool dialogOpen, expanded, claudeMode, intentionalExit;
    POINT? dotClickPt;
    int exitCode = 1;
    int restartCount;
    DateTime startTime;
    Mutex mutex;

    // --- UI elements ---
    Window win;
    Grid root;
    Border dot, panel;
    Canvas canvas;
    Polyline waveLine;
    Ellipse penHead, penGlow1, penGlow2;
    DropShadowEffect penHeadShadow;
    StackPanel listPanel;
    ScaleTransform dotScale;
    List<double> yValues = new List<double>();

    // --- Timers ---
    DispatcherTimer breathTimer, flashTimer, displayChangeTimer, healthTimer;
    int breathTick;

    // --- System event handlers (stored for unsubscribe) ---
    PowerModeChangedEventHandler onPowerMode;
    SessionSwitchEventHandler onSessionSwitch;
    EventHandler onDisplayChanged;

    // --- Brushes ---
    // ECG monitor
    SolidColorBrush C_MonBg, C_MonHover, C_MonBdrBlue, C_MonBdrOrange;
    SolidColorBrush C_WaveBlue, C_WaveOrange, C_WaveGreen;
    SolidColorBrush C_PenBlue, C_PenOrange, C_PenGreen;
    // Panel
    SolidColorBrush C_PanelBg, C_ItemNormal, C_ItemHover;
    SolidColorBrush C_UpNormal, C_UpHover, C_UpText;
    SolidColorBrush C_DelNormal, C_DelHover, C_DelText;
    SolidColorBrush C_AddBg, C_AddHover, C_AddText;
    SolidColorBrush C_Sep;
    // Claude toggle
    SolidColorBrush C_ClaudeOff, C_ClaudeOn, C_ClaudeHoverOff, C_ClaudeHoverOn;
    SolidColorBrush C_ClaudeTextOff, C_ClaudeTextOn, C_ClaudeBdrOff, C_ClaudeBdrOn;

    // =====================================================================
    //  Entry point
    // =====================================================================

    [STAThread]
    static void Main(string[] args)
    {
        if (args.Length > 0)
        {
            string a = args[0].ToLowerInvariant();
            if (a == "--install" || a == "/install") { DoInstall(); return; }
            if (a == "--uninstall" || a == "/uninstall") { DoUninstall(); return; }
        }
        // Parse restart count for crash-loop prevention
        int rc = 0;
        foreach (string arg in args)
            if (arg.StartsWith("--restart-count="))
                int.TryParse(arg.Substring(16), out rc);
        if (rc > 0) Thread.Sleep(2000); // delay before restart attempt
        var qp = new QuickPaths();
        qp.restartCount = rc;
        qp.Run();
    }

    // =====================================================================
    //  Normal run
    // =====================================================================

    void Run()
    {
        // --- Singleton ---
        mutex = new Mutex(false, MUTEX_NAME);
        bool owned = false;
        try { owned = mutex.WaitOne(0); }
        catch (AbandonedMutexException) { owned = true; }
        if (!owned) { Environment.ExitCode = 0; return; }

        // --- Paths ---
        dir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
        dataFile = Path.Combine(dir, "paths.json");
        configFile = Path.Combine(dir, "config.json");
        logFile = Path.Combine(dir, "QuickPaths.log");
        startTime = DateTime.Now;
        Log("QuickPaths starting...");

        // --- Data ---
        LoadPaths();
        AppConfig cfg = LoadConfig();

        // --- Brushes ---
        InitBrushes();

        // --- Window ---
        win = new Window();
        win.WindowStyle = WindowStyle.None;
        win.AllowsTransparency = true;
        win.Background = Brushes.Transparent;
        win.Topmost = true;
        win.ShowInTaskbar = false;
        win.SizeToContent = SizeToContent.WidthAndHeight;
        win.ResizeMode = ResizeMode.NoResize;
        win.Left = cfg.Left;
        win.Top = cfg.Top;

        root = new Grid();

        // --- ECG dot ---
        double initialScale = cfg.Scale > 0 ? cfg.Scale : 1.0;
        dotScale = new ScaleTransform(initialScale, initialScale);

        dot = new Border();
        dot.Width = 88;
        dot.Height = 38;
        dot.CornerRadius = new CornerRadius(6);
        dot.Background = C_MonBg;
        dot.Cursor = Cursors.Hand;
        dot.BorderThickness = new Thickness(1);
        dot.LayoutTransform = dotScale;
        dot.Effect = new DropShadowEffect
        {
            BlurRadius = 12, ShadowDepth = 1, Opacity = 0.35,
            Color = Colors.Black
        };

        canvas = new Canvas();
        canvas.ClipToBounds = false;
        canvas.IsHitTestVisible = false;
        dot.Child = canvas;

        // Waveform polyline (360 points scrolling left)
        double yBot = ECG_PAD_Y + ECG_WAVE_H;
        waveLine = new Polyline();
        waveLine.StrokeThickness = 1.5;
        waveLine.IsHitTestVisible = false;

        yValues.Clear();
        var initPts = new PointCollection();
        for (int i = 0; i < ECG_COUNT; i++)
        {
            yValues.Add(yBot);
            double x = ECG_PAD_X + ECG_WAVE_W * (double)i / (ECG_COUNT - 1);
            initPts.Add(new Point(x, yBot));
        }
        waveLine.Points = initPts;
        canvas.Children.Add(waveLine);

        // Pen head glow layers (outer -> middle -> head, painted in order)
        double glowRX = ECG_PAD_X + ECG_WAVE_W; // right-edge X anchor

        penGlow2 = new Ellipse();
        penGlow2.Width = 30;
        penGlow2.Height = 30;
        penGlow2.IsHitTestVisible = false;
        penGlow2.Opacity = 0.0;
        Canvas.SetLeft(penGlow2, glowRX - 15);
        Canvas.SetTop(penGlow2, yBot - 15);
        canvas.Children.Add(penGlow2);

        penGlow1 = new Ellipse();
        penGlow1.Width = 16;
        penGlow1.Height = 16;
        penGlow1.IsHitTestVisible = false;
        penGlow1.Opacity = 0.0;
        Canvas.SetLeft(penGlow1, glowRX - 8);
        Canvas.SetTop(penGlow1, yBot - 8);
        canvas.Children.Add(penGlow1);

        penHead = new Ellipse();
        penHead.Width = 6;
        penHead.Height = 6;
        penHead.IsHitTestVisible = false;
        penHeadShadow = new DropShadowEffect
        {
            BlurRadius = 15, ShadowDepth = 0, Opacity = 0.8
        };
        penHead.Effect = penHeadShadow;
        Canvas.SetLeft(penHead, glowRX - 3);
        Canvas.SetTop(penHead, yBot - 3);
        canvas.Children.Add(penHead);

        // --- Breath timer (50ms tick, 18s cycle) ---
        breathTimer = new DispatcherTimer();
        breathTimer.Interval = TimeSpan.FromMilliseconds(50);
        breathTimer.Tick += OnBreathTick;

        UpdateDotAppearance(); // sets colors, starts breathTimer

        // --- Panel ---
        panel = new Border();
        panel.CornerRadius = new CornerRadius(12);
        panel.Background = C_PanelBg;
        panel.Padding = new Thickness(10, 8, 10, 8);
        panel.MinWidth = 160;
        panel.MaxWidth = 280;
        panel.Visibility = Visibility.Collapsed;
        panel.Effect = new DropShadowEffect
        {
            BlurRadius = 20, ShadowDepth = 2, Opacity = 0.40,
            Color = Colors.Black
        };

        listPanel = new StackPanel();
        panel.Child = listPanel;

        root.Children.Add(dot);
        root.Children.Add(panel);
        win.Content = root;
        Log("Window built at (" + (int)win.Left + ", " + (int)win.Top + ")");

        // --- Flash timer ---
        flashTimer = new DispatcherTimer();
        flashTimer.Interval = TimeSpan.FromMilliseconds(700);
        flashTimer.Tick += delegate
        {
            try
            {
                if (claudeMode)
                {
                    waveLine.Stroke = C_WaveOrange;
                    penHead.Fill = C_PenOrange;
                }
                else
                {
                    waveLine.Stroke = C_WaveBlue;
                    penHead.Fill = C_PenBlue;
                }
                flashTimer.Stop();
            }
            catch (Exception ex) { Log("flashTimer error: " + ex.Message); flashTimer.Stop(); }
        };

        // --- Dot events ---
        dot.MouseEnter += delegate { try { dot.Background = C_MonHover; } catch { } };
        dot.MouseLeave += delegate { try { dot.Background = C_MonBg; } catch { } };

        dot.MouseLeftButtonDown += delegate
        {
            try
            {
                POINT pt;
                GetCursorPos(out pt);
                dotClickPt = pt;
            }
            catch { }
        };

        dot.MouseMove += delegate(object s, MouseEventArgs e)
        {
            try
            {
                if (e.LeftButton == MouseButtonState.Pressed && dotClickPt.HasValue)
                {
                    POINT cur;
                    GetCursorPos(out cur);
                    if (Math.Abs(cur.X - dotClickPt.Value.X) > 5 ||
                        Math.Abs(cur.Y - dotClickPt.Value.Y) > 5)
                    {
                        dotClickPt = null;
                        win.DragMove();
                        SaveConfig();
                    }
                }
            }
            catch (Exception ex) { Log("dot.MouseMove error: " + ex.Message); }
        };

        dot.MouseLeftButtonUp += delegate
        {
            try
            {
                if (dotClickPt.HasValue) Expand();
                dotClickPt = null;
            }
            catch (Exception ex) { Log("dot.MouseUp error: " + ex.Message); }
        };

        // Scroll wheel to resize dot
        dot.MouseWheel += delegate(object s, MouseWheelEventArgs e)
        {
            try
            {
                double delta = e.Delta > 0 ? 0.1 : -0.1;
                double ns = Math.Round(dotScale.ScaleX + delta, 1);
                ns = Math.Max(0.5, Math.Min(3.0, ns));
                dotScale.ScaleX = ns;
                dotScale.ScaleY = ns;
                SaveConfig();
            }
            catch (Exception ex) { Log("dot.MouseWheel error: " + ex.Message); }
        };

        // Right-click dot to exit
        dot.MouseRightButtonUp += delegate
        {
            try
            {
                var r = MessageBox.Show(
                    "\u9000\u51FA QuickPaths\uFF1F", "\u786E\u8BA4",
                    MessageBoxButton.YesNo, MessageBoxImage.Question);
                if (r == MessageBoxResult.Yes)
                {
                    intentionalExit = true;
                    exitCode = 0;
                    win.Close();
                }
            }
            catch (Exception ex) { Log("RightClick error: " + ex.Message); }
        };

        // Click outside to collapse
        win.Deactivated += delegate
        {
            try { if (expanded && !dialogOpen) Collapse(); }
            catch (Exception ex) { Log("Deactivated error: " + ex.Message); }
        };

        // --- System events ---
        onPowerMode = delegate(object sender, PowerModeChangedEventArgs e)
        {
            try
            {
                if (e.Mode == PowerModes.Resume)
                {
                    Log("System resumed from sleep");
                    ReassertTopmost();
                }
            }
            catch (Exception ex) { Log("PowerModeChanged error: " + ex.Message); }
        };
        SystemEvents.PowerModeChanged += onPowerMode;

        onSessionSwitch = delegate(object sender, SessionSwitchEventArgs e)
        {
            try
            {
                if (e.Reason == SessionSwitchReason.SessionUnlock)
                {
                    Log("Session unlocked");
                    ReassertTopmost();
                }
            }
            catch (Exception ex) { Log("SessionSwitch error: " + ex.Message); }
        };
        SystemEvents.SessionSwitch += onSessionSwitch;

        // Display hotplug (1500ms debounce)
        displayChangeTimer = new DispatcherTimer();
        displayChangeTimer.Interval = TimeSpan.FromMilliseconds(1500);
        displayChangeTimer.Tick += delegate
        {
            try
            {
                displayChangeTimer.Stop();
                double vl = SystemParameters.VirtualScreenLeft;
                double vt = SystemParameters.VirtualScreenTop;
                double vw = SystemParameters.VirtualScreenWidth;
                double vh = SystemParameters.VirtualScreenHeight;
                double wl = win.Left, wt = win.Top;
                Log("DisplayChange settled: Window=(" + (int)wl + "," + (int)wt
                    + ") VirtualScreen=(" + (int)vl + "," + (int)vt + "," + (int)vw + "," + (int)vh + ")");
                if (wl < vl || wl > vl + vw - 30 || wt < vt || wt > vt + vh - 30)
                {
                    Log("Window off-screen, resetting");
                    Rect wa = SystemParameters.WorkArea;
                    win.Left = wa.Width - 60;
                    win.Top = wa.Height * 0.4;
                    SaveConfig();
                }
                win.Topmost = false;
                win.Topmost = true;
            }
            catch (Exception ex) { Log("displayChangeTimer error: " + ex.Message); }
        };

        onDisplayChanged = delegate
        {
            try
            {
                Log("Display settings changed (debouncing...)");
                win.Dispatcher.BeginInvoke((Action)delegate
                {
                    displayChangeTimer.Stop();
                    displayChangeTimer.Start();
                });
            }
            catch (Exception ex) { Log("DisplaySettingsChanged error: " + ex.Message); }
        };
        SystemEvents.DisplaySettingsChanged += onDisplayChanged;

        // Health timer: re-assert Topmost every 5 minutes
        healthTimer = new DispatcherTimer();
        healthTimer.Interval = TimeSpan.FromMinutes(5);
        healthTimer.Tick += delegate
        {
            try { win.Topmost = false; win.Topmost = true; }
            catch (Exception ex) { Log("healthTimer error: " + ex.Message); }
        };
        healthTimer.Start();

        // --- Window.Closed cleanup ---
        win.Closed += delegate
        {
            Log("Window closed (intentional=" + intentionalExit + " exitCode=" + exitCode + ")");
            healthTimer.Stop();
            breathTimer.Stop();
            displayChangeTimer.Stop();
            SaveConfig();
            if (intentionalExit) exitCode = 0;
            try { SystemEvents.PowerModeChanged -= onPowerMode; } catch { }
            try { SystemEvents.SessionSwitch -= onSessionSwitch; } catch { }
            try { SystemEvents.DisplaySettingsChanged -= onDisplayChanged; } catch { }
            try { mutex.ReleaseMutex(); } catch { }
        };

        // --- Launch ---
        Log("Launching...");
        var app = new Application();

        app.DispatcherUnhandledException += delegate(object sender, DispatcherUnhandledExceptionEventArgs e)
        {
            e.Handled = true;
            LogError("Dispatcher.UnhandledException", e.Exception);
            LogCrashContext();
            string msg = e.Exception.GetType().FullName + " " + e.Exception.Message;
            bool fatal = msg.Contains("DUCE") || msg.Contains("Render") || msg.Contains("UCEERR")
                         || msg.Contains("COMException") || msg.Contains("OutOfMemoryException");
            if (fatal)
            {
                Log("FATAL dispatcher error \u2014 closing for restart");
                exitCode = 2;
                try { win.Close(); } catch { }
            }
            else
            {
                Log("Non-fatal dispatcher error \u2014 handled, continuing");
            }
        };

        AppDomain.CurrentDomain.UnhandledException += delegate(object sender, UnhandledExceptionEventArgs e)
        {
            var ex = e.ExceptionObject as Exception;
            if (ex != null) LogError("AppDomain.UnhandledException", ex);
            LogCrashContext();
            exitCode = 3;
            try { win.Close(); } catch { }
        };

        try
        {
            app.Run(win);
        }
        catch (Exception ex)
        {
            LogError("Application.Run", ex);
            LogCrashContext();
            exitCode = 1;
        }
        Log("Exiting with code " + exitCode);

        // Self-restart on crash (up to 5 attempts, then wait for keepalive task)
        if (exitCode != 0 && restartCount < 5)
        {
            Log("Scheduling restart (attempt " + (restartCount + 1) + "/5)...");
            try
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = Assembly.GetExecutingAssembly().Location,
                    Arguments = "--restart-count=" + (restartCount + 1),
                    UseShellExecute = true
                });
            }
            catch (Exception ex) { Log("Restart failed: " + ex.Message); }
        }

        Environment.ExitCode = exitCode;
    }

    // =====================================================================
    //  UI logic
    // =====================================================================

    void Collapse()
    {
        expanded = false;
        panel.Visibility = Visibility.Collapsed;
        dot.Visibility = Visibility.Visible;
    }

    void Expand()
    {
        expanded = true;
        dot.Visibility = Visibility.Collapsed;
        RebuildList();
        panel.Visibility = Visibility.Visible;
    }

    void FlashDot()
    {
        flashTimer.Stop();
        waveLine.Stroke = C_WaveGreen;
        penHead.Fill = C_PenGreen;
        flashTimer.Start();
    }

    void ReassertTopmost()
    {
        win.Dispatcher.BeginInvoke((Action)delegate
        {
            win.Topmost = false;
            win.Topmost = true;
        });
    }

    void UpdateDotAppearance()
    {
        if (claudeMode)
        {
            dot.BorderBrush = C_MonBdrOrange;
            waveLine.Stroke = C_WaveOrange;
            penHead.Fill = C_PenOrange;
            penHeadShadow.Color = Color.FromRgb(217, 119, 87);
            Color cOrange = Color.FromRgb(217, 119, 87);
            penGlow1.Fill = MakeGlowBrush(cOrange, 0.7, 0.0);
            penGlow2.Fill = MakeGlowBrush(cOrange, 0.4, 0.0);
        }
        else
        {
            dot.BorderBrush = C_MonBdrBlue;
            waveLine.Stroke = C_WaveBlue;
            penHead.Fill = C_PenBlue;
            penHeadShadow.Color = Color.FromRgb(106, 155, 204);
            Color cBlue = Color.FromRgb(106, 155, 204);
            penGlow1.Fill = MakeGlowBrush(cBlue, 0.7, 0.0);
            penGlow2.Fill = MakeGlowBrush(cBlue, 0.4, 0.0);
        }
        if (!breathTimer.IsEnabled)
        {
            breathTick = 0;
            breathTimer.Start();
        }
    }

    RadialGradientBrush MakeGlowBrush(Color c, double centerAlpha, double edgeAlpha)
    {
        Color center = Color.FromArgb((byte)Math.Min(255, centerAlpha * 255), c.R, c.G, c.B);
        Color edge = Color.FromArgb((byte)Math.Min(255, edgeAlpha * 255), c.R, c.G, c.B);
        var brush = new RadialGradientBrush();
        brush.GradientStops.Add(new GradientStop(center, 0.0));
        brush.GradientStops.Add(new GradientStop(edge, 1.0));
        return brush;
    }

    // --- Breath timer tick ---
    void OnBreathTick(object sender, EventArgs e)
    {
        try
        {
            breathTick++;
            if (breathTick >= 360) breathTick = 0;
            int k = breathTick;
            double v;
            if (k < 80)
                v = (1 - Math.Cos(k / 80.0 * Math.PI)) / 2; // inhale 4s
            else if (k < 160)
                v = 1.0; // hold 4s
            else if (k < 320)
                v = (1 + Math.Cos((k - 160) / 160.0 * Math.PI)) / 2; // exhale 8s
            else
                v = 0.0; // rest 2s

            double y = ECG_PAD_Y + ECG_WAVE_H * (1 - v);
            yValues.RemoveAt(0);
            yValues.Add(y);

            var pts = new PointCollection();
            for (int i = 0; i < ECG_COUNT; i++)
            {
                double x = ECG_PAD_X + ECG_WAVE_W * (double)i / (ECG_COUNT - 1);
                pts.Add(new Point(x, yValues[i]));
            }
            waveLine.Points = pts;

            Canvas.SetTop(penHead, y - 3);
            Canvas.SetTop(penGlow1, y - 8);
            Canvas.SetTop(penGlow2, y - 15);

            penHeadShadow.Opacity = 0.4 + 0.55 * v;
            penGlow1.Opacity = 0.18 + 0.50 * v;
            penGlow2.Opacity = 0.08 + 0.30 * v;
        }
        catch (Exception ex)
        {
            Log("breathTimer error: " + ex.Message);
            breathTimer.Stop();
        }
    }

    // =====================================================================
    //  Rebuild path list panel
    // =====================================================================

    void RebuildList()
    {
        // Prune invalid paths
        int removed = 0;
        for (int i = paths.Count - 1; i >= 0; i--)
        {
            if (!Directory.Exists(paths[i].ItemPath))
            {
                Log("Pruned: " + paths[i].Name + " (path gone)");
                paths.RemoveAt(i);
                removed++;
            }
        }
        if (removed > 0) SavePaths();

        listPanel.Children.Clear();

        // --- Claude mode toggle ---
        var claudeBtn = new Border();
        claudeBtn.Padding = new Thickness(10, 6, 10, 6);
        claudeBtn.CornerRadius = new CornerRadius(7);
        claudeBtn.Cursor = Cursors.Hand;
        claudeBtn.Margin = new Thickness(0, 0, 0, 4);

        var claudeText = new TextBlock();
        claudeText.Text = "Claude";
        claudeText.FontSize = 13;
        claudeText.FontWeight = FontWeights.Medium;
        claudeText.HorizontalAlignment = HorizontalAlignment.Center;
        claudeText.IsHitTestVisible = false;
        claudeBtn.Child = claudeText;

        if (claudeMode)
        {
            claudeBtn.Background = C_ClaudeOn;
            claudeBtn.BorderBrush = C_ClaudeBdrOn;
            claudeBtn.BorderThickness = new Thickness(1.5);
            claudeText.Foreground = C_ClaudeTextOn;
            claudeBtn.Effect = new DropShadowEffect
            {
                BlurRadius = 20, ShadowDepth = 0, Opacity = 0.7,
                Color = Color.FromRgb(217, 119, 87)
            };
        }
        else
        {
            claudeBtn.Background = C_ClaudeOff;
            claudeBtn.BorderBrush = C_ClaudeBdrOff;
            claudeBtn.BorderThickness = new Thickness(1);
            claudeText.Foreground = C_ClaudeTextOff;
            claudeBtn.Effect = null;
        }

        claudeBtn.MouseEnter += delegate(object s, MouseEventArgs ev)
        {
            ((Border)s).Background = claudeMode ? C_ClaudeHoverOn : C_ClaudeHoverOff;
        };
        claudeBtn.MouseLeave += delegate(object s, MouseEventArgs ev)
        {
            ((Border)s).Background = claudeMode ? C_ClaudeOn : C_ClaudeOff;
        };
        claudeBtn.MouseLeftButtonDown += delegate
        {
            claudeMode = !claudeMode;
            SaveConfig();
            UpdateDotAppearance();
            RebuildList();
        };
        listPanel.Children.Add(claudeBtn);

        // --- Empty hint ---
        if (paths.Count == 0)
        {
            var hint = new TextBlock();
            hint.Text = "\u70B9\u51FB + "; // "点击 + "
            hint.Foreground = Br(80, 140, 140, 140);
            hint.FontSize = 12;
            hint.HorizontalAlignment = HorizontalAlignment.Center;
            hint.Margin = new Thickness(8, 4, 8, 2);
            listPanel.Children.Add(hint);
        }

        // --- Path items ---
        for (int idx = 0; idx < paths.Count; idx++)
        {
            PathItem item = paths[idx];
            var row = new DockPanel();
            row.Margin = new Thickness(0, 2, 0, 2);

            // Delete button (rightmost)
            var del = new Border();
            del.Width = 24;
            del.Height = 24;
            del.CornerRadius = new CornerRadius(12);
            del.Background = C_DelNormal;
            del.Cursor = Cursors.Hand;
            del.Tag = item.ItemPath;
            del.Margin = new Thickness(3, 0, 0, 0);
            DockPanel.SetDock(del, Dock.Right);

            var delText = new TextBlock();
            delText.Text = "\u2212"; // minus sign
            delText.Foreground = C_DelText;
            delText.FontSize = 14;
            delText.FontWeight = FontWeights.Medium;
            delText.HorizontalAlignment = HorizontalAlignment.Center;
            delText.VerticalAlignment = VerticalAlignment.Center;
            delText.IsHitTestVisible = false;
            del.Child = delText;

            del.MouseEnter += delegate(object s, MouseEventArgs ev) { ((Border)s).Background = C_DelHover; };
            del.MouseLeave += delegate(object s, MouseEventArgs ev) { ((Border)s).Background = C_DelNormal; };
            del.MouseLeftButtonDown += delegate(object s, MouseButtonEventArgs ev)
            {
                string target = (string)((Border)s).Tag;
                for (int i = paths.Count - 1; i >= 0; i--)
                {
                    if (paths[i].ItemPath == target) { paths.RemoveAt(i); break; }
                }
                SavePaths();
                RebuildList();
            };
            row.Children.Add(del);

            // Move-up button
            if (idx > 0)
            {
                var up = new Border();
                up.Width = 24;
                up.Height = 24;
                up.CornerRadius = new CornerRadius(12);
                up.Background = C_UpNormal;
                up.Cursor = Cursors.Hand;
                up.Tag = item.ItemPath;
                up.Margin = new Thickness(3, 0, 0, 0);
                DockPanel.SetDock(up, Dock.Right);

                var upText = new TextBlock();
                upText.Text = "\u2191"; // up arrow
                upText.Foreground = C_UpText;
                upText.FontSize = 13;
                upText.HorizontalAlignment = HorizontalAlignment.Center;
                upText.VerticalAlignment = VerticalAlignment.Center;
                upText.IsHitTestVisible = false;
                up.Child = upText;

                up.MouseEnter += delegate(object s, MouseEventArgs ev) { ((Border)s).Background = C_UpHover; };
                up.MouseLeave += delegate(object s, MouseEventArgs ev) { ((Border)s).Background = C_UpNormal; };
                up.MouseLeftButtonDown += delegate(object s, MouseButtonEventArgs ev)
                {
                    string target = (string)((Border)s).Tag;
                    for (int i = 1; i < paths.Count; i++)
                    {
                        if (paths[i].ItemPath == target)
                        {
                            var tmp = paths[i - 1];
                            paths[i - 1] = paths[i];
                            paths[i] = tmp;
                            break;
                        }
                    }
                    SavePaths();
                    RebuildList();
                };
                row.Children.Add(up);
            }

            // Name button (fills remaining space)
            var nameBtn = new Border();
            nameBtn.Padding = new Thickness(10, 7, 10, 7);
            nameBtn.CornerRadius = new CornerRadius(7);
            nameBtn.Background = C_ItemNormal;
            nameBtn.Cursor = Cursors.Hand;
            nameBtn.Tag = item.ItemPath;
            nameBtn.ToolTip = item.ItemPath;

            var nameText = new TextBlock();
            nameText.Text = item.Name;
            nameText.Foreground = Brushes.White;
            nameText.FontSize = 13;
            nameText.TextTrimming = TextTrimming.CharacterEllipsis;
            nameText.IsHitTestVisible = false;
            nameBtn.Child = nameText;

            nameBtn.MouseEnter += delegate(object s, MouseEventArgs ev) { ((Border)s).Background = C_ItemHover; };
            nameBtn.MouseLeave += delegate(object s, MouseEventArgs ev) { ((Border)s).Background = C_ItemNormal; };
            nameBtn.MouseLeftButtonDown += delegate(object s, MouseButtonEventArgs ev)
            {
                string p = (string)((Border)s).Tag;
                if (claudeMode)
                {
                    try
                    {
                        Process.Start(new ProcessStartInfo
                        {
                            FileName = "powershell.exe",
                            Arguments = "-NoExit -Command \"cd '" + p + "'; claude\""
                        });
                    }
                    catch (Exception ex) { Log("Claude launch error: " + ex.Message); }
                }
                else
                {
                    try { Clipboard.SetText(p); }
                    catch (Exception ex) { Log("Clipboard error: " + ex.Message); return; }
                }
                Collapse();
                FlashDot();
            };
            row.Children.Add(nameBtn);
            listPanel.Children.Add(row);
        }

        // --- Separator ---
        if (paths.Count > 0)
        {
            var sep = new Border();
            sep.Height = 1;
            sep.Background = C_Sep;
            sep.Margin = new Thickness(4, 6, 4, 4);
            listPanel.Children.Add(sep);
        }

        // --- Add button ---
        var addBtn = new Border();
        addBtn.Padding = new Thickness(10, 5, 10, 5);
        addBtn.CornerRadius = new CornerRadius(7);
        addBtn.Background = C_AddBg;
        addBtn.Cursor = Cursors.Hand;

        var addText = new TextBlock();
        addText.Text = "+";
        addText.Foreground = C_AddText;
        addText.FontSize = 16;
        addText.HorizontalAlignment = HorizontalAlignment.Center;
        addText.IsHitTestVisible = false;
        addBtn.Child = addText;

        addBtn.MouseEnter += delegate(object s, MouseEventArgs ev) { ((Border)s).Background = C_AddHover; };
        addBtn.MouseLeave += delegate(object s, MouseEventArgs ev) { ((Border)s).Background = C_AddBg; };
        addBtn.MouseLeftButtonDown += delegate
        {
            dialogOpen = true;
            try
            {
                win.Topmost = false;
                // "选择文件夹  (Ctrl+点击多选)"
                string dlgTitle = "\u9009\u62E9\u6587\u4EF6\u5939  (Ctrl+\u70B9\u51FB\u591A\u9009)";
                string[] selected = FolderMultiPicker.ShowDialog(dlgTitle);
                int added = 0;
                foreach (string fp in selected)
                {
                    string fn = Path.GetFileName(fp);
                    bool dup = paths.Any(p => p.ItemPath == fp);
                    if (!dup)
                    {
                        paths.Insert(0, new PathItem { Name = fn, ItemPath = fp });
                        added++;
                    }
                }
                if (added > 0)
                {
                    Log("Added " + added + " folder(s)");
                    SavePaths();
                }
            }
            catch (Exception ex) { Log("Add error: " + ex.Message); }
            finally
            {
                dialogOpen = false;
                win.Topmost = true;
                win.Activate();
                if (expanded) RebuildList();
            }
        };
        listPanel.Children.Add(addBtn);
    }

    // =====================================================================
    //  JSON persistence
    // =====================================================================

    void LoadPaths()
    {
        paths.Clear();
        try
        {
            if (File.Exists(dataFile))
            {
                string json = File.ReadAllText(dataFile, Encoding.UTF8).TrimStart('\uFEFF');
                if (!string.IsNullOrEmpty(json) && json.Trim().Length > 2)
                {
                    var ser = new DataContractJsonSerializer(typeof(List<PathItem>));
                    using (var ms = new MemoryStream(Encoding.UTF8.GetBytes(json)))
                    {
                        paths = (List<PathItem>)ser.ReadObject(ms);
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Log("LoadPaths error: " + ex.Message);
            paths.Clear();
        }
        Log("Loaded " + paths.Count + " paths");
    }

    void SavePaths()
    {
        try
        {
            var noBom = new UTF8Encoding(false);
            string json;
            if (paths.Count == 0)
            {
                json = "[]";
            }
            else
            {
                var ser = new DataContractJsonSerializer(typeof(List<PathItem>));
                using (var ms = new MemoryStream())
                {
                    ser.WriteObject(ms, paths);
                    json = Encoding.UTF8.GetString(ms.ToArray());
                }
            }
            // Safe write: tmp -> delete original -> rename
            string tmp = dataFile + ".tmp";
            File.WriteAllText(tmp, json, noBom);
            if (File.Exists(dataFile)) File.Delete(dataFile);
            File.Move(tmp, dataFile);
        }
        catch (Exception ex)
        {
            Log("SavePaths error: " + ex.Message);
            try
            {
                // Fallback: direct write
                var noBom = new UTF8Encoding(false);
                var ser = new DataContractJsonSerializer(typeof(List<PathItem>));
                using (var ms = new MemoryStream())
                {
                    ser.WriteObject(ms, paths);
                    File.WriteAllText(dataFile, Encoding.UTF8.GetString(ms.ToArray()), noBom);
                }
            }
            catch { }
            try { if (File.Exists(dataFile + ".tmp")) File.Delete(dataFile + ".tmp"); } catch { }
        }
    }

    AppConfig LoadConfig()
    {
        AppConfig cfg = null;
        try
        {
            if (File.Exists(configFile))
            {
                string raw = File.ReadAllText(configFile, Encoding.UTF8).TrimStart('\uFEFF');
                var ser = new DataContractJsonSerializer(typeof(AppConfig));
                using (var ms = new MemoryStream(Encoding.UTF8.GetBytes(raw)))
                {
                    cfg = (AppConfig)ser.ReadObject(ms);
                }
            }
        }
        catch { }

        if (cfg != null) claudeMode = cfg.ClaudeMode;

        double vl = SystemParameters.VirtualScreenLeft;
        double vt = SystemParameters.VirtualScreenTop;
        double vw = SystemParameters.VirtualScreenWidth;
        double vh = SystemParameters.VirtualScreenHeight;
        Log("LoadConfig: saved=(" + (cfg != null ? cfg.Left.ToString() : "null") + ","
            + (cfg != null ? cfg.Top.ToString() : "null") + ") VirtualScreen=("
            + (int)vl + "," + (int)vt + "," + (int)vw + "," + (int)vh + ")");

        if (cfg == null || cfg.Left < vl || cfg.Left > vl + vw - 30
                        || cfg.Top < vt || cfg.Top > vt + vh - 30)
        {
            Rect wa = SystemParameters.WorkArea;
            cfg = new AppConfig
            {
                Left = (int)(wa.Width - 60),
                Top = (int)(wa.Height * 0.4),
                ClaudeMode = claudeMode,
                Scale = 1.0
            };
            Log("LoadConfig: out of bounds, using default (" + cfg.Left + "," + cfg.Top + ")");
        }
        if (cfg.Scale <= 0) cfg.Scale = 1.0;
        return cfg;
    }

    void SaveConfig()
    {
        try
        {
            int l = (int)win.Left, t = (int)win.Top;
            Log("SaveConfig: (" + l + "," + t + ")");
            var cfg = new AppConfig
            {
                Left = l, Top = t,
                ClaudeMode = claudeMode,
                Scale = dotScale.ScaleX
            };
            var ser = new DataContractJsonSerializer(typeof(AppConfig));
            var noBom = new UTF8Encoding(false);
            using (var ms = new MemoryStream())
            {
                ser.WriteObject(ms, cfg);
                File.WriteAllText(configFile, Encoding.UTF8.GetString(ms.ToArray()), noBom);
            }
        }
        catch { }
    }

    // =====================================================================
    //  Logging
    // =====================================================================

    void Log(string msg)
    {
        try
        {
            var fi = new FileInfo(logFile);
            if (fi.Exists && fi.Length > 512 * 1024)
            {
                string bak = logFile + ".old";
                if (File.Exists(bak)) File.Delete(bak);
                File.Move(logFile, bak);
            }
            string ts = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
            File.AppendAllText(logFile, "[" + ts + "] " + msg + "\r\n", Encoding.UTF8);
        }
        catch { }
    }

    void LogError(string context, Exception ex)
    {
        Log("ERROR [" + context + "] " + ex.GetType().FullName + ": " + ex.Message);
        if (ex.StackTrace != null) Log("  StackTrace: " + ex.StackTrace);
        Exception inner = ex.InnerException;
        int depth = 0;
        while (inner != null && depth < 5)
        {
            Log("  Inner[" + depth + "]: " + inner.GetType().FullName + ": " + inner.Message);
            inner = inner.InnerException;
            depth++;
        }
    }

    void LogCrashContext()
    {
        try
        {
            string uptime = startTime != default(DateTime)
                ? (DateTime.Now - startTime).ToString(@"hh\:mm\:ss\.fff") : "N/A";
            Log("  CrashContext: PID=" + Process.GetCurrentProcess().Id
                + " Uptime=" + uptime + " Expanded=" + expanded + " DialogOpen=" + dialogOpen);
            double vl = SystemParameters.VirtualScreenLeft;
            double vt = SystemParameters.VirtualScreenTop;
            double vw = SystemParameters.VirtualScreenWidth;
            double vh = SystemParameters.VirtualScreenHeight;
            Log("  CrashContext: VirtualScreen=(" + (int)vl + "," + (int)vt + "," + (int)vw + "," + (int)vh + ")");
            if (win != null)
                Log("  CrashContext: WindowPos=(" + (int)win.Left + "," + (int)win.Top + ") Topmost=" + win.Topmost);
            double mem = Math.Round(GC.GetTotalMemory(false) / 1048576.0, 1);
            Log("  CrashContext: ManagedMemory=" + mem + "MB");
        }
        catch { }
    }

    // =====================================================================
    //  Brush helpers
    // =====================================================================

    static SolidColorBrush Br(byte a, byte r, byte g, byte b)
    {
        var brush = new SolidColorBrush(Color.FromArgb(a, r, g, b));
        brush.Freeze();
        return brush;
    }

    void InitBrushes()
    {
        // ECG monitor — Anthropic brand palette
        C_MonBg        = Br(240, 18, 20, 26);
        C_MonHover     = Br(250, 28, 30, 38);
        C_MonBdrBlue   = Br(55, 106, 155, 204);
        C_MonBdrOrange = Br(55, 217, 119, 87);
        C_WaveBlue     = Br(190, 106, 155, 204);
        C_WaveOrange   = Br(190, 217, 119, 87);
        C_WaveGreen    = Br(210, 120, 140, 93);
        C_PenBlue      = Br(255, 155, 200, 240);
        C_PenOrange    = Br(255, 245, 160, 125);
        C_PenGreen     = Br(255, 165, 190, 140);
        // Panel
        C_PanelBg    = Br(245, 32, 34, 38);
        C_ItemNormal = Br(22, 255, 255, 255);
        C_ItemHover  = Br(50, 255, 255, 255);
        C_UpNormal   = Brushes.Transparent;
        C_UpHover    = Br(35, 255, 255, 255);
        C_UpText     = Br(100, 180, 180, 180);
        C_DelNormal  = Brushes.Transparent;
        C_DelHover   = Br(40, 255, 80, 80);
        C_DelText    = Br(120, 255, 100, 100);
        C_AddBg      = Brushes.Transparent;
        C_AddHover   = Br(22, 255, 255, 255);
        C_AddText    = Br(110, 180, 180, 180);
        C_Sep        = Br(18, 255, 255, 255);
        // Claude toggle — Anthropic brand orange #d97757
        C_ClaudeOff      = Br(15, 250, 249, 245);
        C_ClaudeOn       = Br(130, 217, 119, 87);
        C_ClaudeHoverOff = Br(28, 250, 249, 245);
        C_ClaudeHoverOn  = Br(150, 217, 119, 87);
        C_ClaudeTextOff  = Br(75, 176, 174, 165);
        C_ClaudeTextOn   = Br(255, 250, 249, 245);
        C_ClaudeBdrOff   = Br(22, 176, 174, 165);
        C_ClaudeBdrOn    = Br(180, 217, 119, 87);
    }

    // =====================================================================
    //  Install / Uninstall
    // =====================================================================

    static void DoInstall()
    {
        string exePath = Assembly.GetExecutingAssembly().Location;
        string exeDir = Path.GetDirectoryName(exePath);

        // 1. Kill old instances
        KillOldInstances();

        // 2. Clean up old Task Scheduler watchdog
        RunHidden("schtasks", "/Delete /TN QuickPaths_Watchdog /F");

        // 3. Clean up old watchdog files
        TryDelete(Path.Combine(exeDir, "watchdog.ps1"));
        TryDelete(Path.Combine(exeDir, "watchdog.vbs"));
        TryDelete(Path.Combine(exeDir, "wrapper.lock"));

        // 4. Remove old Startup VBS
        string startupDir = Environment.GetFolderPath(Environment.SpecialFolder.Startup);
        TryDelete(Path.Combine(startupDir, "QuickPaths.vbs"));

        // 5. Register auto-start via Registry
        try
        {
            using (RegistryKey key = Registry.CurrentUser.OpenSubKey(REG_RUN, true))
            {
                if (key != null)
                    key.SetValue("QuickPaths", "\"" + exePath + "\"");
            }
        }
        catch { }

        // 6. Register keepalive scheduled task (every 5 minutes, singleton mutex prevents duplicates)
        RunHidden("schtasks", "/Create /TN \"QuickPaths_KeepAlive\" /SC MINUTE /MO 5 /TR \"\\\"" + exePath + "\\\"\" /F");

        // 7. Launch QuickPaths (normal mode)
        try
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = exePath,
                UseShellExecute = true
            });
        }
        catch { }

        // 8. Success message
        MessageBox.Show(
            "QuickPaths installed successfully!\n\n"
            + "  Auto-start: Enabled (registry + keepalive task)\n"
            + "  Location: " + exeDir + "\n\n"
            + "The floating dot should appear on your desktop.",
            "QuickPaths Install", MessageBoxButton.OK, MessageBoxImage.Information);
    }

    static void DoUninstall()
    {
        string exePath = Assembly.GetExecutingAssembly().Location;
        string exeDir = Path.GetDirectoryName(exePath);

        // 1. Kill running instances
        KillOldInstances();

        // 2. Remove auto-start registry key
        try
        {
            using (RegistryKey key = Registry.CurrentUser.OpenSubKey(REG_RUN, true))
            {
                if (key != null) key.DeleteValue("QuickPaths", false);
            }
        }
        catch { }

        // 3. Remove old Startup VBS (backward compat)
        string startupDir = Environment.GetFolderPath(Environment.SpecialFolder.Startup);
        TryDelete(Path.Combine(startupDir, "QuickPaths.vbs"));

        // 4. Remove scheduled tasks (keepalive + old watchdog)
        RunHidden("schtasks", "/Delete /TN \"QuickPaths_KeepAlive\" /F");
        RunHidden("schtasks", "/Delete /TN QuickPaths_Watchdog /F");

        // 5. Clean up old files
        TryDelete(Path.Combine(exeDir, "watchdog.ps1"));
        TryDelete(Path.Combine(exeDir, "watchdog.vbs"));
        TryDelete(Path.Combine(exeDir, "wrapper.lock"));

        // 6. Ask about user data
        var r = MessageBox.Show(
            "Delete user data (saved paths and window position)?\n\n"
            + "  paths.json - your saved path list\n"
            + "  config.json - window position\n\n"
            + "Click Yes to delete, No to keep.",
            "QuickPaths Uninstall", MessageBoxButton.YesNo, MessageBoxImage.Question);

        if (r == MessageBoxResult.Yes)
        {
            TryDelete(Path.Combine(exeDir, "paths.json"));
            TryDelete(Path.Combine(exeDir, "config.json"));
        }

        // 7. Done
        MessageBox.Show(
            "QuickPaths uninstalled.\n\n"
            + "  Auto-start: Removed\n"
            + "  Program files: Kept (delete manually if desired)",
            "QuickPaths Uninstall", MessageBoxButton.OK, MessageBoxImage.Information);
    }

    static void KillOldInstances()
    {
        int myPid = Process.GetCurrentProcess().Id;
        // Kill QuickPaths.exe processes (except self)
        foreach (var p in Process.GetProcessesByName("QuickPaths"))
        {
            if (p.Id != myPid)
                try { p.Kill(); } catch { }
        }
        // Kill old VBS wrapper and PS-based instances
        try { RunHidden("wmic", "process where \"name='wscript.exe' and commandline like '%QuickPaths%'\" call terminate"); } catch { }
        try { RunHidden("wmic", "process where \"name='powershell.exe' and commandline like '%QuickPaths.ps1%'\" call terminate"); } catch { }
        Thread.Sleep(500);
    }

    static void RunHidden(string exe, string args)
    {
        try
        {
            var p = Process.Start(new ProcessStartInfo
            {
                FileName = exe,
                Arguments = args,
                WindowStyle = ProcessWindowStyle.Hidden,
                CreateNoWindow = true,
                UseShellExecute = false
            });
            if (p != null) p.WaitForExit(5000);
        }
        catch { }
    }

    static void TryDelete(string path)
    {
        try { if (File.Exists(path)) File.Delete(path); } catch { }
    }
}
