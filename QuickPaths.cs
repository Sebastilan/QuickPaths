// QuickPaths.cs - Desktop floating path quick-copy tool (WinForms)
// Build: build.cmd (csc.exe from .NET Framework 4.x)

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using Microsoft.Win32;

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
        void BindToHandler(IntPtr pbc, ref Guid bhid, ref Guid riid, out IntPtr ppv);
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
            dialog.SetOptions(options | 0x20u | 0x200u | 0x40u);
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
                item.GetDisplayName(0x80058000u, out path);
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

// --- Double-buffered panel for flicker-free GDI+ drawing ---

class BufferedPanel : Panel
{
    public BufferedPanel()
    {
        SetStyle(
            ControlStyles.AllPaintingInWmPaint |
            ControlStyles.UserPaint |
            ControlStyles.OptimizedDoubleBuffer |
            ControlStyles.ResizeRedraw,
            true);
    }
}

// --- Main application ---

class QuickPaths
{
    // --- Constants ---
    const int DOT_W = 88, DOT_H = 38;
    const int ECG_PAD_X = 12, ECG_PAD_Y = 8;
    const int ECG_WAVE_W = 60, ECG_WAVE_H = 22;
    const int ECG_COUNT = 360;
    const string MUTEX_NAME = @"Global\QuickPaths_Singleton";
    const string REG_RUN = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Run";

    // Panel layout
    const int P_PAD = 10, P_PAD_V = 8;
    const int ITEM_H = 32, ITEM_GAP = 4;
    const int BTN_SZ = 24, BTN_GAP = 3;
    const int CLAUDE_H = 30, ADD_H = 30;
    const int SEP_MT = 6, SEP_MB = 4;
    const int PANEL_MIN_W = 160, PANEL_MAX_W = 280;

    // --- State ---
    string dir, dataFile, configFile, logFile;
    List<PathItem> paths = new List<PathItem>();
    bool dialogOpen, expanded, claudeMode, intentionalExit;
    int exitCode = 1;
    int restartCount;
    DateTime startTime;
    Mutex mutex;
    double scale = 1.0;

    // Mouse drag
    bool mouseIsDown, hasDragged;
    Point mouseDownScreen, formPosOnDown;

    // --- UI ---
    Form form;
    BufferedPanel dotPanel;
    Panel expandedPanel;
    ToolTip toolTip;

    // --- Animation ---
    List<double> yValues = new List<double>();
    int breathTick;
    bool dotHovered, flashActive;

    // --- Timers ---
    System.Windows.Forms.Timer breathTimer, flashTimer, displayChangeTimer, healthTimer;

    // --- System event handlers ---
    PowerModeChangedEventHandler onPowerMode;
    SessionSwitchEventHandler onSessionSwitch;
    EventHandler onDisplayChanged;

    // --- Fonts ---
    Font fItem, fAdd, fClaude, fHint;

    // --- Colors ---
    // ECG monitor (alpha-aware, used in GDI+ painting)
    Color C_MonBg, C_MonHover, C_MonBdrBlue, C_MonBdrOrange;
    Color C_WaveBlue, C_WaveOrange, C_WaveGreen;
    Color C_PenBlue, C_PenOrange, C_PenGreen;
    // Panel (opaque, pre-blended for WinForms controls)
    Color C_PanelBg, C_ItemNormal, C_ItemHover;
    Color C_UpNormal, C_UpHover, C_UpText;
    Color C_DelNormal, C_DelHover, C_DelText;
    Color C_AddBg, C_AddHover, C_AddText;
    Color C_Sep;
    // Claude toggle
    Color C_ClaudeOff, C_ClaudeOn, C_ClaudeHoverOff, C_ClaudeHoverOn;
    Color C_ClaudeTextOff, C_ClaudeTextOn;
    // Current dot drawing colors (depend on claudeMode + flash)
    Color curWave, curPen, curBorder, curGlow;

    // =====================================================================
    //  Entry point
    // =====================================================================

    [STAThread]
    static void Main(string[] args)
    {
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);

        if (args.Length > 0)
        {
            string a = args[0].ToLowerInvariant();
            if (a == "--install" || a == "/install") { DoInstall(); return; }
            if (a == "--uninstall" || a == "/uninstall") { DoUninstall(); return; }
        }
        int rc = 0;
        foreach (string arg in args)
            if (arg.StartsWith("--restart-count="))
                int.TryParse(arg.Substring(16), out rc);
        if (rc > 0) Thread.Sleep(2000);
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

        // --- Exception handling (MUST be before any Form creation) ---
        Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);

        // --- Paths ---
        dir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
        dataFile = Path.Combine(dir, "paths.json");
        configFile = Path.Combine(dir, "config.json");
        logFile = Path.Combine(dir, "QuickPaths.log");
        startTime = DateTime.Now;
        Log("QuickPaths starting (WinForms)...");

        // --- Data ---
        LoadPaths();
        AppConfig cfg = LoadConfig();
        scale = cfg.Scale > 0 ? cfg.Scale : 1.0;

        // --- Colors & Fonts ---
        InitColors();
        InitFonts();
        UpdateDotColors();

        // --- Form ---
        form = new Form();
        form.AutoScaleMode = AutoScaleMode.Dpi;
        form.FormBorderStyle = FormBorderStyle.None;
        form.TopMost = true;
        form.ShowInTaskbar = false;
        form.StartPosition = FormStartPosition.Manual;
        form.Location = new Point(cfg.Left, cfg.Top);
        form.BackColor = C_MonBg;
        form.TransparencyKey = Color.FromArgb(1, 0, 1);
        toolTip = new ToolTip();

        // --- Dot panel ---
        dotPanel = new BufferedPanel();
        dotPanel.Dock = DockStyle.Fill;
        dotPanel.Paint += OnDotPaint;
        dotPanel.MouseDown += OnDotMouseDown;
        dotPanel.MouseMove += OnDotMouseMove;
        dotPanel.MouseUp += OnDotMouseUp;
        dotPanel.MouseWheel += OnDotMouseWheel;
        dotPanel.MouseEnter += delegate { dotHovered = true; dotPanel.Invalidate(); };
        dotPanel.MouseLeave += delegate { dotHovered = false; dotPanel.Invalidate(); };
        dotPanel.Cursor = Cursors.Hand;
        form.Controls.Add(dotPanel);

        // Init waveform data
        double yBot = ECG_PAD_Y + ECG_WAVE_H;
        yValues.Clear();
        for (int i = 0; i < ECG_COUNT; i++)
            yValues.Add(yBot);

        ApplyDotSize();
        Log("Window built at (" + cfg.Left + ", " + cfg.Top + ")");

        // --- Timers ---
        breathTimer = new System.Windows.Forms.Timer();
        breathTimer.Interval = 50;
        breathTimer.Tick += OnBreathTick;
        breathTimer.Start();

        flashTimer = new System.Windows.Forms.Timer();
        flashTimer.Interval = 700;
        flashTimer.Tick += delegate
        {
            try
            {
                flashActive = false;
                UpdateDotColors();
                dotPanel.Invalidate();
                flashTimer.Stop();
            }
            catch (Exception ex) { Log("flashTimer error: " + ex.Message); flashTimer.Stop(); }
        };

        displayChangeTimer = new System.Windows.Forms.Timer();
        displayChangeTimer.Interval = 1500;
        displayChangeTimer.Tick += delegate
        {
            try
            {
                displayChangeTimer.Stop();
                Rectangle vs = SystemInformation.VirtualScreen;
                int wl = form.Left, wt = form.Top;
                Log("DisplayChange settled: Window=(" + wl + "," + wt
                    + ") VirtualScreen=(" + vs.X + "," + vs.Y + "," + vs.Width + "," + vs.Height + ")");
                if (wl < vs.Left || wl > vs.Right - 30 || wt < vs.Top || wt > vs.Bottom - 30)
                {
                    Log("Window off-screen, resetting");
                    Rectangle wa = Screen.PrimaryScreen.WorkingArea;
                    form.Left = wa.Right - 60;
                    form.Top = wa.Top + (int)(wa.Height * 0.4);
                    SaveConfig();
                }
                form.TopMost = false;
                form.TopMost = true;
            }
            catch (Exception ex) { Log("displayChangeTimer error: " + ex.Message); }
        };

        healthTimer = new System.Windows.Forms.Timer();
        healthTimer.Interval = 300000;
        healthTimer.Tick += delegate
        {
            try { form.TopMost = false; form.TopMost = true; }
            catch (Exception ex) { Log("healthTimer error: " + ex.Message); }
        };
        healthTimer.Start();

        // --- Click-outside-to-collapse ---
        form.Deactivate += delegate
        {
            try { if (expanded && !dialogOpen) Collapse(); }
            catch (Exception ex) { Log("Deactivate error: " + ex.Message); }
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

        onDisplayChanged = delegate
        {
            try
            {
                Log("Display settings changed (debouncing...)");
                form.BeginInvoke((Action)delegate
                {
                    displayChangeTimer.Stop();
                    displayChangeTimer.Start();
                });
            }
            catch (Exception ex) { Log("DisplaySettingsChanged error: " + ex.Message); }
        };
        SystemEvents.DisplaySettingsChanged += onDisplayChanged;

        // --- Form closed cleanup ---
        form.FormClosed += delegate
        {
            Log("Form closed (intentional=" + intentionalExit + " exitCode=" + exitCode + ")");
            healthTimer.Stop();
            breathTimer.Stop();
            displayChangeTimer.Stop();
            SaveConfig();
            if (intentionalExit) exitCode = 0;
            try { SystemEvents.PowerModeChanged -= onPowerMode; } catch { }
            try { SystemEvents.SessionSwitch -= onSessionSwitch; } catch { }
            try { SystemEvents.DisplaySettingsChanged -= onDisplayChanged; } catch { }
            try { mutex.ReleaseMutex(); } catch { }
            DisposeFonts();
        };

        // --- Exception handling ---
        Application.ThreadException += delegate(object sender, System.Threading.ThreadExceptionEventArgs e)
        {
            LogError("ThreadException", e.Exception);
            LogCrashContext();
            string msg = e.Exception.GetType().FullName + " " + e.Exception.Message;
            bool fatal = msg.Contains("OutOfMemoryException") || msg.Contains("AccessViolation");
            if (fatal)
            {
                Log("FATAL thread error \u2014 closing for restart");
                exitCode = 2;
                try { form.Close(); } catch { }
            }
            else
            {
                Log("Non-fatal thread error \u2014 handled, continuing");
            }
        };

        AppDomain.CurrentDomain.UnhandledException += delegate(object sender, UnhandledExceptionEventArgs e)
        {
            var ex = e.ExceptionObject as Exception;
            if (ex != null) LogError("AppDomain.UnhandledException", ex);
            LogCrashContext();
            exitCode = 3;
            try { form.Close(); } catch { }
        };

        // --- Launch ---
        Log("Launching...");
        try
        {
            Application.Run(form);
        }
        catch (Exception ex)
        {
            LogError("Application.Run", ex);
            LogCrashContext();
            exitCode = 1;
        }
        Log("Exiting with code " + exitCode);

        // Self-restart on crash (up to 5 attempts)
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
    //  Dot sizing & region
    // =====================================================================

    void ApplyDotSize()
    {
        int w = (int)(DOT_W * scale);
        int h = (int)(DOT_H * scale);
        form.ClientSize = new Size(w, h);
        form.Region = null;
        form.BackColor = form.TransparencyKey;
    }

    static Region MakeRoundedRegion(int w, int h, int r)
    {
        if (r < 1) r = 1;
        var path = new GraphicsPath();
        int d = r * 2;
        path.AddArc(0, 0, d, d, 180, 90);
        path.AddArc(w - d, 0, d, d, 270, 90);
        path.AddArc(w - d, h - d, d, d, 0, 90);
        path.AddArc(0, h - d, d, d, 90, 90);
        path.CloseFigure();
        return new Region(path);
    }

    // =====================================================================
    //  Dot painting (GDI+)
    // =====================================================================

    void OnDotPaint(object sender, PaintEventArgs e)
    {
        try
        {
            Graphics g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.Clear(form.TransparencyKey);
            float s = (float)scale;
            int w = dotPanel.Width;
            int h = dotPanel.Height;
            int cr = (int)(6 * s);

            // Background
            Color bg = dotHovered ? C_MonHover : C_MonBg;
            using (var bgPath = MakeRoundedRectPath(0, 0, w, h, cr))
            using (var bgBrush = new SolidBrush(bg))
            {
                g.FillPath(bgBrush, bgPath);
            }

            // Border
            using (var bdrPath = MakeRoundedRectPath(0, 0, w - 1, h - 1, cr))
            using (var bdrPen = new Pen(curBorder, 1f))
            {
                g.DrawPath(bdrPen, bdrPath);
            }

            // Waveform
            float padX = ECG_PAD_X * s;
            float padY = ECG_PAD_Y * s;
            float waveW = ECG_WAVE_W * s;
            float waveH = ECG_WAVE_H * s;

            if (yValues.Count >= 2)
            {
                var pts = new PointF[ECG_COUNT];
                for (int i = 0; i < ECG_COUNT; i++)
                {
                    float x = padX + waveW * (float)i / (ECG_COUNT - 1);
                    float y = (float)(yValues[i] * s);
                    pts[i] = new PointF(x, y);
                }
                using (var wavePen = new Pen(curWave, 1.5f * s))
                {
                    g.DrawLines(wavePen, pts);
                }
            }

            // Pen head position
            float headY = (float)(yValues[yValues.Count - 1] * s);
            float headX = padX + waveW;
            float headR = 3f * s;

            // Glow layers
            float breathVal = (float)GetCurrentBreathValue();

            // Outer glow (30px unscaled)
            float glow2R = 15f * s;
            float glow2Alpha = (float)(0.08 + 0.30 * breathVal);
            DrawGlow(g, headX, headY, glow2R, curGlow, glow2Alpha);

            // Inner glow (16px unscaled)
            float glow1R = 8f * s;
            float glow1Alpha = (float)(0.18 + 0.50 * breathVal);
            DrawGlow(g, headX, headY, glow1R, curGlow, glow1Alpha);

            // Pen head circle
            float shadowAlpha = (float)(0.4 + 0.55 * breathVal);
            // Shadow (slightly larger, semi-transparent)
            using (var shadowBrush = new SolidBrush(Color.FromArgb((int)(shadowAlpha * 120), curGlow)))
            {
                float sr = headR + 4f * s;
                g.FillEllipse(shadowBrush, headX - sr, headY - sr, sr * 2, sr * 2);
            }
            // Head
            using (var headBrush = new SolidBrush(curPen))
            {
                g.FillEllipse(headBrush, headX - headR, headY - headR, headR * 2, headR * 2);
            }
        }
        catch (Exception ex)
        {
            Log("OnDotPaint error: " + ex.Message);
        }
    }

    void DrawGlow(Graphics g, float cx, float cy, float radius, Color color, float alpha)
    {
        if (radius < 1f || alpha < 0.01f) return;
        RectangleF rect = new RectangleF(cx - radius, cy - radius, radius * 2, radius * 2);
        using (var path = new GraphicsPath())
        {
            path.AddEllipse(rect);
            try
            {
                using (var brush = new PathGradientBrush(path))
                {
                    int a = (int)Math.Min(255, alpha * 255);
                    brush.CenterColor = Color.FromArgb(a, color);
                    brush.SurroundColors = new Color[] { Color.FromArgb(0, color) };
                    g.FillPath(brush, path);
                }
            }
            catch { } // PathGradientBrush can fail on degenerate paths
        }
    }

    static GraphicsPath MakeRoundedRectPath(int x, int y, int w, int h, int r)
    {
        var path = new GraphicsPath();
        int d = r * 2;
        if (d > w) d = w;
        if (d > h) d = h;
        path.AddArc(x, y, d, d, 180, 90);
        path.AddArc(x + w - d, y, d, d, 270, 90);
        path.AddArc(x + w - d, y + h - d, d, d, 0, 90);
        path.AddArc(x, y + h - d, d, d, 90, 90);
        path.CloseFigure();
        return path;
    }

    double GetCurrentBreathValue()
    {
        int k = breathTick;
        if (k < 80)
            return (1 - Math.Cos(k / 80.0 * Math.PI)) / 2;
        else if (k < 160)
            return 1.0;
        else if (k < 320)
            return (1 + Math.Cos((k - 160) / 160.0 * Math.PI)) / 2;
        else
            return 0.0;
    }

    // =====================================================================
    //  Breath animation timer
    // =====================================================================

    void OnBreathTick(object sender, EventArgs e)
    {
        try
        {
            breathTick++;
            if (breathTick >= 360) breathTick = 0;

            double v = GetCurrentBreathValue();
            double y = ECG_PAD_Y + ECG_WAVE_H * (1 - v);
            yValues.RemoveAt(0);
            yValues.Add(y);

            if (!expanded)
                dotPanel.Invalidate();
        }
        catch (Exception ex)
        {
            Log("breathTimer error: " + ex.Message);
            breathTimer.Stop();
        }
    }

    // =====================================================================
    //  Dot mouse events
    // =====================================================================

    void OnDotMouseDown(object sender, MouseEventArgs e)
    {
        try
        {
            if (e.Button == MouseButtons.Left)
            {
                mouseIsDown = true;
                mouseDownScreen = Control.MousePosition;
                formPosOnDown = form.Location;
                hasDragged = false;
            }
        }
        catch { }
    }

    void OnDotMouseMove(object sender, MouseEventArgs e)
    {
        try
        {
            if (mouseIsDown && e.Button == MouseButtons.Left)
            {
                Point cur = Control.MousePosition;
                int dx = cur.X - mouseDownScreen.X;
                int dy = cur.Y - mouseDownScreen.Y;
                if (!hasDragged && (Math.Abs(dx) > 5 || Math.Abs(dy) > 5))
                    hasDragged = true;
                if (hasDragged)
                    form.Location = new Point(formPosOnDown.X + dx, formPosOnDown.Y + dy);
            }
        }
        catch (Exception ex) { Log("dot.MouseMove error: " + ex.Message); }
    }

    void OnDotMouseUp(object sender, MouseEventArgs e)
    {
        try
        {
            if (e.Button == MouseButtons.Left)
            {
                if (!hasDragged)
                    Expand();
                else
                    SaveConfig();
                mouseIsDown = false;
            }
            else if (e.Button == MouseButtons.Right)
            {
                var r = MessageBox.Show(
                    "\u9000\u51FA QuickPaths\uFF1F", "\u786E\u8BA4",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (r == DialogResult.Yes)
                {
                    intentionalExit = true;
                    exitCode = 0;
                    form.Close();
                }
            }
        }
        catch (Exception ex) { Log("dot.MouseUp error: " + ex.Message); }
    }

    void OnDotMouseWheel(object sender, MouseEventArgs e)
    {
        try
        {
            double delta = e.Delta > 0 ? 0.1 : -0.1;
            double ns = Math.Round(scale + delta, 1);
            ns = Math.Max(0.5, Math.Min(3.0, ns));
            scale = ns;
            ApplyDotSize();
            dotPanel.Invalidate();
            SaveConfig();
        }
        catch (Exception ex) { Log("dot.MouseWheel error: " + ex.Message); }
    }

    // =====================================================================
    //  UI state transitions
    // =====================================================================

    void Collapse()
    {
        expanded = false;
        form.SuspendLayout();
        if (expandedPanel != null)
        {
            form.Controls.Remove(expandedPanel);
            expandedPanel.Dispose();
            expandedPanel = null;
        }
        dotPanel.Visible = true;
        ApplyDotSize();
        form.ResumeLayout(true);
    }

    void Expand()
    {
        expanded = true;
        form.SuspendLayout();
        dotPanel.Visible = false;
        RebuildList();
        form.ResumeLayout(true);
        form.Activate();
    }

    void FlashDot()
    {
        flashTimer.Stop();
        flashActive = true;
        curWave = C_WaveGreen;
        curPen = C_PenGreen;
        dotPanel.Invalidate();
        flashTimer.Start();
    }

    void ReassertTopmost()
    {
        form.BeginInvoke((Action)delegate
        {
            form.TopMost = false;
            form.TopMost = true;
        });
    }

    void UpdateDotColors()
    {
        if (flashActive) return; // don't override flash
        if (claudeMode)
        {
            curBorder = C_MonBdrOrange;
            curWave = C_WaveOrange;
            curPen = C_PenOrange;
            curGlow = Color.FromArgb(217, 119, 87);
        }
        else
        {
            curBorder = C_MonBdrBlue;
            curWave = C_WaveBlue;
            curPen = C_PenBlue;
            curGlow = Color.FromArgb(106, 155, 204);
        }
    }

    // =====================================================================
    //  Rebuild path list panel
    // =====================================================================

    void PromoteRecentItem(string clickedPath)
    {
        string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
        if (string.Equals(clickedPath, desktopPath, StringComparison.OrdinalIgnoreCase)) return;

        int ci = -1;
        for (int i = 0; i < paths.Count; i++)
            if (paths[i].ItemPath == clickedPath) { ci = i; break; }
        if (ci < 0) return;

        bool desktopAtZero = paths.Count > 0 &&
            string.Equals(paths[0].ItemPath, desktopPath, StringComparison.OrdinalIgnoreCase);
        int insertAt = desktopAtZero ? 1 : 0;
        if (ci == insertAt) return;

        var item = paths[ci];
        paths.RemoveAt(ci);
        if (insertAt > paths.Count) insertAt = paths.Count;
        paths.Insert(insertAt, item);
        SavePaths();
    }

    bool EnsureDesktopFirst()
    {
        string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
        for (int i = 1; i < paths.Count; i++)
        {
            if (string.Equals(paths[i].ItemPath, desktopPath, StringComparison.OrdinalIgnoreCase))
            {
                var desktop = paths[i];
                paths.RemoveAt(i);
                paths.Insert(0, desktop);
                return true;
            }
        }
        return false;
    }

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
        if (EnsureDesktopFirst()) SavePaths();

        // Remove old panel
        if (expandedPanel != null)
        {
            form.Controls.Remove(expandedPanel);
            expandedPanel.Dispose();
        }

        expandedPanel = new Panel();
        expandedPanel.BackColor = C_PanelBg;
        expandedPanel.Dock = DockStyle.Fill;

        // Calculate panel width
        int panelW = PANEL_MIN_W;
        foreach (var p in paths)
        {
            int tw = TextRenderer.MeasureText(p.Name, fItem).Width;
            int rowW = P_PAD + tw + 20 + BTN_GAP + BTN_SZ + BTN_GAP + BTN_SZ + P_PAD;
            if (rowW > panelW) panelW = rowW;
        }
        panelW = Math.Max(PANEL_MIN_W, Math.Min(PANEL_MAX_W, panelW));
        int contentW = panelW - P_PAD * 2;

        int y = P_PAD_V;

        // --- Claude mode toggle ---
        var claudeBtn = new Label();
        claudeBtn.Text = "Claude";
        claudeBtn.Font = fClaude;
        claudeBtn.TextAlign = ContentAlignment.MiddleCenter;
        claudeBtn.Cursor = Cursors.Hand;
        claudeBtn.SetBounds(P_PAD, y, contentW, CLAUDE_H);
        claudeBtn.BackColor = claudeMode ? C_ClaudeOn : C_ClaudeOff;
        claudeBtn.ForeColor = claudeMode ? C_ClaudeTextOn : C_ClaudeTextOff;
        claudeBtn.MouseEnter += delegate(object s, EventArgs ev)
        {
            ((Label)s).BackColor = claudeMode ? C_ClaudeHoverOn : C_ClaudeHoverOff;
        };
        claudeBtn.MouseLeave += delegate(object s, EventArgs ev)
        {
            ((Label)s).BackColor = claudeMode ? C_ClaudeOn : C_ClaudeOff;
        };
        claudeBtn.Click += delegate
        {
            claudeMode = !claudeMode;
            SaveConfig();
            UpdateDotColors();
            RebuildList();
        };
        expandedPanel.Controls.Add(claudeBtn);
        y += CLAUDE_H + ITEM_GAP;

        // --- Empty hint ---
        if (paths.Count == 0)
        {
            var hint = new Label();
            hint.Text = "\u70B9\u51FB + ";
            hint.Font = fHint;
            hint.ForeColor = Color.FromArgb(140, 140, 140);
            hint.BackColor = C_PanelBg;
            hint.TextAlign = ContentAlignment.MiddleCenter;
            hint.SetBounds(P_PAD, y, contentW, 20);
            expandedPanel.Controls.Add(hint);
            y += 20 + ITEM_GAP;
        }

        // --- Path items ---
        for (int idx = 0; idx < paths.Count; idx++)
        {
            PathItem item = paths[idx];
            int btnY = y + (ITEM_H - BTN_SZ) / 2;
            int rightEdge = P_PAD + contentW;

            // Delete button
            var del = new Label();
            del.Text = "\u2212";
            del.Font = fItem;
            del.ForeColor = C_DelText;
            del.BackColor = C_DelNormal;
            del.TextAlign = ContentAlignment.MiddleCenter;
            del.Cursor = Cursors.Hand;
            del.Tag = item.ItemPath;
            del.SetBounds(rightEdge - BTN_SZ, btnY, BTN_SZ, BTN_SZ);
            del.MouseEnter += delegate(object s, EventArgs ev) { ((Label)s).BackColor = C_DelHover; };
            del.MouseLeave += delegate(object s, EventArgs ev) { ((Label)s).BackColor = C_DelNormal; };
            del.Click += delegate(object s, EventArgs ev)
            {
                string target = (string)((Label)s).Tag;
                for (int i = paths.Count - 1; i >= 0; i--)
                {
                    if (paths[i].ItemPath == target) { paths.RemoveAt(i); break; }
                }
                SavePaths();
                RebuildList();
            };
            expandedPanel.Controls.Add(del);
            int nameRight = rightEdge - BTN_SZ - BTN_GAP;

            // Move-up button
            if (idx > 0)
            {
                var up = new Label();
                up.Text = "\u2191";
                up.Font = fItem;
                up.ForeColor = C_UpText;
                up.BackColor = C_UpNormal;
                up.TextAlign = ContentAlignment.MiddleCenter;
                up.Cursor = Cursors.Hand;
                up.Tag = item.ItemPath;
                up.SetBounds(nameRight - BTN_SZ, btnY, BTN_SZ, BTN_SZ);
                up.MouseEnter += delegate(object s, EventArgs ev) { ((Label)s).BackColor = C_UpHover; };
                up.MouseLeave += delegate(object s, EventArgs ev) { ((Label)s).BackColor = C_UpNormal; };
                up.Click += delegate(object s, EventArgs ev)
                {
                    string target = (string)((Label)s).Tag;
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
                expandedPanel.Controls.Add(up);
                nameRight -= BTN_SZ + BTN_GAP;
            }

            // Name button
            var nameBtn = new Label();
            nameBtn.Text = item.Name;
            nameBtn.Font = fItem;
            nameBtn.ForeColor = Color.White;
            nameBtn.BackColor = C_ItemNormal;
            nameBtn.TextAlign = ContentAlignment.MiddleLeft;
            nameBtn.Cursor = Cursors.Hand;
            nameBtn.Tag = item.ItemPath;
            nameBtn.AutoEllipsis = true;
            nameBtn.Padding = new Padding(10, 0, 10, 0);
            nameBtn.SetBounds(P_PAD, y, nameRight - P_PAD, ITEM_H);
            toolTip.SetToolTip(nameBtn, item.ItemPath);
            nameBtn.MouseEnter += delegate(object s, EventArgs ev) { ((Label)s).BackColor = C_ItemHover; };
            nameBtn.MouseLeave += delegate(object s, EventArgs ev) { ((Label)s).BackColor = C_ItemNormal; };
            nameBtn.Click += delegate(object s, EventArgs ev)
            {
                string p = (string)((Label)s).Tag;
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
                PromoteRecentItem(p);
                Collapse();
                FlashDot();
            };
            expandedPanel.Controls.Add(nameBtn);
            y += ITEM_H + ITEM_GAP;
        }

        // --- Separator ---
        if (paths.Count > 0)
        {
            y += SEP_MT - ITEM_GAP;
            var sep = new Label();
            sep.BackColor = C_Sep;
            sep.SetBounds(P_PAD + 4, y, contentW - 8, 1);
            expandedPanel.Controls.Add(sep);
            y += 1 + SEP_MB;
        }

        // --- Add button ---
        var addBtn = new Label();
        addBtn.Text = "+";
        addBtn.Font = fAdd;
        addBtn.ForeColor = C_AddText;
        addBtn.BackColor = C_AddBg;
        addBtn.TextAlign = ContentAlignment.MiddleCenter;
        addBtn.Cursor = Cursors.Hand;
        addBtn.SetBounds(P_PAD, y, contentW, ADD_H);
        addBtn.MouseEnter += delegate(object s, EventArgs ev) { ((Label)s).BackColor = C_AddHover; };
        addBtn.MouseLeave += delegate(object s, EventArgs ev) { ((Label)s).BackColor = C_AddBg; };
        addBtn.Click += delegate
        {
            dialogOpen = true;
            try
            {
                form.TopMost = false;
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
                form.TopMost = true;
                form.Activate();
                if (expanded) RebuildList();
            }
        };
        expandedPanel.Controls.Add(addBtn);
        y += ADD_H + P_PAD_V;

        // --- Set form size ---
        form.ClientSize = new Size(panelW, y);
        form.Region = MakeRoundedRegion(panelW, y, 12);
        form.BackColor = C_PanelBg;

        // Keep panel on screen
        Screen scr = Screen.FromPoint(form.Location);
        Rectangle wa = scr.WorkingArea;
        if (form.Right > wa.Right)
            form.Left = wa.Right - panelW;
        if (form.Bottom > wa.Bottom)
            form.Top = wa.Bottom - y;
        if (form.Left < wa.Left)
            form.Left = wa.Left;
        if (form.Top < wa.Top)
            form.Top = wa.Top;

        form.Controls.Add(expandedPanel);
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

        Rectangle vs = SystemInformation.VirtualScreen;
        Log("LoadConfig: saved=(" + (cfg != null ? cfg.Left.ToString() : "null") + ","
            + (cfg != null ? cfg.Top.ToString() : "null") + ") VirtualScreen=("
            + vs.X + "," + vs.Y + "," + vs.Width + "," + vs.Height + ")");

        if (cfg == null || cfg.Left < vs.Left || cfg.Left > vs.Right - 30
                        || cfg.Top < vs.Top || cfg.Top > vs.Bottom - 30)
        {
            Rectangle wa = Screen.PrimaryScreen.WorkingArea;
            cfg = new AppConfig
            {
                Left = wa.Right - 60,
                Top = wa.Top + (int)(wa.Height * 0.4),
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
            int l = form.Left, t = form.Top;
            Log("SaveConfig: (" + l + "," + t + ")");
            var cfg = new AppConfig
            {
                Left = l, Top = t,
                ClaudeMode = claudeMode,
                Scale = scale
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
            Rectangle vs = SystemInformation.VirtualScreen;
            Log("  CrashContext: VirtualScreen=(" + vs.X + "," + vs.Y + "," + vs.Width + "," + vs.Height + ")");
            if (form != null)
                Log("  CrashContext: WindowPos=(" + form.Left + "," + form.Top + ") Topmost=" + form.TopMost);
            double mem = Math.Round(GC.GetTotalMemory(false) / 1048576.0, 1);
            Log("  CrashContext: ManagedMemory=" + mem + "MB");
        }
        catch { }
    }

    // =====================================================================
    //  Color & Font initialization
    // =====================================================================

    static Color Blend(Color bg, int a, int r, int g, int b)
    {
        float alpha = a / 255f;
        return Color.FromArgb(255,
            (int)(r * alpha + bg.R * (1 - alpha)),
            (int)(g * alpha + bg.G * (1 - alpha)),
            (int)(b * alpha + bg.B * (1 - alpha)));
    }

    void InitColors()
    {
        // ECG monitor (used in GDI+ painting with alpha support)
        C_MonBg        = Color.FromArgb(18, 20, 26);
        C_MonHover     = Color.FromArgb(28, 30, 38);
        C_MonBdrBlue   = Color.FromArgb(80, 106, 155, 204);
        C_MonBdrOrange = Color.FromArgb(80, 217, 119, 87);
        C_WaveBlue     = Color.FromArgb(190, 106, 155, 204);
        C_WaveOrange   = Color.FromArgb(190, 217, 119, 87);
        C_WaveGreen    = Color.FromArgb(210, 120, 140, 93);
        C_PenBlue      = Color.FromArgb(155, 200, 240);
        C_PenOrange    = Color.FromArgb(245, 160, 125);
        C_PenGreen     = Color.FromArgb(165, 190, 140);

        // Panel (pre-blended against panel bg for WinForms controls)
        Color bg = Color.FromArgb(32, 34, 38);
        C_PanelBg    = bg;
        C_ItemNormal = Blend(bg, 22, 255, 255, 255);
        C_ItemHover  = Blend(bg, 50, 255, 255, 255);
        C_UpNormal   = bg;
        C_UpHover    = Blend(bg, 35, 255, 255, 255);
        C_UpText     = Blend(bg, 100, 180, 180, 180);
        C_DelNormal  = bg;
        C_DelHover   = Blend(bg, 40, 255, 80, 80);
        C_DelText    = Blend(bg, 120, 255, 100, 100);
        C_AddBg      = bg;
        C_AddHover   = Blend(bg, 22, 255, 255, 255);
        C_AddText    = Blend(bg, 110, 180, 180, 180);
        C_Sep        = Blend(bg, 18, 255, 255, 255);

        // Claude toggle
        C_ClaudeOff      = Blend(bg, 15, 250, 249, 245);
        C_ClaudeOn       = Blend(bg, 130, 217, 119, 87);
        C_ClaudeHoverOff = Blend(bg, 28, 250, 249, 245);
        C_ClaudeHoverOn  = Blend(bg, 150, 217, 119, 87);
        C_ClaudeTextOff  = Blend(bg, 75, 176, 174, 165);
        C_ClaudeTextOn   = Color.FromArgb(250, 249, 245);
    }

    void InitFonts()
    {
        fItem   = new Font("Segoe UI", 10f);
        fAdd    = new Font("Segoe UI", 12f);
        fClaude = new Font("Segoe UI", 10f);
        fHint   = new Font("Segoe UI", 9f);
    }

    void DisposeFonts()
    {
        if (fItem != null) fItem.Dispose();
        if (fAdd != null) fAdd.Dispose();
        if (fClaude != null) fClaude.Dispose();
        if (fHint != null) fHint.Dispose();
    }

    // =====================================================================
    //  Install / Uninstall
    // =====================================================================

    static void DoInstall()
    {
        string exePath = Assembly.GetExecutingAssembly().Location;
        string exeDir = Path.GetDirectoryName(exePath);

        KillOldInstances();
        RunHidden("schtasks", "/Delete /TN QuickPaths_Watchdog /F");

        TryDelete(Path.Combine(exeDir, "watchdog.ps1"));
        TryDelete(Path.Combine(exeDir, "watchdog.vbs"));
        TryDelete(Path.Combine(exeDir, "wrapper.lock"));

        string startupDir = Environment.GetFolderPath(Environment.SpecialFolder.Startup);
        TryDelete(Path.Combine(startupDir, "QuickPaths.vbs"));

        try
        {
            using (RegistryKey key = Registry.CurrentUser.OpenSubKey(REG_RUN, true))
            {
                if (key != null)
                    key.SetValue("QuickPaths", "\"" + exePath + "\"");
            }
        }
        catch { }

        RunHidden("schtasks", "/Create /TN \"QuickPaths_KeepAlive\" /SC MINUTE /MO 5 /TR \"\\\"" + exePath + "\\\"\" /F");

        try
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = exePath,
                UseShellExecute = true
            });
        }
        catch { }

        MessageBox.Show(
            "QuickPaths installed successfully!\n\n"
            + "  Auto-start: Enabled (registry + keepalive task)\n"
            + "  Location: " + exeDir + "\n\n"
            + "The floating dot should appear on your desktop.",
            "QuickPaths Install", MessageBoxButtons.OK, MessageBoxIcon.Information);
    }

    static void DoUninstall()
    {
        string exePath = Assembly.GetExecutingAssembly().Location;
        string exeDir = Path.GetDirectoryName(exePath);

        KillOldInstances();

        try
        {
            using (RegistryKey key = Registry.CurrentUser.OpenSubKey(REG_RUN, true))
            {
                if (key != null) key.DeleteValue("QuickPaths", false);
            }
        }
        catch { }

        string startupDir = Environment.GetFolderPath(Environment.SpecialFolder.Startup);
        TryDelete(Path.Combine(startupDir, "QuickPaths.vbs"));

        RunHidden("schtasks", "/Delete /TN \"QuickPaths_KeepAlive\" /F");
        RunHidden("schtasks", "/Delete /TN QuickPaths_Watchdog /F");

        TryDelete(Path.Combine(exeDir, "watchdog.ps1"));
        TryDelete(Path.Combine(exeDir, "watchdog.vbs"));
        TryDelete(Path.Combine(exeDir, "wrapper.lock"));

        var r = MessageBox.Show(
            "Delete user data (saved paths and window position)?\n\n"
            + "  paths.json - your saved path list\n"
            + "  config.json - window position\n\n"
            + "Click Yes to delete, No to keep.",
            "QuickPaths Uninstall", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

        if (r == DialogResult.Yes)
        {
            TryDelete(Path.Combine(exeDir, "paths.json"));
            TryDelete(Path.Combine(exeDir, "config.json"));
        }

        MessageBox.Show(
            "QuickPaths uninstalled.\n\n"
            + "  Auto-start: Removed\n"
            + "  Program files: Kept (delete manually if desired)",
            "QuickPaths Uninstall", MessageBoxButtons.OK, MessageBoxIcon.Information);
    }

    static void KillOldInstances()
    {
        int myPid = Process.GetCurrentProcess().Id;
        foreach (var p in Process.GetProcessesByName("QuickPaths"))
        {
            if (p.Id != myPid)
                try { p.Kill(); } catch { }
        }
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
