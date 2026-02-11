<#
    QuickPaths - Desktop floating path quick-copy tool
    Stability: auto-start, sleep/wake recovery, log rotation, singleton
#>

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms

# Multi-folder picker via IFileOpenDialog COM (supports Ctrl+click multi-select)
Add-Type -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

public class FolderMultiPicker {
    [ComImport, Guid("DC1C5A9C-E88A-4DDE-A5A1-60F82A20AEF7")]
    private class FileOpenDialogCOM { }

    [ComImport, Guid("D57C7288-D4AD-4768-BE02-9D969532D960")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IFileOpenDialog {
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
    private interface IShellItem {
        void BindToHandler(IntPtr pbc, ref Guid bhid, ref Guid riid, out IntPtr ppv);
        void GetParent(out IShellItem ppsi);
        void GetDisplayName(uint sigdnName, [MarshalAs(UnmanagedType.LPWStr)] out string ppszName);
        void GetAttributes(uint sfgaoMask, out uint psfgaoAttribs);
        void Compare(IShellItem psi, uint hint, out int piOrder);
    }

    [ComImport, Guid("B63EA76D-1F85-456F-A19C-48159EFA858B")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IShellItemArray {
        void BindToHandler(IntPtr pbc, ref Guid bhid, ref Guid riid, out IntPtr ppvOut);
        void GetPropertyStore(int flags, ref Guid riid, out IntPtr ppv);
        void GetPropertyDescriptionList(IntPtr keyType, ref Guid riid, out IntPtr ppv);
        void GetAttributes(int attribFlags, uint sfgaoMask, out uint psfgaoAttribs);
        void GetCount(out uint pdwNumItems);
        void GetItemAt(uint dwIndex, out IShellItem ppsi);
        void EnumItems(out IntPtr ppenumShellItems);
    }

    public static string[] ShowDialog(string title) {
        IFileOpenDialog dialog = (IFileOpenDialog)new FileOpenDialogCOM();
        try {
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
            List<string> paths = new List<string>();
            for (uint i = 0; i < count; i++) {
                IShellItem item;
                results.GetItemAt(i, out item);
                string path;
                item.GetDisplayName(0x80058000u, out path);
                if (!string.IsNullOrEmpty(path))
                    paths.Add(path);
            }
            return paths.ToArray();
        } finally {
            Marshal.ReleaseComObject(dialog);
        }
    }
}
'@

# Singleton (handle abandoned mutex from killed process)
$script:mutex = New-Object System.Threading.Mutex($false, 'Global\QuickPaths_Singleton')
$script:ownsMutex = $false
try {
    $script:ownsMutex = $script:mutex.WaitOne(0)
} catch [System.Threading.AbandonedMutexException] {
    $script:ownsMutex = $true
}
if (-not $script:ownsMutex) { exit }

# Paths
$script:Dir = $PSScriptRoot
if (-not $script:Dir) { $script:Dir = Split-Path -Parent $MyInvocation.MyCommand.Path }
$script:DataFile   = Join-Path $script:Dir 'paths.json'
$script:ConfigFile = Join-Path $script:Dir 'config.json'
$script:LogFile    = Join-Path $script:Dir 'QuickPaths.log'

# Log with rotation (max 512KB)
function script:Log([string]$msg) {
    try {
        $fi = New-Object System.IO.FileInfo($script:LogFile)
        if ($fi.Exists -and $fi.Length -gt 512KB) {
            $bak = $script:LogFile + '.old'
            if (Test-Path $bak) { Remove-Item $bak -Force }
            Move-Item $script:LogFile $bak -Force
        }
        $ts = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        [System.IO.File]::AppendAllText($script:LogFile, "[$ts] $msg`r`n", [System.Text.Encoding]::UTF8)
    } catch {}
}
Log 'QuickPaths starting...'

# State
$script:paths      = [System.Collections.ArrayList]::new()
$script:dialogOpen = $false
$script:expanded   = $false
$script:dotClickPt = $null

function script:LoadPaths {
    $script:paths.Clear()
    try {
        if (Test-Path $script:DataFile) {
            $json = [System.IO.File]::ReadAllText($script:DataFile, [System.Text.Encoding]::UTF8)
            $json = $json.TrimStart([char]0xFEFF)
            if ($json -and $json.Trim().Length -gt 2) {
                $items = ConvertFrom-Json $json
                foreach ($item in $items) {
                    [void]$script:paths.Add([PSCustomObject]@{
                        name = $item.name
                        path = $item.path
                    })
                }
            }
        }
    } catch {
        Log "LoadPaths error: $_"
        $script:paths.Clear()
    }
    Log "Loaded $($script:paths.Count) paths"
}

function script:SavePaths {
    try {
        $noBom = New-Object System.Text.UTF8Encoding($false)
        if ($script:paths.Count -eq 0) {
            $json = '[]'
        } else {
            $json = @($script:paths) | ConvertTo-Json -Depth 2
            # PS 5.1: ConvertTo-Json unwraps single-element arrays, force array
            if ($script:paths.Count -eq 1) {
                $json = '[' + $json + ']'
            }
        }
        # Safe write: temp file then replace (prevents corruption on crash)
        $tmp = $script:DataFile + '.tmp'
        [System.IO.File]::WriteAllText($tmp, $json, $noBom)
        if (Test-Path $script:DataFile) {
            [System.IO.File]::Replace($tmp, $script:DataFile, $null)
        } else {
            [System.IO.File]::Move($tmp, $script:DataFile)
        }
    } catch {
        Log "SavePaths error: $_"
        $tmp = $script:DataFile + '.tmp'
        if (Test-Path $tmp) { Remove-Item $tmp -Force -ErrorAction SilentlyContinue }
    }
}

function script:LoadConfig {
    $pos = $null
    try {
        if (Test-Path $script:ConfigFile) {
            $raw = [System.IO.File]::ReadAllText($script:ConfigFile, [System.Text.Encoding]::UTF8)
            $raw = $raw.TrimStart([char]0xFEFF)
            $pos = ConvertFrom-Json $raw
        }
    } catch {}
    # VirtualScreen covers ALL monitors (supports negative coords for extended displays)
    $vl = [System.Windows.SystemParameters]::VirtualScreenLeft
    $vt = [System.Windows.SystemParameters]::VirtualScreenTop
    $vw = [System.Windows.SystemParameters]::VirtualScreenWidth
    $vh = [System.Windows.SystemParameters]::VirtualScreenHeight
    Log "LoadConfig: saved=($($pos.left),$($pos.top)) VirtualScreen=($vl,$vt,$vw,$vh)"
    if (-not $pos -or
        $pos.left -lt $vl -or $pos.left -gt ($vl + $vw - 30) -or
        $pos.top  -lt $vt -or $pos.top  -gt ($vt + $vh - 30)) {
        $wa = [System.Windows.SystemParameters]::WorkArea
        $pos = [PSCustomObject]@{ left = [int]($wa.Width - 60); top = [int]($wa.Height * 0.4) }
        Log "LoadConfig: out of bounds, using default ($($pos.left),$($pos.top))"
    }
    return $pos
}

function script:SaveConfig {
    try {
        $l = [int]$script:win.Left
        $t = [int]$script:win.Top
        Log "SaveConfig: ($l,$t)"
        $json = @{ left = $l; top = $t } | ConvertTo-Json
        $noBom = New-Object System.Text.UTF8Encoding($false)
        [System.IO.File]::WriteAllText($script:ConfigFile, $json, $noBom)
    } catch {}
}

LoadPaths

# Colors
function script:Brush([byte]$a, [byte]$r, [byte]$g, [byte]$b) {
    [System.Windows.Media.SolidColorBrush]::new(
        [System.Windows.Media.Color]::FromArgb($a, $r, $g, $b))
}

# Dot: bright cornflower blue - pairs with dark status bar accents
$script:C_DotNormal  = Brush 215 105 145 235
$script:C_DotHover   = Brush 245 125 165 255
$script:C_DotFlash   = Brush 225 80 185 120
# Panel
$script:C_PanelBg    = Brush 245 32 34 38
$script:C_ItemNormal = Brush 22 255 255 255
$script:C_ItemHover  = Brush 50 255 255 255
$script:C_UpNormal   = [System.Windows.Media.Brushes]::Transparent
$script:C_UpHover    = Brush 35 255 255 255
$script:C_UpText     = Brush 100 180 180 180
$script:C_DelNormal  = [System.Windows.Media.Brushes]::Transparent
$script:C_DelHover   = Brush 40 255 80 80
$script:C_DelText    = Brush 120 255 100 100
$script:C_AddBg      = [System.Windows.Media.Brushes]::Transparent
$script:C_AddHover   = Brush 22 255 255 255
$script:C_AddText    = Brush 110 180 180 180
$script:C_Sep        = Brush 18 255 255 255
$script:C_White      = [System.Windows.Media.Brushes]::White
$script:C_Transp     = [System.Windows.Media.Brushes]::Transparent

# Window
$script:win = New-Object System.Windows.Window
$script:win.WindowStyle        = 'None'
$script:win.AllowsTransparency = $true
$script:win.Background         = [System.Windows.Media.Brushes]::Transparent
$script:win.Topmost            = $true
$script:win.ShowInTaskbar      = $false
$script:win.SizeToContent      = 'WidthAndHeight'
$script:win.ResizeMode         = 'NoResize'

$cfg = LoadConfig
$script:win.Left = $cfg.left
$script:win.Top  = $cfg.top

$root = New-Object System.Windows.Controls.Grid

# Dot (30px dark circle)
$script:dot = New-Object System.Windows.Controls.Border
$script:dot.Width        = 30
$script:dot.Height       = 30
$script:dot.CornerRadius = [System.Windows.CornerRadius]::new(15)
$script:dot.Background   = $script:C_DotNormal
$script:dot.Cursor       = [System.Windows.Input.Cursors]::Hand
$script:dot.Effect = New-Object System.Windows.Media.Effects.DropShadowEffect -Property @{
    BlurRadius  = 10
    ShadowDepth = 1
    Opacity     = 0.30
    Color       = [System.Windows.Media.Colors]::Black
}

# Panel
$script:panel = New-Object System.Windows.Controls.Border
$script:panel.CornerRadius = [System.Windows.CornerRadius]::new(12)
$script:panel.Background   = $script:C_PanelBg
$script:panel.Padding      = [System.Windows.Thickness]::new(10,8,10,8)
$script:panel.MinWidth     = 160
$script:panel.MaxWidth     = 280
$script:panel.Visibility   = 'Collapsed'
$script:panel.Effect = New-Object System.Windows.Media.Effects.DropShadowEffect -Property @{
    BlurRadius  = 20
    ShadowDepth = 2
    Opacity     = 0.40
    Color       = [System.Windows.Media.Colors]::Black
}

$script:listPanel = New-Object System.Windows.Controls.StackPanel
$script:panel.Child = $script:listPanel

$root.Children.Add($script:dot)
$root.Children.Add($script:panel)
$script:win.Content = $root
Log "Window built at ($($script:win.Left), $($script:win.Top))"

# UI Logic
function script:Collapse {
    $script:expanded = $false
    $script:panel.Visibility = 'Collapsed'
    $script:dot.Visibility = 'Visible'
}

function script:Expand {
    $script:expanded = $true
    $script:dot.Visibility = 'Collapsed'
    RebuildList
    $script:panel.Visibility = 'Visible'
}

$script:_ft = New-Object System.Windows.Threading.DispatcherTimer
$script:_ft.Interval = [TimeSpan]::FromMilliseconds(700)
$script:_ft.Add_Tick({
    $script:dot.Background = $script:C_DotNormal
    $script:_ft.Stop()
})
function script:FlashDot {
    $script:_ft.Stop()
    $script:dot.Background = $script:C_DotFlash
    $script:_ft.Start()
}

function script:RebuildList {
    $lp = $script:listPanel
    $lp.Children.Clear()

    if ($script:paths.Count -eq 0) {
        $hint = New-Object System.Windows.Controls.TextBlock
        $hint.Text                = [char]0x70B9 + [char]0x51FB + ' + '
        $hint.Foreground          = Brush 80 140 140 140
        $hint.FontSize            = 12
        $hint.HorizontalAlignment = 'Center'
        $hint.Margin              = [System.Windows.Thickness]::new(8,4,8,2)
        $lp.Children.Add($hint)
    }

    for ($idx = 0; $idx -lt $script:paths.Count; $idx++) {
        $item = $script:paths[$idx]
        $row = New-Object System.Windows.Controls.DockPanel
        $row.Margin = [System.Windows.Thickness]::new(0,2,0,2)

        # Delete button (rightmost)
        $del = New-Object System.Windows.Controls.Border
        $del.Width        = 24
        $del.Height       = 24
        $del.CornerRadius = [System.Windows.CornerRadius]::new(12)
        $del.Background   = $script:C_DelNormal
        $del.Cursor       = [System.Windows.Input.Cursors]::Hand
        $del.Tag          = $item.path
        $del.Margin       = [System.Windows.Thickness]::new(3,0,0,0)
        [System.Windows.Controls.DockPanel]::SetDock($del, 'Right')

        $delText = New-Object System.Windows.Controls.TextBlock
        $delText.Text                = [string][char]0x2212
        $delText.Foreground          = $script:C_DelText
        $delText.FontSize            = 14
        $delText.FontWeight          = 'Medium'
        $delText.HorizontalAlignment = 'Center'
        $delText.VerticalAlignment   = 'Center'
        $delText.IsHitTestVisible    = $false
        $del.Child = $delText

        $del.Add_MouseEnter({ param($s) $s.Background = $script:C_DelHover })
        $del.Add_MouseLeave({ param($s) $s.Background = $script:C_DelNormal })
        $del.Add_MouseLeftButtonDown({
            param($s)
            $target = $s.Tag
            for ($i = $script:paths.Count - 1; $i -ge 0; $i--) {
                if ($script:paths[$i].path -eq $target) {
                    $script:paths.RemoveAt($i); break
                }
            }
            SavePaths; RebuildList
        })
        $row.Children.Add($del)

        # Move-up button (to the left of delete)
        if ($idx -gt 0) {
            $up = New-Object System.Windows.Controls.Border
            $up.Width        = 24
            $up.Height       = 24
            $up.CornerRadius = [System.Windows.CornerRadius]::new(12)
            $up.Background   = $script:C_UpNormal
            $up.Cursor       = [System.Windows.Input.Cursors]::Hand
            $up.Tag          = $item.path
            $up.Margin       = [System.Windows.Thickness]::new(3,0,0,0)
            [System.Windows.Controls.DockPanel]::SetDock($up, 'Right')

            $upText = New-Object System.Windows.Controls.TextBlock
            $upText.Text                = [string][char]0x2191
            $upText.Foreground          = $script:C_UpText
            $upText.FontSize            = 13
            $upText.HorizontalAlignment = 'Center'
            $upText.VerticalAlignment   = 'Center'
            $upText.IsHitTestVisible    = $false
            $up.Child = $upText

            $up.Add_MouseEnter({ param($s) $s.Background = $script:C_UpHover })
            $up.Add_MouseLeave({ param($s) $s.Background = $script:C_UpNormal })
            $up.Add_MouseLeftButtonDown({
                param($s)
                $target = $s.Tag
                for ($i = 1; $i -lt $script:paths.Count; $i++) {
                    if ($script:paths[$i].path -eq $target) {
                        $tmp = $script:paths[$i - 1]
                        $script:paths[$i - 1] = $script:paths[$i]
                        $script:paths[$i] = $tmp
                        break
                    }
                }
                SavePaths; RebuildList
            })
            $row.Children.Add($up)
        }

        # Name button (fill remaining)
        $nameBtn = New-Object System.Windows.Controls.Border
        $nameBtn.Padding      = [System.Windows.Thickness]::new(10,7,10,7)
        $nameBtn.CornerRadius = [System.Windows.CornerRadius]::new(7)
        $nameBtn.Background   = $script:C_ItemNormal
        $nameBtn.Cursor       = [System.Windows.Input.Cursors]::Hand
        $nameBtn.Tag          = $item.path
        $nameBtn.ToolTip      = $item.path

        $nameText = New-Object System.Windows.Controls.TextBlock
        $nameText.Text             = $item.name
        $nameText.Foreground       = $script:C_White
        $nameText.FontSize         = 13
        $nameText.TextTrimming     = 'CharacterEllipsis'
        $nameText.IsHitTestVisible = $false
        $nameBtn.Child = $nameText

        $nameBtn.Add_MouseEnter({ param($s) $s.Background = $script:C_ItemHover })
        $nameBtn.Add_MouseLeave({ param($s) $s.Background = $script:C_ItemNormal })
        $nameBtn.Add_MouseLeftButtonDown({
            param($s)
            try { [System.Windows.Clipboard]::SetText($s.Tag) }
            catch { Log "Clipboard error: $_"; return }
            Collapse; FlashDot
        })
        $row.Children.Add($nameBtn)

        $lp.Children.Add($row)
    }

    # Separator
    if ($script:paths.Count -gt 0) {
        $sep = New-Object System.Windows.Controls.Border
        $sep.Height     = 1
        $sep.Background = $script:C_Sep
        $sep.Margin     = [System.Windows.Thickness]::new(4,6,4,4)
        $lp.Children.Add($sep)
    }

    # Add button
    $addBtn = New-Object System.Windows.Controls.Border
    $addBtn.Padding      = [System.Windows.Thickness]::new(10,5,10,5)
    $addBtn.CornerRadius = [System.Windows.CornerRadius]::new(7)
    $addBtn.Background   = $script:C_AddBg
    $addBtn.Cursor       = [System.Windows.Input.Cursors]::Hand

    $addText = New-Object System.Windows.Controls.TextBlock
    $addText.Text                = '+'
    $addText.Foreground          = $script:C_AddText
    $addText.FontSize            = 16
    $addText.HorizontalAlignment = 'Center'
    $addText.IsHitTestVisible    = $false
    $addBtn.Child = $addText

    $addBtn.Add_MouseEnter({ param($s) $s.Background = $script:C_AddHover })
    $addBtn.Add_MouseLeave({ param($s) $s.Background = $script:C_AddBg })
    $addBtn.Add_MouseLeftButtonDown({
        $script:dialogOpen = $true
        try {
            $script:win.Topmost = $false
            $dlgTitle = [char]0x9009 + [char]0x62E9 + [char]0x6587 + [char]0x4EF6 + [char]0x5939 + '  (Ctrl+' + [char]0x70B9 + [char]0x51FB + [char]0x591A + [char]0x9009 + ')'
            $selected = [FolderMultiPicker]::ShowDialog($dlgTitle)
            $added = 0
            foreach ($fp in $selected) {
                $fn = Split-Path $fp -Leaf
                $dup = $false
                foreach ($p in $script:paths) { if ($p.path -eq $fp) { $dup = $true; break } }
                if (-not $dup) {
                    [void]$script:paths.Add([PSCustomObject]@{ name = $fn; path = $fp })
                    $added++
                }
            }
            if ($added -gt 0) {
                Log "Added $added folder(s)"
                SavePaths
            }
        } catch {
            Log "Add error: $_"
        } finally {
            $script:dialogOpen = $false
            $script:win.Topmost = $true
            $script:win.Activate()
            if ($script:expanded) { RebuildList }
        }
    })
    $lp.Children.Add($addBtn)
}

# Dot interaction: click to expand, drag to move
$script:dot.Add_MouseEnter({ $script:dot.Background = $script:C_DotHover })
$script:dot.Add_MouseLeave({ $script:dot.Background = $script:C_DotNormal })

$script:dot.Add_MouseLeftButtonDown({
    param($s, $e)
    $script:dotClickPt = [System.Windows.Forms.Cursor]::Position
})

$script:dot.Add_MouseMove({
    param($s, $e)
    if ($e.LeftButton -eq 'Pressed' -and $script:dotClickPt) {
        $cur = [System.Windows.Forms.Cursor]::Position
        if ([Math]::Abs($cur.X - $script:dotClickPt.X) -gt 5 -or
            [Math]::Abs($cur.Y - $script:dotClickPt.Y) -gt 5) {
            $script:dotClickPt = $null
            $script:win.DragMove()
            SaveConfig
        }
    }
})

$script:dot.Add_MouseLeftButtonUp({
    param($s, $e)
    if ($script:dotClickPt) { Expand }
    $script:dotClickPt = $null
})

# Right-click dot to exit
$script:dot.Add_MouseRightButtonUp({
    $msgTitle = [char]0x786E + [char]0x8BA4
    $msgBody  = [char]0x9000 + [char]0x51FA + ' QuickPaths' + [char]0xFF1F
    $r = [System.Windows.MessageBox]::Show($msgBody, $msgTitle, 'YesNo', 'Question')
    if ($r -eq 'Yes') { $script:win.Close() }
})

# Click outside to collapse
$script:win.Add_Deactivated({
    if ($script:expanded -and -not $script:dialogOpen) { Collapse }
})

# Re-assert Topmost helper
function script:ReassertTopmost {
    $script:win.Dispatcher.Invoke({
        $script:win.Topmost = $false
        $script:win.Topmost = $true
    })
}

# Sleep/wake recovery
[Microsoft.Win32.SystemEvents]::Add_PowerModeChanged({
    param($sender, $e)
    if ($e.Mode -eq [Microsoft.Win32.PowerModes]::Resume) {
        Log 'System resumed from sleep'
        ReassertTopmost
    }
})

# Lock/unlock recovery
[Microsoft.Win32.SystemEvents]::Add_SessionSwitch({
    param($sender, $e)
    if ($e.Reason -eq [Microsoft.Win32.SessionSwitchReason]::SessionUnlock) {
        Log 'Session unlocked'
        ReassertTopmost
    }
})

# Monitor plug/unplug: re-validate position
[Microsoft.Win32.SystemEvents]::Add_DisplaySettingsChanged({
    Log 'Display settings changed'
    $script:win.Dispatcher.Invoke({
        $vl = [System.Windows.SystemParameters]::VirtualScreenLeft
        $vt = [System.Windows.SystemParameters]::VirtualScreenTop
        $vw = [System.Windows.SystemParameters]::VirtualScreenWidth
        $vh = [System.Windows.SystemParameters]::VirtualScreenHeight
        $wl = $script:win.Left
        $wt = $script:win.Top
        if ($wl -lt $vl -or $wl -gt ($vl + $vw - 30) -or
            $wt -lt $vt -or $wt -gt ($vt + $vh - 30)) {
            Log "Window off-screen ($wl,$wt), resetting"
            $wa = [System.Windows.SystemParameters]::WorkArea
            $script:win.Left = $wa.Width - 60
            $script:win.Top  = $wa.Height * 0.4
            SaveConfig
        }
        $script:win.Topmost = $false
        $script:win.Topmost = $true
    })
})

# Periodic health check: re-assert Topmost every 5 minutes
$script:healthTimer = New-Object System.Windows.Threading.DispatcherTimer
$script:healthTimer.Interval = [TimeSpan]::FromMinutes(5)
$script:healthTimer.Add_Tick({
    $script:win.Topmost = $false
    $script:win.Topmost = $true
})
$script:healthTimer.Start()

# Cleanup
$script:win.Add_Closed({
    Log 'Window closed'
    $script:healthTimer.Stop()
    SaveConfig
    try { [Microsoft.Win32.SystemEvents]::Remove_PowerModeChanged($null) } catch {}
    try { [Microsoft.Win32.SystemEvents]::Remove_SessionSwitch($null) } catch {}
    try { [Microsoft.Win32.SystemEvents]::Remove_DisplaySettingsChanged($null) } catch {}
    try { $script:mutex.ReleaseMutex() } catch {}
})

# Launch
Log 'Launching...'
try {
    $app = New-Object System.Windows.Application
    $app.Run($script:win)
} catch {
    Log "FATAL: $($_.Exception.Message)"
}
Log 'Exited.'
