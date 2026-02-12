# QuickPaths

## 项目概述
桌面悬浮路径快捷复制工具。小圆点常驻桌面，点击展开路径列表，单击即复制。

- **技术栈**：PowerShell 5.1 + WPF（零外部依赖）
- **路径**：`C:\Users\ligon\CCA\QuickPaths`
- **当前阶段**：v1 开源就绪。无进行中任务，下一步见 TODO.md「未来可选」

## 文件结构

| 文件 | 用途 |
|------|------|
| `QuickPaths.ps1` | 主程序（WPF 悬浮窗 + 全部逻辑） |
| `Install.cmd` | 一键安装入口（双击即用） |
| `install.ps1` | 安装逻辑（生成 VBS、注册看门狗、启动） |
| `Uninstall.cmd` | 一键卸载入口 |
| `uninstall.ps1` | 卸载逻辑（停进程、清 VBS、清任务） |
| `watchdog.ps1` | 看门狗脚本（Task Scheduler 每 3 分钟检查） |
| `_restart.ps1` | 开发用重启脚本（不入库） |
| `paths.json` | 用户路径数据（运行时生成，不入库） |
| `config.json` | 窗口位置（运行时生成，不入库） |

## 外部依赖文件

| 文件 | 路径 |
|------|------|
| 开机自启 | `%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup\QuickPaths.vbs` |
| 看门狗任务 | Task Scheduler: `QuickPaths_Watchdog` |

## 文档联动表

| 触发事件 | 必须更新 |
|---------|---------|
| 新增/完成功能 | `TODO.md` |
| 踩坑/重要决策 | 本文件「已踩过的坑」 |
| 修改文件结构 | 本文件「文件结构」 |
| 变更外部路径引用 | 本文件「外部依赖文件」+ install.ps1 + uninstall.ps1 + Startup VBS |

## 已踩过的坑

1. **PS 5.1 + 中文 = 必须 UTF-8 BOM**：`.ps1` 文件无 BOM 时中文乱码，`[char]0xHHHH` 可规避
2. **JSON 文件不能有 BOM**：`ConvertFrom-Json` 遇 BOM 会把数组当单对象，读取时 `TrimStart([char]0xFEFF)`
3. **ConvertTo-Json 单元素数组**：PS 5.1 会去掉 `[]` 包裹，需 `Count -eq 1` 时手动加
4. **WPF 逻辑像素 vs WinForms 物理像素**：定位用 `SystemParameters::WorkArea`，不用 `Screen.WorkingArea`
5. **多显示器坐标可为负**：副屏在上方/左方时坐标负值，验证用 `VirtualScreen*` 系列属性
6. **FolderBrowserDialog 导致面板收起**：对话框夺焦触发 Deactivated，用 `$dialogOpen` 标记保护
7. **File.Replace 原子写入**：防止写入中途崩溃丢数据，先写 .tmp 再 Replace
8. **剪贴板可能被占用**：`Clipboard.SetText` 需 try-catch，其他程序锁住时会抛异常
9. **File.Replace + $null 在 PS 5.1 崩溃**：`[System.IO.File]::Replace($src, $dst, $null)` 报"路径形式不合法"，改用 Remove+Move 模式
10. **setup_watchdog.ps1 已废弃**：功能合并到 install.ps1，安装时动态生成 watchdog.ps1 和 Startup VBS
