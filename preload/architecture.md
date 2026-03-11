# 架构：怎么组织的？

## 关键文件/目录

| 路径 | 作用 |
|------|------|
| QuickPaths.cs | 主程序（1534 行），全部逻辑 + UI + 安装/卸载 |
| build.cmd | csc.exe 编译入口 |
| app.manifest | DPI 感知（Per-Monitor V2） |
| setup.cs | exe 安装器源码 |
| setup.ps1 | 一行在线安装脚本（irm \| iex） |
| Install.cmd / Uninstall.cmd | 本地安装/卸载 |
| assets/ | breathing.gif + hero.png |
| _archive/v2.0-wpf/ | 旧 WPF 版本备份 |

## 技术栈

- C# WinForms + GDI+（.NET Framework 4.x，Windows 自带）
- csc.exe 编译（C# 5 语法限制）
- 单实例：Named Mutex `Global\QuickPaths_Singleton`
- 持久化：config.json（窗口位置/缩放）+ paths.json（路径列表）

## 三层保活

1. 注册表 HKCU Run 自启动
2. 崩溃自动重启（exitCode≠0，最多 5 次）
3. 计划任务看门狗（每 5 分钟）

## 版本演进

v1.0（PowerShell+VBS）→ v2.0（WPF，307MB）→ v3.0（WinForms，14MB）
