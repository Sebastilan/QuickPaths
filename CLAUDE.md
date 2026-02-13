# QuickPaths

## 项目概述
桌面悬浮路径快捷复制工具。小圆点常驻桌面，点击展开路径列表，单击即复制。

- **技术栈**：PowerShell 5.1 + WPF（零外部依赖）
- **路径**：`C:\Users\ligon\CCA\QuickPaths`
- **当前阶段**：v1.1 稳定性改造完成。下一步见 TODO.md「未来可选」

## 文件结构

| 文件 | 用途 |
|------|------|
| `QuickPaths.ps1` | 主程序（WPF 悬浮窗 + 全部逻辑） |
| `setup.ps1` | 在线安装脚本（`irm \| iex` 一行命令安装） |
| `setup.cs` | exe 安装器 C# 源码（不入库，用 csc.exe 编译） |
| `QuickPathsSetup.exe` | exe 安装器（不入库，上传到 GitHub Releases） |
| `Install.cmd` | 本地安装入口（clone 后双击） |
| `install.ps1` | 安装逻辑（生成 VBS 崩溃恢复包裹、清理旧看门狗、启动） |
| `Uninstall.cmd` | 一键卸载入口 |
| `uninstall.ps1` | 卸载逻辑（停进程、杀 VBS 包裹、清 VBS、清任务） |
| `_restart.ps1` | 开发用重启脚本（不入库） |
| `paths.json` | 用户路径数据（运行时生成，不入库） |
| `config.json` | 窗口位置（运行时生成，不入库） |
| `wrapper.lock` | VBS 包裹进程锁（运行时生成，不入库） |

## 外部依赖文件

| 文件 | 路径 |
|------|------|
| 开机自启+崩溃恢复 | `%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup\QuickPaths.vbs`（Do...Loop 包裹进程，exit 0 停止，非 0 自动重启） |

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
11. **CSDN 外链图片转存失败**：jsDelivr CDN / GitHub raw URL 都会被 CSDN 拦截（"图片转存失败"）。必须手动上传到 CSDN 图床，用 `i-blog.csdnimg.cn/direct/xxx` 链接
12. **Dispatcher.Invoke 死锁**：`DisplaySettingsChanged` 等非 UI 线程回调中用 `Dispatcher.Invoke`（同步），若渲染管线 `DUCE+Channel.SyncFlush` 崩溃则死锁。**解决**：一律用 `Dispatcher.BeginInvoke`（异步）
13. **显示器变更防抖**：拔/插显示器会连续触发多次 `DisplaySettingsChanged`，每次都做 UI 操作会锤击渲染管线。用 DispatcherTimer 1500ms 防抖，等配置稳定后再校验
14. **Task Scheduler 看门狗已废弃**：改为 Startup VBS Do...Loop 包裹进程，exit 0 停止循环，非 0 两秒后重启。5 次/60s 退避 30s。wrapper.lock 防双重包裹

## CSDN 图床链接

| 图片 | CSDN 图床 URL |
|------|--------------|
| hero.png | `https://i-blog.csdnimg.cn/direct/259b8f03cc454d08af67bc8db0fc6078.png` |
| breathing.gif | `https://i-blog.csdnimg.cn/direct/914ba54968f9426fbda96ede93e65fc7.gif` |
