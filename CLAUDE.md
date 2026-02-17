# QuickPaths

## 项目概述
桌面悬浮路径快捷复制工具。小圆点常驻桌面，点击展开路径列表，单击即复制。

- **技术栈**：C# WinForms + GDI+ 自绘 + csc.exe 编译（零外部依赖，.NET Framework 4.x 自带）
- **路径**：`C:\Users\ligon\CCA\QuickPaths`
- **当前阶段**：v3.0 WinForms 轻量化重写完成（内存 ~14MB，WPF 版备份在 `_archive/v2.0-wpf/`）。下一步见 TODO.md「未来可选」

## 文件结构

| 文件 | 用途 |
|------|------|
| `QuickPaths.cs` | 主程序源码（WinForms + GDI+ 自绘 + 全部逻辑 + install/uninstall） |
| `build.cmd` | csc.exe 编译脚本 |
| `QuickPaths.exe` | 编译产物（不入库，build.cmd 生成或从 Releases 下载） |
| `setup.ps1` | 在线安装脚本（`irm \| iex` 一行命令安装） |
| `setup.cs` | exe 安装器 C# 源码 |
| `QuickPathsSetup.exe` | exe 安装器（不入库，上传到 GitHub Releases） |
| `Install.cmd` | 本地安装入口（clone 后双击） |
| `Uninstall.cmd` | 一键卸载入口 |
| `_restart.ps1` | 开发用重启脚本（不入库） |
| `paths.json` | 用户路径数据（运行时生成，不入库） |
| `config.json` | 窗口位置（运行时生成，不入库） |

## 三模式运行

```
QuickPaths.exe              → 正常运行（悬浮圆点）
QuickPaths.exe --install    → 注册自启动 + 清理旧文件 + 启动 + 弹窗确认
QuickPaths.exe --uninstall  → 停止 + 反注册 + 清理 + 询问删数据 + 弹窗确认
```

## 保活机制（三层保障）

| 层 | 机制 | 覆盖场景 |
|----|------|---------|
| 1 | 注册表 `HKCU\...\Run\QuickPaths` | 开机/登录自启动 |
| 2 | 崩溃自重启（exitCode != 0 → 延迟 2s 重启，上限 5 次） | 运行时崩溃 |
| 3 | 计划任务 `QuickPaths_KeepAlive`（每 5 分钟） | 被杀进程、连续崩溃超限 |

单实例由 Named Mutex `Global\QuickPaths_Singleton` 保证，重复启动无副作用。

## 文档联动表

| 触发事件 | 必须更新 |
|---------|---------|
| 新增/完成功能 | `TODO.md` |
| 踩坑/重要决策 | 本文件「已踩过的坑」 |
| 修改文件结构 | 本文件「文件结构」 |
| 变更外部路径引用 | 本文件 |

## 已踩过的坑

1. **多显示器坐标可为负**：副屏在上方/左方时坐标负值，验证用 `SystemInformation.VirtualScreen`
2. **FolderBrowserDialog 导致面板收起**：对话框夺焦触发 Deactivate，用 `dialogOpen` 标记保护
3. **剪贴板可能被占用**：`Clipboard.SetText` 需 try-catch，其他程序锁住时会抛异常
4. **CSDN 外链图片转存失败**：jsDelivr CDN / GitHub raw URL 都会被 CSDN 拦截。必须手动上传到 CSDN 图床
5. **显示器变更防抖**：拔/插显示器会连续触发多次 `DisplaySettingsChanged`，用 Timer 1500ms 防抖
6. **删文件≠删引用**：废弃外部资源时必须同步清理所有引用方（计划任务、注册表、Startup 快捷方式等）
7. **csc.exe 版本仅支持 C# 5**：不能用 `$""` 字符串插值、`?.` 空条件运算符、`=>` 表达式体成员
8. **WinForms 控件不支持半透明 BackColor**：需预混合 alpha 颜色到面板背景色上（`Blend()` 函数）
9. **Region 裁剪无抗锯齿**：圆角边缘会有锯齿，这是 WinForms 固有限制，可接受
10. **GDI+ PathGradientBrush 退化路径**：极小半径时可能抛异常，需 try-catch 保护
11. **SetUnhandledExceptionMode 时序**：必须在创建任何 Form 之前调用，否则抛 InvalidOperationException 且无法捕获（因为异常处理器尚未注册）

## CSDN 图床链接

| 图片 | CSDN 图床 URL |
|------|--------------|
| hero.png | `https://i-blog.csdnimg.cn/direct/259b8f03cc454d08af67bc8db0fc6078.png` |
| breathing.gif | `https://i-blog.csdnimg.cn/direct/914ba54968f9426fbda96ede93e65fc7.gif` |
