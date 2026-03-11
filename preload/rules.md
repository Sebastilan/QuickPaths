# 规矩：什么别碰？

## 11 大踩坑（CLAUDE.md 已记录）

1. 多显示器负坐标处理
2. FolderBrowserDialog 触发面板折叠
3. 剪贴板占用保护
4. CSDN 外链图片转存失败
5. 显示设置变更防抖
6. 文件删除后引用清理
7. csc.exe 只支持 C# 5（禁止字符串插值/`?.`/`=>`/auto-property 初始化）
8. WinForms 半透明 BackColor 处理
9. Region 裁剪抗锯齿
10. GDI+ PathGradientBrush 退化路径异常
11. SetUnhandledExceptionMode 时序要求

## CSDN 文章同步

大改后需同步 `csdn-auto-publisher/articles/quickpaths-claude-launcher/content.md`。

## 编译注意

Git Bash 内 `/` 开头参数被当路径展开 → 用 `powershell -Command` 或 `cmd.exe /c` 包裹 csc.exe。
