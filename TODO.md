# TODO

## 已完成
- [x] 基础功能：悬浮圆点、展开/收起、路径复制、添加/删除
- [x] 多选添加：IFileOpenDialog COM 接口，支持 Ctrl+点击
- [x] 排序：↑ 上移按钮
- [x] 位置记忆：拖动后自动保存，重启恢复
- [x] 多显示器支持：VirtualScreen 坐标验证
- [x] 自启动：开机自启(Startup VBS) + 看门狗(Task Scheduler)
- [x] 稳定性：睡眠/锁屏恢复、显示器热插拔、原子写入、剪贴板保护
- [x] 项目迁移至 CCA 路径
- [x] 失效路径自动清理：展开面板时检测，不存在的路径自动移除
- [x] 修复 SavePaths 数据丢失（File.Replace + $null 在 PS 5.1 不兼容）
- [x] 开源准备：一键安装/卸载、动态路径、LICENSE、README 英文版
- [x] 滚轮缩放：鼠标滚轮调节圆点大小（0.5x–3.0x），比例持久化
- [x] 笔头光晕：三层径向渐变发光（外30px+中16px+笔头），跟呼吸同步脉动
- [x] 稳定性改造：全局异常处理（Dispatcher/AppDomain）、退出码约定、VBS 崩溃恢复包裹
- [x] Dispatcher.Invoke→BeginInvoke 修复渲染管线死锁
- [x] 显示器变更防抖（1500ms DispatcherTimer）
- [x] 所有事件处理器 try-catch 保护
- [x] 增强日志：毫秒时间戳、LogError 完整堆栈、LogCrashContext 运行上下文
- [x] **v2.0 C# 重写**：PowerShell → 编译型 C# WPF exe，消除 PS 层所有怪癖，砍掉 VBS 崩溃恢复/wrapper.lock，自启动改为注册表，install/uninstall 内置到 exe
- [x] **v3.0 WPF→WinForms 轻量化重写**：内存从 ~307MB 降至 ~14MB。GDI+ 自绘 ECG 动画，Region 裁剪圆角，保持全部功能兼容

## 未来可选
- [ ] 主题切换：深色/浅色自适应
- [ ] 拖拽排序：替代当前的 ↑ 按钮
