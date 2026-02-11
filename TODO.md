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

## 未来可选
- [ ] 开源准备：README 英文版、LICENSE
- [ ] 主题切换：深色/浅色自适应
- [ ] 拖拽排序：替代当前的 ↑ 按钮
