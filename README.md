# M-Line 混流节拍仿真系统

M-Line 混流节拍仿真系统是一个用于多车型混流排程仿真、工位能力节拍校核、整体节拍观察、阻塞风险提示和组合票导出辅助的桌面工具。

当前归档版本：v2.9-result-risk-final

## 主要功能

- 多车型 A/B/C 混流排程仿真
- 工位能力节拍校核
- 下线车辆、达标车辆、达标率、整体节拍、累计阻塞显示
- 风险提示：工位能力超节拍、阻塞工程、整体节拍
- 仿真回放与实时节拍观察
- 车辆明细 / 调试日志
- Windows GitHub Actions 自动打包

## 使用说明

详见：

ticket_export_app/docs/M-Line混流节拍仿真系统_使用说明_v2.9.md

## Windows 打包

GitHub Actions workflow：

.github/workflows/build-windows.yml

生成 artifact：

M-Line_v2.9_windows_package
