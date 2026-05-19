# v2.9 Result Risk Final Release

## 版本信息

- 版本名称：M-Line 混流节拍仿真系统
- 版本号：2.9.0
- tag：v2.9-result-risk-final
- commit hash：以本归档提交为准，最终 hash 见 `git log --oneline --decorate -5`
- commit message：finalize M-Line result metrics and docs

## 打包产物

上一轮已验证打包产物路径：

```text
/Users/zhongwenjie/程序/ticket_export_app/dist/生产指示小工具
```

本轮不重新打包。

## 验证摘要

- 语法检查：`export_ticket_window.py`、`analysis.py`、`tickets.py`、`schedule_v2.py`、`input_parser.py`、`main.py` 已按归档流程检查。
- 核心逻辑：本轮不修改 `core/tickets.py`、`core/schedule_v2.py`、`core/input_parser.py`、`core/analysis.py` 和 `main.py`。
- 模型结果区：保留 5 个客观指标，下线车辆、达标车辆、达标率、整体节拍、累计阻塞。
- 风险提示：按工位能力超节拍、阻塞工程、整体节拍输出；无风险时显示“暂无明显风险”。
- 实时节拍：保留在仿真回放控制栏，仅作为播放观察指标。

## 未处理项

- `checker_ui_app_v2`：保持既有外部 modified 状态，不纳入本次提交。
- `ticket_export_app/docs/reference/`：参考资料目录不纳入本次提交。
- `ticket_export_app/docs/ui_layout_plan.md`：既有未跟踪文档不纳入本次提交。

## 后续建议

- 用户完成打包产物人工点检后，再决定是否补充发布说明或正式分发。
- 如需让文件内记录精确 commit hash，可在后续 release 管理文档中记录，避免提交哈希自引用造成不一致。
- 车辆日志 RESULT 字段仍作为观察信息保留，后续如需要可单独改为与模型结果区完全一致的工位能力口径。
