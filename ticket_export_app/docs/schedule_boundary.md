# Schedule / Tickets Boundary

本文件记录 `core/tickets.py` 与 `core/schedule_v2.py` 的当前职责边界，作为后续架构重塑的低风险参考。

## 当前边界

`core/schedule_v2.py` 当前是安全代理入口，动态调用 `core.tickets.schedule()` 中保留的已验证排程实现。现阶段不要在迁移前改变排程行为。

`core/tickets.py` 当前仍承担：

- `schedule()`：排程 rows 生成。
- `analyze_schedule()`：兼容入口，调用 `core.analysis.analyze_schedule_v2()` 后补齐旧 UI 字段。
- `schedule_and_export()`：排程、分析、导出组合票的旧入口。
- Excel 导出相关逻辑。

## tickets.schedule() 输入

主要参数：

- `step_defs`：岗位定义列表。
- `cars`：需要生成的车辆总数。
- `vehicle_counts`：A/B/C 数量或比例来源，按模式解释。
- `sequence_mode`：`grouped`、`alternate`、`ratio`。
- `max_consecutive`：交替混流时的最大连续台数约束。
- `ratio_pattern`：按比例投车时的 A/B/C 比例块。
- `launch_takt`：固定投车节拍；按比例投车模式下通常等于目标节拍。

## tickets.schedule() 输出

返回：

- `rows`：每车每岗位的排程段列表。
- `max_finish`：所有车辆最后离开时间。

## rows 关键字段

当前 UI、分析与导出主要依赖以下字段：

- `car`：车辆编号。
- `car_type`：车型，通常为 A/B/C。
- `step_seq`：岗位序号。
- `step_display`：岗位展示名称。
- `group`：岗位设备/资源分组。
- `start`：本岗位开始加工时间。
- `svc_finish`：本岗位加工完成时间。
- `depart`：车辆离开本岗位时间，可能包含等待下一岗位的阻塞时间。
- `dur`：本岗位加工工时。
- `block_wait`：加工完成后等待下一岗位/资源的时间。
- `launch_wait`：进入当前岗位前因投车或资源约束产生的等待。
- `resource_key`：资源占用 key。
- `line_no`：车辆当前实际线别。

## tickets.analyze_schedule() 输入输出

输入：

- `rows`
- `max_finish`
- `target_takt`

输出：

- `summary`
- `station_summary` / `stations`
- `car_type_summary` / `car_types`
- `station_stats`
- `type_stats`
- `car_stats`

`tickets.analyze_schedule()` 目前会调用 `core.analysis.analyze_schedule_v2()`，再补齐旧 UI 兼容字段。

## summary 关键字段

当前常用字段包括：

- `max_finish`
- `target_takt`
- `total_wait`
- `avg_wait`
- `average_wait`
- `blocking_result`
- `blocking_station_count`
- `blocking_station_text`
- `total_blocking_time`
- `overflow_vehicle_count`
- `batch_overrun_time`
- `batch_overrun_cars`
- `process_over_takt_root_text`
- `takt_result`
- `over_takt_station_count`

## v2-6D 使用的 summary 字段

按比例投车时间窗口分析会写入并使用：

- `analysis_time_seconds`
- `analysis_time_minutes`
- `theoretical_launch_count`
- `station_count`
- `line_lead_time`
- `planned_output_count_in_window`
- `actual_output_count_in_window`
- `achievement_rate`
- `planned_n_finish_time`
- `actual_n_finish_time`
- `finish_delta`
- `actual_line_takt_in_window`
- `time_window_result`

兼容字段：

- `display_actual_output_count_in_window`
- `actual_equivalent_count_in_window`
- `actual_production_takt_in_window`

v2-6D 最终判定基于完整车辆整数台和第 N 台准时完成，不再基于工程数折算台数、小数台或首尾平均下线节拍。
