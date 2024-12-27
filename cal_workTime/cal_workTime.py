import pandas as pd
import calendar
from datetime import datetime, timedelta

def calculate_daily_work_duration(start_time, end_time):
    """
    计算单日工作时长，单位：小时
    参数：
    - start_time: 开始时间，格式 "HH:MM"
    - end_time: 结束时间，格式 "HH:MM"
    返回：
    - 时长（浮点数），处理跨天情况
    """
    start = datetime.strptime(start_time, "%H:%M")
    end = datetime.strptime(end_time, "%H:%M")
    # 如果结束时间早于开始时间，说明跨天
    if end < start:
        end += timedelta(days=1)
    duration = (end - start).total_seconds() / 3600  # 转换为小时
    return duration

def calculate_daily_difference(start_time, end_time, is_workday):
    """
    计算与标准打卡时长的差异
    参数：
    - start_time: 开始时间，格式 "HH:MM"
    - end_time: 结束时间，格式 "HH:MM"
    - is_workday: 是否为工作日
    返回：
    - 差异时长（浮点数，正数表示超时，负数表示少时）
    """
    if not start_time or not end_time:
        if is_workday:
            return -10, 0  # 工作日缺卡按少 10 小时处理
        return None, 0  # 非工作日缺卡不计时长
    # 实际打卡时长
    actual_hours = calculate_daily_work_duration(start_time, end_time)
    # 标准时长（工作日为 10 小时，非工作日为 0 小时）
    standard_hours = 10 if is_workday else 0
    return actual_hours - standard_hours, actual_hours

def process_punch_data(file_path, output_path):
    """
    处理打卡数据，生成每日差异，并计算总时长信息
    参数：
    - file_path: 输入文件路径
    - output_path: 输出文件路径
    """
    # 读取 Excel 文件
    df = pd.read_excel(file_path, header=0)
    print(df)
    year = int(df.iloc[0, 0])  # 年份
    month = int(df.iloc[0, 1])  # 月份

    # 获取当月总天数和1号是星期几
    start_weekday, num_days = calendar.monthrange(year, month)

    # 初始化统计变量
    total_actual_hours = 0
    total_standard_hours = 0

    # 创建结果数据
    results = []

    for day in range(1, num_days + 1):
        weekday = (start_weekday + day - 1) % 7  # 计算当天是星期几
        is_workday = 0 <= weekday <= 4  # 判断是否为工作日

        # 获取打卡时间
        start_time = df.iloc[0, day + 1] if pd.notna(df.iloc[0, day + 1]) else None
        end_time = df.iloc[1, day + 1] if pd.notna(df.iloc[1, day + 1]) else None

        # 计算差异和实际时长
        difference, actual_hours = calculate_daily_difference(start_time, end_time, is_workday)
        if is_workday:
            total_standard_hours += 10  # 工作日标准时间为 10 小时
        total_actual_hours += actual_hours  # 累计实际打卡时间

        results.append({
            "日期": f"{year}-{month:02d}-{day:02d}",
            "开始时间": start_time,
            "结束时间": end_time,
            "是否工作日": "是" if is_workday else "否",
            "时长差异(小时)": difference
        })

    # 转为 DataFrame 并保存为 Excel
    result_df = pd.DataFrame(results)
    result_df.to_excel(output_path, index=False)

    # 输出统计信息
    print(f"结果已保存到 {output_path}")
    print(f"当前已打卡时长：{total_actual_hours:.2f} 小时")
    print(f"本月规定的打卡时间：{total_standard_hours:.2f} 小时")

if __name__ == "__main__":
    input_file = "punch_time.xlsx"  # 输入文件
    output_file = "punch_time_diff.xlsx"  # 输出文件
    process_punch_data(input_file, output_file)
