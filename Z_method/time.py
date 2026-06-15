from calendar import monthrange
from datetime import datetime, timedelta


def get_month_range(months_ago, reference_date=None):
    """获取 reference_date 所在月往前第 months_ago 个月的 1 号至月末（格式 m.d-m.d）"""
    if reference_date is None:
        reference_date = datetime.now().date()
    elif isinstance(reference_date, datetime):
        reference_date = reference_date.date()

    # 计算目标月份（以 reference_date 所在月的 1 号为基准）
    target_date = reference_date.replace(day=1)
    for _ in range(months_ago):
        target_date = target_date - timedelta(days=1)
        target_date = target_date.replace(day=1)

    # 获取月份的天数
    _, last_day_num = monthrange(target_date.year, target_date.month)

    # 格式化
    start = f"{target_date.month}.{target_date.day}"
    end = f"{target_date.month}.{last_day_num}"

    return f"{start}-{end}"
