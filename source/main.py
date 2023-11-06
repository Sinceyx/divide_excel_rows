import pandas as pd

# 加载Excel文件，假定第一行是表头
df = pd.read_excel('/Users/xiaoyu/Downloads/input.xlsx', header=0)

# 创建一个空的DataFrame用于存储结果
result_df = pd.DataFrame()

# 用于记录当前的组
group = []
# 用于记录当前组的THK最大值和最小值
group_thk_min = None
group_thk_max = None

# 遍历DataFrame，从第二行数据开始
for index, row in df.iterrows():
    thk = row['THK']  # 假设G列的标题是'THK'
    # 如果当前没有组，或者新值能够加入到当前组中
    if not group or ((max(group_thk_max, thk) - min(group_thk_min, thk) <= 7) and len(group) < 3):
        group.append(row)
        group_thk_max = max(group_thk_max or thk, thk)
        group_thk_min = min(group_thk_min or thk, thk)
    else:
        # 如果当前值无法加入到当前组，将当前组的数据添加到result_df中，并重置组
        # 添加当前组数据
        result_df = pd.concat([result_df, pd.DataFrame(group)], ignore_index=True)
        # 添加一个空行
        result_df = pd.concat([result_df, pd.DataFrame([{}])], ignore_index=True)
        group = [row]  # 开始新的组
        group_thk_min, group_thk_max = thk, thk

# 处理最后一组数据，如果存在
if group:
    result_df = pd.concat([result_df, pd.DataFrame(group)], ignore_index=True)
    # 添加一个空行
    result_df = pd.concat([result_df, pd.DataFrame([{}])], ignore_index=True)

# 删除最后一个空行
result_df = result_df[:-1] if result_df.iloc[-1].isnull().all() else result_df

# 将结果写入到新的Excel文件
result_df.to_excel('处理后的文件路径.xlsx', index=False)
