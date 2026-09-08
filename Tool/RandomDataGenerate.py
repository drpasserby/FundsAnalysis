"""
生成随机流水数据
版本：2.0.0
------------------------------------------------------------------------------------------------------------------------
生成随机交易数据并保存为 Excel 文件
格式为：
用户方 | 支出/收入 | 客户方 | 金额（元）

数据特点：
1. 每笔交易成对生成：A 支出 B X 元 与 B 收入 A X 元 一一对应（同一笔资金分别从双方视角记录）；
2. 金额不一定是整数，可为小数，范围 0.01 ~ 20000 元（保留两位小数）；
3. 默认生成 100 笔交易，即 200 条流水记录（每笔两行）。
"""


import pandas as pd
import random

# ---------- 1. 随机名字生成 ----------
surnames = ['张', '李', '王', '刘', '陈', '杨', '黄', '赵', '周', '吴', '徐', '孙', '马', '朱', '胡', '林', '郭', '何', '高', '罗']
given_names = ['伟', '芳', '娜', '秀英', '敏', '静', '丽', '强', '磊', '洋', '勇', '军', '杰', '涛', '明', '超', '平', '刚', '华', '文',
               '斌', '玲', '桂英', '婷', '宇', '浩', '然', '博', '辉', '毅']

def random_name():
    return random.choice(surnames) + random.choice(given_names)


people = set()
while len(people) < 20:
    people.add(random_name())
people = list(people)


records = []
for _ in range(100):  # 100 笔交易，每笔两行 → 200 条流水记录
    a = random.choice(people)
    c = random.choice(people)

    while c == a:
        c = random.choice(people)

    # 金额可为小数，范围 0.01 ~ 20000，保留两位小数
    amount = round(random.uniform(0.01, 20000), 2)

    # 同一笔资金：A 支出 B 与 B 收入 A 金额一致、双方对应
    records.append([a, '支出', c, amount])
    records.append([c, '收入', a, amount])


df = pd.DataFrame(records, columns=['用户方', '支出/收入', '客户方', '金额（元）'])
df.to_excel('流水透视.xlsx', index=False)

print('已生成 流水透视.xlsx：{} 笔交易，共 {} 条流水记录（每笔含支出/收入两条，金额对应一致），{} 个不同人员。'.format(
    len(records) // 2, len(records), len(people)))
