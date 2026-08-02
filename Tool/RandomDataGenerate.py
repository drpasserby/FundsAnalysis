"""
生成随机流水数据
版本：1.0.0
------------------------------------------------------------------------------------------------------------------------
生成随机100条交易数据并保存为 Excel 文件
格式为：
用户方 | 支出/收入 | 客户方 | 金额（元）
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
for _ in range(100):
    a = random.choice(people)
    c = random.choice(people)

    while c == a:
        c = random.choice(people)


    direction = random.choice(['支出', '收入'])
    amount = random.randint(100, 20000)  # 随机金额 100~20000

    records.append([a, direction, c, amount])


df = pd.DataFrame(records, columns=['用户方', '支出/收入', '客户方', '金额（元）'])
df.to_excel('流水透视.xlsx', index=False)

print("已生成 流水透视.xlsx，包含 100 条交易记录，共 {} 个不同人员。".format(len(people)))