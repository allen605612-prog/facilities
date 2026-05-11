import csv, random

surnames = ["王","林","陳","張","李","黃","吳","劉","蔡","鄭","許","謝","洪","邱","楊",
            "周","江","游","方","何","施","盧","戴","蘇","紀","葉","卓","錢","羅","魏",
            "宋","潘","鍾","馬","范","孫","趙","高","曾","石"]
male_names = ["志豪","建宏","俊賢","宗翰","明哲","志偉","建志","冠廷","俊明","志龍",
              "建國","志遠","俊傑","文凱","哲宇","冠宇","彥廷","彥宏","宗諺","宇軒",
              "柏翰","昱廷","聖凱","家豪","文翔","佳翰","宏達","俊宇","哲瑋","宗緯",
              "品睿","冠霖","威廷","仁豪","育誠","冠穎","志遠","承翰","彥霖","博凱"]
female_names = ["雅婷","淑芬","美玲","佳蓉","雅琪","欣怡","淑華","雅雯","佩珊","淑真",
                "雅惠","美慧","雅文","淑玲","雅萍","怡君","欣儀","佩瑩","雅琳","淑娟",
                "美華","雅慧","佩君","雅玲","淑惠","怡萱","欣妤","佩宜","雅筑","淑萍",
                "品妍","佳穎","宜蓁","欣彤","雅涵","淑媛","美瑤","佩蓉","怡蓁","雅晴"]

random.seed(42)
rows = []
for i in range(1, 3001):
    sid = f"S{i:04d}"
    surname = random.choice(surnames)
    if random.random() < 0.5:
        name = surname + random.choice(male_names)
    else:
        name = surname + random.choice(female_names)
    c = random.randint(45, 99)
    m = random.randint(45, 99)
    e = random.randint(45, 99)
    total = c + m + e
    avg = round(total / 3, 1)
    rows.append([sid, name, c, m, e, total, avg])

with open("students.csv", "w", newline="", encoding="utf-8-sig") as f:
    w = csv.writer(f)
    w.writerow(["學號","姓名","國文","數學","英文","總分","平均"])
    w.writerows(rows)

print(f"已產生 {len(rows)} 筆資料 → students.csv")
