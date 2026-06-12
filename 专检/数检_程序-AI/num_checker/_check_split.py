from num_checker.train_crf import split_seed
train, test = split_seed()
print("测试集:")
for t in test:
    label = "数值" if t[1] else "非数值"
    print(f"  [{t[2]}] {label}: {t[0]}")
print("\n训练集中包含的目标词:", sorted(set(t[2] for t in train)))
print("测试集中包含的目标词:", sorted(set(t[2] for t in test)))
