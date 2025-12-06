import pandas as pd
import matplotlib.pyplot as plt
import numpy as np

# 데이터 입력
load_levels = ["idle", "30%", "50%", "70%", "100%", "idle after"]

power_4410y = [276, 336, 348, 360, 360, 288]
power_4510y = [276, 348, 360, 372, 372, 276]

plt.figure(figsize=(8, 5))

# 그래프 그리기
plt.plot(load_levels, power_4410y, marker='o', label='4410Y')
plt.plot(load_levels, power_4510y, marker='o', label='4510Y')

# 숫자 표시
for x, y in zip(load_levels, power_4410y):
    plt.text(x, y + 3, str(y), ha='center', fontsize=18)

for x, y in zip(load_levels, power_4510y):
    plt.text(x, y + 3, str(y), ha='center', fontsize=18)

# 글씨 크기 조정
plt.title("CPU Power Consumption Comparison (4410Y vs 4510Y)", fontsize=22)
plt.xlabel("Load Level", fontsize=18)
plt.ylabel("Power Consumption (W)", fontsize=18)
plt.xticks(fontsize=18)
plt.yticks(fontsize=18)
plt.legend(fontsize=18)

plt.grid(True)
plt.tight_layout()
plt.show()