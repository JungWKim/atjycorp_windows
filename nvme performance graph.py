import pandas as pd
import matplotlib.pyplot as plt
import numpy as np

# 벤더별 데이터 (이미지 기반)
data = {
    "Vendor": [
        "Kunlun 1.6TB",
        "Union 1.6TB",
        "Samsung 1.6TB",
        "Dapustor 1.6TB"
    ],
    "Seq_Read": [7026, 7287, 6932, 6236],
    "Seq_Write": [2710, 2712, 2810, 2660],
    "Rand_Read": [999, 1007, 918, 1044],
    "Rand_Write": [367, 299, 246, 285]
}

df = pd.DataFrame(data)

plt.figure(figsize=(14, 8))
x = np.arange(len(df["Vendor"]))
bar_width = 0.2

bars1 = plt.bar(x - 0.5*bar_width, df["Seq_Read"], width=bar_width, label="128K Seq Read (MB/s)")
#bars2 = plt.bar(x - 0.5*bar_width, df["Seq_Write"], width=bar_width, label="128K Seq Write (MB/s)")
#bars3 = plt.bar(x - 0.5*bar_width, df["Rand_Read"], width=bar_width, label="4K Random Read (kIOPS)")
#bars4 = plt.bar(x - 0.5*bar_width, df["Rand_Write"], width=bar_width, label="4K Random Write (kIOPS)")

#for bars in [bars1, bars2, bars3, bars4]:
for bars in [bars1]:
    for bar in bars:
        height = bar.get_height()
        plt.text(
            bar.get_x() + bar.get_width() / 2,  # x 좌표 (막대 중앙)
            height + (height * 0.01),          # y 좌표 (막대 꼭대기 바로 위)
            f"{height:.0f}",                   # 표시할 숫자 (소수점 없이)
            ha='center', va='bottom', fontsize=13, fontweight='bold', rotation=0
        )

plt.xticks(x, df["Vendor"], rotation=45, ha="right", fontsize=20)
plt.ylabel("Performance (MB/s)", fontsize=20)
plt.title("128K Seq Read", fontsize=25, fontweight='bold')
#plt.legend()
plt.grid(axis="y", linestyle="--", alpha=0.6)
plt.tight_layout()
plt.show()