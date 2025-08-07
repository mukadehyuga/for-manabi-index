import numpy as np
import pandas as pd
import matplotlib.pyplot as plt

# ファイル名
filename = "O73_A152_F1k.csv"

# 読み込み（文字コード変更可）
df = pd.read_csv(filename, encoding="cp932", skiprows=3)

# 列確認と選択
print("\n利用可能な列：")
for i, col in enumerate(df.columns):
    print(f"{i}: {col}")

x_index = int(input("X軸に使う列番号（例: 時間）を入力してください: "))
y_index = int(input("Y軸に使う列番号（例: 電圧）を入力してください: "))

# データ整形
x = df.iloc[:, x_index].astype(float)
y = df.iloc[:, y_index].astype(float)
valid = (~pd.isnull(x)) & (~pd.isnull(y))
x = x[valid]
y = y[valid]

# サンプリング周波数とFFT
dt = x.iloc[1] - x.iloc[0]
fs = 1.0 / dt
N = len(y)
freqs = np.fft.fftfreq(N, d=dt)
fft_vals = np.fft.fft(y)
amplitudes = np.abs(fft_vals) / N

# 正の周波数成分のみ抽出
mask = freqs >= 0
freqs = freqs[mask]
amplitudes = amplitudes[mask]

# ★ ユーザーが範囲を指定 ★
print("\n--- 表示範囲を指定してください（Enterで自動設定） ---")
try:
    f_min = float(input("x軸（周波数）の最小値 [Hz]（例: 0）: ") or freqs.min())
    f_max = float(input("x軸（周波数）の最大値 [Hz]（例: 1000）: ") or freqs.max())
    a_min = float(input("y軸（振幅）の最小値（例: 0）: ") or amplitudes.min())
    a_max = float(input("y軸（振幅）の最大値（例: 0.01）: ") or amplitudes.max())
except ValueError:
    print("無効な値が含まれています。自動範囲に設定します。")
    f_min, f_max = freqs.min(), freqs.max()
    a_min, a_max = amplitudes.min(), amplitudes.max()

# 範囲フィルタ
plot_mask = (freqs >= f_min) & (freqs <= f_max)
plot_freqs = freqs[plot_mask]
plot_amps = amplitudes[plot_mask]

# グラフ描画
plt.figure(figsize=(10, 6))
plt.plot(plot_freqs, plot_amps)
plt.xlabel("Frequency [Hz]")
plt.ylabel("Amplitude")
plt.title(f"FFT of column: {df.columns[y_index]}")
plt.grid(True)
plt.ylim(a_min, a_max)
plt.xlim(f_min, f_max)
plt.tight_layout()
plt.show()

# 保存確認
save = input("このFFTグラフを保存しますか？（はい/いいえ）: ")
if save.strip().lower() in ["はい", "yes", "y"]:
    filename_out = f"fft_{df.columns[y_index]}_zoom.png"
    plt.figure(figsize=(10, 6))
    plt.plot(plot_freqs, plot_amps)
    plt.xlabel("Frequency [Hz]")
    plt.ylabel("Amplitude")
    plt.title(f"FFT of column: {df.columns[y_index]}")
    plt.grid(True)
    plt.ylim(a_min, a_max)
    plt.xlim(f_min, f_max)
    plt.tight_layout()
    plt.savefig(filename_out, dpi=300)
    print(f"{filename_out} として保存しました。")
else:
    print("保存をキャンセルしました。")

# Excel保存確認
excel_save = input("このFFTデータをExcelファイルに保存しますか？（はい/いいえ）: ")
if excel_save.strip().lower() in ["はい", "yes", "y"]:
    fft_df = pd.DataFrame({
        "Frequency [Hz]": plot_freqs,
        "Amplitude": plot_amps
    })
    excel_filename = f"fft_{df.columns[y_index]}_data.xlsx"
    fft_df.to_excel(excel_filename, index=False)
    print(f"{excel_filename} としてExcelに保存しました。")
else:
    print("Excelへの保存をキャンセルしました。")
