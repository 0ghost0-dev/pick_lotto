import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import json
import os
import random
import bisect
import ctypes
import datetime

def extract_numbers_from_file(file_path):
    numbers = []
    bonus_numbers = []
    latest_round = 0  # 최신 회차 초기화

    df = pd.read_excel(file_path, header=None)
    for _, row in df.iterrows():
        row_values = row.dropna().tolist()

        # 두 번째 열(인덱스 1)에 회차 번호가 있다고 가정
        if len(row_values) > 1 and isinstance(row_values[1], (int, float)):
            if int(row_values[1]) > latest_round:
                latest_round = int(row_values[1])

        # 1~45 숫자만 추출
        numeric_values = [int(v) for v in row_values if isinstance(v, (int, float)) and 1 <= v <= 45]
        if len(numeric_values) >= 7:
            winning = numeric_values[-7:]
            numbers.append(winning[:6])
            bonus_numbers.append(winning[6])

    return numbers, bonus_numbers, latest_round

def calculate_frequencies(all_numbers, all_bonus):
    normal_freq = [0] * 45
    bonus_freq = [0] * 45

    for draw in all_numbers:
        for num in draw:
            normal_freq[num - 1] += 1
    for b in all_bonus:
        bonus_freq[b - 1] += 1

    # 이제 두 개 리스트 모두 리턴
    return normal_freq, bonus_freq

def get_hidden_filename():
    if os.name == 'nt':
        base_dir = "C:\\rand_a"
    else:
        base_dir = "/rand_a"

    os.makedirs(base_dir, exist_ok=True)

    filename = "frequencies.json" if os.name == 'nt' else ".frequencies.json"
    return os.path.join(base_dir, filename)

def save_frequencies(normal_freq, bonus_freq, latest_round):
    filename = get_hidden_filename()
    data = {
        "normal_frequencies": normal_freq,
        "bonus_frequencies": bonus_freq,
        "latest_round": latest_round
    }
    with open(filename, "w", encoding="utf-8") as f:
        json.dump(data, f)

def load_frequencies():
    filename = get_hidden_filename()
    if not os.path.exists(filename):
        return None, None, None
    with open(filename, "r", encoding="utf-8") as f:
        data = json.load(f)
        return data.get("normal_frequencies"), data.get("bonus_frequencies"), data.get("latest_round")

def get_hardware_random_bytes(n):
    if os.name == 'nt':
        BCRYPT_USE_SYSTEM_PREFERRED_RNG = 0x00000002
        bcrypt = ctypes.windll.bcrypt
        buffer = ctypes.create_string_buffer(n)
        status = bcrypt.BCryptGenRandom(None, buffer, n, BCRYPT_USE_SYSTEM_PREFERRED_RNG)
        if status != 0:
            raise OSError(f"BCryptGenRandom failed: {status}")
        return buffer.raw
    else:
        return os.urandom(n)

def get_hardware_random_float():
    b = get_hardware_random_bytes(4)
    return int.from_bytes(b, byteorder='little') / 2**32

def build_weighted_table(weights):
    cumulative = []
    total = 0
    for w in weights:
        total += w
        cumulative.append(total)
    return cumulative, total

def weighted_choice(cumulative_weights, total_weight, method='hardware'):
    if method == 'pseudo':
        r = random.uniform(0, total_weight)
    else:
        r = get_hardware_random_float() * total_weight
    return bisect.bisect_left(cumulative_weights, r) + 1

def get_weighted_unique_numbers(frequencies, method='hardware'):
    cumulative_weights, total_weight = build_weighted_table(frequencies)
    numbers = set()
    while len(numbers) < 6:
        num = weighted_choice(cumulative_weights, total_weight, method)
        numbers.add(num)
    return sorted(numbers)

def generate_numbers():
    method = method_var.get()
    normal_freq, bonus_freq, latest_round = load_frequencies()
    if not normal_freq:
        messagebox.showerror("오류", "빈도수 파일이 없습니다. 먼저 분석을 수행하세요.")
        return

    if include_bonus_var.get():
        freqs = [normal_freq[i] + bonus_freq[i] for i in range(45)]
    else:
        freqs = normal_freq

    numbers = get_weighted_unique_numbers(freqs, method)
    result_var.set("🎯 추첨 결과: " + ", ".join(map(str, numbers)))
    if latest_round:
        latest_round_var.set(f"최신 분석 회차: {latest_round}회")

def analyze_and_save():
    files = filedialog.askopenfilenames(title="Excel 파일 선택", filetypes=[("엑셀 파일", "*.xlsx")])
    if not files:
        return

    all_numbers = []
    all_bonus = []
    latest_round = 0

    for file in files:
        try:
            nums, bonus, round_num = extract_numbers_from_file(file)
            all_numbers.extend(nums)
            all_bonus.extend(bonus)
            if round_num > latest_round:
                latest_round = round_num
        except Exception as e:
            messagebox.showerror("에러", f"{file} 처리 중 오류 발생:\n{str(e)}")
            return

    normal_freq, bonus_freq = calculate_frequencies(all_numbers, all_bonus)
    save_frequencies(normal_freq, bonus_freq, latest_round)
    update_freq_status()
    latest_round_var.set(f"최신 분석 회차: {latest_round}회")
    messagebox.showinfo("완료", "빈도수 저장 완료")

def get_freq_file_age_text():
    filename = get_hidden_filename()
    if not os.path.exists(filename):
        return "빈도수 파일 없음", "red"

    last_modified = datetime.datetime.fromtimestamp(os.path.getmtime(filename))
    days_ago = (datetime.datetime.now() - last_modified).days
    days_ago = max(days_ago, 0)  # 음수 방지

    if days_ago >= 30:
        return f"빈도수 파일이 오래되었습니다 ({days_ago}일 전)", "red"
    else:
        return f"빈도수 최신화: {days_ago}일 전", "green"

def update_freq_status():
    text, color = get_freq_file_age_text()
    freq_status_label.config(text=text, fg=color)

def open_manual_entry_popup():
    popup = tk.Toplevel(root)
    popup.title("최근 회차 번호 등록")
    popup.geometry("350x150")
    popup.resizable(False, False)

    entries = []

    def validate_number(P):
        if P == "":
            return True
        if P.isdigit():
            val = int(P)
            return 1 <= val <= 45
        return False

    vcmd = (popup.register(validate_number), '%P')

    lbl_nums = tk.Label(popup, text="1 ~ 6번 번호 입력 (각각 1~45)")
    lbl_nums.pack(pady=(10,5))

    entry_frame = tk.Frame(popup)
    entry_frame.pack()

    for i in range(6):
        e = tk.Entry(entry_frame, width=3, justify='center', validate='key', validatecommand=vcmd)
        e.grid(row=0, column=i, padx=4)
        entries.append(e)

    lbl_bonus = tk.Label(popup, text="보너스 번호 입력 (1~45)")
    lbl_bonus.pack(pady=(10,5))

    bonus_entry = tk.Entry(popup, width=3, justify='center', validate='key', validatecommand=vcmd)
    bonus_entry.pack()

    def on_enter(event=None):
        try:
            nums = [int(e.get()) for e in entries]
            bonus = int(bonus_entry.get())

            if len(set(nums)) != 6:
                messagebox.showerror("입력 오류", "1~6번 번호는 중복되지 않아야 합니다.")
                return
            if bonus in nums:
                messagebox.showerror("입력 오류", "보너스 번호는 1~6번 번호와 중복될 수 없습니다.")
                return

            normal_freq, bonus_freq, latest_round = load_frequencies()
            if normal_freq is None or bonus_freq is None:
                normal_freq = [0]*45
                bonus_freq = [0]*45
                latest_round = 0

            new_round = latest_round + 1

            for n in nums:
                normal_freq[n-1] += 1
            bonus_freq[bonus-1] += 1

            save_frequencies(normal_freq, bonus_freq, new_round)
            latest_round_var.set(f"최신 분석 회차: {new_round}회")
            update_freq_status()

            messagebox.showinfo("성공", f"번호가 등록되었습니다. 회차는 {new_round}회로 업데이트 되었습니다.")
            popup.destroy()
        except ValueError:
            messagebox.showerror("입력 오류", "모든 칸에 숫자를 정확히 입력하세요.")

    popup.bind("<Return>", on_enter)
    entries[0].focus_set()

    btn = tk.Button(popup, text="등록", command=on_enter)
    btn.pack(pady=5)

def load_latest_round_on_start():
    _, _, latest_round = load_frequencies()
    if latest_round:
        latest_round_var.set(f"최신 분석 회차: {latest_round}회")
    else:
        latest_round_var.set("최신 분석 회차: 없음")

# --- GUI 구성 ---
root = tk.Tk()
root.title("로또 분석 및 추첨기")
root.geometry("380x280")
root.resizable(False, False)

freq_status_label = tk.Label(root, font=("Arial", 10))
freq_status_label.pack(pady=5)
update_freq_status()

latest_round_var = tk.StringVar(value="최신 분석 회차: 없음")
tk.Label(root, textvariable=latest_round_var, font=("Arial", 10), fg="blue").pack(pady=2)

load_latest_round_on_start()

btn_frame = tk.Frame(root)
btn_frame.pack(pady=5)

tk.Button(btn_frame, text="📊 Excel 분석 및 저장", font=("Arial", 12), command=analyze_and_save).grid(row=0, column=0, padx=10)
tk.Button(btn_frame, text="📝 최근 회차 번호 등록", font=("Arial", 12), command=open_manual_entry_popup).grid(row=0, column=1, padx=10)

include_bonus_var = tk.BooleanVar(value=True)
tk.Checkbutton(root, text="보너스 번호 빈도 포함", variable=include_bonus_var).pack()

method_var = tk.StringVar(value="hardware")
tk.Radiobutton(root, text="의사 난수", variable=method_var, value="pseudo").pack()
tk.Radiobutton(root, text="하드웨어 난수 (기본)", variable=method_var, value="hardware").pack()

tk.Button(root, text="✨ 번호 추첨하기", font=("Arial", 12), command=generate_numbers).pack(pady=10)

result_var = tk.StringVar()
tk.Label(root, textvariable=result_var, font=("Arial", 14), anchor="center").pack(pady=10)

root.mainloop()
