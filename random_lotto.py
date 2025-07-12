import tkinter as tk
from tkinter import messagebox
import random
import bisect
import ctypes
import os
import sys
import platform

# --- 가중치 데이터 ---
frequencies = [
    192, 181, 191, 188, 172, 190, 192, 173, 151, 182,
    184, 199, 197, 188, 182, 184, 195, 188, 182, 192,
    183, 157, 159, 188, 165, 189, 198, 169, 163, 181,
    185, 171, 197, 203, 184, 180, 188, 191, 185, 187,
    158, 173, 194, 180, 187
]

# 누적합 계산
def build_weighted_table(weights):
    cumulative = []
    total = 0
    for w in weights:
        total += w
        cumulative.append(total)
    return cumulative, total

cumulative_weights, total_weight = build_weighted_table(frequencies)

# --- 하드웨어 난수 함수 (OS별 처리) ---

def get_hardware_random_bytes(n):
    if os.name == 'nt':  # Windows
        BCRYPT_USE_SYSTEM_PREFERRED_RNG = 0x00000002
        bcrypt = ctypes.windll.bcrypt
        buffer = ctypes.create_string_buffer(n)
        status = bcrypt.BCryptGenRandom(None, buffer, n, BCRYPT_USE_SYSTEM_PREFERRED_RNG)
        if status != 0:
            raise OSError(f"BCryptGenRandom failed with status {status}")
        return buffer.raw
    else:  # Unix/Linux/macOS
        return os.urandom(n)

def get_hardware_random_float():
    rand_bytes = get_hardware_random_bytes(4)
    rand_int = int.from_bytes(rand_bytes, byteorder='little')
    return rand_int / 2**32

# --- 가중치 기반 선택 ---
def weighted_choice_pseudo():
    r = random.uniform(0, total_weight)
    return bisect.bisect_left(cumulative_weights, r) + 1

def weighted_choice_hardware():
    r = get_hardware_random_float() * total_weight
    return bisect.bisect_left(cumulative_weights, r) + 1

def get_weighted_unique_numbers(count=6, method='pseudo'):
    numbers = set()
    while len(numbers) < count:
        if method == 'pseudo':
            num = weighted_choice_pseudo()
        else:
            num = weighted_choice_hardware()
        numbers.add(num)
    return sorted(numbers)

# --- GUI 구성 ---
def generate_numbers():
    try:
        method = method_var.get()
        numbers = get_weighted_unique_numbers(6, method=method)
        result_var.set("🎯 추첨 결과: " + ", ".join(map(str, numbers)))
    except Exception as e:
        messagebox.showerror("에러", str(e))

# GUI 창
root = tk.Tk()
root.title("로또 번호 추첨기 (가중치 기반)")
root.geometry("460x260")
root.resizable(False, False)

# 메인 프레임
main_frame = tk.Frame(root)
main_frame.pack(expand=True)

# 라디오 버튼
radio_frame = tk.LabelFrame(main_frame, text="난수 생성 방식", padx=10, pady=5)
radio_frame.grid(row=0, column=0, pady=10)

method_var = tk.StringVar(value="pseudo")
tk.Radiobutton(radio_frame, text="의사 난수 (Pseudo-Random)", variable=method_var, value="pseudo").pack(anchor="w")
tk.Radiobutton(radio_frame, text="하드웨어 난수 (Hardware RNG)", variable=method_var, value="hardware").pack(anchor="w")

# 버튼
tk.Button(main_frame, text="번호 추첨하기", font=("Arial", 13), command=generate_numbers).grid(row=1, column=0, pady=10)

# 결과 출력
result_var = tk.StringVar()
tk.Label(main_frame, textvariable=result_var, font=("Arial", 14), anchor="center", width=40, height=2).grid(row=2, column=0)

root.mainloop()
