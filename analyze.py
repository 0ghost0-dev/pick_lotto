import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import numpy as np
from collections import Counter
import os


class LottoAnalyzer:
    def __init__(self, root):
        self.root = root
        self.root.title("로또 번호 빈도 분석기")
        self.root.geometry("800x600")
        
        self.files = []
        self.analysis_results = None
        
        self.setup_ui()
    
    def setup_ui(self):
        # 메인 프레임
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 파일 선택 섹션
        file_frame = ttk.LabelFrame(main_frame, text="파일 선택", padding="10")
        file_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        ttk.Button(file_frame, text="Excel 파일 선택", command=self.select_files).grid(row=0, column=0, padx=(0, 10))
        ttk.Button(file_frame, text="선택 파일 초기화", command=self.clear_files).grid(row=0, column=1)
        
        # 선택된 파일 목록
        self.file_listbox = tk.Listbox(file_frame, height=5)
        self.file_listbox.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(10, 0))
        
        # 분석 버튼
        ttk.Button(main_frame, text="분석 시작", command=self.analyze_files, 
                  style="Accent.TButton").grid(row=1, column=0, columnspan=2, pady=10)
        
        # 결과 표시 섹션
        result_frame = ttk.LabelFrame(main_frame, text="분석 결과", padding="10")
        result_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))

        ttk.Label(result_frame, text="빈도수 배열 (1~45번):").grid(row=0, column=0, sticky=tk.W, pady=(0, 5))
        self.frequency_text = tk.Text(result_frame, height=3, wrap=tk.WORD)
        self.frequency_text.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))

        ttk.Label(result_frame, text="보너스 번호 빈도수 배열 (1~45번):").grid(row=2, column=0, sticky=tk.W, pady=(0, 5))
        self.bonus_text = tk.Text(result_frame, height=3, wrap=tk.WORD)
        self.bonus_text.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # 전체 번호 빈도 (일반 + 보너스)
        ttk.Label(result_frame, text="전체 번호 빈도수 배열 (일반+보너스, 1~45번):").grid(row=4, column=0, sticky=tk.W, pady=(0, 5))
        self.total_text = tk.Text(result_frame, height=3, wrap=tk.WORD)
        self.total_text.grid(row=5, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # 상세 분석 결과 트리뷰
        ttk.Label(result_frame, text="상세 분석 결과:").grid(row=6, column=0, sticky=tk.W, pady=(0, 5))
        
        columns = ("번호", "일반 빈도", "보너스 빈도", "총 빈도")
        self.tree = ttk.Treeview(result_frame, columns=columns, show="headings", height=10)
        
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=100, anchor=tk.CENTER)
        
        # 스크롤바 추가
        scrollbar = ttk.Scrollbar(result_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        self.tree.grid(row=7, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=7, column=1, sticky=(tk.N, tk.S))
        
        # 내보내기 버튼
        ttk.Button(result_frame, text="결과 저장", command=self.save_results).grid(row=8, column=0, pady=(10, 0))
        
        # 그리드 가중치 설정
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(2, weight=1)
        result_frame.columnconfigure(0, weight=1)
        result_frame.rowconfigure(7, weight=1)
        file_frame.columnconfigure(0, weight=1)
    
    def select_files(self):
        files = filedialog.askopenfilenames(
            title="Excel 파일 선택",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        
        if files:
            self.files.extend(files)
            self.update_file_list()
    
    def clear_files(self):
        self.files = []
        self.update_file_list()
    
    def update_file_list(self):
        self.file_listbox.delete(0, tk.END)
        for file in self.files:
            filename = os.path.basename(file)
            self.file_listbox.insert(tk.END, filename)
    
    def analyze_files(self):
        if not self.files:
            messagebox.showwarning("경고", "분석할 파일을 선택해주세요.")
            return
        
        try:
            # 모든 당첨번호와 보너스번호 수집
            all_numbers = []
            all_bonus = []
            
            for file_path in self.files:
                numbers, bonus = self.extract_numbers_from_file(file_path)
                all_numbers.extend(numbers)
                all_bonus.extend(bonus)
            
            if not all_numbers:
                messagebox.showerror("오류", "당첨번호 데이터를 찾을 수 없습니다.")
                return
            
            # 빈도 분석
            self.analysis_results = self.calculate_frequencies(all_numbers, all_bonus)
            self.display_results()
            
            messagebox.showinfo("완료", f"총 {len(all_numbers)}개 회차의 데이터를 분석했습니다.")
            
        except Exception as e:
            messagebox.showerror("오류", f"분석 중 오류가 발생했습니다:\n{str(e)}")
    
    def extract_numbers_from_file(self, file_path):
        """Excel 파일에서 당첨번호와 보너스번호 추출"""
        numbers = []
        bonus_numbers = []
        
        try:
            # Excel 파일 읽기
            df = pd.read_excel(file_path, header=None)
            
            # 데이터가 있는 행들 찾기 (숫자 데이터가 7개 이상 있는 행)
            for idx, row in df.iterrows():
                # 행의 마지막 부분에서 숫자 찾기
                row_values = row.dropna().tolist()
                
                # 숫자만 필터링
                numeric_values = []
                for val in row_values:
                    try:
                        if isinstance(val, (int, float)) and 1 <= val <= 45:
                            numeric_values.append(int(val))
                    except:
                        continue
                
                # 연속된 7개의 숫자가 있으면 당첨번호로 판단
                if len(numeric_values) >= 7:
                    # 마지막 7개를 당첨번호로 간주
                    winning_nums = numeric_values[-7:]
                    numbers.append(winning_nums[:6])  # 1~6번 번호
                    bonus_numbers.append(winning_nums[6])  # 보너스 번호
        
        except Exception as e:
            print(f"파일 {file_path} 처리 중 오류: {e}")
        
        return numbers, bonus_numbers
    
    def calculate_frequencies(self, all_numbers, all_bonus):
        """빈도수 계산"""
        # 1~45번까지 각 번호의 빈도수 초기화
        normal_freq = [0] * 45
        bonus_freq = [0] * 45
        
        # 일반 번호 빈도수 계산
        for draw_numbers in all_numbers:
            for num in draw_numbers:
                if 1 <= num <= 45:
                    normal_freq[num - 1] += 1
        
        # 보너스 번호 빈도수 계산
        for bonus_num in all_bonus:
            if 1 <= bonus_num <= 45:
                bonus_freq[bonus_num - 1] += 1
        
        return {
            'normal_freq': normal_freq,
            'bonus_freq': bonus_freq,
            'total_freq': [normal_freq[i] + bonus_freq[i] for i in range(45)],
            'total_draws': len(all_numbers)
        }
    
    def display_results(self):
        """결과 표시"""
        if not self.analysis_results:
            return
        
        normal_freq = self.analysis_results['normal_freq']
        bonus_freq = self.analysis_results['bonus_freq']
        total_freq = self.analysis_results['total_freq']

        self.frequency_text.delete(1.0, tk.END)
        freq_str = str(normal_freq)
        self.frequency_text.insert(tk.END, freq_str)

        self.bonus_text.delete(1.0, tk.END)
        bonus_str = str(bonus_freq)
        self.bonus_text.insert(tk.END, bonus_str)
        
        # 전체 빈도수 배열 텍스트 표시
        self.total_text.delete(1.0, tk.END)
        total_str = str(total_freq)
        self.total_text.insert(tk.END, total_str)
        
        # 트리뷰에 상세 결과 표시
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        for i in range(45):
            number = i + 1
            normal_count = normal_freq[i]
            bonus_count = bonus_freq[i]
            total_count = normal_count + bonus_count
            
            self.tree.insert("", "end", values=(number, normal_count, bonus_count, total_count))
    
    def save_results(self):
        """결과를 파일로 저장"""
        if not self.analysis_results:
            messagebox.showwarning("경고", "저장할 분석 결과가 없습니다.")
            return
        
        file_path = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Text files", "*.txt"), ("CSV files", "*.csv"), ("All files", "*.*")]
        )
        
        if file_path:
            try:
                normal_freq = self.analysis_results['normal_freq']
                bonus_freq = self.analysis_results['bonus_freq']
                total_freq = self.analysis_results['total_freq']
                
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write("로또 번호 빈도 분석 결과\n")
                    f.write("=" * 50 + "\n\n")
                    f.write(f"총 분석 회차: {self.analysis_results['total_draws']}회\n\n")
                    
                    f.write("일반 번호 빈도수 배열 (1~45번):\n")
                    f.write(str(normal_freq) + "\n\n")
                    
                    f.write("보너스 번호 빈도수 배열 (1~45번):\n")
                    f.write(str(bonus_freq) + "\n\n")
                    
                    f.write("전체 번호 빈도수 배열 (일반+보너스, 1~45번):\n")
                    f.write(str(total_freq) + "\n\n")
                    
                    f.write("상세 분석 결과:\n")
                    f.write("번호\t일반빈도\t보너스빈도\t총빈도\n")
                    f.write("-" * 40 + "\n")
                    
                    for i in range(45):
                        number = i + 1
                        normal_count = normal_freq[i]
                        bonus_count = bonus_freq[i]
                        total_count = normal_count + bonus_count
                        f.write(f"{number}\t{normal_count}\t\t{bonus_count}\t\t{total_count}\n")
                
                messagebox.showinfo("완료", f"결과가 저장되었습니다:\n{file_path}")
                
            except Exception as e:
                messagebox.showerror("오류", f"저장 중 오류가 발생했습니다:\n{str(e)}")


def main():
    root = tk.Tk()
    app = LottoAnalyzer(root)
    root.mainloop()


if __name__ == "__main__":
    main()