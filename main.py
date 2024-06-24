from collections import Counter
import re
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

def read_text_file(file_path):
    with open(file_path, 'r', encoding='utf-8') as file:
        text = file.read()
    return text

def read_settings_file(settings_file_path):
    patterns = []
    with open(settings_file_path, 'r', encoding='utf-8') as file:
        patterns.extend(line.strip() for line in file if line.strip())
    return patterns

def count_words(text, patterns):
    pattern_counts = Counter()

    for pattern in patterns:
        escaped_pattern = re.escape(pattern).replace(r'\~', '.*?')
        matches = re.findall(escaped_pattern, text, re.IGNORECASE | re.DOTALL)
        pattern_counts[pattern] = len(matches)
        text = re.sub(escaped_pattern, '', text, flags=re.IGNORECASE | re.DOTALL)

    words = re.findall(r'\b\w+\b', text)
    words = [word.lower() for word in words if not re.search(r'[ㄱ-ㅎㅏ-ㅣ가-힣0-9]', word)]
    word_counts = Counter(words)

    word_counts.update(pattern_counts)

    return word_counts

def save_word_counts_to_excel(word_counts, output_file_path):
    df = pd.DataFrame(word_counts.items(), columns=['Word', 'Count'])
    df.to_excel(output_file_path, index=False)

def process_files(text_file_path, settings_file_path=None):
    text = read_text_file(text_file_path)
    patterns = read_settings_file(settings_file_path) if settings_file_path else []
    word_counts = count_words(text, patterns)
    output_file_path = re.sub(r'\.txt$', '.xlsx', text_file_path, flags=re.IGNORECASE)
    save_word_counts_to_excel(word_counts, output_file_path)
    return output_file_path

def select_text_file():
    return filedialog.askopenfilename(
        title="텍스트 파일을 선택하세요",
        filetypes=(("텍스트 파일", "*.txt"), ("모든 파일", "*.*"))
    )

def select_settings_file():
    return filedialog.askopenfilename(
        title="설정 파일(settings.txt)을 선택하세요 (선택 사항)",
        filetypes=(("텍스트 파일", "*.txt"), ("모든 파일", "*.*"))
    )

def select_files():
    text_file_path = select_text_file()
    if text_file_path:
        settings_file_path = select_settings_file()
        try:
            process_files(text_file_path, settings_file_path)
            messagebox.showinfo("완료", "단어 계수 결과를 파일에 저장했습니다.")
        except Exception as e:
            messagebox.showerror("오류", f"파일 처리 중 오류가 발생했습니다: {e}")
            try:
                with open("error_log.txt", "w", encoding="utf-8") as log_file:
                    log_file.write(f"파일 처리 중 오류가 발생했습니다: {e}\n")
            except Exception as log_error:
                messagebox.showerror("오류", f"오류 로그를 기록하는 중 문제가 발생했습니다: {log_error}")

def create_gui():
    root = tk.Tk()
    root.title("단어 계수기")

    frame = tk.Frame(root, padx=10, pady=10)
    frame.pack(padx=10, pady=10)

    label1 = tk.Label(frame, text="텍스트 파일을 선택하세요")
    label1.pack(pady=5)

    button1 = tk.Button(frame, text="파일 선택", command=select_files)
    button1.pack(pady=5)

    root.mainloop()

if __name__ == "__main__":
    create_gui()
