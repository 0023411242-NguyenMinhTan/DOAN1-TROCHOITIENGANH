import tkinter as tk
from tkinter import messagebox, ttk, filedialog
import csv
import os
import json
from datetime import datetime
import random
import docx  # Thư viện mới để đọc file Word

# ---------------------------------------------------------
#Tên người Viết: Nguyễn Minh Tân - 0023411242
# Dự án: Trò chơi tiếng anh cho học sinh tiểu học - Tích hợp đọc file Word
# Đơn vị: Lớp: ĐHSTIN23A - Trường Đại học Đồng Tháp
# Mục tiêu: Tạo một trò chơi trắc nghiệm tiếng Anh cho học sinh tiểu học với giao diện thân thiện, đồng thời cung cấp công cụ quản lý câu hỏi và thống kê kết quả cho giáo viên.
# ---------------------------------------------------------

DATA_FILE = "cau_hoi_tieng_anh.json"
STATS_FILE = "thong_ke_diem_so.csv"

DEFAULT_DATA = {
    "Lớp 3": {"time_per_q": 10, "questions": []}, # Mỗi câu 10 giây
    "Lớp 4": {"time_per_q": 8, "questions": []},  # Mỗi câu 8 giây
    "Lớp 5": {"time_per_q": 6, "questions": []}   # Mỗi câu 6 giây
}

class EnglishGameProWord:
    def __init__(self, root):
        self.root = root
        self.root.title("Trò chơi tiếng anh cho học sinh tiểu học - Trường Đại học Đồng Tháp")
        self.root.geometry("800x600")
        
        self.color_bg = "#E0F7FA"
        self.color_btn = "#FFF59D"
        self.font_title = ("Helvetica", 28, "bold")
        self.font_large = ("Helvetica", 20, "bold")
        self.font_normal = ("Helvetica", 16)
        
        self.root.configure(bg=self.color_bg)
        self.load_data()
        
        self.player_name = ""
        self.current_level = ""
        self.score = 0
        self.timer_id = None
        self.time_left = 0
        self.current_q_idx = 0
        
        self.create_start_screen()

    def load_data(self):
        if not os.path.exists(DATA_FILE):
            with open(DATA_FILE, "w", encoding="utf-8") as f:
                json.dump(DEFAULT_DATA, f, ensure_ascii=False, indent=4)
            self.level_data = DEFAULT_DATA
        else:
            with open(DATA_FILE, "r", encoding="utf-8") as f:
                self.level_data = json.load(f)
            
            
            updated = False
            for level in ["Lớp 3", "Lớp 4", "Lớp 5"]:
                if level in self.level_data and "time_per_q" not in self.level_data[level]:
                    # Phân bổ thời gian theo khối lớp nếu chưa có
                    if level == "Lớp 3": self.level_data[level]["time_per_q"] = 10
                    elif level == "Lớp 4": self.level_data[level]["time_per_q"] = 8
                    else: self.level_data[level]["time_per_q"] = 6
                    updated = True
            
            if updated:
                self.save_data()

    def save_data(self):
        with open(DATA_FILE, "w", encoding="utf-8") as f:
            json.dump(self.level_data, f, ensure_ascii=False, indent=4)

    # ================== GIAO DIỆN CHÍNH & GAME ==================

    def create_start_screen(self):
        self.clear_window()
        tk.Label(self.root, text="CHỌN CẤP ĐỘ ĐỂ BẮT ĐẦU", font=self.font_title, bg=self.color_bg, fg="#0277BD").pack(pady=40)
        
        tk.Label(self.root, text="Nhập tên của em:", font=self.font_large, bg=self.color_bg).pack(pady=10)
        self.entry_name = tk.Entry(self.root, font=self.font_large, justify="center", width=20)
        self.entry_name.pack(pady=10)
        
        tk.Label(self.root, text="Em đang học lớp mấy?", font=self.font_normal, bg=self.color_bg).pack(pady=20)
        
        btn_frame = tk.Frame(self.root, bg=self.color_bg)
        btn_frame.pack(pady=10)
        
        for level in ["Lớp 3", "Lớp 4", "Lớp 5"]:
            tk.Button(btn_frame, text=level, font=self.font_large, bg="#FFCA28", 
                      cursor="hand2", width=10, command=lambda l=level: self.start_game(l)).pack(side=tk.LEFT, padx=15)
            
       
        teacher_frame = tk.Frame(self.root, bg=self.color_bg)
        teacher_frame.pack(side=tk.BOTTOM, pady=30)
        
        tk.Button(teacher_frame, text="👩‍🏫 Mở Giao Diện Quản Lý (Giáo Viên)", font=("Helvetica", 14, "bold"), 
                  bg="#FFCCBC", command=self.open_teacher_dashboard, cursor="hand2").pack(pady=10)
        
    def start_game(self, selected_level):
        name_input = self.entry_name.get().strip()
        if not name_input:
            messagebox.showwarning("Cảnh báo", "Em nhớ nhập tên nhé!")
            return
            
        self.load_data() 
        num_questions = len(self.level_data[selected_level]["questions"])
        
        if num_questions == 0:
            messagebox.showinfo("Thông báo", "Giáo viên chưa cập nhật câu hỏi cho khối lớp này!")
            return

        self.player_name = name_input
        self.current_level = selected_level
        self.score = 0
        self.current_q_idx = 0
        
        # --- TÍNH TOÁN THỜI GIAN LINH HOẠT ---
        # Lấy thời gian mỗi câu của khối lớp (mặc định 10s nếu lỗi)
        time_per_q = self.level_data[selected_level].get("time_per_q", 10) 
        # Tổng thời gian = Số câu hỏi * Thời gian 1 câu
        self.time_left = num_questions * time_per_q 
        
        self.create_game_screen()
        self.load_question()
        self.update_timer()

    def create_game_screen(self):
        self.clear_window()
        top_frame = tk.Frame(self.root, bg=self.color_bg)
        top_frame.pack(fill=tk.X, padx=20, pady=15)
        
        self.lbl_timer = tk.Label(top_frame, text=f"⏳ {self.time_left}s", font=self.font_normal, bg=self.color_bg, fg="red")
        self.lbl_timer.pack(side=tk.LEFT)
        self.lbl_score = tk.Label(top_frame, text=f"⭐ Điểm: {self.score}", font=self.font_normal, bg=self.color_bg, fg="#0277BD")
        self.lbl_score.pack(side=tk.RIGHT)
        
        self.lbl_question = tk.Label(self.root, text="", font=self.font_large, bg=self.color_bg, wraplength=700, justify="center")
        self.lbl_question.pack(pady=40)
        
        self.btn_frame = tk.Frame(self.root, bg=self.color_bg)
        self.btn_frame.pack(pady=20)
        
        self.option_buttons = []
        for i in range(4):
            btn = tk.Button(self.btn_frame, text="", font=self.font_normal, bg=self.color_btn, width=25, height=2,
                            cursor="hand2", command=lambda idx=i: self.check_answer(idx))
            btn.grid(row=i//2, column=i%2, padx=15, pady=15)
            self.option_buttons.append(btn)

    def load_question(self):
        questions = self.level_data[self.current_level]["questions"]
        if self.current_q_idx < len(questions):
            q_data = questions[self.current_q_idx]
            self.lbl_question.config(text=f"Câu {self.current_q_idx + 1}: {q_data['q']}")
            for i in range(4):
                self.option_buttons[i].config(text=q_data['options'][i], bg=self.color_btn, state=tk.NORMAL)
        else:
            self.end_game("Hoàn thành bài thi!")

    def check_answer(self, selected_index):
        if self.timer_id: self.root.after_cancel(self.timer_id)
        for btn in self.option_buttons: btn.config(state=tk.DISABLED)
            
        q_data = self.level_data[self.current_level]["questions"][self.current_q_idx]
        correct_text = q_data['ans']
        
        for btn in self.option_buttons:
            if btn['text'] == correct_text: btn.config(bg="#A5D6A7")
                
        if q_data['options'][selected_index] == correct_text:
            self.score += 10
            self.lbl_score.config(text=f"⭐ Điểm: {self.score}")
        else:
            self.option_buttons[selected_index].config(bg="#EF9A9A")
            
        self.current_q_idx += 1
        self.root.after(1500, self.resume_after_answer)

    def resume_after_answer(self):
        self.update_timer()
        self.load_question()

    def update_timer(self):
        if self.time_left > 0:
            self.time_left -= 1
            self.lbl_timer.config(text=f"⏳ {self.time_left}s")
            self.timer_id = self.root.after(1000, self.update_timer)
        else:
            self.end_game("Hết thời gian!")

    def end_game(self, message):
        if self.timer_id: self.root.after_cancel(self.timer_id)
        self.save_results_to_csv()
        self.clear_window()
        
        tk.Label(self.root, text=message, font=self.font_title, bg=self.color_bg, fg="#FF8F00").pack(pady=40)
        tk.Label(self.root, text=f"{self.player_name} đạt:\n\n{self.score} Điểm", 
                 font=("Helvetica", 36, "bold"), bg=self.color_bg, fg="#0277BD").pack(pady=30)
        tk.Button(self.root, text="QUAY LẠI", font=self.font_large, bg="#FFCA28",
                  command=self.create_start_screen, width=20).pack(pady=40)

    # ================== TÍNH NĂNG GIÁO VIÊN & ĐỌC WORD ==================

    def save_results_to_csv(self):
        file_exists = os.path.isfile(STATS_FILE)
        play_time = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        try:
            with open(STATS_FILE, mode='a', newline='', encoding='utf-8-sig') as f:
                writer = csv.writer(f)
                if not file_exists: writer.writerow(["Thời gian", "Tên học sinh", "Khối Lớp", "Điểm số"])
                writer.writerow([play_time, self.player_name, self.current_level, self.score])
        except Exception: pass

# ================== GIAO DIỆN QUẢN LÝ TỔNG HỢP (DÀNH CHO GIÁO VIÊN) ==================

    def open_teacher_dashboard(self):
        self.mgr_win = tk.Toplevel(self.root)
        self.mgr_win.title("Bảng Điều Khiển Của Giáo Viên 👩‍🏫")
        self.mgr_win.geometry("900x700")
        self.mgr_win.configure(bg="#F0F4F8")

        # Tùy chỉnh phong cách Treeview
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Treeview.Heading", font=("Helvetica", 11, "bold"), background="#90CAF9", foreground="#0D47A1")
        style.configure("Treeview", font=("Helvetica", 11), rowheight=28)

        # Tiêu đề
        tk.Label(self.mgr_win, text="HỆ THỐNG QUẢN LÝ TRÒ CHƠI TIẾNG ANH", font=("Helvetica", 18, "bold"), bg="#F0F4F8", fg="#1565C0").pack(pady=10)

        # Tạo Notebook (Hệ thống thẻ Tab)
        self.notebook = ttk.Notebook(self.mgr_win)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)

        # TẠO TAB 1: QUẢN LÝ CÂU HỎI
        self.tab_q = tk.Frame(self.notebook, bg="white")
        self.notebook.add(self.tab_q, text="📚 Ngân Hàng Câu Hỏi")
        self.build_question_manager()

        # TẠO TAB 2: THỐNG KÊ ĐIỂM SỐ
        self.tab_scores = tk.Frame(self.notebook, bg="white")
        self.notebook.add(self.tab_scores, text="📊 Điểm Số Học Sinh")
        self.build_score_statistics()

    def build_question_manager(self):
        # --- Khung chọn lớp & nút nhập Word ---
        top_frame = tk.Frame(self.tab_q, bg="white")
        top_frame.pack(fill=tk.X, pady=10, padx=10)
        
        tk.Label(top_frame, text="Chọn Khối Lớp:", bg="white", font=self.font_normal, fg="#333333").pack(side=tk.LEFT)
        self.combo_level = ttk.Combobox(top_frame, values=["Lớp 3", "Lớp 4", "Lớp 5"], state="readonly", font=self.font_normal, width=10)
        self.combo_level.set("Lớp 3")
        self.combo_level.pack(side=tk.LEFT, padx=10)
        self.combo_level.bind("<<ComboboxSelected>>", self.refresh_treeview_q)

        tk.Button(top_frame, text="📄 Nhập dữ liệu từ Word", bg="#BA68C8", fg="white", font=("Helvetica", 10, "bold"), 
                  command=self.import_from_word, cursor="hand2").pack(side=tk.RIGHT, padx=10)

        # --- Treeview hiển thị danh sách câu hỏi ---
        columns = ("ID", "Question", "CorrectAns")
        self.tree_q = ttk.Treeview(self.tab_q, columns=columns, show="headings", height=8)
        
        self.tree_q.heading("ID", text="STT")
        self.tree_q.column("ID", width=50, anchor="center")
        self.tree_q.heading("Question", text="Nội Dung Câu Hỏi")
        self.tree_q.column("Question", width=500, anchor="w")
        self.tree_q.heading("CorrectAns", text="Đáp Án Đúng")
        self.tree_q.column("CorrectAns", width=200, anchor="center")
        
        self.tree_q.pack(fill=tk.X, padx=10, pady=5)
        self.tree_q.bind('<<TreeviewSelect>>', self.on_treeview_select)

        # --- Khung nhập liệu (Form) ---
        form_frame = tk.LabelFrame(self.tab_q, text="Chi Tiết Câu Hỏi", bg="white", font=("Helvetica", 12, "bold"), fg="#1565C0")
        form_frame.pack(fill=tk.X, padx=10, pady=10, ipadx=10, ipady=10)
        
        tk.Label(form_frame, text="Câu hỏi:", bg="white", font=self.font_normal).grid(row=0, column=0, sticky="e", pady=5)
        self.entry_q = tk.Entry(form_frame, width=65, font=("Helvetica", 12))
        self.entry_q.grid(row=0, column=1, columnspan=3, pady=5, sticky="w")
        
        self.entries_opt = []
        for i in range(4):
            tk.Label(form_frame, text=f"Đáp án {i+1}:", bg="white", font=self.font_normal).grid(row=1+i//2, column=(i%2)*2, sticky="e", pady=5, padx=5)
            e = tk.Entry(form_frame, width=25, font=("Helvetica", 12))
            e.grid(row=1+i//2, column=(i%2)*2+1, padx=5, pady=5)
            self.entries_opt.append(e)
            
        tk.Label(form_frame, text="Đáp án đúng:", bg="white", font=self.font_normal, fg="#D84315").grid(row=3, column=0, columnspan=2, sticky="e", pady=10)
        self.entry_ans = tk.Entry(form_frame, width=25, font=("Helvetica", 12, "bold"))
        self.entry_ans.grid(row=3, column=2, columnspan=2, sticky="w", pady=10)

        # --- Các nút thao tác ---
        btn_frame = tk.Frame(self.tab_q, bg="white")
        btn_frame.pack(pady=10)
        
        tk.Button(btn_frame, text="➕ Thêm Mới", bg="#81C784", font=("Helvetica", 12, "bold"), width=12, command=self.add_q).pack(side=tk.LEFT, padx=10)
        tk.Button(btn_frame, text="🔄 Cập Nhật", bg="#64B5F6", font=("Helvetica", 12, "bold"), width=12, command=self.update_q).pack(side=tk.LEFT, padx=10)
        tk.Button(btn_frame, text="🗑️ Xóa", bg="#E57373", font=("Helvetica", 12, "bold"), width=12, command=self.delete_q).pack(side=tk.LEFT, padx=10)
        tk.Button(btn_frame, text="🧹 Làm Mới", bg="#E0E0E0", font=("Helvetica", 12), width=12, command=self.clear_form).pack(side=tk.LEFT, padx=10)

        self.refresh_treeview_q()

    def build_score_statistics(self):
        # --- Treeview hiển thị CSV Điểm số ---
        columns = ("Time", "Name", "Level", "Score")
        self.tree_scores = ttk.Treeview(self.tab_scores, columns=columns, show="headings", height=20)
        
        self.tree_scores.heading("Time", text="Thời Gian")
        self.tree_scores.column("Time", width=200, anchor="center")
        self.tree_scores.heading("Name", text="Tên Học Sinh")
        self.tree_scores.column("Name", width=250, anchor="w")
        self.tree_scores.heading("Level", text="Khối Lớp")
        self.tree_scores.column("Level", width=150, anchor="center")
        self.tree_scores.heading("Score", text="Điểm Số")
        self.tree_scores.column("Score", width=150, anchor="center")

        self.tree_scores.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        # Tải dữ liệu từ CSV lên bảng
        if os.path.isfile(STATS_FILE):
            try:
                with open(STATS_FILE, 'r', encoding='utf-8-sig') as f:
                    reader = csv.reader(f)
                    next(reader, None) # Bỏ qua dòng tiêu đề
                    for row in reader:
                        # Bôi màu dựa trên điểm số (Ví dụ: Trên 50 điểm thì màu xanh)
                        tag = "high" if int(row[3]) >= 50 else "low"
                        self.tree_scores.insert("", tk.END, values=row, tags=(tag,))
                
                self.tree_scores.tag_configure("high", background="#E8F5E9")
                self.tree_scores.tag_configure("low", background="#FFEBEE")
            except Exception as e:
                pass

    # --- Các hàm bổ trợ cho Treeview Câu hỏi ---
    def refresh_treeview_q(self, event=None):
        for item in self.tree_q.get_children():
            self.tree_q.delete(item)
            
        level = self.combo_level.get()
        for idx, q_data in enumerate(self.level_data[level]["questions"]):
            self.tree_q.insert("", tk.END, iid=str(idx), values=(idx + 1, q_data['q'], q_data['ans']))
        self.clear_form()

    def on_treeview_select(self, event):
        selected = self.tree_q.selection()
        if not selected: return
        
        idx = int(selected[0]) # ID của dòng (được gán bằng index trong list)
        level = self.combo_level.get()
        q_data = self.level_data[level]["questions"][idx]
        
        self.clear_form()
        self.entry_q.insert(0, q_data["q"])
        for i in range(4): self.entries_opt[i].insert(0, q_data["options"][i])
        self.entry_ans.insert(0, q_data["ans"])
    # ================== CÁC HÀM XỬ LÝ DỮ LIỆU CÂU HỎI ==================

    def import_from_word(self):
        file_path = filedialog.askopenfilename(parent=self.mgr_win, title="Chọn file Word", filetypes=[("Word Documents", "*.docx")])
        if not file_path: return
        
        # --- Hỏi người dùng cách nhập dữ liệu ---
        replace_all = messagebox.askyesno(
            "Tùy chọn nhập dữ liệu", 
            "Bạn có muốn XÓA SẠCH câu hỏi cũ của khối lớp trong file Word để làm mới hoàn toàn không?\n\n- Chọn 'Yes' để Thay thế hoàn toàn.\n- Chọn 'No' để Thêm nối " \
            "tiếp vào cuối danh sách.",
            parent=self.mgr_win
        )
        try:
            doc = docx.Document(file_path)
            current_level = None
            q_data = {}
            options = []
            count = 0
            
            # Biến cờ để đánh dấu các khối lớp đã được xóa dữ liệu cũ (chỉ xóa 1 lần khi bắt đầu đọc lớp đó)
            cleared_levels = set()
            for para in doc.paragraphs:
                text = para.text.strip()
                if not text: continue
                
                # Nhận diện Lớp
                if text in ["[Lớp 3]", "[Lớp 4]", "[Lớp 5]"]:
                    current_level = text.strip("[]")

                    # Nếu chọn Yes và lớp này chưa được dọn dẹp -> Xóa sạch câu hỏi cũ
                    if replace_all and current_level not in cleared_levels:
                        self.level_data[current_level]["questions"] = []
                        cleared_levels.add(current_level)
                        
                # Nhận diện Câu hỏi 
                elif text.lower().startswith("câu hỏi") or text.lower().startswith("câu"):
                    if ":" in text:
                        q_data["q"] = text.split(":", 1)[1].strip()
                    else:
                        q_data["q"] = text 
                    options = [] 
                    
                # Nhận diện các đáp án A, B, C, D
                elif text.upper().startswith(("A.", "B.", "C.", "D.")):
                    options.append(text.split(".", 1)[1].strip())
                    
                # Nhận diện Đáp án đúng và lưu trữ
                elif text.lower().startswith("đáp án:"):
                    q_data["ans"] = text.split(":", 1)[1].strip()
                    q_data["options"] = options
                    
                    if current_level and "q" in q_data and len(q_data["options"]) == 4 and "ans" in q_data:
                        self.level_data[current_level]["questions"].append(q_data.copy())
                        count += 1
                        
                    q_data = {} 
                    
            if count > 0:
                self.save_data()
                self.refresh_treeview_q()
                
                action_text = "thay thế hoàn toàn" if replace_all else "thêm nối tiếp"
                messagebox.showinfo("Thành công", f"Đã {action_text} {count} câu hỏi từ file Word!", parent=self.mgr_win)
            else:
                messagebox.showwarning("Cảnh báo", "Không tìm thấy câu hỏi nào. Hãy đảm bảo format chuẩn.", parent=self.mgr_win)
                
        except Exception as e:
            messagebox.showerror("Lỗi", f"Không thể đọc file Word.\nChi tiết lỗi: {e}", parent=self.mgr_win)

    def get_form_data(self):
        q = self.entry_q.get().strip()
        opts = [e.get().strip() for e in self.entries_opt]
        ans = self.entry_ans.get().strip()
        if not q or not all(opts) or not ans: 
            messagebox.showwarning("Thiếu thông tin", "Vui lòng nhập đầy đủ câu hỏi và đáp án!", parent=self.mgr_win)
            return None
        return {"q": q, "options": opts, "ans": ans}

    def add_q(self):
        data = self.get_form_data()
        if data:
            self.level_data[self.combo_level.get()]["questions"].append(data)
            self.save_data()
            self.refresh_treeview_q() # SỬA LỖI Ở ĐÂY: Gọi hàm cập nhật Treeview mới
            messagebox.showinfo("Thành công", "Đã thêm câu hỏi mới!", parent=self.mgr_win)

    def clear_form(self):
        self.entry_q.delete(0, tk.END)
        self.entry_ans.delete(0, tk.END)
        for e in self.entries_opt: 
            e.delete(0, tk.END)    

    def update_q(self):
        selected = self.tree_q.selection()
        data = self.get_form_data()
        if selected and data:
            idx = int(selected[0])
            self.level_data[self.combo_level.get()]["questions"][idx] = data
            self.save_data()
            self.refresh_treeview_q()
            messagebox.showinfo("Thành công", "Đã cập nhật câu hỏi!", parent=self.mgr_win)

    def delete_q(self):
        selected = self.tree_q.selection()
        if not selected:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn ít nhất một câu hỏi trên bảng để xóa!", parent=self.mgr_win)
            return

        # Báo cáo số lượng câu hỏi đang được chọn để xóa
        msg = f"Bạn có chắc chắn muốn xóa {len(selected)} câu hỏi đã chọn không?"
        if messagebox.askyesno("Xác nhận xóa", msg, parent=self.mgr_win):
            
            # Lấy danh sách ID (chỉ số) và sắp xếp GIẢM DẦN (reverse=True)
            # Việc này rất quan trọng để khi xóa không bị lệch index của List
            indices_to_delete = sorted([int(item) for item in selected], reverse=True)
            
            for idx in indices_to_delete:
                del self.level_data[self.combo_level.get()]["questions"][idx]
                
            self.save_data()
            self.refresh_treeview_q()
            messagebox.showinfo("Thành công", f"Đã xóa thành công {len(selected)} câu hỏi!", parent=self.mgr_win)

    def clear_window(self):
        for widget in self.root.winfo_children(): widget.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    app = EnglishGameProWord(root)
    root.mainloop()

