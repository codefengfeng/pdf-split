import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk
import fitz  # PyMuPDF
from PIL import Image, ImageTk
import os
import threading
import re
import numpy as np
import cv2
from paddleocr import PaddleOCR


class CombinedPDFSplitter:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF 智能拆分工具 (OCR自动 + 手动切割)")
        self.root.geometry("1100x780")

        try:
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            pass

        # OCR 引擎（延迟初始化）
        self.ocr_engine = None

        # 手动模式的核心变量
        self.manual_doc = None
        self.current_page = 0
        self.zoom_factor = 1.0
        self.current_image = None
        self.manual_pdf_path = ""

        self.setup_ui()

    def setup_ui(self):
        # ==================== 顶层 Notebook (选项卡) ====================
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # ---------- Tab 1: OCR 智能切分 ----------
        self.ocr_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.ocr_tab, text="  OCR 智能切分  ")
        self.setup_ocr_tab()

        # ---------- Tab 2: 手动预览切割 ----------
        self.manual_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.manual_tab, text="  手动预览切割  ")
        self.setup_manual_tab()

    # ================================================================
    #                    Tab 1: OCR 智能切分 (原 split_pdf_ocr.py)
    # ================================================================

    def setup_ocr_tab(self):
        main_frame = ttk.Frame(self.ocr_tab, padding=20)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 1. 选择 PDF 文件
        ttk.Label(main_frame, text="1. 选择巨型扫描件 PDF:", font=("微软雅黑", 11, "bold")).pack(anchor=tk.W, pady=(0, 5))
        file_frame = ttk.Frame(main_frame)
        file_frame.pack(fill=tk.X, pady=5)
        self.ocr_pdf_path_var = tk.StringVar()
        ttk.Entry(file_frame, textvariable=self.ocr_pdf_path_var).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        ttk.Button(file_frame, text="浏览", command=self.ocr_load_pdf).pack(side=tk.RIGHT)

        # 2. 选择输出目录
        ttk.Label(main_frame, text="2. 选择拆分结果保存的文件夹:", font=("微软雅黑", 11, "bold")).pack(anchor=tk.W, pady=(15, 5))
        dir_frame = ttk.Frame(main_frame)
        dir_frame.pack(fill=tk.X, pady=5)
        self.ocr_out_dir_var = tk.StringVar()
        ttk.Entry(dir_frame, textvariable=self.ocr_out_dir_var).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        ttk.Button(dir_frame, text="选择", command=self.ocr_select_out_dir).pack(side=tk.RIGHT)

        # 3. 输入大标题关键字
        ttk.Label(main_frame, text="3. 输入切分依据 (大标题关键字，每行一个):", font=("微软雅黑", 11, "bold")).pack(anchor=tk.W, pady=(15, 5))
        ttk.Label(main_frame, text='提示: PaddleOCR 精度极高，输入如"安全交底单"、"隐患告知书"即可', foreground="gray").pack(anchor=tk.W)

        self.titles_text = scrolledtext.ScrolledText(main_frame, height=8, font=("微软雅黑", 10))
        self.titles_text.pack(fill=tk.X, pady=5)
        self.titles_text.insert(tk.END, "安全交底单\n隐患告知书\n电力设施保护方案")

        # 4. 状态与运行按钮
        self.ocr_status_var = tk.StringVar(value="状态：准备就绪。")
        self.ocr_status_label = ttk.Label(main_frame, textvariable=self.ocr_status_var, foreground="#0052cc", font=("微软雅黑", 10))
        self.ocr_status_label.pack(pady=(20, 10))

        self.ocr_run_btn = tk.Button(main_frame, text=" 🚀 启动 PaddleOCR 智能切分 ",
                                     command=self.ocr_start_processing, bg="#28a745", fg="white",
                                     font=("微软雅黑", 12, "bold"), pady=8)
        self.ocr_run_btn.pack(fill=tk.X)

    # ---------- OCR Tab 交互逻辑 ----------

    def ocr_load_pdf(self):
        path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
        if path:
            self.ocr_pdf_path_var.set(path)
            self.ocr_out_dir_var.set(os.path.join(os.path.dirname(path), "拆分结果"))

    def ocr_select_out_dir(self):
        path = filedialog.askdirectory()
        if path:
            self.ocr_out_dir_var.set(path)

    def ocr_start_processing(self):
        pdf_path = self.ocr_pdf_path_var.get()
        out_dir = self.ocr_out_dir_var.get()
        raw_titles = self.titles_text.get("1.0", tk.END).strip().split('\n')
        titles = [re.sub(r'\s+', '', t) for t in raw_titles if t.strip()]

        if not pdf_path or not out_dir or not titles:
            messagebox.showwarning("警告", "请确保已加载PDF、设置输出目录，并输入了关键字！")
            return

        if not os.path.exists(out_dir):
            os.makedirs(out_dir)

        self.ocr_run_btn.config(state=tk.DISABLED)
        self.ocr_status_label.config(foreground="red")
        threading.Thread(target=self.ocr_process_pdf, args=(pdf_path, out_dir, titles)).start()

    # ---------- OCR 核心处理逻辑 ----------

    def ocr_process_pdf(self, pdf_path, out_dir, titles):
        try:
            if self.ocr_engine is None:
                self.ocr_status_var.set("状态：正在加载 PaddleOCR 深度学习模型，请稍候...")
                self.root.update_idletasks()
                self.ocr_engine = PaddleOCR(use_angle_cls=False, lang="ch", show_log=False)

            doc = fitz.open(pdf_path)
            group_counter = 1
            seen_titles_in_current_group = set()
            current_full_name = f"线路{group_counter}-未命名起步文件"
            start_page = 0

            dpi_factor = 2
            roi_height_ratio = 0.25

            for page_num in range(len(doc)):
                self.ocr_status_var.set(f"状态：模型正在极速扫描 第 {page_num + 1} / {len(doc)} 页...")
                self.root.update_idletasks()

                page = doc.load_page(page_num)
                rect = page.rect
                clip_rect = fitz.Rect(rect.x0, rect.y0, rect.x1, rect.height * roi_height_ratio)

                mat = fitz.Matrix(dpi_factor, dpi_factor)
                pix = page.get_pixmap(matrix=mat, clip=clip_rect)

                img_array = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.h, pix.w, pix.n)
                if pix.n == 4:
                    img_array = cv2.cvtColor(img_array, cv2.COLOR_RGBA2RGB)
                elif pix.n == 1:
                    img_array = cv2.cvtColor(img_array, cv2.COLOR_GRAY2RGB)
                elif pix.n == 3:
                    img_array = cv2.cvtColor(img_array, cv2.COLOR_RGB2BGR)

                result = self.ocr_engine.ocr(img_array, cls=False)

                cleaned_text = ""
                if result and result[0]:
                    for line in result[0]:
                        cleaned_text += line[1][0]
                cleaned_text = re.sub(r'\s+', '', cleaned_text)

                found_title = None
                for title in titles:
                    if title in cleaned_text:
                        found_title = title
                        break

                if found_title:
                    if found_title in seen_titles_in_current_group:
                        group_counter += 1
                        seen_titles_in_current_group.clear()

                    seen_titles_in_current_group.add(found_title)
                    new_full_name = f"线路{group_counter}-{found_title}"

                    if page_num > start_page:
                        self.save_pdf_chunk(doc, start_page, page_num - 1, current_full_name, out_dir)

                    current_full_name = new_full_name
                    start_page = page_num

            if start_page < len(doc):
                self.save_pdf_chunk(doc, start_page, len(doc) - 1, current_full_name, out_dir)

            doc.close()
            self.ocr_status_label.config(foreground="green")
            self.ocr_status_var.set("状态：PaddleOCR 拆分全部完成！")
            messagebox.showinfo("成功", f"PDF拆分全部完成！\n已按线路分组排序并保存至:\n{out_dir}")

        except Exception as e:
            self.ocr_status_var.set("状态：发生严重错误！")
            messagebox.showerror("错误", f"处理过程中出现错误:\n{str(e)}\n\n请检查 PaddleOCR 环境是否正常。")
        finally:
            self.ocr_run_btn.config(state=tk.NORMAL)

    def save_pdf_chunk(self, doc, start_page, end_page, title, out_dir):
        safe_title = "".join(c for c in title if c not in r'\/:*?"<>|')
        new_doc = fitz.open()
        new_doc.insert_pdf(doc, from_page=start_page, to_page=end_page)

        base_filename = os.path.join(out_dir, f"{safe_title}.pdf")
        filename = base_filename
        counter = 1
        while os.path.exists(filename):
            filename = os.path.join(out_dir, f"{safe_title}_{counter}.pdf")
            counter += 1

        new_doc.save(filename)
        new_doc.close()

    # ================================================================
    #                    Tab 2: 手动预览切割 (原 split_pdf_tool.py)
    # ================================================================

    def setup_manual_tab(self):
        paned_window = ttk.PanedWindow(self.manual_tab, orient=tk.HORIZONTAL)
        paned_window.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # ==================== 左侧：PDF 预览区 ====================
        left_frame = ttk.Frame(paned_window)
        paned_window.add(left_frame, weight=3)

        toolbar = ttk.Frame(left_frame)
        toolbar.pack(fill=tk.X, pady=(0, 5))

        ttk.Button(toolbar, text="上一页", command=self.manual_prev_page).pack(side=tk.LEFT, padx=5)
        self.manual_page_label = ttk.Label(toolbar, text="页码: 0 / 0", font=("微软雅黑", 10, "bold"))
        self.manual_page_label.pack(side=tk.LEFT, padx=10)
        ttk.Button(toolbar, text="下一页", command=self.manual_next_page).pack(side=tk.LEFT, padx=5)

        ttk.Button(toolbar, text="放大 (+)", command=self.manual_zoom_in).pack(side=tk.RIGHT, padx=5)
        ttk.Button(toolbar, text="缩小 (-)", command=self.manual_zoom_out).pack(side=tk.RIGHT, padx=5)

        canvas_frame = ttk.Frame(left_frame)
        canvas_frame.pack(fill=tk.BOTH, expand=True)

        self.manual_canvas = tk.Canvas(canvas_frame, bg="gray")
        vbar = ttk.Scrollbar(canvas_frame, orient=tk.VERTICAL, command=self.manual_canvas.yview)
        hbar = ttk.Scrollbar(canvas_frame, orient=tk.HORIZONTAL, command=self.manual_canvas.xview)
        self.manual_canvas.configure(yscrollcommand=vbar.set, xscrollcommand=hbar.set)

        vbar.pack(side=tk.RIGHT, fill=tk.Y)
        hbar.pack(side=tk.BOTTOM, fill=tk.X)
        self.manual_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.manual_canvas.bind("<MouseWheel>", self.manual_on_mouse_wheel)

        # ==================== 右侧：参数控制区 ====================
        right_frame = ttk.Frame(paned_window)
        paned_window.add(right_frame, weight=1)

        ttk.Label(right_frame, text="1. 加载文件", font=("微软雅黑", 11, "bold")).pack(anchor=tk.W, pady=(0, 5))
        ttk.Button(right_frame, text="加载 PDF 文件", command=self.manual_load_pdf).pack(fill=tk.X, pady=2)

        self.manual_file_label = ttk.Label(right_frame, text="未选择文件", foreground="blue", wraplength=200)
        self.manual_file_label.pack(anchor=tk.W, pady=(0, 10))

        ttk.Label(right_frame, text="2. 切割参数", font=("微软雅黑", 11, "bold")).pack(anchor=tk.W, pady=(5, 5))

        param_frame = ttk.Frame(right_frame)
        param_frame.pack(fill=tk.X)

        ttk.Label(param_frame, text="起始页码:").grid(row=0, column=0, sticky=tk.W, pady=2)
        self.manual_start_page_var = tk.StringVar()
        ttk.Entry(param_frame, textvariable=self.manual_start_page_var, width=8).grid(row=0, column=1, sticky=tk.W, pady=2)

        ttk.Label(param_frame, text="结束页码:").grid(row=1, column=0, sticky=tk.W, pady=2)
        self.manual_end_page_var = tk.StringVar()
        ttk.Entry(param_frame, textvariable=self.manual_end_page_var, width=8).grid(row=1, column=1, sticky=tk.W, pady=2)

        ttk.Label(param_frame, text="输出文件名:").grid(row=2, column=0, sticky=tk.W, pady=2)
        self.manual_out_name_var = tk.StringVar()
        ttk.Entry(param_frame, textvariable=self.manual_out_name_var, width=18).grid(row=2, column=1, sticky=tk.W, pady=2)
        ttk.Label(param_frame, text=".pdf").grid(row=2, column=2, sticky=tk.W, pady=2)

        ttk.Label(right_frame, text="3. 输出路径", font=("微软雅黑", 11, "bold")).pack(anchor=tk.W, pady=(10, 5))
        self.manual_out_dir_var = tk.StringVar()
        ttk.Entry(right_frame, textvariable=self.manual_out_dir_var).pack(fill=tk.X, pady=2)
        ttk.Button(right_frame, text="更改保存目录", command=self.manual_select_out_dir).pack(fill=tk.X, pady=2)

        ttk.Label(right_frame, text="").pack(pady=10)

        self.manual_split_btn = tk.Button(right_frame, text="立即切割", bg="#007BFF", fg="white",
                                          font=("微软雅黑", 14, "bold"), command=self.manual_split_pdf)
        self.manual_split_btn.pack(fill=tk.X, pady=5, ipady=8)

    # ---------- 手动模式 PDF 交互逻辑 ----------

    def manual_load_pdf(self):
        path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
        if path:
            self.manual_pdf_path = path
            self.manual_file_label.config(text=os.path.basename(path))
            self.manual_out_dir_var.set(os.path.dirname(path))

            if self.manual_doc:
                self.manual_doc.close()
            self.manual_doc = fitz.open(path)
            self.current_page = 0
            self.zoom_factor = 1.0

            self.manual_start_page_var.set("1")
            self.manual_end_page_var.set("1")
            self.manual_render_page()

    def manual_select_out_dir(self):
        path = filedialog.askdirectory()
        if path:
            self.manual_out_dir_var.set(path)

    def manual_render_page(self):
        if not self.manual_doc:
            return
        self.manual_page_label.config(text=f"页码: {self.current_page + 1} / {len(self.manual_doc)}")
        self.manual_end_page_var.set(str(self.current_page + 1))

        page = self.manual_doc.load_page(self.current_page)
        mat = fitz.Matrix(self.zoom_factor * 2, self.zoom_factor * 2)
        pix = page.get_pixmap(matrix=mat)

        mode = "RGBA" if pix.alpha else "RGB"
        img = Image.frombytes(mode, [pix.width, pix.height], pix.samples)
        self.current_image = ImageTk.PhotoImage(img)

        self.manual_canvas.delete("all")
        self.manual_canvas.create_image(0, 0, anchor=tk.NW, image=self.current_image)
        self.manual_canvas.config(scrollregion=self.manual_canvas.bbox(tk.ALL))

    def manual_prev_page(self):
        if self.manual_doc and self.current_page > 0:
            self.current_page -= 1
            self.manual_render_page()

    def manual_next_page(self):
        if self.manual_doc and self.current_page < len(self.manual_doc) - 1:
            self.current_page += 1
            self.manual_render_page()

    def manual_zoom_in(self):
        if self.zoom_factor < 3.0:
            self.zoom_factor += 0.3
            self.manual_render_page()

    def manual_zoom_out(self):
        if self.zoom_factor > 0.4:
            self.zoom_factor -= 0.3
            self.manual_render_page()

    def manual_on_mouse_wheel(self, event):
        self.manual_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def manual_split_pdf(self):
        if not self.manual_doc:
            messagebox.showwarning("警告", "请先加载一个 PDF 文件！")
            return

        try:
            start = int(self.manual_start_page_var.get()) - 1
            end = int(self.manual_end_page_var.get()) - 1
            out_name = self.manual_out_name_var.get().strip()
            out_dir = self.manual_out_dir_var.get()

            if start < 0 or end >= len(self.manual_doc) or start > end:
                messagebox.showerror("错误", f"页码范围无效！\n请输入 1 到 {len(self.manual_doc)} 之间的有效范围。")
                return
            if not out_name:
                messagebox.showerror("错误", "输出文件名不能为空！")
                return
            if not os.path.isdir(out_dir):
                messagebox.showerror("错误", "输出目录不存在！")
                return

            new_doc = fitz.open()
            new_doc.insert_pdf(self.manual_doc, from_page=start, to_page=end)

            safe_name = "".join(c for c in out_name if c not in r'\/:*?"<>|')
            save_path = os.path.join(out_dir, f"{safe_name}.pdf")

            counter = 1
            while os.path.exists(save_path):
                save_path = os.path.join(out_dir, f"{safe_name}_{counter}.pdf")
                counter += 1

            new_doc.save(save_path)
            new_doc.close()

            messagebox.showinfo("成功", f"文件已成功切割并保存至:\n{save_path}")

            # 自动跳到下一段，方便连续切割
            if end + 2 <= len(self.manual_doc):
                self.manual_start_page_var.set(str(end + 2))
                self.manual_out_name_var.set("")

        except ValueError:
            messagebox.showerror("错误", "页码必须是纯数字！")
        except Exception as e:
            messagebox.showerror("错误", f"发生异常: {str(e)}")


if __name__ == "__main__":
    root = tk.Tk()
    app = CombinedPDFSplitter(root)
    root.mainloop()
