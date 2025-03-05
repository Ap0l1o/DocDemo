import tkinter as tk
from tkinter import ttk, filedialog, scrolledtext
from doc_processor import process_doc
import sys
import io

class DocProcessorGUI:
    def __init__(self, root):
        self.root = root
        self.root.title('文档处理工具')
        self.root.geometry('800x600')

        # 创建主框架
        self.main_frame = ttk.Frame(self.root, padding="10")
        self.main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # 文件选择区域
        self.file_frame = ttk.Frame(self.main_frame)
        self.file_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 10))

        self.file_path = tk.StringVar()
        self.file_entry = ttk.Entry(self.file_frame, textvariable=self.file_path, width=60)
        self.file_entry.grid(row=0, column=0, padx=(0, 5))

        self.browse_button = ttk.Button(self.file_frame, text='选择文件', command=self.browse_file)
        self.browse_button.grid(row=0, column=1)

        self.process_button = ttk.Button(self.file_frame, text='处理文档', command=self.process_document)
        self.process_button.grid(row=0, column=2, padx=(5, 0))

        # 结果显示区域
        self.result_text = scrolledtext.ScrolledText(self.main_frame, wrap=tk.WORD, width=80, height=30)
        self.result_text.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # 配置grid权重
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        self.main_frame.columnconfigure(0, weight=1)
        self.main_frame.rowconfigure(1, weight=1)
        self.file_frame.columnconfigure(0, weight=1)

    def browse_file(self):
        file_path = filedialog.askopenfilename(
            title='选择Word文档',
            filetypes=[("Word文档", "*.doc;*.docx")]
        )
        if file_path:
            self.file_path.set(file_path)

    def process_document(self):
        file_path = self.file_path.get()
        if not file_path:
            self.show_result('请先选择要处理的Word文档')
            return

        # 重定向标准输出到StringIO
        output = io.StringIO()
        sys.stdout = output

        try:
            process_doc(file_path)
            # 获取输出结果
            result = output.getvalue()
            self.show_result(result)
        except Exception as e:
            self.show_result(f'处理文档时出错：{str(e)}')
        finally:
            # 恢复标准输出
            sys.stdout = sys.__stdout__
            output.close()

    def show_result(self, text):
        self.result_text.delete('1.0', tk.END)
        self.result_text.insert('1.0', text)

def main():
    root = tk.Tk()
    app = DocProcessorGUI(root)
    root.mainloop()

if __name__ == '__main__':
    main()