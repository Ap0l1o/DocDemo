import tkinter as tk
from tkinter import ttk, filedialog, scrolledtext
from doc_processor import process_doc
import sys
import io
import re

class DocProcessorGUI:
    def __init__(self, root):
        self.root = root
        self.root.title('项目开发类金额校验工具')
        self.root.geometry('1000x500')

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

        self.process_button = ttk.Button(self.file_frame, text='校验文档', command=self.process_document)
        self.process_button.grid(row=0, column=2, padx=(5, 0))

        self.export_button = ttk.Button(self.file_frame, text='导出结果', command=self.export_result, state='disabled')
        self.export_button.grid(row=0, column=3, padx=(5, 0))

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
            self.result_text.delete('1.0', tk.END)

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
            # 启用导出按钮
            self.export_button['state'] = 'normal'
        except Exception as e:
            self.show_result(f'处理文档时出错：{str(e)}')
            # 禁用导出按钮
            self.export_button['state'] = 'disabled'
        finally:
            # 恢复标准输出
            sys.stdout = sys.__stdout__
            output.close()

    def export_result(self):
        if not self.result_text.get('1.0', tk.END).strip():
            self.show_result('没有可导出的处理结果')
            return

        # 获取原文件名
        original_file = self.file_path.get()
        if not original_file:
            self.show_result('请先选择并处理Word文档')
            return

        # 构建默认导出文件名
        base_name = original_file.rsplit('.', 1)[0]
        export_name = f"{base_name}_费用分析结果.txt"

        # 打开文件保存对话框
        file_path = filedialog.asksaveasfilename(
            title='保存分析结果',
            initialfile=export_name.split('/')[-1].split('\\')[-1],
            defaultextension='.txt',
            filetypes=[('文本文件', '*.txt')]
        )

        if file_path:
            try:
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write(self.result_text.get('1.0', tk.END))
                self.show_result('分析结果已成功导出')
            except Exception as e:
                self.show_result(f'导出文件时出错：{str(e)}')

    def show_result(self, text):
        self.result_text.delete('1.0', tk.END)
        self.result_text.tag_configure('red', foreground='red')
        self.result_text.tag_configure('bold', font=('TkDefaultFont', 10, 'bold'))
        # 禁用导出按钮，因为内容被清空了
        self.export_button['state'] = 'disabled'
        
        # 使用正则表达式匹配金额数字
        lines = text.split('\n')
        current_section = 0  # 0: 未知, 1: 第一部分, 2: 第二部分, 3: 第三部分
        
        for line in lines:
            start_pos = self.result_text.index('end-1c')
            self.result_text.insert('end', line + '\n')
            
            # 处理标题加粗
            if '一、文档中包含的金额信息' in line or '二、费用明细表格解析结果' in line or '三、金额对比分析' in line:
                self.result_text.tag_add('bold', start_pos, f"{start_pos.split('.')[0]}.{len(line)}")
                if '一、文档中包含的金额信息' in line:
                    current_section = 1
                elif '二、费用明细表格解析结果' in line:
                    current_section = 2
                elif '三、金额对比分析' in line:
                    current_section = 3
            
            if current_section == 1 and '万元' in line:
                # 第一部分：只标红万元前的数字
                amounts = re.finditer(r'(\d+)(?=万元)', line)
                for match in amounts:
                    start = f"{start_pos.split('.')[0]}.{match.start()}"
                    end = f"{start_pos.split('.')[0]}.{match.end()}"
                    self.result_text.tag_add('red', start, end)
            elif current_section == 2:
                # 第二部分：标红所有金额数字
                if ':' in line:
                    # 费用明细项的金额
                    amounts = re.finditer(r'\d+(?:,\d{3})*\.\d{2}(?=元)', line)
                    for match in amounts:
                        start = f"{start_pos.split('.')[0]}.{match.start()}"
                        end = f"{start_pos.split('.')[0]}.{match.end()}"
                        self.result_text.tag_add('red', start, end)
                elif '费用明细汇总金额' in line or '表格汇总金额' in line:
                    # 标红计算总金额和表格汇总金额
                    amounts = re.finditer(r'\d+(?:,\d{3})*\.\d{2}(?=元)', line)
                    for match in amounts:
                        start = f"{start_pos.split('.')[0]}.{match.start()}"
                        end = f"{start_pos.split('.')[0]}.{match.end()}"
                        self.result_text.tag_add('red', start, end)
                elif '相差' in line:
                    # 标红差异值
                    amounts = re.finditer(r'(?<==\s)\d+(?:,\d{3})*\.\d{2}(?=元)', line)
                    for match in amounts:
                        start = f"{start_pos.split('.')[0]}.{match.start()}"
                        end = f"{start_pos.split('.')[0]}.{match.end()}"
                        self.result_text.tag_add('red', start, end)
            elif current_section == 3 and '相差' in line:
                # 第三部分：只标红差异值
                amounts = re.finditer(r'(?<==\s)\d+(?:,\d{3})*\.\d{2}(?=元)', line)
                for match in amounts:
                    start = f"{start_pos.split('.')[0]}.{match.start()}"
                    end = f"{start_pos.split('.')[0]}.{match.end()}"
                    self.result_text.tag_add('red', start, end)
def main():
    root = tk.Tk()
    app = DocProcessorGUI(root)
    root.mainloop()

if __name__ == '__main__':
    main()