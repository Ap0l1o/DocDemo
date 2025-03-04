import docx
import re
from typing import Dict, List, Tuple

class DocProcessor:
    def __init__(self, file_path: str):
        """初始化文档处理器

        Args:
            file_path (str): Word文档路径
        """
        self.file_path = file_path
        self.doc = docx.Document(file_path)

    def extract_amount_sentences(self) -> List[str]:
        """提取包含'万元'的句子

        Returns:
            List[str]: 包含'万元'的句子列表
        """
        sentences = []
        for paragraph in self.doc.paragraphs:
            text = paragraph.text.strip()
            if '万元' in text:
                # 使用正则表达式匹配包含数字和'万元'的句子
                matches = re.finditer(r'[^。！？.，"]*?\d+万元[^。，！？.]*[，。！？.]?', text)
                for match in matches:
                    sentences.append(match.group().strip())
        return sentences

    def parse_expense_table(self) -> Tuple[Dict[str, float], float]:
        """解析费用明细表格

        Returns:
            Tuple[Dict[str, float], float]: (费用明细字典, 总金额)
        """
        expense_details = {}
        total_amount = 0.0
        table_total = 0.0

        print(f'\n找到 {len(self.doc.tables)} 个表格')
        for i, table in enumerate(self.doc.tables):
            print(f'\n检查第 {i+1} 个表格：')
            print(f'表格行数：{len(table.rows)}')
            if len(table.rows) > 0:
                print(f'第一行内容：{[cell.text for cell in table.rows[0].cells]}')

            # 检查是否是费用明细表格
            if self._is_expense_table(table):
                print('识别为费用明细表格，开始解析...')
                # 获取除表头和最后一行（汇总行）外的所有行
                data_rows = table.rows[1:-1]
                # 解析最后一行（汇总行）
                if len(table.rows) > 2:
                    last_row = table.rows[-1]
                    last_row_amount = last_row.cells[-1].text.strip()
                    total_match = re.search(r'\d+(\.\d+)?', last_row_amount)
                    if total_match:
                        table_total = float(total_match.group())
                        print(f'表格汇总金额：{table_total}万元')

                # 从第二行开始遍历（跳过表头和汇总行）
                processed_cells = set()  # 用于记录已处理的单元格
                for row_idx, row in enumerate(data_rows, 1):
                    try:
                        # 检查第一列单元格是否已处理（合并单元格的情况）
                        cell_key = (row.cells[0]._tc, row.cells[-1]._tc)  # 使用单元格的内部标识作为键
                        if cell_key in processed_cells:
                            print(f'跳过第 {row_idx} 行：合并单元格已处理')
                            continue

                        # 假设第一列是实施内容，最后一列是金额
                        content = row.cells[0].text.strip()
                        amount_text = row.cells[-1].text.strip()
                        print(f'处理第 {row_idx} 行：{content} - {amount_text}')

                        # 提取金额中的数字
                        amount_match = re.search(r'\d+(\.\d+)?', amount_text)
                        if amount_match:
                            amount = float(amount_match.group())
                            expense_details[content] = amount
                            total_amount += amount
                            print(f'成功解析金额：{amount}万元')
                            processed_cells.add(cell_key)  # 标记该单元格已处理
                        else:
                            print(f'警告：无法从文本中提取数字：{amount_text}')

                    except (ValueError, AttributeError) as e:
                        print(f'处理第 {row_idx} 行时出错：{str(e)}')
                        continue

                # 比较计算的总金额与表格汇总金额
                if table_total > 0:
                    if abs(total_amount - table_total) < 0.01:  # 考虑浮点数精度误差
                        print(f'\n计算总金额（{total_amount}万元）与表格汇总金额（{table_total}万元）一致')
                    else:
                        print(f'\n警告：计算总金额（{total_amount}万元）与表格汇总金额（{table_total}万元）不一致！')

        return expense_details, total_amount

    def _is_expense_table(self, table) -> bool:
        """判断是否是费用明细表格

        Args:
            table: docx表格对象

        Returns:
            bool: 是否是费用明细表格
        """
        # 检查表格是否有行
        if not table.rows:
            return False
            
        # 检查表格第一行是否包含'费用明细'字样
        header_text = ' '.join(cell.text.strip() for cell in table.rows[0].cells)
        print(f'表格标题行文本：{header_text}')

        # 如果表格标题包含'费用明细'，认为是费用明细表格
        if '费用明细' in header_text:
            return True

        # 遍历文档中的所有段落，找到当前表格在文档中的位置
        current_element = table._element
        previous_element = current_element.getprevious()
        
        # 向上查找最近的段落
        while previous_element is not None:
            if previous_element.tag.endswith('p'):  # 是段落元素
                paragraph_text = previous_element.xpath('.//w:t')
                if paragraph_text:
                    text = ''.join([t.text for t in paragraph_text if t.text])
                    if '费用明细' in text:
                        print(f'在表格上方找到标题：{text}')
                        return True
                break  # 只检查最近的段落
            previous_element = previous_element.getprevious()

        return False

def process_doc(file_path: str):
    """处理Word文档

    Args:
        file_path (str): Word文档路径
    """
    try:
        print(f'\n开始处理文档：{file_path}')
        processor = DocProcessor(file_path)
        
        # 提取包含'万元'的句子
        amount_sentences = processor.extract_amount_sentences()
        print('\n包含「万元」的句子：')
        for sentence in amount_sentences:
            print(f'- {sentence}')

        # 解析费用明细表格
        expense_details, total_amount = processor.parse_expense_table()
        if expense_details:
            print('\n费用明细：')
            for content, amount in expense_details.items():
                print(f'- {content}: {amount}万元')
            print(f'\n总金额：{total_amount}万元')
        else:
            print('\n未找到费用明细表格或表格为空')

    except Exception as e:
        print(f'处理文档时出错：{str(e)}')

if __name__ == '__main__':
    # 示例使用
    doc_path = input('请输入Word文档路径：')
    process_doc(doc_path)