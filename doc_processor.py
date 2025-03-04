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

        print(f'\n找到 {len(self.doc.tables)} 个表格')
        for i, table in enumerate(self.doc.tables):
            print(f'\n检查第 {i+1} 个表格：')
            print(f'表格行数：{len(table.rows)}')
            if len(table.rows) > 0:
                print(f'第一行内容：{[cell.text for cell in table.rows[0].cells]}')

            # 检查是否是费用明细表格
            if self._is_expense_table(table):
                print('识别为费用明细表格，开始解析...')
                # 从第二行开始遍历（跳过表头）
                for row_idx, row in enumerate(table.rows[1:], 1):
                    try:
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
                        else:
                            print(f'警告：无法从文本中提取数字：{amount_text}')

                    except (ValueError, AttributeError) as e:
                        print(f'处理第 {row_idx} 行时出错：{str(e)}')
                        continue

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

        # 检查表格周围的段落是否包含'费用明细'字样
        for paragraph in self.doc.paragraphs:
            if '费用明细' in paragraph.text:
                # 如果段落中包含费用明细字样，检查其后是否紧跟着这个表格
                if hasattr(paragraph._element.getnext(), 'tbl') and \
                   paragraph._element.getnext().tbl == table._element:
                    print(f'在表格前的段落中找到"费用明细"：{paragraph.text}')
                    return True

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