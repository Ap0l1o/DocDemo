import docx
import re
from typing import Dict, List, Tuple, Tuple

class DocProcessor:
    def compare_amounts(self, text_amt: float, table_calculated: float) -> Dict[str, float]:
        """金额对比验证

        Args:
            text_amt: 文本提取金额
            table_calculated: 表格计算总金额

        Returns:
            Dict[str, float]: 差异分析结果（差异值 = |文本金额 - 表格计算金额|）
        """
        diff_value = abs(text_amt - table_calculated)
        return {
            'text_vs_table': diff_value,
            'calculation_formula': f'差异值 = |{text_amt:.2f} - {table_calculated:.2f}|'
        }
    def __init__(self, file_path: str):
        """初始化文档处理器

        Args:
            file_path (str): Word文档路径
        """
        self.file_path = file_path
        self.doc = docx.Document(file_path)

    def extract_amount_sentences(self) -> Tuple[List[str], List[float]]:
        """提取包含万元的句子及金额

        Returns:
            Tuple[List[str], List[float]]: (句子列表, 金额列表)
        """
        sentences = []
        amounts = []
        pattern = r'(\d+(?:\.\d+)?)万元'
        
        for paragraph in self.doc.paragraphs:
            text = paragraph.text.strip()
            if '万元' in text:
                # 分步处理：先匹配完整句子结构
                sentence_matches = re.findall(r'[^，。！？；]+万元[^，。！？；]*', text)
                for raw_sentence in sentence_matches:
                    print(f'提取到句子：{raw_sentence}')
                    sentences.append(raw_sentence.strip())
                
                # 提取具体金额数值
                amount_matches = re.findall(pattern, text)
                for amt in amount_matches:
                    amounts.append(float(amt))
        
        return sentences, amounts

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
                        table_total = float(total_match.group()) / 10000  # 转换为万元
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
                            amount = round(float(amount_match.group()) / 10000, 4)  # 精确到小数点后四位
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

        return expense_details, total_amount, table_total

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
        amount_sentences, extracted_amounts = processor.extract_amount_sentences()
        print('\n包含「万元」的句子：')
        for sentence in amount_sentences:
            print(f'- {sentence}')

        # 解析费用明细表格
        # 获取所有数据
        amount_sentences, extracted_amounts = processor.extract_amount_sentences()
        expense_details, calculated_total, _ = processor.parse_expense_table()
        
        # 第一部分：提取的万元句子
        print('\n====== 文本提取结果 ======')
        print(f'\n发现 {len(amount_sentences)} 个包含万元的句子：')
        for i, sentence in enumerate(amount_sentences, 1):
            print(f'{i}. {sentence}')
        
        # 第二部分：表格解析结果
        print('\n====== 表格解析结果 ======')
        if expense_details:
            print('\n费用明细列表：')
            for content, amount in expense_details.items():
                print(f'- {content}: {amount}万元')
            
            print(f'\n表格计算总金额：{calculated_total}万元')
            comparison = processor.compare_amounts(amt, calculated_total)
            print(f'对比计算方式：{comparison["calculation_formula"]}')
            print(f'金额差异：{comparison["text_vs_table"]:.2f}万元')
        else:
            print('\n未找到费用明细表格')
        
        # 第三部分：金额对比验证
        print('\n====== 金额一致性验证 ======')
        if expense_details:
            print(f'表格计算总金额：{calculated_total}万元')
            comparison = processor.compare_amounts(amt, calculated_total)
            print(f'对比计算方式：{comparison["calculation_formula"]}')
            print(f'金额差异：{comparison["text_vs_table"]:.2f}万元')
            
            print('\n逐个对比提取金额：')
            for idx, amt in enumerate(extracted_amounts, 1):
                print(f'\n第 {idx} 个提取金额：{amt}万元')
                comparison = processor.compare_amounts(amt, calculated_total)
                print(f'对比计算方式：{comparison["calculation_formula"]}')
                print(f'金额差异：{comparison["text_vs_table"]:.2f}万元')
            
            # 删除重复的对比代码块
            # 遍历所有提取的金额进行对比
            for idx, extracted_amt in enumerate(extracted_amounts, 1):
                print(f'\n第 {idx} 个提取金额：{extracted_amt}万元')
                comparison = processor.compare_amounts(extracted_amt, calculated_total)
                print(f'对比计算方式：{comparison["calculation_formula"]}')
                print(f'金额差异：{comparison["text_vs_table"]:.2f}万元')
                
                if comparison["text_vs_table"] >= 0.01:
                    print(f'发现金额差异：{comparison["text_vs_table"]:.2f}万元')

    except Exception as e:
        print(f'处理文档时出错：{str(e)}')

if __name__ == '__main__':
    # 示例使用
    doc_path = input('请输入Word文档路径：')
    process_doc(doc_path)