import os
import re
from docx import Document
import sys
from pathlib import Path
import nltk
from nltk.tokenize import sent_tokenize

# 下载必要的nltk数据（第一次运行需要）
try:
    nltk.data.find('tokenizers/punkt')
except LookupError:
    nltk.download('punkt')


def insert_split_markers(input_file, output_file, max_length=1000, min_length=300):
    """
    在Word文档中根据算法自动嵌入<!--split-->标记

    使用多种策略确定分隔点:
    1. 标题识别
    2. 语义结构分析
    3. 长度控制
    4. 自然段落边界
    """
    print(f"正在处理文档: {input_file}")

    try:
        doc = Document(input_file)
    except Exception as e:
        print(f"无法打开文档 {input_file}: {str(e)}")
        return False

    # 创建新文档
    new_doc = Document()

    # 第一步: 文档分析阶段
    paragraphs_info = []
    total_text = ""

    for i, paragraph in enumerate(doc.paragraphs):
        text = paragraph.text.strip()
        if not text:  # 跳过空段落
            continue

        # 判断段落类型
        is_heading = paragraph.style.name.startswith(('Heading', '标题'))
        is_list_item = text.startswith(('•', '-', '*', '1.', '2.', '3.')) or (
                    len(text) > 2 and text[0].isdigit() and text[1] == '.')

        # 段落长度分类
        length_category = "short" if len(text) < 50 else "medium" if len(text) < 200 else "long"

        # 尝试判断段落是否为完整句子结束
        ends_with_period = text.endswith(('。', '！', '？', '.', '!', '?'))

        # 将段落信息保存起来
        paragraphs_info.append({
            'index': i,
            'text': text,
            'length': len(text),
            'is_heading': is_heading,
            'is_list_item': is_list_item,
            'ends_with_period': ends_with_period,
            'length_category': length_category
        })

        total_text += text + " "

    # 使用NLTK尝试进行句子拆分（用于辅助判断语义边界）
    try:
        sentences = sent_tokenize(total_text)
        avg_sentence_length = len(total_text) / len(sentences) if sentences else 0
    except:
        avg_sentence_length = 80  # 如果失败，使用默认值

    # 第二步: 确定分隔点
    split_points = []
    current_length = 0
    last_potential_split = -1

    for i, para_info in enumerate(paragraphs_info):
        current_length += para_info['length']

        # 潜在的分隔点评分系统 (0-10分，越高越适合作为分隔点)
        split_score = 0

        # 1. 标题加分
        if para_info['is_heading']:
            split_score += 8

        # 2. 段落结束是完整句子加分
        if para_info['ends_with_period']:
            split_score += 3

        # 3. 长度已达到理想范围加分
        if current_length >= min_length:
            split_score += min(5, (current_length - min_length) // 100)
        elif current_length < min_length * 0.7:  # 如果太短则减分
            split_score -= 3

        # 4. 列表项开始或结束是好的分隔点
        if i > 0 and ((para_info['is_list_item'] and not paragraphs_info[i - 1]['is_list_item']) or
                      (not para_info['is_list_item'] and paragraphs_info[i - 1]['is_list_item'])):
            split_score += 2

        # 5. 避免相邻分隔点
        if split_points and i - split_points[-1] < 3:
            split_score -= 5

        # 6. 如果长度超过最大值，增加分割倾向
        if current_length > max_length:
            split_score += 4

        # 记录潜在的分隔点
        if split_score >= 7 and i > 0:  # 至少要7分以上才考虑作为分隔点
            split_points.append(i)
            current_length = para_info['length']  # 重置长度计数
            last_potential_split = i
        elif current_length > max_length * 1.5 and i - last_potential_split > 3:
            # 太长了，没有找到理想分隔点，强制分隔
            split_points.append(i)
            current_length = para_info['length']
            last_potential_split = i

    # 第三步: 创建带有分隔符的新文档
    split_marker_count = 0
    current_para_index = 0

    for i, paragraph in enumerate(doc.paragraphs):
        text = paragraph.text.strip()
        if not text:  # 空段落也要保留
            new_doc.add_paragraph("")
            continue

        # 检查是否是分隔点
        if current_para_index in split_points:
            # 添加分隔符
            new_para = new_doc.add_paragraph("<!--split-->")
            split_marker_count += 1

        # 添加当前段落
        new_para = new_doc.add_paragraph(text)

        # 复制格式
        try:
            if paragraph.style:
                new_para.style = paragraph.style
            new_para.alignment = paragraph.alignment

            # 复制段落内的文本格式
            for j in range(min(len(paragraph.runs), len(new_para.runs))):
                src_run = paragraph.runs[j]
                dst_run = new_para.runs[j]

                for attr in ['bold', 'italic', 'underline']:
                    if hasattr(src_run, attr):
                        setattr(dst_run, attr, getattr(src_run, attr))

                # 复制字体属性
                if hasattr(src_run, 'font') and hasattr(dst_run, 'font'):
                    for attr in ['size', 'name', 'color']:
                        try:
                            if hasattr(src_run.font, attr):
                                src_value = getattr(src_run.font, attr)
                                if src_value:
                                    setattr(dst_run.font, attr, src_value)
                        except:
                            pass
        except:
            pass

        current_para_index += 1

    # 处理表格
    try:
        for table in doc.tables:
            try:
                rows = len(table.rows)
                cols = 0
                if rows > 0:
                    cols = len(table.rows[0].cells)

                if rows > 0 and cols > 0:
                    new_table = new_doc.add_table(rows=rows, cols=cols)

                    try:
                        new_table.style = table.style
                    except:
                        pass

                    for i, row in enumerate(table.rows):
                        for j, cell in enumerate(row.cells):
                            if i < len(new_table.rows) and j < len(new_table.rows[i].cells):
                                try:
                                    new_cell = new_table.rows[i].cells[j]
                                    if cell.text:
                                        new_cell.text = cell.text
                                except:
                                    pass
            except Exception as e:
                print(f"  警告: 处理表格时出错: {str(e)}")
    except Exception as e:
        print(f"  警告: 处理文档中的表格时出错: {str(e)}")

    # 保存文档
    output_dir = os.path.dirname(output_file)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir, exist_ok=True)

    try:
        new_doc.save(output_file)
        print(f"已保存: {output_file}，插入了 {split_marker_count} 个分隔符")
        return True
    except Exception as e:
        print(f"保存文档 {output_file} 时出错: {str(e)}")
        return False


def process_all_documents():
    """处理当前目录下所有Word文档（排除"双碳输出"文件夹）"""
    # 获取脚本当前路径
    current_dir = os.path.dirname(os.path.abspath(sys.argv[0]))
    if not current_dir:  # 如果为空，使用当前工作目录
        current_dir = os.getcwd()

    # 创建输出目录
    output_base_dir = os.path.join(current_dir, "双碳输出")
    os.makedirs(output_base_dir, exist_ok=True)

    total_files = 0
    processed_files = 0
    failed_files = []

    # 遍历当前目录及子目录
    for root, dirs, files in os.walk(current_dir):
        # 跳过"双碳输出"文件夹
        if "双碳输出" in root:
            continue

        # 创建相对路径
        rel_path = os.path.relpath(root, current_dir)
        if rel_path == ".":  # 当前目录
            rel_path = ""

        # 处理当前目录下的所有Word文档
        for file in files:
            if file.endswith(('.docx', '.doc')) and not file.startswith('~$'):  # 排除临时文件
                total_files += 1

                # 构建输入和输出路径
                input_path = os.path.join(root, file)

                # 构建输出路径，保持原始目录结构
                if rel_path:
                    output_dir = os.path.join(output_base_dir, rel_path)
                    os.makedirs(output_dir, exist_ok=True)
                else:
                    output_dir = output_base_dir

                output_path = os.path.join(output_dir, file)

                # 处理文档
                try:
                    if insert_split_markers(input_path, output_path):
                        processed_files += 1
                    else:
                        failed_files.append(input_path)
                except Exception as e:
                    print(f"处理 {input_path} 时出错: {str(e)}")
                    failed_files.append(input_path)

    return total_files, processed_files, failed_files


def main():
    """主函数"""
    print("开始处理Word文档...")
    print("该脚本将根据智能算法在Word文档中插入<!--split-->分隔符")

    # 检查是否安装了所需库
    try:
        import nltk
    except ImportError:
        print("警告: 未安装nltk库，将使用简化的分隔算法")
        print("建议安装nltk以获得更好的分隔效果: pip install nltk")

    total_files, processed_files, failed_files = process_all_documents()

    print(f"\n处理完成! 共找到 {total_files} 个Word文档，成功处理 {processed_files} 个，失败 {len(failed_files)} 个。")
    print("处理后的文档已保存在当前目录下的'双碳输出'文件夹中。")

    if failed_files:
        print("\n以下文件处理失败:")
        for file in failed_files:
            print(f" - {file}")

    input("按Enter键退出...")  # 添加这一行使窗口不会立即关闭


if __name__ == "__main__":
    main()