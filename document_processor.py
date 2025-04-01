#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
文档处理核心模块
处理Word文档，在适当位置插入分隔标记
"""

from docx import Document
import os
from text_analysis import (
    is_sentence_boundary,
    find_nearest_sentence_boundary,
    analyze_document
)


def insert_split_markers(input_file, output_file, config):
    """
    在Word文档中根据算法自动嵌入<!--split-->标记

    使用多种策略确定分隔点:
    1. 标题识别
    2. 语义结构分析
    3. 长度控制
    4. 自然段落边界
    5. 句子完整性保证
    """
    # 从配置中提取参数
    doc_settings = config["document_settings"]
    proc_options = config["processing_options"]
    adv_settings = config["advanced_settings"]

    max_length = doc_settings["max_length"]
    min_length = doc_settings["min_length"]
    sentence_integrity_weight = doc_settings["sentence_integrity_weight"]
    debug_mode = proc_options["debug_mode"]
    search_window = adv_settings["search_window"]
    min_split_score = adv_settings["min_split_score"]
    heading_score_bonus = adv_settings["heading_score_bonus"]
    sentence_end_score_bonus = adv_settings["sentence_end_score_bonus"]
    length_score_factor = adv_settings["length_score_factor"]

    if debug_mode:
        print(f"正在处理文档: {input_file}")

    # 如果配置为跳过已存在文件，且输出文件已存在，则跳过处理
    if proc_options["skip_existing"] and os.path.exists(output_file):
        if debug_mode:
            print(f"跳过已存在文件: {output_file}")
        return True

    try:
        doc = Document(input_file)
    except Exception as e:
        print(f"无法打开文档 {input_file}: {str(e)}")
        return False

    # 创建新文档
    new_doc = Document()

    # 文档分析阶段
    paragraphs_info, semantic_blocks = analyze_document(doc, debug_mode)

    if debug_mode:
        print(f"文档共有 {len(paragraphs_info)} 个段落，组合成 {len(semantic_blocks)} 个语义块")

    # 确定分隔点
    split_points = find_split_points(
        paragraphs_info,
        max_length,
        min_length,
        sentence_integrity_weight,
        search_window,
        min_split_score,
        heading_score_bonus,
        sentence_end_score_bonus,
        length_score_factor,
        debug_mode
    )

    # 后处理：检查所有分割点确保不会打断句子
    final_split_points = refine_split_points(
        paragraphs_info,
        split_points,
        search_window,
        debug_mode
    )

    if debug_mode:
        print(f"最终分割点: {final_split_points}")

    # 创建带有分隔符的新文档
    result = create_output_document(
        doc,
        new_doc,
        final_split_points,
        output_file,
        debug_mode
    )

    return result


def find_split_points(paragraphs_info, max_length, min_length,
                      sentence_integrity_weight, search_window,
                      min_split_score, heading_score_bonus,
                      sentence_end_score_bonus, length_score_factor,
                      debug_mode):
    """确定文档的分隔点"""
    split_points = []
    current_length = 0
    last_potential_split = -1

    for i, para_info in enumerate(paragraphs_info):
        if para_info['length'] == 0:  # 跳过空段落
            continue

        current_length += para_info['length']

        # 潜在的分隔点评分系统 (0-10分，越高越适合作为分隔点)
        split_score = calculate_split_score(
            i, para_info, paragraphs_info,
            current_length, min_length, max_length,
            sentence_integrity_weight, heading_score_bonus,
            sentence_end_score_bonus, length_score_factor,
            split_points
        )

        if debug_mode and split_score >= 0:
            print(f"段落 {i}: 得分={split_score:.1f}, 文本预览: '{para_info['text'][:50]}...'")

        # 记录潜在的分隔点
        if split_score >= min_split_score and i > 0:  # 至少要设定分以上才考虑作为分隔点
            if debug_mode:
                print(f"选择分割点: {i}, 得分: {split_score:.1f}")

            split_points.append(i)
            current_length = para_info['length']  # 重置长度计数
            last_potential_split = i
        elif current_length > max_length * 1.5:
            # 长度已经超过最大限制的1.5倍，需要找一个合适的分割点
            best_boundary_index = find_nearest_sentence_boundary(paragraphs_info, i, search_window)

            if best_boundary_index >= 0 and (not split_points or best_boundary_index > split_points[-1]):
                if debug_mode:
                    print(f"超长分段，选择最近的句子边界: {best_boundary_index}")

                split_points.append(best_boundary_index)
                # 重新计算当前长度
                if best_boundary_index < i:
                    current_length = sum(p['length'] for p in paragraphs_info[best_boundary_index:i + 1])
                else:
                    current_length = para_info['length']
                last_potential_split = best_boundary_index
            elif i - last_potential_split > 3:
                # 实在找不到合适的句子边界，只能在当前位置分割
                if debug_mode:
                    print(f"警告：在段落 {i} 处强制分割，未找到合适的句子边界")

                split_points.append(i)
                current_length = para_info['length']
                last_potential_split = i

    return split_points


def calculate_split_score(i, para_info, paragraphs_info, current_length,
                          min_length, max_length, sentence_integrity_weight,
                          heading_score_bonus, sentence_end_score_bonus,
                          length_score_factor, split_points):
    """计算段落作为分割点的得分"""
    split_score = 0

    # 1. 标题是最佳分隔点
    if para_info['is_heading']:
        split_score += heading_score_bonus

    # 2. 句子完整性检查
    if para_info['ends_with_period']:
        split_score += sentence_end_score_bonus

    # 3. 确保分割点在句子边界
    if i > 0 and is_sentence_boundary(paragraphs_info[i - 1]['text'], para_info['text']):
        split_score += 8 * sentence_integrity_weight / 8.0  # 使用传入的权重调整
    else:
        split_score -= 10  # 严重惩罚非句子边界的分割

    # 4. 长度已达到理想范围加分
    if current_length >= min_length:
        split_score += min(4, (current_length - min_length) // length_score_factor)
    elif current_length < min_length * 0.7:  # 如果太短则减分
        split_score -= 5

    # 5. 列表项开始或结束是好的分隔点
    if i > 0 and ((para_info['is_list_item'] and not paragraphs_info[i - 1]['is_list_item']) or
                  (not para_info['is_list_item'] and paragraphs_info[i - 1]['is_list_item'] and
                   paragraphs_info[i - 1]['ends_with_period'])):
        split_score += 3

    # 6. 避免相邻分隔点
    if split_points and i - split_points[-1] < 3:
        split_score -= 8

    # 7. 如果长度超过最大值，增加分割倾向
    if current_length > max_length:
        split_score += 4

    return split_score


def refine_split_points(paragraphs_info, split_points, search_window, debug_mode):
    """修正分割点，确保分割点在句子边界上"""
    final_split_points = []

    for split_point in split_points:
        # 获取分割点前后的文本
        before_text = paragraphs_info[split_point - 1]['text'] if split_point > 0 else ""
        after_text = paragraphs_info[split_point]['text']

        # 检查是否为句子边界
        if is_sentence_boundary(before_text, after_text):
            final_split_points.append(split_point)
        else:
            # 寻找附近更合适的分割点
            best_point = find_nearest_sentence_boundary(paragraphs_info, split_point, search_window)
            if best_point >= 0 and best_point not in final_split_points:
                if debug_mode:
                    print(f"修正分割点: {split_point} -> {best_point}")
                final_split_points.append(best_point)
            else:
                # 无法找到更好的点，保留原分割点但记录警告
                if debug_mode:
                    print(f"警告: 无法找到比 {split_point} 更好的句子边界")
                final_split_points.append(split_point)

    # 确保分割点有序
    return sorted(list(set(final_split_points)))



def copy_single_table(table, new_doc, debug_mode):
    """复制单个表格"""
    try:
        rows = len(table.rows)
        cols = len(table.rows[0].cells) if rows > 0 else 0

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
        if debug_mode:
            print(f"  警告: 处理表格时出错: {str(e)}")


def create_output_document(doc, new_doc, split_points, output_file, debug_mode):
    """根据原文档和分割点创建新的输出文档"""
    split_marker_count = 0
    paragraph_index = 0
    paragraph_map = {}
    table_map = {}

    # 创建映射
    for i, para in enumerate(doc.paragraphs):
        paragraph_map[para._element] = (para, i)

    for i, table in enumerate(doc.tables):
        table_map[table._element] = table

    # 获取所有段落和表格元素，按它们在文档中的顺序
    # 未来可以扩展为: './/w:p | .//w:tbl | .//w:drawing' 来包含图片
    all_elements = doc._element.xpath('.//w:p | .//w:tbl')

    # 处理每个元素
    for element in all_elements:
        if element.tag.endswith('p'):  # 段落
            if element in paragraph_map:
                para, para_index = paragraph_map[element]

                # 如果是分隔点，添加分隔符
                if para_index in split_points:
                    new_doc.add_paragraph("<!--split-->")
                    split_marker_count += 1

                # 复制段落
                copy_paragraph(para, new_doc, debug_mode)

        elif element.tag.endswith('tbl'):  # 表格
            if element in table_map:
                table = table_map[element]
                try:
                    copy_single_table(table, new_doc, debug_mode)
                except Exception as e:
                    if debug_mode:
                        print(f"  警告: 处理表格时出错: {str(e)}")

    # 保存文档
    output_dir = os.path.dirname(output_file)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir, exist_ok=True)

    try:
        new_doc.save(output_file)
        if debug_mode:
            print(f"已保存: {output_file}，插入了 {split_marker_count} 个分隔符")
        return True
    except Exception as e:
        print(f"保存文档 {output_file} 时出错: {str(e)}")
        return False


def copy_paragraph(src_para, new_doc, debug_mode):
    """复制段落内容和格式"""
    text = src_para.text
    new_para = new_doc.add_paragraph(text)

    # 复制格式
    try:
        if src_para.style:
            new_para.style = src_para.style
        new_para.alignment = src_para.alignment

        # 复制段落内的文本格式
        for j in range(min(len(src_para.runs), len(new_para.runs))):
            src_run = src_para.runs[j]
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
    except Exception as e:
        if debug_mode:
            print(f"  警告: 复制格式时出错: {str(e)}")

