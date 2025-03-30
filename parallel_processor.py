#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
并行处理模块
实现多线程并行处理文档
"""

import os
import sys
import multiprocessing
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from document_processor import insert_split_markers
from utils import get_script_dir


def collect_files_to_process(config):
    """收集需要处理的文件列表"""
    # 获取脚本当前路径
    current_dir = get_script_dir()

    # 创建输出目录
    output_folder = config["processing_options"]["output_folder"]
    output_base_dir = os.path.join(current_dir, output_folder)
    os.makedirs(output_base_dir, exist_ok=True)

    files_to_process = []

    # 遍历当前目录及子目录
    for root, dirs, files in os.walk(current_dir):
        # 跳过输出文件夹
        if output_folder in root:
            continue

        # 创建相对路径
        rel_path = os.path.relpath(root, current_dir)
        if rel_path == ".":  # 当前目录
            rel_path = ""

        # 收集当前目录下的所有Word文档
        for file in files:
            if file.endswith(('.docx', '.doc')) and not file.startswith('~$'):  # 排除临时文件
                # 构建输入和输出路径
                input_path = os.path.join(root, file)

                # 构建输出路径，保持原始目录结构
                if rel_path:
                    output_dir = os.path.join(output_base_dir, rel_path)
                    os.makedirs(output_dir, exist_ok=True)
                else:
                    output_dir = output_base_dir

                output_path = os.path.join(output_dir, file)

                # 添加到待处理文件列表
                files_to_process.append((input_path, output_path))

    return files_to_process


def process_all_documents(config):
    """并行处理所有Word文档"""
    debug_mode = config["processing_options"]["debug_mode"]

    # 获取性能设置
    perf_settings = config.get("performance_settings", {})
    use_parallel = perf_settings.get("parallel_processing", True)

    if not use_parallel:
        # 使用传统的单线程处理
        return process_sequentially(config)

    # 获取CPU核心数，确定并行进程数
    num_workers = perf_settings.get("num_workers", 0)

    if num_workers <= 0:
        # 自动选择工作进程数: CPU核心数-1，至少为1
        num_workers = max(1, multiprocessing.cpu_count() - 1)

    if debug_mode:
        print(f"启用并行处理，使用 {num_workers} 个工作线程")

    # 收集所有需要处理的文件
    files_to_process = collect_files_to_process(config)
    total_files = len(files_to_process)

    if total_files == 0:
        print("没有找到需要处理的Word文档")
        return 0, 0, []

    processed_files = 0
    failed_files = []

    # 使用线程池并行处理文件
    with ThreadPoolExecutor(max_workers=num_workers) as executor:
        # 提交所有处理任务到线程池
        futures = {
            executor.submit(insert_split_markers, input_path, output_path, config):
                (input_path, output_path)
            for input_path, output_path in files_to_process
        }

        # 实时显示处理进度
        completed = 0
        for future in as_completed(futures):
            input_path, _ = futures[future]
            try:
                result = future.result()
                completed += 1

                # 计算并显示进度
                progress = completed / total_files * 100

                if result:
                    processed_files += 1
                    status = "成功"
                else:
                    failed_files.append(input_path)
                    status = "失败"

                if debug_mode:
                    print(
                        f"进度: [{completed}/{total_files}] {progress:.1f}% - {os.path.basename(input_path)} {status}")
                else:
                    # 非调试模式下，只显示一行进度信息并刷新
                    sys.stdout.write(f"\r进度: [{completed}/{total_files}] {progress:.1f}%")
                    sys.stdout.flush()
            except Exception as e:
                print(f"处理 {input_path} 时出错: {str(e)}")
                failed_files.append(input_path)
                completed += 1

    # 处理完成后换行
    if not debug_mode and total_files > 0:
        print()

    return total_files, processed_files, failed_files


def process_sequentially(config):
    """顺序处理所有Word文档（非并行方式）"""
    debug_mode = config["processing_options"]["debug_mode"]

    # 收集所有需要处理的文件
    files_to_process = collect_files_to_process(config)
    total_files = len(files_to_process)

    if total_files == 0:
        print("没有找到需要处理的Word文档")
        return 0, 0, []

    if debug_mode:
        print(f"使用单线程顺序处理 {total_files} 个文件")

    processed_files = 0
    failed_files = []

    # 顺序处理文件
    for i, (input_path, output_path) in enumerate(files_to_process):
        try:
            # 计算并显示进度
            progress = (i + 1) / total_files * 100

            result = insert_split_markers(input_path, output_path, config)
            if result:
                processed_files += 1
                status = "成功"
            else:
                failed_files.append(input_path)
                status = "失败"

            if debug_mode:
                print(f"进度: [{i + 1}/{total_files}] {progress:.1f}% - {os.path.basename(input_path)} {status}")
            else:
                # 非调试模式下，只显示一行进度信息并刷新
                sys.stdout.write(f"\r进度: [{i + 1}/{total_files}] {progress:.1f}%")
                sys.stdout.flush()

        except Exception as e:
            print(f"\n处理 {input_path} 时出错: {str(e)}")
            failed_files.append(input_path)

    # 处理完成后换行
    if not debug_mode and total_files > 0:
        print()

    return total_files, processed_files, failed_files