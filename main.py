#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Word文档智能分割工具
主程序入口
"""

import time
import sys
from config_manager import load_config, show_config, edit_config, save_config
from parallel_processor import process_all_documents
from utils import check_dependencies


def main():
    """主函数"""
    print("VerbaAurea")
    print("============================")
    print("该工具会根据智能算法在Word文档中插入<!--split-->分隔符")

    # 检查必要的库
    check_dependencies()

    # 加载配置
    config = load_config()

    # 确保性能配置存在
    if "performance_settings" not in config:
        config["performance_settings"] = {
            "parallel_processing": True,
            "num_workers": 0,  # 0表示自动选择
            "cache_size": 1024,  # 缓存大小(MB)
            "batch_size": 50
        }
        save_config(config)

    # 显示主菜单
    while True:
        print("\n=== 主菜单 ===")
        print("1. 开始处理文档")
        print("2. 查看当前配置")
        print("3. 编辑配置")
        print("4. 退出")

        choice = input("\n请选择操作 [1-4]: ").strip()

        if choice == '1':
            # 开始处理文档
            print("\n开始处理文档...")
            start_time = time.time()

            total_files, processed_files, failed_files = process_all_documents(config)

            elapsed_time = time.time() - start_time

            print(
                f"\n处理完成! 共找到 {total_files} 个Word文档，成功处理 {processed_files} 个，失败 {len(failed_files)} 个。")
            print(f"总处理时间: {elapsed_time:.2f}秒")

            if processed_files > 0:
                avg_time = elapsed_time / processed_files
                print(f"平均每个文档处理时间: {avg_time:.2f}秒")

            print(f"处理后的文档已保存在当前目录下的'{config['processing_options']['output_folder']}'文件夹中。")

            if failed_files:
                print("\n以下文件处理失败:")
                for file in failed_files:
                    print(f" - {file}")

            input("\n按Enter键继续...")

        elif choice == '2':
            # 查看当前配置
            show_config()

            # 显示性能设置
            if "performance_settings" in config:
                perf = config["performance_settings"]
                print("\n性能设置:")
                print(f"  并行处理: {'开启' if perf.get('parallel_processing', True) else '关闭'}")
                print(f"  工作线程数: {perf.get('num_workers', 0)} (0表示自动选择)")
                print(f"  缓存大小: {perf.get('cache_size', 1024)}MB")
                print(f"  批处理大小: {perf.get('batch_size', 50)}")

            input("\n按Enter键继续...")

        elif choice == '3':
            # 编辑配置
            edit_config()
            # 重新加载配置以确保使用最新的值
            config = load_config()

            # 编辑性能设置
            print("\n性能设置:")
            if "performance_settings" not in config:
                config["performance_settings"] = {
                    "parallel_processing": True,
                    "num_workers": 0,
                    "cache_size": 1024,
                    "batch_size": 50
                }

            perf = config["performance_settings"]
            parallel = input(f"  并行处理 (y/n) [{'y' if perf.get('parallel_processing', True) else 'n'}]: ")
            workers = input(f"  工作线程数 (0=自动) [{perf.get('num_workers', 0)}]: ")
            cache = input(f"  缓存大小 (MB) [{perf.get('cache_size', 1024)}]: ")
            batch = input(f"  批处理大小 [{perf.get('batch_size', 50)}]: ")

            if parallel.strip():
                perf['parallel_processing'] = parallel.lower() == 'y'

            if workers.strip():
                try:
                    perf['num_workers'] = int(workers)
                except ValueError:
                    print("  输入无效，保留原值")

            if cache.strip():
                try:
                    perf['cache_size'] = int(cache)
                except ValueError:
                    print("  输入无效，保留原值")

            if batch.strip():
                try:
                    perf['batch_size'] = int(batch)
                except ValueError:
                    print("  输入无效，保留原值")

            # 保存更新后的配置
            save_config(config)
            print("配置已保存")

        elif choice == '4':
            # 退出
            print("\n谢谢使用！")
            break

        else:
            print("\n无效选择，请重试")


if __name__ == "__main__":
    main()