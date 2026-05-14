#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
快速转换脚本 - 自动转换仓库中的所有 Word 文件
"""

import os
import glob
from word_to_ppt import WordToPptConverter


def convert_all_word_files():
    """
    转换当前目录中的所有 .docx 文件
    """
    # 创建输出目录
    output_dir = "output"
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
        print(f"✓ 创建输出目录: {output_dir}")
    
    # 查找所有 .docx 文件
    word_files = glob.glob("*.docx")
    
    if not word_files:
        print("未找到 .docx 文件")
        return
    
    print(f"找到 {len(word_files)} 个 Word 文件\n")
    
    for word_file in word_files:
        print(f"正在处理: {word_file}")
        
        try:
            # 生成输出文件名
            base_name = os.path.splitext(word_file)[0]
            output_file = os.path.join(output_dir, f"{base_name}.pptx")
            
            # 转换
            converter = WordToPptConverter(word_file)
            converter.convert()
            converter.save(output_file)
            
            print(f"  ✓ 已保存到: {output_file}\n")
            
        except Exception as e:
            print(f"  ✗ 转换失败: {e}\n")
            import traceback
            traceback.print_exc()


if __name__ == "__main__":
    convert_all_word_files()
