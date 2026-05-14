#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Word 转 PowerPoint 转换工具

功能：
- 读取 Word 文档内容
- 按照结构转换为 PowerPoint 幻灯片
- 保留文本格式和基本结构
"""

import os
import re
from docx import Document
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor


class WordToPptConverter:
    """Word 到 PPT 的转换器"""
    
    def __init__(self, word_file):
        """
        初始化转换器
        
        Args:
            word_file (str): Word 文件路径
        """
        self.word_file = word_file
        self.doc = Document(word_file)
        self.prs = Presentation()
        self.prs.slide_width = Inches(10)
        self.prs.slide_height = Inches(7.5)
        
    def add_title_slide(self, title, subtitle=""):
        """
        添加标题幻灯片
        
        Args:
            title (str): 标题
            subtitle (str): 副标题
        """
        slide_layout = self.prs.slide_layouts[6]  # 空白布局
        slide = self.prs.slides.add_slide(slide_layout)
        
        # 添加背景色
        background = slide.background
        fill = background.fill
        fill.solid()
        fill.fore_color.rgb = RGBColor(0, 51, 102)
        
        # 添加标题
        left = Inches(1)
        top = Inches(2.5)
        width = Inches(8)
        height = Inches(1.5)
        
        title_box = slide.shapes.add_textbox(left, top, width, height)
        title_frame = title_box.text_frame
        title_frame.text = title
        title_frame.word_wrap = True
        
        # 格式化标题
        for paragraph in title_frame.paragraphs:
            paragraph.font.size = Pt(60)
            paragraph.font.bold = True
            paragraph.font.color.rgb = RGBColor(255, 255, 255)
            paragraph.alignment = PP_ALIGN.CENTER
        
        # 添加副标题
        if subtitle:
            subtitle_box = slide.shapes.add_textbox(left, top + Inches(2), width, Inches(1))
            subtitle_frame = subtitle_box.text_frame
            subtitle_frame.text = subtitle
            for paragraph in subtitle_frame.paragraphs:
                paragraph.font.size = Pt(32)
                paragraph.font.color.rgb = RGBColor(200, 200, 200)
                paragraph.alignment = PP_ALIGN.CENTER
    
    def add_content_slide(self, title, content_list):
        """
        添加内容幻灯片
        
        Args:
            title (str): 标题
            content_list (list): 内容列表
        """
        slide_layout = self.prs.slide_layouts[6]  # 空白布局
        slide = self.prs.slides.add_slide(slide_layout)
        
        # 白色背景
        background = slide.background
        fill = background.fill
        fill.solid()
        fill.fore_color.rgb = RGBColor(255, 255, 255)
        
        # 添加标题
        left = Inches(0.5)
        top = Inches(0.5)
        width = Inches(9)
        height = Inches(0.8)
        
        title_box = slide.shapes.add_textbox(left, top, width, height)
        title_frame = title_box.text_frame
        title_frame.text = title
        
        for paragraph in title_frame.paragraphs:
            paragraph.font.size = Pt(44)
            paragraph.font.bold = True
            paragraph.font.color.rgb = RGBColor(0, 51, 102)
        
        # 添加分隔线
        line = slide.shapes.add_connector(1, left, top + height, left + width, top + height)
        line.line.color.rgb = RGBColor(0, 102, 204)
        line.line.width = Pt(2)
        
        # 添加内容
        content_top = top + height + Inches(0.5)
        content_box = slide.shapes.add_textbox(left + Inches(0.3), content_top, width - Inches(0.6), Inches(6))
        text_frame = content_box.text_frame
        text_frame.word_wrap = True
        
        for i, item in enumerate(content_list):
            if i > 0:
                text_frame.add_paragraph()
            p = text_frame.paragraphs[i]
            p.text = f"• {item}"
            p.font.size = Pt(24)
            p.font.color.rgb = RGBColor(51, 51, 51)
            p.space_before = Pt(6)
            p.space_after = Pt(6)
    
    def convert(self):
        """
        从 Word 文档转换为 PPT
        
        Returns:
            Presentation: 生成的演示文稿
        """
        paragraphs = self.doc.paragraphs
        
        if not paragraphs:
            print("Word 文档为��")
            return self.prs
        
        # 第一个段落作为标题
        first_text = paragraphs[0].text.strip()
        if first_text:
            self.add_title_slide(first_text)
        
        # 处理其他段落
        current_title = ""
        current_content = []
        
        for para in paragraphs[1:]:
            text = para.text.strip()
            
            if not text:
                continue
            
            # 判断是否为标题（较短的文本或特殊格式）
            # 这里使用简单的启发式方法：短文本作为标题
            if len(text) < 50 and current_content:
                # 保存之前的内容幻灯片
                if current_title:
                    self.add_content_slide(current_title, current_content)
                current_title = text
                current_content = []
            else:
                if not current_title:
                    current_title = text if len(text) < 50 else "主要内容"
                else:
                    current_content.append(text)
        
        # 添加最后一个幻灯片
        if current_title and current_content:
            self.add_content_slide(current_title, current_content)
        elif current_title:
            self.add_title_slide(current_title)
        
        return self.prs
    
    def save(self, output_file):
        """
        保存 PPT 文件
        
        Args:
            output_file (str): 输出文件路径
        """
        self.prs.save(output_file)
        print(f"✓ PPT 已保存到: {output_file}")


def main():
    """
    主函数
    """
    import sys
    
    if len(sys.argv) < 2:
        print("使用方法: python word_to_ppt.py <Word文件路径> [输出PPT路径]")
        print("示例: python word_to_ppt.py document.docx output.pptx")
        return
    
    word_file = sys.argv[1]
    output_file = sys.argv[2] if len(sys.argv) > 2 else word_file.replace('.docx', '.pptx')
    
    if not os.path.exists(word_file):
        print(f"错误: 文件 '{word_file}' 不存在")
        return
    
    print(f"正在转换: {word_file}")
    
    try:
        converter = WordToPptConverter(word_file)
        converter.convert()
        converter.save(output_file)
        print(f"✓ 转换成功！")
    except Exception as e:
        print(f"✗ 转换失败: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main()
