---
name: ppt-template-refiner
description: 专门用于读取标准 PPT 模板（.pptx），提取设计规范并根据用户输入润色排版输出 PPT。
---

# Skill: PPT Template Refiner & Polisher

## Overview
该技能专门用于读取用户提供的标准 PPT 模板（.pptx），提取其中的版式设计、字体、色彩规范、占位符，并将用户输入的散乱文本内容进行润色，最终严格按照模板的视觉规范输出排版的 PPT 文件。

## Requirements
- python-pptx >= 0.6.21
- python-docx (可选，用于处理复杂的文本输入)

## Execution Steps

1. **模版解析 (Template Parsing)**
   - 读取用户上传的 `.pptx` 模板文件，识别幻灯片母版（Slide Masters）以及常用的布局页面（Layouts），特别是封面页、目录页、过渡页、图文内容页 and 结尾页的占位符（Placeholders）。

2. **内容润色 (Content Polishing)**
   - 针对用户提供的文本，进行结构化提炼：将大段文字提炼为金句标题、核心观点（不超过 20 字）以及 3-5 点精简的 Bullet Points。
   - 确保文风专业、有逻辑，符合商务/设计领域的标准。

3. **代码生成与渲染 (Code Execution)**
   - 在 Manus 虚拟机环境中，编写并执行一个 Python 脚本。
   - 使用 `python-pptx` 的 `prs.slides.add_slide(prs.slide_layouts[i])` 方法调用模板中对应的页面布局。
   - 通过 `shape.text_frame` 严格将润色后的文本写入模板的既定占位符中，**绝对不要**随意新建悬浮文本框，确保字体大小、颜色、行距继承自模板设置。

4. **格式检查 (Quality Assurance)**
   - 检查文本是否溢出（Overflow），如果字数过多导致出界，自动触发缩减字数或精简表达。
   - 检查页面前后的逻辑色调是否与模板保持一致。

## Outputs
- 返回优化润色后的文字大纲供用户确认。
- 生成并提供最终下载链接的规范化 `.pptx` 文件。
