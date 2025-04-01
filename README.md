# 📚 VerbaAurea 🌟

中文 [中文](./README.md) | 英文 [English](./README_EN.md)

VerbaAurea 是一个智能文档预处理工具，致力于将原始文档转化为"黄金"般的知识，为知识库构建提供高质量的文本数据。它专注于文档智能分割，确保语义完整性，为知识库检索和大语言模型微调提供优质素材。

## 项目特点

- **智能文档分割** - 基于句子边界和语义完整性进行精准分段
- **多维度评分系统** - 考虑标题、句子完整性、段落长度等多种因素决定最佳分割点
- **语义完整性保护** - 优先保证句子和语义单元的完整，避免在句子中间断开
- **可配置化设计** - 通过配置文件灵活调整分割策略，无需修改代码
- **多语言支持** - 针对中英文文本采用不同的句子分割策略
- **格式保留** - 保留原始文档的格式信息，包括样式、字体和表格

## 应用场景

- **知识库构建** - 为检索式问答系统提供合适粒度的文本单元

- **语料库准备** - 为大语言模型微调准备高质量的训练数据

- **文档索引** - 优化文档检索系统的索引单元

- **内容管理** - 改进内容管理系统中的文档组织方式

  
## 项目结构如下
```
├── main.py                 # 主程序入口
├── config_manager.py       # 配置管理
├── document_processor.py   # 文档处理核心
├── text_analysis.py        # 文本分析功能
├── parallel_processor.py   # 并行处理实现
├── utils.py                # 工具函数
├── config.json   # 自动生成的配置文件
├── requirements.txt   # 项目所需库
├── README.md   # 中文文档
├── README_EN.md   # 英文文档
├── LICENSE   # 开源许可证
└── 文档或文档所在文件夹...
```



## 核心功能

- **句子边界检测** - 结合规则和NLP技术，精确识别句子边界
- **分割点评分系统** - 多维度评分，选择最佳分割点
- **语义块分析** - 分析文档结构，保留段落间的语义联系
- **自适应长度控制** - 根据配置自动调整文本片段长度
- **格式保留处理** - 在分割的同时保留文档原始格式

## 安装说明

### 环境要求

- Python 3.6 或更高版本
- 支持 Windows、macOS 和 Linux 系统



### 安装步骤

1. 克隆项目到本地

```bash
git clone https://github.com/yourusername/VerbaAurea.git
cd VerbaAurea
```

2. 安装依赖

```bash
pip install -r requirements.txt
```

## 使用指南

### 基本用法

1. 将需要处理的Word文档放在脚本所在目录或子目录中
2. 运行主脚本

```bash
python main.py
```

1. 根据菜单选择操作:
   - 选择 `1` 开始处理文档
   - 选择 `2` 查看当前配置
   - 选择 `3` 编辑配置
   - 选择 `4` 退出程序

1. 处理后的文档将保存在`双碳输出`(默认)或自定义的输出文件夹中

### 配置说明

可以通过菜单编辑或直接修改`split_config.json`文件定制分割参数:

#### 文档设置

- `max_length`: 最大段落长度 (字符数)
- `min_length`: 最小段落长度 (字符数)
- `sentence_integrity_weight`: 句子完整性权重 (值越大，越避免在非句子边界处分割)

#### 处理选项

- `debug_mode`: 是否启用调试输出
- `output_folder`: 输出文件夹名称
- `skip_existing`: 是否跳过已存在的文件

#### 高级设置

- `min_split_score`: 最小分割得分 (决定分割阈值)
- `heading_score_bonus`: 标题加分值
- `sentence_end_score_bonus`: 句子结束加分值
- `length_score_factor`: 长度评分因子
- `search_window`: 搜索句子边界的窗口大小

### 最佳实践

- **设置合理的长度范围** - 根据知识库或应用需求，设置合适的最大和最小段落长度
- **调整句子完整性权重** - 如果出现句子被分割的情况，提高此权重
- **启用调试模式** - 处理重要文档时启用调试模式，观察分割点的选择过程
- **标题规范化** - 确保文档中的标题使用标准样式，以便更好地识别分割点

## 工作原理

1. **文档解析** - 解析Word文档，提取文本、样式和结构信息
2. **段落分析** - 分析每个段落的特征，如长度、是否为标题、是否以句号结尾等
3. **评分计算** - 为每个潜在分割点计算综合评分
4. **分割点选择** - 基于评分和配置选择最佳分割点
5. **句子边界校正** - 调整分割点位置，确保在句子边界处分割
6. **分割标记插入** - 在选定的位置插入`<!--split-->`标记
7. **格式保留** - 保留原文档的格式信息并保存为新文档

## 开发计划

-  添加对更多文档格式的支持
-  实现图形用户界面
-  增强语义分析能力，使用更先进的NLP模型
-  添加批量处理进度条和统计报告
-  支持更多中文分词和句子边界检测算法

## 常见问题

**Q: 为什么分割后的文档中有些段落太短或太长？**

A: 尝试调整配置文件中的 `max_length` 和 `min_length` 参数，平衡分割粒度。

**Q: 如何避免句子被分割在中间？**

A: 提高 `sentence_integrity_weight` 参数值，默认值为 8.0，可以尝试设置为 10.0 或更高。

**Q: 如何处理特殊格式的文档？**

A: 对于特殊格式，可以通过调整高级设置中的评分参数来适应不同的文档结构。

## 贡献指南

欢迎对VerbaAurea项目做出贡献! 您可以通过以下方式参与:

1. 报告Bug或提出功能建议
2. 提交Pull Request改进代码
3. 完善文档和使用示例
4. 分享您使用VerbaAurea的经验和案例

## Star History

<a href="https://www.star-history.com/#AEPAX/VerbaAurea&Date">
 <picture>
   <source media="(prefers-color-scheme: dark)" srcset="https://api.star-history.com/svg?repos=AEPAX/VerbaAurea&type=Date&theme=dark" />
   <source media="(prefers-color-scheme: light)" srcset="https://api.star-history.com/svg?repos=AEPAX/VerbaAurea&type=Date" />
   <img alt="Star History Chart" src="https://api.star-history.com/svg?repos=AEPAX/VerbaAurea&type=Date" />
 </picture>
</a>

本项目使用 [CC BY 4.0](https://creativecommons.org/licenses/by/4.0/) 许可协议。
