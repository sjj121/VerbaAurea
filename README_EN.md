# 📚 VerbaAurea 🌟

🇨🇳 [中文](./README.md) | 🇺🇸 [English](./README_EN.md)

VerbaAurea is an intelligent document preprocessing tool dedicated to transforming raw documents into "golden" knowledge for high-quality text data in knowledge base construction. It focuses on intelligent document segmentation, ensuring semantic integrity, and providing quality material for knowledge base retrieval and large language model fine-tuning.

## Project Features

- **Intelligent Document Segmentation** - Precise segmentation based on sentence boundaries and semantic integrity
- **Multi-dimensional Scoring System** - Considers multiple factors such as headings, sentence completeness, and paragraph length to determine optimal split points
- **Semantic Integrity Protection** - Prioritizes sentence and semantic unit completeness, avoiding breaks in the middle of sentences
- **Configurable Design** - Flexible adjustment of segmentation strategies through configuration files without code modification
- **Multi-language Support** - Different sentence splitting strategies for Chinese and English texts
- **Format Preservation** - Maintains original document formatting, including styles, fonts, and tables

## Application Scenarios

- **Knowledge Base Construction** - Provides text units of appropriate granularity for retrieval-based QA systems
- **Corpus Preparation** - Prepares high-quality training data for fine-tuning large language models
- **Document Indexing** - Optimizes index units for document retrieval systems
- **Content Management** - Improves document organization in content management systems

## Project Structure:
```
├── main.py                 # Main program entry point
├── config_manager.py       # Configuration management
├── document_processor.py   # Document processing core
├── text_analysis.py        # Text analysis functions
├── parallel_processor.py   # Parallel processing implementation
├── utils.py                # Utility functions
├── config.json             # Auto-generated configuration file
├── requirements.txt        # Project dependencies
├── README.md               # Chinese documentation
├── README_EN.md            # English documentation
├── LICENSE                 # Open-source license
└── Documents or document folders...
```


## Core Functions

- **Sentence Boundary Detection** - Precisely identifies sentence boundaries using rules and NLP techniques
- **Split Point Scoring System** - Multi-dimensional scoring to select optimal split points
- **Semantic Block Analysis** - Analyzes document structure to preserve semantic connections between paragraphs
- **Adaptive Length Control** - Automatically adjusts text fragment length according to configuration
- **Format Preservation Processing** - Preserves original document formatting during segmentation

## Installation Guide

### Requirements

- Python 3.6 or higher
- Supports Windows, macOS, and Linux systems

### Installation Steps

1. Clone the project locally

```bash
git clone https://github.com/yourusername/VerbaAurea.git
cd VerbaAurea
```


2. Install dependencies

```bash
pip install -r requirements.txt
```


## Usage Guide

### Basic Usage

1. Place Word documents to be processed in the script directory or subdirectory
2. Run the main script

```bash
python main.py
```


3. Choose operations from the menu:
   - Select `1` to start processing documents
   - Select `2` to view current configuration
   - Select `3` to edit configuration
   - Select `4` to exit the program

4. Processed documents will be saved in the `双碳输出` (default) or custom output folder

### Configuration Details

You can customize segmentation parameters by editing through the menu or directly modifying the `split_config.json` file:

#### Document Settings

- `max_length`: Maximum paragraph length (character count)
- `min_length`: Minimum paragraph length (character count)
- `sentence_integrity_weight`: Sentence integrity weight (higher values avoid splitting at non-sentence boundaries)

#### Processing Options

- `debug_mode`: Whether to enable debug output
- `output_folder`: Output folder name
- `skip_existing`: Whether to skip existing files

#### Advanced Settings

- `min_split_score`: Minimum split score (determines splitting threshold)
- `heading_score_bonus`: Heading bonus score
- `sentence_end_score_bonus`: Sentence ending bonus score
- `length_score_factor`: Length scoring factor
- `search_window`: Window size for searching sentence boundaries

### Best Practices

- **Set reasonable length ranges** - Set appropriate maximum and minimum paragraph lengths based on knowledge base or application requirements
- **Adjust sentence integrity weight** - Increase this weight if sentences are being split
- **Enable debug mode** - Enable debug mode when processing important documents to observe the split point selection process
- **Standardize headings** - Ensure headings use standard styles for better split point identification

## How It Works

1. **Document Parsing** - Parses Word documents to extract text, style, and structure information
2. **Paragraph Analysis** - Analyzes paragraph characteristics such as length, heading status, and sentence ending
3. **Score Calculation** - Calculates comprehensive scores for potential split points
4. **Split Point Selection** - Selects optimal split points based on scores and configuration
5. **Sentence Boundary Correction** - Adjusts split point positions to ensure splits occur at sentence boundaries
6. **Split Mark Insertion** - Inserts `<!--split-->` markers at selected positions
7. **Format Preservation** - Preserves original document formatting and saves as new documents

## Development Plan

- Add support for more document formats
- Implement a graphical user interface
- Enhance semantic analysis capabilities using advanced NLP models
- Add batch processing progress bars and statistical reports
- Support more Chinese word segmentation and sentence boundary detection algorithms

## Frequently Asked Questions

**Q: Why are some paragraphs too short or too long after splitting?**
A: Try adjusting the `max_length` and `min_length` parameters in the configuration file to balance segmentation granularity.

**Q: How can I prevent sentences from being split in the middle?**
A: Increase the `sentence_integrity_weight` parameter value; the default is 8.0, try setting it to 10.0 or higher.

**Q: How do I handle documents with special formatting?**
A: For special formats, adjust the scoring parameters in advanced settings to adapt to different document structures.

## Contribution Guidelines

Contributions to the VerbaAurea project are welcome! You can participate by:

1. Reporting bugs or suggesting features
2. Submitting Pull Requests to improve the code
3. Enhancing documentation and usage examples
4. Sharing your experiences and case studies using VerbaAurea

## Star History

<a href="https://www.star-history.com/#AEPAX/VerbaAurea&Date">
 <picture>
   <source media="(prefers-color-scheme: dark)" srcset="https://api.star-history.com/svg?repos=AEPAX/VerbaAurea&type=Date&theme=dark" />
   <source media="(prefers-color-scheme: light)" srcset="https://api.star-history.com/svg?repos=AEPAX/VerbaAurea&type=Date" />
   <img alt="Star History Chart" src="https://api.star-history.com/svg?repos=AEPAX/VerbaAurea&type=Date" />
 </picture>
</a>

This project is licensed under the [CC BY 4.0](https://creativecommons.org/licenses/by/4.0/) license.
