# ExcelProcessor: End-to-End RAG Anything with Progress Tracking

A comprehensive RAG (Retrieval-Augmented Generation) solution that processes documents and enables intelligent querying using either OpenAI or AWS Bedrock. This project provides end-to-end document processing with real-time progress tracking and supports both simple OCR and advanced multimodal RAG capabilities.

## 🚀 Quick Start

### Prerequisites
- Python 3.7+
- Either OpenAI API access OR AWS Bedrock access
- Required Python packages (see requirements files)

### Installation

1. **Install dependencies:**
```bash
# For OpenAI version
pip install -r requirements.txt

# For AWS Bedrock version  
pip install -r requirements_pdf.txt
```

2. **Set your API credentials:**

**For OpenAI:**
```bash
export OPENAI_API_KEY='your-openai-api-key'
# Optional: Set custom base URL
export OPENAI_BASE_URL='https://api.openai.com/v1'
```

**For AWS Bedrock:**
```bash
export AWS_ACCESS_KEY_ID='your-access-key'
export AWS_SECRET_ACCESS_KEY='your-secret-key'
export AWS_BEARER_TOKEN_BEDROCK='your-bearer-token'
export BEDROCK_REGION='ap-southeast-1'
```

3. **Run with progress tracking:**
```bash
# OpenAI version (default)
python rag_anything_implementation/run_with_progress.py --pdf input/your-document.pdf

# AWS Bedrock version
python rag_anything_implementation/run_with_progress.py --aws --pdf input/your-document.pdf
```

## 📖 Usage Guide

### End-to-End RAG Anything

The main functionality is provided by the `run_with_progress.py` script, which offers real-time progress tracking and supports both OpenAI and AWS Bedrock backends.

#### Basic Usage
```bash
# Process a document with OpenAI (default)
python rag_anything_implementation/run_with_progress.py --pdf input/your-document.pdf

# Process a document with AWS Bedrock
python rag_anything_implementation/run_with_progress.py --aws --pdf input/your-document.pdf

# Specify custom output directory
python rag_anything_implementation/run_with_progress.py --pdf input/your-document.pdf --output-dir custom_output

# Quiet mode (only show progress, no child output)
python rag_anything_implementation/run_with_progress.py --pdf input/your-document.pdf --quiet-child
```

#### Advanced Usage
```bash
# Use custom script path
python rag_anything_implementation/run_with_progress.py --script /path/to/custom/script.py --pdf input/your-document.pdf

# Explicitly specify OpenAI provider
python rag_anything_implementation/run_with_progress.py --openai --pdf input/your-document.pdf
```

### What the RAG Anything Pipeline Does

The end-to-end pipeline will:
- **Parse your document** using advanced parsers (MinerU or Docling)
- **Extract multimodal content** including text, images, tables, and equations
- **Build a knowledge graph** with entities and relationships
- **Create vector embeddings** for semantic search
- **Enable intelligent querying** with both text and multimodal queries
- **Provide real-time progress tracking** with elapsed time and status updates

#### Output Structure

The RAG Anything pipeline creates a comprehensive output structure:

```
output/
├── auto/                           # Auto-generated content
│   ├── {document_name}_layout.pdf # Layout analysis
│   ├── images/                     # Extracted images
│   ├── {document_name}.json       # Structured data
│   └── {document_name}.md         # Markdown summary
└── rag_storage/                   # RAG storage (OpenAI)
    └── rag_storage_1024/          # RAG storage (AWS, 1024-dim)
```

#### Query Examples

After processing, you can query the document:

```python
# Text-based queries
result = await rag.aquery("What are the main topics in this document?")

# Multimodal queries with equations
result = await rag.aquery_with_multimodal(
    "Explain this equation in context",
    multimodal_content=[{
        "type": "equation",
        "latex": "P(d|q) = \\frac{P(q|d) \\cdot P(d)}{P(q)}",
        "equation_caption": "Document relevance probability"
    }]
)
```

### Legacy OCR Features

The project also includes simple OCR capabilities for basic image processing:

#### Simple OCR with `simple_ocr_example.py`

```bash
# Basic OCR mode
python simple_ocr_example.py

# Advanced OCR with sub-image extraction
python simple_ocr_example.py --extract-images
```

This provides:
- Text extraction from images
- Visual element identification
- Sub-image extraction
- Markdown output with image placeholders

## 🔧 Technical Details

### RAG Anything Configuration

#### OpenAI Backend
- **LLM Model:** `gpt-4o-mini` (configurable via `OPENAI_LLM_MODEL`)
- **Vision Model:** `gpt-4o` (configurable via `OPENAI_VISION_MODEL`)
- **Embedding Model:** `text-embedding-3-large` (configurable via `OPENAI_EMBEDDING_MODEL`)
- **Embedding Dimension:** 3072

#### AWS Bedrock Backend
- **LLM Model:** `global.anthropic.claude-sonnet-4-5-20250929-v1:0`
- **Vision Model:** Same as LLM model
- **Embedding Model:** `amazon.titan-embed-text-v2:0` (with fallback to v1)
- **Embedding Dimension:** 1024 (configurable via `BEDROCK_EMBEDDING_DIM`)
- **Region:** `ap-southeast-1` (configurable via `BEDROCK_REGION`)

### Authentication

**OpenAI:**
```python
# Uses standard OpenAI API key authentication
api_key = os.getenv("OPENAI_API_KEY")
base_url = os.getenv("OPENAI_BASE_URL", "https://api.openai.com/v1")
```

**AWS Bedrock:**
```python
session = boto3.Session(
    aws_access_key_id=os.environ.get('AWS_ACCESS_KEY_ID'),
    aws_secret_access_key=os.environ.get('AWS_SECRET_ACCESS_KEY'),
    aws_session_token=os.environ.get('AWS_BEARER_TOKEN_BEDROCK'),
    region_name=region
)
```

### Supported Document Formats
- PDF documents (`.pdf`)
- Excel files (`.xlsx`, `.xls`)
- Images (`.png`, `.jpg`, `.jpeg`)
- Text documents (`.txt`, `.md`)

## 📁 Project Structure

```
ExcelProcessor-1/
├── rag_anything_implementation/    # Main RAG Anything implementation
│   ├── run_with_progress.py       # Progress tracker with provider selection
│   ├── end-2-end-rag-anything.py  # OpenAI backend
│   ├── end-2-end-rag-anything-aws.py # AWS Bedrock backend
│   ├── optimize_pdf.py            # PDF optimization utilities
│   ├── input/                     # Input documents
│   └── output/                    # Generated outputs
│       ├── auto/                  # Auto-generated content
│       └── rag_storage*/          # RAG storage directories
├── simple_ocr_example.py          # Legacy OCR script
├── bedrock_ocr.py                 # Legacy OCR class module
├── excelprocessor.py              # Excel extraction utilities
├── requirements.txt               # OpenAI dependencies
├── requirements_pdf.txt           # AWS Bedrock dependencies
└── output/                        # Legacy outputs
    ├── images/                    # Extracted images
    └── ocr_results.json          # OCR results
```

## 🎯 Use Cases

### 1. Document Intelligence with RAG
```bash
# Process any document for intelligent querying
python rag_anything_implementation/run_with_progress.py --pdf input/your-document.pdf

# Query the processed document
# The system will automatically handle:
# - Document parsing and structure extraction
# - Knowledge graph construction
# - Vector embedding generation
# - Multimodal content processing
```

### 2. Multimodal Document Analysis
```bash
# Process documents with images, tables, and equations
python rag_anything_implementation/run_with_progress.py --aws --pdf input/complex-document.pdf

# Query with multimodal context
# - Ask about specific images or charts
# - Query mathematical equations
# - Analyze table data
# - Understand document structure
```

### 3. Batch Document Processing
```bash
# Process multiple documents
for doc in input/*.pdf; do
    python rag_anything_implementation/run_with_progress.py --pdf "$doc"
done
```

### 4. Legacy OCR Processing
```bash
# Simple image OCR (legacy feature)
python simple_ocr_example.py

# Advanced image analysis with sub-image extraction
python simple_ocr_example.py --extract-images
```

## 🐛 Troubleshooting

### Common Issues

**1. API Key Not Set (OpenAI)**
```
OPENAI_API_KEY not found in environment variables
```
**Solution:**
```bash
export OPENAI_API_KEY='your-openai-api-key'
```

**2. AWS Credentials Not Set (Bedrock)**
```
WARNING: AWS_BEARER_TOKEN_BEDROCK not set in environment
```
**Solution:**
```bash
export AWS_ACCESS_KEY_ID='your-access-key'
export AWS_SECRET_ACCESS_KEY='your-secret-key'
export AWS_BEARER_TOKEN_BEDROCK='your-bearer-token'
```

**3. Document Not Found**
```
Error: PDF not found at /path/to/document.pdf
```
**Solution:**
- Check if the document file exists
- Use absolute paths or ensure relative paths are correct
- Verify file permissions

**4. Import Errors**
```
ModuleNotFoundError: No module named 'raganything'
```
**Solution:**
```bash
# For OpenAI version
pip install -r requirements.txt

# For AWS Bedrock version
pip install -r requirements_pdf.txt
```

**5. Progress Tracking Issues**
```
Error: script not found at /path/to/script.py
```
**Solution:**
- Ensure you're running from the correct directory
- Use `--script` to specify custom script path
- Check that provider flags (`--openai`, `--aws`) are used correctly

### Performance Tips

1. **Use OpenAI backend** for faster processing and better model availability
2. **Use AWS Bedrock backend** for cost-effective processing with Claude models
3. **Process smaller documents** for faster results
4. **Use `--quiet-child`** to reduce output noise during processing
5. **Monitor progress** with the built-in elapsed time tracking

## 📚 Additional Documentation

- **[AWS_BEDROCK_SETUP_SUMMARY.md](AWS_BEDROCK_SETUP_SUMMARY.md)** - AWS Bedrock setup guide
- **[REQUIREMENTS.md](REQUIREMENTS.md)** - Detailed requirements and dependencies

## 🎉 Getting Started

1. **Set up your environment:**
   ```bash
   # For OpenAI
   export OPENAI_API_KEY='your-openai-api-key'
   pip install -r requirements.txt
   
   # For AWS Bedrock
   export AWS_ACCESS_KEY_ID='your-access-key'
   export AWS_SECRET_ACCESS_KEY='your-secret-key'
   export AWS_BEARER_TOKEN_BEDROCK='your-bearer-token'
   pip install -r requirements_pdf.txt
   ```

2. **Run your first RAG Anything processing:**
   ```bash
   # OpenAI version (default)
   python rag_anything_implementation/run_with_progress.py --pdf input/your-document.pdf
   
   # AWS Bedrock version
   python rag_anything_implementation/run_with_progress.py --aws --pdf input/your-document.pdf
   ```

3. **Try advanced features:**
   ```bash
   # Process with custom output directory
   python rag_anything_implementation/run_with_progress.py --pdf input/your-document.pdf --output-dir custom_output
   
   # Quiet mode for cleaner output
   python rag_anything_implementation/run_with_progress.py --pdf input/your-document.pdf --quiet-child
   ```

4. **Explore the codebase:**
   - Read `rag_anything_implementation/run_with_progress.py` for the main entry point
   - Check `rag_anything_implementation/end-2-end-rag-anything.py` for OpenAI backend
   - Review `rag_anything_implementation/end-2-end-rag-anything-aws.py` for AWS backend
   - Explore legacy OCR features in `simple_ocr_example.py`

Happy RAG-ing! 🎊
