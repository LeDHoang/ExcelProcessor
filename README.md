# ExcelProcessor-1: AWS Bedrock OCR Integration

A comprehensive OCR solution using AWS Bedrock with Claude Sonnet 4.5 for processing Excel images and documents. This project provides both simple and advanced OCR capabilities with bearer token authentication.

## 🚀 Quick Start

### Prerequisites
- Python 3.7+
- AWS Bedrock access with bearer token
- Required Python packages

### Installation

1. **Install dependencies:**
```bash
pip install boto3 pillow
```

2. **Set your AWS bearer token:**
```bash
export AWS_BEARER_TOKEN_BEDROCK='your-bearer-token-here'
```

3. **Test the setup:**
```bash
python simple_ocr_example.py
```

## 📖 Usage Guide

### Simple OCR with `simple_ocr_example.py`

The `simple_ocr_example.py` script provides an easy way to perform OCR on images using AWS Bedrock. It supports two modes:

#### Basic OCR Mode
```bash
python simple_ocr_example.py
```

This will:
- Process the image specified in `image_path` (default: `Sheet2.png`)
- Extract all text content
- Save results to a markdown file
- Use streaming responses for real-time output

#### Advanced OCR with Sub-Image Extraction
```bash
python simple_ocr_example.py --extract-images
```

This enhanced mode will:
- Extract all text content preserving hierarchy
- Identify visual elements (charts, diagrams, icons, photos)
- Extract sub-images automatically based on AI analysis
- Create a comprehensive markdown document
- Save sub-images to a dedicated directory

#### Configuration

Edit the script to customize:

```python
# Change the input image
image_path = 'your_image.png'

# The script will automatically:
# - Detect image format (PNG/JPEG)
# - Encode to base64
# - Send to Claude Sonnet 4.5
# - Generate output filename based on input
```

#### Output Files

**Basic Mode:**
- `{input_filename}_out_ocr_simple.md` - Markdown with extracted text

**Advanced Mode:**
- `{input_filename}_out_ocr_simple.md` - Complete markdown with images
- `{input_filename}_subimages/` - Directory with extracted sub-images
- Individual sub-image files: `subimg_01_top_left.png`, etc.

#### Example Usage

```bash
# Basic OCR on Sheet2.png
python simple_ocr_example.py

# Advanced OCR with sub-image extraction
python simple_ocr_example.py --extract-images

# Process a different image (edit the script first)
# Change image_path = 'your_image.png' in the script
python simple_ocr_example.py
```

### Advanced Features

#### Custom Prompts
The script uses different prompts based on mode:

**Basic Mode:**
```
"Please extract all text from this image. Preserve the structure and formatting."
```

**Advanced Mode:**
```
Analyze this image and create a structured markdown document with:
1. Extract all text content preserving hierarchy
2. Identify ALL visual elements (charts, diagrams, icons, photos)
3. For each visual element, provide position and percentage coordinates
4. Insert image placeholders where they appear in document flow
5. Maintain reading order: top to bottom, left to right
```

#### Sub-Image Extraction Logic

When using `--extract-images`, the script:

1. **Analyzes the image** using Claude Sonnet 4.5 vision capabilities
2. **Identifies visual elements** with descriptions and positions
3. **Calculates bounding boxes** based on percentage coordinates
4. **Estimates sizes** based on element type:
   - Icons/logos: 10% width, 8% height
   - Charts/graphs: 35% width, 25% height
   - Other elements: 20% width, 15% height
5. **Extracts and saves** sub-images with descriptive filenames

#### Output Format

**Markdown Structure:**
```markdown
# OCR Analysis with Sub-Images

**Source Image:** `Sheet2.png`  
**Processed:** 2024-01-15 14:30:25  
**Model:** Claude Sonnet 4.5 (Bedrock)  
**Sub-images extracted:** 5  
**Sub-images directory:** `Sheet2_subimages/`

---

[Extracted content with image placeholders]

---

## Extracted Sub-Images

| # | Description | Position | File |
|---|-------------|----------|------|
| 1 | Company Logo | top-left | `subimg_01_top_left.png` |
| 2 | Sales Chart | center | `subimg_02_center.png` |
```

## 🔧 Technical Details

### Authentication
Uses bearer token authentication following the same pattern as `testBedrockOcr.py`:

```python
session = boto3.Session(
    aws_access_key_id=os.environ.get('AWS_ACCESS_KEY_ID'),
    aws_secret_access_key=os.environ.get('AWS_SECRET_ACCESS_KEY'),
    aws_session_token=os.environ.get('AWS_BEARER_TOKEN_BEDROCK'),
    region_name='ap-southeast-1'
)
```

### Model Configuration
- **Model:** `global.anthropic.claude-sonnet-4-5-20250929-v1:0`
- **Region:** `ap-southeast-1` (Singapore)
- **Max Tokens:** 4096 (basic) / 8192 (advanced)
- **Response:** Streaming (basic) / Non-streaming (advanced)

### Supported Image Formats
- PNG (`.png`)
- JPEG (`.jpg`, `.jpeg`)
- Automatic format detection

## 📁 Project Structure

```
ExcelProcessor-1/
├── simple_ocr_example.py          # Main OCR script
├── bedrock_ocr.py                 # OCR class module
├── test_bedrock_ocr_integration.py # Integration tests
├── excelprocessor.py              # Excel extraction
├── START_HERE.md                  # Quick start guide
├── README_BEDROCK_OCR.md          # Detailed documentation
├── QUICKSTART_OCR.md              # 3-step quick start
├── output/                        # Generated outputs
│   ├── images/                    # Extracted Excel images
│   ├── ocr_results.json          # OCR results
│   └── ocr_outline.txt           # Text outline
└── {filename}_subimages/         # Sub-image directories
```

## 🎯 Use Cases

### 1. Document OCR
```bash
# Process a scanned document
python simple_ocr_example.py
```

### 2. Excel Sheet Analysis
```bash
# Extract and analyze Excel sheet images
python excelprocessor.py -i Sheet1.xlsx -o output
python simple_ocr_example.py  # Process the extracted images
```

### 3. Visual Content Extraction
```bash
# Extract both text and visual elements
python simple_ocr_example.py --extract-images
```

### 4. Batch Processing
```python
# Use the BedrockOCR class for batch processing
from bedrock_ocr import BedrockOCR

ocr = BedrockOCR()
results = ocr.batch_ocr(['image1.png', 'image2.png'])
```

## 🐛 Troubleshooting

### Common Issues

**1. Token Not Set**
```
WARNING: AWS_BEARER_TOKEN_BEDROCK not set in environment
```
**Solution:**
```bash
export AWS_BEARER_TOKEN_BEDROCK='your-token'
```

**2. Image Not Found**
```
Image not found: Sheet2.png
```
**Solution:**
- Check if the image file exists
- Update `image_path` variable in the script
- Extract images from Excel first: `python excelprocessor.py -i Sheet1.xlsx -o output`

**3. Authentication Error**
```
Error: Unable to locate credentials
```
**Solution:**
- Verify your bearer token is valid
- Check token expiration
- Ensure AWS credentials are properly set

**4. Import Errors**
```
ModuleNotFoundError: No module named 'boto3'
```
**Solution:**
```bash
pip install boto3 pillow
```

### Performance Tips

1. **Use basic mode** for simple text extraction
2. **Use advanced mode** only when you need visual element analysis
3. **Process smaller images** for faster results
4. **Batch similar images** together when possible

## 📚 Additional Documentation

- **[START_HERE.md](START_HERE.md)** - Complete project overview
- **[README_BEDROCK_OCR.md](README_BEDROCK_OCR.md)** - Detailed technical documentation
- **[QUICKSTART_OCR.md](QUICKSTART_OCR.md)** - 3-step quick start guide
- **[AWS_BEDROCK_SETUP_SUMMARY.md](AWS_BEDROCK_SETUP_SUMMARY.md)** - Setup guide

## 🎉 Getting Started

1. **Set up your environment:**
   ```bash
   export AWS_BEARER_TOKEN_BEDROCK='your-token'
   pip install boto3 pillow
   ```

2. **Run your first OCR:**
   ```bash
   python simple_ocr_example.py
   ```

3. **Try advanced features:**
   ```bash
   python simple_ocr_example.py --extract-images
   ```

4. **Explore the codebase:**
   - Read `simple_ocr_example.py` for implementation details
   - Check `bedrock_ocr.py` for the OCR class
   - Review test files for integration examples

Happy OCR-ing! 🎊
