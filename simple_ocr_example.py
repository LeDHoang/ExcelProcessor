"""
Simple OCR Example using AWS Bedrock with Bearer Token
This follows the same pattern as testBedrockOcr.py
Exports results to markdown file

Usage:
  python simple_ocr_example.py                    # Basic OCR
  python simple_ocr_example.py --extract-images   # OCR with sub-image extraction
"""

import boto3
import json
import os
import re
import sys
import base64
from datetime import datetime
from pathlib import Path
from PIL import Image

# Setup AWS session with bearer token (same pattern as testBedrockOcr.py)
session = boto3.Session(
    aws_access_key_id=os.environ.get('AWS_ACCESS_KEY_ID'),
    aws_secret_access_key=os.environ.get('AWS_SECRET_ACCESS_KEY'),
    aws_session_token=os.environ.get('AWS_BEARER_TOKEN_BEDROCK'),
    region_name='ap-southeast-1'
)

bedrock = session.client('bedrock-runtime')

# Claude Sonnet 4.5 model with vision capabilities
model_id = 'global.anthropic.claude-sonnet-4-5-20250929-v1:0'

# Check if user wants to extract sub-images
EXTRACT_SUBIMAGES = '--extract-images' in sys.argv

# Path to image you want to OCR
image_path = 'sheet2.png'

# Check if image exists
if not os.path.exists(image_path):
    print(f"Image not found: {image_path}")
    print("Please update the image_path variable or extract images first:")
    print("  python excelprocessor.py -i Sheet1.xlsx -o output")
    exit(1)

# Create output directory for sub-images if needed
input_filename = Path(image_path).stem
if EXTRACT_SUBIMAGES:
    output_dir = f"{input_filename}_subimages"
    os.makedirs(output_dir, exist_ok=True)
    print(f"Sub-image extraction enabled")
    print(f"Sub-images will be saved to: {output_dir}/")
    
    # Load image for sub-image extraction
    img = Image.open(image_path)
    img_width, img_height = img.size
    print(f"Image dimensions: {img_width} x {img_height} pixels")
    print()

# Read and encode image to base64
with open(image_path, 'rb') as img_file:
    image_data = base64.b64encode(img_file.read()).decode('utf-8')

# Determine image type
if image_path.endswith('.png'):
    media_type = 'image/png'
elif image_path.endswith('.jpg') or image_path.endswith('.jpeg'):
    media_type = 'image/jpeg'
else:
    media_type = 'image/png'

# Prepare prompt based on mode
if EXTRACT_SUBIMAGES:
    ocr_prompt = """Analyze this image and create a structured markdown document with the following:

1. **Extract all text content** preserving hierarchy (headings, paragraphs, lists)

2. **Identify ALL visual elements** (charts, diagrams, icons, photos, illustrations, graphs, tables)
   
3. **For each visual element**, provide in this EXACT format:
   [IMAGE: Brief description | Position: top-left/top-center/top-right/middle-left/center/middle-right/bottom-left/bottom-center/bottom-right | ApproxPercent: X% from top, Y% from left]

4. **Insert image placeholders EXACTLY where they appear** in the document flow

5. **Maintain reading order**: top to bottom, left to right

6. Use markdown formatting: `#` for headings, `##` for subheadings, `-` for bullets, `**bold**` for emphasis

Example:
```
# Title
[IMAGE: Logo | Position: top-left | ApproxPercent: 5% from top, 10% from left]

## Section
Text content...
[IMAGE: Chart | Position: center | ApproxPercent: 40% from top, 50% from left]
```

Provide complete markdown with all text and image placeholders."""
else:
    ocr_prompt = "Please extract all text from this image. Preserve the structure and formatting."

# Prepare request body for Claude with vision
request_body = {
    "anthropic_version": "bedrock-2023-05-31",
    "max_tokens": 8192 if EXTRACT_SUBIMAGES else 4096,
    "messages": [
        {
            "role": "user",
            "content": [
                {
                    "type": "image",
                    "source": {
                        "type": "base64",
                        "media_type": media_type,
                        "data": image_data
                    }
                },
                {
                    "type": "text",
                    "text": ocr_prompt
                }
            ]
        }
    ]
}

print(f"Processing image: {image_path}")
print(f"Mode: {'OCR with sub-image extraction' if EXTRACT_SUBIMAGES else 'Basic OCR'}")
print("Sending request to AWS Bedrock...")
print()

if EXTRACT_SUBIMAGES:
    # Use non-streaming for easier parsing when extracting images
    response = bedrock.invoke_model(
        modelId=model_id,
        body=json.dumps(request_body),
        contentType='application/json'
    )
    
    response_body = json.loads(response['body'].read())
    extracted_text = ""
    if 'content' in response_body:
        for content_block in response_body['content']:
            if content_block.get('type') == 'text':
                extracted_text += content_block.get('text', '')
    
    print("=== Extracted Content ===")
    print(extracted_text)
    print()
    print("=== Analysis Complete ===")
    
else:
    # Use streaming for basic OCR (same as testBedrockOcr.py)
    response = bedrock.invoke_model_with_response_stream(
        modelId=model_id,
        body=json.dumps(request_body),
        contentType='application/json'
    )
    
    print("=== Extracted Text ===")
    print()
    
    # Stream the response and collect text
    extracted_text = ""
    for event in response['body']:
        chunk = json.loads(event['chunk']['bytes'].decode())
        
        if chunk['type'] == 'content_block_delta':
            if 'text' in chunk['delta']:
                text_chunk = chunk['delta']['text']
                print(text_chunk, end='', flush=True)
                extracted_text += text_chunk
        elif chunk['type'] == 'message_stop':
            break
    
    print()
    print()
    print("=== OCR Complete ===")

# Generate output filename from input filename
output_filename = f"{input_filename}_out_ocr_simple.md"

# Process sub-images if enabled
sub_images_info = []
final_content = extracted_text

if EXTRACT_SUBIMAGES:
    print()
    print("Extracting sub-images...")
    print()
    
    # Parse image placeholders
    image_pattern = r'\[IMAGE:\s*([^\|]+)\s*\|\s*Position:\s*([^\|]+)\s*\|\s*ApproxPercent:\s*(\d+)%\s*from top,\s*(\d+)%\s*from left\]'
    image_matches = re.finditer(image_pattern, extracted_text, re.IGNORECASE)
    
    image_counter = 0
    for match in image_matches:
        description = match.group(1).strip()
        position = match.group(2).strip()
        percent_top = int(match.group(3))
        percent_left = int(match.group(4))
        
        image_counter += 1
        
        print(f"Sub-image {image_counter}: {description}")
        print(f"  Position: {position} ({percent_top}% from top, {percent_left}% from left)")
        
        # Calculate bounding box
        center_x = int(img_width * percent_left / 100)
        center_y = int(img_height * percent_top / 100)
        
        # Estimate size based on description
        if 'icon' in description.lower() or 'logo' in description.lower():
            box_width = int(img_width * 0.10)
            box_height = int(img_height * 0.08)
        elif 'chart' in description.lower() or 'graph' in description.lower() or 'diagram' in description.lower():
            box_width = int(img_width * 0.35)
            box_height = int(img_height * 0.25)
        else:
            box_width = int(img_width * 0.20)
            box_height = int(img_height * 0.15)
        
        left = max(0, center_x - box_width // 2)
        top = max(0, center_y - box_height // 2)
        right = min(img_width, center_x + box_width // 2)
        bottom = min(img_height, center_y + box_height // 2)
        
        # Extract and save sub-image
        try:
            sub_img = img.crop((left, top, right, bottom))
            sub_img_filename = f"subimg_{image_counter:02d}_{position.replace('-', '_')}.png"
            sub_img_path = os.path.join(output_dir, sub_img_filename)
            sub_img.save(sub_img_path)
            
            print(f"  ✓ Saved: {sub_img_path} ({right-left}x{bottom-top} px)")
            
            sub_images_info.append({
                'original_placeholder': match.group(0),
                'description': description,
                'position': position,
                'filename': sub_img_filename,
                'relative_path': f"{output_dir}/{sub_img_filename}"
            })
        except Exception as e:
            print(f"  ✗ Error: {e}")
        
        print()
    
    # Replace placeholders with markdown images
    for img_info in sub_images_info:
        markdown_img = f"\n\n![{img_info['description']}]({img_info['relative_path']})\n\n*{img_info['description']}*\n\n"
        final_content = final_content.replace(img_info['original_placeholder'], markdown_img)
    
    # Handle any remaining simple [IMAGE: ...] patterns
    simple_image_pattern = r'\[IMAGE:\s*([^\]]+)\]'
    final_content = re.sub(
        simple_image_pattern,
        lambda m: f"\n\n> 📷 **Visual Element:** {m.group(1).strip()}\n\n",
        final_content
    )

# Create markdown content
if EXTRACT_SUBIMAGES:
    markdown_content = f"""# OCR Analysis with Sub-Images

**Source Image:** `{image_path}`  
**Processed:** {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}  
**Model:** Claude Sonnet 4.5 (Bedrock)  
**Sub-images extracted:** {len(sub_images_info)}  
**Sub-images directory:** `{output_dir}/`

---

{final_content}

---

## Extracted Sub-Images

"""
    
    if sub_images_info:
        markdown_content += "\n| # | Description | Position | File |\n"
        markdown_content += "|---|-------------|----------|------|\n"
        for idx, img_info in enumerate(sub_images_info, 1):
            markdown_content += f"| {idx} | {img_info['description']} | {img_info['position']} | `{img_info['filename']}` |\n"
    else:
        markdown_content += "\nNo sub-images were automatically extracted.\n"
    
    markdown_content += f"\n\n*Generated by AWS Bedrock OCR with Sub-Image Extraction*\n"
    
else:
    markdown_content = f"""# OCR Results

**Source Image:** `{image_path}`  
**Processed:** {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}  
**Model:** Claude Sonnet 4.5 (Bedrock)

---

## Extracted Text

{final_content}

---

*Generated by AWS Bedrock OCR*
"""

# Save to markdown file
with open(output_filename, 'w', encoding='utf-8') as f:
    f.write(markdown_content)

print()
print("=" * 80)
print("✓ Processing Complete!")
print("=" * 80)
print(f"Markdown output: {output_filename}")
if EXTRACT_SUBIMAGES and sub_images_info:
    print(f"Sub-images directory: {output_dir}/")
    print(f"Sub-images extracted: {len(sub_images_info)}")
print("=" * 80)

