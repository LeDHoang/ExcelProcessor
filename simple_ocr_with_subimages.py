"""
Advanced OCR with Sub-Image Extraction using AWS Bedrock
This script:
1. Analyzes the image to identify sub-images (charts, diagrams, icons, photos)
2. Extracts those sub-images and saves them
3. Creates a markdown output with embedded sub-images in their proper positions
"""

import boto3
import json
import os
import re
import base64
from datetime import datetime
from pathlib import Path
from PIL import Image
import io

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

# Path to image you want to OCR
image_path = 'sheet2.png'

# Check if image exists
if not os.path.exists(image_path):
    print(f"Image not found: {image_path}")
    print("Please update the image_path variable or extract images first:")
    print("  python excelprocessor.py -i Sheet1.xlsx -o output")
    exit(1)

# Create output directory for sub-images
input_filename = Path(image_path).stem
output_dir = f"{input_filename}_subimages"
os.makedirs(output_dir, exist_ok=True)

print(f"Processing image: {image_path}")
print(f"Sub-images will be saved to: {output_dir}/")
print()

# Load the image for processing
img = Image.open(image_path)
img_width, img_height = img.size
print(f"Image dimensions: {img_width} x {img_height} pixels")

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

# Enhanced prompt to identify sub-images with position information
analysis_prompt = """Analyze this image and create a structured markdown document with the following:

1. **Extract all text content** preserving hierarchy (headings, paragraphs, lists)

2. **Identify ALL visual elements** (charts, diagrams, icons, photos, illustrations, graphs, tables)
   
3. **For each visual element**, provide in this EXACT format:
   ```
   [IMAGE: Brief description | Position: top-left/top-center/top-right/middle-left/center/middle-right/bottom-left/bottom-center/bottom-right | ApproxPercent: X% from top, Y% from left]
   ```

4. **Insert image placeholders EXACTLY where they appear** in the document flow

5. **Maintain reading order**: top to bottom, left to right

6. **Use markdown formatting**:
   - `#` for main headings
   - `##` for subheadings
   - `-` for bullet points
   - `**bold**` for emphasis
   - Tables in markdown format if present

Example format:
```
# Main Title

[IMAGE: Company logo | Position: top-left | ApproxPercent: 5% from top, 10% from left]

## Section 1

Some text content here...

[IMAGE: Bar chart showing sales data | Position: center | ApproxPercent: 40% from top, 50% from left]

More text content...
```

Provide the complete markdown with all text and image placeholders in their correct positions."""

# Prepare request body for Claude with vision
request_body = {
    "anthropic_version": "bedrock-2023-05-31",
    "max_tokens": 8192,
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
                    "text": analysis_prompt
                }
            ]
        }
    ]
}

print("Analyzing image and identifying sub-images...")
print()

# Call Bedrock (non-streaming for easier parsing)
response = bedrock.invoke_model(
    modelId=model_id,
    body=json.dumps(request_body),
    contentType='application/json'
)

# Parse response
response_body = json.loads(response['body'].read())
extracted_content = ""
if 'content' in response_body:
    for content_block in response_body['content']:
        if content_block.get('type') == 'text':
            extracted_content += content_block.get('text', '')

print("=== Analysis Complete ===")
print()
print("Extracted content preview:")
print(extracted_content[:500] + "..." if len(extracted_content) > 500 else extracted_content)
print()

# Parse image placeholders from the content
# Pattern: [IMAGE: description | Position: position | ApproxPercent: X% from top, Y% from left]
image_pattern = r'\[IMAGE:\s*([^\|]+)\s*\|\s*Position:\s*([^\|]+)\s*\|\s*ApproxPercent:\s*(\d+)%\s*from top,\s*(\d+)%\s*from left\]'
image_matches = re.finditer(image_pattern, extracted_content, re.IGNORECASE)

# Extract and save sub-images
sub_images_info = []
image_counter = 0

for match in image_matches:
    description = match.group(1).strip()
    position = match.group(2).strip()
    percent_top = int(match.group(3))
    percent_left = int(match.group(4))
    
    image_counter += 1
    
    print(f"Found sub-image {image_counter}: {description}")
    print(f"  Position: {position} ({percent_top}% from top, {percent_left}% from left)")
    
    # Calculate approximate bounding box
    # Use a heuristic to estimate the size of the sub-image
    # Default: 20% width, 15% height around the center point
    center_x = int(img_width * percent_left / 100)
    center_y = int(img_height * percent_top / 100)
    
    # Estimate box size based on position
    if 'icon' in description.lower() or 'logo' in description.lower():
        box_width = int(img_width * 0.10)  # 10% width for icons
        box_height = int(img_height * 0.08)  # 8% height
    elif 'chart' in description.lower() or 'graph' in description.lower() or 'diagram' in description.lower():
        box_width = int(img_width * 0.35)  # 35% width for charts
        box_height = int(img_height * 0.25)  # 25% height
    else:
        box_width = int(img_width * 0.20)  # 20% width default
        box_height = int(img_height * 0.15)  # 15% height
    
    # Calculate bounding box
    left = max(0, center_x - box_width // 2)
    top = max(0, center_y - box_height // 2)
    right = min(img_width, center_x + box_width // 2)
    bottom = min(img_height, center_y + box_height // 2)
    
    # Extract sub-image
    try:
        sub_img = img.crop((left, top, right, bottom))
        sub_img_filename = f"subimg_{image_counter:02d}_{position.replace('-', '_')}.png"
        sub_img_path = os.path.join(output_dir, sub_img_filename)
        sub_img.save(sub_img_path)
        
        print(f"  ✓ Saved to: {sub_img_path}")
        print(f"  Dimensions: {right-left} x {bottom-top} pixels")
        
        # Store info for markdown generation
        sub_images_info.append({
            'original_placeholder': match.group(0),
            'description': description,
            'position': position,
            'filename': sub_img_filename,
            'relative_path': f"{output_dir}/{sub_img_filename}"
        })
        
    except Exception as e:
        print(f"  ✗ Error extracting sub-image: {e}")
    
    print()

# Replace image placeholders with actual markdown image references
final_markdown = extracted_content

for img_info in sub_images_info:
    # Replace the placeholder with markdown image syntax
    markdown_img = f"![{img_info['description']}]({img_info['relative_path']})\n\n*{img_info['description']}*"
    final_markdown = final_markdown.replace(img_info['original_placeholder'], markdown_img)

# Also handle any remaining simple [IMAGE: ...] patterns without coordinates
simple_image_pattern = r'\[IMAGE:\s*([^\]]+)\]'
final_markdown = re.sub(
    simple_image_pattern,
    lambda m: f"\n\n> 📷 **Visual Element:** {m.group(1).strip()}\n\n",
    final_markdown
)

# Generate output filename
output_filename = f"{input_filename}_out_ocr_simple.md"

# Create comprehensive markdown content
markdown_content = f"""# OCR Analysis with Sub-Images

**Source Image:** `{image_path}`  
**Processed:** {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}  
**Model:** Claude Sonnet 4.5 (Bedrock)  
**Sub-images extracted:** {len(sub_images_info)}  
**Sub-images directory:** `{output_dir}/`

---

{final_markdown}

---

## Extracted Sub-Images

"""

# Add a table of all sub-images
if sub_images_info:
    markdown_content += "\n| # | Description | Position | File |\n"
    markdown_content += "|---|-------------|----------|------|\n"
    for idx, img_info in enumerate(sub_images_info, 1):
        markdown_content += f"| {idx} | {img_info['description']} | {img_info['position']} | `{img_info['filename']}` |\n"
else:
    markdown_content += "\nNo sub-images were automatically extracted.\n"

markdown_content += f"""

---

## Technical Details

- **Original image dimensions:** {img_width} x {img_height} pixels
- **Total tokens used:** {response_body.get('usage', {}).get('total_tokens', 'N/A')}
- **Processing method:** AWS Bedrock with Claude Sonnet 4.5 vision analysis

*Generated by AWS Bedrock OCR with Sub-Image Extraction*
"""

# Save to markdown file
with open(output_filename, 'w', encoding='utf-8') as f:
    f.write(markdown_content)

print("=" * 80)
print("✓ Processing Complete!")
print("=" * 80)
print(f"Markdown output: {output_filename}")
print(f"Sub-images directory: {output_dir}/")
print(f"Sub-images extracted: {len(sub_images_info)}")
print("=" * 80)

