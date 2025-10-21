"""
Test script for AWS Bedrock OCR integration with Excel Processor
Demonstrates how to use AWS_BEARER_TOKEN_BEDROCK for OCR on extracted Excel images
"""

import os
import json
from pathlib import Path
from bedrock_ocr import BedrockOCR


def test_single_image_ocr():
    """
    Test OCR on a single image
    """
    print("=== Testing Single Image OCR ===\n")
    
    # Initialize OCR with bearer token from environment
    # Make sure AWS_BEARER_TOKEN_BEDROCK is set
    bearer_token = os.getenv('AWS_BEARER_TOKEN_BEDROCK')
    if not bearer_token:
        print("WARNING: AWS_BEARER_TOKEN_BEDROCK not set in environment")
        print("Please set it using: export AWS_BEARER_TOKEN_BEDROCK='your-token-here'")
        return
    
    ocr = BedrockOCR(region_name='ap-southeast-1', bearer_token=bearer_token)
    
    # Test with one of the extracted Excel images
    test_image = "sheet2.png"
    
    if not os.path.exists(test_image):
        print(f"Test image not found: {test_image}")
        print("Please run excelprocessor.py first to extract images")
        return
    
    print(f"Processing: {test_image}")
    result = ocr.perform_ocr(test_image)
    
    if result['success']:
        print("\n✓ OCR Successful!")
        print(f"\nExtracted Text ({len(result['text'])} characters):")
        print("-" * 60)
        print(result['text'])
        print("-" * 60)
        print(f"\nModel: {result['model_id']}")
        print(f"Usage: {json.dumps(result['metadata']['usage'], indent=2)}")
    else:
        print(f"\n✗ OCR Failed: {result['error']}")


def test_batch_ocr():
    """
    Test batch OCR on multiple Excel images
    """
    print("\n\n=== Testing Batch OCR ===\n")
    
    bearer_token = os.getenv('AWS_BEARER_TOKEN_BEDROCK')
    if not bearer_token:
        print("WARNING: AWS_BEARER_TOKEN_BEDROCK not set")
        return
    
    ocr = BedrockOCR(region_name='ap-southeast-1', bearer_token=bearer_token)
    
    # Get all images from output directory
    images_dir = "output/images"
    if not os.path.exists(images_dir):
        print(f"Images directory not found: {images_dir}")
        return
    
    image_files = sorted([
        os.path.join(images_dir, f) 
        for f in os.listdir(images_dir) 
        if f.endswith(('.png', '.jpg', '.jpeg'))
    ])
    
    if not image_files:
        print("No images found in output/images")
        return
    
    print(f"Found {len(image_files)} images to process")
    
    # Process only first 3 for testing (to avoid high costs)
    test_images = image_files[:3]
    print(f"Processing first {len(test_images)} images for testing...\n")
    
    results = ocr.batch_ocr(test_images, verbose=True)
    
    # Save results to file
    output_file = "output/ocr_results.json"
    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(results, f, ensure_ascii=False, indent=2)
    
    print(f"\n✓ Results saved to: {output_file}")
    
    # Print summary
    successful = sum(1 for r in results if r['success'])
    print(f"\nSummary: {successful}/{len(results)} images processed successfully")


def test_streaming_ocr():
    """
    Test streaming OCR response
    """
    print("\n\n=== Testing Streaming OCR ===\n")
    
    bearer_token = os.getenv('AWS_BEARER_TOKEN_BEDROCK')
    if not bearer_token:
        print("WARNING: AWS_BEARER_TOKEN_BEDROCK not set")
        return
    
    ocr = BedrockOCR(region_name='ap-southeast-1', bearer_token=bearer_token)
    
    test_image = "output/images/S__img_1.png"
    
    if not os.path.exists(test_image):
        print(f"Test image not found: {test_image}")
        return
    
    print(f"Processing (streaming): {test_image}")
    print("\nExtracted text (streaming):")
    print("-" * 60)
    
    text = ocr.perform_ocr_streaming(test_image)
    
    print(text)
    print("-" * 60)
    print(f"\n✓ Extracted {len(text)} characters")


def process_excel_with_ocr():
    """
    Complete workflow: Extract Excel images and perform OCR
    """
    print("\n\n=== Complete Excel + OCR Workflow ===\n")
    
    # Check for bearer token
    bearer_token = os.getenv('AWS_BEARER_TOKEN_BEDROCK')
    if not bearer_token:
        print("ERROR: AWS_BEARER_TOKEN_BEDROCK not set")
        print("Set it using: export AWS_BEARER_TOKEN_BEDROCK='your-token-here'")
        return
    
    # Check if Excel file exists
    excel_file = "Sheet1.xlsx"
    if not os.path.exists(excel_file):
        print(f"Excel file not found: {excel_file}")
        print("Please provide an Excel file to process")
        return
    
    # Step 1: Extract images from Excel (if needed)
    import subprocess
    import sys
    
    output_dir = "output"
    if not os.path.exists(f"{output_dir}/images") or not os.listdir(f"{output_dir}/images"):
        print("Step 1: Extracting images from Excel...")
        result = subprocess.run(
            [sys.executable, "excelprocessor.py", "-i", excel_file, "-o", output_dir],
            capture_output=True,
            text=True
        )
        print(result.stdout)
        if result.returncode != 0:
            print(f"Error: {result.stderr}")
            return
    else:
        print("Step 1: Images already extracted, skipping...")
    
    # Step 2: Perform OCR on extracted images
    print("\nStep 2: Performing OCR on extracted images...")
    
    ocr = BedrockOCR(region_name='ap-southeast-1', bearer_token=bearer_token)
    
    images_dir = f"{output_dir}/images"
    image_files = sorted([
        os.path.join(images_dir, f) 
        for f in os.listdir(images_dir) 
        if f.endswith(('.png', '.jpg', '.jpeg'))
    ])
    
    # Process all images
    print(f"Found {len(image_files)} images to process\n")
    
    results = ocr.batch_ocr(
        image_files,
        prompt="Extract all text from this image. Preserve formatting and structure.",
        verbose=True
    )
    
    # Step 3: Save OCR results
    ocr_output_file = f"{output_dir}/ocr_results.json"
    with open(ocr_output_file, 'w', encoding='utf-8') as f:
        json.dump(results, f, ensure_ascii=False, indent=2)
    
    print(f"\n✓ OCR results saved to: {ocr_output_file}")
    
    # Step 4: Generate text outline with OCR data
    print("\nStep 3: Generating text outline with OCR data...")
    
    outline_lines = []
    for idx, result in enumerate(results, 1):
        if result['success']:
            image_name = Path(result['image_path']).name
            outline_lines.append(f"\n=== Image {idx}: {image_name} ===")
            outline_lines.append(result['text'])
    
    outline_file = f"{output_dir}/ocr_outline.txt"
    with open(outline_file, 'w', encoding='utf-8') as f:
        f.write("\n".join(outline_lines))
    
    print(f"✓ Text outline saved to: {outline_file}")
    
    # Summary
    successful = sum(1 for r in results if r['success'])
    total_chars = sum(len(r['text']) for r in results if r['success'])
    
    print(f"\n{'='*60}")
    print("SUMMARY")
    print(f"{'='*60}")
    print(f"Excel file: {excel_file}")
    print(f"Images extracted: {len(image_files)}")
    print(f"OCR processed: {successful}/{len(results)}")
    print(f"Total characters extracted: {total_chars}")
    print(f"Results: {ocr_output_file}")
    print(f"Outline: {outline_file}")
    print(f"{'='*60}")


if __name__ == "__main__":
    import sys
    
    # Check if bearer token is set
    if not os.getenv('AWS_BEARER_TOKEN_BEDROCK'):
        print("=" * 70)
        print("ERROR: AWS_BEARER_TOKEN_BEDROCK environment variable not set")
        print("=" * 70)
        print("\nPlease set your bearer token:")
        print("  export AWS_BEARER_TOKEN_BEDROCK='your-bearer-token-here'")
        print("\nOr set it in your script:")
        print("  os.environ['AWS_BEARER_TOKEN_BEDROCK'] = 'your-token'")
        print("=" * 70)
        sys.exit(1)
    
    # Run tests
    if len(sys.argv) > 1:
        if sys.argv[1] == "single":
            test_single_image_ocr()
        elif sys.argv[1] == "batch":
            test_batch_ocr()
        elif sys.argv[1] == "streaming":
            test_streaming_ocr()
        elif sys.argv[1] == "full":
            process_excel_with_ocr()
        else:
            print("Usage: python test_bedrock_ocr_integration.py [single|batch|streaming|full]")
    else:
        # Run all tests
        test_single_image_ocr()
        test_batch_ocr()
        test_streaming_ocr()

