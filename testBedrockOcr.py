import boto3
import json

session = boto3.Session(
    region_name='ap-southeast-1'
)
bedrock = session.client('bedrock-runtime')

model_id = 'global.anthropic.claude-sonnet-4-5-20250929-v1:0'

request_body = {
    "anthropic_version": "bedrock-2023-05-31",
    "max_tokens": 1000,
    "messages": [
        {
            "role": "user",
            "content": "Hello! Can you tell me a short joke?"
        }
    ]
}

response = bedrock.invoke_model_with_response_stream(
    modelId=model_id,
    body=json.dumps(request_body),
    contentType='application/json'
)

print("Response: ", end='')
for event in response['body']:
    chunk = json.loads(event['chunk']['bytes'].decode())

    if chunk['type'] == 'content_block_delta':
        if 'text' in chunk['delta']:
            print(chunk['delta']['text'], end='', flush=True)
    elif chunk['type'] == 'message_stop':
        break

print()
 