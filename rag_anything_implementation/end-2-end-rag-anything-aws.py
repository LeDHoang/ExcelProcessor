import asyncio
import os
import json
import base64
from dotenv import load_dotenv
from raganything import RAGAnything, RAGAnythingConfig
from lightrag.utils import EmbeddingFunc
import boto3
import botocore

# Load environment variables from .env file
load_dotenv()

async def main():
    # Set up AWS Bedrock configuration from .env file
    region = os.getenv("BEDROCK_REGION", "ap-southeast-1")
    llm_model_id = os.getenv(
        "BEDROCK_LLM_MODEL_ID",
        "global.anthropic.claude-sonnet-4-5-20250929-v1:0",
    )
    vision_model_id = os.getenv(
        "BEDROCK_VISION_MODEL_ID",
        llm_model_id,
    )
    embedding_model_id = os.getenv(
        "BEDROCK_EMBEDDING_MODEL_ID",
        "amazon.titan-embed-text-v2:0",
    )
    embedding_dim = int(os.getenv("BEDROCK_EMBEDDING_DIM", "1024"))

    # Initialize Bedrock client using bearer token (session token)
    session = boto3.Session(
        aws_access_key_id=os.environ.get("AWS_ACCESS_KEY_ID"),
        aws_secret_access_key=os.environ.get("AWS_SECRET_ACCESS_KEY"),
        aws_session_token=os.environ.get("AWS_BEARER_TOKEN_BEDROCK"),
        region_name=region,
    )
    bedrock = session.client("bedrock-runtime")

    print(f"Using Bedrock region: {region}")
    print(f"Using LLM model ID: {llm_model_id}")
    print(f"Using vision model ID: {vision_model_id}")
    print(f"Using embedding model ID: {embedding_model_id}")
    print(f"Embedding dimension: {embedding_dim}")
    print(
        f"AWS bearer token provided: {'yes' if os.environ.get('AWS_BEARER_TOKEN_BEDROCK') else 'no'}"
    )

    # Create RAGAnything configuration
    # Use a per-embedding-dimension working dir to avoid index dim mismatch with old caches
    config = RAGAnythingConfig(
        working_dir=f"./rag_storage_{embedding_dim}",
        parser="mineru",  # Parser selection: mineru or docling
        parse_method="auto",  # Parse method: auto, ocr, or txt
        enable_image_processing=True,
        enable_table_processing=True,
        enable_equation_processing=True,
    )

    # Helpers to invoke Anthropic (Claude) via AWS Bedrock
    def _extract_text_from_anthropic_response(resp_json):
        text_out = ""
        for block in resp_json.get("content", []):
            if block.get("type") == "text":
                text_out += block.get("text", "")
        return text_out

    def _normalize_history_to_anthropic(history_messages):
        messages = []
        for m in history_messages or []:
            role = m.get("role", "user")
            content = m.get("content", "")
            # Flatten OpenAI-style content into plain text blocks
            if isinstance(content, list):
                parts = []
                for p in content:
                    if isinstance(p, dict) and p.get("type") == "text":
                        parts.append(p.get("text", ""))
                    elif isinstance(p, str):
                        parts.append(p)
                content_text = "\n".join([t for t in parts if t])
            else:
                content_text = str(content)
            messages.append(
                {
                    "role": "assistant" if role == "assistant" else "user",
                    "content": [{"type": "text", "text": content_text}],
                }
            )
        return messages

    def _invoke_anthropic_chat(model_id, system, messages, **kwargs):
        body = {
            "anthropic_version": "bedrock-2023-05-31",
            "max_tokens": int(kwargs.get("max_tokens", 1024)),
            "messages": messages,
        }
        if system:
            body["system"] = system
        if "temperature" in kwargs and kwargs["temperature"] is not None:
            body["temperature"] = float(kwargs["temperature"])
        if "top_p" in kwargs and kwargs["top_p"] is not None:
            body["top_p"] = float(kwargs["top_p"])

        response = bedrock.invoke_model(
            modelId=model_id,
            body=json.dumps(body).encode("utf-8"),
            contentType="application/json",
            accept="application/json",
        )
        resp_json = json.loads(response["body"].read())
        return _extract_text_from_anthropic_response(resp_json)

    # Define LLM model function (text-only)
    async def llm_model_func(prompt, system_prompt=None, history_messages=[], **kwargs):
        messages = _normalize_history_to_anthropic(history_messages)
        # Append current user prompt
        user_blocks = []
        if prompt:
            user_blocks.append({"type": "text", "text": prompt})
        messages.append({"role": "user", "content": user_blocks or [{"type": "text", "text": ""}]})
        return await asyncio.to_thread(_invoke_anthropic_chat, llm_model_id, system_prompt, messages, **kwargs)

    # Define vision model function for image processing
    async def vision_model_func(
        prompt, system_prompt=None, history_messages=[], image_data=None, messages=None, **kwargs
    ):
        def _parse_openai_messages_to_anthropic(openai_messages):
            anthro_messages = []
            system_chunks = []
            for m in openai_messages or []:
                role = m.get("role", "user")
                content = m.get("content", "")

                if role == "system":
                    if isinstance(content, list):
                        for c in content:
                            if isinstance(c, dict) and c.get("type") == "text":
                                system_chunks.append(c.get("text", ""))
                            elif isinstance(c, str):
                                system_chunks.append(c)
                    else:
                        system_chunks.append(str(content))
                    continue

                blocks = []
                if isinstance(content, str):
                    blocks.append({"type": "text", "text": content})
                elif isinstance(content, list):
                    for c in content:
                        if isinstance(c, dict) and c.get("type") == "text":
                            blocks.append({"type": "text", "text": c.get("text", "")})
                        elif isinstance(c, dict) and c.get("type") == "image_url":
                            img = c.get("image_url")
                            url = img.get("url") if isinstance(img, dict) else img
                            media_type = "image/jpeg"
                            data_b64 = None
                            if isinstance(url, str) and url.startswith("data:"):
                                # Format: data:image/png;base64,XXXXX
                                try:
                                    header, b64data = url.split(",", 1)
                                    if ";base64" in header:
                                        mt = header.split(":", 1)[1].split(";")[0]
                                        media_type = mt
                                        data_b64 = b64data
                                except Exception:
                                    data_b64 = None
                            if data_b64:
                                blocks.append(
                                    {
                                        "type": "image",
                                        "source": {
                                            "type": "base64",
                                            "media_type": media_type,
                                            "data": data_b64,
                                        },
                                    }
                                )
                        # Ignore other block types for now
                anthro_messages.append(
                    {
                        "role": "assistant" if role == "assistant" else "user",
                        "content": blocks or [{"type": "text", "text": ""}],
                    }
                )
            system_text = "\n".join([s for s in system_chunks if s]) if system_chunks else None
            return anthro_messages, system_text

        # If messages format is provided (for multimodal VLM enhanced query), convert and invoke
        if messages:
            anthro_messages, sys_from_msgs = _parse_openai_messages_to_anthropic(messages)
            final_system = system_prompt or sys_from_msgs
            return await asyncio.to_thread(_invoke_anthropic_chat, vision_model_id, final_system, anthro_messages, **kwargs)

        # Traditional single image format
        if image_data:
            media_type = kwargs.get("media_type", "image/jpeg")
            content_blocks = []
            if prompt:
                content_blocks.append({"type": "text", "text": prompt})
            content_blocks.append(
                {
                    "type": "image",
                    "source": {
                        "type": "base64",
                        "media_type": media_type,
                        "data": image_data,
                    },
                }
            )
            messages_payload = [
                {"role": "user", "content": content_blocks},
            ]
            return await asyncio.to_thread(_invoke_anthropic_chat, vision_model_id, system_prompt, messages_payload, **kwargs)

        # Pure text format
        return await llm_model_func(prompt, system_prompt, history_messages, **kwargs)

    # Define embedding function (Titan v2 via Bedrock)
    # Resolve embedding model with graceful fallback to v1 if v2 is unavailable in-region
    selected_embedding_model_id = embedding_model_id
    supports_dimensions = "titan-embed-text-v2" in selected_embedding_model_id

    def _bedrock_embed_titan_v2(texts):
        vectors = []
        for t in texts:
            nonlocal selected_embedding_model_id, supports_dimensions
            body = {"inputText": t}
            # Only Titan v2 supports configurable dimensions
            if supports_dimensions and embedding_dim:
                body["dimensions"] = embedding_dim
            try:
                resp = bedrock.invoke_model(
                    modelId=selected_embedding_model_id,
                    body=json.dumps(body).encode("utf-8"),
                    contentType="application/json",
                    accept="application/json",
                )
            except botocore.exceptions.ClientError as e:
                err = e.response.get("Error", {})
                code = err.get("Code", "")
                message = (err.get("Message", "") or str(e)).lower()
                invalid_model = "model identifier is invalid" in message or code in (
                    "ValidationException",
                    "ResourceNotFoundException",
                )
                if invalid_model and selected_embedding_model_id != "amazon.titan-embed-text-v1":
                    # Fallback to Titan v1 (no dimensions parameter)
                    selected_embedding_model_id = "amazon.titan-embed-text-v1"
                    supports_dimensions = False
                    body = {"inputText": t}
                    resp = bedrock.invoke_model(
                        modelId=selected_embedding_model_id,
                        body=json.dumps(body).encode("utf-8"),
                        contentType="application/json",
                        accept="application/json",
                    )
                else:
                    raise

            resp_json = json.loads(resp["body"].read())
            vec = resp_json.get("embedding") or resp_json.get("embeddings")
            if isinstance(vec, list) and vec and isinstance(vec[0], list):
                # Some providers return [[...]]
                vec = vec[0]
            # Normalize vector length to configured embedding_dim to avoid downstream mismatches
            if isinstance(vec, list) and embedding_dim:
                try:
                    if len(vec) > embedding_dim:
                        vec = vec[:embedding_dim]
                    elif len(vec) < embedding_dim:
                        vec = vec + [0.0] * (embedding_dim - len(vec))
                except Exception:
                    # If vec has no len or isn't list-like, leave as is
                    pass
            vectors.append(vec)
        return vectors

    async def _bedrock_embed_async(texts):
        # Wrap sync Bedrock calls in a thread so caller can await
        return await asyncio.to_thread(_bedrock_embed_titan_v2, texts)

    embedding_func = EmbeddingFunc(
        embedding_dim=embedding_dim,
        max_token_size=8192,
        func=_bedrock_embed_async,
    )

    # Initialize RAGAnything
    rag = RAGAnything(
        config=config,
        llm_model_func=llm_model_func,
        vision_model_func=vision_model_func,
        embedding_func=embedding_func,
    )

    # Process a document
    await rag.process_document_complete(
        file_path="input/fullsheets-vba_optimized.pdf",
        output_dir="output",
        parse_method="auto"
    )

    # Query the processed content
    # Pure text query - for basic knowledge base search
    text_result = await rag.aquery(
        "Tell me the list of countries inside the sanctioned country list in the document",
        mode="hybrid"
    )
    print("Text query result:", text_result)

    # Multimodal query with specific multimodal content
    multimodal_result = await rag.aquery_with_multimodal(
        "Tell me the list of countries inside the sanctioned country list in the document",
        multimodal_content=[
            {
                "type": "equation",
                "latex": "P(d|q) = \\frac{P(q|d) \\cdot P(d)}{P(q)}",
                "equation_caption": "Document relevance probability",
            }
        ],
        mode="hybrid",
    )
    print("Multimodal query result:", multimodal_result)

if __name__ == "__main__":
    asyncio.run(main())