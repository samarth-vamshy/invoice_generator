"""Azure OpenAI fallback for general conversation only."""

from __future__ import annotations

from openai import AsyncAzureOpenAI

from invoice_settings import settings


SYSTEM_PROMPT = (
    "You are the Invoice Generator assistant. Answer concise, general questions about "
    "creating test or demonstration invoices. Do not claim that any invoice exists, was "
    "generated, or was saved. Do not invent facts about stored invoices or calculate invoice "
    "totals. If asked to create an invoice, direct the user to Generate Invoices. If asked "
    "about stored invoices, direct the user to Search Invoices; do not answer from memory."
)


async def answer_general_question(message: str) -> str:
    """Use the project's existing Azure endpoint, key, and deployment names."""
    if not all((settings.azure_openai_api_key, settings.azure_openai_endpoint, settings.azure_openai_deployment_name)):
        return "I can help create a test invoice. Choose Generate Invoices, or ask for Help."
    client = AsyncAzureOpenAI(
        api_version="2024-12-01-preview",
        api_key=settings.azure_openai_api_key,
        azure_endpoint=settings.azure_openai_endpoint,
        timeout=25.0,
    )
    try:
        result = await client.chat.completions.create(
            model=settings.azure_openai_deployment_name,
            messages=[{"role": "system", "content": SYSTEM_PROMPT}, {"role": "user", "content": message}],
        )
        return (result.choices[0].message.content or "").strip() or "I couldn't answer that. Try Help for available actions."
    finally:
        await client.close()
