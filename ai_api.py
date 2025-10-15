import openai
import json

def get_enrichment_prompt(full_csv_data, pre_extracted_products, shipper_list, brand_keywords, category_keywords):
    """
    Generates a prompt for the AI to enrich pre-extracted data.
    The AI's job is to find global info and add semantic tags (brand/category) to products.
    """
    # Convert the list of product dicts to a compact JSON string for the prompt
    products_json_str = json.dumps(pre_extracted_products, ensure_ascii=False, indent=2)

    prompt = f"""
You are an expert data enrichment AI. I have already processed an Excel file and extracted the core product data. Your task is to analyze this pre-extracted data along with the full context of the original file to add semantic information.

**CONTEXT: FULL ORIGINAL FILE (in CSV format):**
```csv
{full_csv_data}
```

**PRE-EXTRACTED PRODUCTS:**
```json
{products_json_str}
```

**YOUR TASKS:**

1.  **Find Global Information:** From the **FULL ORIGINAL FILE CONTEXT** above, find the following:
    *   `寄件廠商`: Find a cell that **exactly matches** one of the names in the "Valid Shipper List".
    *   `結單日期`: Find a cell containing keywords like '結單日', '結單日期', '訂購截止日', '最後回單日'. Extract its corresponding date value and **format it as "YYYY-MM-DD"**.

2.  **Enrich Product Data:** For each product in the **PRE-EXTRACTED PRODUCTS** list, perform the following analysis based on all available information:
    *   `預計發售月份`: Analyze the value of this field. It can be in various formats (e.g., "2026年3月底", "2025-11-01 00:00:00", "2025.11"). Your task is to parse it and **replace its original value** with the standardized `YYYY-MM` format.
    *   `偵測到的品牌`: Find the most specific and correct brand name from the "Valid Brand Keyword List". Prioritize the actual manufacturer over the distributor (e.g., 'FREEing' over 'Good Smile Company' if both are present).
    *   `ai_matched_category_keyword`: Analyze the "Valid Category Keyword List". Find the best keyword where the product information is associated with ALL the words in the category keyword.

**OUTPUT FORMAT:**

*   Return a single JSON object.
*   The JSON object must have two top-level keys:
    1.  `global_info`: An object containing the `寄件廠商` and `結單日期` you found.
    2.  `products`: An array of objects. Each object must be one of the products from the input. You will add `偵測到的品牌` and `ai_matched_category_keyword`. You **must** also update the `預計發售月份` field with the normalized value. **Do not alter any other original fields.**
*   **CRITICAL**: Your entire response must be ONLY the JSON object, with no other text, explanations, or markdown formatting.

**VALID LISTS FOR MATCHING:**

**Valid Shipper List:**
{shipper_list}
"""
    if brand_keywords:
        prompt += f"""
**Valid Brand Keyword List:**
{brand_keywords}
"""

    if category_keywords:
        prompt += f"""
**Valid Category Keyword List:**
{category_keywords}
"""
    return prompt

def call_ai_for_enrichment(client, full_csv_data, pre_extracted_products, shipper_list, brand_keywords, category_keywords, logger, debug_path_prefix=None):
    """Calls the AI to enrich pre-extracted product data."""
    if not client:
        logger("OpenAI client not configured. Please set your OPENAI_API_KEY.")
        return None
    if not pre_extracted_products:
        logger("No pre-extracted products to enrich.")
        return None

    prompt = get_enrichment_prompt(full_csv_data, pre_extracted_products, shipper_list, brand_keywords, category_keywords)

    if debug_path_prefix:
        try:
            prompt_filename = f"{debug_path_prefix}_enrichment_prompt.txt"
            with open(prompt_filename, 'w', encoding='utf-8') as f:
                f.write(prompt)
            logger(f"AI enrichment prompt saved to: {prompt_filename}")
        except Exception as e:
            logger(f"Error saving AI enrichment prompt: {e}")
    
    logger("Calling OpenAI API for data enrichment...")

    try:
        response = client.chat.completions.create(
            model="gpt-4o",
            messages=[
                {"role": "system", "content": "You are an AI assistant that enriches structured JSON data based on context and rules."},
                {"role": "user", "content": prompt},
            ],
            temperature=0,
            response_format={"type": "json_object"},
        )
        logger("Successfully received response from AI for enrichment.")
        return response.choices[0].message.content
    except Exception as e:
        logger(f"Error calling OpenAI API for enrichment: {e}")
        return None
