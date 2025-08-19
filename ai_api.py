import openai

def get_extraction_prompt(csv_data, shipper_list):
    """Generates the prompt for the AI to extract data from a CSV string."""
    return f"""
You are an expert data extraction AI. Your task is to analyze the following CSV data, which represents an entire Excel sheet, and extract structured information based on the rules provided.

**Rules:**

1.  **Identify Global Information:**
    *   `寄件廠商`: Scan the entire CSV. Find a cell that **exactly matches** one of the names in the provided "Valid Shipper List". This is the global shipper.
    *   `結單日期`: Scan the entire CSV. Find a cell containing keywords like '結單日', '結單日期', '訂購截止日', '最後回單日' and extract its corresponding date value.

2.  **Identify Product Rows:**
    *   First, determine the header row. The header contains titles like '品名', '貨號', '條碼', '東海成本'.
    *   Process each row below the header as a potential product.

3.  **Extract Product Data:** For each product row, extract the following based on the header and keywords:
    *   `國際條碼`: Keywords ['國際條碼', 'jan code', '條碼']
    *   `貨號`: Keywords ['sku', '貨號', '商品貨號']
    *   `品名`: Keywords ['品名', '商品名', '品項', '商品', '中文品名']
    *   `預計發售月份`: Keywords ['發售日', '預定到貨', '預計上市日', '發貨日']
    *   `備註`: Keywords ['備註', '備考', '附註', '註']
    *   `起始進價`: Keywords ['東海成本']
    *   `建議售價`: Keywords ['東海售價']

4.  **Filter Data:**
    *   Only keep product entries where both `起始進價` and `建議售價` have valid, non-empty values. Discard all others.

5.  **Output Format:**
    *   Return a single JSON object.
    *   The JSON object must have two top-level keys:
        1.  `global_info`: An object containing the `寄件廠商` and `結單日期` you found.
        2.  `products`: An array of objects, where each object is a product that passed the filtering rule.
    *   **CRITICAL**: Your entire response must be ONLY the JSON object, with no other text, explanations, or markdown formatting.

**Valid Shipper List:**
{shipper_list}

**CSV Data to Process:**
```csv
{csv_data}
```
"""

def call_ai_to_extract_data(client, csv_data, shipper_list, logger):
    """Calls the AI to extract structured JSON from CSV data."""
    if not client:
        logger("OpenAI client not configured. Please set your OPENAI_API_KEY environment variable.")
        return None
    if not csv_data:
        logger("No CSV data to process.")
        return None

    prompt = get_extraction_prompt(csv_data, shipper_list)
    logger("Calling OpenAI API for data extraction...")

    try:
        response = client.chat.completions.create(
            model="gpt-4o",
            messages=[
                {"role": "system", "content": "You are an AI assistant that extracts structured JSON from raw CSV data according to user-provided rules."},
                {"role": "user", "content": prompt},
            ],
            temperature=0,
            response_format={"type": "json_object"},
        )
        logger("Successfully received response from AI.")
        return response.choices[0].message.content
    except Exception as e:
        logger(f"Error calling OpenAI API: {e}")
        return None
