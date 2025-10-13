import openai

def get_extraction_prompt(csv_data, shipper_list, brand_keywords, category_keywords):
    """Generates the prompt for the AI to extract data from a CSV string."""
    prompt = f"""
You are an expert data extraction AI. Your task is to analyze the following CSV data, which represents an entire Excel sheet, and extract structured information based on the rules provided.

**Rules:**

1.  **Identify Global Information:**
    *   `寄件廠商`: Scan the entire CSV. Find a cell that **exactly matches** one of the names in the provided "Valid Shipper List". This is the global shipper.
    *   `結單日期`: Scan the entire CSV. Find a cell containing keywords like '結單日', '結單日期', '訂購截止日', '最後回單日' and extract its corresponding date value.

2.  **Identify Product Rows:**
    *   First, determine the header row. The header contains titles like '品名', '貨號', '條碼', '東海成本'.
    *   Process each row below the header as a potential product.

3.  **Extract Product Data:** For each product row, extract the following based on the header. For each field, find the column with a matching header and extract the **value** from that column for the current product row. **Never return the header text itself.**
    *   `國際條碼`: Keywords [\'國際條碼\', \'jan code\', \'條碼\']
    *   `貨號`: Keywords [\'sku\', \'貨號\', \'商品貨號\']
    *   `品名`: Keywords [\'品名\', \'商品名\', \'品項\', \'商品\', \'中文品名\']
    *   `預計發售月份`: Keywords [\'發售日\', \'預定到貨\', \'預計上市日\', \'發貨日\']
    *   `備註`: Keywords [\'備註\', \'備考\', \'附註\', \'註\']
    *   `起始進價`: Find the column with a header like \'東海成本\' and extract its numeric value.
    *   `建議售價`: Find the column with a header like \'東海售價\' and extract its numeric value.
    *   `偵測到的品牌`: Analyze all columns for the product row. Find the most specific and correct brand name from the "Valid Brand Keyword List". Some products might mention both a manufacturer and a distributor (e.g., a \'FREEing\' product distributed by \'Good Smile Company\'). In such cases, prioritize the actual manufacturer (\'FREEing\') over the distributor. If only one brand from the list is mentioned, use that one. If no match, omit this field.
    *   `ai_matched_category_keyword`: This is a new, critical rule. Analyze the "Valid Category Keyword List". For each product, determine if its information (across all its columns) is associated with ALL the words in a given category keyword. For example, if a category keyword is "Good Smile Company 1/6", the product must be related to BOTH "Good Smile Company" AND "1/6" to be considered a match. If you find a confident match, return the original, full keyword (e.g., "Good Smile Company 1/6") in this field. Otherwise, omit this field.

4.  **Output Format:**
    *   Return a single JSON object.
    *   The JSON object must have two top-level keys:
        1.  `global_info`: An object containing the `寄件廠商` and `結單日期` you found.
        2.  `products`: An array of objects, where each object is a product that passed the filtering rule.
    *   **CRITICAL**: Your entire response must be ONLY the JSON object, with no other text, explanations, or markdown formatting.

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

    prompt += f"""
**CSV Data to Process:**
```csv
{csv_data}
```
"""
    return prompt

def call_ai_to_extract_data(client, csv_data, shipper_list, brand_keywords, category_keywords, logger):
    """Calls the AI to extract structured JSON from CSV data."""
    if not client:
        logger("OpenAI client not configured. Please set your OPENAI_API_KEY environment variable.")
        return None
    if not csv_data:
        logger("No CSV data to process.")
        return None

    prompt = get_extraction_prompt(csv_data, shipper_list, brand_keywords, category_keywords)
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