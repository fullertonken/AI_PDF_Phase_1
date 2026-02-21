"""
prompts.py — All Ollama prompt templates for Phase 2.

These prompts are designed for a VISION model (qwen2.5-vl:7b) that receives
page images directly. The model can see the document layout, tables, logos,
and spatial structure — much better than working from OCR text alone.

Keeping prompts in one file makes them easy to tune and version.
"""

# ---------------------------------------------------------------------------
# System prompt (shared across all calls)
# ---------------------------------------------------------------------------
SYSTEM_PROMPT = """\
You are a document analysis assistant with vision capabilities. You analyze \
images of PDF pages from invoices and business documents. You can see the \
full document layout including tables, headers, logos, and text positioning. \
You always respond with valid JSON only — no markdown, no explanation, no \
extra text. Just the JSON object."""


# ---------------------------------------------------------------------------
# Step 1: Page Classification (vision — image sent alongside prompt)
# ---------------------------------------------------------------------------
CLASSIFY_PAGE_PROMPT = """\
Look at this document page image and classify it.

Respond with a JSON object:
{{
  "page_type": "<one of: invoice_first_page, invoice_continuation, supporting_document, unknown>",
  "confidence": <float 0.0 to 1.0>,
  "invoice_number": "<invoice number if visible, otherwise null>",
  "page_indicator": "<e.g. 'Page 1 of 3' if visible, otherwise null>",
  "reasoning": "<brief one-line explanation>"
}}

Classification rules:
- "invoice_first_page": The page is the first page of an invoice. It typically has \
a prominent "INVOICE" header, an invoice number, date, billing/shipping addresses, \
and line items or totals. Look for company logos, "Bill To" / "Ship To" sections, \
and tax breakdowns.
- "invoice_continuation": The page is a continuation of an invoice (e.g. page 2 of 3). \
It may have continued line items, subtotals, or totals but lacks a new invoice header. \
Look for "continued" labels, page numbers like "Page 2 of 3", or tables that appear \
to continue from a previous page.
- "supporting_document": The page is a packing slip, delivery receipt, purchase order, \
remittance advice, credit note, statement, or other document that is NOT an invoice. \
Look for headers like "Packing Slip", "Purchase Order", "Delivery Note", "Statement".
- "unknown": Cannot determine the page type from the image.

Respond with ONLY the JSON object."""


# ---------------------------------------------------------------------------
# Step 1b: Classification with OCR text backup
# When the vision model also has OCR text available for extra context.
# ---------------------------------------------------------------------------
CLASSIFY_PAGE_WITH_TEXT_PROMPT = """\
Look at this document page image and classify it. I have also extracted the \
following text from the page via OCR for additional context:

OCR TEXT:
---
{page_text}
---

Respond with a JSON object:
{{
  "page_type": "<one of: invoice_first_page, invoice_continuation, supporting_document, unknown>",
  "confidence": <float 0.0 to 1.0>,
  "invoice_number": "<invoice number if visible, otherwise null>",
  "page_indicator": "<e.g. 'Page 1 of 3' if visible, otherwise null>",
  "reasoning": "<brief one-line explanation>"
}}

Classification rules:
- "invoice_first_page": First page of an invoice — has invoice number, date, \
billing addresses, and line items or totals. Look for "INVOICE" header.
- "invoice_continuation": Continuation page — continued line items, subtotals, \
page numbers like "Page 2 of 3". No new invoice header.
- "supporting_document": Packing slip, PO, delivery note, statement, credit note, \
or other non-invoice document.
- "unknown": Cannot determine.

Respond with ONLY the JSON object."""


# ---------------------------------------------------------------------------
# Step 2: Invoice Grouping (text-only — uses classification results)
# This step doesn't need vision since it works on structured metadata.
# ---------------------------------------------------------------------------
GROUP_PAGES_PROMPT = """\
You are given a list of classified PDF pages in sequential order. \
Group them into invoices and supporting documents.

PAGES:
{pages_json}

Each page has: page_id, page_type, invoice_number, page_indicator, source_pdf, original_page.

Rules for grouping:
1. An "invoice_first_page" starts a new invoice group.
2. "invoice_continuation" pages following an invoice_first_page belong to the same invoice, \
UNLESS they have a different invoice_number.
3. "supporting_document" pages are attached to the most recent preceding invoice group.
4. If a supporting_document appears before any invoice, put it in a "pre_invoice_documents" group.
5. Pages from different source PDFs should generally be separate invoices unless \
invoice numbers match.

Respond with a JSON object:
{{
  "groups": [
    {{
      "group_id": 1,
      "group_type": "<invoice or supporting_documents>",
      "invoice_number": "<if known, otherwise null>",
      "page_ids": ["page_id_1", "page_id_2"]
    }}
  ]
}}

Respond with ONLY the JSON object."""


# ---------------------------------------------------------------------------
# Step 3: Field Extraction (vision — image(s) sent alongside prompt)
# For multi-page invoices, images are sent in sequence.
# ---------------------------------------------------------------------------
EXTRACT_FIELDS_PROMPT = """\
Look at this invoice document image(s) and extract all the details you can see.
If multiple images are provided, they are consecutive pages of the same invoice.

Respond with a JSON object containing these fields:
{{
  "invoice_number": "<string or null>",
  "invoice_date": "<YYYY-MM-DD format or null>",
  "due_date": "<YYYY-MM-DD format or null>",
  "supplier_name": "<string or null>",
  "supplier_address": "<full address string or null>",
  "customer_name": "<string or null>",
  "customer_address": "<full address string or null>",
  "purchase_order_number": "<string or null>",
  "currency": "<CAD, USD, etc. or null>",
  "line_items": [
    {{
      "description": "<string>",
      "quantity": <number or null>,
      "unit_price": <number or null>,
      "amount": <number or null>
    }}
  ],
  "subtotal": <number or null>,
  "shipping": <number or null>,
  "gst": <number or null>,
  "pst": <number or null>,
  "hst": <number or null>,
  "other_taxes": <number or null>,
  "total": <number or null>,
  "notes": "<any special notes, payment terms, or other relevant info — or null>",
  "extraction_confidence": <float 0.0 to 1.0>,
  "fields_uncertain": ["<list of field names where you are not confident>"]
}}

Rules:
- For dollar amounts, use plain numbers without $ signs (e.g. 1523.45 not "$1,523.45").
- If a field is not present or cannot be determined, use null.
- GST is the Canadian federal Goods and Services Tax (5%).
- PST is Provincial Sales Tax (varies by province).
- HST is Harmonized Sales Tax (used in some provinces instead of GST+PST).
- Look carefully at the table layout to correctly read line items with their \
quantities, unit prices, and amounts. Pay attention to column alignment.
- extraction_confidence should reflect how complete and reliable the extraction is.
- fields_uncertain should list any field names where you had to guess or \
the text was hard to read.

Respond with ONLY the JSON object."""


# ---------------------------------------------------------------------------
# Step 3b: Extraction with OCR text backup
# ---------------------------------------------------------------------------
EXTRACT_FIELDS_WITH_TEXT_PROMPT = """\
Look at this invoice document image(s) and extract all the details. \
I have also extracted the following text via OCR for reference:

OCR TEXT:
---
{invoice_text}
---

Use BOTH the image (for layout and visual context) AND the OCR text \
(for exact strings) to produce the most accurate extraction.

Respond with a JSON object containing these fields:
{{
  "invoice_number": "<string or null>",
  "invoice_date": "<YYYY-MM-DD format or null>",
  "due_date": "<YYYY-MM-DD format or null>",
  "supplier_name": "<string or null>",
  "supplier_address": "<full address string or null>",
  "customer_name": "<string or null>",
  "customer_address": "<full address string or null>",
  "purchase_order_number": "<string or null>",
  "currency": "<CAD, USD, etc. or null>",
  "line_items": [
    {{
      "description": "<string>",
      "quantity": <number or null>,
      "unit_price": <number or null>,
      "amount": <number or null>
    }}
  ],
  "subtotal": <number or null>,
  "shipping": <number or null>,
  "gst": <number or null>,
  "pst": <number or null>,
  "hst": <number or null>,
  "other_taxes": <number or null>,
  "total": <number or null>,
  "notes": "<any special notes, payment terms, or other relevant info — or null>",
  "extraction_confidence": <float 0.0 to 1.0>,
  "fields_uncertain": ["<list of field names where you are not confident>"]
}}

Rules:
- Dollar amounts as plain numbers (e.g. 1523.45).
- null for missing fields.
- GST = Canadian federal 5%, PST = provincial, HST = harmonized.
- Read the table carefully — use column alignment from the image.

Respond with ONLY the JSON object."""


# ---------------------------------------------------------------------------
# Fallback: simpler extraction for when full extraction fails
# ---------------------------------------------------------------------------
EXTRACT_KEY_FIELDS_PROMPT = """\
Look at this invoice image and extract only the key fields. Keep it simple.

Respond with a JSON object:
{{
  "invoice_number": "<string or null>",
  "invoice_date": "<YYYY-MM-DD or null>",
  "supplier_name": "<string or null>",
  "customer_name": "<string or null>",
  "subtotal": <number or null>,
  "gst": <number or null>,
  "pst": <number or null>,
  "total": <number or null>
}}

Respond with ONLY the JSON object."""