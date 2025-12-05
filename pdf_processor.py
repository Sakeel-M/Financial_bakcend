"""
PDF processor for bank statements with AI-powered extraction
Supports text-based and image-based PDFs with intelligent data extraction
"""
import json
import re
from datetime import datetime
import os
import openai

try:
    import pdfplumber
    PDF_PLUMBER_AVAILABLE = True
except ImportError:
    PDF_PLUMBER_AVAILABLE = False

try:
    import PyPDF2
    PYPDF2_AVAILABLE = True
except ImportError:
    PYPDF2_AVAILABLE = False

try:
    import fitz  # PyMuPDF
    PYMUPDF_AVAILABLE = True
except ImportError:
    PYMUPDF_AVAILABLE = False

# Set OpenAI API key
openai.api_key = os.getenv('OPENAI_API_KEY')


class BankStatementPDFProcessor:
    def __init__(self):
        self.bank_patterns = {
            'ADCB': ['abu dhabi commercial bank', 'adcb'],
            'FAB': ['first abu dhabi bank', 'fab'],
            'ENBD': ['emirates nbd', 'enbd'],
            'MASHREQ': ['mashreq bank', 'mashreq'],
            'CBD': ['commercial bank of dubai', 'cbd'],
            'HSBC': ['hsbc'],
            'RAKBANK': ['rak bank', 'rakbank'],
            'ADIB': ['abu dhabi islamic bank', 'adib'],
            'BOA': ['bank of america', 'bofa', 'boa', 'b of a'],
            'CHASE': ['chase', 'jp morgan chase', 'jpmorgan'],
            'WELLS': ['wells fargo', 'wells'],
            'CITI': ['citibank', 'citi'],
            'BARCLAYS': ['barclays'],
            'LLOYDS': ['lloyds'],
            'SBI': ['state bank of india', 'sbi'],
            'HDFC': ['hdfc bank', 'hdfc'],
            'ICICI': ['icici bank', 'icici']
        }

        self.categories = {
            'Food & Dining': [
                'carrefour', 'lulu', 'spinneys', 'choithrams', 'union coop', 'waitrose',
                'restaurant', 'cafe', 'kfc', 'mcdonald', 'pizza', 'subway', 'dominos',
                'starbucks', 'costa', 'dunkin', 'burger', 'food', 'dining', 'eat',
                'grocery', 'supermarket', 'hypermarket', 'bakery', 'deli', 'bistro',
                'catering', 'takeaway', 'delivery', 'zomato', 'talabat', 'deliveroo',
                'doordash', 'grubhub', 'uber eats', 'postmates', 'chipotle', 'panera',
                'whole foods', 'trader joe', 'safeway', 'kroger', 'target', 'walmart',
                'donuts', 'coffee', 'espresso', 'bacio di latte', 'butchery', 'zinque',
                'ralphs', 'albertsons', 'vons', 'pavilions', 'publix', 'wegmans'
            ],
            'Transportation': [
                'adnoc', 'eppco', 'enoc', 'petrol', 'fuel', 'gas', 'gasoline',
                'taxi', 'uber', 'careem', 'metro', 'bus', 'rta', 'parking',
                'salik', 'toll', 'car wash', 'transport', 'emirates', 'etihad',
                'flydubai', 'air arabia', 'airline', 'flight', 'airport',
                'chevron', 'shell', 'bp', '76', 'exxon', 'mobil', 'lyft'
            ],
            'Shopping & Retail': [
                'mall', 'centrepoint', 'max', 'home centre', 'ikea', 'ace',
                'sharaf dg', 'jumbo', 'electronics', 'clothing', 'fashion',
                'shop', 'store', 'retail', 'amazon', 'noon', 'souq', 'namshi',
                'h&m', 'zara', 'nike', 'adidas', 'apple', 'samsung', 'virgin',
                'uniqlo', 'costco', 'best buy', 'macy', 'nordstrom'
            ],
            'Healthcare': [
                'hospital', 'clinic', 'pharmacy', 'medical', 'doctor', 'health',
                'dental', 'medicare', 'aster', 'nmc', 'mediclinic', 'life pharmacy',
                'boots', 'cvs', 'walgreens'
            ],
            'Utilities & Bills': [
                'dewa', 'addc', 'sewa', 'fewa', 'etisalat', 'du', 'internet',
                'mobile', 'telecom', 'electricity', 'water', 'utility', 'bill',
                'wifi', 'broadband', 'verizon', 'at&t', 't-mobile'
            ],
            'Entertainment': [
                'cinema', 'movie', 'vox', 'reel', 'netflix', 'osn', 'gaming',
                'entertainment', 'park', 'beach', 'attraction', 'ticket', 'event',
                'spotify', 'youtube', 'disney', 'hulu', 'disneyland'
            ],
            'Subscriptions & Digital Services': [
                'netflix', 'spotify', 'youtube premium', 'amazon prime', 'disney+',
                'adobe', 'microsoft', 'google', 'icloud', 'dropbox', 'zoom',
                'subscription', 'saas', 'office 365', 'rocket money'
            ],
            'ATM & Cash Withdrawals': [
                'atm', 'cash withdrawal', 'withdrawal', 'cash advance'
            ],
            'Banking & Finance': [
                'transfer', 'fee', 'charge', 'finance', 'loan', 'interest',
                'bank fee', 'service charge', 'wire transfer', 'zelle'
            ],
            'Personal Care': [
                'salon', 'spa', 'barbershop', 'beauty', 'cosmetics', 'skincare',
                'gym', 'fitness', '24hourfitness', 'planet fitness'
            ]
        }

    def detect_bank(self, text):
        """Detect bank from PDF content"""
        text_lower = text.lower()
        for bank_code, patterns in self.bank_patterns.items():
            for pattern in patterns:
                if pattern in text_lower:
                    return {
                        'ADCB': 'Abu Dhabi Commercial Bank',
                        'FAB': 'First Abu Dhabi Bank',
                        'ENBD': 'Emirates NBD',
                        'MASHREQ': 'Mashreq Bank',
                        'CBD': 'Commercial Bank of Dubai',
                        'HSBC': 'HSBC Bank Middle East',
                        'RAKBANK': 'RAKBank',
                        'ADIB': 'Abu Dhabi Islamic Bank',
                        'BOA': 'Bank of America',
                        'CHASE': 'Chase Bank',
                        'WELLS': 'Wells Fargo',
                        'CITI': 'Citibank',
                        'BARCLAYS': 'Barclays',
                        'LLOYDS': 'Lloyds Bank',
                        'SBI': 'State Bank of India',
                        'HDFC': 'HDFC Bank',
                        'ICICI': 'ICICI Bank'
                    }.get(bank_code, f'{bank_code} Bank')
        return 'Unknown Bank'

    def categorize_transaction(self, description):
        """Categorize transaction based on description"""
        desc_lower = description.lower().strip()

        # Priority-based categorization
        category_priority = [
            'ATM & Cash Withdrawals',
            'Subscriptions & Digital Services',
            'Food & Dining',
            'Transportation',
            'Healthcare',
            'Utilities & Bills',
            'Entertainment',
            'Shopping & Retail',
            'Personal Care',
            'Banking & Finance'
        ]

        for category in category_priority:
            keywords = self.categories[category]
            for keyword in keywords:
                if keyword in desc_lower:
                    return category

        return 'Other Expenses'

    def extract_text_from_pdf(self, pdf_file):
        """Extract text from PDF using multiple methods"""
        text_content = ""
        tables_data = []

        try:
            # Method 1: Try pdfplumber (best for tables)
            if PDF_PLUMBER_AVAILABLE:
                try:
                    pdf_file.seek(0)
                    with pdfplumber.open(pdf_file) as pdf:
                        for page_num, page in enumerate(pdf.pages, 1):
                            # Extract text
                            page_text = page.extract_text()
                            if page_text:
                                text_content += f"\n--- PAGE {page_num} ---\n{page_text}\n"

                            # Extract tables
                            tables = page.extract_tables()
                            if tables:
                                for table in tables:
                                    tables_data.append({
                                        'page': page_num,
                                        'data': table
                                    })

                    if text_content or tables_data:
                        print(f"[PDF] Extracted {len(pdf.pages)} pages using pdfplumber")
                        return text_content, tables_data
                except Exception as e:
                    print(f"[PDF] pdfplumber failed: {e}")

            # Method 2: Try PyMuPDF (good for text extraction)
            if PYMUPDF_AVAILABLE and not text_content:
                try:
                    pdf_file.seek(0)
                    doc = fitz.open(stream=pdf_file.read(), filetype="pdf")
                    for page_num in range(len(doc)):
                        page = doc[page_num]
                        text_content += f"\n--- PAGE {page_num + 1} ---\n{page.get_text()}\n"
                    doc.close()

                    if text_content:
                        print(f"[PDF] Extracted {len(doc)} pages using PyMuPDF")
                        return text_content, tables_data
                except Exception as e:
                    print(f"[PDF] PyMuPDF failed: {e}")

            # Method 3: Fallback to PyPDF2
            if PYPDF2_AVAILABLE and not text_content:
                try:
                    pdf_file.seek(0)
                    pdf_reader = PyPDF2.PdfReader(pdf_file)
                    for page_num in range(len(pdf_reader.pages)):
                        page = pdf_reader.pages[page_num]
                        text_content += f"\n--- PAGE {page_num + 1} ---\n{page.extract_text()}\n"

                    if text_content:
                        print(f"[PDF] Extracted {len(pdf_reader.pages)} pages using PyPDF2")
                        return text_content, tables_data
                except Exception as e:
                    print(f"[PDF] PyPDF2 failed: {e}")

        except Exception as e:
            print(f"[PDF] All text extraction methods failed: {e}")
            return "", []

        return text_content, tables_data

    def ai_extract_transactions(self, pdf_text, tables_data):
        """Use OpenAI to extract structured transaction data from PDF text"""
        if not openai.api_key:
            print("[AI] OpenAI API key not available")
            return None

        try:
            # Prepare comprehensive prompt with both text and table data
            table_info = ""
            if tables_data:
                table_info = f"\n\nEXTRACTED TABLES ({len(tables_data)} tables):\n"
                for i, table in enumerate(tables_data[:5], 1):  # Limit to first 5 tables
                    table_info += f"\nTable {i} (Page {table['page']}):\n"
                    for row in table['data'][:20]:  # Limit rows
                        table_info += f"{row}\n"

            prompt = f"""You are an expert at extracting transaction data from bank statements. Extract ALL transactions from this PDF bank statement.

PDF CONTENT:
{pdf_text[:8000]}
{table_info}

INSTRUCTIONS:
1. Extract ALL transactions with: Date, Description, Amount
2. Determine if amount is Debit (expense, negative) or Credit (income, positive)
3. Parse dates to YYYY-MM-DD format
4. Extract account holder name and account number if visible
5. Identify the bank name
6. Determine currency (AED, USD, EUR, GBP, INR, etc.)

CRITICAL RULES:
- Debits/Withdrawals/Payments = NEGATIVE amounts (use -100.00 format)
- Credits/Deposits/Income = POSITIVE amounts (use 100.00 format)
- Skip opening/closing balance rows
- Skip headers and summary rows
- Parse all date formats correctly (DD/MM/YYYY, MM/DD/YYYY, etc.)
- Include transaction reference/check numbers in description if available

OUTPUT FORMAT (JSON only, no markdown):
{{
  "bank_info": {{
    "bank_name": "Bank Name",
    "account_holder": "Holder Name or Unknown",
    "account_number": "Last 4 digits or Unknown",
    "currency": "USD or AED or EUR etc"
  }},
  "transactions": [
    {{
      "date": "2024-01-15",
      "description": "Transaction description",
      "amount": -50.00,
      "type": "Debit"
    }},
    {{
      "date": "2024-01-20",
      "description": "Salary deposit",
      "amount": 5000.00,
      "type": "Credit"
    }}
  ],
  "total_transactions": 0
}}

Extract ALL transactions you can find. Return ONLY valid JSON."""

            print("[AI] Calling OpenAI to extract transactions from PDF...")
            response = openai.chat.completions.create(
                model="gpt-4o-mini",
                messages=[
                    {
                        "role": "system",
                        "content": "You are a financial data extraction expert. Extract transaction data accurately and return ONLY valid JSON. No markdown, no explanations."
                    },
                    {"role": "user", "content": prompt}
                ],
                max_tokens=4000,
                temperature=0.1
            )

            ai_response = response.choices[0].message.content.strip()

            # Clean response
            if '```json' in ai_response:
                start_idx = ai_response.find('```json') + 7
                end_idx = ai_response.find('```', start_idx)
                ai_response = ai_response[start_idx:end_idx].strip()
            elif '```' in ai_response:
                ai_response = ai_response.replace('```', '').strip()

            # Parse JSON
            extracted_data = json.loads(ai_response)
            print(f"[AI] Successfully extracted {len(extracted_data.get('transactions', []))} transactions")

            return extracted_data

        except json.JSONDecodeError as e:
            print(f"[AI] JSON parsing error: {e}")
            print(f"[AI] Response preview: {ai_response[:200] if 'ai_response' in locals() else 'No response'}")
            return None
        except Exception as e:
            print(f"[AI] Error during AI extraction: {e}")
            return None

    def process_pdf_file(self, pdf_file):
        """Main PDF processing function"""
        print("[PDF] Starting PDF processing...")

        # Check if any PDF library is available
        if not (PDF_PLUMBER_AVAILABLE or PYPDF2_AVAILABLE or PYMUPDF_AVAILABLE):
            return None, "No PDF processing library available. Install pdfplumber, PyPDF2, or PyMuPDF."

        try:
            # Extract text and tables from PDF
            pdf_text, tables_data = self.extract_text_from_pdf(pdf_file)

            if not pdf_text and not tables_data:
                return None, "Could not extract text from PDF. File may be encrypted or image-based."

            # Use AI to extract structured transaction data
            extracted_data = self.ai_extract_transactions(pdf_text, tables_data)

            if not extracted_data or 'transactions' not in extracted_data:
                # Fallback: Try basic pattern matching
                print("[PDF] AI extraction failed, trying basic pattern matching...")
                extracted_data = self.fallback_extraction(pdf_text)

            if not extracted_data or not extracted_data.get('transactions'):
                return None, "Could not extract transactions from PDF. Please check the file format."

            # Process and categorize transactions
            transactions = []
            for tx in extracted_data['transactions']:
                try:
                    # Categorize
                    category = self.categorize_transaction(tx.get('description', ''))

                    transaction = {
                        'Date': tx['date'],
                        'Amount': float(tx['amount']),
                        'Description': tx['description'],
                        'Category': category,
                        'Subcategory': category.split(' ')[0] if category != 'Other Expenses' else 'Miscellaneous'
                    }
                    transactions.append(transaction)
                except Exception as e:
                    print(f"[PDF] Error processing transaction: {e}")
                    continue

            # Get bank info
            bank_info = extracted_data.get('bank_info', {
                'bank_name': 'Unknown Bank',
                'account_holder': 'Account Holder',
                'account_number': 'XXXX',
                'currency': 'USD'
            })

            # Apply AI categorization for better accuracy
            if len(transactions) > 0:
                print(f"[AI] Applying AI categorization to {len(transactions)} transactions...")
                transactions = self.ai_categorize_transactions(transactions)

            print(f"[SUCCESS] PDF processing complete: {len(transactions)} transactions extracted")

            return {
                'transactions': transactions,
                'bank_info': bank_info,
                'total_rows': len(transactions),
                'processing_mode': 'AI-Powered PDF Processing'
            }, None

        except Exception as e:
            import traceback
            print(f"[ERROR] PDF processing failed: {str(e)}")
            traceback.print_exc()
            return None, f"Error processing PDF: {str(e)}"

    def fallback_extraction(self, pdf_text):
        """Basic pattern matching fallback if AI fails"""
        print("[PDF] Using fallback extraction method...")

        transactions = []
        lines = pdf_text.split('\n')

        # Try to detect bank
        bank_name = self.detect_bank(pdf_text[:2000])

        # Simple pattern matching for common transaction formats
        # Pattern: Date Amount Description
        date_patterns = [
            r'(\d{1,2}[/-]\d{1,2}[/-]\d{2,4})',  # DD/MM/YYYY or MM/DD/YYYY
            r'(\d{4}[/-]\d{1,2}[/-]\d{1,2})'      # YYYY-MM-DD
        ]

        for line in lines:
            line = line.strip()
            if not line or len(line) < 10:
                continue

            # Skip headers and totals
            if any(word in line.upper() for word in ['TOTAL', 'BALANCE', 'OPENING', 'CLOSING', 'DATE', 'DESCRIPTION']):
                continue

            # Try to find date, amount, description
            for pattern in date_patterns:
                match = re.search(pattern, line)
                if match:
                    date_str = match.group(1)
                    # Look for amount
                    amount_match = re.search(r'[\-]?\d+[,.]?\d*\.?\d{2}', line)
                    if amount_match:
                        amount_str = amount_match.group(0).replace(',', '')
                        try:
                            amount = float(amount_str)
                            # Get description (remaining text)
                            description = re.sub(r'[\d/\-,.\s]+', ' ', line).strip()

                            if description:
                                transactions.append({
                                    'date': self.normalize_date(date_str),
                                    'description': description[:100],
                                    'amount': amount
                                })
                        except:
                            continue
                    break

        return {
            'bank_info': {
                'bank_name': bank_name,
                'account_holder': 'Unknown',
                'account_number': 'Unknown',
                'currency': 'USD'
            },
            'transactions': transactions[:100]  # Limit fallback results
        }

    def normalize_date(self, date_str):
        """Normalize date to YYYY-MM-DD format"""
        patterns = [
            ('%d/%m/%Y', r'\d{1,2}/\d{1,2}/\d{4}'),
            ('%m/%d/%Y', r'\d{1,2}/\d{1,2}/\d{4}'),
            ('%Y-%m-%d', r'\d{4}-\d{1,2}-\d{1,2}'),
            ('%d-%m-%Y', r'\d{1,2}-\d{1,2}-\d{4}'),
        ]

        for fmt, pattern in patterns:
            if re.match(pattern, date_str):
                try:
                    dt = datetime.strptime(date_str, fmt)
                    return dt.strftime('%Y-%m-%d')
                except:
                    continue

        return datetime.now().strftime('%Y-%m-%d')

    def ai_categorize_transactions(self, transactions):
        """Use AI to categorize transactions (same as Excel processor)"""
        if not openai.api_key or not transactions:
            return transactions

        categories = [
            'Food & Dining', 'Transportation', 'Shopping & Retail', 'Healthcare',
            'Utilities & Bills', 'Entertainment', 'Subscriptions & Digital Services',
            'ATM & Cash Withdrawals', 'Banking & Finance', 'Personal Care',
            'Travel', 'Income', 'Other Expenses'
        ]

        try:
            batch_size = 50
            categorized_transactions = []

            for i in range(0, len(transactions), batch_size):
                batch = transactions[i:i + batch_size]
                descriptions = [f"{idx}. {tx['Description']}" for idx, tx in enumerate(batch, start=i)]

                prompt = f"""Categorize these transactions into: {', '.join(categories)}

Transactions:
{chr(10).join(descriptions[:30])}

Return ONLY a JSON array of category names."""

                try:
                    response = openai.chat.completions.create(
                        model="gpt-4o-mini",
                        messages=[
                            {"role": "system", "content": "Return only JSON array of categories."},
                            {"role": "user", "content": prompt}
                        ],
                        max_tokens=500,
                        temperature=0.1
                    )

                    ai_response = response.choices[0].message.content.strip()
                    if '```json' in ai_response:
                        ai_response = ai_response.replace('```json', '').replace('```', '').strip()

                    ai_categories = json.loads(ai_response)

                    for idx, tx in enumerate(batch):
                        if idx < len(ai_categories):
                            tx['Category'] = ai_categories[idx]
                            tx['Subcategory'] = ai_categories[idx].split(' ')[0]
                        categorized_transactions.append(tx)

                except Exception as e:
                    print(f"[AI] Categorization error: {e}")
                    categorized_transactions.extend(batch)

            return categorized_transactions

        except Exception as e:
            print(f"[AI] Categorization failed: {e}")
            return transactions


# Global processor instance
pdf_processor = BankStatementPDFProcessor()
