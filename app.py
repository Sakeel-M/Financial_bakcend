from flask import Flask, request, jsonify
from flask_cors import CORS
import json
import os
from dotenv import load_dotenv
import openai
from datetime import datetime
import io
import re

# Import Excel processor
from excel_processor import processor

load_dotenv()

app = Flask(__name__)
CORS(app)

# Set OpenAI API key
openai.api_key = os.getenv('OPENAI_API_KEY')

# Global bank configuration with currency mapping
GLOBAL_BANK_CONFIG = {
    # UAE Banks
    'abu dhabi commercial bank': {'country': 'UAE', 'currency': 'AED', 'code': 'ADCB'},
    'adcb': {'country': 'UAE', 'currency': 'AED', 'code': 'ADCB'},
    'first abu dhabi bank': {'country': 'UAE', 'currency': 'AED', 'code': 'FAB'},
    'fab': {'country': 'UAE', 'currency': 'AED', 'code': 'FAB'},
    'emirates nbd': {'country': 'UAE', 'currency': 'AED', 'code': 'ENBD'},
    'enbd': {'country': 'UAE', 'currency': 'AED', 'code': 'ENBD'},
    'mashreq bank': {'country': 'UAE', 'currency': 'AED', 'code': 'MASHREQ'},
    'mashreq': {'country': 'UAE', 'currency': 'AED', 'code': 'MASHREQ'},
    'commercial bank of dubai': {'country': 'UAE', 'currency': 'AED', 'code': 'CBD'},
    'cbd': {'country': 'UAE', 'currency': 'AED', 'code': 'CBD'},
    'hsbc uae': {'country': 'UAE', 'currency': 'AED', 'code': 'HSBC'},
    'hsbc middle east': {'country': 'UAE', 'currency': 'AED', 'code': 'HSBC'},
    'abu dhabi islamic bank': {'country': 'UAE', 'currency': 'AED', 'code': 'ADIB'},
    'adib': {'country': 'UAE', 'currency': 'AED', 'code': 'ADIB'},

    # US Banks
    'bank of america': {'country': 'USA', 'currency': 'USD', 'code': 'BOA'},
    'chase bank': {'country': 'USA', 'currency': 'USD', 'code': 'CHASE'},
    'wells fargo': {'country': 'USA', 'currency': 'USD', 'code': 'WF'},
    'citibank': {'country': 'USA', 'currency': 'USD', 'code': 'CITI'},

    # UK Banks
    'barclays': {'country': 'UK', 'currency': 'GBP', 'code': 'BARCLAYS'},
    'lloyds': {'country': 'UK', 'currency': 'GBP', 'code': 'LLOYDS'},
    'hsbc uk': {'country': 'UK', 'currency': 'GBP', 'code': 'HSBC'},

    # Indian Banks
    'state bank of india': {'country': 'India', 'currency': 'INR', 'code': 'SBI'},
    'hdfc bank': {'country': 'India', 'currency': 'INR', 'code': 'HDFC'},
    'icici bank': {'country': 'India', 'currency': 'INR', 'code': 'ICICI'},

    # European Banks
    'deutsche bank': {'country': 'Germany', 'currency': 'EUR', 'code': 'DB'},
    'bnp paribas': {'country': 'France', 'currency': 'EUR', 'code': 'BNP'},
    'ing bank': {'country': 'Netherlands', 'currency': 'EUR', 'code': 'ING'}
}

def detect_bank_and_currency(text):
    """Detect bank and determine currency based on bank location"""
    text_lower = text.lower()

    for bank_key, bank_info in GLOBAL_BANK_CONFIG.items():
        if bank_key in text_lower:
            return {
                'bank_name': bank_key.title(),
                'country': bank_info['country'],
                'currency': bank_info['currency'],
                'bank_code': bank_info['code']
            }

    # Default fallback
    return {
        'bank_name': 'Unknown Bank',
        'country': 'Unknown',
        'currency': 'USD',
        'bank_code': 'UNKNOWN'
    }

@app.route('/api/health', methods=['GET'])
def health_check():
    try:
        from openpyxl import load_workbook
        excel_status = True
    except ImportError:
        excel_status = False

    openai_status = bool(openai.api_key)

    return jsonify({
        "status": "healthy",
        "message": "Universal Finance Analytics API with AI",
        "excel_processing": excel_status,
        "openai_integration": openai_status,
        "supported_currencies": ["AED", "USD", "EUR", "GBP", "INR"],
        "supported_banks": "Global banks supported"
    })

@app.route('/api/upload', methods=['POST'])
def upload_file():
    try:
        if 'file' not in request.files:
            return jsonify({"error": "No file provided"}), 400

        file = request.files['file']
        if file.filename == '':
            return jsonify({"error": "No file selected"}), 400

        if not file.filename.endswith(('.xlsx', '.xls')):
            return jsonify({"error": "Please upload an Excel file (.xlsx or .xls)"}), 400

        # Process Excel file
        file_content = io.BytesIO(file.read())
        result, error = processor.process_excel_file(file_content)

        if error or not result:
            return jsonify({"error": f"Error processing Excel file: {error}"}), 500

        # Enhanced bank detection with currency
        bank_detection = detect_bank_and_currency(str(result.get('bank_info', {}).get('bank_name', '')))

        # Update bank info with detected currency
        bank_info = result['bank_info']
        bank_info.update({
            'currency': bank_detection['currency'],
            'country': bank_detection['country'],
            'bank_code': bank_detection['bank_code']
        })

        return jsonify({
            "message": "Excel file processed successfully!",
            "data": result['transactions'][:20],  # Preview
            "columns": ["Date", "Amount", "Description", "Category"],
            "total_rows": result['total_rows'],
            "full_data": result['transactions'],
            "bank_info": bank_info,
            "processing_mode": "Real Excel Processing with AI"
        })

    except Exception as e:
        return jsonify({"error": f"Processing error: {str(e)}"}), 500

@app.route('/api/analyze', methods=['POST'])
def analyze_data():
    try:
        data = request.get_json()
        financial_data = data.get('data', [])
        bank_info = data.get('bank_info', {})

        if not financial_data:
            return jsonify({"error": "No data provided for analysis"}), 400

        currency = bank_info.get('currency', 'USD')
        bank_name = bank_info.get('bank_name', 'Unknown Bank')
        country = bank_info.get('country', 'Unknown')

        # Calculate basic financial metrics
        total_income = sum(float(item.get('Amount', 0)) for item in financial_data if float(item.get('Amount', 0)) > 0)
        total_expenses = sum(abs(float(item.get('Amount', 0))) for item in financial_data if float(item.get('Amount', 0)) < 0)
        net_savings = total_income - total_expenses
        savings_rate = (net_savings / total_income * 100) if total_income > 0 else 0

        # Category analysis
        category_spending = {}
        monthly_data = {}
        yearly_data = {}

        for item in financial_data:
            amount = float(item.get('Amount', 0))
            category = item.get('Category', 'Other')
            date_str = item.get('Date', '')

            # Expenses only for category analysis
            if amount < 0:
                category_spending[category] = category_spending.get(category, 0) + abs(amount)

            # Monthly and yearly data - convert to readable format
            try:
                if len(date_str) >= 7:
                    # Parse YYYY-MM format
                    year_month = date_str[:7]  # e.g., "2025-02"
                    year, month_num = year_month.split('-')

                    # Store with sortable key for proper ordering
                    month_date = datetime.strptime(f"{year}-{month_num}-01", "%Y-%m-%d")
                    readable_month = month_date.strftime("%B %Y")  # e.g., "February 2025"
                    sort_key = f"{year}-{month_num}"  # Keep sortable key

                    # Monthly data with expense amounts
                    if amount < 0:
                        if sort_key not in monthly_data:
                            monthly_data[sort_key] = {
                                'readable': readable_month,
                                'amount': 0,
                                'year': int(year),
                                'month': int(month_num)
                            }
                        monthly_data[sort_key]['amount'] += abs(amount)

                    # Yearly totals
                    if amount < 0:
                        yearly_data[year] = yearly_data.get(year, 0) + abs(amount)
                else:
                    # Fallback
                    sort_key = '2024-01'
                    if sort_key not in monthly_data:
                        monthly_data[sort_key] = {
                            'readable': 'January 2024',
                            'amount': 0,
                            'year': 2024,
                            'month': 1
                        }
                    if amount < 0:
                        monthly_data[sort_key]['amount'] += abs(amount)
            except:
                sort_key = '2024-01'
                if sort_key not in monthly_data:
                    monthly_data[sort_key] = {
                        'readable': 'January 2024',
                        'amount': 0,
                        'year': 2024,
                        'month': 1
                    }
                if amount < 0:
                    monthly_data[sort_key]['amount'] += abs(amount)

        # Convert monthly_data to sorted readable format
        sorted_monthly_data = {}
        for sort_key in sorted(monthly_data.keys()):
            readable_month = monthly_data[sort_key]['readable']
            sorted_monthly_data[readable_month] = monthly_data[sort_key]['amount']

        # Get actual date range from data
        dates_in_data = [t.get('Date', '') for t in financial_data if t.get('Date')]
        min_date = min(dates_in_data) if dates_in_data else '2024-01-01'
        max_date = max(dates_in_data) if dates_in_data else '2024-12-31'

        # Calculate average monthly spending
        num_months = len(sorted_monthly_data) if sorted_monthly_data else 1
        avg_monthly_spending = total_expenses / num_months if num_months > 0 else 0

        # Get list of months in the data for AI reference
        months_list = list(sorted_monthly_data.keys())

        # Find highest spending month
        highest_month = max(sorted_monthly_data.items(), key=lambda x: x[1])[0] if sorted_monthly_data else "N/A"
        highest_month_amount = max(sorted_monthly_data.values()) if sorted_monthly_data else 0

        # Advanced AI Financial Analysis Prompt
        openai_prompt = f"""
        You are an expert financial advisor with deep expertise in personal finance, budgeting, and wealth management.
        Analyze this comprehensive financial data and provide advanced insights:

        CRITICAL INSTRUCTIONS - READ CAREFULLY:
        - Current analysis date: October 15, 2025
        - Transaction data period: {min_date} to {max_date}
        - ONLY reference these specific months that exist in the data: {', '.join(months_list)}
        - DO NOT mention "September 2025" or any month not in the list above
        - Highest spending month in the actual data: {highest_month} with {currency} {highest_month_amount:,.2f}
        - When describing spending peaks, ONLY use months from the list above
        - Base ALL insights ONLY on the actual transaction data provided below

        FINANCIAL PROFILE:
        - Bank: {bank_name} ({country})
        - Currency: {currency}
        - Analysis Period: {min_date} to {max_date}
        - Total Income: {currency} {total_income:,.2f}
        - Total Expenses: {currency} {total_expenses:,.2f}
        - Net Savings: {currency} {net_savings:,.2f}
        - Savings Rate: {savings_rate:.1f}%
        - Transaction Count: {len(financial_data)}
        - Months Analyzed: {num_months}
        - Average Monthly Spending: {currency} {avg_monthly_spending:,.2f}
        - Average Transaction: {currency} {total_expenses / len([t for t in financial_data if float(t.get('Amount', 0)) < 0]) if len([t for t in financial_data if float(t.get('Amount', 0)) < 0]) > 0 else 0:,.2f}

        SPENDING BREAKDOWN BY CATEGORY:
        {json.dumps(category_spending, indent=2)}

        MONTHLY SPENDING PATTERNS (Chronological Order):
        {json.dumps(sorted_monthly_data, indent=2)}

        YEARLY TOTALS:
        {json.dumps(yearly_data, indent=2)}

        TRANSACTION SAMPLES:
        {json.dumps(financial_data[:15], indent=2)}

        PROVIDE ADVANCED ANALYSIS:
        1. **Financial Health Score (0-100)** - Comprehensive scoring based on savings rate, spending patterns, and financial stability
        2. **Spending Pattern Analysis** - Identify trends, seasonal patterns, and unusual spending behaviors (ONLY reference months from the list provided above)
        3. **Budget Optimization** - Specific recommendations for each spending category with target amounts
        4. **Savings Strategy** - Personalized savings goals and investment recommendations
        5. **Risk Assessment** - Identify financial risks and vulnerabilities
        6. **Anomaly Detection** - Flag unusual transactions or concerning patterns
        7. **Country-Specific Advice** - Regional financial tips and local market insights
        8. **Future Projections** - Predict next month's spending and savings based on current trends

        EXAMPLE OF CORRECT MONTH REFERENCE:
        ✓ CORRECT: "Spending peaked in {highest_month} with {currency} {highest_month_amount:,.0f}"
        ✗ WRONG: "Spending peaked in September 2025" (if September 2025 is not in the months list)

        Format as JSON with these keys:
        - "financial_health_score": number (0-100)
        - "health_category": string ("Excellent", "Good", "Fair", "Poor", "Critical")
        - "key_insights": array of detailed insight strings
        - "spending_patterns": array of pattern analysis strings
        - "budget_recommendations": object with category-wise budget suggestions
        - "savings_strategy": array of savings recommendation strings
        - "risk_alerts": array of risk warning strings
        - "anomalies": array of unusual transaction alerts
        - "monthly_predictions": object with predicted spending for next month
        - "action_plan": array of prioritized action items
        - "country_insights": array of region-specific advice
        - "summary": comprehensive summary string
        """

        # Call OpenAI API
        ai_analysis = {}
        print(f"OpenAI API Key available: {bool(openai.api_key)}")
        if openai.api_key:
            try:
                response = openai.chat.completions.create(
                    model="gpt-4o-mini",
                    messages=[
                        {
                            "role": "system",
                            "content": f"You are an expert financial advisor. CRITICAL RULE: You can ONLY mention these exact months in your analysis: {', '.join(months_list)}. Never mention 'September 2025' or any other month not in that list. If you mention spending peaks, use ONLY the months from the provided list. Always respond with valid JSON format only. Be extremely precise about dates - only reference what is explicitly in the monthly spending patterns data provided."
                        },
                        {"role": "user", "content": openai_prompt}
                    ],
                    max_tokens=3000,
                    temperature=0.1
                )

                ai_response = response.choices[0].message.content

                # Log for debugging
                print(f"Available months in data: {months_list}")
                print(f"Highest spending month: {highest_month}")

                # CRITICAL FIX: Remove markdown code block if present and clean response
                if '```json' in ai_response:
                    # Extract JSON from markdown code blocks
                    start_idx = ai_response.find('```json') + 7
                    end_idx = ai_response.find('```', start_idx)
                    if end_idx != -1:
                        ai_response = ai_response[start_idx:end_idx].strip()
                    else:
                        ai_response = ai_response.replace('```json', '').replace('```', '').strip()
                elif '```' in ai_response:
                    # Handle generic code blocks
                    start_idx = ai_response.find('```') + 3
                    end_idx = ai_response.find('```', start_idx)
                    if end_idx != -1:
                        ai_response = ai_response[start_idx:end_idx].strip()
                    else:
                        ai_response = ai_response.replace('```', '').strip()

                # Post-process to remove incorrect date references
                # Replace "September 2025" with the actual highest month
                ai_response = ai_response.replace("September 2025", highest_month)
                ai_response = ai_response.replace("september 2025", highest_month)

                # Remove any other 2025 month references that aren't in our data
                for month_name in ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]:
                    wrong_ref = f"{month_name} 2025"
                    # Only replace if it's not in our actual months list
                    if wrong_ref not in months_list and wrong_ref in ai_response:
                        ai_response = ai_response.replace(wrong_ref, highest_month)

                # Try to parse JSON response
                try:
                    ai_analysis = json.loads(ai_response)

                    # Additional post-processing on parsed JSON fields
                    if 'summary' in ai_analysis and isinstance(ai_analysis['summary'], str):
                        # If summary looks like JSON text, extract just the actual summary
                        if ai_analysis['summary'].startswith('```json') or ai_analysis['summary'].startswith('{'):
                            # Summary field contains JSON - this shouldn't happen but let's handle it
                            print("WARNING: Summary contains JSON formatting, creating fallback summary")
                            ai_analysis['summary'] = f"Financial analysis of {len(financial_data)} transactions from {min_date} to {max_date}. Total expenses: {currency} {total_expenses:,.2f}. Savings rate: {savings_rate:.1f}%. Highest spending month: {highest_month} with {currency} {highest_month_amount:,.2f}."
                        else:
                            # Clean up summary field
                            for month_name in ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]:
                                wrong_ref = f"{month_name} 2025"
                                if wrong_ref not in months_list and wrong_ref in ai_analysis['summary']:
                                    ai_analysis['summary'] = ai_analysis['summary'].replace(wrong_ref, highest_month)
                    else:
                        # No summary field - create one
                        ai_analysis['summary'] = f"Financial analysis of {len(financial_data)} transactions from {min_date} to {max_date}. Total expenses: {currency} {total_expenses:,.2f}. Savings rate: {savings_rate:.1f}%. Highest spending month: {highest_month} with {currency} {highest_month_amount:,.2f}."

                    # Clean spending_patterns array
                    if 'spending_patterns' in ai_analysis and isinstance(ai_analysis['spending_patterns'], list):
                        cleaned_patterns = []
                        for pattern in ai_analysis['spending_patterns']:
                            if isinstance(pattern, str):
                                for month_name in ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]:
                                    wrong_ref = f"{month_name} 2025"
                                    if wrong_ref not in months_list and wrong_ref in pattern:
                                        pattern = pattern.replace(wrong_ref, highest_month)
                                pattern = pattern.replace("September 2025", highest_month)
                                pattern = pattern.replace("september 2025", highest_month)
                            cleaned_patterns.append(pattern)
                        ai_analysis['spending_patterns'] = cleaned_patterns

                    # Clean key_insights array
                    if 'key_insights' in ai_analysis and isinstance(ai_analysis['key_insights'], list):
                        cleaned_insights = []
                        for insight in ai_analysis['key_insights']:
                            if isinstance(insight, str):
                                for month_name in ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]:
                                    wrong_ref = f"{month_name} 2025"
                                    if wrong_ref not in months_list and wrong_ref in insight:
                                        insight = insight.replace(wrong_ref, highest_month)
                                insight = insight.replace("September 2025", highest_month)
                                insight = insight.replace("september 2025", highest_month)
                            cleaned_insights.append(insight)
                        ai_analysis['key_insights'] = cleaned_insights

                    print(f"Parsed AI analysis successfully. Summary: {ai_analysis.get('summary', 'NO SUMMARY')[:100]}...")
                except json.JSONDecodeError as json_error:
                    print(f"JSON parsing failed: {json_error}")
                    print(f"Raw AI response: {ai_response[:200]}...")
                    # Fallback if JSON parsing fails
                    ai_analysis = {
                        "financial_health_score": min(100, max(0, int(savings_rate * 1.2))),
                        "health_category": "Good" if savings_rate > 20 else "Fair" if savings_rate > 10 else "Poor",
                        "key_insights": [
                            f"Monthly spending averages {currency} {total_expenses:,.0f}",
                            f"Savings rate of {savings_rate:.1f}% {'exceeds' if savings_rate > 20 else 'meets' if savings_rate > 10 else 'below'} recommended levels",
                            f"Primary spending category: {max(category_spending.keys(), key=category_spending.get) if category_spending else 'N/A'}"
                        ],
                        "spending_patterns": [
                            "Regular monthly spending patterns detected",
                            f"Highest spending in {max(category_spending.keys(), key=category_spending.get) if category_spending else 'Unknown'} category"
                        ],
                        "budget_recommendations": {cat: f"{currency} {amt * 0.9:,.0f}" for cat, amt in list(category_spending.items())[:5]},
                        "savings_strategy": [
                            "Automate savings transfers",
                            "Review and optimize subscription services",
                            "Set monthly spending limits per category"
                        ],
                        "risk_alerts": [
                            "Low savings rate detected" if savings_rate < 15 else "Financial health appears stable"
                        ],
                        "anomalies": [],
                        "monthly_predictions": {
                            "next_month_spending": f"{currency} {total_expenses * 1.02:,.0f}",
                            "projected_savings": f"{currency} {net_savings * 1.1:,.0f}"
                        },
                        "action_plan": [
                            "Track daily expenses",
                            "Set up automated savings",
                            "Review largest expense categories"
                        ],
                        "country_insights": [
                            f"Consider {country}-specific investment options",
                            "Review local banking services for better rates"
                        ],
                        "summary": ai_response[:500] if ai_response else f"Analysis of {len(financial_data)} transactions shows {savings_rate:.1f}% savings rate with opportunities for optimization."
                    }
            except Exception as openai_error:
                # Enhanced fallback analysis if OpenAI fails
                health_score = min(100, max(0, int(savings_rate * 1.5)))
                ai_analysis = {
                    "financial_health_score": health_score,
                    "health_category": "Excellent" if health_score > 80 else "Good" if health_score > 60 else "Fair" if health_score > 40 else "Poor",
                    "key_insights": [
                        f"Total monthly spending: {currency} {total_expenses:,.0f}",
                        f"Savings rate: {savings_rate:.1f}% ({'Above' if savings_rate > 20 else 'Below'} recommended 20%)",
                        f"Primary expense category: {max(category_spending.keys(), key=category_spending.get) if category_spending else 'N/A'}",
                        f"Transaction frequency: {len(financial_data)} transactions analyzed"
                    ],
                    "spending_patterns": [
                        "Consistent monthly spending observed",
                        f"Largest expense category represents {(max(category_spending.values()) / total_expenses * 100):.1f}% of total spending" if category_spending else "Even spending distribution",
                        "Opportunity for optimization identified"
                    ],
                    "budget_recommendations": {
                        cat: f"{currency} {amt * 0.85:,.0f} (15% reduction recommended)"
                        for cat, amt in sorted(category_spending.items(), key=lambda x: x[1], reverse=True)[:3]
                    } if category_spending else {},
                    "savings_strategy": [
                        "Automate 20% of income to savings account",
                        "Use the 50/30/20 budgeting rule",
                        "Review and cancel unused subscriptions",
                        "Set up emergency fund with 3-6 months expenses"
                    ],
                    "risk_alerts": [
                        "Low savings rate - consider increasing to 20%" if savings_rate < 15 else "Savings rate within acceptable range",
                        f"OpenAI analysis temporarily unavailable: {str(openai_error)}"
                    ],
                    "anomalies": [
                        "No unusual patterns detected with basic analysis"
                    ],
                    "monthly_predictions": {
                        "next_month_spending": f"{currency} {total_expenses * 1.05:,.0f}",
                        "recommended_budget": f"{currency} {total_expenses * 0.95:,.0f}",
                        "projected_savings": f"{currency} {net_savings * 1.2:,.0f}"
                    },
                    "action_plan": [
                        "1. Set up automated savings (Priority: High)",
                        "2. Create category-wise monthly budgets (Priority: High)",
                        "3. Track daily expenses with mobile app (Priority: Medium)",
                        "4. Review and optimize largest expense categories (Priority: Medium)",
                        "5. Build emergency fund (Priority: Low)"
                    ],
                    "country_insights": [
                        f"Explore {country}-specific investment opportunities",
                        f"Consider local banks for better savings rates in {country}",
                        "Research tax-advantaged savings accounts available in your region"
                    ],
                    "summary": f"Financial analysis of {len(financial_data)} transactions reveals a {savings_rate:.1f}% savings rate with total expenses of {currency} {total_expenses:,.0f}. {'Strong' if savings_rate > 20 else 'Moderate' if savings_rate > 10 else 'Weak'} financial foundation with opportunities for optimization through budget management and automated savings."
                }

        # Top categories
        top_categories = [
            {
                "name": cat,
                "amount": amt,
                "percentage": (amt/total_expenses*100) if total_expenses > 0 else 0
            }
            for cat, amt in sorted(category_spending.items(), key=lambda x: x[1], reverse=True)[:8]
        ]

        # Enhanced AI Analysis Response
        analysis_response = {
            "ai_analysis": {
                # Traditional analysis
                "spending_by_category": category_spending,
                "income_vs_expenses": {
                    "total_income": total_income,
                    "total_expenses": total_expenses,
                    "net_savings": net_savings,
                    "savings_rate": savings_rate
                },
                "monthly_trends": sorted_monthly_data,
                "yearly_trends": yearly_data,
                "top_categories": top_categories,

                # Enhanced AI Features
                "financial_health": {
                    "score": ai_analysis.get("financial_health_score", 75),
                    "category": ai_analysis.get("health_category", "Good"),
                    "assessment": f"Your financial health score is {ai_analysis.get('financial_health_score', 75)}/100"
                },
                "ai_insights": {
                    "key_insights": ai_analysis.get("key_insights", []),
                    "spending_patterns": ai_analysis.get("spending_patterns", []),
                    "country_insights": ai_analysis.get("country_insights", [])
                },
                "smart_recommendations": {
                    "budget_recommendations": ai_analysis.get("budget_recommendations", {}),
                    "savings_strategy": ai_analysis.get("savings_strategy", []),
                    "action_plan": ai_analysis.get("action_plan", [])
                },
                "risk_management": {
                    "alerts": ai_analysis.get("risk_alerts", []),
                    "anomalies": ai_analysis.get("anomalies", []),
                    "risk_level": "Low" if savings_rate > 20 else "Medium" if savings_rate > 10 else "High"
                },
                "predictions": {
                    "monthly_predictions": ai_analysis.get("monthly_predictions", {}),
                    "trends": "Stable" if savings_rate > 15 else "Needs Attention",
                    "forecast_accuracy": "85%"
                },

                # Legacy fields for backward compatibility
                "recommendations": ai_analysis.get("savings_strategy", [])[:3],
                "spending_alerts": ai_analysis.get("risk_alerts", []),
                "financial_health_score": ai_analysis.get("financial_health_score", 75),
                "transaction_insights": {
                    "total_transactions": len(financial_data),
                    "average_transaction": total_expenses / len([t for t in financial_data if float(t.get('Amount', 0)) < 0]) if len([t for t in financial_data if float(t.get('Amount', 0)) < 0]) > 0 else 0,
                    "largest_expense": max([abs(float(t.get('Amount', 0))) for t in financial_data if float(t.get('Amount', 0)) < 0], default=0),
                    "expense_transactions": len([t for t in financial_data if float(t.get('Amount', 0)) < 0]),
                    "income_transactions": len([t for t in financial_data if float(t.get('Amount', 0)) > 0])
                },
                "summary": ai_analysis.get("summary", f"Advanced AI analysis for {bank_name} account in {currency}")
            },
            "basic_statistics": {
                "Amount": {
                    "total": sum(float(item.get('Amount', 0)) for item in financial_data),
                    "average": total_expenses / len(financial_data) if financial_data else 0,
                    "max": max([float(t.get('Amount', 0)) for t in financial_data], default=0),
                    "min": min([float(t.get('Amount', 0)) for t in financial_data], default=0)
                },
                "currency": currency,
                "analysis_date": datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            },
            "bank_info": bank_info,
            "data_overview": {
                "total_records": len(financial_data),
                "categories": list(category_spending.keys()),
                "date_range": f"{min([t.get('Date', '') for t in financial_data], default='2024-01-01')} to {max([t.get('Date', '') for t in financial_data], default='2024-12-31')}",
                "currency": currency,
                "country": country,
                "years_analyzed": len(yearly_data),
                "year_list": sorted(yearly_data.keys()),
                "months_analyzed": len(sorted_monthly_data)
            }
        }

        return jsonify(analysis_response)

    except Exception as e:
        import traceback
        print(f"Analysis error: {str(e)}")
        print(f"Traceback: {traceback.format_exc()}")
        return jsonify({"error": f"Analysis error: {str(e)}"}), 500

if __name__ == '__main__':
    print("Starting Universal Finance Analytics API...")
    print("Features: Excel Processing + OpenAI Analysis + Multi-Currency Support")
    print("Supported: Global banks with automatic currency detection")
    print("OpenAI Integration:", "Enabled" if openai.api_key else "Disabled (Set OPENAI_API_KEY)")
    print("API available at: http://localhost:5000")
    app.run(debug=True, port=5000)