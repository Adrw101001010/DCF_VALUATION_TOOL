import pandas as pd
import numpy as np
import os

# Ask for Excel file
file_name = input("Enter the Excel file name (with .xlsx): ")
df_is = pd.read_excel(file_name, sheet_name='Income Statement')
df_cf = pd.read_excel(file_name, sheet_name='Cash Flow Statement')

def extract_item(df, item_name):
    row = df.loc[df['Item'] == item_name]
    if not row.empty:
        return [round(float(v), 2) for v in row.iloc[0, 1:].values]
    else:
        print(f"Warning: '{item_name}' not found.")
        return None

financial_data = {
    'Revenue': extract_item(df_is, 'Revenue'),
    'Net Income': extract_item(df_is, 'Net Income'),
    'CapEx': extract_item(df_cf, 'CapEx'),
    'Cash Flow From Operations': extract_item(df_cf, 'Cash Flow From Operations'),
    'Debt Issued': extract_item(df_cf, 'Debt Issued'),
    'Debt Repaid': extract_item(df_cf, 'Debt Repaid')
}

rev_list = financial_data["Revenue"]
ni_list = financial_data["Net Income"]
capex_list = financial_data["CapEx"]
cfo_list = financial_data["Cash Flow From Operations"]
debt_issued_list = financial_data["Debt Issued"]
debt_repaid_list = financial_data["Debt Repaid"]
net_borrowing_list = [round(debt_issued_list[i] - debt_repaid_list[i], 2) for i in range(len(debt_issued_list))]

def rev_growth():
    global rev_list
    growth = sum([(rev_list[i+1] / rev_list[i]) - 1 for i in range(len(rev_list) - 1)])
    return round(growth / (len(rev_list) - 1), 4)

def rev_projection(growth_factor):
    for _ in range(5):
        next_rev = round(rev_list[-1] * (1 + growth_factor), 2)
        rev_list.append(next_rev)
    return rev_list

def fcfe_calc():
    return [round(cfo_list[i] - capex_list[i] + net_borrowing_list[i], 2) for i in range(len(cfo_list))]

def avg_net_income_margin():
    ni_margins = [round(ni_list[i] / rev_list[i], 4) for i in range(len(ni_list))]
    return ni_margins, round(sum(ni_margins) / len(ni_margins), 4)

def avg_FCFE_ratio():
    global FCFE_list, ni_list
    ratios = [round(min(max(FCFE_list[i] / ni_list[i], 0), 3), 4) for i in range(len(ni_list))]
    return ratios, round(sum(ratios) / len(ratios), 4)

def net_income_projection():
    for _ in range(len(rev_list) - len(ni_list)):
        ni_list.append(round(rev_list[len(ni_list)] * avg_ni_margin, 2))
    return ni_list

def FCFE_projection():
    for _ in range(len(ni_list) - len(FCFE_list)):
        FCFE_list.append(round(ni_list[len(FCFE_list)] * avg_FCFE_ratio * reliability_factor, 2))
    return FCFE_list

ni_margin_list, avg_ni_margin = avg_net_income_margin()
FCFE_list = fcfe_calc()
FCFE_ratio_list, avg_FCFE_ratio = avg_FCFE_ratio()
rev_growth_rate = rev_growth()
rev_list = rev_projection(rev_growth_rate)
ni_list = net_income_projection()

print("""
RELIABILITY FACTOR GUIDE:
- 1.0: Very high confidence
- 0.8–0.9: Moderately high confidence
- 0.6–0.7: Uncertain firms
- 0.4–0.5: Speculative
- <0.4: Highly speculative
""")

reliability_factor = float(input("Enter a reliability factor for FCFE projections: "))
FCFE_list = FCFE_projection()
discount_rate = float(input("Enter a discount rate: "))
perp_growth = float(input("Enter a perpetual growth rate: "))
shares_outstanding = float(input("Enter number of shares outstanding (same units as Excel): "))
current_share_price = float(input("Enter current share price: "))

terminal_value = round(((FCFE_list[-1]) * (1 + perp_growth)) / (discount_rate - perp_growth), 4)
DCFV = terminal_value / ((1 + discount_rate) ** 5)
for i in range(5):
    DCFV += FCFE_list[-5 + i] / ((1 + discount_rate) ** (i + 1))
share_intrinsic = round(DCFV / shares_outstanding, 2)

print(f"\nThe intrinsic value is ${share_intrinsic} per share.")
if share_intrinsic < (0.8 * current_share_price):
    print("The stock may be overvalued")
elif share_intrinsic > (1.21 * current_share_price):
    print("The stock may be undervalued")
else:
    print("The stock may be fairly valued")

import openai
import re

openai.api_key = "ADD YOUR OWN API" 

def extract_manual_updates(user_input):
    q = user_input.lower()
    changes = {}

    match = re.search(r"(growth|perpetual).*?(\d*\.?\d+)", q)
    if match:
        changes['growth_rate'] = float(match.group(2))

    match = re.search(r"discount.*?(\d*\.?\d+)", q)
    if match:
        changes['discount_rate'] = float(match.group(1))

    match = re.search(r"(reliability|confidence).*?(\d*\.?\d+)", q)
    if match:
        changes['reliability'] = float(match.group(2))

    return changes

def ask_gpt_natural_language(user_input, summary):
    completion = openai.chat.completions.create(
        model="gpt-4o",
        messages=[
            {"role": "system", "content": "You are a helpful financial analyst assisting with a DCF valuation."},
            {"role": "user", "content": f"The user previously calculated this:\n{summary}\n\nThey are now asking: {user_input}"}
        ]
    )
    return completion.choices[0].message.content.strip()

# Final summary string (used for GPT context)
summary_template = (
    f"The intrinsic value of the stock is ${round(share_intrinsic, 2)} per share. "
    f"The current share price is ${current_share_price}. "
    f"The inputs used were: discount rate = {discount_rate}, perpetual growth rate = {perp_growth}, "
    f"reliability factor = {reliability_factor}."
)

# Interactive loop
while True:
    user_input = input("\nAsk a question about the valuation (or type 'exit'): ").strip()
    if user_input.lower() == 'exit':
        print("Goodbye!")
        break

    updates = extract_manual_updates(user_input)

    if updates:
        if 'growth_rate' in updates:
            perp_growth = updates['growth_rate']
        if 'discount_rate' in updates:
            discount_rate = updates['discount_rate']
        if 'reliability' in updates:
            reliability_factor = updates['reliability']

        # Re-run projection
        FCFE_list = FCFE_list[:5]  # Reset FCFE projection
        FCFE_list = FCFE_projection()

        # Re-run valuation
        terminal_value = round(((FCFE_list[-1]) * (1 + perp_growth)) / (discount_rate - perp_growth), 4)
        DCFV = terminal_value / ((1 + discount_rate) ** 5)
        for i in range(5):
            DCFV += FCFE_list[-5 + i] / ((1 + discount_rate) ** (i + 1))

        share_intrinsic = DCFV / shares_outstanding
        valuation = "undervalued" if share_intrinsic > 1.2 * current_share_price else "overvalued" if share_intrinsic < 0.8 * current_share_price else "fairly valued"

        print(f"\nNew intrinsic value: ${round(share_intrinsic, 2)} — the stock may be {valuation}")
    else:
        print("\nGPT Answer:")
        print(ask_gpt_natural_language(user_input, summary_template))
