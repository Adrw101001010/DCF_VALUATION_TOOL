import pandas as pd
import numpy as np

#ASK FOR FILE NAME
file_name = input("Enter the Excel file name (with .xlsx): ")

# READ IS AND CFS FROM FILE
df_is = pd.read_excel(file_name, sheet_name='Income Statement')
df_cf = pd.read_excel(file_name, sheet_name='Cash Flow Statement')

# EXTRACT ITEMS FROM EXCEL SHEET
def extract_item(df, item_name):
    row = df.loc[df['Item'] == item_name]
    if not row.empty:
        return [round(float(v), 2) for v in row.iloc[0, 1:].values]
    else:
        print(f"Warning: '{item_name}' not found.")
        return None

# VALUES STORED IN DICTIONARY
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


print("\nExtracted Financial Data:")
for key, values in financial_data.items():
    if values is not None:
        print(f"{key}: {values}")
    else:
        print(f"{key}: Not found")

# CALCULATION FUNCTIONS

def rev_growth():
    global rev_list
    rev_growth_aggregate = 0
    for i in range(len(rev_list) - 1):
        rev_growth_aggregate += (rev_list[i+1] / rev_list[i]) - 1
    rev_growth_avg = rev_growth_aggregate / (len(rev_list) - 1)
    rev_growth_avg = round(rev_growth_avg, 4)
    print("The average revenue growth rate is:", rev_growth_avg)
    return rev_growth_avg

def rev_projection(growth_factor):
    global rev_list
    j = len(rev_list) - 1
    for i in range(5):
        next_rev = round(rev_list[j] * (1 + growth_factor), 2)
        print(f"A revenue of {next_rev} has been projected")
        rev_list.append(next_rev)
        j = len(rev_list) - 1
    return rev_list

def fcfe_calc():
    global capex_list, cfo_list, net_borrowing_list
    FCFE_list = []
    for i in range(len(cfo_list)):
        FCFE = round(cfo_list[i] - capex_list[i] + net_borrowing_list[i], 2)
        FCFE_list.append(FCFE)
    return FCFE_list

def avg_net_income_margin():
    global ni_list, rev_list
    ni_margin_list = []
    ni_margin_aggregate = 0
    for i in range(len(ni_list)):
        margin = round(ni_list[i] / rev_list[i], 4)
        ni_margin_list.append(margin)
        ni_margin_aggregate += margin
    avg_margin = round(ni_margin_aggregate / len(ni_list), 4)
    return ni_margin_list, avg_margin
def avg_FCFE_ratio():
    global FCFE_list, ni_list
    FCFE_ratio_list = []
    FCFE_ratio_aggregate = 0
    for i in range(len(ni_list)):
        raw_ratio = FCFE_list[i] / ni_list[i]
        capped_ratio = round(min(raw_ratio, 1.25), 4)  # Cap at 1.75
        FCFE_ratio_list.append(capped_ratio)
        FCFE_ratio_aggregate += capped_ratio
    avg_ratio = round(FCFE_ratio_aggregate / len(FCFE_list), 4)
    return FCFE_ratio_list, avg_ratio
def net_income_projection():
    global ni_list, rev_list, avg_ni_margin
    nLen = len(ni_list)
    for i in range(len(rev_list) - nLen):
        ni_proj = round(rev_list[nLen] * avg_ni_margin, 2)
        ni_list.append(ni_proj)
        print("A net income of:", ni_proj, "has been projected")
        nLen += 1
    return ni_list

def FCFE_projection():
    global ni_list, avg_FCFE_ratio, FCFE_list, reliability_factor
    nLen = len(FCFE_list)
    for i in range(len(ni_list) - nLen):
        fcfe_proj = round(ni_list[nLen] * avg_FCFE_ratio * reliability_factor, 2)
        FCFE_list.append(fcfe_proj)
        print("A FCFE of:", fcfe_proj, "has been projected")
        nLen += 1
    return FCFE_list

# EXECUTE CALCULATIONS

print("\nNet income margins for the given years have been calculated:")
ni_margin_list, avg_ni_margin = avg_net_income_margin()
print(f"List of the net income margins: {ni_margin_list}")
print(f"Average net income margin: {avg_ni_margin}")

print("\nGiven CFO, CapEx, debt issued, and debt repaid, FCFE for the given years has been calculated:")
FCFE_list = fcfe_calc()

print("FCFE-to-net-income-ratios have been calculated:")
FCFE_ratio_list, avg_FCFE_ratio = avg_FCFE_ratio()
print(f"List of the FCFE-to-net-income-ratios: {FCFE_ratio_list}")
print(f"Average FCFE-to-net-income-ratio: {avg_FCFE_ratio}")

print("\nGiven the past revenues, revenue growth rate has been calculated:")
rev_growth_rate = rev_growth()

print("\nUsing the growth rate, five years of revenue projections have been made:")
rev_list = rev_projection(rev_growth_rate)
print("Here are the past revenues + projections:", rev_list)

print("\nGiven the projected revenues and the past net income margins, net income projections have been made:")
ni_list = net_income_projection()
print("Here are the past net incomes + projections:", ni_list)

#SETTING RELIABILITY FACTOR FOR FCFE PROJECTIONS
print("""
RELIABILITY FACTOR GUIDE:
- 1.0: Very high confidence (e.g. large, mature, stable companies like JNJ, MSFT, or major banks)
- 0.8–0.9: Moderately high confidence (e.g. consistent revenue-generating firms in non-cyclical sectors)
- 0.6–0.7: Uncertain firms (e.g. recent margin pressure, volatile historical performance)
- 0.4–0.5: Speculative or fast-growth companies with limited track record
- <0.4: Highly speculative, minimal reliability (e.g. early-stage or loss-making firms)
""")
reliability_factor = float(input("Enter a reliability factor for FCFE projections : "))
print("\nGiven the projected net incomes, FCFE projections have been made:")
FCFE_list = FCFE_projection()
print("Here are the past FCFE + projections:", FCFE_list)

#SET USER INPUT RATES
discount_rate = float(input("enter a discount rate: "))
perp_growth = float(input("enter a perpetual growth rate: "))
shares_outstanding = float(input("enter the number of shares outstanding (in same units as financial statement values. eg. in millions):  "))
current_share_price = float(input("enter the current share price:"))

#DCF starts here
terminal_value = round(((FCFE_list[len(FCFE_list)-1])* (1+perp_growth))/(discount_rate - perp_growth),4)
print(f"ther terminal value at projection year 5: {terminal_value}")

DCFV = terminal_value/((1+discount_rate)**5)
for i in range(5):
    DCFV += FCFE_list[len(FCFE_list) - 5 + i] * ((1+discount_rate)**(i+1))
print(f"The sum of the discounted projected FCFE and terminal value is: {DCFV}" )
share_intrinsic = DCFV/shares_outstanding
print(f"The intrinsic value is ${share_intrinsic} per share.")
if share_intrinsic <  (0.8*current_share_price):
    print("the stock may be overvalued")
elif share_intrinsic > (1.21* current_share_price):
    print("the stock may be undervalued")
else:
    print("the stock may be fairly valued")

