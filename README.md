# DCF_VALUATION_TOOL
A Python tool for projecting FCFE and calculating intrinsic share value using DCF, based on Excel financials and customizable assumptions.

Requirements: Make sure you have pandas and numpy installed. You’ll also need openpyxl for reading .xlsx Excel files
### Features:
- Automatically extracts financials from income and cash flow statement sheets
- Projects revenue, net income, and FCFE over 5 years
- Applies reliability factor to adjust FCFE projections based on company stability
- Computes intrinsic value per share using user-defined discount rate and growth rate
- Built for use across any public company (tested on TSLA, GOOGL, AMZN)

### ⚠️ Excel Sheet Formatting Requirements

For the program to extract financial data correctly, your Excel file must:

- Have **two sheets**:
  - `Income Statement`
  - `Cash Flow Statement`

- Use **exact item names** (case-sensitive) in the first column:
  - In the Income Statement sheet:  
    - `Revenue`  
    - `Net Income`
  - In the Cash Flow Statement sheet:  
    - `CapEx`  
    - `Cash Flow From Operations`  
    - `Debt Issued`  
    - `Debt Repaid'
   
Built by Andrew.K
