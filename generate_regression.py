"""
Beacon Regression DB Generator - FULL VERSION
==============================================
Covers: Company, Lenders, Accounts, Sanctions,
        TL, STL, WCDL, CC, OD, LOC, ECB, ICD,
        Re-Finance, Direct Assignment, NCD, CP,
        TDS, Fees, Prepayment

Usage:
    python generate_regression.py
    python generate_regression.py input_config.xlsx
"""

import pandas as pd
import random
import sys
import os
from datetime import datetime, timedelta
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# ============================================================
# CONSTANTS
# ============================================================

COMPANY = {
    "Company_Code": "CMP-001", "Name": "Wipro", "Full_Name": "Wipro Ltd",
    "PAN": "AAACW1234A", "CIN": "L12345MH1945PLC004520",
    "GSTIN": "27AAACW1234A1Z5", "Address": "Doddakannelli, Sarjapur Road",
    "PIN": "560035", "Base_Currency": "INR", "Group_Affiliation": "Wipro Group",
    "MD_Name": "Thierry Delaporte", "CEO_Name": "Thierry Delaporte",
    "Reg_No": "004520", "Email": "info@wipro.com", "Phone": "080-28440011",
    "State": "Karnataka", "City": "Bangalore", "Country": "India",
    "CGST": "9%", "SGST": "9%", "IGST": "18%", "TDS_Rate": "10%", "TDS_Section": "194A"
}

LENDERS = [
    {"Lender_ID":"LEN-001","Is_Bank":"Yes","Bank_ID":"BNK-001","Short_Name":"HDFC","Full_Name":"HDFC Bank Ltd","Location":"Mumbai","PAN":"AAACH1234E","Group":"HDFC Group","Lender_Class":"Bank","Lender_Type":"Private Bank","Lender_Sub_Type":"Scheduled Commercial Bank","TAN":"MUMA12345B","LEI":"LEI001"},
    {"Lender_ID":"LEN-002","Is_Bank":"Yes","Bank_ID":"BNK-002","Short_Name":"AXIS","Full_Name":"Axis Bank Ltd","Location":"Mumbai","PAN":"AAACA1234H","Group":"Axis Group","Lender_Class":"Bank","Lender_Type":"Private Bank","Lender_Sub_Type":"Scheduled Commercial Bank","TAN":"MUMA12348E","LEI":"LEI004"},
    {"Lender_ID":"LEN-003","Is_Bank":"Yes","Bank_ID":"BNK-003","Short_Name":"SBI","Full_Name":"State Bank of India","Location":"Mumbai","PAN":"AAACS1234G","Group":"SBI Group","Lender_Class":"Bank","Lender_Type":"PSU Bank","Lender_Sub_Type":"Scheduled Commercial Bank","TAN":"MUMA12347D","LEI":"LEI003"},
    {"Lender_ID":"LEN-004","Is_Bank":"No","Bank_ID":"","Short_Name":"WIPRO","Full_Name":"Wipro Ltd","Location":"Bangalore","PAN":"AAACW1234A","Group":"Wipro Group","Lender_Class":"Corporate","Lender_Type":"NBFC","Lender_Sub_Type":"NBFC-ICC","TAN":"BLRA12345A","LEI":"LEI004X"},
    {"Lender_ID":"LEN-005","Is_Bank":"No","Bank_ID":"","Short_Name":"INFY","Full_Name":"Infosys Ltd","Location":"Bangalore","PAN":"AAACI1234B","Group":"Infosys Group","Lender_Class":"Corporate","Lender_Type":"NBFC","Lender_Sub_Type":"NBFC-ICC","TAN":"BLRA12346I","LEI":"LEI011"},
    {"Lender_ID":"LEN-006","Is_Bank":"No","Bank_ID":"","Short_Name":"TCS","Full_Name":"Tata Consultancy Services Ltd","Location":"Mumbai","PAN":"AAACT1234C","Group":"Tata Group","Lender_Class":"Corporate","Lender_Type":"NBFC","Lender_Sub_Type":"NBFC-ICC","TAN":"MUMA12347J","LEI":"LEI012"},
    {"Lender_ID":"LEN-007","Is_Bank":"No","Bank_ID":"","Short_Name":"RIL","Full_Name":"Reliance Industries Ltd","Location":"Mumbai","PAN":"AAACR1234D","Group":"Reliance Group","Lender_Class":"Corporate","Lender_Type":"NBFC","Lender_Sub_Type":"NBFC-ICC","TAN":"MUMA12348K","LEI":"LEI013"},
    {"Lender_ID":"LEN-008","Is_Bank":"No","Bank_ID":"","Short_Name":"XYZF","Full_Name":"XYZ Finance Ltd","Location":"Bangalore","PAN":"AAAXYZ1234K","Group":"XYZ Group","Lender_Class":"NBFC","Lender_Type":"NBFC","Lender_Sub_Type":"NBFC-ICC","TAN":"BLRA12345H","LEI":"LEI007"},
]

INSTRUMENTS = {
    "Term Loan":                     {"revolving": False, "sub_limits": True,  "currency": "INR"},
    "Working Capital Demand Loan":   {"revolving": False, "sub_limits": True,  "currency": "INR"},
    "Inter Company Deposit":         {"revolving": False, "sub_limits": False, "currency": "INR"},
    "Direct Assignment":             {"revolving": False, "sub_limits": False, "currency": "INR"},
    "Pass Through Certificate":      {"revolving": False, "sub_limits": False, "currency": "INR"},
    "External Commercial Borrowing": {"revolving": False, "sub_limits": True,  "currency": "USD"},
    "Re-Finance":                    {"revolving": False, "sub_limits": True,  "currency": "INR"},
    "Line Of Credit":                {"revolving": True,  "sub_limits": True,  "currency": "INR"},
    "Short Term Loan":               {"revolving": True,  "sub_limits": False, "currency": "INR"},
    "Cash Credit":                   {"revolving": True,  "sub_limits": True,  "currency": "INR"},
    "Overdraft":                     {"revolving": True,  "sub_limits": False, "currency": "INR"},
}

DCC_OPTIONS  = ["Actual/365","Actual/360","Actual/Actual By Payment Date","US 30/360","European 30/360"]
COUPON_TYPES = ["Fixed","Floating"]
COUPON_FREQ  = ["1M","3M","6M","12M"]
TDS_TYPES    = ["No TDS","Normal Tax Rate","Tax Rate"]
FEE_TYPES    = ["Processing Fee","Advisory / Arranger / Broker Fee","Renewal Fee","Stamp Duty","Bank Charges","Commitment Fee"]
INDICES      = ["Repo Rate","MCLR","LIBOR","SOFR","T-Bill"]
NATURES      = ["Secured","Unsecured"]

# ============================================================
# HELPERS
# ============================================================

def _id(p, n, w=5): return f"{p}-{str(n).zfill(w)}"
def _add(d, days):
    if isinstance(d, str): d = datetime.strptime(d, "%Y-%m-%d")
    return (d + timedelta(days=days)).strftime("%Y-%m-%d")
def _pick(l): return random.choice(l)
def _nature(inst):
    s = INSTRUMENTS.get(inst, {})
    return random.choice(["Revolving","Non-Revolving"]) if s.get("revolving") else "Non-Revolving"
def _curr(inst): return INSTRUMENTS.get(inst, {}).get("currency", "INR")
def _sub_ok(inst): return INSTRUMENTS.get(inst, {}).get("sub_limits", False)

# ============================================================
# GENERATORS
# ============================================================

def gen_accounts():
    out, idx = [], 1
    for l in LENDERS:
        for use in ["Current","Loan"]:
            out.append({
                "Account_ID": _id("ACC", idx, 3), "Company_Code": COMPANY["Company_Code"],
                "Lender_ID": l["Lender_ID"], "PAN": COMPANY["PAN"],
                "Account_Name": COMPANY["Full_Name"], "Branch": f"{l['Location']} Main",
                "IFSC": f"{l['Short_Name'][:4].upper():4s}000{idx:04d}",
                "MICR": f"{560240000+idx}", "Account_Use": use,
                "Account_Number": str(10000+idx if use=="Current" else 20000+idx),
                "Account_Type": use
            })
            idx += 1
    return out

def gen_sanctions(n=12, fy=26, start="2026-01-05"):
    basic, main, sub = [], [], []
    counter = 1
    insts = list(INSTRUMENTS.keys())
    for i in range(n):
        l = LENDERS[i % len(LENDERS)]
        sref = f"SM-{fy}-{str(counter).zfill(5)}"
        counter += 1
        picked = random.sample(insts, random.choice([1,2,2,3]))
        basic.append({
            "Sanction_Ref": sref, "Lender_Full_Name": l["Full_Name"],
            "Sanction_Date": _add(start, i*2), "Company_Full_Name": COMPANY["Full_Name"]
        })
        for j, inst in enumerate(picked):
            nat, cur = _nature(inst), _curr(inst)
            amt = random.choice([1000000,2500000,5000000,7500000,10000000])
            has_sub = "Yes" if (_sub_ok(inst) and random.random()<0.4) else "No"
            main.append({
                "Sanction_Ref": sref, "Main_Breakup_No": j+1, "Instrument": inst,
                "Nature_of_Breakup": nat, "Sanction_Amount": amt,
                "Validity_Date": _add(start, 360), "Has_Sub_Limits": has_sub, "Currency": cur
            })
            if has_sub == "Yes":
                sub.append({
                    "Sanction_Ref": sref, "Main_Breakup_No": j+1, "Instrument": inst,
                    "Nature_of_Breakup": nat, "Sanction_Amount": amt,
                    "Validity_Date": _add(start, 360), "Currency": cur
                })
    return basic, main, sub

def _lender_for_sanction(sref, basic):
    s = next((b for b in basic if b["Sanction_Ref"]==sref), None)
    return next((l for l in LENDERS if l["Full_Name"]==s["Lender_Full_Name"]), LENDERS[0]) if s else LENDERS[0]

def _accts(lid, accounts):
    cur = next((a for a in accounts if a["Account_Use"]=="Current" and a["Lender_ID"]==lid), accounts[0])
    lo  = next((a for a in accounts if a["Account_Use"]=="Loan"    and a["Lender_ID"]==lid), accounts[-1])
    return cur, lo

def gen_deal(deal_id, prefix, inst_filter, accounts, basic, main, i):
    pool = [m for m in main if m["Instrument"] in inst_filter]
    if not pool: return None, None
    m = pool[i % len(pool)]
    l = _lender_for_sanction(m["Sanction_Ref"], basic)
    cur, lo = _accts(l["Lender_ID"], accounts)
    nat = _pick(NATURES)
    settle = _add("2026-02-01", i*3)
    repay  = _add(settle, 90)
    base = {
        "ID": _id(prefix, i+1, 3), "Is_Rollover": _pick(["Yes","No"]),
        "Rollover_Type": _pick(["With Interest","Without Interest"]),
        "Settlement_Date": settle, "Lender_Full_Name": l["Full_Name"],
        "Current_Bank_Account": cur["Account_Number"], "Loan_Bank_Account": lo["Account_Number"],
        "Nature": nat, "Security_Cover": 1 if nat=="Secured" else 0,
        "Company_Full_Name": COMPANY["Full_Name"],
        "Unsecured_Loan_Description": "Unsecured Loan" if nat=="Unsecured" else "",
        "Sanction_ID": m["Sanction_Ref"], "Sanction_Facility_ID": f"{prefix}-Rev-{str(i+1).zfill(4)}",
        "Is_EMI": _pick(["Yes","No"]),
        "Disbursement_Date": settle, "Disbursement_Amount": m["Sanction_Amount"]//10,
        "Repayment_Date": repay, "Repayment_Amount": m["Sanction_Amount"]//10,
    }
    # Interest details for STL-type deals
    ctype = _pick(COUPON_TYPES)
    extra = {
        "Coupon_Type": ctype, "Coupon_Frequency": _pick(COUPON_FREQ),
        "Fixed_Rate": round(random.uniform(0.06,0.12),4) if ctype=="Fixed" else "",
        "Index": _pick(INDICES) if ctype=="Floating" else "",
        "Spread": round(random.uniform(0.01,0.05),4) if ctype=="Floating" else "",
        "DCC": _pick(DCC_OPTIONS),
        "First_Interest_Due_Date": _add(settle, 3),
        "Next_Rate_Review_Date": _add(settle, 33),
    }
    return base, extra

def gen_interest_details(deals):
    out = []
    for d in deals:
        out.append({
            "ID": d["ID"], "Period (Coupon Start Date)": d["Settlement_Date"],
            "Coupon Type": d.get("Coupon_Type","Fixed"),
            "Coupon Payment Frequency": d.get("Coupon_Frequency","1M"),
            "First Interest Due Date": d.get("First_Interest_Due_Date",""),
            "Fixed Rate": d.get("Fixed_Rate",""), "Index (Benchmark Name)": d.get("Index",""),
            "Spread": d.get("Spread",""), "DCC": d.get("DCC","Actual/365"),
            "Next Rate Review Date": d.get("Next_Rate_Review_Date",""),
        })
    return out

def gen_disbursement(deals):
    return [{"ID": d["ID"], "Date": d["Disbursement_Date"], "Amount": d["Disbursement_Amount"],
             "Payment Instructions": "No", "Remarks": "Auto"} for d in deals]

def gen_repayment(deals):
    return [{"ID": d["ID"], "Date": d["Repayment_Date"], "Payment Date": d["Repayment_Date"],
             "Amount": d["Repayment_Amount"], "Invoice Number": f"INV-{d['ID']}", "Remarks": "Auto"}
            for d in deals]

def gen_tds(deals):
    return [{"ID": d["ID"], "TDS": _pick(TDS_TYPES), "Start Date": d["Disbursement_Date"]} for d in deals]

def gen_fees(deals):
    out = []
    for d in deals:
        for i, ft in enumerate(random.sample(FEE_TYPES, random.choice([1,2]))):
            out.append({
                "ID": d["ID"], "Invoice Date": d["Disbursement_Date"],
                "Payment Due Date": _add(d["Disbursement_Date"], 30), "Type": ft,
                "Amount": _pick([200,500,1000,2000,5000]),
                "% of Disbursement": round(random.uniform(0.05,1.0),2),
                "Counterparty Type": _pick(["Lender","Arranger","Other"]),
                "Counterparty Name": d["Lender_Full_Name"],
                "Is GST Applicable": _pick(["Yes","No"]), "GSTIN": "27AAACH1234E1Z5",
                "Include for IndAS": _pick(["Yes","No"]), "TDS Type": _pick(TDS_TYPES),
                "Is RCM Applicable": _pick(["Yes","No"]), "Invoice Num": f"INV-{d['ID']}-{i+1}",
                "Bank Account": d["Current_Bank_Account"],
            })
    return out

def gen_prepay(deals):
    out = []
    for d in deals:
        nop = _pick(["Yes","No"])
        out.append({
            "ID": d["ID"], "Upto Months": _pick([3,6,9,12]), "No Prepayment": nop,
            "Applicable Rate(%)": 0 if nop=="Yes" else _pick([1.0,1.5,2.0,2.5,3.0]),
            "Remarks": _pick(["Flat Prepayment Penalty","Proportional to Prepayment Days","No Prepayment"]),
        })
    return out

# ============================================================
# MAIN
# ============================================================

def main():
    random.seed(42)
    accounts = gen_accounts()
    basic, main, sub = gen_sanctions(n=12)

    # Generate per instrument type deals
    deal_types = [
        ("TL",  "Term Loan",                     "06_TL_Deal",               4),
        ("WC",  "Working Capital Demand Loan",   "07_WCDL_Deal",             3),
        ("CC",  "Cash Credit",                   "08_Cash_Credit_Deal",      3),
        ("STL", "Short Term Loan",               "09_Short_Term_Loan_Deal",  3),
        ("OD",  "Overdraft",                     "10_Overdraft_Deal",        2),
        ("LOC", "Line Of Credit",                "11_Line_Of_Credit_Deal",   2),
        ("ICD", "Inter Company Deposit",         "12_ICD_Deal",              2),
        ("ECB", "External Commercial Borrowing", "13_ECB_Deal",              2),
        ("RFN", "Re-Finance",                    "14_Re_Finance_Deal",       2),
        ("DA",  "Direct Assignment",             "15_Direct_Assignment_Deal",1),
    ]

    sheets = {
        "01_Company_Master": pd.DataFrame([COMPANY]),
        "02_Regulatory_Settings": pd.DataFrame([{
            "Company_Code": COMPANY["Company_Code"], "CGST": COMPANY["CGST"],
            "SGST": COMPANY["SGST"], "IGST": COMPANY["IGST"],
            "TDS_Rate": COMPANY["TDS_Rate"], "TDS_Section": COMPANY["TDS_Section"],
            "State": COMPANY["State"], "City": COMPANY["City"], "Country": COMPANY["Country"]
        }]),
        "03_Lender_Entity_Master": pd.DataFrame(LENDERS),
        "04_Bank_Account_Master": pd.DataFrame(accounts),
        "05A_Sanction_Basic": pd.DataFrame(basic),
        "05B_Sanction_Main_Breakup": pd.DataFrame(main),
        "05C_Sanction_Sub_Breakup": pd.DataFrame(sub),
    }

    all_deals = []
    for prefix, inst, sheet_name, n in deal_types:
        deals = []
        for i in range(n):
            base, extra = gen_deal(prefix, prefix, [inst], accounts, basic, main, i)
            if base:
                base.update(extra or {})
                deals.append(base)
        if deals:
            sheets[sheet_name] = pd.DataFrame(deals)
            all_deals.extend(deals)

    # Consolidated sheets
    sheets["16_Disbursement_Schedule"] = pd.DataFrame(gen_disbursement(all_deals))
    sheets["17_Repayment_Schedule"]    = pd.DataFrame(gen_repayment(all_deals))
    sheets["18_TDS_Master"]            = pd.DataFrame(gen_tds(all_deals))
    sheets["19_Fees_Charges"]          = pd.DataFrame(gen_fees(all_deals))
    sheets["20_Prepayment_Penalty"]    = pd.DataFrame(gen_prepay(all_deals))

    out = "Regression_DB_Output.xlsx"
    with pd.ExcelWriter(out, engine="openpyxl") as w:
        for name, df in sheets.items():
            df.to_excel(w, sheet_name=name[:31], index=False)

    # Styling
    wb = load_workbook(out)
    hf = Font(bold=True, color="FFFFFF", size=11)
    hfill = PatternFill(start_color="2F5496", end_color="2F5496", fill_type="solid")
    bd = Border(left=Side(style="thin",color="CCCCCC"), right=Side(style="thin",color="CCCCCC"),
                top=Side(style="thin",color="CCCCCC"), bottom=Side(style="thin",color="CCCCCC"))
    for ws in wb.worksheets:
        for c in ws[1]:
            c.font, c.fill = hf, hfill
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.border = bd
        for col in ws.columns:
            ml = max((len(str(c.value)) for c in col if c.value), default=10)
            ws.column_dimensions[col[0].column_letter].width = min(ml+3, 30)
        ws.freeze_panes = "A2"
    wb.save(out)

    print("\n" + "="*60)
    print(f"✅  Output File: {out}")
    print("="*60)
    for name, df in sheets.items():
        print(f"   {name:35s} → {len(df):3d} rows")
    print("="*60)
    print("\n📋 IMPORT ORDER:")
    print("   1. Company Master  →  2. Regulatory  →  3. Lender  →  4. Account")
    print("   5. Sanction Basic  →  Main  →  Sub")
    print("   6. TL  →  7. WCDL  →  8. CC  →  9. STL  →  10. OD")
    print("   11. LOC  →  12. ICD  →  13. ECB  →  14. Re-Fin  →  15. DA")
    print("   16-20. Disbursement, Repayment, TDS, Fees, Prepay")
    print("="*60 + "\n")

if __name__ == "__main__":
    main()