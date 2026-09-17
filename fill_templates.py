"""
Beacon Regression Template Filler
==================================
Takes Beacon import templates and fills them with regression
test data. Output preserves exact template format.

Outputs:
    Lender_Regression_Filled.xlsx
    Sanction_Regression_Filled.xlsx
    WCDL_Regression_Filled.xlsx

Usage:
    python fill_templates.py
"""

import pandas as pd
import random
from datetime import datetime, timedelta
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

random.seed(42)

# ==========================================================
# SHARED DATA
# ==========================================================
COMPANY = {
    "name": "Wipro Ltd", "short": "WIPRO",
    "pan": "AAACW1234A", "gstin": "27AAACW1234A1Z5",
    "cin": "L12345MH1945PLC004520",
    "address": "Doddakannelli, Sarjapur Road",
    "city": "Bangalore", "state": "Karnataka",
    "pin": "560035", "country": "India",
    "group": "Wipro Group",
}

BANKS = [
    {"is_bank":"Yes","bank_id":"BNK-001","short":"HDFC","full":"HDFC Bank Ltd","loc":"Mumbai","pan":"AAACH1234E","grp":"HDFC Group","cls":"Bank","typ":"Private Bank","sub":"Scheduled Commercial Bank"},
    {"is_bank":"Yes","bank_id":"BNK-002","short":"AXIS","full":"Axis Bank Ltd","loc":"Mumbai","pan":"AAACA1234H","grp":"Axis Group","cls":"Bank","typ":"Private Bank","sub":"Scheduled Commercial Bank"},
    {"is_bank":"Yes","bank_id":"BNK-003","short":"SBI","full":"State Bank of India","loc":"Mumbai","pan":"AAACS1234G","grp":"SBI Group","cls":"Bank","typ":"PSU Bank","sub":"Scheduled Commercial Bank"},
]

CORP_LENDERS = [
    {"is_bank":"No","bank_id":"","short":"INFY","full":"Infosys Ltd","loc":"Bangalore","pan":"AAACI1234B","grp":"Infosys Group","cls":"Corporate","typ":"NBFC","sub":"NBFC-ICC"},
    {"is_bank":"No","bank_id":"","short":"TCS","full":"Tata Consultancy Services Ltd","loc":"Mumbai","pan":"AAACT1234C","grp":"Tata Group","cls":"Corporate","typ":"NBFC","sub":"NBFC-ICC"},
    {"is_bank":"No","bank_id":"","short":"RIL","full":"Reliance Industries Ltd","loc":"Mumbai","pan":"AAACR1234D","grp":"Reliance Group","cls":"Corporate","typ":"NBFC","sub":"NBFC-ICC"},
    {"is_bank":"No","bank_id":"","short":"XYZF","full":"XYZ Finance Ltd","loc":"Bangalore","pan":"AAAXYZ1234K","grp":"XYZ Group","cls":"NBFC","typ":"NBFC","sub":"NBFC-ICC"},
]

ALL_LENDERS = BANKS + CORP_LENDERS

INST_RULES = {
    "Term Loan":                     {"rev": False, "sub": True,  "curr": "INR"},
    "Working Capital Demand Loan":   {"rev": False, "sub": True,  "curr": "INR"},
    "Inter Company Deposit":         {"rev": False, "sub": False, "curr": "INR"},
    "Direct Assignment":             {"rev": False, "sub": False, "curr": "INR"},
    "External Commercial Borrowing": {"rev": False, "sub": True,  "curr": "USD"},
    "Re-Finance":                    {"rev": False, "sub": True,  "curr": "INR"},
    "Line Of Credit":                {"rev": True,  "sub": True,  "curr": "INR"},
    "Short Term Loan":               {"rev": True,  "sub": False, "curr": "INR"},
    "Cash Credit":                   {"rev": True,  "sub": True,  "curr": "INR"},
    "Overdraft":                     {"rev": True,  "sub": False, "curr": "INR"},
}
ALL_INSTRUMENTS = list(INST_RULES.keys())

DCC = ["Actual/365","Actual/360","Actual/Actual By Payment Date","US 30/360","European 30/360"]
COUPON_TYPES = ["Fixed","Floating"]
COUPON_FREQ = ["1M","3M","6M","12M"]
TDS_TYPES = ["No TDS","Normal Tax Rate","Tax Rate"]
FEE_TYPES = ["Processing Fee","Advisory / Arranger Fee","Renewal Fee","Stamp Duty","Bank Charges","Commitment Fee"]
INDICES = ["Repo Rate","MCLR","LIBOR","SOFR","T-Bill"]

def add_days(date_str, days):
    if isinstance(date_str, str):
        date_str = datetime.strptime(date_str, "%Y-%m-%d")
    return (date_str + timedelta(days=days)).strftime("%Y-%m-%d")

def pick(lst): return random.choice(lst)

# ==========================================================
# FILE 1: LENDER MASTER
# ==========================================================
def fill_lender_template():
    """Generates Lender_Regression_Filled.xlsx with 7 lenders"""
    wb = Workbook()
    wb.remove(wb.active)

    # ---- Sheet 1: Basic Client Details ----
    ws = wb.create_sheet("Basic Client Details")
    headers = ["Is Bank","Bank ID","Short Name","Full Name","Location","PAN",
               "Group","Related Party","Related Party Dropdown","Lender Class",
               "Lender Type","Lender Sub Type","TAN","Legal Entity Identifier","ID"]
    ws.append(headers)
    for i, l in enumerate(ALL_LENDERS, 1):
        ws.append([
            l["is_bank"], l["bank_id"], l["short"], l["full"], l["loc"], l["pan"],
            l["grp"], "No", "Parent - Holding", l["cls"], l["typ"], l["sub"],
            f"TAN{i:05d}", f"LEI{i:05d}", f"LEN-{i:03d}"
        ])

    # ---- Sheet 2: Demat Details ----
    ws2 = wb.create_sheet("Demat Details")
    ws2.append(["DP Name","DP ID","Demat Account","Scheme Name","Client ID","Depository","ID"])
    for i, l in enumerate(ALL_LENDERS, 1):
        ws2.append([f"{l['short']} DP", f"DP{i:04d}", f"DEMAT{i:06d}",
                    "Scheme A", f"CL{i:05d}", "NSDL", f"LEN-{i:03d}"])

    # ---- Sheet 3: Contact Details ----
    ws3 = wb.create_sheet("Contact Details")
    ws3.append(["Contact Person","Contact Number","Landline Number","Email Address","Designation","ID"])
    for i, l in enumerate(ALL_LENDERS, 1):
        ws3.append([f"Contact Person {i}", f"98{i:08d}", f"080-2{i:07d}",
                    f"contact{i}@test.com", "Manager", f"LEN-{i:03d}"])

    # ---- Sheet 4: Bank Account Details ----
    ws4 = wb.create_sheet("Bank Account Details")
    ws4.append(["PAN","Account Name","Name","Branch","Address","IFSC Code",
                "MICR Code","Account Use","Account Number","ID"])
    for i, l in enumerate(ALL_LENDERS, 1):
        ws4.append([l["pan"], l["full"], l["full"], f"{l['loc']} Main",
                    f"{l['loc']} Address", f"{l['short'].upper():4s}000{i:04d}",
                    f"56024{i:05d}", "Current", f"{10000+i}", f"LEN-{i:03d}"])

    # ---- Sheet 5: Address Segment Details ----
    ws5 = wb.create_sheet("Address Segment Details")
    ws5.append(["Office Location Type","Mark as Default","Office DUNS Number",
                "Address Line 1","Address Line 2","Address Line 3","City / Town",
                "District","State / Union Territory","Country","Pin Code","GSTIN",
                "Mobile Number","Email Address","Telephone Area Code",
                "Telephone Number(s)","Fax Area Code","Fax Number(s)","ID"])
    for i, l in enumerate(ALL_LENDERS, 1):
        ws5.append(["Registered Office","Yes",f"DUNS{i:05d}",
                    l["loc"], "Address Line 2", "", l["loc"], l["loc"],
                    "Maharashtra" if l["loc"]=="Mumbai" else "Karnataka",
                    "India", "400001" if l["loc"]=="Mumbai" else "560001",
                    l["pan"][:2]+"XXXXX"+"1Z5", f"98{i:08d}",
                    f"info{i}@test.com", "080", f"2{i:07d}", "080", f"2{i:07d}",
                    f"LEN-{i:03d}"])

    wb.save("Lender_Regression_Filled.xlsx")
    print(f"✅ Lender_Regression_Filled.xlsx — {len(ALL_LENDERS)} lenders")

# ==========================================================
# FILE 2: SANCTION MASTER
# ==========================================================
def fill_sanction_template():
    """Generates Sanction_Regression_Filled.xlsx"""
    wb = Workbook()
    wb.remove(wb.active)

    basic, main, sub = [], [], []
    counter = 1

    # 12 sanctions - mix of banks and corporates
    for i in range(12):
        lender = ALL_LENDERS[i % len(ALL_LENDERS)]
        sref = f"SM-26-{str(counter).zfill(5)}"
        counter += 1

        picked = random.sample(ALL_INSTRUMENTS, random.choice([1, 2, 2, 3]))
        basic.append({
            "Sanction_Ref": sref,
            "Lender": lender["full"],
            "Date": add_days("2026-01-05", i*2),
            "Company": COMPANY["name"],
        })

        for j, inst in enumerate(picked, 1):
            rule = INST_RULES[inst]
            nature = pick(["Revolving","Non-Revolving"]) if rule["rev"] else "Non-Revolving"
            amt = pick([1000000, 2500000, 5000000, 7500000, 10000000])
            has_sub = "Yes" if (rule["sub"] and random.random() < 0.4) else "No"
            main.append({
                "Sanction_Ref": sref, "Breakup_No": j, "Instrument": inst,
                "Nature": nature, "Amount": amt,
                "Validity": add_days("2026-01-05", 360),
                "Has_Sub": has_sub, "Currency": rule["curr"]
            })
            if has_sub == "Yes":
                sub.append({
                    "Sanction_Ref": sref, "Breakup_No": j, "Instrument": inst,
                    "Nature": nature, "Amount": amt,
                    "Validity": add_days("2026-01-05", 360),
                    "Currency": rule["curr"]
                })

    # Sheet 1: Basic Details
    ws = wb.create_sheet("Basic Details")
    ws.append(["Sanction Reference*","Lender Full Name*","Sanction Date*","Company Full Name"])
    for b in basic:
        ws.append([b["Sanction_Ref"], b["Lender"], b["Date"], b["Company"]])

    # Sheet 2: Main Breakup Details
    ws2 = wb.create_sheet("Main Breakup Details")
    ws2.append(["Sanction Reference*","Main Breakup Number(Used for mapping sub limits)",
                "Instrument*","Nature of breakup*","Sanction Amount*",
                "Availability/Validity Date*","Has Sub Limits*","Currency*"])
    for m in main:
        ws2.append([m["Sanction_Ref"], m["Breakup_No"], m["Instrument"],
                    m["Nature"], m["Amount"], m["Validity"], m["Has_Sub"], m["Currency"]])

    # Sheet 3: Sub Breakup Details
    ws3 = wb.create_sheet("Sub Breakup Details")
    ws3.append(["Sanction Reference*","Main Breakup Number(Used for mapping sub limits)",
                "Instrument*","Nature of breakup*","Sanction Amount",
                "Availability/Validity Date*","Currency*"])
    for s in sub:
        ws3.append([s["Sanction_Ref"], s["Breakup_No"], s["Instrument"],
                    s["Nature"], s["Amount"], s["Validity"], s["Currency"]])

    # Sheet 4: dropdowns
    ws4 = wb.create_sheet("dropdowns")
    ws4.append(["Nature of Sanction","Has Sub Limits","Instruments","","Currency"])
    ws4.append(["Non-Revolving","Yes","Term Loan","","USD"])
    ws4.append(["Revolving","No","Working Capital Demand Loan","","JPY"])
    ws4.append(["","","Inter Company Deposit","","EUR"])
    ws4.append(["","","Direct Assignment","","INR"])
    ws4.append(["","","External Commercial Borrowing","","GBP"])
    ws4.append(["","","Re-Finance","","CHF"])
    ws4.append(["","","Line Of Credit","","QAR"])
    ws4.append(["","","Short Term Loan","","CAD"])
    ws4.append(["","","Cash Credit","","SEK"])
    ws4.append(["","","Overdraft","","SGD"])

    wb.save("Sanction_Regression_Filled.xlsx")
    print(f"✅ Sanction_Regression_Filled.xlsx — {len(basic)} sanctions, {len(main)} breakups, {len(sub)} sub-limits")

# ==========================================================
# FILE 3: WCDL / STL DEAL
# ==========================================================
def fill_wcdl_template():
    """Generates WCDL_Regression_Filled.xlsx with all sub-sheets"""
    wb = Workbook()
    wb.remove(wb.active)

    # Pre-generate sanctions for reference
    sanctions = []
    for i in range(10):
        l = ALL_LENDERS[i % len(ALL_LENDERS)]
        sanctions.append({
            "ref": f"SM-26-{str(i+1).zfill(5)}",
            "lender": l,
            "facility_wc": f"WC-Rev-{str(i+1).zfill(4)}",
        })

    deals = []
    for i in range(6):
        s = sanctions[i]
        ctype = pick(COUPON_TYPES)
        settle = add_days("2026-02-01", i*3)
        deals.append({
            "ID": f"WC-{str(i+1).zfill(3)}",
            "Is_Rollover": pick(["Yes","No"]),
            "Rollover_Type": pick(["With Interest","Without Interest"]),
            "Settlement_Date": settle,
            "Lender": s["lender"]["full"],
            "Curr_Acc": f"{10001 + (i*2)}",
            "Loan_Acc": f"{20001 + (i*2)}",
            "Loan_Acc_No": f"LAN{i+1:08d}",
            "Nature": pick(["Secured","Unsecured"]),
            "Security_Cover": pick([0,1,1,1]),
            "Company": COMPANY["name"],
            "Sanction_ID": s["ref"],
            "Facility_ID": s["facility_wc"],
            "Is_EMI": "No",
            "Disb_Date": settle,
            "Disb_Amt": pick([20000, 50000, 100000, 150000]),
            "Repay_Date": add_days(settle, 90),
            "Repay_Amt": 20000,
            "Coupon_Type": ctype,
            "Coupon_Freq": pick(COUPON_FREQ),
            "Fixed_Rate": round(random.uniform(0.06,0.12),4) if ctype=="Fixed" else "",
            "Index": pick(INDICES) if ctype=="Floating" else "",
            "Spread": round(random.uniform(0.01,0.05),4) if ctype=="Floating" else "",
            "DCC": pick(DCC),
            "First_Int": add_days(settle, 3),
            "Next_Review": add_days(settle, 33),
            "TDS": pick(TDS_TYPES),
        })

    # ---- Facility Details ----
    ws = wb.create_sheet("Facility Details")
    ws.append(["ID*","Is Rollover","Rollover Type","Rollover ID","Settlement Date*",
               "Lender Full Name*","Current Bank Account*","Loan Bank Account",
               "Loan Account Number","Nature","Details Of Security","Security Cover",
               "Company Full Name*","Unsecured Loan Description",
               "Scheme Name (Re-finance)","Scheme Details (Re-finance)"])
    for d in deals:
        ws.append([d["ID"], d["Is_Rollover"], d["Rollover_Type"] if d["Is_Rollover"]=="Yes" else "",
                   "", d["Settlement_Date"], d["Lender"], d["Curr_Acc"], d["Loan_Acc"],
                   d["Loan_Acc_No"], d["Nature"], "", d["Security_Cover"],
                   d["Company"], "Unsecured Loan" if d["Nature"]=="Unsecured" else "",
                   "None", ""])

    # ---- Ratings ----
    ws2 = wb.create_sheet("Ratings")
    ws2.append(["ID*","Rating Reference*","Remarks"])
    for d in deals:
        ws2.append([d["ID"], pick(["AAA","AA+","AA","A+","Unrated"]), "Sample"])

    # ---- Contact Master ----
    ws3 = wb.create_sheet("Contact Master")
    ws3.append(["ID*","Contact Person Name","Comments"])
    for d in deals:
        ws3.append([d["ID"], f"Contact {d['ID']}", "Sample"])

    # ---- Sanction Details ----
    ws4 = wb.create_sheet("Sanction Details")
    ws4.append(["ID*","Sanction ID*","Sanction Facility ID*","Is EMI*",
                "EMI Override DCC","Maturity for EMI","PRINCIPAL MORATORIUM INSTALLMENT"])
    for d in deals:
        ws4.append([d["ID"], d["Sanction_ID"], d["Facility_ID"], "No", "", "", ""])

    # ---- Disbursement Schedule ----
    ws5 = wb.create_sheet("Disbursement Schedule")
    ws5.append(["ID*","Date*","Amount*","Payment Instructions","Remarks"])
    for d in deals:
        ws5.append([d["ID"], d["Disb_Date"], d["Disb_Amt"], "No", "Auto-generated"])

    # ---- Repayment Schedule ----
    ws6 = wb.create_sheet("Repayment Schedule")
    ws6.append(["ID*","Date*","Payment Date","Forward Rate (ECB)","Forward Date(ECB)",
                "Amount*","Invoice Number","Remarks"])
    for d in deals:
        ws6.append([d["ID"], d["Repay_Date"], d["Repay_Date"], "", "",
                    d["Repay_Amt"], f"INV-{d['ID']}", "Auto-generated"])

    # ---- Interest details ----
    ws7 = wb.create_sheet("Interest details")
    ws7.append(["ID*","Period (Coupon Start Date)*","Coupon Type*","Coupon Payment Frequency*",
                "Custom Payment Frequency (Value)","Custom Payment Frequency (Frequency)",
                "First Interest Due Date*","Fixed Rate","Next Rate Review Date",
                "Index (Benchmark Name)","Spread","Resets","First Reset Date",
                "Reset - N days","T-nth day Reset","Exclude -N Days for Settlement Date",
                "Next Expected Spread Reset Date","Cap","Floor","Remarks",
                "No Month-End Coupon Dates Assumption","DCC*","Exclude Payment Date for Coupon Accrual",
                "Interest on Repayment Date","Runnig Balance Loan","Capitalize Short Interest Payment",
                "Interest on Closing Balance","Coupon Adjustment in Repayment",
                "Fixed Instalments (Late Payment)","Fixed Instalments (Early Payment)",
                "TDS Calulation Type","Round Interest Due","Interest Round-Off Type",
                "Round Upto","Principal Payable","Interest Payable",
                "Principal & Interest Payable on Same Day","Holiday List",
                "Interest Payment (+n)(business days)",
                "Interest Calculation Changes if Principal Date on Holiday",
                "Coupon Accrual Date Changes if Interest on Holiday",
                "Compound Interest","Net Interest Compounding",
                "Update EMI Amount (incase of IsEMI)"])
    for d in deals:
        ws7.append([d["ID"], d["Settlement_Date"], d["Coupon_Type"], d["Coupon_Freq"],
                    "1","W", d["First_Int"], d["Fixed_Rate"], d["Next_Review"],
                    d["Index"], d["Spread"], "", "", "Yes","10","Yes","",
                    "", "", "Sample","Yes", d["DCC"], "Yes","Yes","Yes","Yes","Yes",
                    "Adjust Amount from Next Repayment","Yes","Yes","Due Date Basis",
                    "Yes","Round-Closest","1","Previous","Previous",
                    "Respective Convention","","2","Yes","Yes","Yes","Yes",""])

    # ---- TDS ----
    ws8 = wb.create_sheet("TDS")
    ws8.append(["ID*","TDS","Start Date"])
    for d in deals:
        ws8.append([d["ID"], d["TDS"], d["Settlement_Date"]])

    # ---- Custom Compounding ----
    ws9 = wb.create_sheet("Custom Compounding")
    ws9.append(["ID*","First Compounding Date","No Compounding","Frequency"])
    for d in deals:
        ws9.append([d["ID"], "", "", ""])

    # ---- Fees,Charges ----
    ws10 = wb.create_sheet("Fees,Charges")
    ws10.append(["ID*","Invoice Date","Payment Due Date","Type","Amount",
                 "% of Disbursement","Counterparty Type","Counterparty Name",
                 "Is GST Applicable","GSTIN","Include for IndAS","TDS Type",
                 "Is RCM Applicable","Invoice Num","Details","Bank Account","Remarks"])
    for d in deals:
        for j in range(random.choice([1, 2])):
            ws10.append([d["ID"], d["Disb_Date"], add_days(d["Disb_Date"], 30),
                         pick(FEE_TYPES), pick([200,500,1000,2000,5000]),
                         round(random.uniform(0.05,1.0),2),
                         pick(["Lender","Arranger","Other"]), d["Lender"],
                         pick(["Yes","No"]), "27AAACH1234E1Z5",
                         pick(["Yes","No"]), pick(TDS_TYPES),
                         pick(["Yes","No"]), f"INV-{d['ID']}-{j+1}",
                         "Auto", d["Curr_Acc"], ""])

    # ---- PutCall Option ----
    ws11 = wb.create_sheet("PutCall Option")
    ws11.append(["ID*","Date","Notification Start Date","Notification End Date","Put","Call"])
    for d in deals:
        ws11.append([d["ID"], "", "", "", "No", "No"])

    # ---- Penalty Details ----
    ws12 = wb.create_sheet("Penalty Details")
    ws12.append(["ID*","Penalty Offset","Penalty Offset Interest",
                 "Interest On Overdue Interest Applicable","Non-Compliance Default Rate",
                 "Cure Period Days","Compound Penalty","Penalty Compounding Frequency",
                 "Remarks","Prepayment Penalty Convention",
                 "Allow Prepayment prior to Reset Date","No. of Days"])
    for d in deals:
        ws12.append([d["ID"], "", "", "No", "", "", "No", "",
                     "", pick(["Flat Prepayment Penalty","Proportional to Prepayment Days"]),
                     "Yes", ""])

    # ---- Pre Payment Penalty ----
    ws13 = wb.create_sheet("Pre Payment Penalty")
    ws13.append(["ID*","Upto Months","No Prepayment","Applicable Rate(%)","Remarks"])
    for d in deals:
        nop = pick(["Yes","No"])
        ws13.append([d["ID"], pick([3,6,9,12]), nop,
                     0 if nop=="Yes" else pick([1.0,1.5,2.0,2.5,3.0]),
                     "Auto"])

    wb.save("WCDL_Regression_Filled.xlsx")
    print(f"✅ WCDL_Regression_Filled.xlsx — {len(deals)} deals across 13 sheets")

# ==========================================================
# MAIN
# ==========================================================
def main():
    print("\n" + "="*60)
    print("  BEACON REGRESSION TEMPLATE FILLER")
    print("="*60 + "\n")

    fill_lender_template()
    fill_sanction_template()
    fill_wcdl_template()

    print("\n" + "="*60)
    print("✅ DONE! 3 files generated in current folder:")
    print("   1. Lender_Regression_Filled.xlsx")
    print("   2. Sanction_Regression_Filled.xlsx")
    print("   3. WCDL_Regression_Filled.xlsx")
    print("="*60 + "\n")

if __name__ == "__main__":
    main()