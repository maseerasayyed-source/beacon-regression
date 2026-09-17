"""
Beacon Regression Template Filler - V2
=======================================
Reads input_config.xlsx and generates filled templates.

Usage:
    python fill_templates_v2.py                    # uses input_config.xlsx
    python fill_templates_v2.py my_config.xlsx     # uses custom config
"""

import pandas as pd
import random
import sys
import os
from datetime import datetime, timedelta
from openpyxl import Workbook

# ==========================================================
# DEFAULT VALUES (agar input file nahi mile)
# ==========================================================
DEFAULTS = {
    "company_full_name": "Wipro Ltd",
    "n_sanctions": 12,
    "n_wcdl": 6,
    "n_tl": 5,
    "fiscal_year": 26,
    "start_date": "2026-01-05",
    "random_seed": 42,
}

DEFAULT_COMPANY = {
    "name": "Wipro Ltd", "short": "WIPRO",
    "pan": "AAACW1234A", "gstin": "27AAACW1234A1Z5",
    "cin": "L12345MH1945PLC004520",
    "address": "Doddakannelli, Sarjapur Road",
    "city": "Bangalore", "state": "Karnataka",
    "pin": "560035", "group": "Wipro Group",
}

DEFAULT_LENDERS = [
    {"use":"Yes","is_bank":"Yes","bank_id":"BNK-001","short":"HDFC","full":"HDFC Bank Ltd","loc":"Mumbai","pan":"AAACH1234E","grp":"HDFC Group","cls":"Bank","typ":"Private Bank","sub":"Scheduled Commercial Bank"},
    {"use":"Yes","is_bank":"Yes","bank_id":"BNK-002","short":"AXIS","full":"Axis Bank Ltd","loc":"Mumbai","pan":"AAACA1234H","grp":"Axis Group","cls":"Bank","typ":"Private Bank","sub":"Scheduled Commercial Bank"},
    {"use":"Yes","is_bank":"Yes","bank_id":"BNK-003","short":"SBI","full":"State Bank of India","loc":"Mumbai","pan":"AAACS1234G","grp":"SBI Group","cls":"Bank","typ":"PSU Bank","sub":"Scheduled Commercial Bank"},
    {"use":"Yes","is_bank":"No","bank_id":"","short":"INFY","full":"Infosys Ltd","loc":"Bangalore","pan":"AAACI1234B","grp":"Infosys Group","cls":"Corporate","typ":"NBFC","sub":"NBFC-ICC"},
    {"use":"Yes","is_bank":"No","bank_id":"","short":"TCS","full":"Tata Consultancy Services Ltd","loc":"Mumbai","pan":"AAACT1234C","grp":"Tata Group","cls":"Corporate","typ":"NBFC","sub":"NBFC-ICC"},
    {"use":"Yes","is_bank":"No","bank_id":"","short":"RIL","full":"Reliance Industries Ltd","loc":"Mumbai","pan":"AAACR1234D","grp":"Reliance Group","cls":"Corporate","typ":"NBFC","sub":"NBFC-ICC"},
    {"use":"No","is_bank":"Yes","bank_id":"BNK-004","short":"KOTAK","full":"Kotak Mahindra Bank","loc":"Mumbai","pan":"AAACK1234I","grp":"Kotak Group","cls":"Bank","typ":"Private Bank","sub":"Scheduled Commercial Bank"},
]

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

DCC = ["Actual/365","Actual/360","Actual/Actual By Payment Date","US 30/360","European 30/360"]
COUPON_TYPES = ["Fixed","Floating"]
COUPON_FREQ = ["1M","3M","6M","12M"]
TDS_TYPES = ["No TDS","Normal Tax Rate","Tax Rate"]
FEE_TYPES = ["Processing Fee","Advisory / Arranger Fee","Renewal Fee","Stamp Duty","Bank Charges","Commitment Fee"]
INDICES = ["Repo Rate","MCLR","LIBOR","SOFR","T-Bill"]

def add_days(d, n):
    if isinstance(d, str): d = datetime.strptime(d, "%Y-%m-%d")
    return (d + timedelta(days=n)).strftime("%Y-%m-%d")

def pick(l): return random.choice(l)

# ==========================================================
# READ INPUT CONFIG
# ==========================================================
def read_input(path="input_config.xlsx"):
    cfg = dict(DEFAULTS)
    company = dict(DEFAULT_COMPANY)
    lenders = None
    instruments = None

    if not os.path.exists(path):
        print(f"⚠️  '{path}' nahi mila — DEFAULT data use kar raha hoon")
        return cfg, company, lenders, instruments

    print(f"📖 Reading: {path}")
    xls = pd.ExcelFile(path)

    # Settings
    if "Settings" in xls.sheet_names:
        df = pd.read_excel(path, sheet_name="Settings")
        for _, row in df.iterrows():
            key = str(row["Setting"]).strip()
            val = row["Value"]
            if key in cfg:
                if isinstance(cfg[key], int):
                    try: cfg[key] = int(val)
                    except: pass
                else:
                    cfg[key] = val
            elif key in company:
                company[key] = val
        print(f"   ✅ Settings loaded")

    # Lenders
    if "Lenders" in xls.sheet_names:
        df = pd.read_excel(path, sheet_name="Lenders")
        df = df[df["Use"].astype(str).str.lower() == "yes"]
        if len(df) > 0:
            # Match with default lenders to get extra fields
            result = []
            for _, row in df.iterrows():
                full = row["Full_Name"]
                match = next((l for l in DEFAULT_LENDERS if l["full"].lower() == full.lower()), None)
                if match:
                    result.append({**match, "use": "Yes"})
                else:
                    result.append({
                        "use":"Yes","is_bank":"Yes" if "bank" in str(row.get("Lender_Class","")).lower() else "No",
                        "bank_id": f"BNK-{len(result)+1:03d}" if "bank" in str(row.get("Lender_Class","")).lower() else "",
                        "short": full[:6].upper(),
                        "full": full, "loc": "Mumbai", "pan": f"AAXX{len(result):04d}A",
                        "grp": f"{full[:6]} Group",
                        "cls": row.get("Lender_Class","Bank"),
                        "typ": row.get("Lender_Type","Private Bank"),
                        "sub": "Scheduled Commercial Bank"
                    })
            lenders = result
            print(f"   ✅ Lenders loaded: {len(lenders)}")

    # Instruments
    if "Instruments" in xls.sheet_names:
        df = pd.read_excel(path, sheet_name="Instruments")
        df = df[df["Use"].astype(str).str.lower() == "yes"]
        if len(df) > 0:
            instruments = df["Instrument"].tolist()
            print(f"   ✅ Instruments loaded: {len(instruments)}")

    return cfg, company, lenders, instruments

# ==========================================================
# SAVE INPUT TEMPLATE
# ==========================================================
def create_input_template():
    """Creates input_config_template.xlsx for user to fill"""
    wb = Workbook()
    wb.remove(wb.active)

    # Settings
    ws = wb.create_sheet("Settings")
    ws.append(["Setting", "Value", "Description"])
    for k, v in DEFAULTS.items():
        desc = {
            "company_full_name": "Company name for deals",
            "n_sanctions": "How many sanctions to generate",
            "n_wcdl": "How many WCDL deals",
            "n_tl": "How many Term Loan deals",
            "fiscal_year": "FY code (26 = 2026)",
            "start_date": "Start date YYYY-MM-DD",
            "random_seed": "Seed for random (same seed = same output)",
        }.get(k, "")
        ws.append([k, v, desc])

    # Lenders
    ws2 = wb.create_sheet("Lenders")
    ws2.append(["Use", "Full_Name", "Lender_Class", "Lender_Type"])
    for l in DEFAULT_LENDERS:
        ws2.append([l["use"], l["full"], l["cls"], l["typ"]])

    # Instruments
    ws3 = wb.create_sheet("Instruments")
    ws3.append(["Use", "Instrument"])
    for inst in INST_RULES:
        ws3.append(["Yes", inst])

    # Scenarios
    ws4 = wb.create_sheet("Scenarios")
    ws4.append(["Deal_Type", "Count", "Nature", "Coupon_Type", "Rollover"])
    ws4.append(["WCDL", 3, "Secured", "Fixed", "No"])
    ws4.append(["WCDL", 2, "Unsecured", "Floating", "No"])
    ws4.append(["WCDL", 1, "Secured", "Fixed", "Yes"])
    ws4.append(["TL", 4, "Secured", "-", "No"])
    ws4.append(["TL", 1, "Unsecured", "-", "No"])

    wb.save("input_config_template.xlsx")
    print("✅ Created: input_config_template.xlsx")
    print("   → Yeh file bharo, phir 'input_config.xlsx' naam se save karo")

# ==========================================================
# GENERATE: LENDER FILE
# ==========================================================
def fill_lender_template(lenders, company):
    wb = Workbook()
    wb.remove(wb.active)

    ws = wb.create_sheet("Basic Client Details")
    ws.append(["Is Bank","Bank ID","Short Name","Full Name","Location","PAN",
               "Group","Related Party","Related Party Dropdown","Lender Class",
               "Lender Type","Lender Sub Type","TAN","Legal Entity Identifier","ID"])
    for i, l in enumerate(lenders, 1):
        ws.append([l["is_bank"], l["bank_id"], l["short"], l["full"], l["loc"], l["pan"],
                   l["grp"], "No", "Parent - Holding", l["cls"], l["typ"], l["sub"],
                   f"TAN{i:05d}", f"LEI{i:05d}", f"LEN-{i:03d}"])

    ws2 = wb.create_sheet("Demat Details")
    ws2.append(["DP Name","DP ID","Demat Account","Scheme Name","Client ID","Depository","ID"])
    for i, l in enumerate(lenders, 1):
        ws2.append([f"{l['short']} DP", f"DP{i:04d}", f"DEMAT{i:06d}",
                    "Scheme A", f"CL{i:05d}", "NSDL", f"LEN-{i:03d}"])

    ws3 = wb.create_sheet("Contact Details")
    ws3.append(["Contact Person","Contact Number","Landline Number","Email Address","Designation","ID"])
    for i, l in enumerate(lenders, 1):
        ws3.append([f"Contact {i}", f"98{i:08d}", f"080-2{i:07d}",
                    f"contact{i}@test.com", "Manager", f"LEN-{i:03d}"])

    ws4 = wb.create_sheet("Bank Account Details")
    ws4.append(["PAN","Account Name","Name","Branch","Address","IFSC Code",
                "MICR Code","Account Use","Account Number","ID"])
    for i, l in enumerate(lenders, 1):
        ws4.append([l["pan"], l["full"], l["full"], f"{l['loc']} Main",
                    f"{l['loc']} Address", f"{l['short'][:4].upper():4s}000{i:04d}",
                    f"56024{i:05d}", "Current", f"{10000+i}", f"LEN-{i:03d}"])

    ws5 = wb.create_sheet("Address Segment Details")
    ws5.append(["Office Location Type","Mark as Default","Office DUNS Number",
                "Address Line 1","Address Line 2","Address Line 3","City / Town",
                "District","State / Union Territory","Country","Pin Code","GSTIN",
                "Mobile Number","Email Address","Telephone Area Code",
                "Telephone Number(s)","Fax Area Code","Fax Number(s)","ID"])
    for i, l in enumerate(lenders, 1):
        ws5.append(["Registered Office","Yes",f"DUNS{i:05d}",l["loc"],"Line 2","",
                    l["loc"],l["loc"],"Maharashtra" if l["loc"]=="Mumbai" else "Karnataka",
                    "India","400001" if l["loc"]=="Mumbai" else "560001",
                    "27XXXXX1Z5",f"98{i:08d}",f"info{i}@test.com","080",f"2{i:07d}",
                    "080",f"2{i:07d}",f"LEN-{i:03d}"])

    wb.save("Lender_Regression_Filled.xlsx")
    print(f"✅ Lender_Regression_Filled.xlsx — {len(lenders)} lenders")

# ==========================================================
# GENERATE: SANCTION FILE
# ==========================================================
def fill_sanction_template(cfg, company, lenders, instruments):
    wb = Workbook()
    wb.remove(wb.active)
    basic, main, sub = [], [], []
    counter = 1
    start = cfg["start_date"]
    fy = cfg["fiscal_year"]

    for i in range(cfg["n_sanctions"]):
        l = lenders[i % len(lenders)]
        sref = f"SM-{fy}-{str(counter).zfill(5)}"
        counter += 1
        picked = random.sample(instruments, min(random.choice([1,2,2,3]), len(instruments)))
        basic.append({"ref": sref, "lender": l["full"],
                      "date": add_days(start, i*2), "company": company["name"]})
        for j, inst in enumerate(picked, 1):
            rule = INST_RULES[inst]
            nature = pick(["Revolving","Non-Revolving"]) if rule["rev"] else "Non-Revolving"
            amt = pick([1000000,2500000,5000000,7500000,10000000])
            has_sub = "Yes" if (rule["sub"] and random.random()<0.4) else "No"
            main.append({"ref":sref,"no":j,"inst":inst,"nat":nature,"amt":amt,
                         "valid":add_days(start,360),"has_sub":has_sub,"curr":rule["curr"]})
            if has_sub=="Yes":
                sub.append({"ref":sref,"no":j,"inst":inst,"nat":nature,"amt":amt,
                            "valid":add_days(start,360),"curr":rule["curr"]})

    ws = wb.create_sheet("Basic Details")
    ws.append(["Sanction Reference*","Lender Full Name*","Sanction Date*","Company Full Name"])
    for b in basic: ws.append([b["ref"], b["lender"], b["date"], b["company"]])

    ws2 = wb.create_sheet("Main Breakup Details")
    ws2.append(["Sanction Reference*","Main Breakup Number(Used for mapping sub limits)",
                "Instrument*","Nature of breakup*","Sanction Amount*",
                "Availability/Validity Date*","Has Sub Limits*","Currency*"])
    for m in main:
        ws2.append([m["ref"],m["no"],m["inst"],m["nat"],m["amt"],m["valid"],m["has_sub"],m["curr"]])

    ws3 = wb.create_sheet("Sub Breakup Details")
    ws3.append(["Sanction Reference*","Main Breakup Number(Used for mapping sub limits)",
                "Instrument*","Nature of breakup*","Sanction Amount",
                "Availability/Validity Date*","Currency*"])
    for s in sub:
        ws3.append([s["ref"],s["no"],s["inst"],s["nat"],s["amt"],s["valid"],s["curr"]])

    ws4 = wb.create_sheet("dropdowns")
    ws4.append(["Nature of Sanction","Has Sub Limits","Instruments","","Currency"])
    for row in [["Non-Revolving","Yes","Term Loan","","USD"],
                ["Revolving","No","Working Capital Demand Loan","","JPY"],
                ["","","Inter Company Deposit","","EUR"],
                ["","","Direct Assignment","","INR"],
                ["","","External Commercial Borrowing","","GBP"],
                ["","","Re-Finance","","CHF"],
                ["","","Line Of Credit","","QAR"],
                ["","","Short Term Loan","","CAD"],
                ["","","Cash Credit","","SEK"],
                ["","","Overdraft","","SGD"]]:
        ws4.append(row)

    wb.save("Sanction_Regression_Filled.xlsx")
    print(f"✅ Sanction_Regression_Filled.xlsx — {len(basic)} sanctions, {len(main)} breakups, {len(sub)} sub-limits")

# ==========================================================
# GENERATE: WCDL FILE
# ==========================================================
def fill_wcdl_template(cfg, company, lenders, n_deals):
    wb = Workbook()
    wb.remove(wb.active)

    sanctions = [{"ref": f"SM-{cfg['fiscal_year']}-{str(i+1).zfill(5)}",
                  "lender": lenders[i % len(lenders)],
                  "facility": f"WC-Rev-{str(i+1).zfill(4)}"}
                 for i in range(n_deals)]

    deals = []
    for i in range(n_deals):
        s = sanctions[i]
        ctype = pick(COUPON_TYPES)
        settle = add_days("2026-02-01", i*3)
        deals.append({
            "ID": f"WC-{str(i+1).zfill(3)}",
            "Is_Rollover": pick(["Yes","No"]),
            "Rollover_Type": pick(["With Interest","Without Interest"]),
            "Settlement_Date": settle,
            "Lender": s["lender"]["full"],
            "Curr_Acc": str(10001 + (i*2)),
            "Loan_Acc": str(20001 + (i*2)),
            "Loan_Acc_No": f"LAN{i+1:08d}",
            "Nature": pick(["Secured","Unsecured"]),
            "Security_Cover": pick([0,1,1,1]),
            "Company": company["name"],
            "Sanction_ID": s["ref"],
            "Facility_ID": s["facility"],
            "Disb_Date": settle,
            "Disb_Amt": pick([20000,50000,100000,150000]),
            "Repay_Date": add_days(settle, 90),
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

    ws = wb.create_sheet("Facility Details")
    ws.append(["ID*","Is Rollover","Rollover Type","Rollover ID","Settlement Date*",
               "Lender Full Name*","Current Bank Account*","Loan Bank Account",
               "Loan Account Number","Nature","Details Of Security","Security Cover",
               "Company Full Name*","Unsecured Loan Description",
               "Scheme Name (Re-finance)","Scheme Details (Re-finance)"])
    for d in deals:
        ws.append([d["ID"],d["Is_Rollover"],
                   d["Rollover_Type"] if d["Is_Rollover"]=="Yes" else "","",
                   d["Settlement_Date"],d["Lender"],d["Curr_Acc"],d["Loan_Acc"],
                   d["Loan_Acc_No"],d["Nature"],"",d["Security_Cover"],d["Company"],
                   "Unsecured Loan" if d["Nature"]=="Unsecured" else "","None",""])

    ws2 = wb.create_sheet("Ratings")
    ws2.append(["ID*","Rating Reference*","Remarks"])
    for d in deals: ws2.append([d["ID"], pick(["AAA","AA+","AA","A+","Unrated"]), "Sample"])

    ws3 = wb.create_sheet("Contact Master")
    ws3.append(["ID*","Contact Person Name","Comments"])
    for d in deals: ws3.append([d["ID"], f"Contact {d['ID']}", "Sample"])

    ws4 = wb.create_sheet("Sanction Details")
    ws4.append(["ID*","Sanction ID*","Sanction Facility ID*","Is EMI*",
                "EMI Override DCC","Maturity for EMI","PRINCIPAL MORATORIUM INSTALLMENT"])
    for d in deals: ws4.append([d["ID"],d["Sanction_ID"],d["Facility_ID"],"No","","",""])

    ws5 = wb.create_sheet("Disbursement Schedule")
    ws5.append(["ID*","Date*","Amount*","Payment Instructions","Remarks"])
    for d in deals: ws5.append([d["ID"],d["Disb_Date"],d["Disb_Amt"],"No","Auto"])

    ws6 = wb.create_sheet("Repayment Schedule")
    ws6.append(["ID*","Date*","Payment Date","Forward Rate (ECB)","Forward Date(ECB)",
                "Amount*","Invoice Number","Remarks"])
    for d in deals:
        ws6.append([d["ID"],d["Repay_Date"],d["Repay_Date"],"","",
                    d["Disb_Amt"],f"INV-{d['ID']}","Auto"])

    ws7 = wb.create_sheet("Interest details")
    ws7.append(["ID*","Period (Coupon Start Date)*","Coupon Type*","Coupon Payment Frequency*",
                "First Interest Due Date*","Fixed Rate","Next Rate Review Date",
                "Index (Benchmark Name)","Spread","DCC*"])
    for d in deals:
        ws7.append([d["ID"],d["Settlement_Date"],d["Coupon_Type"],d["Coupon_Freq"],
                    d["First_Int"],d["Fixed_Rate"],d["Next_Review"],
                    d["Index"],d["Spread"],d["DCC"]])

    ws8 = wb.create_sheet("TDS")
    ws8.append(["ID*","TDS","Start Date"])
    for d in deals: ws8.append([d["ID"],d["TDS"],d["Settlement_Date"]])

    ws10 = wb.create_sheet("Fees,Charges")
    ws10.append(["ID*","Invoice Date","Payment Due Date","Type","Amount",
                 "% of Disbursement","Counterparty Type","Counterparty Name",
                 "Is GST Applicable","GSTIN","Include for IndAS","TDS Type",
                 "Is RCM Applicable","Invoice Num","Details","Bank Account","Remarks"])
    for d in deals:
        for j in range(random.choice([1,2])):
            ws10.append([d["ID"],d["Disb_Date"],add_days(d["Disb_Date"],30),
                         pick(FEE_TYPES),pick([200,500,1000,2000,5000]),
                         round(random.uniform(0.05,1.0),2),
                         pick(["Lender","Arranger","Other"]),d["Lender"],
                         pick(["Yes","No"]),"27AAACH1234E1Z5",
                         pick(["Yes","No"]),pick(TDS_TYPES),
                         pick(["Yes","No"]),f"INV-{d['ID']}-{j+1}",
                         "Auto",d["Curr_Acc"],""])

    ws13 = wb.create_sheet("Pre Payment Penalty")
    ws13.append(["ID*","Upto Months","No Prepayment","Applicable Rate(%)","Remarks"])
    for d in deals:
        nop = pick(["Yes","No"])
        ws13.append([d["ID"],pick([3,6,9,12]),nop,
                     0 if nop=="Yes" else pick([1.0,1.5,2.0,2.5,3.0]),"Auto"])

    wb.save("WCDL_Regression_Filled.xlsx")
    print(f"✅ WCDL_Regression_Filled.xlsx — {len(deals)} deals")

# ==========================================================
# MAIN
# ==========================================================
def main():
    # If no argument, create input template
    if len(sys.argv) > 1 and sys.argv[1] == "--create-input":
        create_input_template()
        return

    input_file = sys.argv[1] if len(sys.argv) > 1 else "input_config.xlsx"

    print("\n" + "="*60)
    print("  BEACON REGRESSION TEMPLATE FILLER v2")
    print("="*60 + "\n")

    cfg, company, lenders, instruments = read_input(input_file)

    if lenders is None:
        lenders = [l for l in DEFAULT_LENDERS if l.get("use") == "Yes"]
    if instruments is None:
        instruments = list(INST_RULES.keys())

    random.seed(cfg.get("random_seed", 42))

    fill_lender_template(lenders, company)
    fill_sanction_template(cfg, company, lenders, instruments)
    fill_wcdl_template(cfg, company, lenders, int(cfg["n_wcdl"]))

    print("\n" + "="*60)
    print("✅ DONE! Files generated:")
    print("   1. Lender_Regression_Filled.xlsx")
    print("   2. Sanction_Regression_Filled.xlsx")
    print("   3. WCDL_Regression_Filled.xlsx")
    print("="*60 + "\n")

if __name__ == "__main__":
    main()