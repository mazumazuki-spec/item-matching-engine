print("START RUNNING... [VERSION 2.8 | SAFE PARTIAL + STRICT RULE FIX]")

import pandas as pd
import re
import time

file_path = "Stock Item SNT_sample.xlsx"
output_path = "output_matched_sample.xlsx"

SHEET_MASTER = "Master SNT"
SHEET_SYY = "SYY"
SHEET_NOTFOUND = "SYY-Not Found"

def clean_text(x):
    if pd.isna(x):
        return ""
    x = str(x).lower().strip()
    x = re.sub(r"[^a-z0-9ก-๙\s]", " ", x)
    x = re.sub(r"\s+", " ", x)
    return x

def get_words_from_clean(cleaned):
    return {
        w for w in cleaned.split()
        if len(w) >= 2
        and not w.isdigit()
        and w not in ["ml","cm","mm","kg","g","oz","pack","pcs","pc"]
    }

# =========================
# VERSION 2.8 RULE OVERRIDE
# เรียงจาก Specific ไป General
# =========================
RULES = [
    # BATTERY / เครื่องมือที่มีคำว่าแบตเตอรี่ แต่ไม่ใช่ Battery Spare Part
    (["ขั้วแบตเตอรี่"], "MAINTENANCE", "ENGINE SPARE PART", "CAR / PICKUP / VAN SPARE PART"),
    (["เครื่องตัดหญ้า", "แบตเตอรี่"], "ELECTRONIC EQUIPMENT", "ELECTRONIC EQUIPMENT", "ELECTRICAL EQUIPMENT"),
    (["เครื่องเป่าลม", "แบตเตอรี่"], "ELECTRONIC EQUIPMENT", "ELECTRONIC EQUIPMENT", "ELECTRICAL EQUIPMENT"),
    (["เครื่องพ่น", "ulv"], "ELECTRONIC EQUIPMENT", "ELECTRONIC EQUIPMENT", "ELECTRICAL EQUIPMENT"),
    (["เครื่องพ่นยา"], "GENERAL SUPPLY", "GARDEN", "OTHER GARDEN"),
    (["เครื่องเลื่อยโซ่"], "TOOL/MACHINERY/ENGINE", "TOOL", "ELECTRICAL TOOL"),
    (["เลื่อยไฟฟ้า"], "TOOL/MACHINERY/ENGINE", "TOOL", "ELECTRICAL TOOL"),
    (["แม็กยิงไร้สาย"], "TOOL/MACHINERY/ENGINE", "TOOL", "HAND TOOL"),
    (["ตู้ชารจแบตเตอรี่"], "ELECTRONIC EQUIPMENT", "ELECTRONIC EQUIPMENT", "OTHER ELECTRONIC EQUIPMENT"),
    (["ที่ชารจแบตเตอรี่"], "ELECTRONIC EQUIPMENT", "ELECTRONIC EQUIPMENT", "ELECTRICAL EQUIPMENT"),
    (["สายชารจแบตเตอรี่"], "ELECTRONIC EQUIPMENT", "ELECTRONIC EQUIPMENT SPARE PART", "BATTERY EQUIPMENT SPARE PART"),

    # PLUG แยกบริบท
    (["ปลั๊กสวมเร็ว"], "CONSTRUCTION", "MATERIALS", "WATER WORK"),
    (["ปลั๊กอุดท่อ"], "CONSTRUCTION", "MATERIALS", "WATER WORK"),
    (["ปลั๊กหัวเทียน"], "MAINTENANCE", "ENGINE SPARE PART", "MOTORCYCLE SPARE PART"),
    (["ปลั๊กไฟ"], "CONSTRUCTION", "MATERIALS", "ELECTRICAL&COMUNICATION WORK"),
    (["ปลั๊กพ่วง"], "CONSTRUCTION", "MATERIALS", "ELECTRICAL&COMUNICATION WORK"),
    (["เต้ารับ"], "CONSTRUCTION", "MATERIALS", "ELECTRICAL&COMUNICATION WORK"),

    # เคสที่ Partial เคยลากผิด
    (["พัดลมดูด"], "ELECTRONIC EQUIPMENT", "ELECTRONIC EQUIPMENT", "ELECTRICAL EQUIPMENT"),
    (["กีวี"], "FOOD", "FRESH/FROZEN FOOD", "FRUIT"),
    (["กุญแจ"], "GENERAL SUPPLY", "OTHER GENERAL", "GENERAL EQUIPMENT"),
    (["เกลือสระ"], "MAINTENANCE", "WATER TREATMENT / TESTING", "CHEMICAL"),
    (["คลิปหนีบ"], "GENERAL SUPPLY", "STATIONERY / PRINTING", "OTHER"),
    (["ค่าบริการเช่า"], "CHARGES/EXPENSES", "RENT", "RENT"),
    (["ค่าบริการ"], "CHARGES/EXPENSES", "EXPENSES/WAGE/CHARGE/REPAIR", "SERVICE CHARGE"),
    (["กระเป๋าใส่เงิน"], "GENERAL SUPPLY", "TEXTILE / CLOTHING", "ACCESSORIES"),

    # FOOD / BEVERAGE
    (["ไอศกรีม"], "FOOD", "FRESH/FROZEN FOOD", "ICE CREAM / SHERBET"),
    (["ice cream"], "FOOD", "FRESH/FROZEN FOOD", "ICE CREAM / SHERBET"),
    (["sherbet"], "FOOD", "FRESH/FROZEN FOOD", "ICE CREAM / SHERBET"),
    (["granola"], "FOOD", "DRY FOOD", "SNACK / MINI BAR"),
    (["กราโนล่า"], "FOOD", "DRY FOOD", "SNACK / MINI BAR"),
    (["แก้วมังกร"], "FOOD", "FRESH/FROZEN FOOD", "FRUIT"),
    (["ชากรีนที"], "BEVERAGE", "NON-ALCOHOL", "TEA"),
    (["ชาเขียว"], "BEVERAGE", "NON-ALCOHOL", "TEA"),

    # GENERAL
    (["กระดาษไข"], "GENERAL SUPPLY", "OTHER GENERAL", "GENERAL CONSUMABLE"),
    (["ตะกร้าผ้า"], "GENERAL SUPPLY", "OTHER GENERAL", "GENERAL EQUIPMENT"),
    (["ตะกร้าใส่ผัก"], "GENERAL SUPPLY", "KITCHENWARE", "OTHER KITCHENWARE"),
    (["ต้นเดฟ"], "GENERAL SUPPLY", "GARDEN", "PLANTS / TREE / SEED / FLOWER"),
    (["ต้นไม้"], "GENERAL SUPPLY", "GARDEN", "PLANTS / TREE / SEED / FLOWER"),
    (["แผ่นยางรองแก้ว"], "GENERAL SUPPLY", "OTHER GENERAL", "SOUVENIR"),

    # TOOL
    (["เทปวัดระยะ"], "TOOL/MACHINERY/ENGINE", "TOOL", "MEASUREMENT TOOL"),
    (["ตลับเมตร"], "TOOL/MACHINERY/ENGINE", "TOOL", "MEASUREMENT TOOL"),
    (["โฮลซอ"], "TOOL/MACHINERY/ENGINE", "TOOL", "TOOL PART / OTHER"),
    (["hole saw"], "TOOL/MACHINERY/ENGINE", "TOOL", "TOOL PART / OTHER"),
    (["makita", "เครื่องเจียร"], "TOOL/MACHINERY/ENGINE", "TOOL", "ELECTRICAL TOOL"),
    (["makita", "เครื่องเป่าลม"], "TOOL/MACHINERY/ENGINE", "TOOL", "ELECTRICAL TOOL"),
    (["makita", "เครื่องเลื่อย"], "TOOL/MACHINERY/ENGINE", "TOOL", "ELECTRICAL TOOL"),
    (["makita", "แท่นเลื่อย"], "TOOL/MACHINERY/ENGINE", "TOOL", "ELECTRICAL TOOL"),

    # FIRST AID
    (["bactoclav"], "GENERAL SUPPLY", "FIRST AID", "INTERNAL USE"),
    (["amoxicillin"], "GENERAL SUPPLY", "FIRST AID", "INTERNAL USE"),
    (["antibiotic"], "GENERAL SUPPLY", "FIRST AID", "INTERNAL USE"),

    # GLASS
    (["แก้ว", "wine"], "GENERAL SUPPLY", "KITCHENWARE", "GLASS WARE"),
    (["madison", "red wine"], "GENERAL SUPPLY", "KITCHENWARE", "GLASS WARE"),

    # ENGINE
    (["ไฟท้าย"], "ENGINE SPARE PART", "ENGINE SPARE PART", "CAR / PICKUP / VAN SPARE PART"),
]

def rule_override(cleaned_text):
    text = cleaned_text
    for keywords, main_group, sub_group, minor_group in RULES:
        if all(clean_text(k) in text for k in keywords):
            return main_group, sub_group, minor_group, " & ".join(keywords)
    return None

start_time = time.time()

print("Step 1: Reading Excel...")
master = pd.read_excel(file_path, sheet_name=SHEET_MASTER, dtype=str)
syy = pd.read_excel(file_path, sheet_name=SHEET_SYY, dtype=str)
notfound = pd.read_excel(file_path, sheet_name=SHEET_NOTFOUND, dtype=str)

print("Step 2: Setup columns...")

master_item_col = master.columns[11]
syy_match_col = syy.columns[7]
notfound_match_col = notfound.columns[7]

col_t_syy = syy.columns[19]
col_v_syy = syy.columns[21]
col_x_syy = syy.columns[23]

col_t_nf = notfound.columns[19]
col_v_nf = notfound.columns[21]
col_x_nf = notfound.columns[23]

out_ae = master.columns[30]
out_ag = master.columns[32]
out_ai = master.columns[34]
out_ak = master.columns[36]

print("Step 3: Pre-clean data...")

master["_clean"] = master[master_item_col].apply(clean_text)
master["_nospace"] = master["_clean"].str.replace(" ", "", regex=False)
master["_words"] = master["_clean"].apply(get_words_from_clean)

syy["_clean"] = syy[syy_match_col].apply(clean_text)
syy["_nospace"] = syy["_clean"].str.replace(" ", "", regex=False)
syy["_words"] = syy["_clean"].apply(get_words_from_clean)

notfound["_clean"] = notfound[notfound_match_col].apply(clean_text)
notfound["_nospace"] = notfound["_clean"].str.replace(" ", "", regex=False)
notfound["_words"] = notfound["_clean"].apply(get_words_from_clean)

print("Step 4: Build index...")
syy_dict = syy.drop_duplicates("_clean").set_index("_clean")
notfound_dict = notfound.drop_duplicates("_clean").set_index("_clean")

def phrase_match_fast(master_clean, ref_df):
    if len(master_clean) < 4:
        return None

    best_idx = None
    best_score = 0

    for idx, ref_clean in ref_df["_clean"].items():
        if master_clean in ref_clean or ref_clean in master_clean:
            score = min(len(master_clean), len(ref_clean))
            if score > best_score:
                best_score = score
                best_idx = idx

    return ref_df.loc[best_idx] if best_idx is not None else None

def keyword_match_fast(master_words, ref_df):
    best_idx = None
    best_score = 0

    for idx, ref_words in ref_df["_words"].items():
        common = master_words & ref_words
        strong = [w for w in common if len(w) >= 5]

        if len(common) >= 2 or len(strong) >= 1:
            score = len(common) * 10 + sum(len(w) for w in strong)
            if score > best_score:
                best_score = score
                best_idx = idx

    return ref_df.loc[best_idx] if best_idx is not None else None

def partial_match_fast(master_ns, ref_df):
    if len(master_ns) < 8:
        return None

    # คำเสี่ยง: ห้ามให้ Partial ลากเอง
    risky_words = [
        "แบตเตอรี่", "ปลั๊ก", "กุญแจ", "คลิปหนีบ",
        "ค่าบริการ", "เกลือ", "กระเป๋า", "พัดลม"
    ]

    for w in risky_words:
        if clean_text(w).replace(" ", "") in master_ns:
            return None

    best_idx = None
    best_score = 0
    second_score = 0
    best_longest = 0

    for idx, ref_ns in ref_df["_nospace"].items():
        if len(ref_ns) < 8:
            continue

        score = 0
        longest = 0

        for size in range(10, 5, -1):
            for j in range(len(master_ns) - size + 1):
                part = master_ns[j:j+size]
                if part in ref_ns:
                    score += size
                    longest = max(longest, size)

        if score > best_score:
            second_score = best_score
            best_score = score
            best_idx = idx
            best_longest = longest
        elif score > second_score:
            second_score = score

    if best_idx is not None:
        if best_score >= 35 and best_longest >= 8 and (best_score - second_score) >= 10:
            return ref_df.loc[best_idx]

    return None

print("Step 5: Matching...")

count_rule = 0
count_exact_syy = 0
count_exact_nf = 0
count_phrase_syy = 0
count_phrase_nf = 0
count_keyword_syy = 0
count_keyword_nf = 0
count_partial_syy = 0
count_partial_nf = 0
count_na = 0

total = len(master)

for i, row in master.iterrows():
    if i % 100 == 0:
        print(f"Processing row {i + 1}/{total}")

    key = row["_clean"]
    words = row["_words"]
    nospace = row["_nospace"]
    source = "N/A"

    rule_result = rule_override(key)

    if rule_result is not None:
        main_group, sub_group, minor_group, keyword_hit = rule_result
        master.at[i, out_ae] = main_group
        master.at[i, out_ag] = sub_group
        master.at[i, out_ai] = minor_group
        source = f"Rule Override: {keyword_hit}"
        count_rule += 1

    elif key in syy_dict.index:
        r = syy_dict.loc[key]
        master.at[i, out_ae] = r[col_t_syy]
        master.at[i, out_ag] = r[col_v_syy]
        master.at[i, out_ai] = r[col_x_syy]
        source = "Exact Match: SYY"
        count_exact_syy += 1

    elif key in notfound_dict.index:
        r = notfound_dict.loc[key]
        master.at[i, out_ae] = r[col_t_nf]
        master.at[i, out_ag] = r[col_v_nf]
        master.at[i, out_ai] = r[col_x_nf]
        source = "Exact Match: SYY_notfound"
        count_exact_nf += 1

    else:
        r = phrase_match_fast(key, syy)

        if r is not None:
            master.at[i, out_ae] = r[col_t_syy]
            master.at[i, out_ag] = r[col_v_syy]
            master.at[i, out_ai] = r[col_x_syy]
            source = "Phrase Match: SYY"
            count_phrase_syy += 1

        else:
            r = phrase_match_fast(key, notfound)

            if r is not None:
                master.at[i, out_ae] = r[col_t_nf]
                master.at[i, out_ag] = r[col_v_nf]
                master.at[i, out_ai] = r[col_x_nf]
                source = "Phrase Match: SYY_notfound"
                count_phrase_nf += 1

            else:
                r = keyword_match_fast(words, syy)

                if r is not None:
                    master.at[i, out_ae] = r[col_t_syy]
                    master.at[i, out_ag] = r[col_v_syy]
                    master.at[i, out_ai] = r[col_x_syy]
                    source = "Keyword Match: SYY"
                    count_keyword_syy += 1

                else:
                    r = keyword_match_fast(words, notfound)

                    if r is not None:
                        master.at[i, out_ae] = r[col_t_nf]
                        master.at[i, out_ag] = r[col_v_nf]
                        master.at[i, out_ai] = r[col_x_nf]
                        source = "Keyword Match: SYY_notfound"
                        count_keyword_nf += 1

                    else:
                        r = partial_match_fast(nospace, syy)

                        if r is not None:
                            master.at[i, out_ae] = r[col_t_syy]
                            master.at[i, out_ag] = r[col_v_syy]
                            master.at[i, out_ai] = r[col_x_syy]
                            source = "Partial Match: SYY"
                            count_partial_syy += 1

                        else:
                            r = partial_match_fast(nospace, notfound)

                            if r is not None:
                                master.at[i, out_ae] = r[col_t_nf]
                                master.at[i, out_ag] = r[col_v_nf]
                                master.at[i, out_ai] = r[col_x_nf]
                                source = "Partial Match: SYY_notfound"
                                count_partial_nf += 1

                            else:
                                count_na += 1

    master.at[i, out_ak] = source

print("Step 6: Cleaning helper columns...")
master = master.drop(columns=["_clean", "_nospace", "_words"])
syy = syy.drop(columns=["_clean", "_nospace", "_words"])
notfound = notfound.drop(columns=["_clean", "_nospace", "_words"])

print("Step 7: Saving output file...")
with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
    master.to_excel(writer, sheet_name=SHEET_MASTER, index=False)
    syy.to_excel(writer, sheet_name=SHEET_SYY, index=False)
    notfound.to_excel(writer, sheet_name=SHEET_NOTFOUND, index=False)

end_time = time.time()

print("================================")
print("DONE")
print("Output file:", output_path)
print("Rule Override:", count_rule)
print("Exact Match SYY:", count_exact_syy)
print("Exact Match SYY_notfound:", count_exact_nf)
print("Phrase Match SYY:", count_phrase_syy)
print("Phrase Match SYY_notfound:", count_phrase_nf)
print("Keyword Match SYY:", count_keyword_syy)
print("Keyword Match SYY_notfound:", count_keyword_nf)
print("Partial Match SYY:", count_partial_syy)
print("Partial Match SYY_notfound:", count_partial_nf)
print("Not Matched:", count_na)
print(f"Time used: {end_time - start_time:.2f} seconds")
print("================================")