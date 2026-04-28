print("Starting script")
import logging
logging.basicConfig(filename='debug.log', level=logging.INFO,
                    format='%(asctime)s - %(levelname)s - %(message)s')
logging.info("Starting script")

import time
start_time = time.time()

import pandas as pd
import warnings
warnings.filterwarnings("ignore", message="Could not infer format")
import os

INPUT_FILE = os.path.join(os.path.dirname(__file__), "відомість.csv")
OUTPUT_FILE = os.path.join(os.path.dirname(__file__), "відомість_результат.xlsx")

# --------------------------------------------------
# Максимальна кількість рядків у Excel-файлі (.xlsx)
MAX_EXCEL_ROWS = 1048576
max_data_rows = MAX_EXCEL_ROWS - 8

# --------------------------------------------------
# 1. Зчитуємо ПЕРШІ 8 РЯДКІВ (заголовок)
# --------------------------------------------------
try:
    with open(INPUT_FILE, encoding="cp1251", errors="ignore") as f:
        header_lines = [next(f).rstrip("\n").split(";") for _ in range(8)]
    print("✅ Header read successfully")
except Exception as e:
    print(f"❌ Error reading header: {e}")
    exit(1)

# --------------------------------------------------
# 2. Зчитуємо ОСНОВНІ ДАНІ з 9-го рядка
# --------------------------------------------------
print("📖 Reading main data...")
try:
    df = pd.read_csv(
        INPUT_FILE,
        sep=';',
        encoding='cp1251',
        skiprows=8,
        engine='python',
        on_bad_lines='warn',
        dtype=str
    )
    print(f"✅ Data read: {len(df)} rows, {len(df.columns)} columns")
except Exception as e:
    print(f"❌ Error reading data: {e}")
    exit(1)

# Видалення рядка з номерами колонок (перший рядок даних)
df = df.iloc[1:].reset_index(drop=True)

df.columns = df.columns.str.strip()
df = df.apply(lambda x: x.str.strip() if x.dtype == "object" else x)

# --------------------------------------------------
# 3. Перетворення типів даних
# --------------------------------------------------
print("🔄 Processing data transformations...")

# Система → число
df["Система"] = pd.to_numeric(df["Система"], errors="coerce")

# Дати
date_columns = [col for col in df.columns if "дата" in col.lower()]
for col in date_columns:
    df[col] = pd.to_datetime(df[col], format='%d.%m.%Y', errors='coerce')

# Числові колонки (десятковий роздільник — кома)
non_date_cols = [
    c for c in df.columns
    if c not in date_columns and c not in ["Система", "Ознака виду заборгованості"]
]
for col in non_date_cols:
    converted = (
        df[col]
        .astype(str)
        .str.replace(r"\s+", "", regex=True)
        .str.replace(",", ".", regex=False)
    )
    converted = pd.to_numeric(converted, errors="coerce")
    if converted.notna().sum() > df[col].notna().sum() * 0.5:
        df[col] = converted

# --------------------------------------------------
# 4. Заповнення пустих значень
# --------------------------------------------------
print("🔄 Filling empty values...")
df["Система"] = df["Система"].ffill()
if "Ідентифікаційний код/номер (ЕДРПОУ)" in df.columns:
    df["Ідентифікаційний код/номер (ЕДРПОУ)"] = df["Ідентифікаційний код/номер (ЕДРПОУ)"].ffill()

# --------------------------------------------------
# 5. Дебіторська заборгованість → ВИДАЛЯЄМО РЯДОК
# --------------------------------------------------
print("🗑️ Removing дебіторська заборгованість...")
mask_debit = (
    df["Ознака виду заборгованості"]
    .fillna("")
    .str.lower()
    .str.contains("дебіторська заборгованість", na=False)
)
df = df.loc[~mask_debit]

# --------------------------------------------------
# 6. Видалення по системам
# --------------------------------------------------
print("🗑️ Removing by systems...")

# 2000, 5000, 6000, 7101 + зобов'язання
mask_1 = (
    df["Система"].isin([2000, 5000, 6000, 7101]) &
    (
        df["Ознака виду заборгованості"].isna() |
        df["Ознака виду заборгованості"]
        .fillna("")
        .str.lower()
        .str.contains("зобов'язання", na=False)
    )
)
df = df.loc[~mask_1]

# 7001, 7003 + кредитна заборгованість
mask_2 = (
    df["Система"].isin([7001, 7003]) &
    df["Ознака виду заборгованості"]
    .fillna("")
    .str.lower()
    .str.contains("кредитна заборгованість", na=False)
)
df = df.loc[~mask_2]

# --------------------------------------------------
# 7. Видалення рядка "ИТОГО"
# --------------------------------------------------
print("🗑️ Removing ИТОГО rows...")
df = df[
    ~df.apply(
        lambda r: r.astype(str)
        .str.contains("ИТОГО", case=False)
        .any(),
        axis=1
    )
]

print(f"📊 Final dataset: {len(df)} rows")

# --------------------------------------------------
# 8. Запис Excel-файлу
# --------------------------------------------------
print("📝 Writing Excel output...")
with pd.ExcelWriter(OUTPUT_FILE, engine="xlsxwriter",
                     datetime_format='dd.mm.yyyy') as writer:
    if len(df) <= max_data_rows:
        df.to_excel(writer, sheet_name="відомість", index=False, startrow=8)
        workbook = writer.book
        ws = writer.sheets["відомість"]

        # Запис шапки (рядки 1–8, розбиття по комірках)
        for i, parts in enumerate(header_lines):
            for j, val in enumerate(parts):
                if val:
                    ws.write(i, j, val)

        # Форматування
        date_format = workbook.add_format({'num_format': 'dd.mm.yyyy'})
        number_format = workbook.add_format({'num_format': '#.##########'})
        for idx, col in enumerate(df.columns):
            if "дата" in col.lower():
                ws.set_column(idx, idx, None, date_format)
            elif pd.api.types.is_numeric_dtype(df[col]):
                ws.set_column(idx, idx, None, number_format)
    else:
        # Аркуш 1
        df1 = df.iloc[:max_data_rows]
        df2 = df.iloc[max_data_rows:]

        df1.to_excel(writer, sheet_name="відомість", index=False, startrow=8)
        workbook = writer.book
        ws = writer.sheets["відомість"]

        for i, parts in enumerate(header_lines):
            for j, val in enumerate(parts):
                if val:
                    ws.write(i, j, val)

        date_format = workbook.add_format({'num_format': 'dd.mm.yyyy'})
        number_format = workbook.add_format({'num_format': '#.##########'})
        for idx, col in enumerate(df.columns):
            if "дата" in col.lower():
                ws.set_column(idx, idx, None, date_format)
            elif pd.api.types.is_numeric_dtype(df[col]):
                ws.set_column(idx, idx, None, number_format)

        # Аркуш 2
        df2.to_excel(writer, sheet_name="Аркуш2", index=False)
        ws2 = writer.sheets["Аркуш2"]
        for idx, col in enumerate(df.columns):
            if "дата" in col.lower():
                ws2.set_column(idx, idx, None, date_format)
            elif pd.api.types.is_numeric_dtype(df[col]):
                ws2.set_column(idx, idx, None, number_format)

end_time = time.time()
print(f"⏱️ Execution time: {end_time - start_time:.2f} seconds")
print("✅ Готово")
print("📄 Основний файл:", OUTPUT_FILE)
