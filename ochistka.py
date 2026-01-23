import pandas as pd
from openpyxl.utils import get_column_letter
from openpyxl.styles import numbers

INPUT_FILE = "/content/відомість.csv"
OUTPUT_FILE = "/content/відомість_результат.xlsx"
CHECK_FILE = "/content/відомість_переперевірка.xlsx"

# --------------------------------------------------
# 1. Зчитуємо ПЕРШІ 8 РЯДКІВ (заголовок)
# --------------------------------------------------
header_part = pd.read_csv(
    INPUT_FILE,
    sep=';',
    engine="python",
    encoding='cp1251',
    header=None,
    nrows=8
)

# --------------------------------------------------
# 2. Зчитуємо ОСНОВНІ ДАНІ з 9-го рядка
# --------------------------------------------------
df = pd.read_csv(
    INPUT_FILE,
    sep=';',
    engine="python",
    encoding='cp1251',
    skiprows=8,
    header=0,
    dtype=str
)

df.columns = df.columns.str.strip()
df = df.apply(lambda x: x.str.strip() if x.dtype == "object" else x)


# --- ПРАВИЛЬНОЕ ПРЕОБРАЗОВАНИЕ ЧИСЕЛ И ДАТ ---

# Система → число
df["Система"] = pd.to_numeric(df["Система"], errors="coerce")

# Даты
date_columns = [col for col in df.columns if "дата" in col.lower()]
for col in date_columns:
    df[col] = pd.to_datetime(
        df[col],
        errors="coerce",
        dayfirst=True
    )

# Числовые колонки
for col in df.columns:
    if col not in ["Ознака виду заборгованості"] and col not in date_columns:
        df[col] = pd.to_numeric(df[col], errors="ignore")



# --------------------------------------------------
# 4. Заповнення пустих значень у "Система"
# --------------------------------------------------
df["Система"] = df["Система"].ffill()

# --------------------------------------------------
# 5. Дебіторська заборгованість → ВИДАЛЯЄМО РЯДОК
# --------------------------------------------------
mask_debit = (
    df["Ознака виду заборгованості"]
    .str.lower()
    .str.contains("дебіторська заборгованість", na=False)
)
df = df.loc[~mask_debit]

# --------------------------------------------------
# 6. Видалення по системам
# --------------------------------------------------

# 2000, 5000, 6000, 7101 + зобов'язання
mask_1 = (
    df["Система"].isin([2000, 5000, 6000, 7101]) &
    (
        df["Ознака виду заборгованості"].isna() |
        df["Ознака виду заборгованості"]
        .str.lower()
        .str.contains("зобов'язання", na=False)
    )
)
df = df.loc[~mask_1]

# 7001, 7003 + кредитна заборгованість
mask_2 = (
    df["Система"].isin([7001, 7003]) &
    df["Ознака виду заборгованості"]
    .str.lower()
    .str.contains("кредитна заборгованість", na=False)
)
df = df.loc[~mask_2]

# --------------------------------------------------
# 7. ПЕРЕВІРКА
# --------------------------------------------------
systems_to_check = [
    7004, 2000, 5000, 6000,
    7001, 7003, 7101, 7150,
    7160, 9002
]

checks = {
    str(sys): df[df["Система"] == sys]
    for sys in systems_to_check
}

# --------------------------------------------------
# 8. Видалення рядка "ИТОГО"
# --------------------------------------------------
df = df[
    ~df.apply(
        lambda r: r.astype(str)
        .str.contains("ИТОГО", case=False)
        .any(),
        axis=1
    )
]

# --------------------------------------------------
# 9. Запис ОСНОВНОГО ФАЙЛУ з розбиттям
# --------------------------------------------------
MAX_EXCEL_ROWS = 1_048_576
SPLIT_LIMIT = 1_000_000

with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:

    # повертаємо заголовок (рядки 1–8)
    header_part.to_excel(
        writer,
        sheet_name="Аркуш1",
        index=False,
        header=False,
        startrow=0
    )

    start_row = len(header_part)

    if len(df) <= MAX_EXCEL_ROWS:
        df.to_excel(
            writer,
            sheet_name="Аркуш1",
            index=False,
            startrow=start_row
        )
    else:
        df.iloc[:SPLIT_LIMIT].to_excel(
            writer,
            sheet_name="Аркуш1",
            index=False,
            startrow=start_row
        )
        df.iloc[SPLIT_LIMIT:].to_excel(
            writer,
            sheet_name="Аркуш2",
            index=False
        )

    # --- ПРАВИЛЬНОЕ ФОРМАТИРОВАНИЕ EXCEL ---
    workbook = writer.book
    ws = workbook["Аркуш1"] # Get the worksheet object

    for idx, col in enumerate(df.columns, start=1):
        # ДАТЫ (по названию колонки)
        if "дата" in col.lower():
            for row_of_cells in ws.iter_rows(
                min_row=start_row + 2,
                min_col=idx,
                max_col=idx
            ):
                for cell in row_of_cells:
                    if cell.value is not None:
                        cell.number_format = 'dd.mm.yyyy' # Corrected format string

        # ЧИСЛА
        elif pd.api.types.is_numeric_dtype(df[col]):
            for row_of_cells in ws.iter_rows(
                min_row=start_row + 2,
                min_col=idx,
                max_col=idx
            ):
                for cell in row_of_cells:
                    if cell.value is not None:
                        cell.number_format = "#,##0.00"

    # Handle the second sheet if it exists
    if len(df) > MAX_EXCEL_ROWS:
        ws2 = workbook["Аркуш2"]
        for idx, col in enumerate(df.columns, start=1):
            if "дата" in col.lower():
                for row_of_cells in ws2.iter_rows(
                    min_row=2, # Data starts from row 2 (row 1 is pandas header)
                    min_col=idx,
                    max_col=idx
                ):
                    for cell in row_of_cells:
                        if cell.value is not None:
                            cell.number_format = 'dd.mm.yyyy'
            elif pd.api.types.is_numeric_dtype(df[col]):
                for row_of_cells in ws2.iter_rows(
                    min_row=2, # Data starts from row 2
                    min_col=idx,
                    max_col=idx
                ):
                    for cell in row_of_cells:
                        if cell.value is not None:
                            cell.number_format = numbers.FORMAT_NUMBER


# --------------------------------------------------
# 10. Запис ФАЙЛУ ПЕРЕПЕРЕВІРКИ
# --------------------------------------------------
with pd.ExcelWriter(CHECK_FILE, engine="openpyxl") as writer:
    for name, data in checks.items():
        if not data.empty:
            data.to_excel(writer, sheet_name=name, index=False)

print("✅ Готово")
print("📄 Основний файл:", OUTPUT_FILE)
print("🔍 Файл переперевірки:", CHECK_FILE)