# ============================================================
# FASHION RETAIL SALES — ФИНАЛЬНЫЙ ПАЙПЛАЙН (РУССКИЙ)
# Очистка + флаги + RFM + гипотезы + корреляции + Excel-дашборд
# ============================================================

import pandas as pd
import numpy as np
import sqlite3
from scipy import stats
import warnings
warnings.filterwarnings('ignore')

from openpyxl.styles import Font
from openpyxl.chart import BarChart, LineChart, PieChart, Reference

CSV_PATH = 'Fashion_Retail_Sales.csv'
EXCEL_PATH = 'Мода.xlsx'

# ============================================================
# СЛОВАРИ ПЕРЕВОДА
# ============================================================
COLUMN_MAP = {
    'customer_id':    'ID клиента',
    'item':           'Товар',
    'amount':         'Сумма',
    'purchase_date':  'Дата покупки',
    'rating':         'Рейтинг',
    'payment_method': 'Способ оплаты',
    'amount_missing': 'Пропуск суммы',
    'rating_missing': 'Пропуск рейтинга',
    'purchases':      'Количество покупок',
    'revenue':        'Выручка',
    'avg_check':      'Средний чек',
    'avg_rating':     'Средний рейтинг',
    'month':          'Месяц',
    'pct_share':      'Доля, %',
    'recency':        'Давность',
    'frequency':      'Частота',
    'monetary':       'Сумма',
    'R_score':        'Оценка давности',
    'F_score':        'Оценка частоты',
    'M_score':        'Оценка суммы',
    'RFM_score':      'Общая оценка RFM',
    'segment':        'Сегмент',
    'clients':        'Клиентов',
    'avg_recency':    'Средняя давность',
    'avg_frequency':  'Средняя частота',
    'avg_monetary':   'Средняя сумма',
    'total_revenue':  'Общая выручка',
    'hypothesis':     'Гипотеза',
    'test':           'Тест',
    'statistic':      'Статистика',
    'p_value':        'p-значение',
    'result':         'Результат',
    'cnt':            'Количество',
    'avg_amount':     'Средняя сумма',
    'total_rows':     'Всего строк',
    'unique_customers': 'Уникальных клиентов',
    'unique_items':   'Уникальных товаров',
    'first_purchase': 'Первая покупка',
    'last_purchase':  'Последняя покупка',
    'rows_with_missing_amount': 'Строк с пропуском суммы',
    'rows_with_missing_rating': 'Строк с пропуском рейтинга',
}

SEGMENT_MAP = {
    'Champions': 'Чемпионы',
    'Loyal':     'Лояльные',
    'New':       'Новые',
    'At Risk':   'В зоне риска',
    'Lost':      'Потерянные',
    'Others':    'Прочие',
}

PAYMENT_MAP = {
    'Credit Card': 'Кредитная карта',
    'Cash':        'Наличные',
}

# ============================================================
# 1. ЗАГРУЗКА CSV
# ============================================================
df = pd.read_csv(CSV_PATH)
conn = sqlite3.connect(':memory:')
df.to_sql('fashion_sales', conn, index=False, if_exists='replace')

print("=" * 60)
print("1. ИСХОДНЫЕ ДАННЫЕ")
print("=" * 60)
print(f"Строк: {len(df)} | Колонок: {len(df.columns)}")

# ============================================================
# 2. ОЧИЩЕННАЯ ТАБЛИЦА С ФЛАГАМИ
# ============================================================
conn.execute("""
CREATE TABLE fashion_sales_clean AS
SELECT
    "Customer Reference ID" AS customer_id,
    "Item Purchased"        AS item,
    "Purchase Amount (USD)" AS amount,
    "Date Purchase"         AS purchase_date,
    "Review Rating"         AS rating,
    "Payment Method"        AS payment_method,
    CASE WHEN "Purchase Amount (USD)" IS NULL THEN 1 ELSE 0 END AS amount_missing,
    CASE WHEN "Review Rating"         IS NULL THEN 1 ELSE 0 END AS rating_missing
FROM fashion_sales;
""")
conn.commit()

print("\n" + "=" * 60)
print("2. ФЛАГИ ПРОПУСКОВ")
print("=" * 60)
print(pd.read_sql("""
SELECT
    SUM(amount_missing) AS amount_missing,
    SUM(rating_missing) AS rating_missing,
    COUNT(*)            AS total
FROM fashion_sales_clean;
""", conn))

# ============================================================
# 3. УДАЛЯЕМ ВЫБРОСЫ
# ============================================================
conn.execute("""
DELETE FROM fashion_sales_clean
WHERE amount IS NOT NULL AND (amount <= 0 OR amount > 500);
""")
conn.commit()

# ============================================================
# 4. ЗАПОЛНЯЕМ ПРОПУСКИ
# ============================================================
conn.execute("""
UPDATE fashion_sales_clean AS t
SET amount = COALESCE(
    (SELECT AVG(s.amount) FROM fashion_sales_clean AS s
     WHERE s.item = t.item AND s.amount IS NOT NULL),
    (SELECT AVG(amount) FROM fashion_sales_clean WHERE amount IS NOT NULL)
)
WHERE t.amount IS NULL;
""")
conn.execute("""
UPDATE fashion_sales_clean AS t
SET rating = COALESCE(
    (SELECT AVG(s.rating) FROM fashion_sales_clean AS s
     WHERE s.item = t.item AND s.rating IS NOT NULL),
    (SELECT AVG(rating) FROM fashion_sales_clean WHERE rating IS NOT NULL)
)
WHERE t.rating IS NULL;
""")
conn.commit()

# ============================================================
# 5. ДАТА → ISO
# ============================================================
conn.execute("""
UPDATE fashion_sales_clean
SET purchase_date =
    SUBSTR(purchase_date, 7, 4) || '-' ||
    SUBSTR(purchase_date, 4, 2) || '-' ||
    SUBSTR(purchase_date, 1, 2);
""")
conn.commit()

clean_df = pd.read_sql(
    "SELECT * FROM fashion_sales_clean",
    conn,
    parse_dates=['purchase_date']
)

print("\n" + "=" * 60)
print("3. ДАННЫЕ ОЧИЩЕНЫ")
print("=" * 60)
print(f"Строк: {len(clean_df)} | Пропусков: {clean_df.isna().sum().sum()}")

# ============================================================
# 6. ТОП ТОВАРОВ
# ============================================================
top_items = pd.read_sql("""
SELECT
    item,
    COUNT(*)              AS purchases,
    ROUND(SUM(amount), 2) AS revenue,
    ROUND(AVG(amount), 2) AS avg_check,
    ROUND(AVG(rating), 2) AS avg_rating
FROM fashion_sales_clean
GROUP BY item
ORDER BY revenue DESC
LIMIT 10;
""", conn)

# ============================================================
# 7. ВЫРУЧКА ПО МЕСЯЦАМ
# ============================================================
monthly = pd.read_sql("""
SELECT
    STRFTIME('%Y-%m', purchase_date) AS month,
    COUNT(*)                          AS purchases,
    ROUND(SUM(amount), 2)             AS revenue,
    ROUND(AVG(amount), 2)             AS avg_check,
    ROUND(AVG(rating), 2)             AS avg_rating
FROM fashion_sales_clean
WHERE purchase_date < '2023-10-01'
GROUP BY month
ORDER BY month;
""", conn)

# ============================================================
# 8. СПОСОБЫ ОПЛАТЫ
# ============================================================
payment = pd.read_sql("""
SELECT
    payment_method,
    COUNT(*)              AS purchases,
    ROUND(SUM(amount), 2) AS revenue,
    ROUND(AVG(amount), 2) AS avg_check,
    ROUND(AVG(rating), 2) AS avg_rating,
    ROUND(100.0 * COUNT(*) /
          (SELECT COUNT(*) FROM fashion_sales_clean), 2) AS pct_share
FROM fashion_sales_clean
GROUP BY payment_method
ORDER BY revenue DESC;
""", conn)

# ============================================================
# 9. RFM
# ============================================================
rfm = pd.read_sql("""
WITH snapshot AS (
    SELECT DATE(MAX(purchase_date), '+1 day') AS snap_date
    FROM fashion_sales_clean
)
SELECT
    customer_id,
    CAST(JULIANDAY((SELECT snap_date FROM snapshot)) -
         JULIANDAY(MAX(purchase_date)) AS INT) AS recency,
    COUNT(*)              AS frequency,
    ROUND(SUM(amount), 2) AS monetary
FROM fashion_sales_clean
GROUP BY customer_id;
""", conn)

rfm['R_score'] = pd.qcut(rfm['recency'], 5, labels=[5,4,3,2,1]).astype(int)
rfm['F_score'] = pd.qcut(rfm['frequency'].rank(method='first'), 5,
                         labels=[1,2,3,4,5]).astype(int)
rfm['M_score'] = pd.qcut(rfm['monetary'], 5, labels=[1,2,3,4,5]).astype(int)
rfm['RFM_score'] = rfm['R_score'] + rfm['F_score'] + rfm['M_score']

def segment(row):
    r, f, m = row['R_score'], row['F_score'], row['M_score']
    if r >= 4 and f >= 4 and m >= 4: return 'Champions'
    elif r >= 3 and f >= 4:          return 'Loyal'
    elif r >= 4 and f <= 2:          return 'New'
    elif r <= 2 and f >= 4:          return 'At Risk'
    elif r <= 2 and f <= 2:          return 'Lost'
    else:                            return 'Others'

rfm['segment'] = rfm.apply(segment, axis=1)

rfm_segments = rfm.groupby('segment').agg(
    clients=('customer_id', 'count'),
    avg_recency=('recency', 'mean'),
    avg_frequency=('frequency', 'mean'),
    avg_monetary=('monetary', 'mean'),
    total_revenue=('monetary', 'sum')
).round(2).reset_index().sort_values('avg_monetary', ascending=False)

print("\n" + "=" * 60)
print("4. RFM-СЕГМЕНТАЦИЯ")
print("=" * 60)
print(rfm_segments.to_string(index=False))

# ============================================================
# 10. ГИПОТЕЗЫ
# ============================================================
hypotheses = []

cc   = clean_df[clean_df['payment_method'] == 'Credit Card']['amount']
cash = clean_df[clean_df['payment_method'] == 'Cash']['amount']
t, p = stats.ttest_ind(cc, cash, equal_var=False)
hypotheses.append({'hypothesis': 'H1: Средний чек отличается у Credit Card vs Cash',
                   'test': 'Welch t-test', 'statistic': round(t, 3),
                   'p_value': round(p, 4),
                   'result': 'Отвергаем H0' if p < 0.05 else 'Не отвергаем H0'})

cc_r   = clean_df[clean_df['payment_method'] == 'Credit Card']['rating']
cash_r = clean_df[clean_df['payment_method'] == 'Cash']['rating']
t, p = stats.ttest_ind(cc_r, cash_r, equal_var=False)
hypotheses.append({'hypothesis': 'H2: Средний рейтинг отличается у Credit Card vs Cash',
                   'test': 'Welch t-test', 'statistic': round(t, 3),
                   'p_value': round(p, 4),
                   'result': 'Отвергаем H0' if p < 0.05 else 'Не отвергаем H0'})

groups = [g['amount'].values for _, g in clean_df.groupby('item') if len(g) >= 2]
f, p = stats.f_oneway(*groups)
hypotheses.append({'hypothesis': 'H3: Средний чек отличается между товарами',
                   'test': 'One-way ANOVA', 'statistic': round(f, 3),
                   'p_value': round(p, 4),
                   'result': 'Отвергаем H0' if p < 0.05 else 'Не отвергаем H0'})

groups = [g['rating'].values for _, g in clean_df.groupby('item') if len(g) >= 2]
f, p = stats.f_oneway(*groups)
hypotheses.append({'hypothesis': 'H4: Средний рейтинг отличается между товарами',
                   'test': 'One-way ANOVA', 'statistic': round(f, 3),
                   'p_value': round(p, 4),
                   'result': 'Отвергаем H0' if p < 0.05 else 'Не отвергаем H0'})

ct = pd.crosstab(clean_df['payment_method'], clean_df['item'])
chi2, p, dof, _ = stats.chi2_contingency(ct)
hypotheses.append({'hypothesis': 'H5: Способ оплаты связан с товаром',
                   'test': 'Chi-square', 'statistic': round(chi2, 3),
                   'p_value': round(p, 4),
                   'result': 'Отвергаем H0' if p < 0.05 else 'Не отвергаем H0'})

miss   = clean_df[clean_df['amount_missing'] == 1]['amount']
nomiss = clean_df[clean_df['amount_missing'] == 0]['amount']
t, p = stats.ttest_ind(miss, nomiss, equal_var=False)
hypotheses.append({'hypothesis': 'H6: Средний чек отличается у строк с пропуском и без',
                   'test': 'Welch t-test', 'statistic': round(t, 3),
                   'p_value': round(p, 4),
                   'result': 'Отвергаем H0' if p < 0.05 else 'Не отвергаем H0'})

hypotheses_df = pd.DataFrame(hypotheses)

# ============================================================
# 11. КОРРЕЛЯЦИИ
# ============================================================
pearson_r, pearson_p   = stats.pearsonr(clean_df['amount'], clean_df['rating'])
spearman_r, spearman_p = stats.spearmanr(clean_df['amount'], clean_df['rating'])

def cramers_v(confusion_matrix):
    chi2 = stats.chi2_contingency(confusion_matrix)[0]
    n = confusion_matrix.sum().sum()
    phi2 = chi2 / n
    r, k = confusion_matrix.shape
    phi2corr = max(0, phi2 - ((k-1)*(r-1))/(n-1))
    rcorr = r - ((r-1)**2)/(n-1)
    kcorr = k - ((k-1)**2)/(n-1)
    return np.sqrt(phi2corr / min((kcorr-1), (rcorr-1)))

v = cramers_v(pd.crosstab(clean_df['payment_method'], clean_df['item']))
corr_matrix = clean_df[['amount', 'rating']].corr().round(4)

# ============================================================
# 12. ВЛИЯНИЕ ПРОПУСКОВ
# ============================================================
missing_impact = pd.read_sql("""
SELECT
    amount_missing,
    COUNT(*)              AS cnt,
    ROUND(AVG(amount), 2) AS avg_amount,
    ROUND(AVG(rating), 2) AS avg_rating
FROM fashion_sales_clean
GROUP BY amount_missing;
""", conn)

# ============================================================
# 13. СВОДКА
# ============================================================
summary = pd.read_sql("""
SELECT
    COUNT(*)                    AS total_rows,
    COUNT(DISTINCT customer_id) AS unique_customers,
    COUNT(DISTINCT item)        AS unique_items,
    ROUND(SUM(amount), 2)       AS total_revenue,
    ROUND(AVG(amount), 2)       AS avg_check,
    ROUND(AVG(rating), 2)       AS avg_rating,
    MIN(purchase_date)          AS first_purchase,
    MAX(purchase_date)          AS last_purchase,
    SUM(amount_missing)         AS rows_with_missing_amount,
    SUM(rating_missing)         AS rows_with_missing_rating
FROM fashion_sales_clean;
""", conn)

# ============================================================
# 14. ПЕРЕВОД НА РУССКИЙ
# ============================================================
clean_df       = clean_df.rename(columns=COLUMN_MAP)
top_items      = top_items.rename(columns=COLUMN_MAP)
monthly        = monthly.rename(columns=COLUMN_MAP)
payment        = payment.rename(columns=COLUMN_MAP)
rfm            = rfm.rename(columns=COLUMN_MAP)
rfm_segments   = rfm_segments.rename(columns=COLUMN_MAP)
hypotheses_df  = hypotheses_df.rename(columns=COLUMN_MAP)
missing_impact = missing_impact.rename(columns=COLUMN_MAP)
corr_matrix    = corr_matrix.rename(columns=COLUMN_MAP, index=COLUMN_MAP)
summary        = summary.rename(columns=COLUMN_MAP)

# Перевод значений
rfm['Сегмент']          = rfm['Сегмент'].map(SEGMENT_MAP)
rfm_segments['Сегмент'] = rfm_segments['Сегмент'].map(SEGMENT_MAP)
clean_df['Способ оплаты'] = clean_df['Способ оплаты'].map(PAYMENT_MAP)
payment['Способ оплаты']  = payment['Способ оплаты'].map(PAYMENT_MAP)

print("\n" + "=" * 60)
print("5. ПЕРЕВОД НА РУССКИЙ ВЫПОЛНЕН")
print("=" * 60)
print("Столбцы clean_data:")
print(list(clean_df.columns))

# ============================================================
# 15. EXCEL С ДАШБОРДОМ
# ============================================================
with pd.ExcelWriter(EXCEL_PATH, engine='openpyxl') as writer:
    clean_df.to_excel(writer,       sheet_name='Данные',          index=False)
    top_items.to_excel(writer,      sheet_name='Топ товаров',     index=False)
    monthly.to_excel(writer,        sheet_name='По месяцам',      index=False)
    payment.to_excel(writer,        sheet_name='Оплата',          index=False)
    rfm.to_excel(writer,            sheet_name='RFM',             index=False)
    rfm_segments.to_excel(writer,   sheet_name='RFM сегменты',    index=False)
    hypotheses_df.to_excel(writer,  sheet_name='Гипотезы',        index=False)
    corr_matrix.to_excel(writer,    sheet_name='Корреляции',      index=True)
    missing_impact.to_excel(writer, sheet_name='Пропуски',        index=False)
    summary.to_excel(writer,        sheet_name='Сводка',          index=False)

    wb = writer.book
    ws = wb.create_sheet('Дашборд', 0)

    bold16 = Font(bold=True, size=16)
    bold12 = Font(bold=True, size=12)
    bold   = Font(bold=True)

    ws['A1'] = 'FASHION RETAIL SALES — ДАШБОРД'
    ws['A1'].font = bold16

    ws['A3'] = 'КЛЮЧЕВЫЕ МЕТРИКИ'
    ws['A3'].font = bold12

    metrics = [
        ('Всего строк',              summary['Всего строк'][0]),
        ('Уникальных клиентов',      summary['Уникальных клиентов'][0]),
        ('Уникальных товаров',       summary['Уникальных товаров'][0]),
        ('Общая выручка (USD)',      summary['Общая выручка'][0]),
        ('Средний чек (USD)',        summary['Средний чек'][0]),
        ('Средний рейтинг',          summary['Средний рейтинг'][0]),
        ('Первая покупка',           str(summary['Первая покупка'][0])),
        ('Последняя покупка',        str(summary['Последняя покупка'][0])),
        ('Пропусков в сумме',        summary['Строк с пропуском суммы'][0]),
        ('Пропусков в рейтинге',     summary['Строк с пропуском рейтинга'][0]),
    ]
    for i, (k, v) in enumerate(metrics, start=4):
        ws.cell(row=i, column=1, value=k).font = bold
        ws.cell(row=i, column=2, value=v)

    # Формулы с русскими названиями листов и столбцов
    ws['D3'] = 'ФОРМУЛЫ С УЧЁТОМ ФЛАГОВ'
    ws['D3'].font = bold12

    ws['D4'] = 'Средний чек без пропусков'
    ws['E4'] = "=AVERAGEIFS(Данные!C:C, Данные!G:G, 0)"

    ws['D5'] = 'Средний чек с пропусками'
    ws['E5'] = "=AVERAGEIFS(Данные!C:C, Данные!G:G, 1)"

    ws['D6'] = 'Доля строк с пропуском суммы'
    ws['E6'] = "=SUM(Данные!G:G)/COUNTA(Данные!G:G)"
    ws['E6'].number_format = '0.00%'

    ws['D7'] = 'Доля строк с пропуском рейтинга'
    ws['E7'] = "=SUM(Данные!H:H)/COUNTA(Данные!H:H)"
    ws['E7'].number_format = '0.00%'

    ws['D8'] = 'Выручка Чемпионов'
    ws['E8'] = "=SUMIF('RFM'!G:G, \"Чемпионы\", 'RFM'!D:D)"

    ws['D9'] = 'Клиентов в Чемпионах'
    ws['E9'] = "=COUNTIF('RFM'!G:G, \"Чемпионы\")"

    # Гипотезы
    ws['A16'] = 'ПРОВЕРКА ГИПОТЕЗ'
    ws['A16'].font = bold12
    for j, col in enumerate(hypotheses_df.columns, start=1):
        ws.cell(row=17, column=j, value=col).font = bold
    for i, row in hypotheses_df.iterrows():
        for j, val in enumerate(row, start=1):
            ws.cell(row=18+i, column=j, value=val)

    # Корреляции
    ws['A26'] = 'КОРРЕЛЯЦИИ'
    ws['A26'].font = bold12
    ws['A27'] = 'Pearson Сумма vs Рейтинг'
    ws['B27'] = round(float(pearson_r), 4)
    ws['A28'] = 'Spearman Сумма vs Рейтинг'
    ws['B28'] = round(float(spearman_r), 4)
    ws['A29'] = "Cramér's V Оплата vs Товар"
    ws['B29'] = round(float(v), 4)

    # RFM-сегменты
    ws['D16'] = 'RFM-СЕГМЕНТЫ'
    ws['D16'].font = bold12
    for j, col in enumerate(rfm_segments.columns, start=1):
        ws.cell(row=17, column=3+j, value=col).font = bold
    for i, row in rfm_segments.iterrows():
        for j, val in enumerate(row, start=1):
            ws.cell(row=18+i, column=3+j, value=val)

    # Автоширина
    for col in ws.columns:
        max_len = max((len(str(c.value)) for c in col if c.value), default=10)
        ws.column_dimensions[col[0].column_letter].width = min(max_len + 2, 40)

    # Диаграммы
    chart1 = LineChart()
    chart1.title = 'Выручка по месяцам'
    chart1.y_axis.title = 'USD'
    chart1.x_axis.title = 'Месяц'
    chart1.height = 8
    chart1.width = 16
    ms = wb['По месяцам']
    data = Reference(ms, min_col=3, min_row=1, max_row=len(monthly)+1)
    cats = Reference(ms, min_col=1, min_row=2, max_row=len(monthly)+1)
    chart1.add_data(data, titles_from_data=True)
    chart1.set_categories(cats)
    ws.add_chart(chart1, 'H3')

    chart2 = BarChart()
    chart2.title = 'Топ-10 товаров по выручке'
    chart2.y_axis.title = 'USD'
    chart2.x_axis.title = 'Товар'
    chart2.height = 8
    chart2.width = 16
    ts = wb['Топ товаров']
    data2 = Reference(ts, min_col=3, min_row=1, max_row=len(top_items)+1)
    cats2 = Reference(ts, min_col=1, min_row=2, max_row=len(top_items)+1)
    chart2.add_data(data2, titles_from_data=True)
    chart2.set_categories(cats2)
    ws.add_chart(chart2, 'H22')

    chart3 = PieChart()
    chart3.title = 'Доли способов оплаты'
    chart3.height = 8
    chart3.width = 12
    ps = wb['Оплата']
    data3 = Reference(ps, min_col=6, min_row=1, max_row=len(payment)+1)
    cats3 = Reference(ps, min_col=1, min_row=2, max_row=len(payment)+1)
    chart3.add_data(data3, titles_from_data=True)
    chart3.set_categories(cats3)
    ws.add_chart(chart3, 'H41')

    # ========================================================
    # ЛИСТ СО СЛОВАРЁМ
    # ========================================================
    ws_dict = wb.create_sheet('Словарь')

    ws_dict['A1'] = 'СЛОВАРЬ ПЕРЕВОДА'
    ws_dict['A1'].font = bold16

    ws_dict['A3'] = 'Столбцы'
    ws_dict['A3'].font = bold12
    ws_dict['A4'] = 'English'
    ws_dict['B4'] = 'Русский'
    ws_dict['A4'].font = bold
    ws_dict['B4'].font = bold

    for i, (en, ru) in enumerate(COLUMN_MAP.items(), start=5):
        ws_dict.cell(row=i, column=1, value=en)
        ws_dict.cell(row=i, column=2, value=ru)

    start_row = 5 + len(COLUMN_MAP) + 2
    ws_dict.cell(row=start_row, column=1, value='Сегменты RFM').font = bold12
    ws_dict.cell(row=start_row+1, column=1, value='English').font = bold
    ws_dict.cell(row=start_row+1, column=2, value='Русский').font = bold
    for i, (en, ru) in enumerate(SEGMENT_MAP.items(), start=start_row+2):
        ws_dict.cell(row=i, column=1, value=en)
        ws_dict.cell(row=i, column=2, value=ru)

    start_row2 = start_row + 2 + len(SEGMENT_MAP) + 2
    ws_dict.cell(row=start_row2, column=1, value='Способы оплаты').font = bold12
    ws_dict.cell(row=start_row2+1, column=1, value='English').font = bold
    ws_dict.cell(row=start_row2+1, column=2, value='Русский').font = bold
    for i, (en, ru) in enumerate(PAYMENT_MAP.items(), start=start_row2+2):
        ws_dict.cell(row=i, column=1, value=en)
        ws_dict.cell(row=i, column=2, value=ru)

    ws_dict.column_dimensions['A'].width = 35
    ws_dict.column_dimensions['B'].width = 35

print("\n" + "=" * 60)
print(f"✅ EXCEL СОХРАНЁН: {EXCEL_PATH}")
print("=" * 60)
print("Листы:")
for s in ['Дашборд', 'Данные', 'Топ товаров', 'По месяцам', 'Оплата',
          'RFM', 'RFM сегменты', 'Гипотезы', 'Корреляции',
          'Пропуски', 'Сводка', 'Словарь']:
    print(f"  • {s}")

conn.close()
