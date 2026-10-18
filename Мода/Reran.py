# ============================================================
# FASHION RETAIL SALES — АНАЛИЗ ПРОДАЖ
# ============================================================
#
# Проект: анализ продаж одежды за 2022-2023
# Автор: Кудрявцева Диана
# Курс: "Аналитик данных с нуля", итоговый проект
# Дата: сентябрь 2026
#
# Что делает скрипт:
#   1. читает data/Fashion_Retail_Sales.csv
#   2. чистит данные (пропуски, выбросы, даты)
#   3. считает RFM-сегментацию
#   4. проверяет 6 гипотез (t-test, ANOVA, chi-square)
#   5. собирает Excel-отчёт с 12 листами и дашбордом
#
# Как запустить:
#   1. pip install -r requirements.txt
#   2. положить Fashion_Retail_Sales.csv в папку data/
#   3. python Мода.py
#   4. открыть output/Мода.xlsx
#
# Что не получилось / что дальше:
#   - прогноз выручки ненадёжен (R² < 0.05), тренда нет
#   - хочу попробовать K-means для сегментации
#   - дашборд пока статичный, можно добавить срезы
#   - в "Прочих" 45 клиентов, правило сегментации слишком узкое
#
# ============================================================

import pandas as pd
import numpy as np
import sqlite3
from scipy import stats
import warnings
import os
warnings.filterwarnings('ignore')

from openpyxl.styles import Font
from openpyxl.chart import BarChart, LineChart, PieChart, Reference

# пути
CSV_PATH = 'data/Fashion_Retail_Sales.csv'
EXCEL_PATH = 'output/Мода.xlsx'

# создаю папку output, если её нет — а то скрипт упадёт
os.makedirs('output', exist_ok=True)

# словарь перевода колонок, долго писала, зато потом красиво
perevod = {
    'customer_id': 'ID клиента',
    'item': 'Товар',
    'amount': 'Сумма',
    'purchase_date': 'Дата покупки',
    'rating': 'Рейтинг',
    'payment_method': 'Способ оплаты',
    'amount_missing': 'Пропуск суммы',
    'rating_missing': 'Пропуск рейтинга',
    'purchases': 'Количество покупок',
    'revenue': 'Выручка',
    'avg_check': 'Средний чек',
    'avg_rating': 'Средний рейтинг',
    'month': 'Месяц',
    'pct_share': 'Доля, %',
    'recency': 'Давность',
    'frequency': 'Частота',
    'monetary': 'Сумма',
    'R_score': 'Оценка давности',
    'F_score': 'Оценка частоты',
    'M_score': 'Оценка суммы',
    'RFM_score': 'Общая оценка RFM',
    'segment': 'Сегмент',
    'clients': 'Клиентов',
    'avg_recency': 'Средняя давность',
    'avg_frequency': 'Средняя частота',
    'avg_monetary': 'Средняя сумма',
    'total_revenue': 'Общая выручка',
    'hypothesis': 'Гипотеза',
    'test': 'Тест',
    'statistic': 'Статистика',
    'p_value': 'p-значение',
    'result': 'Результат',
    'cnt': 'Количество',
    'avg_amount': 'Средняя сумма',
    'total_rows': 'Всего строк',
    'unique_customers': 'Уникальных клиентов',
    'unique_items': 'Уникальных товаров',
    'first_purchase': 'Первая покупка',
    'last_purchase': 'Последняя покупка',
    'rows_with_missing_amount': 'Строк с пропуском суммы',
    'rows_with_missing_rating': 'Строк с пропуском рейтинга',
}

# сегменты отдельно, чтобы не запутаться
seg_ru = {
    'Champions': 'Чемпионы',
    'Loyal': 'Лояльные',
    'New': 'Новые',
    'At Risk': 'В зоне риска',
    'Lost': 'Потерянные',
    'Others': 'Прочие',
}

pay_ru = {
    'Credit Card': 'Кредитная карта',
    'Cash': 'Наличные',
}


# ============================================================
# ЭТАП 1: ЗАГРУЗКА И ЧИСТКА
# ============================================================

# --- читаю csv ---
print("читаю csv...")
df = pd.read_csv(CSV_PATH)
print(df.shape)
print(df.head())  # посмотреть глазами

# заливаю в sqlite, так удобнее агрегировать
conn = sqlite3.connect(':memory:')
df.to_sql('fashion_sales', conn, index=False, if_exists='replace')

print("исходно строк:", len(df))

# --- чистая таблица с флагами ---
# флаги нужны, чтобы потом сравнить средние с пропусками и без
conn.execute("""
CREATE TABLE fashion_sales_clean AS
SELECT
    "Customer Reference ID" AS customer_id,
    "Item Purchased" AS item,
    "Purchase Amount (USD)" AS amount,
    "Date Purchase" AS purchase_date,
    "Review Rating" AS rating,
    "Payment Method" AS payment_method,
    CASE WHEN "Purchase Amount (USD)" IS NULL THEN 1 ELSE 0 END AS amount_missing,
    CASE WHEN "Review Rating" IS NULL THEN 1 ELSE 0 END AS rating_missing
FROM fashion_sales
""")
conn.commit()

print(pd.read_sql("select sum(amount_missing), sum(rating_missing), count(*) from fashion_sales_clean", conn))

# --- выбросы ---
# 500 взяла примерно, посмотрела describe
# сумма <= 0 или > 500 — это явно ошибки ввода
conn.execute("""
delete from fashion_sales_clean
where amount is not null and (amount <= 0 or amount > 500)
""")
conn.commit()

# --- пропуски ---
# тут сначала хотела просто заполнить средним по всей колонке,
# но подумала, что у платьев и ботинок разные цены — будет криво.
# в итоге заполняю средним по товару, а если по товару всё пусто —
# общим средним. COALESCE как раз для этого.
conn.execute("""
update fashion_sales_clean as t
set amount = coalesce(
    (select avg(s.amount) from fashion_sales_clean as s where s.item = t.item and s.amount is not null),
    (select avg(amount) from fashion_sales_clean where amount is not null)
)
where t.amount is null
""")

conn.execute("""
update fashion_sales_clean as t
set rating = coalesce(
    (select avg(s.rating) from fashion_sales_clean as s where s.item = t.item and s.rating is not null),
    (select avg(rating) from fashion_sales_clean where rating is not null)
)
where t.rating is null
""")
conn.commit()

# --- даты ---
# в исходнике DD/MM/YYYY, sqlite хочет YYYY-MM-DD
# разбираю строку вручную — на курсе показывали такой способ
conn.execute("""
update fashion_sales_clean
set purchase_date =
    substr(purchase_date, 7, 4) || '-' ||
    substr(purchase_date, 4, 2) || '-' ||
    substr(purchase_date, 1, 2)
""")
conn.commit()

clean_df = pd.read_sql("select * from fashion_sales_clean", conn, parse_dates=['purchase_date'])
print("после чистки строк:", len(clean_df))
print("пропусков осталось:", clean_df.isna().sum().sum())


# ============================================================
# ЭТАП 2: АГРЕГАТЫ
# ============================================================

# --- топ товаров ---
top_items = pd.read_sql("""
select item,
       count(*) as purchases,
       round(sum(amount), 2) as revenue,
       round(avg(amount), 2) as avg_check,
       round(avg(rating), 2) as avg_rating
from fashion_sales_clean
group by item
order by revenue desc
limit 10
""", conn)

# --- по месяцам ---
# октябрь 2023 отрезаю — там данные обрезаны, резкий провал на графике
monthly = pd.read_sql("""
select strftime('%Y-%m', purchase_date) as month,
       count(*) as purchases,
       round(sum(amount), 2) as revenue,
       round(avg(amount), 2) as avg_check,
       round(avg(rating), 2) as avg_rating
from fashion_sales_clean
where purchase_date < '2023-10-01'
group by month
order by month
""", conn)

# --- оплата ---
payment = pd.read_sql("""
select payment_method,
       count(*) as purchases,
       round(sum(amount), 2) as revenue,
       round(avg(amount), 2) as avg_check,
       round(avg(rating), 2) as avg_rating,
       round(100.0 * count(*) / (select count(*) from fashion_sales_clean), 2) as pct_share
from fashion_sales_clean
group by payment_method
order by revenue desc
""", conn)


# ============================================================
# ЭТАП 3: RFM И ГИПОТЕЗЫ
# ============================================================

# --- RFM ---
# snapshot = день после последней покупки
print("считаю RFM...")
rfm = pd.read_sql("""
with snapshot as (
    select date(max(purchase_date), '+1 day') as snap_date
    from fashion_sales_clean
)
select customer_id,
       cast(julianday((select snap_date from snapshot)) - julianday(max(purchase_date)) as int) as recency,
       count(*) as frequency,
       round(sum(amount), 2) as monetary
from fashion_sales_clean
group by customer_id
""", conn)

# скоры 1..5
# для R чем меньше давность — тем выше скор (5 — лучший)
rfm['R_score'] = pd.qcut(rfm['recency'], 5, labels=[5,4,3,2,1]).astype(int)
# для F и M наоборот: чем больше, тем лучше
# rank(method='first') нужен, чтобы qcut не ругался на одинаковые границы
rfm['F_score'] = pd.qcut(rfm['frequency'].rank(method='first'), 5, labels=[1,2,3,4,5]).astype(int)
rfm['M_score'] = pd.qcut(rfm['monetary'], 5, labels=[1,2,3,4,5]).astype(int)
rfm['RFM_score'] = rfm['R_score'] + rfm['F_score'] + rfm['M_score']

# правила сегментов взяла из статьи на Habr, они стандартные
# TODO: проверить, почему в "Прочих" получилось 45 клиентов — правило слишком узкое
def segment(row):
    r = row['R_score']
    f = row['F_score']
    m = row['M_score']
    if r >= 4 and f >= 4 and m >= 4:
        return 'Champions'
    elif r >= 3 and f >= 4:
        return 'Loyal'
    elif r >= 4 and f <= 2:
        return 'New'
    elif r <= 2 and f >= 4:
        return 'At Risk'
    elif r <= 2 and f <= 2:
        return 'Lost'
    else:
        return 'Others'

rfm['segment'] = rfm.apply(segment, axis=1)

rfm_segments = rfm.groupby('segment').agg(
    clients=('customer_id', 'count'),
    avg_recency=('recency', 'mean'),
    avg_frequency=('frequency', 'mean'),
    avg_monetary=('monetary', 'mean'),
    total_revenue=('monetary', 'sum')
).round(2).reset_index().sort_values('avg_monetary', ascending=False)

print(rfm_segments)

# --- гипотезы ---
# если p < 0.05 — отвергаем H0
print("проверяю гипотезы...")
hypotheses = []

cc = clean_df[clean_df['payment_method'] == 'Credit Card']['amount']
cash = clean_df[clean_df['payment_method'] == 'Cash']['amount']
t, p = stats.ttest_ind(cc, cash, equal_var=False)
hypotheses.append({
    'hypothesis': 'H1: средний чек отличается у Credit Card и Cash',
    'test': 'Welch t-test',
    'statistic': round(t, 3),
    'p_value': round(p, 4),
    'result': 'Отвергаем H0' if p < 0.05 else 'Не отвергаем H0'
})

cc_r = clean_df[clean_df['payment_method'] == 'Credit Card']['rating']
cash_r = clean_df[clean_df['payment_method'] == 'Cash']['rating']
t, p = stats.ttest_ind(cc_r, cash_r, equal_var=False)
hypotheses.append({
    'hypothesis': 'H2: средний рейтинг отличается у Credit Card и Cash',
    'test': 'Welch t-test',
    'statistic': round(t, 3),
    'p_value': round(p, 4),
    'result': 'Отвергаем H0' if p < 0.05 else 'Не отвергаем H0'
})

groups = [g['amount'].values for _, g in clean_df.groupby('item') if len(g) >= 2]
f, p = stats.f_oneway(*groups)
hypotheses.append({
    'hypothesis': 'H3: средний чек отличается между товарами',
    'test': 'ANOVA',
    'statistic': round(f, 3),
    'p_value': round(p, 4),
    'result': 'Отвергаем H0' if p < 0.05 else 'Не отвергаем H0'
})

groups = [g['rating'].values for _, g in clean_df.groupby('item') if len(g) >= 2]
f, p = stats.f_oneway(*groups)
hypotheses.append({
    'hypothesis': 'H4: средний рейтинг отличается между товарами',
    'test': 'ANOVA',
    'statistic': round(f, 3),
    'p_value': round(p, 4),
    'result': 'Отвергаем H0' if p < 0.05 else 'Не отвергаем H0'
})

ct = pd.crosstab(clean_df['payment_method'], clean_df['item'])
chi2, p, dof, _ = stats.chi2_contingency(ct)
hypotheses.append({
    'hypothesis': 'H5: способ оплаты связан с товаром',
    'test': 'Chi-square',
    'statistic': round(chi2, 3),
    'p_value': round(p, 4),
    'result': 'Отвергаем H0' if p < 0.05 else 'Не отвергаем H0'
})

miss = clean_df[clean_df['amount_missing'] == 1]['amount']
nomiss = clean_df[clean_df['amount_missing'] == 0]['amount']
t, p = stats.ttest_ind(miss, nomiss, equal_var=False)
hypotheses.append({
    'hypothesis': 'H6: средний чек отличается у строк с пропуском и без',
    'test': 'Welch t-test',
    'statistic': round(t, 3),
    'p_value': round(p, 4),
    'result': 'Отвергаем H0' if p < 0.05 else 'Не отвергаем H0'
})

hypotheses_df = pd.DataFrame(hypotheses)
print(hypotheses_df)

# --- корреляции ---
pearson_r, pearson_p = stats.pearsonr(clean_df['amount'], clean_df['rating'])
spearman_r, spearman_p = stats.spearmanr(clean_df['amount'], clean_df['rating'])

# cramer's V формулу нашла в интернете, честно
def cramers_v(cm):
    chi2 = stats.chi2_contingency(cm)[0]
    n = cm.sum().sum()
    phi2 = chi2 / n
    r, k = cm.shape
    phi2corr = max(0, phi2 - ((k-1)*(r-1))/(n-1))
    rcorr = r - ((r-1)**2)/(n-1)
    kcorr = k - ((k-1)**2)/(n-1)
    return np.sqrt(phi2corr / min((kcorr-1), (rcorr-1)))

v = cramers_v(pd.crosstab(clean_df['payment_method'], clean_df['item']))
corr_matrix = clean_df[['amount', 'rating']].corr().round(4)

# --- влияние пропусков ---
missing_impact = pd.read_sql("""
select amount_missing,
       count(*) as cnt,
       round(avg(amount), 2) as avg_amount,
       round(avg(rating), 2) as avg_rating
from fashion_sales_clean
group by amount_missing
""", conn)

# --- сводка ---
summary = pd.read_sql("""
select count(*) as total_rows,
       count(distinct customer_id) as unique_customers,
       count(distinct item) as unique_items,
       round(sum(amount), 2) as total_revenue,
       round(avg(amount), 2) as avg_check,
       round(avg(rating), 2) as avg_rating,
       min(purchase_date) as first_purchase,
       max(purchase_date) as last_purchase,
       sum(amount_missing) as rows_with_missing_amount,
       sum(rating_missing) as rows_with_missing_rating
from fashion_sales_clean
""", conn)


# ============================================================
# ЭТАП 4: EXCEL-ОТЧЁТ
# ============================================================

# перевожу всё на русский — чтобы в Excel было читаемо
print("перевожу на русский...")
clean_df = clean_df.rename(columns=perevod)
top_items = top_items.rename(columns=perevod)
monthly = monthly.rename(columns=perevod)
payment = payment.rename(columns=perevod)
rfm = rfm.rename(columns=perevod)
rfm_segments = rfm_segments.rename(columns=perevod)
hypotheses_df = hypotheses_df.rename(columns=perevod)
missing_impact = missing_impact.rename(columns=perevod)
corr_matrix = corr_matrix.rename(columns=perevod, index=perevod)
summary = summary.rename(columns=perevod)

rfm['Сегмент'] = rfm['Сегмент'].map(seg_ru)
rfm_segments['Сегмент'] = rfm_segments['Сегмент'].map(seg_ru)
clean_df['Способ оплаты'] = clean_df['Способ оплаты'].map(pay_ru)
payment['Способ оплаты'] = payment['Способ оплаты'].map(pay_ru)

# --- собираю excel ---
# тут долго разбиралась с openpyxl, оставила как есть
print("собираю Excel...")

with pd.ExcelWriter(EXCEL_PATH, engine='openpyxl') as writer:
    clean_df.to_excel(writer, sheet_name='Данные', index=False)
    top_items.to_excel(writer, sheet_name='Топ товаров', index=False)
    monthly.to_excel(writer, sheet_name='По месяцам', index=False)
    payment.to_excel(writer, sheet_name='Оплата', index=False)
    rfm.to_excel(writer, sheet_name='RFM', index=False)
    rfm_segments.to_excel(writer, sheet_name='RFM сегменты', index=False)
    hypotheses_df.to_excel(writer, sheet_name='Гипотезы', index=False)
    corr_matrix.to_excel(writer, sheet_name='Корреляции', index=True)
    missing_impact.to_excel(writer, sheet_name='Пропуски', index=False)
    summary.to_excel(writer, sheet_name='Сводка', index=False)

    wb = writer.book
    ws = wb.create_sheet('Дашборд', 0)

    ws['A1'] = 'FASHION RETAIL SALES — ДАШБОРД'
    ws['A1'].font = Font(bold=True, size=16)

    ws['A3'] = 'КЛЮЧЕВЫЕ МЕТРИКИ'
    ws['A3'].font = Font(bold=True, size=12)

    # метрики прописываю руками, так виднее
    ws['A4'] = 'Всего строк'
    ws['B4'] = summary['Всего строк'][0]
    ws['A5'] = 'Уникальных клиентов'
    ws['B5'] = summary['Уникальных клиентов'][0]
    ws['A6'] = 'Уникальных товаров'
    ws['B6'] = summary['Уникальных товаров'][0]
    ws['A7'] = 'Общая выручка (USD)'
    ws['B7'] = summary['Общая выручка'][0]
    ws['A8'] = 'Средний чек (USD)'
    ws['B8'] = summary['Средний чек'][0]
    ws['A9'] = 'Средний рейтинг'
    ws['B9'] = summary['Средний рейтинг'][0]
    ws['A10'] = 'Первая покупка'
    ws['B10'] = str(summary['Первая покупка'][0])
    ws['A11'] = 'Последняя покупка'
    ws['B11'] = str(summary['Последняя покупка'][0])
    ws['A12'] = 'Пропусков в сумме'
    ws['B12'] = summary['Строк с пропуском суммы'][0]
    ws['A13'] = 'Пропусков в рейтинге'
    ws['B13'] = summary['Строк с пропуском рейтинга'][0]

    # формулы — это уже "продвинутая" часть, но очень хотелось попробовать
    ws['D3'] = 'ФОРМУЛЫ С УЧЁТОМ ФЛАГОВ'
    ws['D3'].font = Font(bold=True, size=12)
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

    # гипотезы
    ws['A16'] = 'ПРОВЕРКА ГИПОТЕЗ'
    ws['A16'].font = Font(bold=True, size=12)
    for j, col in enumerate(hypotheses_df.columns, start=1):
        ws.cell(row=17, column=j, value=col).font = Font(bold=True)
    for i, row in hypotheses_df.iterrows():
        for j, val in enumerate(row, start=1):
            ws.cell(row=18 + i, column=j, value=val)

    # корреляции
    ws['A26'] = 'КОРРЕЛЯЦИИ'
    ws['A26'].font = Font(bold=True, size=12)
    ws['A27'] = 'Pearson (сумма vs рейтинг)'
    ws['B27'] = round(float(pearson_r), 4)
    ws['A28'] = 'Spearman (сумма vs рейтинг)'
    ws['B28'] = round(float(spearman_r), 4)
    ws['A29'] = "Cramer's V (оплата vs товар)"
    ws['B29'] = round(float(v), 4)

    # rfm сегменты
    ws['D16'] = 'RFM-СЕГМЕНТЫ'
    ws['D16'].font = Font(bold=True, size=12)
    for j, col in enumerate(rfm_segments.columns, start=1):
        ws.cell(row=17, column=3 + j, value=col).font = Font(bold=True)
    for i, row in rfm_segments.iterrows():
        for j, val in enumerate(row, start=1):
            ws.cell(row=18 + i, column=3 + j, value=val)

    # автоширина — нашла пример в интернете
    for col in ws.columns:
        max_len = max((len(str(c.value)) for c in col if c.value), default=10)
        ws.column_dimensions[col[0].column_letter].width = min(max_len + 2, 40)

    # графики
    chart1 = LineChart()
    chart1.title = 'Выручка по месяцам'
    chart1.y_axis.title = 'USD'
    chart1.x_axis.title = 'Месяц'
    chart1.height = 8
    chart1.width = 16
    ms = wb['По месяцам']
    data = Reference(ms, min_col=3, min_row=1, max_row=len(monthly) + 1)
    cats = Reference(ms, min_col=1, min_row=2, max_row=len(monthly) + 1)
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
    data2 = Reference(ts, min_col=3, min_row=1, max_row=len(top_items) + 1)
    cats2 = Reference(ts, min_col=1, min_row=2, max_row=len(top_items) + 1)
    chart2.add_data(data2, titles_from_data=True)
    chart2.set_categories(cats2)
    ws.add_chart(chart2, 'H22')

    chart3 = PieChart()
    chart3.title = 'Доли способов оплаты'
    chart3.height = 8
    chart3.width = 12
    ps = wb['Оплата']
    data3 = Reference(ps, min_col=6, min_row=1, max_row=len(payment) + 1)
    cats3 = Reference(ps, min_col=1, min_row=2, max_row=len(payment) + 1)
    chart3.add_data(data3, titles_from_data=True)
    chart3.set_categories(cats3)
    ws.add_chart(chart3, 'H41')

    # словарь на отдельном листе, чтобы не забыть что как переводила
    ws_dict = wb.create_sheet('Словарь')
    ws_dict['A1'] = 'СЛОВАРЬ ПЕРЕВОДА'
    ws_dict['A1'].font = Font(bold=True, size=16)

    ws_dict['A3'] = 'Столбцы'
    ws_dict['A3'].font = Font(bold=True, size=12)
    ws_dict['A4'] = 'English'
    ws_dict['B4'] = 'Русский'
    ws_dict['A4'].font = Font(bold=True)
    ws_dict['B4'].font = Font(bold=True)
    row = 5
    for en, ru in perevod.items():
        ws_dict.cell(row=row, column=1, value=en)
        ws_dict.cell(row=row, column=2, value=ru)
        row += 1

    row += 2
    ws_dict.cell(row=row, column=1, value='Сегменты RFM').font = Font(bold=True, size=12)
    row += 1
    ws_dict.cell(row=row, column=1, value='English').font = Font(bold=True)
    ws_dict.cell(row=row, column=2, value='Русский').font = Font(bold=True)
    row += 1
    for en, ru in seg_ru.items():
        ws_dict.cell(row=row, column=1, value=en)
        ws_dict.cell(row=row, column=2, value=ru)
        row += 1

    row += 2
    ws_dict.cell(row=row, column=1, value='Способы оплаты').font = Font(bold=True, size=12)
    row += 1
    ws_dict.cell(row=row, column=1, value='English').font = Font(bold=True)
    ws_dict.cell(row=row, column=2, value='Русский').font = Font(bold=True)
    row += 1
    for en, ru in pay_ru.items():
        ws_dict.cell(row=row, column=1, value=en)
        ws_dict.cell(row=row, column=2, value=ru)
        row += 1

    ws_dict.column_dimensions['A'].width = 35
    ws_dict.column_dimensions['B'].width = 35

print("готово, файл:", EXCEL_PATH)
conn.close()
