# ============================================================
# ЭТАП 1. PYTHON — ТОЛЬКО СЛОЖНОЕ
# Очистка + гипотезы + корреляции + K-means + прогноз
# На выходе: 5 CSV-файлов для Excel
# ============================================================

import pandas as pd
import numpy as np
import sqlite3
import os
from scipy import stats
from sklearn.cluster import KMeans
from sklearn.preprocessing import StandardScaler
from sklearn.linear_model import LinearRegression
import warnings
warnings.filterwarnings('ignore')

# ============================================================
# НАСТРОЙКИ — ПОЛНЫЙ ПУТЬ К CSV
# ============================================================
CSV_IN = r'C:\Users\Диана\Desktop\Проекты\Мода\Fashion_Retail_Sales.csv'

# Проверка, что файл существует
if not os.path.exists(CSV_IN):
    print(f"❌ Файл не найден: {CSV_IN}")
    print("Проверьте, что CSV лежит в папке проекта.")
    input("Нажмите Enter для выхода...")
    exit()

# ============================================================
# 1. ЗАГРУЗКА CSV
# ============================================================
print("=" * 60)
print("ЭТАП 1. PYTHON — СЛОЖНЫЕ РАСЧЁТЫ")
print("=" * 60)
print("\n[1/6] Загрузка данных...")

df = pd.read_csv(CSV_IN)
conn = sqlite3.connect(':memory:')
df.to_sql('sales', conn, index=False, if_exists='replace')
print(f"      Строк: {len(df)} | Колонок: {len(df.columns)}")

# ============================================================
# 2. ОЧИСТКА ДАННЫХ
# ============================================================
print("\n[2/6] Очистка данных...")

# 2.1. Создаём очищенную таблицу
conn.execute("""
CREATE TABLE clean AS
SELECT
    "Customer Reference ID" AS customer_id,
    "Item Purchased"        AS item,
    "Purchase Amount (USD)" AS amount,
    "Date Purchase"         AS purchase_date,
    "Review Rating"         AS rating,
    "Payment Method"        AS payment_method
FROM sales;
""")
conn.commit()

# 2.2. Удаляем выбросы (amount > 500 или <= 0)
conn.execute("""
DELETE FROM clean
WHERE amount IS NOT NULL AND (amount <= 0 OR amount > 500);
""")
conn.commit()

# 2.3. Заполняем пропуски в amount средним по товару
conn.execute("""
UPDATE clean AS t
SET amount = COALESCE(
    (SELECT AVG(s.amount) FROM clean AS s
     WHERE s.item = t.item AND s.amount IS NOT NULL),
    (SELECT AVG(amount) FROM clean WHERE amount IS NOT NULL)
)
WHERE t.amount IS NULL;
""")

# 2.4. Заполняем пропуски в rating средним по товару
conn.execute("""
UPDATE clean AS t
SET rating = COALESCE(
    (SELECT AVG(s.rating) FROM clean AS s
     WHERE s.item = t.item AND s.rating IS NOT NULL),
    (SELECT AVG(rating) FROM clean WHERE rating IS NOT NULL)
)
WHERE t.rating IS NULL;
""")
conn.commit()

# 2.5. Дата → ISO
conn.execute("""
UPDATE clean
SET purchase_date =
    SUBSTR(purchase_date, 7, 4) || '-' ||
    SUBSTR(purchase_date, 4, 2) || '-' ||
    SUBSTR(purchase_date, 1, 2);
""")
conn.commit()

# 2.6. Забираем в pandas
clean = pd.read_sql(
    "SELECT * FROM clean",
    conn,
    parse_dates=['purchase_date']
)

print(f"      После очистки: {len(clean)} строк")
print(f"      Пропусков: {clean.isna().sum().sum()}")

# 2.7. Сохраняем очищенные данные
clean.to_csv('data_clean.csv', index=False, encoding='utf-8-sig', sep=';', decimal=',')
print("      ✅ data_clean.csv")

# ============================================================
# 3. ГИПОТЕЗЫ
# ============================================================
print("\n[3/6] Проверка гипотез...")

hypotheses = []

# H1: t-test — средний чек CC vs Cash
cc   = clean[clean['payment_method'] == 'Credit Card']['amount']
cash = clean[clean['payment_method'] == 'Cash']['amount']
t, p = stats.ttest_ind(cc, cash, equal_var=False)
hypotheses.append({
    'Гипотеза': 'H1: Средний чек отличается у Credit Card vs Cash',
    'Тест': 'Welch t-test',
    'Статистика': round(t, 3),
    'p-значение': round(p, 4),
    'Результат': 'Отвергаем H0' if p < 0.05 else 'Не отвергаем H0'
})

# H2: t-test — средний рейтинг CC vs Cash
cc_r   = clean[clean['payment_method'] == 'Credit Card']['rating']
cash_r = clean[clean['payment_method'] == 'Cash']['rating']
t, p = stats.ttest_ind(cc_r, cash_r, equal_var=False)
hypotheses.append({
    'Гипотеза': 'H2: Средний рейтинг отличается у CC vs Cash',
    'Тест': 'Welch t-test',
    'Статистика': round(t, 3),
    'p-значение': round(p, 4),
    'Результат': 'Отвергаем H0' if p < 0.05 else 'Не отвергаем H0'
})

# H3: ANOVA — средний чек между товарами
groups = [g['amount'].values for _, g in clean.groupby('item') if len(g) >= 2]
f, p = stats.f_oneway(*groups)
hypotheses.append({
    'Гипотеза': 'H3: Средний чек отличается между товарами',
    'Тест': 'One-way ANOVA',
    'Статистика': round(f, 3),
    'p-значение': round(p, 4),
    'Результат': 'Отвергаем H0' if p < 0.05 else 'Не отвергаем H0'
})

# H4: ANOVA — средний рейтинг между товарами
groups = [g['rating'].values for _, g in clean.groupby('item') if len(g) >= 2]
f, p = stats.f_oneway(*groups)
hypotheses.append({
    'Гипотеза': 'H4: Средний рейтинг отличается между товарами',
    'Тест': 'One-way ANOVA',
    'Статистика': round(f, 3),
    'p-значение': round(p, 4),
    'Результат': 'Отвергаем H0' if p < 0.05 else 'Не отвергаем H0'
})

# H5: Chi-square — способ оплаты × товар
ct = pd.crosstab(clean['payment_method'], clean['item'])
chi2, p, dof, _ = stats.chi2_contingency(ct)
hypotheses.append({
    'Гипотеза': 'H5: Способ оплаты связан с товаром',
    'Тест': 'Chi-square',
    'Статистика': round(chi2, 3),
    'p-значение': round(p, 4),
    'Результат': 'Отвергаем H0' if p < 0.05 else 'Не отвергаем H0'
})

# H6: t-test — пропуски влияют на чек
orig_sales = pd.read_sql("""
SELECT
    "Purchase Amount (USD)" AS amount,
    CASE WHEN "Purchase Amount (USD)" IS NULL THEN 1 ELSE 0 END AS amount_missing
FROM sales
WHERE "Purchase Amount (USD)" IS NULL
   OR "Purchase Amount (USD)" BETWEEN 1 AND 500;
""", conn)

miss_amt = orig_sales[orig_sales['amount_missing'] == 1]['amount'].dropna()
nomiss_amt = orig_sales[orig_sales['amount_missing'] == 0]['amount'].dropna()

if len(miss_amt) >= 2 and len(nomiss_amt) >= 2:
    t, p = stats.ttest_ind(miss_amt, nomiss_amt, equal_var=False)
    hypotheses.append({
        'Гипотеза': 'H6: Средний чек отличается у строк с пропуском',
        'Тест': 'Welch t-test',
        'Статистика': round(t, 3),
        'p-значение': round(p, 4),
        'Результат': 'Отвергаем H0' if p < 0.05 else 'Не отвергаем H0'
    })

# H7: ANOVA — рейтинг по месяцам
groups = [g['rating'].values for _, g in clean.groupby(
    clean['purchase_date'].dt.to_period('M')) if len(g) >= 2]
f, p = stats.f_oneway(*groups)
hypotheses.append({
    'Гипотеза': 'H7: Средний рейтинг отличается по месяцам',
    'Тест': 'One-way ANOVA',
    'Статистика': round(f, 3),
    'p-значение': round(p, 4),
    'Результат': 'Отвергаем H0' if p < 0.05 else 'Не отвергаем H0'
})

# H8: ANOVA — чек по дням недели
groups = [g['amount'].values for _, g in clean.groupby(
    clean['purchase_date'].dt.dayofweek) if len(g) >= 2]
f, p = stats.f_oneway(*groups)
hypotheses.append({
    'Гипотеза': 'H8: Средний чек отличается по дням недели',
    'Тест': 'One-way ANOVA',
    'Статистика': round(f, 3),
    'p-значение': round(p, 4),
    'Результат': 'Отвергаем H0' if p < 0.05 else 'Не отвергаем H0'
})

hypotheses_df = pd.DataFrame(hypotheses)
hypotheses_df.to_csv('hypotheses.csv', index=False, encoding='utf-8-sig')
print("      ✅ hypotheses.csv")
print(f"      Гипотез: {len(hypotheses_df)}")
for _, row in hypotheses_df.iterrows():
    print(f"         {row['Гипотеза'][:60]:60} → {row['Результат']}")

# ============================================================
# 4. КОРРЕЛЯЦИИ
# ============================================================
print("\n[4/6] Расчёт корреляций...")

# Pearson
pearson_r, pearson_p = stats.pearsonr(clean['amount'], clean['rating'])

# Spearman
spearman_r, spearman_p = stats.spearmanr(clean['amount'], clean['rating'])

# Cramér's V
def cramers_v(confusion_matrix):
    chi2 = stats.chi2_contingency(confusion_matrix)[0]
    n = confusion_matrix.sum().sum()
    phi2 = chi2 / n
    r, k = confusion_matrix.shape
    phi2corr = max(0, phi2 - ((k-1)*(r-1))/(n-1))
    rcorr = r - ((r-1)**2)/(n-1)
    kcorr = k - ((k-1)**2)/(n-1)
    return np.sqrt(phi2corr / min((kcorr-1), (rcorr-1)))

v = cramers_v(pd.crosstab(clean['payment_method'], clean['item']))

# Point-biserial
clean['payment_bin'] = (clean['payment_method'] == 'Credit Card').astype(int)
pb_r, pb_p = stats.pointbiserialr(clean['payment_bin'], clean['amount'])

correlations_df = pd.DataFrame([
    {'Метрика': 'Pearson: Сумма vs Рейтинг',
     'Коэффициент': round(pearson_r, 4),
     'p-значение': round(pearson_p, 4),
     'Интерпретация': 'Связи нет' if abs(pearson_r) < 0.3 else 'Слабая связь'},
    {'Метрика': 'Spearman: Сумма vs Рейтинг',
     'Коэффициент': round(spearman_r, 4),
     'p-значение': round(spearman_p, 4),
     'Интерпретация': 'Связи нет' if abs(spearman_r) < 0.3 else 'Слабая связь'},
    {'Метрика': "Cramér's V: Оплата vs Товар",
     'Коэффициент': round(v, 4),
     'p-значение': None,
     'Интерпретация': 'Независимы' if v < 0.3 else 'Связаны'},
    {'Метрика': 'Point-biserial: Оплата vs Сумма',
     'Коэффициент': round(pb_r, 4),
     'p-значение': round(pb_p, 4),
     'Интерпретация': 'Связи нет' if abs(pb_r) < 0.3 else 'Слабая связь'},
])
correlations_df.to_csv('correlations.csv', index=False, encoding='utf-8-sig')
print("      ✅ correlations.csv")

# ============================================================
# 5. K-MEANS КЛАСТЕРИЗАЦИЯ
# ============================================================
print("\n[5/6] K-means кластеризация...")

# 5.1. RFM-показатели
rfm = pd.read_sql("""
WITH snap AS (SELECT DATE(MAX(purchase_date), '+1 day') AS d FROM clean)
SELECT
    customer_id,
    CAST(JULIANDAY((SELECT d FROM snap)) -
         JULIANDAY(MAX(purchase_date)) AS INT) AS recency,
    COUNT(*)              AS frequency,
    ROUND(SUM(amount), 2) AS monetary
FROM clean
GROUP BY customer_id;
""", conn)

# 5.2. Масштабирование
X = rfm[['recency', 'frequency', 'monetary']].copy()
scaler = StandardScaler()
X_scaled = scaler.fit_transform(X)

# 5.3. K-means
kmeans = KMeans(n_clusters=4, random_state=42, n_init=10)
rfm['cluster'] = kmeans.fit_predict(X_scaled)

# 5.4. Расшифровка кластеров
cluster_summary = rfm.groupby('cluster').agg(
    clients=('customer_id', 'count'),
    avg_recency=('recency', 'mean'),
    avg_frequency=('frequency', 'mean'),
    avg_monetary=('monetary', 'mean')
).round(2).reset_index()

# 5.5. Присваиваем понятные названия
def cluster_name(row):
    r, f, m = row['avg_recency'], row['avg_frequency'], row['avg_monetary']
    if r < 15 and f > 22 and m > 2400:
        return 'VIP'
    elif r < 20 and f > 18:
        return 'Активные'
    elif r > 35:
        return 'Потерянные'
    else:
        return 'Средние'

cluster_summary['Название'] = cluster_summary.apply(cluster_name, axis=1)

# 5.6. Присваиваем названия клиентам
cluster_map = dict(zip(cluster_summary['cluster'], cluster_summary['Название']))
rfm['cluster_name'] = rfm['cluster'].map(cluster_map)

rfm.to_csv('clusters.csv', index=False, encoding='utf-8-sig')
print("      ✅ clusters.csv")
print(f"      Клиентов: {len(rfm)}")
print("\n      Кластеры:")
print(cluster_summary.to_string(index=False))

# ============================================================
# 6. ПРОГНОЗ ВЫРУЧКИ
# ============================================================
print("\n[6/6] Прогноз выручки...")

# 6.1. Месячная выручка
monthly = pd.read_sql("""
SELECT
    STRFTIME('%Y-%m', purchase_date) AS month,
    ROUND(SUM(amount), 2) AS revenue
FROM clean
WHERE purchase_date < '2023-10-01'
GROUP BY month
ORDER BY month;
""", conn)

# 6.2. Линейная регрессия
monthly['t'] = range(len(monthly))
model = LinearRegression()
model.fit(monthly[['t']], monthly['revenue'])

# 6.3. Прогноз на 3 месяца
future = pd.DataFrame({'t': range(len(monthly), len(monthly)+3)})
future['month'] = ['2023-10', '2023-11', '2023-12']
future['forecast'] = model.predict(future[['t']]).round(2)

# 6.4. Добавляем R²
r2 = model.score(monthly[['t']], monthly['revenue'])

forecast_df = pd.concat([
    monthly[['month', 'revenue']].rename(columns={'revenue': 'fact'}),
    future[['month', 'forecast']]
], ignore_index=True)

forecast_df.to_csv('forecast.csv', index=False, encoding='utf-8-sig')
print("      ✅ forecast.csv")
print(f"      R² модели: {r2:.4f}")
print("\n      Прогноз:")
print(future[['month', 'forecast']].to_string(index=False))

# ============================================================
# ИТОГ
# ============================================================
print("\n" + "=" * 60)
print("✅ ЭТАП 1 ЗАВЕРШЁН")
print("=" * 60)
print("\nСозданы файлы:")
print("  • data_clean.csv    — очищенные данные")
print("  • hypotheses.csv    — гипотезы")
print("  • correlations.csv  — корреляции")
print("  • clusters.csv      — K-means кластеры")
print("  • forecast.csv      — прогноз на 3 месяца")
print("\nТеперь эти файлы нужно подключить в Excel (ЭТАП 2).")

conn.close()

input("\nНажмите Enter для выхода...")
