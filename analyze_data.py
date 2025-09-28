import pandas as pd
from sklearn.linear_model import LogisticRegression
from sklearn.preprocessing import StandardScaler
import numpy as np
import datetime
import random

# --- Load Excel file ---
df = pd.read_excel("demo_leads.xlsx", engine='openpyxl')

# --- Extract relevant columns (J=Previous Purchases, K=Time Since Last Purchase, L=Average Purchase Value) ---
X = df.iloc[:, [9, 10, 11]]  # J=9, K=10, L=11 (0-indexed)
X.columns = ['Previous Purchases', 'Time Since Last Purchase', 'Average Purchase Value (SEK)']

# --- Step 1: Create target variable for Purchase Score (binary) ---
y_purchase = X['Time Since Last Purchase'].apply(lambda x: 1 if x < 200 else 0)

# --- Step 2: Scale features for logistic regression ---
scaler = StandardScaler()
X_scaled = scaler.fit_transform(X)

# --- Step 3: Train logistic regression ---
purchase_model = LogisticRegression()
purchase_model.fit(X_scaled, y_purchase)

# --- Step 4: Predict Purchase Score ---
purchase_scores = np.round(purchase_model.predict_proba(X_scaled)[:, 1], 2)

# --- Step 5: Calculate normalized LTV ---
historical_ltv = df.iloc[:, 9] * df.iloc[:, 11]  # Previous Purchases * Average Purchase Value
ltv_normalized = (historical_ltv - historical_ltv.min()) / (historical_ltv.max() - historical_ltv.min())
median = np.median(ltv_normalized)
ltv_normalized = np.clip(ltv_normalized / (2 * median), 0, 1)
ltv_normalized = np.round(ltv_normalized, 2)

# --- Step 6: Compute Lead Score with dynamic weighting ---
lead_score = []
for p_score, ltv_score in zip(purchase_scores, ltv_normalized):
    diff = abs(p_score - ltv_score)
    adjustment = 0.3 * diff
    if p_score > ltv_score:
        weight_p = 0.5 + adjustment
        weight_ltv = 0.5 - adjustment
    elif ltv_score > p_score:
        weight_ltv = 0.5 + adjustment
        weight_p = 0.5 - adjustment
    else:
        weight_p = weight_ltv = 0.5
    lead_score.append(round(p_score * weight_p + ltv_score * weight_ltv, 2))
lead_score = np.array(lead_score)

# --- Step 7: Remove existing columns if they exist ---
columns_to_remove = ['Purchase Score', 'Lifetime Value', 'Lead Score',
                     'Last Contact Date', 'Next Follow-up Date',
                     'Promo 1 Date','Promo 2 Date','Promo 3 Date','Promo 4 Date','Promo 5 Date','Promo 6 Date','Promo 7 Date',
                     'Education Date','Feedback Date','Welcome Date','Swedish/English']
for col in columns_to_remove:
    if col in df.columns:
        df.drop(columns=col, inplace=True)

# --- Step 8: Insert new columns ---
avg_col_index = df.columns.get_loc('Average Purchase Value (SEK)')
df.insert(avg_col_index + 1, 'Purchase Score', purchase_scores)
df.insert(avg_col_index + 2, 'Lifetime Value', ltv_normalized)
df.insert(avg_col_index + 3, 'Lead Score', lead_score)

# --- Step 9: Format dates and update based on Lead Score ---
today = datetime.date.today()

def get_random_date(days_range):
    return today - datetime.timedelta(days=random.randint(1, days_range))

# Initialize columns
promo_cols = [f'Promo {i} Date' for i in range(1, 8)]
df['Last Contact Date'] = np.nan
df['Next Follow-up Date'] = np.nan
df['Education Date'] = np.nan
df['Feedback Date'] = np.nan
df['Welcome Date'] = 'N/A'
df['Swedish/English'] = np.nan
for col in promo_cols:
    df[col] = 'N/A'

for idx, score in enumerate(lead_score):
    last_contact, next_followup = np.nan, np.nan
    if 0.8 <= score <= 1.0:
        last_contact = get_random_date(5)
        next_followup = last_contact + datetime.timedelta(days=5)
        step = 5
        promo_schedule = [last_contact + datetime.timedelta(days=step * i) for i in range(1, 9)]
        df.at[idx, 'Education Date'] = promo_schedule[1]
        df.at[idx, 'Feedback Date'] = promo_schedule[5]
        for i, col in enumerate(promo_cols):
            df.at[idx, col] = promo_schedule[i]
    elif 0.7 <= score <= 0.79:
        last_contact = get_random_date(7)
        next_followup = last_contact + datetime.timedelta(days=7)
        step = 7
        promo_schedule = [last_contact + datetime.timedelta(days=step * i) for i in range(1, 9)]
        df.at[idx, 'Education Date'] = promo_schedule[1]
        df.at[idx, 'Feedback Date'] = promo_schedule[3]
        for i, col in enumerate(promo_cols):
            df.at[idx, col] = promo_schedule[i]
    elif 0.6 <= score <= 0.69:
        last_contact = get_random_date(10)
        next_followup = last_contact + datetime.timedelta(days=10)
        df.at[idx, 'Education Date'] = last_contact + datetime.timedelta(days=20)
        df.at[idx, 'Feedback Date'] = last_contact + datetime.timedelta(days=30)
        promo_days = [10, 30, 30, 30, 30, 30, 30]
        prev_date = last_contact
        for i, col in enumerate(promo_cols):
            prev_date = prev_date + datetime.timedelta(days=promo_days[i])
            df.at[idx, col] = prev_date
    elif 0.4 <= score <= 0.59:
        last_contact = get_random_date(15)
        next_followup = last_contact + datetime.timedelta(days=15)
        df.at[idx, 'Education Date'] = last_contact + datetime.timedelta(days=15)
        df.at[idx, 'Feedback Date'] = df.at[idx, 'Education Date'] + datetime.timedelta(days=15)
    else:  # 0-0.39
        last_contact = get_random_date(30)
        next_followup = last_contact + datetime.timedelta(days=30)
        df.at[idx, 'Education Date'] = last_contact + datetime.timedelta(days=30)
        df.at[idx, 'Feedback Date'] = 'N/A'

    df.at[idx, 'Last Contact Date'] = last_contact
    df.at[idx, 'Next Follow-up Date'] = next_followup

    # Swedish/English column: 60% chance Swedish, 40% English
    df.at[idx, 'Swedish/English'] = 'Swedish' if random.random() < 0.6 else 'English'

    # Welcome Date based on Time Since Last Purchase
    time_since = df.at[idx, 'Time Since Last Purchase']
    if time_since == 1:
        df.at[idx, 'Welcome Date'] = today + datetime.timedelta(days=1)
    elif time_since == 2:
        df.at[idx, 'Welcome Date'] = today

# --- Step 10: Ensure all date columns are object dtype to allow 'N/A' ---
all_date_cols = ['Date Added', 'Last Contact Date', 'Next Follow-up Date',
                 'Education Date', 'Feedback Date', 'Welcome Date'] + promo_cols
for col in all_date_cols:
    if col in df.columns:
        df[col] = df[col].astype(object)

# --- Step 11: Convert actual datetime values to date only, leave 'N/A' ---
for col in all_date_cols:
    if col in df.columns:
        for idx, val in enumerate(df[col]):
            if isinstance(val, (datetime.datetime, datetime.date)):
                df.at[idx, col] = val.date() if isinstance(val, datetime.datetime) else val
            elif pd.isna(val):
                df.at[idx, col] = 'N/A'

# --- Step 12: Save back to Excel ---
df.to_excel("demo_leads_scored.xlsx", index=False, engine='openpyxl')

print("All scores, dates, and language assignments have been updated in 'demo_leads_scored.xlsx'.")
