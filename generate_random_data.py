# save as generate_demo_leads_v3.py
import pandas as pd
from faker import Faker
import random
from datetime import datetime, timedelta

fake = Faker("sv_SE")
industries = ["IT-tjänster", "Konsult", "Detaljhandel", "Bygg", "Marknadsföring"]
lead_sources = ["Webbplats", "Rekommendation", "LinkedIn", "Mässa", "Kallkontakt"]

rows = []
for i in range(750):
    fn = fake.first_name()
    ln = fake.last_name()
    email = f"{fn.lower()}.{ln.lower()}@example.com"
    phone = random.randint(0, 9)
    company = fake.company()
    industry = random.choice(industries)
    city = fake.city()
    country = "Sweden"
    lead_source = random.choice(lead_sources)

    # Previous Purchases distribution
    p = random.random()
    if p < 0.6:
        previous_purchases = random.randint(1, 10)
    elif p < 0.9:
        previous_purchases = random.randint(11, 50)
    else:
        previous_purchases = random.randint(51, 100)

    # Time Since Last Purchase and Average Purchase Value
    time_since_last_purchase = random.randint(1, 400)
    avg_purchase_value = random.randint(400, 10000)

    # Date Added must be at least as old as Time Since Last Purchase
    days_ago_added = random.randint(time_since_last_purchase, time_since_last_purchase + 365)
    date_added = datetime.today() - timedelta(days=days_ago_added)
    
    # Last Contact Date after Date Added
    last_contact = fake.date_between(start_date=date_added, end_date='today')
    next_followup = fake.date_between(start_date='today', end_date='+90d')

    rows.append({
        "First Name": fn,
        "Last Name": ln,
        "Email": email,
        "Phone": phone,
        "Company": company,
        "Industry": industry,
        "City": city,
        "Country": country,
        "Lead Source": lead_source,
        "Previous Purchases": previous_purchases,
        "Time Since Last Purchase": time_since_last_purchase,
        "Average Purchase Value (SEK)": avg_purchase_value,
        "Purchase Score": "",  # empty
        "Lifetime Value": "",  # empty
        "Lead Score": "",      # empty
        "Date Added": date_added.date(),
        "Last Contact Date": "",
        "Next Follow-up Date": "",
        "Swedish/English": "",
        "Promo 1 Date": "",
        "Promo 2 Date": "",
        "Promo 3 Date": "",
        "Promo 4 Date": "",
        "Promo 5 Date": "",
        "Promo 6 Date": "",
        "Promo 7 Date": "",
        "Education Date": "",
        "Feedback Date": "",
        "Welcome Date": ""
    })

df = pd.DataFrame(rows)
df.to_excel("demo_leads.xlsx", index=False)
print("Created demo_leads.xlsx with enhanced purchase data (all fake data)")
