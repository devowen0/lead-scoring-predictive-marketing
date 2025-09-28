# Lead Scoring & Predictive Marketing Automation

## Overview
This project is a complete lead scoring and follow-up automation system that demonstrates skills in data science, machine learning, marketing analytics, process automation, and reporting.

It takes raw lead data, analyzes purchase patterns, calculates Purchase Score, Lifetime Value, and Lead Score, assigns dynamic follow-up schedules, facilitates email campaign execution, and generates PDF reports with visual insights. 

Lead Score is a 0–1 metric that prioritizes leads based on their potential future revenue impact. It combines Purchase Score (likelihood of buying again) and Lifetime Value (historical spending) to ensure we target the people most likely to generate meaningful revenue — not just those likely to buy, but those whose purchases will have the biggest impact.

This simulates a real-world marketing workflow that could be used by sales teams to prioritize leads, reduce churn, and increase conversion rates.

---
## First do:
pip install -r requirements.txt
## or:
pip3 install -r requirements.txt

## How to Run:
## 1. Generate synthetic leads
python generate_random_data.py

## 2. Analyze & score leads
python analyze_data.py

## 3. View campaign matches and send emails for a given date
python message.py

## 4. Generate PDF report
python pdf.py

## Concepts Demonstrated

- **Data Science & Machine Learning**
  - Built a logistic regression model to estimate purchase probability from historical purchase behavior.
  - Applied feature scaling using `StandardScaler` for better model performance.
  - Engineered **dynamic weighting** between purchase probability and lifetime value to compute a nuanced **Lead Score**.

- **Automation**
  - Automates lead generation with realistic synthetic data.
  - Automates date assignments dynamically based on lead score.
  - Automated email sending workflow** with clipboard-based copy-paste assistance.

- **Data Visualization & Reporting**
  - Created **pie charts, bar charts, and summary tables** to visualize lead distribution, average scores, and revenue.
  - Generated **professional PDF reports** summarizing insights for decision-makers.

- **Software Engineering**
  - Wrote modular Python scripts.
  - Used `pandas`, `scikit-learn`, `matplotlib`, `fpdf`, and `openpyxl`.
  - Saved data back into Excel with added insights and status tracking.

---

## Core Concepts

### Purchase Score
The **Purchase Score** predicts how likely a lead is to make a purchase in the near future.

Using logistic regression, we transform purchase-related variables into a probability between `0` and `1`.

**Variables used:**
- Previous Purchases  
- Time Since Last Purchase  
- Average Purchase Value (SEK)

**Equation:**

```math
PurchaseScore = 1 / (1 + e^{-(β₀ + β₁x₁ + β₂x₂ + β₃x₃)})
```

## Where:

x₁ = Previous Purchases

x₂ = Time Since Last Purchase

x₃ = Average Purchase Value (scaled)

β₀ ... β₃ = model coefficients learned by logistic regression

This allows the score to represent probability of conversion, a fundamental metric for lead prioritization.

### Lifetime Value (LTV)
LTV measures the historical monetary value of a lead. It captures how much revenue a lead has generated in the past and is normalized to combine fairly with Purchase Score.

**Calculation:**
LTV_normalized = (Previous Purchases × Average Purchase Value − min) / (max − min)

This ensures the Lead Score considers both likelihood to buy **and** the potential revenue impact of that purchase.

### Lead Score
The Lead Score is a weighted combination of Purchase Score and Lifetime Value. 

- If Purchase Score > LTV, give more weight to Purchase Score.  
- If LTV > Purchase Score, give more weight to LTV.  

This prioritizes leads with the **highest potential revenue impact**, balancing both likelihood to buy and historical spending.

### Trend Analysis
Before scoring, the dataset is analyzed to:
- Understand purchase frequency distributions.
- Normalize spending behavior.
- Identify outliers or anomalies.

This ensures that the scoring model is **data-driven, unbiased, and robust**.

### Follow-Up Scheduling
- **High-score leads (≥ 0.8):** Frequent promotions, multiple touchpoints, education + feedback campaigns.  
- **Medium-score leads:** Fewer, spaced-out follow-ups.  
- **Low-score leads:** Minimal contact to reduce marketing spend.  

Follow-up dates are dynamically generated based on lead quality, ensuring efficient resource allocation and minimizing email fatigue.

### A/B Testing
To improve engagement:
- Send different versions of emails to leads in the same group.
- Track open rates and click-through rates.
- Use results to refine subject lines, content, and scheduling.
- Avoid email fatigue by optimizing the frequency of messaging.

### Reporting
- Pie chart of leads by score range.
- Bar chart of average Lead Score by lead source.
- Tables for average revenue per customer by industry and city.
- Automatically generated PDF report summarizing all metrics.

