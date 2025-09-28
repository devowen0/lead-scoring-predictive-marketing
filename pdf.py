import pandas as pd
import matplotlib.pyplot as plt
from fpdf import FPDF, XPos, YPos
from datetime import datetime

# -----------------------------
# Step 1: Read Excel file
# -----------------------------
file_name = "demo_leads_scored.xlsx"
df = pd.read_excel(file_name)

# -----------------------------
# Step 2: Filter Lead Score between 0 and 1
# -----------------------------
df = df[(df['Lead Score'] >= 0) & (df['Lead Score'] <= 1)]

# -----------------------------
# Step 3: Define bins for each tenth
# -----------------------------
bins = [i/10 for i in range(11)]
labels = [f"{bins[i]:.1f}-{bins[i+1]:.1f}" for i in range(len(bins)-1)]

# -----------------------------
# Step 4: Calculate percentages
# -----------------------------
df['Score Range'] = pd.cut(df['Lead Score'], bins=bins, labels=labels, include_lowest=True)
percentages = df['Score Range'].value_counts(normalize=True).sort_index() * 100

# -----------------------------
# Step 5: Create pie chart of score ranges
# -----------------------------
colors = ['#4E79A7', '#F28E2B', '#E15759', '#76B7B2', '#59A14F',
          '#EDC948', '#B07AA1', '#FF9DA7', '#9C755F', '#BAB0AC']

plt.figure(figsize=(6, 6))
plt.pie(
    percentages,
    labels=labels,
    autopct='%1.1f%%',
    startangle=90,
    colors=colors,
    wedgeprops={'edgecolor': 'white'}
)
plt.title("Percentage of leads by score range")
score_chart_file = "lead_score_pie_chart.png"
plt.savefig(score_chart_file, bbox_inches='tight')
plt.close()

# -----------------------------
# Step 6: Calculate averages
# -----------------------------
industry_avg = df.groupby('Industry')['Lead Score'].mean().sort_values(ascending=False) if 'Industry' in df.columns else pd.Series(dtype=float)
city_avg = df.groupby('City')['Lead Score'].mean().sort_values(ascending=False) if 'City' in df.columns else pd.Series(dtype=float)
source_avg = df.groupby('Lead Source')['Lead Score'].mean().sort_values(ascending=False) if 'Lead Source' in df.columns else pd.Series(dtype=float)

# -----------------------------
# Step 7: Calculate Revenue
# -----------------------------
if 'Previous Purchases' in df.columns and 'Average Purchase Value (SEK)' in df.columns:
    df['Revenue'] = df['Previous Purchases'] * df['Average Purchase Value (SEK)']
    revenue_by_industry = df.groupby('Industry')['Revenue'].mean().sort_values(ascending=False) if 'Industry' in df.columns else pd.Series(dtype=float)
    revenue_by_city = df.groupby('City')['Revenue'].mean().sort_values(ascending=False) if 'City' in df.columns else pd.Series(dtype=float)
    revenue_by_source = df.groupby('Lead Source')['Revenue'].mean().sort_values(ascending=False) if 'Lead Source' in df.columns else pd.Series(dtype=float)
else:
    revenue_by_industry = pd.Series(dtype=float)
    revenue_by_city = pd.Series(dtype=float)
    revenue_by_source = pd.Series(dtype=float)

# -----------------------------
# Step 8: Create charts for Lead Source
# -----------------------------
if not source_avg.empty:
    plt.figure(figsize=(6, 4))
    source_avg.plot(kind='bar', color='#4E79A7')
    plt.ylabel("Average Lead Score")
    plt.title("Average Lead Score by Lead Source")
    plt.xticks(rotation=45, ha='right')
    plt.tight_layout()
    leadscore_chart_file = "leadscore_by_source_bar.png"
    plt.savefig(leadscore_chart_file)
    plt.close()

if not revenue_by_source.empty:
    plt.figure(figsize=(6, 6))
    plt.pie(
        revenue_by_source,
        labels=revenue_by_source.index,
        autopct='%1.1f%%',
        startangle=90,
        wedgeprops={'edgecolor': 'white'}
    )
    plt.title("Average Revenue per Customer by Lead Source")
    revenue_chart_file = "revenue_by_source_pie.png"
    plt.savefig(revenue_chart_file, bbox_inches='tight')
    plt.close()

# -----------------------------
# Step 9: Create PDF
# -----------------------------
today = datetime.today().strftime("%d %B %Y")
pdf_file_name = f"{today} Report.pdf"

pdf = FPDF()
pdf.add_page()

# Title
pdf.set_font("Helvetica", "B", 16)
pdf.cell(0, 10, f"{today} Report", new_x=XPos.LMARGIN, new_y=YPos.NEXT, align="C")

# Pie chart: Score ranges
pdf.image(score_chart_file, x=30, w=150)

# Explanatory text (replace en dash with normal hyphen to avoid Unicode issues)
explanatory_text = (
    "Lead Score is a 0-1 metric that ranks existing customers by future revenue potential, "
    "dynamically combining their likelihood of buying again (Purchase Score) and their historical "
    "spending level (Lifetime Value) to help prioritize retention, reactivation, and upselling efforts."
)
pdf.ln(8)
pdf.set_font("Helvetica", "", 11)
pdf.multi_cell(0, 6, explanatory_text)

# Percentages
pdf.ln(4)
pdf.set_font("Helvetica", "B", 12)
pdf.cell(0, 8, "Percentage of Leads by Score Range:", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
pdf.set_font("Helvetica", "", 12)
for label, pct in percentages.items():
    pdf.cell(0, 6, f"{label}: {pct:.1f}%", new_x=XPos.LMARGIN, new_y=YPos.NEXT)

# Average Lead Score by Industry
if not industry_avg.empty:
    pdf.ln(2)
    pdf.set_font("Helvetica", "B", 12)
    pdf.cell(0, 8, "Average Lead Score by Industry:", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.set_font("Helvetica", "", 12)
    for industry, avg_score in industry_avg.items():
        pdf.cell(0, 6, f"{industry}: {avg_score:.2f}", new_x=XPos.LMARGIN, new_y=YPos.NEXT)

# Average Revenue Per Customer by Industry
if not revenue_by_industry.empty:
    pdf.ln(2)
    pdf.set_font("Helvetica", "B", 12)
    pdf.cell(0, 8, "Average Revenue Per Customer by Industry:", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.set_font("Helvetica", "", 12)
    for industry, revenue in revenue_by_industry.items():
        pdf.cell(0, 6, f"{industry}: {int(round(revenue)):,} SEK", new_x=XPos.LMARGIN, new_y=YPos.NEXT)

# Average Lead Score by City
if not city_avg.empty:
    pdf.ln(2)
    pdf.set_font("Helvetica", "B", 12)
    pdf.cell(0, 8, "Average Lead Score by City:", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.set_font("Helvetica", "", 12)
    for city, avg_score in city_avg.items():
        pdf.cell(0, 6, f"{city}: {avg_score:.2f}", new_x=XPos.LMARGIN, new_y=YPos.NEXT)

# Average Revenue Per Customer by City
if not revenue_by_city.empty:
    pdf.ln(2)
    pdf.set_font("Helvetica", "B", 12)
    pdf.cell(0, 8, "Average Revenue Per Customer by City:", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.set_font("Helvetica", "", 12)
    for city, revenue in revenue_by_city.items():
        pdf.cell(0, 6, f"{city}: {int(round(revenue)):,} SEK", new_x=XPos.LMARGIN, new_y=YPos.NEXT)

# Bar chart: Average Lead Score by Lead Source
if not source_avg.empty:
    pdf.ln(2)
    pdf.image(leadscore_chart_file, x=25, w=160)

# Average Lead Score by Lead Source
if not source_avg.empty:
    pdf.ln(2)
    pdf.set_font("Helvetica", "B", 12)
    pdf.cell(0, 8, "Average Lead Score by Lead Source:", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.set_font("Helvetica", "", 12)
    for source, avg_score in source_avg.items():
        pdf.cell(0, 6, f"{source}: {avg_score:.2f}", new_x=XPos.LMARGIN, new_y=YPos.NEXT)

# Pie chart: Average Revenue per Customer by Lead Source
if not revenue_by_source.empty:
    pdf.ln(2)
    pdf.image(revenue_chart_file, x=30, w=150)

# Average Revenue Per Customer by Lead Source
if not revenue_by_source.empty:
    pdf.ln(2)
    pdf.set_font("Helvetica", "B", 12)
    pdf.cell(0, 8, "Average Revenue Per Customer by Lead Source:", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.set_font("Helvetica", "", 12)
    for source, revenue in revenue_by_source.items():
        pdf.cell(0, 6, f"{source}: {int(round(revenue)):,} SEK", new_x=XPos.LMARGIN, new_y=YPos.NEXT)

# Save PDF
pdf.output(pdf_file_name)
print(f"Report saved as '{pdf_file_name}'")
