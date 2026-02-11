# 1. Setup Environment
!pip install pandas matplotlib seaborn openpyxl python-docx -q

import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np
import os
import io
import base64
from google.colab import files
from docx import Document

def run_master_dashboard_suite():
    print("💎 Initializing Master Intelligence Suite...")
    
    # --- STEP 1: FILE LOADING ---
    uploaded = files.upload()
    if not uploaded: return
    fname = list(uploaded.keys())[0]
    df = pd.read_csv(io.BytesIO(uploaded[fname])) if fname.endswith('.csv') else pd.read_excel(io.BytesIO(uploaded[fname]))

    # --- STEP 2: SMART DETECTION & CALCULATIONS ---
    text_cols = df.select_dtypes(include=['object']).columns.tolist()
    name_col = text_cols[0] if text_cols else "Student"
    num_cols = df.select_dtypes(include=[np.number]).columns.tolist()
    att_col = next((c for c in num_cols if 'attendance' in c.lower()), None)
    subjects = [c for c in num_cols if c != att_col and 'percentage' not in c.lower() and 'gpa' not in c.lower()]

    df['Percentage'] = (df[subjects].sum(axis=1) / (len(subjects) * 100)) * 100
    df['GPA'] = df['Percentage'].apply(lambda p: round(min(4.0, (p/25)), 2))
    df['Status'] = df.apply(lambda r: '🔴 CRITICAL' if (r['Percentage'] < 50 or (r[att_col] < 75 if att_col else False)) else '🟢 STABLE', axis=1)

    # --- STEP 3: TREND VISUALIZATION (Encoded for HTML) ---
    def fig_to_base64(fig):
        img = io.BytesIO()
        fig.savefig(img, format='png', bbox_inches='tight')
        img.seek(0)
        return base64.b64encode(img.getvalue()).decode()

    # Create Trend Charts
    plt.style.use('ggplot')
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(16, 6))
    
    # Chart 1: Subject Trends
    df[subjects].mean().sort_values().plot(kind='barh', color='teal', ax=ax1)
    ax1.set_title("Class Average per Subject")
    
    # Chart 2: Attendance vs Performance
    if att_col:
        sns.regplot(x=att_col, y='Percentage', data=df, ax=ax2, color='orange')
        ax2.set_title("Correlation: Attendance vs Performance")

    charts_base64 = fig_to_base64(fig)
    plt.close()

    # --- STEP 4: SUBJECT LEADERS ---
    sub_experts = {sub: df.loc[df[sub].idxmax(), name_col] for sub in subjects}

    # --- STEP 5: GENERATE HTML MASTER DASHBOARD ---
    html_content = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <title>Master Analytics Dashboard</title>
        <style>
            body {{ font-family: 'Segoe UI', Arial; margin: 0; background: #f0f2f5; color: #333; }}
            .header {{ background: #2c3e50; color: white; padding: 20px; text-align: center; }}
            .container {{ padding: 30px; max-width: 1200px; margin: auto; }}
            .grid {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(250px, 1fr)); gap: 20px; margin-bottom: 30px; }}
            .card {{ background: white; padding: 20px; border-radius: 12px; box-shadow: 0 4px 15px rgba(0,0,0,0.05); }}
            .stat {{ font-size: 2em; font-weight: bold; color: #3498db; }}
            .trend-img {{ width: 100%; border-radius: 8px; margin-top: 20px; }}
            table {{ width: 100%; border-collapse: collapse; margin-top: 20px; background: white; }}
            th, td {{ padding: 15px; text-align: left; border-bottom: 1px solid #eee; }}
            th {{ background: #3498db; color: white; }}
            .expert-tag {{ background: #f1c40f; padding: 5px 10px; border-radius: 20px; font-size: 0.8em; font-weight: bold; }}
        </style>
    </head>
    <body>
        <div class="header"><h1>🏫 Master Educational Analytics Suite</h1></div>
        <div class="container">
            <div class="grid">
                <div class="card"><h3>Total Students</h3><div class="stat">{len(df)}</div></div>
                <div class="card"><h3>Class Average</h3><div class="stat">{df['Percentage'].mean():.1f}%</div></div>
                <div class="card"><h3>Average GPA</h3><div class="stat">{df['GPA'].mean():.2f}</div></div>
                <div class="card"><h3>Risk Cases</h3><div class="stat" style="color: #e74c3c;">{len(df[df['Status'] == '🔴 CRITICAL'])}</div></div>
            </div>

            <div class="card">
                <h2>📈 Performance Trends & Correlations</h2>
                <img src="data:image/png;base64,{charts_base64}" class="trend-img">
            </div>

            <div class="card" style="margin-top: 20px;">
                <h2>🏆 Subject Experts (Peer Tutors)</h2>
                <div style="display: flex; flex-wrap: wrap; gap: 10px;">
                    {"".join([f"<div class='card' style='flex:1; border: 1px solid #eee;'><b>{sub}</b><br>{leader} <span class='expert-tag'>TOP</span></div>" for sub, leader in sub_experts.items()])}
                </div>
            </div>

            <div class="card" style="margin-top: 20px;">
                <h2>📋 Detailed Student Audit</h2>
                {df[[name_col, 'Percentage', 'GPA', 'Status']].sort_values(by='Percentage', ascending=False).to_html(classes='table', index=False)}
            </div>
        </div>
    </body>
    </html>
    """

    with open("Master_Report_2026.html", "w", encoding='utf-8') as f:
        f.write(html_content)

    # --- STEP 6: EXPORTS ---
    df.to_excel("Comprehensive_Data_Analysis.xlsx", index=False)
    print("\n✅ MASTER SUITE COMPLETE")
    files.download("Master_Report_2026.html")
    files.download("Comprehensive_Data_Analysis.xlsx")

run_master_dashboard_suite()