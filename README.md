# 📦 Smart Supply Chain Analytics Dashboard

![Python](https://img.shields.io/badge/Python-3.10+-3776AB?style=for-the-badge&logo=python&logoColor=white)
![Streamlit](https://img.shields.io/badge/Streamlit-FF4B4B?style=for-the-badge&logo=streamlit&logoColor=white)
![Plotly](https://img.shields.io/badge/Plotly-3F4F75?style=for-the-badge&logo=plotly&logoColor=white)
![Pandas](https://img.shields.io/badge/Pandas-150458?style=for-the-badge&logo=pandas&logoColor=white)
![Status](https://img.shields.io/badge/Status-Completed-00C853?style=for-the-badge)

> An end-to-end Business Intelligence project analyzing **180,000+ supply chain orders** using the DataCo Smart Supply Chain dataset. Covers the full data lifecycle — from raw ingestion to an interactive dark-themed dashboard — delivering actionable insights on fulfillment, fraud, customer behavior, and shipping performance.

---

---

## 📁 Project Structure

```
smart-supply-chain-dashboard/
│
├── 1_Raw_Data/                   # Original DataCo Excel & CSV files
│   └── DataCo_StarSchema.xlsx
│
├── 2_Cleaned_Data/               # Processed & merged datasets
│   └── dataco_cleaned.csv
│
├── 3_Python_EDA/                 # Exploratory Data Analysis notebooks
│   ├── EDA.ipynb
│
├── 4_Data_Modeling/
│   ├── DataCo_StarSchema.xlsx
│
├── 5_Dashboard/                  # Streamlit app
│   ├── supply_chain_dashboard.py
│   ├── requirements.txt
│
└── README.md
```

---

## 🎯 Business Objectives

| # | Objective |
|---|---|
| 1 | Identify the key drivers of **late deliveries** and shipping inefficiencies |
| 2 | Detect and quantify **fraudulent orders** across segments and regions |
| 3 | Analyze **customer profitability** by segment, city, and country |
| 4 | Uncover **discount patterns** and their impact on profit margins |
| 5 | Build a fully interactive dashboard for real-time supply chain monitoring |

---

## 📊 Dashboard Pages

### Page 1 — Order & Fulfillment Performance
- KPIs: Total Orders · Total Sales · Fraud Rate % · On-Time Delivery %
- Visuals: Orders & Sales over time · Revenue by Order Status · Shipping Mode Donut · Delivery Status breakdown

### Page 2 — Customer & Market Insights
- KPIs: Total Customers · Top Segment · Avg Sales per Customer
- Visuals: Top 10 Customers · Sales by Segment · Country bar chart · City Sales Treemap

### Page 3 — Product & Profitability
- KPIs: Units Sold · Top Product · Total Profit · Avg Profit per Order
- Visuals: Quarterly Profit Trend · Top 5 Products · Department Profit · Discount Sunburst (Segment → Shipping → Delivery)

### Page 4 — Shipping & Delivery Performance
- KPIs: Avg Actual Ship Days · Avg Scheduled Days · Fraud Sales $ · Fraud Orders Count
- Visuals: Grouped Bar (Actual vs Scheduled) · On-Time Gauge · Delivery Heatmap · Fraud by Segment

### Page 5 — Time Intelligence & Trends
- Visuals: Rolling 3-Month Orders · Rolling 12-Month Revenue · YoY Sales vs Profit · Dept Profit Trend · Fraud Trend by Segment

---

## 🔑 Key Insights

- 🚨 **~19% of all orders** are flagged as Suspected Fraud, concentrated in the Consumer segment
- 📦 **Standard Class** shipping has the largest gap between actual and scheduled delivery days
- 💰 **Fitness & Technology** departments drive the highest profit margins
- 🌎 Orders are heavily concentrated in **Puerto Rico and the US**, with Puerto Rico showing the highest volume per city
- 📉 **Late delivery risk** correlates with higher discount rates, suggesting fulfillment pressure on discounted orders

---

## 🗄️ Data Model (Star Schema)

```
                    ┌─────────────┐
                    │  Dim_Date   │
                    │  YearMonth  │
                    │  Month/Year │
                    └──────┬──────┘
                           │
┌──────────────┐    ┌──────┴───────┐    ┌──────────────┐
│ Dim_Customer │────│  Fact_Orders │────│  Dim_Product │
│ Customer Id  │    │  Order Id    │    │ Product Card │
│ Segment      │    │  Sales       │    │ Product Name │
│ City/Country │    │  Profit      │    │ Category     │
└──────────────┘    │  Discount    │    └──────────────┘
                    │  Ship Days   │
┌──────────────┐    │  Fraud Flag  │    ┌──────────────┐
│Dim_Shipping  │────│              │────│ Dim_Dept     │
│ Shipping Mode│    └──────────────┘    │ Department   │
│Delivery Stat │                        │ Name         │
└──────────────┘                        └──────────────┘
```

---

## ⚙️ Tech Stack

| Layer | Tools |
|---|---|
| Data Storage | Excel (Star Schema), CSV |
| Data Cleaning | Python · Pandas · NumPy |
| EDA | Jupyter Notebook · Matplotlib · Seaborn |
| BI Modeling | DAX (Power BI measures) |
| Dashboard | Streamlit · Plotly |
| Version Control | Git · GitHub |

---

## 🚀 Getting Started

### 1. Clone the repository
```bash
git clone https://github.com/Mohammed-Dahi/smart-supply-chain-dashboard
cd smart-supply-chain-dashboard
```

### 2. Install dependencies
```bash
pip install -r 5_Dashboard/requirements.txt
```

### 3. Run the dashboard
```bash
streamlit run 5_Dashboard/supply_chain_dashboard.py
```

### 4. Upload your data
When the app opens in your browser, upload `DataCo_StarSchema.xlsx` from the sidebar. All 6 sheets are loaded and joined automatically.

---

## 📦 Requirements

```
streamlit>=1.32.0
plotly>=5.18.0
pandas>=2.1.0
openpyxl>=3.1.2
numpy>=1.26.0
```

---

## 📂 Dataset

**DataCo Smart Supply Chain for Big Data Analysis**  
- Source: [Kaggle — DataCo Smart Supply Chain](https://www.kaggle.com/datasets/shashwatwork/dataco-smart-supply-chain-for-big-data-analysis)
- Records: ~180,000 orders
- Period: 2015 – 2018
- Tables: 6 (Star Schema)

> ⚠️ The raw dataset is not included in this repository due to size. Download it from Kaggle and place it in the `1_Raw_Data/` folder.

---

## 👤 Author

**Mohammed Dahi**  
Data Analyst | BI Developer  
[![LinkedIn](https://img.shields.io/badge/LinkedIn-0A66C2?style=flat&logo=linkedin&logoColor=white)](https://www.linkedin.com/in/mohammeddahi/)
[![GitHub](https://img.shields.io/badge/GitHub-181717?style=flat&logo=github&logoColor=white)](https://github.com/Mohammed-Dahi)

---

## 📄 License

This project is licensed under the **MIT License** — see the [LICENSE](LICENSE) file for details.

---

<p align="center">
  Built with ❤️ using Python · Streamlit · Plotly
</p>
```
