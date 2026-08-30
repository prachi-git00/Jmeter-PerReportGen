# 📊 JMeter Performance Report Generator

Transforms JMeter `.jtl` result files into professional performance reports with response time analytics, throughput, pass/fail ratios, and SLA tracking.

## ✨ Features

- Response time analytics (Min, Max, Avg, 90th/95th/99th percentiles)
- Pass/Fail analysis with error rates
- Throughput calculations (Mbps)
- Common page handling (login, setup) with overall duration stats
- Response time categorization (1-3s, 3-5s, 5-10s ranges)
- Infrastructure monitoring (DB, Pod utilization)
- Project management with pretest changes tracking
- Professional Excel reports with formatting

## 🚀 Quick Start

### Installation

```bash
cd Jmeter-PerformanceReportGen
python -m venv .venv
.venv\Scripts\Activate.ps1           # Windows
# or: source .venv/bin/activate      # Linux/Mac
pip install -r requirements.txt
streamlit run app.py
```
Open browser at `http://localhost:8501`

## 📖 Usage Steps

1. **Create Project** - Set up a new test project
2. **Download Template** - Click "Download Sample Template" button
3. **Fill Template** - Add transaction mappings (Type, Module, Scenario, Page, Yes/No, Target TPH)
    Note: Jtl transction and template transaction/pages name must match.
          Yes/No column is for the Transaction to be included in passFail sheet.
4. **Upload Files** - Add JTL files and transaction template
5. **Set Time Window** - Select peak test period (From/To dates)
6. **Generate Report** - Click "Generate Report" and download Excel file

## 📁 Key Files

- `app.py` - Main Streamlit UI
- `Cal_TPH_RespTime_jtl.py` - Response time & throughput calculations
- `ExeSummary.py` - Executive summary sheet
- `db_manager.py` - Project management

## 📊 Response Time Behavior

| Type | Time Window | Success Only |
|------|-------------|--------------|
| Regular Transactions | Peak window (From/To dates) | ✓ Yes |
| Common Pages* | Full JTL duration | ✓ Yes |
| PassFail Analysis | Peak window | All samples |

*Mark common pages by adding "Common" in Module or Scenario column

## 💾 Template Columns

| Type | Module | Scenario | Page | Yes/No | Target TPH |
|------|--------|----------|------|--------|------------|
| UI | Common | Setup | Login | yes | |
| UI | Orders | Checkout | Place Order | yes | 500 |
| API | API | Orders | Place Order | yes | 5000 |

## ⚙️ Configuration

- **From/To Date/Time**: Peak test window for regular transactions
- **Common Pages**: Detected by "common" keyword in Module/Scenario - appear first in report, excluded from PassFail
- **Template Match**: JTL labels must match "Page" column (case-insensitive)

## 📝 Requirements

- Python 3.8+
- streamlit, pandas, openpyxl, plotly

## 🔄 Example Workflow

1. Create project "Load Test v1"
2. Download & populate template with your transactions
3. Upload JTL file and template
4. Set analysis window: 10:00 AM - 10:30 AM
5. Click "Generate Report"
6. Download Excel with all metrics & SLA status

---

**Version**: 1.0  |  **Last Updated**: August 2026
