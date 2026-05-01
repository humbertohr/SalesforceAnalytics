# Salesforce Analytics Dashboard

Interactive CRM analytics dashboard built with Streamlit and Plotly, replicating core Salesforce reporting functionality. Data is loaded from a local Excel export instead of the Salesforce API, making it fully portable and easy to run locally or deploy to the cloud.

---

## Live Demo

[View the app on Streamlit Community Cloud](https://your-app-url.streamlit.app)

---

## Screenshots

<!-- Add screenshots here -->

---

## Features

- **Overview** — Portfolio KPIs, pipeline stage distribution, and revenue by industry at a glance
- **Monthly Closed Won Revenue** — Area chart with peak month annotation and month-over-month trend
- **Top 10 Deals** — Largest opportunities ranked by deal size with stage color coding
- **Revenue vs Employees** — Bubble scatter plot with outlier isolation tab
- **Account Rating Analysis** — Dodged bar chart by industry and rating tier, revenue concentration donut, and risk insights
- **Contacts per Account** — Relationship coverage risk scoring with CRM rating segmentation
- **Open Pipeline by Account** — Stacked bar chart of active opportunities by stage
- **Revenue vs Pipeline Quadrant** — Four-quadrant scatter analysis for account prioritization

---

## Tech Stack

| Library | Purpose |
|---|---|
| `streamlit` | Web app framework |
| `plotly` | Interactive charts |
| `pandas` | Data manipulation |
| `numpy` | Numerical operations |
| `openpyxl` | Excel file reading |

---

## Project Structure

    ├── app2.py               # Main Streamlit application
    ├── salesforce.xlsx       # Excel data export (8 sheets)
    ├── salesforce.png        # Salesforce logo for sidebar
    ├── requirements.txt      # Python dependencies
    └── README.md

---

## Excel File Structure

The app reads from a local `salesforce.xlsx` file with the following sheets:

| Sheet | Key Columns Used |
|---|---|
| `accounts` | Name, Industry, AnnualRevenue, NumberOfEmployees, Rating |
| `contacts` | FirstName, LastName, AccountId |
| `opportunities` | StageName, Amount, CloseDate, AccountId |
| `leads` | — |
| `cases` | — |
| `tasks` | — |
| `products` | — |
| `pricebook_entries` | — |

---

## Getting Started

**1. Clone the repository**

    git clone https://github.com/your-username/your-repo-name.git
    cd your-repo-name

**2. Install dependencies**

    pip install -r requirements.txt

**3. Run the app**

    streamlit run app2.py

The app will open automatically at `http://localhost:8501`

---

## Deployment

This app is deployed on [Streamlit Community Cloud](https://share.streamlit.io). To deploy your own instance:

1. Push the repository to GitHub
2. Go to [share.streamlit.io](https://share.streamlit.io) and sign in with GitHub
3. Select the repository, set the main file to `app2.py`, and click Deploy

---

## Author

**Humberto** — Freelance Financial Analyst & R Developer  
[Upwork Profile](https://www.upwork.com) · [GitHub](https://github.com/your-username)

---

## License

This project is intended for portfolio demonstration purposes.
