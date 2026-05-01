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

```bash
git clone https://github.com/your-username/your-repo-name.git
cd your-repo-name
```

**2. Install dependencies**

```bash
pip install -r requirements.txt
```

**3. Run the app**

```bash
streamlit run app2.py
```

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
