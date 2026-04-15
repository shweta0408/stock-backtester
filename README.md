# 📊 NSE Stock Backtesting System

A powerful and easy-to-use stock backtesting tool built with **Streamlit** that allows users to upload screener results and evaluate stock performance using real market data.

---

## 🚀 Live Demo

👉 Try it here: https://your-app-link.streamlit.app

⏳ Note: App may take ~15–20 seconds to wake up if idle.

## 🚀 Features

* 📂 Upload Excel/CSV screener files
* 📈 Fetch historical stock data (Yahoo Finance)
* 📊 Calculate:

  * Entry Price
  * RSI (14)
  * 1W / 2W / 3W Returns
  * Delivery %
* 🎯 Signal classification:

  * 🟢 Momentum
  * 🟡 Neutral
  * 🔴 Weak
* 🏆 Win/Loss analysis (based on 1-week return)
* 📥 Download formatted Excel report
* 🌐 Deployed on Streamlit Cloud (accessible via URL)

---

## 🧾 Input File Format

Your Excel/CSV file must contain:

| Column        | Example       |
| ------------- | ------------- |
| Symbol        | RELIANCE, TCS |
| Sector        | Technology    |
| Market Cap    | Largecap      |
| Date of Entry | 04-09-2025    |

⚠️ **Important:**
Dates should be in **DD-MM-YYYY** or similar format.
The system uses `dayfirst=True` to correctly interpret dates like `4.9.25` as **4 September 2025**.

---

## 🧠 How It Works

1. Upload your screener file
2. System fetches historical stock data
3. Adjusts entry date to nearest trading day (if market closed)
4. Calculates RSI and forward returns
5. Classifies signal and win/loss
6. Displays results + allows Excel download

---

## ⚙️ Tech Stack

* Python 🐍
* Streamlit 🎨
* Pandas & NumPy 📊
* Yahoo Finance API 📡
* OpenPyXL 📄

---

## 🛠️ Installation (Local)

```bash
pip install -r requirements.txt
streamlit run app.py
```

---

## 🌐 Deployment

This app can be deployed easily using:

👉 Streamlit Community Cloud

Steps:

1. Push code to GitHub
2. Connect repo to Streamlit Cloud
3. Deploy `app.py`

---

## 📁 Project Structure

```
stock-backtester/
 ├── app.py          # Streamlit UI
 ├── backtest.py     # Core logic
 ├── requirements.txt
 └── README.md
```

---

## 📌 Notes

* If a date falls on a weekend/holiday, it is automatically adjusted to the **nearest previous trading day**
* This ensures realistic backtesting

---

## 💡 Future Improvements

* 📊 Interactive charts (price + RSI)
* 🔐 User authentication
* ☁️ Save backtest history
* 🎨 Enhanced UI/UX
* 📅 NSE holiday calendar integration

---

## 🤝 Contributing

Feel free to fork the repo and improve the project!

---

## 📧 Contact

For queries or collaboration, feel free to reach out.
> [Connect on LinkedIn](https://www.linkedin.com/in/shweta-mishra-4777681a4)
---

## ⭐ If you like this project

Give it a ⭐ on GitHub — it helps a lot!
BUILT BY SHWETA MISHRA 
