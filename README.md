# BRIDeal

**BRIDeal** is a modular desktop application designed for agricultural equipment dealerships (e.g., John Deere) to manage, track, and streamline equipment sales deals. It integrates sales tools, inventory management, pricing engines, and external APIs into a single PyQt5-based interface tailored for sales teams and equipment managers.

---

## 🚜 Key Features

- **🔧 Deal Management**  
  Create and manage complex equipment deals involving new/used machinery, trades, parts, and pricing breakdowns. Automatically generate CSV reports and email notifications.

- **📊 Home Dashboard**  
  View live data including:
  - Weather conditions
  - Exchange rates (CAD/USD)
  - Commodity prices (Wheat, Canola, Bitcoin)

- **📁 Recent Deals**  
  Browse previously created deals, reload forms, or regenerate data exports.

- **📘 Price Book**  
  Access up-to-date equipment pricing from SharePoint with margin/markup calculations using real-time FX rates.

- **🚜 Used Inventory**  
  Display and manage used equipment listings synced with SharePoint.

- **🧮 Calculator Tool**  
  Integrated financial calculator for margin, markup, and FX conversions.

- **🛠 CSV Editors**  
  GUI-based editors for maintaining product, parts, customers, and salesperson datasets.

- **🚚 Receiving Module**  
  Integration with external traffic systems for managing incoming stock.

- **🔗 John Deere Portal Integration**  
  Interface directly with JD Quote and JD Portal services for equipment configuration and quoting.

- **📤 SharePoint Integration**  
  Read/write operations to SharePoint-hosted Excel files, enabling centralized data sync and email delivery via Microsoft Graph.

---

## 💻 Tech Stack

- **Python 3.x**
- **PyQt5** – GUI framework
- **MSAL** – Microsoft Graph Authentication
- **Microsoft Graph API** – SharePoint and Email integration
- **REST APIs** – Live data: Weather, Exchange, Commodities
- **CSV/JSON** – Local data handling
- **Excel (via SharePoint)** – Remote structured storage

---

## 🔧 Installation

> **Note**: This app is intended to be used internally by dealership employees. Ensure you have access to your organization's SharePoint and required API keys.

```bash
# Clone the repo
git clone https://github.com/your-org/BRIDeal.git
cd BRIDeal

# Create virtual environment
python -m venv venv
source venv/bin/activate  # or venv\Scripts\activate on Windows

# Install dependencies
pip install -r requirements.txt
```

---

## 🚀 Running the App

```bash
python main.py
```

---

## 🛡 Authentication Setup

BRIDeal uses **MSAL** to authenticate via Microsoft Azure AD and access SharePoint + Graph resources.

> You’ll need:
- A registered app in Azure AD
- Client ID / Tenant ID
- Redirect URI (for desktop flows)
- API permissions: `Files.ReadWrite`, `Mail.Send`, `Sites.Read.All`

These credentials should be placed in a `config.json` or securely injected via environment variables.

---

## 🧪 Development Notes

- **Modules** are loaded dynamically via stacked widgets
- **Signal-slot architecture** is used for inter-module communication
- Code follows **PEP8** styling and modular architecture
- APIs are decoupled and mocked for offline development



## 🤝 Contributing

This is a private/internal-use app. Contributions are limited to approved team members.

---

## 📄 License

 `[Internal Use Only]

---

