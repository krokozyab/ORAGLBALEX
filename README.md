# OraGlBalEx Excel-DNA Add-In


**OraGlBalEx** is a powerful Excel-DNA add-in developed in F# that seamlessly integrates Oracle General Ledger (GL) Balances data directly into Microsoft Excel. Leveraging asynchronous operations, secure credentials management, and structured logging with Serilog, OraGlBalEx ensures a responsive and efficient user experience while fetching and processing financial data from Oracle Fusion REST APIs.

---

## 📄 Table of Contents

- [🚀 Features](#-features)
- [🛠 Prerequisites](#-prerequisites)
- [📝 Installation](#-installation)
- [⚙️ Configuration](#-configuration)
    - [Secrets Management](#secrets-management)
    - [Logging Configuration](#logging-configuration)
- [🖥 Usage](#-usage)
    - [StoreSecret Function](#storesecret-function)
    - [GetSecret Function](#getsecret-function)
    - [DeleteSecret Function](#deletesecret-function)
    - [WriteBalancesFieldsToExcel Function](#writebalancesfieldstoexcel-function)
- [📜 Logging](#-logging)
- [🔧 Project Structure](#-project-structure)
- [🛠 Troubleshooting](#-troubleshooting)
- [🤝 Contributing](#-contributing)
- [📝 License](#-license)
- [📫 Contact](#-contact)
- [📚 Additional Resources](#-additional-resources)

---

## 🚀 Features

- **Asynchronous Data Fetching:** Retrieve Oracle GL Balances without freezing Excel, ensuring a responsive user interface.
- **Structured Logging:** Utilize Serilog for comprehensive logging, aiding in monitoring and debugging.
- **Dynamic Data Processing:** Automatically segment and organize `DetailAccountCombination` fields into separate columns.
- **Secure Credentials Management:** Safely manage and encode credentials using Windows Credential Manager.
- **Custom JSON Converters:** Handle complex JSON deserialization scenarios with custom converters.
- **Error Handling:** Gracefully handle API errors and deserialization issues, providing informative feedback within Excel.
- **Configurable Logging:** Store logs in user-specific directories following Windows best practices.

---

## 🛠 Prerequisites

Before installing OraGlBalEx, ensure you have the following prerequisites:

- **Microsoft Excel:** Version compatible with Excel-DNA add-ins (Excel 2007 and later).
- **.NET 6.0:** Ensure that the .NET 6.0 SDK is installed on your machine.
- **Oracle Fusion REST API Access:** Valid credentials and necessary permissions to access Oracle GL Balances APIs.
- **PowerShell (Optional):** For managing credentials and environment setup.

---

## 📝 Installation

### 1. **Clone the Repository**

Clone the OraGlBalEx repository to your local machine:

```bash
git clone https://github.com/yourusername/OraGlBalEx.git
