# OraGlBalEx Excel-DNA Add-In

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

## 🚀 Features
- **OraGlBalEx** is a Excel-DNA add-in developed in F# that seamlessly integrates Oracle [Fusion General Ledger (GL) API Balance](https://docs.oracle.com/en/cloud/saas/financials/24d/farfa/op-ledgerbalances-get.html) data directly into Microsoft Excel. By eliminating intermediate steps between Oracle Fusion and Excel, OraGlBalEx streamlines the data retrieval process, enabling to access and analyze GL balances efficiently.
Utilizing a single, intuitive Excel function, OraGlBalEx fetches data from the Oracle GL API and enhances it by parsing separation key flexfields into distinct columns. This automation fit for various ad-hoc reporting scenarios, including the generation of trial balances and other financial statements.
Below is the possible use case scenario.
![Watch the Demonstration Video](gifs/glbalexdemo.gif)
---

## 🛠 Prerequisites

Before installing OraGlBalEx, ensure you have the following prerequisites:

- **Microsoft Excel:** Version compatible with Excel-DNA add-ins (Excel 2007 and later).
- **.NET 6.0:** Ensure that the [.NET 6.0 runtime](https://dotnet.microsoft.com/en-us/download/dotnet/6.0) is installed on your machine.
- **Oracle Fusion REST API Access:** Valid credentials and necessary permissions to access Oracle GL Balances APIs (only basic authorization here).
- **PowerShell (Optional):** For managing credentials and environment setup.

---

## 📝 Installation

### 1. **Clone the Repository**

Clone the OraGlBalEx repository to your local machine:

git clone https://github.com/yourusername/OraGlBalEx.git



## 📚 Additional Resources
- Excel-DNA Documentation:
[Excel-DNA Official Documentation](https://excel-dna.net/)




