# CSV Convertor – Shopify Order Processing Tool

This Java desktop application was built to streamline order data reconciliation for a Shopify-based business. It automates the previously manual and error-prone process of matching purchase and shipment records, helping reduce human error and prevent costly duplicate shipments.

---

## 🔧 Features

- ✅ **Automatic CSV parsing and transformation**  
  Converts raw Shopify export files (order and shipment data) into reconciled, error-free final reports.

- ✅ **Rule-based validation engine**  
  Built-in logic to match and verify item quantity per order ID, flag mismatches, and resolve inconsistencies.

- ✅ **Persistent path preferences**  
  Users can save and reuse frequently accessed folder paths via a simple local configuration system.

- ✅ **One-click batch processing**  
  Supports processing hundreds of orders in seconds via intuitive UI.

---

## 💼 Business Impact

Before this tool was implemented:

- Manual cross-checking across multiple CSV files
- Frequent human errors (e.g., duplicate shipments)
- At least one full-time admin required for daily tasks
- Revenue loss due to unrecoverable shipping mistakes

After implementing this app:

- Processing became **fully automated** and **zero-error**
- Hours of work reduced to **under one minute**
- No further incidents of incorrect shipping

---

## 🖥️ Technologies Used

- Java (Swing GUI)
- OpenCSV (CSV parsing and writing)
- Custom logic modules for validation and transformation
- Local file preference storage

---

## 📂 Project Structure

- `Main.java` – Entry point of the application
- `ConverterGUI.java` – GUI logic and event handling
- `ShopifyTracker.java` – Manages file parsing and data flow
- `TextProcessor.java` – Handles text parsing and cleaning
- `Orders.java` – Internal order data model
- `PathPreference.java` – Stores/retrieves path history

---

## 🚀 How to Run

1. Clone this repository:
   ```bash
   git clone https://github.com/chengcindyy/csv-convertor.git
   ```
   Open the project in IntelliJ, Eclipse, or any Java IDE.

2. Run Main.java to start the application.
3. Select input folders and click the button to process files.

## 🧠 Learning & Context
This project originated from solving a real-world business challenge at a small e-commerce company. It refactors and builds on logic from a prior document-scanning project (Image Convertor) to optimize order processing through reliable automation.
