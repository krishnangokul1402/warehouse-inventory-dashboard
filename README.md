# Warehouse Inventory Tracking Agent — Python + Excel

This starter project runs a warehouse-level control tower with an Excel workbook and a Streamlit chat-style agent. **Sample data are fictional.**

## What the agent tracks

- **Goods at warehouse**: physically received on-hand units; available = on-hand − reserved.
- **Goods at plant**: ready but unshipped, allocated to a destination warehouse.
- **Goods in transit (GIT)**: pending = shipped − received for shipments marked `In Transit`.
- **ETA and delays**: a shipment is late if ETA is before today's date and its pending quantity is positive.
- **7-day projection**: available now + in-transit quantities with ETA between today and 7 days ahead. This intentionally excludes plant goods and future consumption.

Do not add plant stock or in-transit goods to on-hand warehouse stock until receipt is recorded. Upon dispatch reduce plant allocation and add a shipment; upon receipt update `Qty_Received` *and* increase the warehouse on-hand record. Set status `Delivered` when fully reconciled.

## Run on Windows (PowerShell)

1. Unzip the project and open the folder in PyCharm or VS Code.
2. Install **Python 3.10+**. Open a terminal in the project folder.
3. (Recommended) Create a virtual environment:

   ```powershell
   py -m venv .venv
   .\.venv\Scripts\Activate.ps1
   ```

4. Install dependencies:

   ```powershell
   python -m pip install -r requirements.txt
   ```

5. Start the app:

   ```powershell
   python -m streamlit run app.py
   ```

6. The terminal displays a local URL (usually `http://localhost:8501`). Open it in a browser.
7. Change records in **Warehouse_Inventory_Sample.xlsx** (preserving sheet/column names), then upload your saved file via the sidebar.
8. Choose warehouse and SKU filters or ask a natural-language question in the **Ask the agent** tab.

### Sample questions

- How much stock is available in Chennai?
- Show stock status for SKU-102 in Bengaluru.
- Which shipments are delayed?
- What is the ETA for SKU-101?
- How many units are arriving in the next 7 days?
- How much stock is still at plant for Hyderabad?
- Which products are below reorder level?

## Optional OpenAI explanation

You can run all calculations and basic question answering **without** an API key. To enable optional AI rewording, create an OpenAI API key in your own account, set it as a local environment variable, and enable the toggle in the sidebar:

```powershell
$env:OPENAI_API_KEY = "your-own-key"
python -m streamlit run app.py
```

**Privacy note:** Enabling that toggle sends your typed question and the calculated answer to the OpenAI API. Keep it off for confidential company stock data unless your organization permits sharing it. Do not put the API key inside the Excel workbook or commit it to source control.

## Required worksheets and core columns

| Worksheet | Fields |
| --- | --- |
| Warehouses | Warehouse_ID, Warehouse_Name, City, Region, Active |
| Products | SKU, Product_Name, Category, Reorder_Level |
| Warehouse_Stock | Warehouse_ID, SKU, On_Hand, Reserved, Last_Updated |
| Plant_Stock | Plant_ID, Plant_Name, Destination_Warehouse_ID, SKU, Qty_Ready, Planned_Dispatch_Date, Last_Updated |
| Shipments | Shipment_ID, Origin_Plant, Destination_Warehouse_ID, SKU, Qty_Shipped, Qty_Received, Dispatch_Date, ETA, Status, Last_Updated |

Use dates such as `2026-11-03` (Excel dates are best). All quantities must be nonnegative. The same warehouse IDs and SKUs must exist in their master sheets.

## Important limitations for production

- ETA is supplied by your workbook, not predicted from GPS. For real-time updates, connect an ERP/WMS/transport management data feed.
- This is a **read-only dashboard**; no automated inventory posting or movement reconciliation is performed.
- The 7-day figure is an inbound-only projection. Add demand, outbound orders, and purchase orders for a complete available-to-promise (ATP) forecast.
- Updates use your latest uploaded Excel snapshot. A production app should implement access control, audit trails, schedule/refresh, and data validation before operational use.
- Duplicate same-SKU stock rows for a warehouse are grouped for reporting; shipment IDs must be unique.

## Files

- `Warehouse_Inventory_Sample.xlsx`: editable workbook with example data and an Excel dashboard.
- `inventory_agent.py`: Excel validation, quantities, delay logic, local natural-language query handling.
- `app.py`: Streamlit control tower and chat interface.
- `requirements.txt`: dependencies.
- `test_agent.py`: automated calculation checks using demo data.
