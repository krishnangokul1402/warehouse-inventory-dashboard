"""Validated inventory calculations for the Warehouse Control Tower.

The calculations are deterministic. No LLM is allowed to invent stock or ETA values.
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date, timedelta
import re

import pandas as pd


REQUIRED = {
    "Warehouses": ["Warehouse_ID", "Warehouse_Name", "City", "Region", "Active"],
    "Products": ["SKU", "Product_Name", "Category", "Reorder_Level"],
    "Warehouse_Stock": ["Warehouse_ID", "SKU", "On_Hand", "Reserved", "Last_Updated"],
    "Plant_Stock": ["Plant_ID", "Plant_Name", "Destination_Warehouse_ID", "SKU", "Qty_Ready", "Planned_Dispatch_Date", "Last_Updated"],
    "Shipments": ["Shipment_ID", "Origin_Plant", "Destination_Warehouse_ID", "SKU", "Qty_Shipped", "Qty_Received", "Dispatch_Date", "ETA", "Status", "Last_Updated"],
}


@dataclass
class Inventory:
    warehouses: pd.DataFrame
    products: pd.DataFrame
    stock: pd.DataFrame
    plant: pd.DataFrame
    shipments: pd.DataFrame
    snapshot: date


def _validate_quantity(df: pd.DataFrame, col: str, tab: str) -> None:
    df[col] = pd.to_numeric(df[col], errors="coerce")
    if df[col].isna().any() or (df[col] < 0).any():
        raise ValueError(f"{tab}.{col} must contain non-negative numeric values on all rows.")


def load_inventory(source, as_of: date | None = None) -> Inventory:
    """Load the five input sheets from an XLSX path or a Streamlit upload."""
    sheets = pd.read_excel(source, sheet_name=list(REQUIRED), engine="openpyxl")
    for tab, fields in REQUIRED.items():
        missing = sorted(set(fields) - set(sheets[tab].columns))
        if missing:
            raise ValueError(f"{tab} is missing columns: {', '.join(missing)}")
        sheets[tab] = sheets[tab][fields].dropna(how="all").copy()
    warehouse = sheets["Warehouses"]
    products = sheets["Products"]
    stock = sheets["Warehouse_Stock"]
    plant = sheets["Plant_Stock"]
    shipments = sheets["Shipments"]

    for df, cols in [
        (warehouse, ["Warehouse_ID", "Warehouse_Name", "City"]),
        (products, ["SKU", "Product_Name"]),
        (stock, ["Warehouse_ID", "SKU"]),
        (plant, ["Plant_ID", "Destination_Warehouse_ID", "SKU"]),
        (shipments, ["Shipment_ID", "Destination_Warehouse_ID", "SKU", "Status"]),
    ]:
        for col in cols:
            df[col] = df[col].astype("string").str.strip()
    for col in ["On_Hand", "Reserved"]:
        _validate_quantity(stock, col, "Warehouse_Stock")
    _validate_quantity(plant, "Qty_Ready", "Plant_Stock")
    for col in ["Qty_Shipped", "Qty_Received"]:
        _validate_quantity(shipments, col, "Shipments")
    _validate_quantity(products, "Reorder_Level", "Products")

    if (stock["Reserved"] > stock["On_Hand"]).any():
        raise ValueError("Reserved must not exceed On_Hand in Warehouse_Stock.")
    if (shipments["Qty_Received"] > shipments["Qty_Shipped"]).any():
        raise ValueError("Qty_Received must not exceed Qty_Shipped in Shipments.")
    if warehouse["Warehouse_ID"].duplicated().any():
        raise ValueError("Duplicate Warehouse_ID in Warehouses.")
    if products["SKU"].duplicated().any():
        raise ValueError("Duplicate SKU in Products.")
    if shipments["Shipment_ID"].duplicated().any():
        raise ValueError("Duplicate Shipment_ID in Shipments.")
    wh_ids = set(warehouse["Warehouse_ID"])
    sku_ids = set(products["SKU"])
    for tab, frame, wh_col in [
        ("Warehouse_Stock", stock, "Warehouse_ID"),
        ("Plant_Stock", plant, "Destination_Warehouse_ID"),
        ("Shipments", shipments, "Destination_Warehouse_ID"),
    ]:
        if not set(frame[wh_col]).issubset(wh_ids) or not set(frame["SKU"]).issubset(sku_ids):
            raise ValueError(f"{tab} contains an unknown warehouse or SKU. Update Warehouses/Products first.")

    for df, cols in [
        (stock, ["Last_Updated"]),
        (plant, ["Planned_Dispatch_Date", "Last_Updated"]),
        (shipments, ["Dispatch_Date", "ETA", "Last_Updated"]),
    ]:
        for col in cols:
            df[col] = pd.to_datetime(df[col], errors="coerce").dt.normalize()

    shipments["Status_Normalized"] = shipments["Status"].str.lower().str.strip()
    valid_statuses = {"in transit", "delivered", "cancelled", "canceled"}
    if not set(shipments["Status_Normalized"]).issubset(valid_statuses):
        raise ValueError("Shipments.Status must be In Transit, Delivered, or Cancelled.")
    shipments["Pending_Qty"] = shipments["Qty_Shipped"] - shipments["Qty_Received"]
    shipments["Active_Transit"] = (shipments["Status_Normalized"] == "in transit") & (shipments["Pending_Qty"] > 0)
    ref = pd.Timestamp(as_of or date.today())
    shipments["Delayed"] = shipments["Active_Transit"] & (shipments["ETA"] < ref)
    shipments["Within_7_Days"] = shipments["Active_Transit"] & shipments["ETA"].between(ref, ref + pd.Timedelta(days=7))
    stock["Available"] = stock["On_Hand"] - stock["Reserved"]
    return Inventory(warehouse, products, stock, plant, shipments, ref.date())


def stock_summary(inv: Inventory, warehouse_id: str | None = None, sku: str | None = None) -> pd.DataFrame:
    """One row per (warehouse, SKU); plant stock is NOT counted in available."""
    key = ["Warehouse_ID", "SKU"]
    stock = inv.stock.groupby(key, as_index=False)[["On_Hand", "Reserved", "Available"]].sum()
    plants = inv.plant.rename(columns={"Destination_Warehouse_ID": "Warehouse_ID"})
    plants = plants.groupby(key, as_index=False)["Qty_Ready"].sum().rename(columns={"Qty_Ready": "At_Plant"})
    transit = inv.shipments.loc[inv.shipments["Active_Transit"]].rename(columns={"Destination_Warehouse_ID": "Warehouse_ID"})
    incoming = transit.groupby(key, as_index=False)["Pending_Qty"].sum().rename(columns={"Pending_Qty": "In_Transit"})
    due = transit.loc[transit["Within_7_Days"]].groupby(key, as_index=False)["Pending_Qty"].sum().rename(columns={"Pending_Qty": "Due_7_Days"})
    keys = pd.concat([stock[key], plants[key], incoming[key]], ignore_index=True).drop_duplicates()
    for frame in (stock, plants, incoming, due):
        keys = keys.merge(frame, on=key, how="left")
    qty_cols = ["On_Hand", "Reserved", "Available", "At_Plant", "In_Transit", "Due_7_Days"]
    keys[qty_cols] = keys[qty_cols].fillna(0).astype(int)
    keys["Projected_Available_7_Days"] = keys["Available"] + keys["Due_7_Days"]
    keys = keys.merge(inv.warehouses[["Warehouse_ID", "Warehouse_Name", "City"]], on="Warehouse_ID", how="left")
    keys = keys.merge(inv.products[["SKU", "Product_Name", "Reorder_Level"]], on="SKU", how="left")
    keys["Low_Stock"] = keys["Available"] < keys["Reorder_Level"]
    if warehouse_id:
        keys = keys.loc[keys["Warehouse_ID"] == warehouse_id]
    if sku:
        keys = keys.loc[keys["SKU"] == sku]
    return keys.sort_values(["Warehouse_ID", "SKU"]).reset_index(drop=True)


def active_shipments(inv: Inventory, warehouse_id: str | None = None, sku: str | None = None) -> pd.DataFrame:
    frame = inv.shipments.loc[inv.shipments["Active_Transit"]].copy()
    if warehouse_id:
        frame = frame.loc[frame["Destination_Warehouse_ID"] == warehouse_id]
    if sku:
        frame = frame.loc[frame["SKU"] == sku]
    frame = frame.merge(inv.warehouses[["Warehouse_ID", "Warehouse_Name"]], left_on="Destination_Warehouse_ID", right_on="Warehouse_ID", how="left")
    frame = frame.merge(inv.products[["SKU", "Product_Name"]], on="SKU", how="left")
    frame["ETA_Status"] = frame["ETA"].apply(lambda x: "ETA unknown" if pd.isna(x) else ("Delayed" if x.date() < inv.snapshot else "Scheduled"))
    return frame.sort_values("ETA", na_position="last").reset_index(drop=True)


def plant_allocations(inv: Inventory, warehouse_id: str | None = None, sku: str | None = None) -> pd.DataFrame:
    frame = inv.plant.copy()
    if warehouse_id:
        frame = frame.loc[frame["Destination_Warehouse_ID"] == warehouse_id]
    if sku:
        frame = frame.loc[frame["SKU"] == sku]
    frame = frame.merge(inv.warehouses[["Warehouse_ID", "Warehouse_Name"]], left_on="Destination_Warehouse_ID", right_on="Warehouse_ID", how="left")
    frame = frame.merge(inv.products[["SKU", "Product_Name"]], on="SKU", how="left")
    return frame.sort_values("Planned_Dispatch_Date").reset_index(drop=True)


def metrics(inv: Inventory, warehouse_id: str | None = None, sku: str | None = None) -> dict:
    report = stock_summary(inv, warehouse_id, sku)
    shipments = active_shipments(inv, warehouse_id, sku)
    return {
        "on_hand": int(report["On_Hand"].sum()),
        "reserved": int(report["Reserved"].sum()),
        "available": int(report["Available"].sum()),
        "plant": int(report["At_Plant"].sum()),
        "transit": int(report["In_Transit"].sum()),
        "incoming_7d": int(report["Due_7_Days"].sum()),
        "projected_7d": int(report["Projected_Available_7_Days"].sum()),
        "delayed_count": int(shipments["Delayed"].sum()),
    }


def _detect_filters(question: str, inv: Inventory, warehouse_id: str | None, sku: str | None):
    lower = question.casefold()
    for _, record in inv.warehouses.iterrows():
        terms = [str(record["Warehouse_ID"]), str(record["Warehouse_Name"]), str(record["City"])]
        if any(t and t.casefold() in lower for t in terms):
            warehouse_id = record["Warehouse_ID"]
            break
    for _, record in inv.products.iterrows():
        terms = [str(record["SKU"]), str(record["Product_Name"])]
        if any(t and t.casefold() in lower for t in terms):
            sku = record["SKU"]
            break
    return warehouse_id, sku


def _rows_as_lines(frame: pd.DataFrame, formatter, limit: int = 9) -> str:
    if frame.empty:
        return "No matching records."
    lines = ["- " + formatter(row) for _, row in frame.head(limit).iterrows()]
    if len(frame) > limit:
        lines.append(f"- …and {len(frame)-limit} more. Use the tables below for all records.")
    return "\n".join(lines)


def _fmt_date(value) -> str:
    return "Not supplied" if pd.isna(value) else value.strftime("%d %b %Y")


def answer_question(question: str, inv: Inventory, selected_warehouse: str | None = None, selected_sku: str | None = None) -> str:
    """Grounded local agent. It answers from deterministic functions, not from invented estimates."""
    warehouse_id, sku = _detect_filters(question, inv, selected_warehouse, selected_sku)
    q = question.casefold()
    desc = ", ".join(x for x in [warehouse_id or "all warehouses", sku] if x)
    report = stock_summary(inv, warehouse_id, sku)
    shipments = active_shipments(inv, warehouse_id, sku)
    plant = plant_allocations(inv, warehouse_id, sku)
    m = metrics(inv, warehouse_id, sku)

    if any(term in q for term in ["delay", "late", "overdue", "missed eta"]):
        late = shipments.loc[shipments["Delayed"]]
        result = _rows_as_lines(late, lambda r: f"**{r.Shipment_ID}** • {r.Warehouse_Name} • {r.SKU} • {int(r.Pending_Qty)} units • ETA {_fmt_date(r.ETA)}")
        return f"**Delayed shipments for {desc}: {len(late)}** (as of {inv.snapshot}).\n\n{result}"

    if any(term in q for term in ["shortage", "low stock", "reorder", "below minimum"]):
        low = report.loc[report["Low_Stock"]]
        result = _rows_as_lines(low, lambda r: f"**{r.Warehouse_Name} / {r.SKU}**: available {int(r.Available)}; reorder level {int(r.Reorder_Level)}; in transit {int(r.In_Transit)}")
        return f"**Below-reorder-level warehouse/SKU lines for {desc}: {len(low)}**.\n\n{result}\n\nReorder levels are demo values unless replaced in Products."

    if any(term in q for term in ["eta", "when will", "arrival", "arrive", "arriving", "delivery", "deliver", "shipment", "transit", "in-transit", "git"]):
        window_match = re.search(r"(?:next|within|in the next)\s+(\d+)\s+days?", q)
        days = int(window_match.group(1)) if window_match else (7 if any(term in q for term in ["7 days", "next week", "this week", "upcoming week"]) else None)
        if days is not None:
            days = min(days, 365)
            end = pd.Timestamp(inv.snapshot) + pd.Timedelta(days=days)
            shipments = shipments.loc[shipments["ETA"].between(pd.Timestamp(inv.snapshot), end)]
        elif "tomorrow" in q:
            day = pd.Timestamp(inv.snapshot) + pd.Timedelta(days=1)
            shipments = shipments.loc[shipments["ETA"] == day]
        elif "today" in q:
            shipments = shipments.loc[shipments["ETA"] == pd.Timestamp(inv.snapshot)]
        label = f"Arrivals expected in the next {days} days" if days is not None else ("Arrivals expected tomorrow" if "tomorrow" in q else "Active shipments and ETAs")
        result = _rows_as_lines(shipments, lambda r: f"**{r.Shipment_ID}** → {r.Warehouse_Name} | {r.SKU} | {int(r.Pending_Qty)} pending | ETA **{_fmt_date(r.ETA)}** ({r.ETA_Status})")
        units = int(shipments["Pending_Qty"].sum())
        return f"**{label} for {desc}: {units} pending units across {len(shipments)} shipments.**\n\n{result}\n\nETAs come from the Excel file; they are not live vehicle predictions."

    if any(term in q for term in ["plant", "manufactur", "factory", "ready to dispatch"]):
        result = _rows_as_lines(plant, lambda r: f"**{r.Plant_Name}** → {r.Warehouse_Name}: {int(r.Qty_Ready)} units of {r.SKU}; planned dispatch {_fmt_date(r.Planned_Dispatch_Date)}")
        return f"**Unshipped goods at plants allocated to {desc}: {m['plant']} units.**\n\n{result}\n\nThese units are not yet available at the warehouse."

    if any(term in q for term in ["forecast", "projected", "next 7 days", "after 7 days", "expected stock"]):
        return (f"**7-day warehouse stock projection for {desc}: {m['projected_7d']:,} units.**\n\n"
                f"Available now: {m['available']:,} + In-transit with ETA within 7 days: {m['incoming_7d']:,}. "
                "This is a simple inbound-only projection; it excludes future consumption, outgoing shipments and unshipped plant stock.")

    if any(term in q for term in ["available", "stock", "inventory", "quantity", "warehouse", "how much", "status", "show"]):
        breakdown = _rows_as_lines(report, lambda r: f"**{r.Warehouse_Name} / {r.SKU}**: available {int(r.Available)}; in transit {int(r.In_Transit)}; at plant {int(r.At_Plant)}")
        return (f"**Stock position for {desc} (as of {inv.snapshot}):**\n\n"
                f"- Warehouse on hand: **{m['on_hand']:,}** units\n"
                f"- Reserved: **{m['reserved']:,}** units\n"
                f"- Available at warehouse: **{m['available']:,}** units\n"
                f"- In transit: **{m['transit']:,}** units\n"
                f"- Still at plant: **{m['plant']:,}** units\n"
                f"- Delayed shipments: **{m['delayed_count']}**\n\n{breakdown}")

    return (f"I can help with warehouse stock, unshipped plant goods, goods in transit, ETAs, delivery delays, "
            f"and reorder alerts. For example: **How much stock is available in Chennai?** "
            f"or **Which shipments are delayed?**")
