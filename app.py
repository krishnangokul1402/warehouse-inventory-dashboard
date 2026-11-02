"""Warehouse Inventory Agent - Streamlit user interface.
Run: streamlit run app.py
"""

from __future__ import annotations

import os
from pathlib import Path

import pandas as pd
import streamlit as st

from inventory_agent import (
    active_shipments,
    answer_question,
    load_inventory,
    metrics,
    plant_allocations,
    stock_summary,
)

st.set_page_config(page_title="Warehouse Control Tower", page_icon="📦", layout="wide")
st.title("📦 Warehouse Inventory Control Tower")
st.caption("Warehouse stock · Goods in plant · Goods in transit · Delivery ETA · Delayed shipments")

sample_path = Path(__file__).resolve().parent / "Warehouse_Inventory_Sample.xlsx"
with st.sidebar:
    st.header("Data source")
    uploaded = st.file_uploader("Upload updated inventory workbook", type=["xlsx"])
    st.caption("If no workbook is uploaded, fictional sample data is used.")
    ai_enabled = st.checkbox("Use optional OpenAI explanation", value=False, help="Off: all queries run locally. On: the calculated answer and your question are sent to OpenAI.")
    st.caption("AI mode requires an OPENAI_API_KEY environment variable. Numeric results are always calculated locally.")

try:
    inv = load_inventory(uploaded if uploaded is not None else sample_path)
except Exception as exc:
    st.error(f"Could not read workbook: {exc}")
    st.info("Use the supplied Excel template and keep the original worksheet and column names.")
    st.stop()

warehouse_options = {"All warehouses": None}
warehouse_options.update({f"{r.Warehouse_Name} ({r.Warehouse_ID})": r.Warehouse_ID for _, r in inv.warehouses.iterrows()})
product_options = {"All products": None}
product_options.update({f"{r.SKU} — {r.Product_Name}": r.SKU for _, r in inv.products.iterrows()})

with st.sidebar:
    st.header("Filters")
    wh_name = st.selectbox("Destination warehouse", list(warehouse_options))
    product_name = st.selectbox("Product / SKU", list(product_options))
    warehouse_id = warehouse_options[wh_name]
    sku = product_options[product_name]

m = metrics(inv, warehouse_id, sku)
st.caption(f"Based on uploaded data | Current date for ETA checks: {inv.snapshot:%d %b %Y} | {'UPLOADED WORKBOOK' if uploaded else 'DEMO DATA'}")
r1 = st.columns(3)
r1[0].metric("Warehouse on hand", f"{m['on_hand']:,}")
r1[1].metric("Available (less reserved)", f"{m['available']:,}")
r1[2].metric("Goods in transit", f"{m['transit']:,}")
r2 = st.columns(3)
r2[0].metric("Goods ready at plant", f"{m['plant']:,}")
r2[1].metric("Arriving within 7 days", f"{m['incoming_7d']:,}")
r2[2].metric("Delayed shipments", f"{m['delayed_count']:,}")

summary = stock_summary(inv, warehouse_id, sku)
shipments = active_shipments(inv, warehouse_id, sku)
plants = plant_allocations(inv, warehouse_id, sku)

tab1, tab2, tab3, tab4 = st.tabs(["🏭 Warehouse stock", "🚚 Goods in transit & ETA", "🏗️ Goods at plant", "💬 Ask the agent"])
with tab1:
    st.subheader("Warehouse-wise stock availability")
    st.caption("Projected 7-day available = available now + pending in-transit goods due within 7 days; does not subtract future demand.")
    columns = ["Warehouse_Name", "SKU", "Product_Name", "On_Hand", "Reserved", "Available", "In_Transit", "At_Plant", "Due_7_Days", "Projected_Available_7_Days", "Reorder_Level", "Low_Stock"]
    st.dataframe(summary[columns], width="stretch", hide_index=True)
    st.download_button("Download stock report (CSV)", summary[columns].to_csv(index=False).encode(), file_name="warehouse_stock_status.csv", mime="text/csv")
    grouped = summary.groupby("Warehouse_Name", as_index=True)[["Available", "In_Transit", "At_Plant"]].sum()
    if not grouped.empty:
        st.bar_chart(grouped)

with tab2:
    st.subheader("Open shipments and delivery ETAs")
    st.caption("Pending = shipped − received. A shipment is delayed when its ETA is before today and goods remain in transit.")
    cols = ["Shipment_ID", "Warehouse_Name", "SKU", "Origin_Plant", "Qty_Shipped", "Qty_Received", "Pending_Qty", "Dispatch_Date", "ETA", "ETA_Status"]
    st.dataframe(shipments[cols], hide_index=True, width="stretch")
    if shipments["Delayed"].any():
        st.warning(f"{int(shipments['Delayed'].sum())} shipment(s) have passed their stated ETA.")

with tab3:
    st.subheader("Goods still at plant — not yet dispatched")
    cols = ["Plant_ID", "Plant_Name", "Warehouse_Name", "SKU", "Product_Name", "Qty_Ready", "Planned_Dispatch_Date"]
    st.dataframe(plants[cols], hide_index=True, width="stretch")

with tab4:
    st.subheader("Ask a warehouse question")
    st.caption("Try: 'How much stock is available in Chennai?', 'ETA for SKU-101', or 'Which shipments are delayed?'")
    for row in st.session_state.get("messages", []):
        with st.chat_message(row["role"]):
            st.markdown(row["content"])
    question = st.chat_input("Ask about inventory, plant quantities, transit or ETA")
    if question:
        st.session_state.setdefault("messages", []).append({"role": "user", "content": question})
        response = answer_question(question, inv, warehouse_id, sku)
        if ai_enabled:
            api_key = os.getenv("OPENAI_API_KEY")
            if api_key:
                try:
                    from openai import OpenAI
                    client = OpenAI(api_key=api_key)
                    ai_response = client.responses.create(
                        model="gpt-4.1-mini",
                        instructions=("Write a concise, clear warehouse operations answer using ONLY the "
                                      "verified facts below. Keep every quantity, date, and qualifier intact. "
                                      "Do not make up ETAs or availability. This is an explanation, not a prediction."),
                        input=f"User question: {question}\n\nVerified calculation:\n{response}",
                    )
                    if ai_response.output_text.strip():
                        response = ai_response.output_text
                except Exception as exc:
                    response = response + f"\n\n_AI explanation unavailable ({type(exc).__name__}); local answer shown._"
            else:
                response += "\n\n_AI mode is selected but OPENAI_API_KEY is missing; local answer shown._"
        st.session_state["messages"].append({"role": "assistant", "content": response})
        st.rerun()

st.divider()
st.caption("Sample data are fictional. Update plant stock when dispatching and warehouse stock when receiving goods to avoid double counting. All ETAs originate from your workbook, not live GPS.")
