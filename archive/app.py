
from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st


# ============================================================
# 1. PAGE CONFIGURATION
# ============================================================
st.set_page_config(
    page_title="Ecommerce Analytics Dashboard",
    page_icon="🛒",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.title("🛒 Ecommerce Analytics Dashboard")
st.write(
    "Explore ecommerce sales, orders, customers, locations, "
    "product categories, refunds, and profit/loss."
)


# ============================================================
# 2. DATASET LOADING
# ============================================================
project = Path(__file__).resolve().parent.parent
files = list(project.rglob("*.xlsx")) + list(project.rglob("*.csv"))
files = [
    file for file in files
    if not file.name.startswith("~$")
    and file.name.lower() not in {"ecommerce_analysis.csv"}
]

uploaded = st.sidebar.file_uploader(
    "Upload your dataset",
    type=["csv", "xlsx"],
)


def load_csv(source):
    """Read CSV files with common encodings."""
    last_error = None
    for encoding in ("utf-8-sig", "utf-8", "cp1252", "latin1"):
        try:
            if hasattr(source, "seek"):
                source.seek(0)
            return pd.read_csv(source, encoding=encoding)
        except UnicodeDecodeError as error:
            last_error = error
    raise ValueError(f"Could not decode the CSV file: {last_error}")


def load_dataset(source):
    """Load CSV or Excel data from an uploaded file or disk path."""
    filename = source.name if hasattr(source, "name") else str(source)
    filename = filename.lower()
    if filename.endswith(".csv"):
        return load_csv(source)
    if filename.endswith((".xlsx", ".xls")):
        return pd.read_excel(source)
    raise ValueError("Please use a CSV or Excel (.xlsx/.xls) file.")


df = None
try:
    if uploaded is not None:
        df = load_dataset(uploaded)
    else:
        preferred_names = {
            "india_ecommerce_transactions_2023_2025.csv",
            "india_ecommerce_transactions_2023_2025.xlsx",
            "ecommerce-sales-v1.0.0-analysis.xlsx",
        }
        preferred = [
            file for file in files
            if file.name.lower() in preferred_names
        ]
        selected_file = preferred[0] if preferred else (files[0] if files else None)
        if selected_file is not None:
            df = load_dataset(selected_file)
except Exception as error:
    st.error(f"Could not load dataset: {error}")

if df is None or df.empty:
    st.warning(
        "No usable dataset was loaded. Upload a CSV or XLSX file in the sidebar, "
        "or place the dataset inside your project folder."
    )
    st.stop()


# ============================================================
# 3. NORMALIZE COLUMN NAMES AND CLEAN DATA
# ============================================================
df.columns = (
    df.columns.astype(str)
    .str.strip()
    .str.lower()
    .str.replace(r"[^a-z0-9]+", "_", regex=True)
    .str.strip("_")
)
df = df.loc[:, ~df.columns.duplicated()].copy()

st.sidebar.success(f"Loaded {len(df):,} records")
st.caption(f"Dataset: {len(df):,} rows × {len(df.columns):,} columns")

columns = set(df.columns)


def find_column(candidates):
    return next((name for name in candidates if name in columns), None)


date_col = find_column(["ordered_at", "order_date", "purchase_date", "date", "created_at"])
order_col = find_column(["order_id", "order_number", "order_no", "transaction_id"])
customer_col = find_column(["customer_id", "customer", "user_id"])
sales_col = find_column(["net_revenue", "net_sales", "sales", "gross_payment", "total_sales", "subtotal"])
gross_col = find_column(["gross_payment", "gross_revenue", "total_amount", "order_total"])
refund_col = find_column(["refund_amount", "refunds", "returned_amount"])
status_col = find_column(["order_status", "status", "transaction_status"])
country_col = find_column(["country_code", "country"])
state_col = find_column(["state", "state_name", "province"])
city_col = find_column(["city", "city_name", "town"])
channel_col = find_column(["acquisition_channel", "channel", "sales_channel"])
segment_col = find_column(["customer_segment", "segment"])
category_col = find_column(["product_category", "category", "category_name"])
product_col = find_column(["product_name", "product", "item_name"])
units_col = find_column(["units", "quantity", "qty"])
profit_col = find_column(["profit_loss", "profit", "net_profit", "profit_amount"])
cogs_col = find_column(["cost_of_goods_sold", "cogs", "cost", "product_cost"])
expenses_col = find_column(["operating_expenses", "operating_costs", "expenses", "expense_amount"])

numeric_candidates = [
    "subtotal", "tax_amount", "shipping_amount", "gross_payment",
    "refund_amount", "net_revenue", "sales", "units", "line_count",
    "unit_price", "cost_of_goods_sold", "operating_expenses",
    "profit_loss", "profit", "net_profit", "quantity", "qty",
]
for column in numeric_candidates:
    if column in df.columns:
        cleaned = (
            df[column].astype(str)
            .str.replace(",", "", regex=False)
            .str.replace("₹", "", regex=False)
            .str.replace("$", "", regex=False)
            .str.strip()
        )
        df[column] = pd.to_numeric(cleaned, errors="coerce")

if date_col:
    df[date_col] = pd.to_datetime(df[date_col], errors="coerce")

# If profit_loss is absent but revenue, COGS, and expenses exist, calculate it.
if profit_col is None and sales_col and cogs_col:
    if expenses_col:
        df["calculated_profit_loss"] = (
            pd.to_numeric(df[sales_col], errors="coerce").fillna(0)
            - pd.to_numeric(df[cogs_col], errors="coerce").fillna(0)
            - pd.to_numeric(df[expenses_col], errors="coerce").fillna(0)
        )
    else:
        df["calculated_profit_loss"] = (
            pd.to_numeric(df[sales_col], errors="coerce").fillna(0)
            - pd.to_numeric(df[cogs_col], errors="coerce").fillna(0)
        )
    profit_col = "calculated_profit_loss"

# Prefer a true city column. A mapping is only a fallback for datasets that
# contain state but do not contain city-level records.
CITY_STATE_MAPPING = {
    "Delhi": "Delhi",
    "Mumbai": "Maharashtra",
    "Pune": "Maharashtra",
    "Nagpur": "Maharashtra",
    "Nashik": "Maharashtra",
    "Lucknow": "Uttar Pradesh",
    "Prayagraj": "Uttar Pradesh",
    "Kanpur": "Uttar Pradesh",
    "Varanasi": "Uttar Pradesh",
    "Agra": "Uttar Pradesh",
    "Jaipur": "Rajasthan",
    "Jodhpur": "Rajasthan",
    "Udaipur": "Rajasthan",
    "Chennai": "Tamil Nadu",
    "Coimbatore": "Tamil Nadu",
    "Madurai": "Tamil Nadu",
    "Kolkata": "West Bengal",
    "Howrah": "West Bengal",
    "Siliguri": "West Bengal",
    "Visakhapatnam": "Andhra Pradesh",
    "Vijayawada": "Andhra Pradesh",
    "Tirupati": "Andhra Pradesh",
    "Ahmedabad": "Gujarat",
    "Surat": "Gujarat",
    "Vadodara": "Gujarat",
}
INDIA_STATES = [
    "Andhra Pradesh", "Delhi", "Gujarat", "Maharashtra",
    "Rajasthan", "Tamil Nadu", "Uttar Pradesh", "West Bengal",
]


# ============================================================
# 4. SIDEBAR FILTERS
# ============================================================
st.sidebar.header("🔎 Dashboard Filters")
filtered = df.copy()

if date_col and df[date_col].notna().any():
    valid_dates = df[date_col].dropna()
    min_date, max_date = valid_dates.min().date(), valid_dates.max().date()
    date_selection = st.sidebar.date_input(
        "Order date range",
        value=(min_date, max_date),
        min_value=min_date,
        max_value=max_date,
    )
    if isinstance(date_selection, (tuple, list)) and len(date_selection) == 2:
        start_date, end_date = date_selection
        filtered = filtered[
            filtered[date_col].notna()
            & filtered[date_col].dt.date.between(start_date, end_date)
        ].copy()

if country_col:
    options = sorted(filtered[country_col].dropna().astype(str).unique().tolist())
    selected = st.sidebar.multiselect("Country", options, default=options)
    filtered = filtered[filtered[country_col].astype(str).isin(selected)].copy()

# State selection plus city search.
if state_col:
    st.sidebar.subheader("📍 Location")
    states_available = sorted(filtered[state_col].dropna().astype(str).unique().tolist())
    state_options = states_available if states_available else INDIA_STATES
    selected_states = st.sidebar.multiselect(
        "Select states",
        options=state_options,
        default=state_options,
        key="state_filter",
    )

    city_search = st.sidebar.text_input(
        "Search by city",
        placeholder="e.g. Mumbai, Lucknow, Jaipur",
        help="Searches the City column when available. If your dataset has no City column, this can only narrow results to the mapped state.",
    ).strip()

    filtered = filtered[filtered[state_col].astype(str).isin(selected_states)].copy()

    if city_search:
        if city_col:
            city_matches = filtered[city_col].fillna("").astype(str).str.contains(
                city_search, case=False, regex=False
            )
            filtered = filtered[city_matches].copy()
            if not city_matches.any():
                st.sidebar.caption("No matching city found in the current selection.")
        else:
            matching_states = [
                state for city, state in CITY_STATE_MAPPING.items()
                if city_search.lower() in city.lower()
            ]
            if matching_states:
                filtered = filtered[
                    filtered[state_col].astype(str).isin(matching_states)
                ].copy()
                st.sidebar.warning(
                    "This dataset has no City column. Search narrows results by state, "
                    "not by individual city. Add a city column for true city-level filtering."
                )
            else:
                filtered = filtered.iloc[0:0].copy()
                st.sidebar.info("No matching city found in the built-in city list.")

if channel_col:
    options = sorted(filtered[channel_col].dropna().astype(str).unique().tolist())
    selected = st.sidebar.multiselect("Acquisition Channel", options, default=options)
    filtered = filtered[filtered[channel_col].astype(str).isin(selected)].copy()

if status_col:
    options = sorted(filtered[status_col].dropna().astype(str).unique().tolist())
    selected = st.sidebar.multiselect("Order Status", options, default=options)
    filtered = filtered[filtered[status_col].astype(str).isin(selected)].copy()

if category_col:
    options = sorted(filtered[category_col].dropna().astype(str).unique().tolist())
    selected = st.sidebar.multiselect("Product Category", options, default=options)
    filtered = filtered[filtered[category_col].astype(str).isin(selected)].copy()

st.sidebar.caption(f"Showing {len(filtered):,} of {len(df):,} records")

if filtered.empty:
    st.warning("No records match your selected filters. Adjust the sidebar filters.")
    st.stop()


# ============================================================
# 5. KEY PERFORMANCE INDICATORS
# ============================================================
st.subheader("📊 Key Performance Indicators")
total_records = len(filtered)
total_orders = filtered[order_col].nunique() if order_col else total_records
total_customers = filtered[customer_col].nunique() if customer_col else None
total_revenue = pd.to_numeric(filtered[sales_col], errors="coerce").fillna(0).sum() if sales_col else None
average_order_value = total_revenue / total_orders if total_revenue is not None and total_orders else None
total_units = pd.to_numeric(filtered[units_col], errors="coerce").fillna(0).sum() if units_col else None
total_profit = pd.to_numeric(filtered[profit_col], errors="coerce").fillna(0).sum() if profit_col else None
total_refunds = pd.to_numeric(filtered[refund_col], errors="coerce").fillna(0).sum() if refund_col else None

kpis = st.columns(4)
kpis[0].metric("Total Orders", f"{total_orders:,}")
kpis[1].metric("Unique Customers", f"{total_customers:,}" if total_customers is not None else "N/A")
kpis[2].metric("Net Revenue", f"₹{total_revenue:,.2f}" if total_revenue is not None else "N/A")
kpis[3].metric("Average Order Value", f"₹{average_order_value:,.2f}" if average_order_value is not None else "N/A")

kpis2 = st.columns(4)
kpis2[0].metric("Total Records", f"{total_records:,}")
kpis2[1].metric("Units Sold", f"{total_units:,.0f}" if total_units is not None else "N/A")
kpis2[2].metric("Total Profit / Loss", f"₹{total_profit:,.2f}" if total_profit is not None else "N/A")
kpis2[3].metric("Total Refunds", f"₹{total_refunds:,.2f}" if total_refunds is not None else "N/A")

if profit_col:
    st.caption(
        "Profit/loss uses the dataset's profit column"
        + (" (calculated from revenue minus cost and expenses where available)." if profit_col == "calculated_profit_loss" else ".")
    )
else:
    st.info(
        "Profit/loss chart is unavailable because the dataset has no profit column "
        "and no sufficient cost columns to calculate one."
    )


# ============================================================
# 6. MONTHLY REVENUE TREND
# ============================================================
if date_col and sales_col:
    st.subheader("📈 Monthly Revenue Trend")
    trend = filtered.dropna(subset=[date_col]).copy()
    trend["Month"] = trend[date_col].dt.to_period("M").dt.to_timestamp()
    trend = trend.groupby("Month", as_index=False)[sales_col].sum().sort_values("Month")
    if not trend.empty:
        fig = px.line(
            trend, x="Month", y=sales_col, markers=True,
            title="Monthly Revenue",
            labels={"Month": "Month", sales_col: "Revenue"},
        )
        fig.update_layout(hovermode="x unified")
        st.plotly_chart(fig, use_container_width=True)


# ============================================================
# 7. MONTHLY PROFIT AND LOSS
# ============================================================
if date_col and profit_col:
    st.subheader("💹 Monthly Profit and Loss")
    pl_data = filtered.dropna(subset=[date_col]).copy()
    pl_data["Month"] = pl_data[date_col].dt.to_period("M").dt.to_timestamp()
    pl_data[profit_col] = pd.to_numeric(pl_data[profit_col], errors="coerce").fillna(0)
    monthly_pl = pl_data.groupby("Month", as_index=False)[profit_col].sum().sort_values("Month")
    if not monthly_pl.empty:
        fig = px.line(
            monthly_pl, x="Month", y=profit_col, markers=True,
            title="Monthly Profit and Loss",
            labels={"Month": "Month", profit_col: "Profit / Loss (₹)"},
        )
        fig.add_hline(y=0, line_dash="dash")
        fig.update_layout(hovermode="x unified")
        st.plotly_chart(fig, use_container_width=True)

# Profit/loss by state
if state_col and profit_col:
    st.subheader("🗺️ Profit and Loss by State")
    state_pl = filtered.copy()
    state_pl[profit_col] = pd.to_numeric(state_pl[profit_col], errors="coerce").fillna(0)
    state_pl = state_pl.groupby(state_col, as_index=False)[profit_col].sum().sort_values(profit_col, ascending=False)
    if not state_pl.empty:
        fig = px.bar(
            state_pl, x=state_col, y=profit_col, color=profit_col,
            title="Total Profit / Loss by State",
            labels={state_col: "State", profit_col: "Profit / Loss (₹)"},
        )
        fig.add_hline(y=0, line_dash="dash")
        st.plotly_chart(fig, use_container_width=True)


# ============================================================
# 8. REVENUE BY ACQUISITION CHANNEL
# ============================================================
if channel_col and sales_col:
    st.subheader("🥧 Revenue by Acquisition Channel")
    channel_data = filtered.dropna(subset=[channel_col]).groupby(channel_col, as_index=False)[sales_col].sum()
    channel_data = channel_data.sort_values(sales_col, ascending=False)
    if not channel_data.empty:
        fig = px.pie(channel_data, names=channel_col, values=sales_col, title="Revenue Distribution by Channel", hole=0.4)
        fig.update_traces(textposition="inside", textinfo="percent+label")
        st.plotly_chart(fig, use_container_width=True)


# ============================================================
# 9. ORDERS BY STATUS
# ============================================================
if status_col:
    st.subheader("📦 Orders by Status")
    status_data = (
        filtered[status_col].fillna("Unknown").astype(str).value_counts()
        .rename_axis("Order Status").reset_index(name="Number of Records")
    )
    if not status_data.empty:
        fig = px.bar(
            status_data, x="Order Status", y="Number of Records",
            color="Order Status", title="Order Status Distribution",
            text="Number of Records",
        )
        st.plotly_chart(fig, use_container_width=True)


# ============================================================
# 10. ORDERS BY COUNTRY / STATE / CITY
# ============================================================
if country_col and order_col:
    st.subheader("🌍 Orders by Country")
    country_data = filtered.groupby(country_col)[order_col].nunique().reset_index(name="Number of Orders")
    country_data = country_data.sort_values("Number of Orders", ascending=False).head(15)
    if not country_data.empty:
        fig = px.bar(
            country_data, x=country_col, y="Number of Orders",
            color=country_col, title="Top Countries by Orders",
            text="Number of Orders",
        )
        st.plotly_chart(fig, use_container_width=True)

if state_col and order_col:
    st.subheader("📍 Orders by State")
    state_orders = filtered.groupby(state_col)[order_col].nunique().reset_index(name="Number of Orders")
    state_orders = state_orders.sort_values("Number of Orders", ascending=False)
    if not state_orders.empty:
        fig = px.bar(
            state_orders, x=state_col, y="Number of Orders",
            color=state_col, title="Orders by State",
            text="Number of Orders",
        )
        st.plotly_chart(fig, use_container_width=True)

if city_col and order_col:
    st.subheader("🏙️ Orders by City")
    city_orders = filtered.groupby(city_col)[order_col].nunique().reset_index(name="Number of Orders")
    city_orders = city_orders.sort_values("Number of Orders", ascending=False).head(20)
    if not city_orders.empty:
        fig = px.bar(
            city_orders, x=city_col, y="Number of Orders",
            title="Top 20 Cities by Orders", text="Number of Orders",
        )
        st.plotly_chart(fig, use_container_width=True)


# ============================================================
# 11. CUSTOMERS BY SEGMENT
# ============================================================
if segment_col and customer_col:
    st.subheader("👥 Customers by Segment")
    segment_data = filtered.groupby(segment_col)[customer_col].nunique().reset_index(name="Unique Customers")
    segment_data = segment_data.sort_values("Unique Customers", ascending=False)
    if not segment_data.empty:
        fig = px.bar(
            segment_data, x=segment_col, y="Unique Customers",
            color=segment_col, title="Unique Customers by Segment",
            text="Unique Customers",
        )
        st.plotly_chart(fig, use_container_width=True)


# ============================================================
# 12. REVENUE BY STATE / CATEGORY
# ============================================================
if state_col and sales_col:
    st.subheader("💰 Revenue by State")
    state_revenue = filtered.groupby(state_col, as_index=False)[sales_col].sum().sort_values(sales_col, ascending=False)
    if not state_revenue.empty:
        fig = px.bar(
            state_revenue, x=state_col, y=sales_col,
            color=state_col, title="Revenue by State",
            labels={sales_col: "Revenue (₹)"},
        )
        st.plotly_chart(fig, use_container_width=True)

if category_col and sales_col:
    st.subheader("🛍️ Revenue by Product Category")
    category_data = filtered.groupby(category_col, as_index=False)[sales_col].sum().sort_values(sales_col, ascending=False)
    if not category_data.empty:
        fig = px.bar(
            category_data, x=category_col, y=sales_col,
            color=category_col, title="Revenue by Product Category",
            labels={sales_col: "Revenue (₹)"},
        )
        st.plotly_chart(fig, use_container_width=True)

if product_col and sales_col:
    st.subheader("🏆 Top 10 Products by Revenue")
    product_data = filtered.groupby(product_col, as_index=False)[sales_col].sum().sort_values(sales_col, ascending=False).head(10)
    if not product_data.empty:
        fig = px.bar(
            product_data, x=sales_col, y=product_col,
            orientation="h", title="Top 10 Products by Revenue",
            labels={sales_col: "Revenue (₹)", product_col: "Product"},
        )
        st.plotly_chart(fig, use_container_width=True)


# ============================================================
# 13. UNITS SOLD BY CHANNEL
# ============================================================
if channel_col and units_col:
    st.subheader("📦 Units Sold by Acquisition Channel")
    units_data = filtered.groupby(channel_col, as_index=False)[units_col].sum().sort_values(units_col, ascending=False)
    if not units_data.empty:
        fig = px.bar(
            units_data, x=channel_col, y=units_col,
            color=channel_col, title="Units Sold by Channel",
            text=units_col,
        )
        st.plotly_chart(fig, use_container_width=True)


# ============================================================
# 14. FINANCIAL BREAKDOWN
# ============================================================
financial_columns = [
    column for column in [
        "subtotal", "tax_amount", "shipping_amount",
        "gross_payment", "refund_amount", "net_revenue",
        "cost_of_goods_sold", "operating_expenses", "profit_loss",
    ]
    if column in filtered.columns
]
if financial_columns:
    st.subheader("💵 Revenue, Refund, Cost and Expense Breakdown")
    financial_data = pd.DataFrame({
        "Component": financial_columns,
        "Amount": [
            pd.to_numeric(filtered[column], errors="coerce").fillna(0).sum()
            for column in financial_columns
        ],
    })
    fig = px.bar(
        financial_data, x="Component", y="Amount", color="Component",
        title="Financial Components", text="Amount",
    )
    fig.update_traces(texttemplate="%{text:,.2f}", textposition="outside", cliponaxis=False)
    st.plotly_chart(fig, use_container_width=True)


# ============================================================
# 15. TOP CUSTOMERS
# ============================================================
if customer_col and sales_col:
    st.subheader("🏆 Top 10 Customers by Revenue")
    customers = filtered.groupby(customer_col, as_index=False)[sales_col].sum()
    customers = customers.sort_values(sales_col, ascending=False).head(10)
    if not customers.empty:
        fig = px.bar(
            customers, x=customer_col, y=sales_col,
            title="Top 10 Customers by Revenue",
            labels={customer_col: "Customer", sales_col: "Revenue (₹)"},
        )
        st.plotly_chart(fig, use_container_width=True)
        st.dataframe(customers, use_container_width=True, hide_index=True)


# ============================================================
# 16. FILTERED DATA AND EXPORT
# ============================================================
st.subheader("📋 Filtered Dataset")
st.dataframe(filtered, use_container_width=True, hide_index=True)

st.subheader("📥 Export Data")
csv_data = filtered.to_csv(index=False).encode("utf-8-sig")
st.download_button(
    label="⬇️ Download Filtered CSV",
    data=csv_data,
    file_name="ecommerce_analysis_filtered.csv",
    mime="text/csv",
)

# Optional Excel export.
from io import BytesIO
excel_buffer = BytesIO()
try:
    with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
        filtered.to_excel(writer, index=False, sheet_name="Filtered Data")
    st.download_button(
        label="⬇️ Download Filtered Excel",
        data=excel_buffer.getvalue(),
        file_name="ecommerce_analysis_filtered.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
except Exception:
    st.caption("Excel export needs openpyxl installed. Install it with: pip install openpyxl")


# ============================================================
# 17. FOOTER
# ============================================================
with st.expander("ℹ️ About This Dashboard"):
    st.write(
        "This interactive dashboard uses Python, Pandas, Streamlit, and Plotly. "
        "Filters update the KPIs, charts, data preview, and exports."
    )
    st.write(
        "City-level filtering and charts require a City column in the dataset. "
        "If the City column is missing, the city search can only narrow results "
        "by the mapped state and displays a warning."
    )
    st.write(
        "Profit/loss is taken from the dataset's profit/loss column when available. "
        "Otherwise it is calculated from revenue minus cost of goods sold and "
        "operating expenses, where those columns are available."
    )

st.caption("Built with Python, Pandas, Streamlit, and Plotly.")
