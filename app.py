import os
import sqlite3
import io
import csv
from datetime import datetime
from functools import wraps

from flask import (
    Flask, render_template, request, redirect, url_for, session, jsonify, Response
)
from werkzeug.security import generate_password_hash, check_password_hash

try:
    from openpyxl import Workbook
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

try:
    from reportlab.lib.pagesizes import letter, landscape
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.lib import colors
    HAS_REPORTLAB = True
except ImportError:
    HAS_REPORTLAB = False

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "sales-dashboard-secret-key-2025")

DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "sales.db")


def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if "user_id" not in session:
            return redirect(url_for("login"))
        return f(*args, **kwargs)
    return decorated_function


def build_where_clause(args):
    conditions = []
    params = []

    year = args.get("year", "")
    if year:
        conditions.append("strftime('%Y', order_date) = ?")
        params.append(year)

    month = args.get("month", "")
    if month:
        conditions.append("strftime('%m', order_date) = ?")
        params.append(month.zfill(2))

    start_date = args.get("start_date", "")
    if start_date:
        conditions.append("order_date >= ?")
        params.append(start_date)

    end_date = args.get("end_date", "")
    if end_date:
        conditions.append("order_date <= ?")
        params.append(end_date)

    region = args.get("region", "")
    if region:
        conditions.append("region = ?")
        params.append(region)

    state = args.get("state", "")
    if state:
        conditions.append("state = ?")
        params.append(state)

    city = args.get("city", "")
    if city:
        conditions.append("city = ?")
        params.append(city)

    category = args.get("category", "")
    if category:
        conditions.append("category = ?")
        params.append(category)

    sub_category = args.get("sub_category", "")
    if sub_category:
        conditions.append("sub_category = ?")
        params.append(sub_category)

    segment = args.get("segment", "")
    if segment:
        conditions.append(
            "customer_name IN (SELECT customer_name FROM customers WHERE segment = ?)"
        )
        params.append(segment)

    search = args.get("search", "")
    if search:
        conditions.append("product_name LIKE ?")
        params.append(f"%{search}%")

    where = " AND ".join(conditions) if conditions else "1=1"
    return where, params


# ---------- Auth Routes ----------

@app.route("/login", methods=["GET", "POST"])
def login():
    if "user_id" in session:
        return redirect(url_for("dashboard"))

    error = None
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "")
        conn = get_db()
        user = conn.execute(
            "SELECT * FROM users WHERE username = ?", (username,)
        ).fetchone()
        conn.close()

        if user and check_password_hash(user["password_hash"], password):
            session["user_id"] = user["id"]
            session["username"] = user["username"]
            session["role"] = user["role"]
            return redirect(url_for("dashboard"))
        else:
            error = "Invalid username or password."

    return render_template("login.html", error=error)


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


# ---------- Page Routes ----------

@app.route("/")
@login_required
def dashboard():
    return render_template("dashboard.html")


@app.route("/products")
@login_required
def products_page():
    return render_template("products.html")


@app.route("/customers")
@login_required
def customers_page():
    return render_template("customers.html")


@app.route("/reports")
@login_required
def reports_page():
    return render_template("reports.html")


@app.route("/analytics")
@login_required
def analytics_page():
    return render_template("analytics.html")


# ---------- API Routes ----------

@app.route("/api/dashboard")
@login_required
def api_dashboard():
    where, params = build_where_clause(request.args)
    conn = get_db()

    row = conn.execute(
        f"SELECT COALESCE(SUM(sales),0) as total_sales, COALESCE(SUM(profit),0) as total_profit, "
        f"COUNT(*) as total_orders, COUNT(DISTINCT customer_name) as total_customers, "
        f"COALESCE(AVG(sales),0) as avg_order_value "
        f"FROM orders WHERE {where}", params
    ).fetchone()

    today = datetime.now().strftime("%Y-%m-%d")
    today_row = conn.execute(
        "SELECT COALESCE(SUM(sales),0) as today_sales FROM orders WHERE order_date = ?",
        (today,)
    ).fetchone()

    margin = (row["total_profit"] / row["total_sales"] * 100) if row["total_sales"] > 0 else 0

    years = [r[0] for r in conn.execute(
        "SELECT DISTINCT strftime('%Y', order_date) as y FROM orders ORDER BY y"
    ).fetchall()]

    regions = [r[0] for r in conn.execute("SELECT DISTINCT region FROM orders ORDER BY region").fetchall()]
    states = [r[0] for r in conn.execute("SELECT DISTINCT state FROM orders ORDER BY state").fetchall()]
    cities = [r[0] for r in conn.execute("SELECT DISTINCT city FROM orders ORDER BY city").fetchall()]
    categories = [r[0] for r in conn.execute("SELECT DISTINCT category FROM orders ORDER BY category").fetchall()]
    sub_categories = [r[0] for r in conn.execute("SELECT DISTINCT sub_category FROM orders ORDER BY sub_category").fetchall()]
    segments = [r[0] for r in conn.execute("SELECT DISTINCT segment FROM customers ORDER BY segment").fetchall()]

    conn.close()

    return jsonify({
        "total_sales": round(row["total_sales"], 2),
        "total_profit": round(row["total_profit"], 2),
        "total_orders": row["total_orders"],
        "total_customers": row["total_customers"],
        "avg_order_value": round(row["avg_order_value"], 2),
        "profit_margin": round(margin, 2),
        "today_sales": round(today_row["today_sales"], 2),
        "years": years,
        "regions": regions,
        "states": states,
        "cities": cities,
        "categories": categories,
        "sub_categories": sub_categories,
        "segments": segments,
    })


@app.route("/api/monthly-sales")
@login_required
def api_monthly_sales():
    where, params = build_where_clause(request.args)
    conn = get_db()
    rows = conn.execute(
        f"SELECT strftime('%Y-%m', order_date) as month, "
        f"SUM(sales) as total_sales, SUM(profit) as total_profit, COUNT(*) as order_count "
        f"FROM orders WHERE {where} GROUP BY month ORDER BY month", params
    ).fetchall()
    conn.close()
    return jsonify([dict(r) for r in rows])


@app.route("/api/category-sales")
@login_required
def api_category_sales():
    where, params = build_where_clause(request.args)
    conn = get_db()
    rows = conn.execute(
        f"SELECT category, SUM(sales) as total_sales, SUM(profit) as total_profit, COUNT(*) as order_count "
        f"FROM orders WHERE {where} GROUP BY category ORDER BY total_sales DESC", params
    ).fetchall()
    conn.close()
    return jsonify([dict(r) for r in rows])


@app.route("/api/region-sales")
@login_required
def api_region_sales():
    where, params = build_where_clause(request.args)
    conn = get_db()
    rows = conn.execute(
        f"SELECT region, SUM(sales) as total_sales, SUM(profit) as total_profit, COUNT(*) as order_count "
        f"FROM orders WHERE {where} GROUP BY region ORDER BY total_sales DESC", params
    ).fetchall()
    conn.close()
    return jsonify([dict(r) for r in rows])


@app.route("/api/state-sales")
@login_required
def api_state_sales():
    where, params = build_where_clause(request.args)
    conn = get_db()
    rows = conn.execute(
        f"SELECT state, SUM(sales) as total_sales, SUM(profit) as total_profit, COUNT(*) as order_count "
        f"FROM orders WHERE {where} GROUP BY state ORDER BY total_sales DESC", params
    ).fetchall()
    conn.close()
    return jsonify([dict(r) for r in rows])


@app.route("/api/customer-segment")
@login_required
def api_customer_segment():
    where, params = build_where_clause(request.args)
    conn = get_db()
    rows = conn.execute(
        f"SELECT c.segment, COUNT(DISTINCT o.customer_name) as customer_count, "
        f"SUM(o.sales) as total_sales, SUM(o.profit) as total_profit "
        f"FROM orders o JOIN customers c ON o.customer_name = c.customer_name "
        f"WHERE {where} GROUP BY c.segment ORDER BY total_sales DESC", params
    ).fetchall()
    conn.close()
    return jsonify([dict(r) for r in rows])


@app.route("/api/top-products")
@login_required
def api_top_products():
    where, params = build_where_clause(request.args)
    limit = request.args.get("limit", 10, type=int)
    order = request.args.get("order", "desc")
    sort_col = "total_sales" if order == "desc" else "total_sales"

    conn = get_db()
    rows = conn.execute(
        f"SELECT product_name, category, sub_category, "
        f"SUM(sales) as total_sales, SUM(profit) as total_profit, "
        f"SUM(quantity) as total_quantity, COUNT(*) as order_count "
        f"FROM orders WHERE {where} GROUP BY product_name "
        f"ORDER BY {sort_col} {'DESC' if order == 'desc' else 'ASC'} LIMIT ?",
        params + [limit]
    ).fetchall()
    conn.close()
    return jsonify([dict(r) for r in rows])


@app.route("/api/category-subcategory")
@login_required
def api_category_subcategory():
    where, params = build_where_clause(request.args)
    conn = get_db()
    rows = conn.execute(
        f"SELECT category, sub_category, SUM(sales) as total_sales, SUM(profit) as total_profit "
        f"FROM orders WHERE {where} GROUP BY category, sub_category ORDER BY total_sales DESC", params
    ).fetchall()
    conn.close()
    return jsonify([dict(r) for r in rows])


@app.route("/api/orders")
@login_required
def api_orders():
    where, params = build_where_clause(request.args)
    page = request.args.get("page", 1, type=int)
    per_page = request.args.get("per_page", 20, type=int)
    sort = request.args.get("sort", "order_date")
    direction = request.args.get("direction", "desc")
    search = request.args.get("search", "")

    allowed_sorts = {"order_date", "sales", "profit", "quantity", "customer_name", "product_name", "category", "region"}
    if sort not in allowed_sorts:
        sort = "order_date"
    if direction not in ("asc", "desc"):
        direction = "desc"

    count_params = list(params)
    conn = get_db()
    total = conn.execute(
        f"SELECT COUNT(*) FROM orders WHERE {where}", count_params
    ).fetchone()[0]

    offset = (page - 1) * per_page
    rows = conn.execute(
        f"SELECT * FROM orders WHERE {where} ORDER BY {sort} {direction} LIMIT ? OFFSET ?",
        params + [per_page, offset]
    ).fetchall()
    conn.close()

    return jsonify({
        "orders": [dict(r) for r in rows],
        "total": total,
        "page": page,
        "per_page": per_page,
        "pages": max(1, (total + per_page - 1) // per_page),
    })


@app.route("/api/products")
@login_required
def api_products():
    page = request.args.get("page", 1, type=int)
    per_page = request.args.get("per_page", 20, type=int)
    search = request.args.get("search", "")
    sort = request.args.get("sort", "product_name")
    direction = request.args.get("direction", "asc")

    allowed_sorts = {"product_name", "category", "sub_category", "unit_price"}
    if sort not in allowed_sorts:
        sort = "product_name"
    if direction not in ("asc", "desc"):
        direction = "asc"

    conditions = []
    params = []
    if search:
        conditions.append("product_name LIKE ?")
        params.append(f"%{search}%")
        or_cat = "OR category LIKE ?"
        conditions[-1] = f"(product_name LIKE ? OR category LIKE ? OR sub_category LIKE ?)"
        params = [f"%{search}%", f"%{search}%", f"%{search}%"]

    where = " AND ".join(conditions) if conditions else "1=1"

    conn = get_db()
    total = conn.execute(f"SELECT COUNT(*) FROM products WHERE {where}", params).fetchone()[0]

    offset = (page - 1) * per_page
    rows = conn.execute(
        f"SELECT p.*, "
        f"COALESCE(SUM(o.sales),0) as total_sales, "
        f"COALESCE(SUM(o.profit),0) as total_profit, "
        f"COALESCE(SUM(o.quantity),0) as total_sold, "
        f"COUNT(o.id) as order_count "
        f"FROM products p LEFT JOIN orders o ON p.product_name = o.product_name "
        f"WHERE {where} GROUP BY p.id ORDER BY p.{sort} {direction} LIMIT ? OFFSET ?",
        params + [per_page, offset]
    ).fetchall()
    conn.close()

    return jsonify({
        "products": [dict(r) for r in rows],
        "total": total,
        "page": page,
        "per_page": per_page,
        "pages": max(1, (total + per_page - 1) // per_page),
    })


@app.route("/api/customers")
@login_required
def api_customers():
    page = request.args.get("page", 1, type=int)
    per_page = request.args.get("per_page", 20, type=int)
    search = request.args.get("search", "")
    sort = request.args.get("sort", "customer_name")
    direction = request.args.get("direction", "asc")

    allowed_sorts = {"customer_name", "segment", "city", "state", "region"}
    if sort not in allowed_sorts:
        sort = "customer_name"
    if direction not in ("asc", "desc"):
        direction = "asc"

    conditions = []
    params = []
    if search:
        conditions.append(
            "(c.customer_name LIKE ? OR c.segment LIKE ? OR c.city LIKE ? OR c.state LIKE ?)"
        )
        params = [f"%{search}%"] * 4

    where = " AND ".join(conditions) if conditions else "1=1"

    conn = get_db()
    total = conn.execute(
        f"SELECT COUNT(*) FROM customers c WHERE {where}", params
    ).fetchone()[0]

    offset = (page - 1) * per_page
    rows = conn.execute(
        f"SELECT c.*, "
        f"COALESCE(SUM(o.sales),0) as total_spent, "
        f"COALESCE(SUM(o.profit),0) as total_profit, "
        f"COUNT(o.id) as order_count "
        f"FROM customers c LEFT JOIN orders o ON c.customer_name = o.customer_name "
        f"WHERE {where} GROUP BY c.id ORDER BY c.{sort} {direction} LIMIT ? OFFSET ?",
        params + [per_page, offset]
    ).fetchall()
    conn.close()

    return jsonify({
        "customers": [dict(r) for r in rows],
        "total": total,
        "page": page,
        "per_page": per_page,
        "pages": max(1, (total + per_page - 1) // per_page),
    })


@app.route("/api/analytics")
@login_required
def api_analytics():
    conn = get_db()

    highest_region = conn.execute(
        "SELECT region, SUM(sales) as total FROM orders GROUP BY region ORDER BY total DESC LIMIT 1"
    ).fetchone()

    highest_category = conn.execute(
        "SELECT category, SUM(profit) as total FROM orders GROUP BY category ORDER BY total DESC LIMIT 1"
    ).fetchone()

    best_product = conn.execute(
        "SELECT product_name, SUM(sales) as total FROM orders GROUP BY product_name ORDER BY total DESC LIMIT 1"
    ).fetchone()

    worst_product = conn.execute(
        "SELECT product_name, SUM(sales) as total FROM orders GROUP BY product_name ORDER BY total ASC LIMIT 1"
    ).fetchone()

    total_sales = conn.execute("SELECT COALESCE(SUM(sales),0) FROM orders").fetchone()[0]
    total_profit = conn.execute("SELECT COALESCE(SUM(profit),0) FROM orders").fetchone()[0]

    monthly_growth = conn.execute(
        "SELECT strftime('%Y-%m', order_date) as month, SUM(sales) as total "
        "FROM orders GROUP BY month ORDER BY month DESC LIMIT 2"
    ).fetchall()

    growth_rate = 0
    if len(monthly_growth) == 2 and monthly_growth[1]["total"] > 0:
        growth_rate = ((monthly_growth[0]["total"] - monthly_growth[1]["total"]) / monthly_growth[1]["total"]) * 100

    top_customer = conn.execute(
        "SELECT customer_name, SUM(sales) as total FROM orders "
        "GROUP BY customer_name ORDER BY total DESC LIMIT 1"
    ).fetchone()

    top_state = conn.execute(
        "SELECT state, SUM(sales) as total FROM orders GROUP BY state ORDER BY total DESC LIMIT 1"
    ).fetchone()

    profit_margin = (total_profit / total_sales * 100) if total_sales > 0 else 0

    avg_order_value = conn.execute("SELECT COALESCE(AVG(sales),0) FROM orders").fetchone()[0]

    low_margin_products = conn.execute(
        "SELECT product_name, SUM(sales) as total_sales, SUM(profit) as total_profit, "
        "ROUND(SUM(profit)*100.0/SUM(sales), 2) as margin "
        "FROM orders GROUP BY product_name HAVING SUM(sales) > 0 ORDER BY margin ASC LIMIT 5"
    ).fetchall()

    high_discount = conn.execute(
        "SELECT category, AVG(discount) as avg_discount, SUM(sales) as total_sales "
        "FROM orders GROUP BY category ORDER BY avg_discount DESC"
    ).fetchall()

    recommendations = []
    if profit_margin < 20:
        recommendations.append("Overall profit margin is below 20%. Consider reviewing pricing strategy and cost controls.")
    if growth_rate < 0:
        recommendations.append("Monthly sales are declining. Investigate market conditions and consider promotional campaigns.")
    for p in low_margin_products:
        if p["margin"] < 5:
            recommendations.append(
                f"'{p['product_name']}' has a very low margin ({p['margin']}%). Consider repricing or bundling."
            )
    for c in high_discount:
        if c["avg_discount"] > 0.2:
            recommendations.append(
                f"Average discount in '{c['category']}' is {round(c['avg_discount']*100, 1)}%. "
                f"Review discount policy to protect margins."
            )
    if not recommendations:
        recommendations.append("All KPIs are within healthy ranges. Keep monitoring for changes.")

    conn.close()

    return jsonify({
        "highest_region": dict(highest_region) if highest_region else None,
        "highest_category": dict(highest_category) if highest_category else None,
        "best_product": dict(best_product) if best_product else None,
        "worst_product": dict(worst_product) if worst_product else None,
        "monthly_growth_rate": round(growth_rate, 2),
        "profit_margin": round(profit_margin, 2),
        "avg_order_value": round(avg_order_value, 2),
        "top_customer": dict(top_customer) if top_customer else None,
        "top_state": dict(top_state) if top_state else None,
        "total_sales": round(total_sales, 2),
        "total_profit": round(total_profit, 2),
        "low_margin_products": [dict(p) for p in low_margin_products],
        "recommendations": recommendations,
    })


@app.route("/api/export/csv")
@login_required
def export_csv():
    where, params = build_where_clause(request.args)
    conn = get_db()
    rows = conn.execute(
        f"SELECT * FROM orders WHERE {where} ORDER BY order_date DESC", params
    ).fetchall()
    conn.close()

    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow(["ID", "Order Date", "Customer", "Product", "Category", "Sub Category",
                     "Region", "State", "City", "Quantity", "Sales", "Profit", "Discount"])
    for r in rows:
        writer.writerow([r["id"], r["order_date"], r["customer_name"], r["product_name"],
                         r["category"], r["sub_category"], r["region"], r["state"], r["city"],
                         r["quantity"], r["sales"], r["profit"], r["discount"]])

    return Response(
        output.getvalue(),
        mimetype="text/csv",
        headers={"Content-Disposition": "attachment;filename=orders_export.csv"}
    )


@app.route("/api/export/excel")
@login_required
def export_excel():
    if not HAS_OPENPYXL:
        return jsonify({"error": "openpyxl not installed"}), 500

    where, params = build_where_clause(request.args)
    conn = get_db()
    rows = conn.execute(
        f"SELECT * FROM orders WHERE {where} ORDER BY order_date DESC", params
    ).fetchall()
    conn.close()

    wb = Workbook()
    ws = wb.active
    ws.title = "Orders"

    headers = ["ID", "Order Date", "Customer", "Product", "Category", "Sub Category",
               "Region", "State", "City", "Quantity", "Sales", "Profit", "Discount"]
    ws.append(headers)

    for r in rows:
        ws.append([r["id"], r["order_date"], r["customer_name"], r["product_name"],
                    r["category"], r["sub_category"], r["region"], r["state"], r["city"],
                    r["quantity"], r["sales"], r["profit"], r["discount"]])

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    return Response(
        output.getvalue(),
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": "attachment;filename=orders_export.xlsx"}
    )


@app.route("/api/export/pdf")
@login_required
def export_pdf():
    if not HAS_REPORTLAB:
        return jsonify({"error": "reportlab not installed"}), 500

    where, params = build_where_clause(request.args)
    conn = get_db()
    rows = conn.execute(
        f"SELECT * FROM orders WHERE {where} ORDER BY order_date DESC LIMIT 500", params
    ).fetchall()
    conn.close()

    output = io.BytesIO()
    doc = SimpleDocTemplate(output, pagesize=landscape(letter))
    styles = getSampleStyleSheet()
    elements = []

    elements.append(Paragraph("Sales Dashboard - Orders Report", styles["Title"]))
    elements.append(Spacer(1, 20))
    elements.append(Paragraph(
        f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')} | Total Records: {len(rows)}",
        styles["Normal"]
    ))
    elements.append(Spacer(1, 20))

    table_data = [["ID", "Date", "Customer", "Product", "Region", "State", "Qty", "Sales", "Profit"]]
    for r in rows[:200]:
        table_data.append([
            str(r["id"]), r["order_date"], r["customer_name"][:15],
            r["product_name"][:20], r["region"], r["state"],
            str(r["quantity"]), f"${r['sales']:,.2f}", f"${r['profit']:,.2f}"
        ])

    table = Table(table_data, repeatRows=1)
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1a73e8")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTSIZE", (0, 0), (-1, 0), 8),
        ("FONTSIZE", (0, 1), (-1, -1), 7),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f0f4ff")]),
    ]))

    elements.append(table)
    doc.build(elements)

    return Response(
        output.getvalue(),
        mimetype="application/pdf",
        headers={"Content-Disposition": "attachment;filename=orders_report.pdf"}
    )


# ---------- Run ----------

if __name__ == "__main__":
    app.run(debug=True, port=5000)
