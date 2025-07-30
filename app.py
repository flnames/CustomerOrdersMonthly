from flask import Flask, request, jsonify, abort, url_for
import pandas as pd
import os

EXCEL_FILE = "CustomerOrders.xlsx"
PER_PAGE = 5000

app = Flask(__name__)

# 📦 Load Excel data once at startup
try:
    if not os.path.isfile(EXCEL_FILE):
        print(f"[ERROR] '{EXCEL_FILE}' not found.")
        data = []
    else:
        df = pd.read_excel(EXCEL_FILE)
        data = df.to_dict(orient="records")
        print(f"[INFO] Loaded {len(data)} records from '{EXCEL_FILE}'")
except Exception as e:
    print(f"[ERROR] Failed to read Excel file: {e}")
    data = []

# 🟢 Health check
@app.route("/")
def index():
    return "OK", 200

# 📊 CustomerOrders endpoint with optional month+year filtering
@app.route("/CustomerOrders", strict_slashes=False)
def get_data():
    # Parse page
    try:
        page = int(request.args.get("page", 1))
        if page <= 0:
            raise ValueError
    except ValueError:
        return jsonify({"error": "Invalid 'page' parameter"}), 400

    # Get filter parameters
    month = request.args.get("month")
    year = request.args.get("year")

    # Enforce month+year must be provided together
    if (month and not year) or (year and not month):
        return jsonify({"error": "Both 'month' and 'year' parameters are required together"}), 400

    # Filter if both provided
    filtered_data = data
    if month and year:
        try:
            month = int(month)
            year = int(year)

            if not (1 <= month <= 12):
                raise ValueError("Month must be between 1 and 12.")

            df_filtered = pd.DataFrame(data)
            df_filtered["OrderDate"] = pd.to_datetime(df_filtered["OrderDate"], errors='coerce')

            df_filtered = df_filtered[
                (df_filtered["OrderDate"].dt.month == month) &
                (df_filtered["OrderDate"].dt.year == year)
            ]

            filtered_data = df_filtered.to_dict(orient="records")
        except Exception as e:
            return jsonify({"error": f"Invalid month/year filter: {str(e)}"}), 400

    # Pagination
    total_rows = len(filtered_data)
    start = (page - 1) * PER_PAGE
    end = start + PER_PAGE
    paginated = filtered_data[start:end]
    has_more = end < total_rows

    return jsonify({
        "page": page,
        "per_page": PER_PAGE,
        "total_rows": total_rows,
        "has_more": has_more,
        "next_page": url_for('get_data', page=page + 1, month=month, year=year, _external=True) if has_more and month and year else None,
        "data": paginated
    })


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))  # Render compatibility
    app.run(host="0.0.0.0", port=port)
