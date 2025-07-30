from flask import Flask, request, jsonify, abort, url_for
import pandas as pd
import os
from datetime import datetime

EXCEL_FILE = "CustomerOrders.xlsx"
PER_PAGE = 500

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

# 📊 CustomerOrders endpoint with start_date & end_date filtering
@app.route("/CustomerOrders", strict_slashes=False)
def get_data():
    # Parse page
    try:
        page = int(request.args.get("page", 1))
        if page <= 0:
            raise ValueError
    except ValueError:
        return jsonify({"error": "Invalid 'page' parameter"}), 400

    start_date_str = request.args.get("start_date")
    end_date_str = request.args.get("end_date")

    filtered_data = data

    try:
        df_filtered = pd.DataFrame(data)
        df_filtered["OrderDate"] = pd.to_datetime(df_filtered["OrderDate"], errors='coerce')

        if start_date_str and not end_date_str:
            return jsonify({"error": "'end_date' is required when 'start_date' is provided"}), 400

        elif start_date_str and end_date_str:
            start_date = pd.to_datetime(start_date_str)
            end_date = pd.to_datetime(end_date_str)
            df_filtered = df_filtered[
                (df_filtered["OrderDate"] >= start_date) &
                (df_filtered["OrderDate"] <= end_date)
            ]

        elif not start_date_str and end_date_str:
            end_date = pd.to_datetime(end_date_str)
            df_filtered = df_filtered[df_filtered["OrderDate"] <= end_date]

        # else: no filters → return all

        filtered_data = df_filtered.to_dict(orient="records")

    except Exception as e:
        return jsonify({"error": f"Invalid date format: {str(e)}"}), 400

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
        "next_page": url_for('get_data', page=page + 1, start_date=start_date_str, end_date=end_date_str, _external=True) if has_more else None,
        "data": paginated
    })


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))  # Render compatibility
    app.run(host="0.0.0.0", port=port)
