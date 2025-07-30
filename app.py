from flask import Flask, request, jsonify, abort, url_for
import pandas as pd
import os
from datetime import datetime, timedelta

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

# 📊 CustomerOrders endpoint with date range logic
@app.route("/CustomerOrders", strict_slashes=False)
def get_data():
    # Parse page
    try:
        page = int(request.args.get("page", 1))
        if page <= 0:
            raise ValueError
    except ValueError:
        return jsonify({"error": "Invalid 'page' parameter"}), 400

    last_load = request.args.get("last_load")
    end_date_str = request.args.get("end_date")

    # Ensure both are provided
    if (last_load and not end_date_str) or (end_date_str and not last_load):
        return jsonify({"error": "Both 'last_load' and 'end_date' must be provided"}), 400

    filtered_data = data

    if last_load and end_date_str:
        try:
            # Compute range
            last_loaded_date = pd.to_datetime(last_load)
            start_date = last_loaded_date + timedelta(days=1)
            end_date = pd.to_datetime(end_date_str)

            # Filter data
            df_filtered = pd.DataFrame(data)
            df_filtered["OrderDate"] = pd.to_datetime(df_filtered["OrderDate"], errors='coerce')

            df_filtered = df_filtered[
                (df_filtered["OrderDate"] >= start_date) &
                (df_filtered["OrderDate"] <= end_date)
            ]

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
        "next_page": url_for('get_data', page=page + 1, last_load=last_load, end_date=end_date_str, _external=True) if has_more and last_load and end_date_str else None,
        "data": paginated
    })


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))  # Render compatibility
    app.run(host="0.0.0.0", port=port)
