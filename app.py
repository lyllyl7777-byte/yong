import io
import os
from flask import Flask, render_template, request, send_file
from generator import generate_contract

app = Flask(__name__)


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/generate", methods=["POST"])
def generate():
    data = {
        # 사업주 정보
        "owner_name": request.form.get("owner_name", ""),
        "business_reg_no": request.form.get("business_reg_no", ""),
        # 근로자 정보
        "worker_name": request.form.get("worker_name", ""),
        "worker_id_no": request.form.get("worker_id_no", ""),
        "worker_address": request.form.get("worker_address", ""),
        "worker_phone": request.form.get("worker_phone", ""),
        "worker_relation": request.form.get("worker_relation", ""),
        # 제1조
        "job_duties": request.form.get("job_duties", ""),
        # 제2조
        "contract_start": request.form.get("contract_start", ""),
        "contract_end": request.form.get("contract_end", ""),
        "no_fixed_term": "no_fixed_term" in request.form,
        "probation_none": request.form.get("probation") == "none",
        "probation_months": request.form.get("probation_months", ""),
        "probation_wage": request.form.get("probation_wage", ""),
        # 제3조
        "work_start": request.form.get("work_start", ""),
        "work_end": request.form.get("work_end", ""),
        "daily_hours": request.form.get("daily_hours", ""),
        "weekly_hours": request.form.get("weekly_hours", ""),
        "break_start": request.form.get("break_start", ""),
        "break_end": request.form.get("break_end", ""),
        "work_days_per_week": request.form.get("work_days_per_week", ""),
        "weekly_holiday": request.form.get("weekly_holiday", ""),
        # 제4조
        "basic_wage": request.form.get("basic_wage", ""),
        "hourly_wage": request.form.get("hourly_wage", ""),
        "meal_allowance": request.form.get("meal_allowance", ""),
        "other_allowance": request.form.get("other_allowance", ""),
        "total_wage": request.form.get("total_wage", ""),
        "pay_day": request.form.get("pay_day", ""),
        "pay_direct": request.form.get("pay_method") == "direct",
        "pay_bank": request.form.get("pay_method") == "bank",
        "bank_name": request.form.get("bank_name", ""),
        "account_no": request.form.get("account_no", ""),
        # 제6조 사회보험
        "ins_employment": "ins_employment" in request.form,
        "ins_industrial": "ins_industrial" in request.form,
        "ins_pension": "ins_pension" in request.form,
        "ins_health": "ins_health" in request.form,
        # 제7조
        "severance_yes": request.form.get("severance") == "yes",
        "severance_no": request.form.get("severance") == "no",
        "special_terms": request.form.get("special_terms", ""),
        # 계약체결일
        "contract_year": request.form.get("contract_year", ""),
        "contract_month": request.form.get("contract_month", ""),
        "contract_day": request.form.get("contract_day", ""),
    }

    docx_bytes = generate_contract(data)

    worker = data["worker_name"] or "근로자"
    filename = f"근로계약서_{worker}.docx"

    return send_file(
        io.BytesIO(docx_bytes),
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        as_attachment=True,
        download_name=filename,
    )


if __name__ == "__main__":
    app.run(debug=True, port=5000)
