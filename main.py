        return {"success": False, "type": "unknown",
                "message": parsed.get("reason", "Could not understand that message.")}

    if parsed.get("type") == "query":
        result = InventoryService.answer_query(wb, parsed)
        return {"success": True, **result}

    # movement
    try:
        result = InventoryService.apply_movement(wb, parsed, data.message, data.entered_by)
    except ValueError as e:
        return {"success": False, "type": "movement_rejected", "message": str(e)}

    try:
        ExcelService.save(wb, f"{result['action']}: {data.message[:60]} ({data.entered_by})")
    except Exception as e:
        raise HTTPException(500, f"Could not save the workbook to GitHub: {e}")

    return {"success": True, "type": "movement", **result}


@app.post("/api/undo")
def undo():
    wb = ExcelService.load_live_workbook()
    try:
        result = ExcelService.undo_last(wb)
    except ValueError as e:
        raise HTTPException(400, str(e))
    try:
        ExcelService.save(wb, "Undo last transaction")
    except Exception as e:
        raise HTTPException(500, f"Could not save the workbook to GitHub: {e}")
    return {"success": True, "reversed": result}


@app.get("/api/transactions")
def transactions(limit: int = 50):
    wb = ExcelService.load_live_workbook()
    return {"transactions": ExcelService.get_log_rows(wb, limit)}


@app.get("/api/products")
def products():
    """Current stock, grouped by product then location - powers the Reports table."""
    wb = ExcelService.load_live_workbook()
    summary = ExcelService.stock_summary(wb)
    out = []
    batches = ExcelService.all_batches(wb)
    for product, locs in sorted(summary.items()):
        serials = []
        for b in batches:
            if str(b["Product"]).strip().lower() == product.strip().lower() and (b["Qty Remaining"] or 0) > 0:
                sn = str(b.get("Serial Number") or "").strip()
                if sn:
                    serials.append({"serial_number": sn, "location": b["Location"], "quantity": b["Qty Remaining"]})
        out.append({"product": product, "total": round(sum(locs.values()), 4),
                    "locations": [{"location": l, "quantity": q} for l, q in sorted(locs.items())],
                    "serial_numbers": serials})
    return {"products": out}


@app.get("/api/aging-report")
def aging_report(threshold_days: Optional[int] = None):
    wb = ExcelService.load_live_workbook()
    rows = ExcelService.aging_batches(wb, threshold_days or AGING_THRESHOLD_DAYS)
    return {"threshold_days": threshold_days or AGING_THRESHOLD_DAYS, "batches": rows}


@app.get("/api/stock-by-location")
def stock_by_location():
    wb = ExcelService.load_live_workbook()
    summary = ExcelService.stock_summary(wb)
    totals: Dict[str, float] = {}
    for product, locs in summary.items():
        for loc, q in locs.items():
            totals[loc] = round(totals.get(loc, 0) + q, 4)
    return {"locations": [{"location": l, "quantity": q} for l, q in sorted(totals.items(), key=lambda x: -x[1])]}


@app.get("/api/dashboard-summary")
def dashboard_summary():
    wb = ExcelService.load_live_workbook()
    summary = ExcelService.stock_summary(wb)
    all_batches = [b for b in ExcelService.all_batches(wb) if (b["Qty Remaining"] or 0) > 0]
    aging = ExcelService.aging_batches(wb)
    total_qty = round(sum(b["Qty Remaining"] for b in all_batches), 4)
    return {
        "summary": {
            "total_products": len(summary),
            "total_locations": len(ExcelService.known_locations(wb)),
            "total_quantity": total_qty,
            "active_batches": len(all_batches),
            "aging_alerts": len(aging),
            "aging_threshold_days": AGING_THRESHOLD_DAYS,
        }
    }


@app.get("/api/download")
def download_excel():
    try:
        _, raw = _get_file_meta()
        if raw is None:
            wb = ExcelService.new_workbook()
            buf = io.BytesIO()
            wb.save(buf)
            raw = buf.getvalue()
    except Exception as e:
        raise HTTPException(500, f"Could not read the workbook: {e}")
    return StreamingResponse(
        io.BytesIO(raw),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": "attachment; filename=inventory_report.xlsx"},
    )


@app.get("/api/download-url")
def download_url_route():
    return {"url": get_download_url()}
