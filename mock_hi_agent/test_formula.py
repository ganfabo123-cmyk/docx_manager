import requests


def handler(params):
    server_url = params.get("server_url", "http://localhost:5000")

    # Step 1: process-formulas
    resp = requests.get(f"{server_url}/process-formulas")
    if resp.status_code != 200:
        return {"error": resp.text, "process_status": resp.status_code}

    results = resp.json().get("results", [])

    summary = [
        {
            "id":          item.get("id"),
            "text_before": item.get("text_before", ""),
            "label":       item.get("label", ""),
            "text_after":  item.get("text_after", ""),
            "omath_len":   len(item.get("omath", "")),
            "error":       item.get("error"),
        }
        for item in results
    ]

    # Step 2: backfill-formulas
    backfill = None
    if results:
        bf_resp = requests.post(
            f"{server_url}/backfill-formulas",
            json={"results": results},
        )
        backfill = bf_resp.json()

    return {
        "formula_count": len(results),
        "formulas":      summary,
        "backfill":      backfill,
    }

