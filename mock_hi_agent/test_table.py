import requests


def handler(params):
    server_url = params.get("server_url", "http://localhost:5000")

    # Step 1: process-tables
    resp = requests.get(f"{server_url}/process-tables")
    if resp.status_code != 200:
        return {"error": resp.text, "process_status": resp.status_code}

    results = resp.json().get("results", [])

    summary = []
    for item in results:
        blocks_info = []
        for block in item.get("blocks", []):
            btype = block.get("type")
            if btype == "body":
                blocks_info.append({"type": "body", "title": block.get("content", "")})
            elif btype == "table":
                style = block.get("style", {})
                blocks_info.append({
                    "type": "table",
                    "rows": style.get("rows"),
                    "cols": style.get("cols"),
                    "content": block.get("content", []),
                })
        summary.append({"ids": item.get("ids"), "blocks": blocks_info})

    # Step 2: backfill-tables
    backfill = None
    if results:
        bf_resp = requests.post(
            f"{server_url}/backfill-tables",
            json={"results": results},
        )
        backfill = bf_resp.json()

    return {
        "table_count": len(results),
        "tables":      summary,
        "backfill":    backfill,
    }
