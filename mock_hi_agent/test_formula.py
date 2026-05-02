import requests
import json

SERVER_URL = "http://localhost:5000"


def test_process_formulas():
    print("=" * 50)
    print("调用 GET /process-formulas")
    print("=" * 50)

    response = requests.get(f"{SERVER_URL}/process-formulas")

    print(f"状态码: {response.status_code}")

    if response.status_code != 200:
        print(f"错误: {response.text}")
        return []

    data = response.json()
    results = data.get("results", [])

    print(f"共返回 {len(results)} 条公式结果\n")

    for i, item in enumerate(results, 1):
        print(f"--- [{i}] id: {item['id']} ---")
        if item.get("text_before"):
            print(f"  text_before : {item['text_before']}")
        if item.get("label"):
            print(f"  label       : {item['label']}")
        if item.get("text_after"):
            print(f"  text_after  : {item['text_after']}")
        if item.get("error"):
            print(f"  [ERROR]     : {item['error']}")
        omath = item.get("omath", "")
        print(f"  omath 长度  : {len(omath)} chars")
        print(f"  omath 预览  : {omath[:120]}...")
        print()

    return results


def test_backfill_formulas(results):
    print("=" * 50)
    print("调用 POST /backfill-formulas")
    print("=" * 50)

    response = requests.post(
        f"{SERVER_URL}/backfill-formulas",
        json={"results": results},
    )
    print(f"状态码: {response.status_code}")
    print(response.json())


if __name__ == "__main__":
    formula_results = test_process_formulas()
    if formula_results:
        test_backfill_formulas(formula_results)
