import requests
import json

SERVER_URL = "http://localhost:5000"


def test_process_tables():
    print("=" * 50)
    print("调用 GET /process-tables")
    print("=" * 50)

    response = requests.get(f"{SERVER_URL}/process-tables")

    print(f"状态码: {response.status_code}")

    if response.status_code != 200:
        print(f"错误: {response.text}")
        return []

    data = response.json()
    results = data.get("results", [])

    print(f"共返回 {len(results)} 个表格结果\n")

    for i, item in enumerate(results, 1):
        print(f"--- [{i}] id: {item['ids']} ---")
        for j, block in enumerate(item.get("blocks", [])):
            btype = block.get("type")
            if btype == "body":
                print(f"  [body]  title   : {block.get('content', '')}")
            elif btype == "table":
                style = block.get("style", {})
                content = block.get("content", [])
                print(f"  [table] rows={style.get('rows', '?')}  cols={style.get('cols', '?')}")
                for row in content:
                    print(f"          {row}")
        print()

    return results


def test_backfill_tables(results):
    print("=" * 50)
    print("调用 POST /backfill-tables")
    print("=" * 50)

    response = requests.post(
        f"{SERVER_URL}/backfill-tables",
        json={"results": results},
    )
    print(f"状态码: {response.status_code}")
    print(response.json())


if __name__ == "__main__":
    table_results = test_process_tables()
    if table_results:
        test_backfill_tables(table_results)
