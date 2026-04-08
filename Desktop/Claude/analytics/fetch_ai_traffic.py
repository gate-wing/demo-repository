"""
AI参照トラフィック横断データ取得スクリプト

「株式会社ぐっとくるダイニング」の全GA4プロパティを自動検出し、
ChatGPT / Gemini / Perplexity などAIエンジンからの参照セッション数を
月別・プロパティ別に取得して ai_traffic_data.json に保存します。

実行方法:
  python -X utf8 fetch_ai_traffic.py
"""

import sys, os, json
if sys.platform == "win32":
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

import yaml
import requests
from datetime import datetime, date
from dateutil.relativedelta import relativedelta
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials

SCOPES = [
    "https://www.googleapis.com/auth/analytics.readonly",
]

BASE_DIR = os.path.dirname(os.path.abspath(__file__))


# ── 設定読み込み ───────────────────────────────────────────────
def load_config():
    path = os.path.join(BASE_DIR, "ai_traffic_config.yaml")
    with open(path, encoding="utf-8") as f:
        return yaml.safe_load(f)


# ── Google 認証 ────────────────────────────────────────────────
def get_credentials(cfg):
    creds_path = os.path.join(BASE_DIR, cfg["google_auth"]["credentials_json_path"].lstrip("./"))
    token_path  = os.path.join(BASE_DIR, cfg["google_auth"]["token_json_path"].lstrip("./"))

    creds = None
    if os.path.exists(token_path):
        creds = Credentials.from_authorized_user_file(token_path, SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(creds_path, SCOPES)
            creds = flow.run_local_server(port=0)
        with open(token_path, "w") as f:
            f.write(creds.to_json())
    return creds


def auth_headers(creds):
    return {"Authorization": f"Bearer {creds.token}"}


# ── GA4 Admin API: アカウント・プロパティ取得 ──────────────────
def list_ga4_accounts(headers):
    url = "https://analyticsadmin.googleapis.com/v1beta/accounts"
    res = requests.get(url, headers=headers)
    if res.status_code != 200:
        print(f"  [GA4 accounts error] {res.status_code}: {res.text[:300]}")
        return []
    return res.json().get("accounts", [])


def list_ga4_properties(account_name, headers):
    url = "https://analyticsadmin.googleapis.com/v1beta/properties"
    res = requests.get(url, headers=headers, params={"filter": f"parent:{account_name}"})
    if res.status_code != 200:
        print(f"  [GA4 properties error] {res.status_code}: {res.text[:300]}")
        return []
    return res.json().get("properties", [])


def find_target_account(cfg, headers):
    """対象アカウント名（部分一致）でGA4アカウントを検索する"""
    target = cfg["target_account_name"]
    accounts = list_ga4_accounts(headers)
    matched = [a for a in accounts if target in a.get("displayName", "")]
    if not matched:
        print(f"  [警告] アカウント名 '{target}' が見つかりませんでした。")
        print(f"  利用可能なアカウント: {[a.get('displayName') for a in accounts]}")
    return matched


# ── 分析期間の計算 ─────────────────────────────────────────────
def get_month_ranges(months):
    """直近 N ヶ月分の (start, end) リストを返す（降順）"""
    today = date.today()
    # 当月を含まず、直近の完了済み月から遡る
    end_month = today.replace(day=1) - relativedelta(days=1)  # 先月末
    ranges = []
    for i in range(months):
        m_start = end_month.replace(day=1) - relativedelta(months=i)
        m_end   = (m_start + relativedelta(months=1)) - relativedelta(days=1)
        ranges.append({
            "start": m_start.strftime("%Y-%m-%d"),
            "end":   m_end.strftime("%Y-%m-%d"),
            "label": m_start.strftime("%Y-%m"),
        })
    return list(reversed(ranges))  # 古い順


# ── フィルタ条件のビルド ───────────────────────────────────────
def build_ai_filter(ai_sources):
    """ai_sources設定からGA4 dimensionFilter (orGroup) を構築する"""
    expressions = []
    for source in ai_sources:
        for domain in source["domains"]:
            expressions.append({
                "filter": {
                    "fieldName": "sessionSource",
                    "stringFilter": {
                        "matchType": "CONTAINS",
                        "value": domain,
                        "caseSensitive": False,
                    }
                }
            })
    return {"orGroup": {"expressions": expressions}}


# ── GA4 Data API: AI参照トラフィック取得 ──────────────────────
def fetch_ai_sessions_for_month(property_id, month_range, ai_filter, headers):
    """1プロパティ・1ヶ月分のAI参照セッションを取得する"""
    url = f"https://analyticsdata.googleapis.com/v1beta/properties/{property_id}:runReport"
    body = {
        "dateRanges": [{"startDate": month_range["start"], "endDate": month_range["end"]}],
        "dimensions": [{"name": "sessionSource"}],
        "metrics":    [{"name": "sessions"}, {"name": "activeUsers"}],
        "dimensionFilter": ai_filter,
        "orderBys": [{"metric": {"metricName": "sessions"}, "desc": True}],
        "limit": 100,
    }
    res = requests.post(url, headers=headers, json=body)
    if res.status_code != 200:
        print(f"    [GA4 Data error] {res.status_code}: {res.text[:200]}")
        return []

    rows = []
    for row in res.json().get("rows", []):
        rows.append({
            "source":       row["dimensionValues"][0]["value"],
            "sessions":     int(row["metricValues"][0]["value"]),
            "active_users": int(row["metricValues"][1]["value"]),
        })
    return rows


def classify_source(source, ai_sources):
    """参照元ドメインをAIエンジン名に分類する"""
    for engine in ai_sources:
        for domain in engine["domains"]:
            if domain.lower() in source.lower():
                return engine["name"]
    return "その他"


# ── メイン ────────────────────────────────────────────────────
def main():
    print("=== AI参照トラフィック 横断データ取得 ===\n")

    cfg   = load_config()
    creds = get_credentials(cfg)
    hdrs  = auth_headers(creds)

    # ① 対象アカウント検索
    print(f"▶ 対象アカウント検索: '{cfg['target_account_name']}'")
    target_accounts = find_target_account(cfg, hdrs)
    if not target_accounts:
        print("  対象アカウントが見つかりません。ai_traffic_config.yaml を確認してください。")
        return

    # ② 全プロパティ取得
    all_properties = []
    for acc in target_accounts:
        print(f"  アカウント: {acc.get('displayName')} ({acc['name']})")
        props = list_ga4_properties(acc["name"], hdrs)
        for p in props:
            prop_id = p["name"].replace("properties/", "")
            all_properties.append({
                "property_id":   prop_id,
                "display_name":  p.get("displayName", ""),
                "account_name":  acc.get("displayName", ""),
            })
            print(f"    プロパティ: {p.get('displayName')} (ID: {prop_id})")

    if not all_properties:
        print("  プロパティが見つかりません。")
        return
    print(f"\n  合計 {len(all_properties)} プロパティを対象とします。\n")

    # ③ 分析期間の計算
    months = cfg["period"]["months"]
    month_ranges = get_month_ranges(months)
    print(f"▶ 分析期間: {month_ranges[0]['label']} 〜 {month_ranges[-1]['label']} ({months}ヶ月)\n")

    # ④ AI参照フィルタ構築
    ai_filter = build_ai_filter(cfg["ai_sources"])
    ai_source_names = [s["name"] for s in cfg["ai_sources"]]

    # ⑤ プロパティ × 月別でデータ取得
    result_properties = []
    prop_count = len(all_properties)

    for prop_idx, prop in enumerate(all_properties, 1):
        pid  = prop["property_id"]
        name = prop["display_name"]
        print(f"▶ [{prop_idx}/{prop_count}] {name}")

        monthly_data = []
        for mr in month_ranges:
            print(f"  {mr['label']}: 取得中...", end="", flush=True)
            rows = fetch_ai_sessions_for_month(pid, mr, ai_filter, hdrs)

            # AIエンジン別に集計
            by_engine = {engine: {"sessions": 0, "active_users": 0} for engine in ai_source_names}
            total_sessions = 0
            for row in rows:
                engine = classify_source(row["source"], cfg["ai_sources"])
                if engine in by_engine:
                    by_engine[engine]["sessions"]     += row["sessions"]
                    by_engine[engine]["active_users"] += row["active_users"]
                total_sessions += row["sessions"]

            monthly_data.append({
                "month":          mr["label"],
                "start":          mr["start"],
                "end":            mr["end"],
                "total_sessions": total_sessions,
                "by_engine":      by_engine,
                "raw_rows":       rows,
            })
            print(f"\r  {mr['label']}: AI参照 {total_sessions:,} セッション")

        result_properties.append({
            "property_id":  pid,
            "display_name": name,
            "account_name": prop["account_name"],
            "monthly_data": monthly_data,
        })

    # ⑥ JSON保存
    output = {
        "fetched_at":    datetime.now().isoformat(),
        "target_account": cfg["target_account_name"],
        "period_months": months,
        "ai_engines":    ai_source_names,
        "properties":    result_properties,
    }

    out_path = os.path.join(BASE_DIR, "ai_traffic_data.json")
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, indent=2)

    print(f"\n✅ データを保存しました: ai_traffic_data.json")
    print(f"   プロパティ数: {len(result_properties)}")
    print("次のステップ: python -X utf8 export_ai_report.py")


if __name__ == "__main__":
    main()
