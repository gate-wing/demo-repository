import pandas as pd
import re
import os
from pathlib import Path

# 東京23区のリスト
WARDS_23 = [
    '千代田区', '中央区', '港区', '新宿区', '文京区', '台東区',
    '墨田区', '江東区', '品川区', '目黒区', '大田区', '世田谷区',
    '渋谷区', '中野区', '杉並区', '豊島区', '北区', '荒川区',
    '板橋区', '練馬区', '足立区', '葛飾区', '江戸川区'
]


def extract_area(address):
    """住所からエリア（都道府県）を抽出する"""
    if not address or pd.isna(address):
        return '不明'

    address = str(address)

    match = re.search(r'(東京都|北海道|(?:大阪|京都)府|\S{2,3}県)', address)
    if not match:
        return '不明'

    pref = match.group(1)

    if pref == '東京都':
        for ward in WARDS_23:
            if ward in address:
                return '東京都（23区）'
        return '東京都（23区以外）'

    return pref


def load_csv(filepath):
    """CSVファイルを読み込む（2行目の説明行をスキップ）"""
    df = pd.read_csv(filepath, encoding='utf-8-sig', skiprows=[1], dtype=str)
    return df


def main():
    data_dir = Path('data')
    output_file = 'analysis_result.xlsx'

    if not data_dir.exists():
        print('「data」フォルダが見つかりません。dataフォルダを作成してCSVファイルを入れてください。')
        return

    csv_files = sorted(data_dir.glob('*.csv'))
    if not csv_files:
        print('dataフォルダにCSVファイルがありません。')
        return

    all_data = []

    for csv_file in csv_files:
        month = csv_file.stem
        print(f'{month} を読み込み中...')

        try:
            df = load_csv(csv_file)
            df['月'] = month
            all_data.append(df)
        except Exception as e:
            print(f'{csv_file.name} の読み込みに失敗: {e}')

    if not all_data:
        print('読み込めるデータがありませんでした。')
        return

    combined = pd.concat(all_data, ignore_index=True)

    num_cols = [
        'Google 検索 - モバイル', 'Google 検索 - パソコン',
        'Google マップ - モバイル', 'Google マップ - パソコン',
        '通話', 'メッセージ', '予約', 'ルート', 'ウェブサイトのクリック',
        '料理の注文', 'フードメニューのクリック'
    ]

    for col in num_cols:
        if col in combined.columns:
            combined[col] = pd.to_numeric(combined[col], errors='coerce').fillna(0)

    combined['表示回数'] = (
        combined.get('Google 検索 - モバイル', 0) +
        combined.get('Google 検索 - パソコン', 0) +
        combined.get('Google マップ - モバイル', 0) +
        combined.get('Google マップ - パソコン', 0)
    )

    combined['アクション数'] = (
        combined.get('通話', 0) +
        combined.get('メッセージ', 0) +
        combined.get('予約', 0) +
        combined.get('ルート', 0) +
        combined.get('ウェブサイトのクリック', 0) +
        combined.get('料理の注文', 0) +
        combined.get('フードメニューのクリック', 0)
    )

    combined['エリア'] = combined['住所'].apply(extract_area)

    combined['アクション率(%)'] = (
        combined['アクション数'] / combined['表示回数'].replace(0, float('nan')) * 100
    ).round(2)

    # 月別集計
    monthly = combined.groupby('月').agg(
        表示回数合計=('表示回数', 'sum'),
        アクション数合計=('アクション数', 'sum'),
        店舗数=('ビジネス名', 'count')
    ).reset_index()
    monthly['アクション率(%)'] = (
        monthly['アクション数合計'] / monthly['表示回数合計'] * 100
    ).round(2)

    # エリア別集計
    area = combined.groupby('エリア').agg(
        表示回数合計=('表示回数', 'sum'),
        アクション数合計=('アクション数', 'sum'),
        店舗数=('ビジネス名', 'count')
    ).reset_index().sort_values('表示回数合計', ascending=False)
    area['アクション率(%)'] = (
        area['アクション数合計'] / area['表示回数合計'] * 100
    ).round(2)

    # エリア×月別集計
    area_monthly = combined.groupby(['エリア', '月']).agg(
        表示回数合計=('表示回数', 'sum'),
        アクション数合計=('アクション数', 'sum')
    ).reset_index()
    area_monthly['アクション率(%)'] = (
        area_monthly['アクション数合計'] / area_monthly['表示回数合計'] * 100
    ).round(2)

    # 相関分析
    correlation = combined[['表示回数', 'アクション数']].corr()

    # Excelに出力
    print(f'\n結果を {output_file} に出力中...')
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        monthly.to_excel(writer, sheet_name='月別集計', index=False)
        area.to_excel(writer, sheet_name='エリア別集計', index=False)
        area_monthly.to_excel(writer, sheet_name='エリア×月別集計', index=False)
        correlation.to_excel(writer, sheet_name='相関分析')
        combined[['月', 'エリア', 'ビジネス名', '住所', '表示回数', 'アクション数', 'アクション率(%)']].to_excel(
            writer, sheet_name='全データ', index=False
        )

    print(f'\n完了！{output_file} を開いて結果を確認してください。')
    print(f'\n概要:')
    print(f'  分析期間: {combined["月"].min()} 〜 {combined["月"].max()}')
    print(f'  総店舗数: {combined["ビジネス名"].nunique()} 店舗')
    print(f'  エリア数: {combined["エリア"].nunique()} エリア')


if __name__ == '__main__':
    main()
