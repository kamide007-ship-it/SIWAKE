# 銀行明細 → Excel自動変換システム

## 起動方法
```bash
pip install -r requirements.txt
python app.py
```
ブラウザで http://localhost:7860 を開く

## 対応CSV形式
- 列: 日付,摘要,入金金額,出金金額,残高,メモ
- 日付形式: YYYYMMDD / YYYY/MM/DD / YYYY-MM-DD
- エンコード: Shift-JIS / UTF-8 自動判定

## 出力Excelの構成
- 📊年間サマリー: 月別集計表
- 各月シート: 月別明細（前月繰越〜合計まで）
- 🏥経営健康診断: スコア・改善ポイント
