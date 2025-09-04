# presence_alert

ネットワーク接続（OmadaコントローラのWebhook）をトリガに、スタッフの在室・出退勤を自動記録し、Discordへ通知、Google スプレッドシートへ保存・集計する仕組みです。Google Apps Script（GAS）で構成します。

## 概要 / アーキテクチャ
- Omada → Webhook → `gas1.gs`（受信・判定）
  - 端末MACと状態（ONLINE/OFFLINE）を解析、`data`/`log` シートへ保存
  - `facility` シートでサイト名を表示名に解決
  - Discordへ通知、出勤簿用 `gas2` へイベント転送
- `gas2.gs`（出勤簿・集計）
  - `出勤簿` に出社/退社、`raw_data` に生ログ、`月次集計` を生成
  - 退勤後8時間以内のONLINEは「外出復帰」とみなし退勤取消、最終退勤時に外出分を控除

## スプレッドシート構成（主要）
- `mac`：A=MAC, B=デバイス名, C=gas2 URL, D=ユーザー名, E=Discord Webhook
- `facility`：A=Site名, B=表示名（未登録時はSite名を使用）
- `data`：タイムスタンプ/MAC/表示名/状態/IP/施設名/元JSON/判定
- `出勤簿`：日付/ユーザー/施設/出社/退社/休憩/実働/備考

## デプロイ
1. `gas1.gs` を新規GASに配置→「デプロイ > ウェブアプリ」URLをOmadaに設定
2. `gas2.gs` を別GASに配置→同様にデプロイし、そのURLを `mac` シートC列へ

## 開発・テスト
- `gas1.gs`：`testSetup()` で初期化、`testWebhook()` でONLINEを擬似送信
- `gas2.gs`：`setupAttendanceSheets()` で初期化、`reprocessRawData()` で時系列再構築
- 並行実行対策として `LockService` を使用。タイムスタンプは `timestampMs`（ms）優先

## 運用メモ
- 複数端末: 他端末がONLINEの場合はOFFLINEを退勤扱いにしない
- 外出復帰: 退勤から8時間以内のONLINEで退勤取消、備考に `[OUTING_MINUTES=…]` 追加、実働から控除
- 月次: 当月のみ再生成。他月は保持

## セキュリティ
- Webhookやgas2 URLはリポジトリに直書きせず、`mac` シートに保持
- Discord通知のスクショ共有時はMAC/IPを秘匿
