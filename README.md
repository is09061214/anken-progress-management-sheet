# 案件進捗管理シート（anken-progress-management-sheet）

「iMuseLLC 案件管理表」スプレッドシートを中心とした、案件進捗管理の仕組み一式をまとめたリポジトリです。

スプレッドシートが頭脳（判定・データ管理）、ダッシュボードが顔（見える化）、という役割分担で、この2つは運命共同体として一緒に管理します。

## フォルダ構成

```
anken-progress-management-sheet/
├── dashboard/   進捗信号ダッシュボード（Cloud Run で稼働する Web アプリ）
├── sheet/       スプレッドシート側の道具箱（Apps Script・図解ガイド・手順書）
└── .github/     自動化（自動デプロイ・Discord 通知）
```

### dashboard/ — 進捗信号ダッシュボード

スプレッドシートの「ダッシュボード」シート（判定済みの表）をそのまま読み取り、赤／黄／青／灰の信号で表示する Web アプリ。詳細・起動方法は [dashboard/README.md](dashboard/README.md) を参照。

- 本番: Cloud Run サービス `youtube-progress-dashboard`（asia-northeast1）
- **`main` にプッシュすると自動でデプロイされる**（`dashboard/` 配下の変更時のみ発火）

### sheet/ — スプレッドシート側の道具箱

- `*.gs` … Apps Script（シート生成・数式修復・カレンダー自動追加・リマインダー等）。
  実体はスプレッドシート側の Apps Script プロジェクトにあり、ここはそのマスターコピー
- `*-guide/`, `v2-kaizen/`, `v3-full-guide/` … チーム向け図解ガイド（HTML、surge.sh で公開）
- `セットアップ手順_*.md` … コピペで進められる導入手順書

### .github/workflows/ — 自動化

| ワークフロー | 内容 |
|---|---|
| `cloud_run_deploy.yml` | `dashboard/` 変更時に Cloud Run へ自動デプロイ |
| `discord_morning.yml` | 毎朝 7:00 JST 頃、進捗サマリを Discord に通知（@everyone） |
| `film_monday.yml` | 毎週月曜朝、今後14日間の撮影予定を Discord に通知 |
| `schedule_keepalive.yml` | cron が止まらないよう週1でタイムスタンプをコミット |

## 信号の意味（判定はスプレッドシート側の数式で実施）

- 🔴 **赤（要対応）**: 公開まで 7 日を切った
- 🟡 **黄（もうすぐ）**: 今の工程の締切を過ぎた、または 破綻（CL提出締切＞公開リミット）
- 🔵 **青（順調）**: それ以外・完了
- ⚪ **灰（情報不足）**: タイトル・投稿予定日・担当のいずれかが未入力

判定ルールを変えたいときはスプレッドシートを編集すればよく、アプリの再デプロイは不要。
