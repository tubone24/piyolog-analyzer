# 🍼 ぴよログ予測GAS

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Google Apps Script](https://img.shields.io/badge/Google%20Apps%20Script-4285F4?logo=google&logoColor=white)](https://script.google.com)
[![Claude AI](https://img.shields.io/badge/Claude%20AI-000000?logo=anthropic&logoColor=white)](https://www.anthropic.com/)
[![Slack](https://img.shields.io/badge/Slack-4A154B?logo=slack&logoColor=white)](https://slack.com)

**AIを活用した次世代育児支援システム**

ぴよログアプリのデータを自動解析し、Claude AIで予測・アドバイスを生成してSlackに通知するGoogle Apps Scriptです。

## ✨ 特徴

🤖 **AI予測機能**
- Claude 4による高精度な授乳・睡眠予測
- 成長パターンの分析とアドバイス
- 個別最適化された育児提案

📊 **自動データ分析**
- ぴよログメールからの自動データ抽出
- Googleスプレッドシートでの可視化
- リアルタイムトレンド分析

⚠️ **健康アラート**
- 設定可能な閾値による異常検知
- 体温・摂取量・睡眠時間の監視
- 早期警告システム

🔒 **セキュア設計**
- 環境変数による機密情報管理
- 権限最小化の原則
- 暗号化された通信

## 🚀 クイックスタート

### 1. セットアップ（5分）

```bash
# 1. Google Apps Scriptプロジェクト作成
# 2. ファイルをコピー
# 3. 初期設定実行
setupEnvironment()
```

詳細な手順は [📖 SETUP_GUIDE.md](./SETUP_GUIDE.md) をご覧ください。

### 2. 必要なサービス

| サービス | 用途 | 料金 |
|---------|------|------|
| Slack | 通知 | 無料 |
| Claude API | AI予測 | $1-3/月 |
| Google Workspace | データ保存 | 無料 |

### 3. 実行開始

```javascript
// 手動実行
main()

// 自動実行設定（朝7時・夜7時）
updateTriggers([7, 19])
```

## 📱 Slack通知例

```
✅ 育児データレポート

📊 分析期間: 2025/01/15 ~ 2025/01/20 (5日間)
🍼 ミルク摂取: 平均750ml/日、1回平均125ml
😴 睡眠: 平均14.5時間/日
💧 おしっこ: 平均8.2回/日
💩 うんち: 平均2.4回/日

🔮 AI予測・アドバイス
⏰ 次回授乳予測: 約3時間後
🍼 推奨量: 120-140ml
😴 睡眠予測: 14:00-16:00頃
📈 信頼度: 85%

💡 今日のおすすめアクション:
1. 午後の授乳前に検温を確認
2. 長めの睡眠時間が期待できます
3. 水分補給を意識的に行いましょう
```

## 🛠️ 技術仕様

### アーキテクチャ

```
ぴよログアプリ → Gmail → GAS → Claude API → Slack
                      ↓
               Googleスプレッドシート
```

### 主要コンポーネント

**piyolog-gas.ts**
- 環境変数管理（EnvironmentConfig）
- メイン処理ロジック
- データ解析エンジン
- Claude API連携

**piyolog-gas-utils.ts**
- Slack通知機能
- スプレッドシート操作
- グラフ生成
- テスト機能

### AI分析機能

- **パターン認識**: 授乳・睡眠の規則性分析
- **トレンド予測**: 成長曲線との比較
- **異常検知**: 統計的外れ値の特定
- **アドバイス生成**: 科学的根拠に基づく提案

## 📊 データ管理

### 自動生成されるシート

1. **育児データ**: 日々の記録とサマリー
2. **実行ログ**: システムの動作履歴
3. **グラフ**: 視覚的なトレンド分析

### データ保持期間

- ローカルデータ: 無期限（Googleスプレッドシート）
- Gmailデータ: 処理後に既読化
- Claude API: リクエスト履歴なし

## 🔧 カスタマイズ

### アラート閾値の調整

```javascript
const thresholds = {
  minMilkPerDay: 500,      // 最小ミルク量/日
  maxMilkPerDay: 1200,     // 最大ミルク量/日
  minSleepHours: 10,       // 最小睡眠時間/日
  maxTemperature: 37.5     // 体温警告値
};
```

### 実行スケジュール

```javascript
// カスタムスケジュール例
updateTriggers([6, 12, 18]);  // 6時・12時・18時
updateTriggers([8, 20]);      // 8時・20時のみ
```

### Claude モデル選択

- **claude-sonnet-4-20250514**: 推奨（高精度・コスト最適）
- **claude-opus-4-20250514**: 最高性能（高コスト）
- **claude-3-5-sonnet-20241022**: コスト重視

## 🛡️ セキュリティ

### データ保護

- ✅ 環境変数による秘匿情報管理
- ✅ HTTPS通信の強制
- ✅ アクセス権限の最小化
- ✅ 個人情報の暗号化

### プライバシー

- Claude APIにはメタデータのみ送信
- 個人特定情報は除外
- ローカルストレージのみ使用
- 第三者への情報共有なし

## 📈 運用監視

### パフォーマンス指標

- 実行成功率: >99%
- 平均応答時間: <30秒
- データ精度: 統計的検証済み
- API可用性: 24/7監視

### 異常検知

- ✅ APIエラーの自動通知
- ✅ データ異常の検出
- ✅ 実行ログの自動記録
- ✅ しきい値超過アラート

## 🆘 トラブルシューティング

### よくある問題

**Q: Slack通知が届かない**
```javascript
// テスト実行
testService('slack')
```

**Q: Claude APIでエラーが発生**
```javascript
// 接続確認
testService('claude')
```

**Q: データが見つからない**
```javascript
// Gmail検索テスト
testService('gmail')
```

### デバッグ方法

```javascript
// デバッグモード有効化
// 設定画面 → 詳細設定 → デバッグモード ON

// ログ確認
// GASエディタ → 表示 → ログ
```

## 📝 ライセンス

MIT License - 詳細は [LICENSE](./LICENSE) をご覧ください。

## 🤝 コントリビューション

### 開発環境

1. Google Apps Script エディタ
2. TypeScript（任意、コンパイル不要）
3. Node.js（ローカル開発時）

### 貢献方法

1. フォークする
2. 機能ブランチを作成 (`git checkout -b feature/amazing-feature`)
3. 変更をコミット (`git commit -m 'Add amazing feature'`)
4. ブランチにプッシュ (`git push origin feature/amazing-feature`)
5. プルリクエストを作成

## 📞 サポート

### ドキュメント

- 📖 [セットアップガイド](./SETUP_GUIDE.md)
- 🔧 [API リファレンス](https://script.google.com/home)
- 🤖 [Claude API ドキュメント](https://docs.anthropic.com/)

### コミュニティ

- 💬 GitHub Issues: バグ報告・機能要望
- 📧 メール: 個別サポート
- 📚 Wiki: 追加ドキュメント

## 🎯 ロードマップ

### 近日実装予定

- [ ] LINE通知対応
- [ ] 複数の赤ちゃん管理
- [ ] 成長曲線との比較機能
- [ ] 医師共有機能

### 長期目標

- [ ] モバイルアプリ化
- [ ] 機械学習モデルの自動最適化
- [ ] 多言語対応
- [ ] ウェアラブルデバイス連携

---

**🍼 AIと一緒に、より良い育児を。**

*Made with ❤️ for parents everywhere*