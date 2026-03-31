# 社会基盤ユニットQMS管理システム v3.1

## セットアップ

1. **GASプロジェクト作成** → [script.google.com](https://script.google.com/)
2. **ファイル配置** → `Code.gs` / `Index`(HTML) / `Stylesheet`(HTML) / `JavaScript`(HTML)
3. **Drive API v3 有効化** → サービス → Drive API
4. **進捗管理表シート作成** → 耐震技術部 / 技術開発部 / 土木設計部 / 風力技術部 / バックエンド技術部 / 地下開発技術部
5. **件名テンプレート配置** → `1kcAYs0mXtCc2qsei9NNiI8fY_GwZzLi3` フォルダに「件名_Default」を配置
6. **デプロイ** → ウェブアプリとしてデプロイ

## v3.1 変更点

### 件名フォルダの自動作成
件名読込時に `{部署名}/{件名コード}_{件名}` フォルダを自動作成。
件名スプレッドシート・承認済みPDFともにこのフォルダに保存。

### 議事録画面の縦並びレイアウト
議事録No → 件名コード → 件名 → 年月日 → 開始時間 → 終了時間 → 場所の順に縦並び表記。

### 作成者の氏名表示
ログインメールアドレスから電話帳を検索し、氏名で議事録に記載（`getNameByEmail()`）。

### 承認確認画面の修正
承認依頼メールのURLに `dept`（部署名）、`row`（行番号）、`code`（件名コード）、`kenmei`（件名）をパラメータとして付加。
承認画面では `ssId` + `no` で議事録シートから直接データを取得するため、トークン不要（GASのdoGetはデプロイ時のアクセス権限で動作）。

### PDF出力の改善
画面表示と同じ縦並びレイアウトでPDFを生成。
作成者・実施責任者の下に作成日・承認日を記載。

## フォルダ構成

```
件名フォルダ (1kcAYs0mXtCc2qsei9NNiI8fY_GwZzLi3)
├── 件名_Default（テンプレート）
├── 耐震技術部/
│   ├── ABC001_○○橋梁設計/
│   │   ├── ABC001_○○橋梁設計.gsheet（件名SS）
│   │   └── 議事録_No1.pdf
│   └── ...
├── 技術開発部/
│   └── ...
└── ...
```

## 承認ワークフロー

1. 作成者が議事録を保存
2. 「承認依頼を送信」→ 実施責任者にメール送信
3. メール内URLパラメータ: `?mode=approve&ssId=...&no=...&dept=...&row=...&code=...&kenmei=...`
4. 承認 → 承認日記録 + PDF生成（件名フォルダに保存）+ 進捗管理表更新
5. 否決 → コメント付きで作成者に差戻メール

## 参照ID
- 進捗管理表: `1fhMHbLWHeSIF4HRd9d44Aqmw0aS7j4WYKI3FsGQyKKE`
- 電話帳: `1x6Uy711HFPwdLPFNxyCvMk0Fo77XC4MJNEMh29_n0Lo`
- 受注データフォルダ: `10azgUkgwEKMxfmv5O9GVmwuos9KiAB-y`
- 件名フォルダ: `1kcAYs0mXtCc2qsei9NNiI8fY_GwZzLi3`
- AY列（51列目）: 件名スプレッドシートID格納
