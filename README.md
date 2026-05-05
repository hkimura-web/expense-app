# 経費精算システム

関東学院六浦中学校・高等学校 向け 経費精算Webアプリ。

## 機能

| タブ | 内容 |
|------|------|
| 出張費精算書 | 交通費（最大4路線）・日当の自動合計 |
| 自家用車使用届 | 車種・区間・使用目的など |
| 駐車場代精算書 | 現金払い精算（立替／仮払） |

- **📊 Excel出力** — 3シートまとめて `.xlsx` でダウンロード
- **🖨️ 印刷 / PDF保存** — 印刷プレビューからPDF保存も可

---

## ローカルで動かす

```bash
# 依存インストール
npm install

# 開発サーバー起動（http://localhost:5173）
npm run dev
```

## GitHub Pages にデプロイする手順

### 1. GitHubにリポジトリを作る

https://github.com/new でリポジトリを新規作成（例：`expense-app`）

### 2. このコードをプッシュ

```bash
git init
git add .
git commit -m "initial commit"
git branch -M main
git remote add origin https://github.com/あなたのID/expense-app.git
git push -u origin main
```

### 3. GitHub Actions でビルド＆自動デプロイ

`.github/workflows/deploy.yml` が含まれているので、`main` ブランチにプッシュするたびに
自動的に GitHub Pages へデプロイされます。

**初回のみ：** GitHubのリポジトリ → Settings → Pages → Source を
`GitHub Actions` に変更してください。

デプロイ後、以下のURLでアクセスできます：
```
https://あなたのGitHubID.github.io/expense-app/
```

---

## 技術スタック

- [React 18](https://react.dev/)
- [Vite 5](https://vitejs.dev/)
- [SheetJS (xlsx)](https://sheetjs.com/) — Excel出力
