# Git セットアップ・運用ガイド

このドキュメントは、`powerpoint-genAI-update` リポジトリのGit設定と運用方法をまとめたものです。

## 目次

1. [リポジトリ情報](#リポジトリ情報)
2. [認証設定](#認証設定)
3. [基本操作](#基本操作)
4. [ブランチ戦略](#ブランチ戦略)
5. [コミットメッセージ規約](#コミットメッセージ規約)
6. [トラブルシューティング](#トラブルシューティング)

---

## リポジトリ情報

### リポジトリURL

- **HTTPS**: `https://github.com/toshimichi-rakuten/powerpoint-genAI-update.git`
- **SSH**: `git@github.com:toshimichi-rakuten/powerpoint-genAI-update.git`

### ブランチ構成

- **main**: 本番ブランチ（安定版）
- **develop**: 開発ブランチ（機能開発用）
- **feature/***: 機能追加ブランチ
- **fix/***: バグ修正ブランチ

### ローカルパス

```
/Users/toshimichi.suzuki/Desktop/Chrome拡張機能/パワポ自動化/既存合体＋API/powerpoint-genAI-update
```

---

## 認証設定

### 方法1: Personal Access Token (HTTPS) - 推奨

#### トークンの作成

1. GitHubにログイン（toshimichi-rakutenアカウント）
2. Settings → Developer settings → Personal access tokens → Tokens (classic)
3. **Generate new token (classic)** をクリック
4. 設定:
   - **Note**: `Git Push Token`
   - **Expiration**: 90 days（または任意）
   - **Scopes**: ✅ **repo**（全てにチェック）
5. **Generate token** → トークンをコピー（後で見れないので保存）

#### リモートURLの設定

```bash
cd /Users/toshimichi.suzuki/Desktop/Chrome拡張機能/パワポ自動化/既存合体＋API/powerpoint-genAI-update

# トークンを含めてリモートURLを設定
git remote set-url origin https://YOUR_TOKEN@github.com/toshimichi-rakuten/powerpoint-genAI-update.git
```

**設定例**:
```
https://YOUR_TOKEN@github.com/toshimichi-rakuten/powerpoint-genAI-update.git
```

**注意**: `YOUR_TOKEN` の部分を実際のPersonal Access Tokenに置き換えてください。

### 方法2: SSH鍵（複数アカウント対応）

#### SSH鍵の作成

```bash
# toshimichi-rakutenアカウント用の鍵を作成
ssh-keygen -t ed25519 -C "toshimichi-rakuten@github.com" -f ~/.ssh/id_ed25519_rakuten
```

作成された鍵:
- **秘密鍵**: `~/.ssh/id_ed25519_rakuten`
- **公開鍵**: `~/.ssh/id_ed25519_rakuten.pub`

#### SSH設定ファイル

`~/.ssh/config` に以下を追加:

```
# Default GitHub account (Guy0121)
Host github.com
  HostName github.com
  User git
  IdentityFile ~/.ssh/id_ed25519
  IdentitiesOnly yes

# toshimichi-rakuten account
Host github.com-rakuten
  HostName github.com
  User git
  IdentityFile ~/.ssh/id_ed25519_rakuten
  IdentitiesOnly yes
  PreferredAuthentications publickey
```

#### GitHubに公開鍵を登録

1. 公開鍵を表示:
   ```bash
   cat ~/.ssh/id_ed25519_rakuten.pub
   ```

2. GitHubで登録:
   - Settings → SSH and GPG keys → New SSH key
   - Title: `Macbook Pro`
   - Key: 公開鍵の内容を貼り付け

3. 接続テスト:
   ```bash
   ssh -T git@github.com-rakuten
   # 成功: "Hi toshimichi-rakuten! You've successfully authenticated..."
   ```

#### リモートURLの設定（SSH）

```bash
git remote set-url origin git@github.com-rakuten:toshimichi-rakuten/powerpoint-genAI-update.git
```

---

## 基本操作

### 初回クローン

```bash
# HTTPS（トークン使用）
git clone https://YOUR_TOKEN@github.com/toshimichi-rakuten/powerpoint-genAI-update.git

# SSH
git clone git@github.com-rakuten:toshimichi-rakuten/powerpoint-genAI-update.git
```

### 日常的な操作

#### 1. 変更を確認

```bash
git status
```

#### 2. ファイルをステージング

```bash
# 特定のファイル
git add path/to/file

# 全ての変更
git add .

# フォルダごと
git add sandbox/icon/
```

#### 3. コミット

```bash
git commit -m "コミットメッセージ"
```

#### 4. プッシュ

```bash
# mainブランチにプッシュ
git push origin main

# 現在のブランチにプッシュ
git push
```

#### 5. プル（リモートの変更を取得）

```bash
git pull origin main
```

### よく使うコマンド

```bash
# 現在のブランチを確認
git branch

# 新しいブランチを作成して切り替え
git checkout -b feature/new-feature

# ブランチを切り替え
git checkout main

# コミット履歴を確認
git log --oneline -10

# リモートURLを確認
git remote -v

# 差分を確認
git diff
```

---

## ブランチ戦略

### ブランチの種類

| ブランチ | 用途 | 命名規則 | 例 |
|---------|------|---------|-----|
| main | 本番環境（安定版） | `main` | - |
| develop | 開発環境 | `develop` | - |
| feature | 新機能開発 | `feature/機能名` | `feature/template-save` |
| fix | バグ修正 | `fix/バグ名` | `fix/preview-layout` |
| hotfix | 緊急修正 | `hotfix/修正内容` | `hotfix/api-error` |

### ブランチ運用フロー

#### 1. 機能開発の流れ

```bash
# 1. mainから最新を取得
git checkout main
git pull origin main

# 2. 機能ブランチを作成
git checkout -b feature/new-icon-system

# 3. 開発・コミット
git add .
git commit -m "Add new icon system"

# 4. プッシュ
git push origin feature/new-icon-system

# 5. GitHub上でPull Requestを作成
# 6. レビュー後、mainにマージ
```

#### 2. バグ修正の流れ

```bash
# 1. mainから最新を取得
git checkout main
git pull origin main

# 2. 修正ブランチを作成
git checkout -b fix/preview-bug

# 3. 修正・コミット
git add .
git commit -m "Fix preview layout bug"

# 4. プッシュ
git push origin fix/preview-bug

# 5. Pull Request作成・マージ
```

---

## コミットメッセージ規約

### 基本フォーマット

```
<type>: <subject>

<body>

<footer>
```

### Type（種類）

| Type | 説明 | 例 |
|------|------|-----|
| feat | 新機能追加 | `feat: Add template preview feature` |
| fix | バグ修正 | `fix: Fix modal close button issue` |
| docs | ドキュメント変更 | `docs: Update README` |
| style | コードフォーマット | `style: Format code with Prettier` |
| refactor | リファクタリング | `refactor: Simplify template save logic` |
| test | テスト追加・修正 | `test: Add unit tests for API client` |
| chore | ビルド・設定変更 | `chore: Update dependencies` |

### 例

#### シンプルなコミット

```bash
git commit -m "feat: Add icon folder with Font Awesome icons"
```

#### 詳細なコミット

```bash
git commit -m "$(cat <<'EOF'
feat: Add template save modal with custom naming

- Add modal dialog for template naming
- Implement save button disable during processing
- Add preview display on template usage page
- Update button colors (blue for back button)

Closes #123
EOF
)"
```

#### Claude Codeでの推奨フォーマット

```bash
git commit -m "$(cat <<'EOF'
feat: Add new feature

詳細な説明をここに記載

🤖 Generated with [Claude Code](https://claude.com/claude-code)

Co-Authored-By: Claude <noreply@anthropic.com>
EOF
)"
```

---

## トラブルシューティング

### 1. Permission denied (publickey)

**症状**: SSH接続時に権限エラー

**解決策**:

```bash
# 接続テスト
ssh -T git@github.com-rakuten

# 鍵が正しいか確認
ls -la ~/.ssh/id_ed25519_rakuten*

# SSH設定を確認
cat ~/.ssh/config
```

### 2. Authentication failed (HTTPS)

**症状**: プッシュ時に認証エラー

**解決策**:

```bash
# Personal Access Tokenを含むURLに変更
git remote set-url origin https://YOUR_TOKEN@github.com/toshimichi-rakuten/powerpoint-genAI-update.git
```

### 3. fatal: Could not read from remote repository

**症状**: リポジトリへのアクセスエラー

**原因**:
- SSH鍵がGitHubに登録されていない
- リポジトリへの権限がない
- 間違ったアカウントで接続している

**解決策**:

```bash
# 現在のリモートURLを確認
git remote -v

# 接続テスト
ssh -T git@github.com-rakuten

# 認証されているアカウントを確認（出力の "Hi USERNAME!" を確認）
```

### 4. Your branch is ahead of 'origin/main'

**症状**: ローカルがリモートより進んでいる

**解決策**:

```bash
# プッシュして同期
git push origin main
```

### 5. .DS_Store や一時ファイルが追加される

**解決策**:

`.gitignore` ファイルを作成:

```bash
# macOSシステムファイル
.DS_Store
.AppleDouble
.LSOverride

# npm
node_modules/
npm-debug.log

# 環境変数
.env
.env.local

# IDE
.vscode/
.idea/

# ビルド成果物
dist/
build/
```

### 6. コミットを取り消したい

```bash
# 直前のコミットを取り消し（変更は保持）
git reset --soft HEAD^

# 直前のコミットを完全に取り消し（変更も削除）
git reset --hard HEAD^

# プッシュ済みの場合は慎重に
git revert HEAD
```

---

## 参考リンク

- [GitHub公式ドキュメント](https://docs.github.com/)
- [Pro Git Book（日本語版）](https://git-scm.com/book/ja/v2)
- [GitHub Personal Access Token作成](https://github.com/settings/tokens)
- [SSH鍵の設定](https://docs.github.com/ja/authentication/connecting-to-github-with-ssh)

---

## 更新履歴

| 日付 | 変更内容 | 担当者 |
|------|---------|--------|
| 2025-10-28 | 初版作成 | toshimichi.suzuki |

