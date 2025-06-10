# Git 版本控制工作流程

## 分支結構
- `master` - 主分支，包含穩定的發布版本
- `develop` - 開發分支，用於日常開發工作

## 基本工作流程

### 1. 日常開發
```bash
# 切換到開發分支
git checkout develop

# 拉取最新變更（如果有遠端倉庫）
git pull origin develop

# 進行開發工作...
# 修改檔案後

# 查看變更狀態
git status

# 添加變更到暫存區
git add .

# 或者添加特定檔案
git add 檔案名稱

# 提交變更
git commit -m "描述變更內容"
```

### 2. 發布新版本
```bash
# 切換到主分支
git checkout master

# 合併開發分支
git merge develop

# 建立版本標籤
git tag -a v1.6.0 -m "版本 1.6.0 - 新功能描述"

# 推送到遠端（如果有設定）
git push origin master --tags
```

### 3. 常用命令
```bash
# 查看提交歷史
git log --oneline

# 查看所有分支
git branch -a

# 查看所有標籤
git tag

# 查看目前狀態
git status

# 查看變更差異
git diff

# 撤銷工作區變更
git checkout -- 檔案名稱

# 撤銷暫存區變更
git reset HEAD 檔案名稱
```

### 4. 回滾版本
```bash
# 回滾到特定提交
git reset --hard 提交ID

# 回滾到特定標籤
git reset --hard v1.5.0
```

## 注意事項
1. 在提交前務必測試程式功能
2. 提交訊息要清楚描述變更內容
3. 定期將開發分支的變更合併到主分支
4. 重要版本要建立標籤以便追蹤
5. 避免直接在主分支上進行開發工作

## 建議的提交訊息格式
- `feat: 新增功能描述`
- `fix: 修復問題描述`
- `docs: 文件更新`
- `style: 程式碼格式調整`
- `refactor: 重構程式碼`
- `test: 測試相關變更`
- `chore: 建置或工具變更`
