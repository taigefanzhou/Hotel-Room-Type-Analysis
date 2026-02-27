# 上传到GitHub指南

## 第1步：创建GitHub仓库

1. 访问 https://github.com/new
2. 仓库名称：`hotel-cad-analyzer`
3. 选择 "Public"（公开）或 "Private"（私有）
4. **不要**勾选 "Initialize this repository with a README"
5. 点击 "Create repository"

## 第2步：上传代码

### 方法A：使用Git命令行

```bash
# 进入项目目录
cd hotel-cad-analyzer

# 初始化Git仓库
git init

# 添加所有文件
git add .

# 提交
git commit -m "Initial commit"

# 添加远程仓库（替换为你的仓库地址）
git remote add origin https://github.com/你的用户名/hotel-cad-analyzer.git

# 推送
git push -u origin main
```

### 方法B：直接拖拽（更简单）

1. 在GitHub仓库页面，点击 "uploading an existing file" 链接
2. 拖拽本文件夹中的所有文件到网页
3. 点击 "Commit changes"

### 方法C：使用GitHub Desktop

1. 下载GitHub Desktop：https://desktop.github.com
2. File → Add local repository
3. 选择本文件夹
4. 填写提交信息，点击Commit
5. 点击Publish repository

## 第3步：等待自动构建

上传后，GitHub会自动开始构建：

1. 进入仓库页面的 "Actions" 标签
2. 你会看到 "Build Windows EXE" 工作流正在运行
3. 等待5-10分钟（绿色✅表示成功）

## 第4步：下载EXE文件

构建完成后：

1. 进入 "Actions" 页面
2. 点击最新的成功构建
3. 在 "Artifacts" 部分下载 `windows-exe`
4. 解压后得到 `酒店CAD分析器.exe`

或者：

1. 进入 "Releases" 页面
2. 下载最新版本的EXE文件

## 第5步：发给同事

把下载的 `酒店CAD分析器.exe` 发给同事，双击即可使用！

---

## 🔄 后续更新

如果修改了代码：

```bash
git add .
git commit -m "更新说明"
git push
```

GitHub会自动重新构建EXE！

---

## ❓ 常见问题

**Q: Actions构建失败？**  
A: 点击失败的构建，查看日志，通常是依赖问题

**Q: 如何触发重新构建？**  
A: 进入Actions → Build Windows EXE → Run workflow

**Q: 生成的EXE在哪里？**  
A: Actions → 最新构建 → Artifacts → windows-exe