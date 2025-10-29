# Excel 识别与列增量小工具

一个基于 Tkinter + pandas 的桌面小应用：
- 读取本地 Excel 文件
- 选择需要的列生成新表
- 指定某一列为每个单元格增加固定增量（默认 0.01，可改）
- 预览结果并导出为新的 Excel

## 运行环境
- Python 3.9+

## 安装依赖
```bash
pip install -r requirements.txt
```

## 启动
```bash
python main.py
```

## 打包与分发

### macOS 打包
方式一：直接命令
```bash
pip install pyinstaller
pyinstaller --noconfirm --windowed --name "Excel处理小工具" main.py
```
方式二：脚本
```bash
zsh build_mac.sh
```
生成的应用位于 `dist/Excel处理小工具.app`，可直接拷贝给其他 mac 使用。

首次启动可能被 Gatekeeper 拦截：右键该 `.app` → 选择“打开”即可放行。如果需要对外分发给大量用户，建议进行签名与公证（codesign + notarize）。

### Windows 打包
方式一：直接命令
```powershell
pip install pyinstaller
pyinstaller --noconfirm --windowed --name ExcelTool main.py
```
方式二：脚本
```bat
build_win.bat
```
生成的可执行程序在 `dist/` 目录下（单文件或文件夹形式）。可直接拷贝到无 Python 的电脑运行。

### 在 mac 上生成 Windows 程序（借助 GitHub Actions）

#### 第一步：上传代码到 GitHub

**方法一：使用脚本（推荐）**
```bash
# 在项目目录下运行
zsh setup_git.sh
# 然后按照提示操作
```

**方法二：手动操作**

1. **创建 GitHub 仓库**
   - 访问 https://github.com/new
   - 填写仓库名称（例如：`excel-tool`）
   - 选择 Public 或 Private
   - 不要勾选 "Initialize this repository with a README"（因为本地已有）
   - 点击 "Create repository"

2. **初始化并推送代码**
   在终端执行（替换 `你的用户名` 和 `仓库名`）：
   ```bash
   cd /Users/jzh/Desktop/gui
   
   # 初始化 git（如果还没初始化）
   git init
   
   # 添加所有文件
   git add .
   
   # 提交
   git commit -m "初始提交：Excel 处理小工具"
   
   # 连接到 GitHub 仓库（替换成你的仓库地址）
   git remote add origin https://github.com/你的用户名/仓库名.git
   
   # 设置为 main 分支
   git branch -M main
   
   # 推送到 GitHub
   git push -u origin main
   ```
   
   如果提示需要登录，按照提示操作（可能需要输入 GitHub 用户名和 Personal Access Token）。

#### 第二步：使用 GitHub Actions 打包 Windows 程序

代码推送成功后：
1. 打开你的 GitHub 仓库页面
2. 点击顶部的 **Actions** 标签
3. 左侧选择 **"Build Windows EXE"** 工作流
4. 点击右侧 **"Run workflow"** 按钮 → 选择 `main` 分支 → 点击绿色的 **"Run workflow"**
5. 等待几分钟（会显示构建进度）
6. 构建完成后，在页面底部找到 **Artifacts** 区域
7. 下载 **`ExcelTool-windows-exe`** 压缩包
8. 解压后即可得到 `ExcelTool.exe`，可在任何 Windows 电脑上运行

### 常见问题
- 打包后体积较大属正常，因为包含了运行所需的 Python 运行时与依赖。
- 如果 Excel 读取报错，确保 `requirements.txt` 中的 `openpyxl/xlrd` 已被正确包含；也可在 spec 中通过 `datas` 手动添加资源。
- 如需应用图标，可在 PyInstaller 加上 `--icon your.ico`（Windows）或 `.icns`（macOS）。

## 使用步骤
1. 点击“选择Excel文件”，选择要处理的文件。
2. 在“工作表”下拉框中选择需要处理的 Sheet。
3. 在左侧列表中多选需要提取的列。
4. 在右侧选择要加增量的目标列，并设置增量（默认 0.01）。
5. 点击“应用提取与增量”，在下方预览结果（最多显示前 200 行）。
6. 点击“导出为Excel...”保存新表格。

## 说明
- 增量仅对可数值单元格生效，非数值会原样保留。
- 如果增量列未包含在提取列中，程序会自动将该列加入到结果中。
- 支持 .xlsx/.xls/.xlsm 读取，导出为 .xlsx。

