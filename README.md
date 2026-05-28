# 工资单批量生成

网页工具，用于上传每月工资表，生成每位员工的 PDF 工资单，并打包为 ZIP 下载。

## 本地启动

```bash
./run.sh
```

打开：

```text
http://127.0.0.1:5001
```

如果 5001 端口被占用，可以这样换端口：

```bash
PORT=5010 ./run.sh
```

## 部署到 Render

1. 在 GitHub 新建一个仓库，比如 `payslip-generator`。
2. 把本项目文件提交并推送到这个仓库。
3. 登录 Render，选择 `New` -> `Blueprint`。
4. 连接刚才的 GitHub 仓库。
5. Render 会读取 `render.yaml`，自动创建 Python Web Service。
6. 部署完成后，Render 会给你一个 `https://xxx.onrender.com` 网址，把这个网址发给同事即可。

如果不用 Blueprint，也可以在 Render 里手动创建 `Web Service`：

- Runtime: `Python`
- Build Command: `pip install -r requirements.txt`
- Start Command: `python app.py`

Render 会自动提供 `PORT` 环境变量，本项目会监听 `0.0.0.0:$PORT`。

## 当前规则

- 不计算工资，只读取 Excel 中已经保存的单元格值。
- 只读取第一个工作表。
- 从第 5 行开始识别员工数据，遇到“合计”行停止。
- 空金额显示为 `0.00 元`。
- 每人生成一个 PDF。
- ZIP 内文件名格式为 `4月工资单-员工名字.pdf`。
- 同名员工文件名自动追加 `1`、`2`、`3`。
- 缺姓名或缺月份的行会跳过，并在页面和 ZIP 内的 `生成提示.txt` 中提示。
- 生成的 ZIP 是临时文件；服务每次生成时会清理 1 小时以前的 ZIP。
- 如需使用“方正兰亭超细黑简体”，请把完整字体文件放到 `static/FZLTCXHJW.ttf` 或 `static/方正兰亭超细黑简体.ttf` 后重新部署。

## 数据提醒

部署到 Render 后，HR 上传的工资表会发送到 Render 服务器生成 PDF。当前项目不会使用数据库，也不会保存上传的 Excel；ZIP 会短暂保存在服务临时目录中，供下载使用。
