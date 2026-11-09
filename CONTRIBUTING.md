# 贡献指南

欢迎提 Issue 和 Pull Request。这个项目保持「小工具」定位：无框架、无构建步骤、只依赖标准库 + `openpyxl` / `xlrd`，请尽量让改动保持简单。

## 本地开发

```bash
git clone https://github.com/NewObject-Han/excel-a2b-converter.git
cd excel-a2b-converter
pip install -r scripts/requirements.txt
python web/serve.py          # 打开 http://127.0.0.1:8765/
```

## 提交前自查

```bash
python -m py_compile scripts/excel_a2b.py web/serve.py
node --check web/assets/app.js
python scripts/excel_a2b.py convert --config examples/mapping.example.json
python scripts/excel_a2b.py presets --json
```

## 几条约定

1. **不做业务假设**：A→B 的字段对应必须完全由配置决定，不要在引擎里写死任何行业规则或列名。
2. **双端一致**：`scripts/excel_a2b.py` 与 `web/assets/app.js` 各有一份变换实现。新增 / 修改变换时，两边必须**同输入同输出**（含边界值、非法输入、日期、舍入），并各自跑一遍比对。
3. **预设要同步**：新增变换预设时，`excel_a2b.py` 的 `TRANSFORM_PRESETS`、`app.js` 的 `TRANSFORM_PRESETS`、`presets --json` 输出、文档表格四处要一致。
4. **向后兼容**：`row_filter` 单条写法、旧的 `finalize(value, rule)` 调用、旧 `mapping.json` 都要能继续工作。
5. **不要提交数据文件**：`.gitignore` 已排除 `*.xls*` / `*.zip` / `转换结果*.xlsx` 等。请**不要**把真实业务表格、客户数据、打包好的 exe 提交进仓库。
6. **提交信息**：简短中文或英文均可，说明改了什么、为什么。

## 报告问题

请附上：源文件格式（`.xls` / `.xlsx`）、表头行、相关 `mapping.json` 片段、期望结果与实际结果。**不要粘贴含真实客户数据的表格内容**，用一小段脱敏样例即可。
