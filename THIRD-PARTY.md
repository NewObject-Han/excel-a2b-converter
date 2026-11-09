# 第三方依赖与许可

本仓库遵循 MIT 许可，但包含 / 使用以下第三方组件，各自适用其原始许可。

| 组件 | 版本 | 用途 | 许可 | 是否随源码分发 |
|---|---|---|---|---|
| [SheetJS Community Edition](https://sheetjs.com/) (`xlsx.full.min.js`) | 0.20.3 | 纯浏览器模式下读写 Excel，位于 `web/vendor/` | Apache-2.0 | 是（离线可用，已 vendored） |
| [openpyxl](https://openpyxl.readthedocs.io/) | >=3.1.0 | Python 引擎读写 `.xlsx` | MIT | 否（`pip install`） |
| [xlrd](https://xlrd.readthedocs.io/) | >=2.0.1 | Python 引擎读取老式 `.xls` | BSD-3-Clause | 否（`pip install`） |
| [PyInstaller](https://pyinstaller.org/) | — | 仅 `build_mac.sh` / `build_windows.sh` 打包时使用 | GPLv2 + bootloader exception | 否（仅开发/打包期） |

说明：

- `web/vendor/xlsx.full.min.js` 为 SheetJS 社区版压缩产物，随仓库分发以便在 `file://` 模式下完全离线工作；其版权与许可归原作者所有。
- PyInstaller 不参与运行时，仅在用户主动执行打包脚本时安装；生成的可执行程序的许可由使用者自行评估。
- 本项目不使用任何在线服务，转换过程全部在本地完成，不上传表格数据。
