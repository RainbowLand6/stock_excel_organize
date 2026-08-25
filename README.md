# 库存工作台

用于整理库存 Excel 文件的 Windows 桌面应用，支持：

- 单品牌库存对比：对比同一品牌的上周/本周库存，并导出差异表。
- 多品牌库存汇总，支持按“品牌 + 源仓库名称”映射至任意对应仓库名称，默认目标为“德清仓”。
- 汇总对比：比较两张库存汇总表的合计主单位，并标记新增物料。
- 特殊品牌拆分：按仓库名称中的“品牌名-仓库名”拆分合并库存表，支持清除末行总计、下载拆分结果与清空本次会话。
- 可选本机诊断日志，便于排查上传、导出与处理问题。
- 在本机保存导出配置、映射关系和导出记录，不会修改原始 Excel 文件。

详细操作见 [安装与使用说明](docs/安装与使用说明.md)。

## 项目结构

```text
app.py                 Python 后端与 Excel 处理逻辑
web/                   浏览器界面
desktop/               Electron 桌面壳
模板总模板.xlsx         库存汇总导出模板
lucide.js              本地化图标库
scripts/               可复现的文档与核心验证脚本
docs/                  使用说明和需求文档
prototypes/            早期界面原型，仅供设计参考
```

运行产生的上传缓存、用户配置、导出文件、依赖目录和安装包均不会纳入 Git。`local-data/` 仅用于保留本机 Excel 样例，也不会上传。

## 本地运行

要求：Windows、Python 3.9–3.12。

```powershell
py -3.12 -m venv .venv
.\.venv\Scripts\python -m pip install -r requirements.txt
.\启动库存工作台.bat
```

也可以直接运行：

```powershell
.\.venv\Scripts\python app.py
```

默认地址为 `http://127.0.0.1:4174`。应用数据默认保存在项目目录下的 `.inventory_data/`；该目录仅供本机使用。

## 桌面开发与打包

```powershell
npm ci
npm run desktop:dev
```

生成 Windows 安装包前，请准备 Python 打包环境：

```powershell
py -3.12 -m venv .venv-build
.\.venv-build\Scripts\python -m pip install -r requirements-build.txt
npm run package:win
```

安装包输出到 `release/`，后端临时产物输出到 `build/` 和 `dist/`。

## 验证

核心 Excel 处理逻辑的验证脚本会自行生成临时工作簿和结果：

```powershell
.\.venv\Scripts\python scripts\verify_engine.py
```

重新生成需求文档时，额外安装文档依赖：

```powershell
.\.venv\Scripts\python -m pip install -r requirements-docs.txt
.\.venv\Scripts\python scripts\build_requirements_doc.py
```
