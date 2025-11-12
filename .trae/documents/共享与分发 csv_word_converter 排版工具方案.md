## 使用场景与推荐路径
- 面向开发者：提供可安装的 wheel 包（pip 安装），即可在任意支持 Python 的电脑使用。
- 面向普通用户（Windows）：提供单文件/单目录 EXE（PyInstaller），无需安装 Python。
- 面向团队/跨平台：部署 Web 服务（Flask，`--port`）或提供 Docker 镜像。
- 便携场景：打包好虚拟环境（venv）与脚本的压缩包，解压即用。

## 方案A：打包为 pip 包（wheel）
- 入口与脚本：`csv2word`、`csv-word-convert`（setup.py:103-108，pyproject.toml:85-88）；模块入口：`python -m csv_word_converter`（src/csv_word_converter/__main__.py:17-24）。
- 构建步骤（本地）：
  1) 安装构建工具：`pip install build`
  2) 在项目根运行：`python -m build`，生成 `dist/*.whl` 与 `dist/*.tar.gz`
  3) 分发给用户：用户执行 `pip install <path\to\whl>`，使用 `csv2word` 或 `python -m csv_word_converter`
- 私有/公开发布：
  - 私有 PyPI（公司内网索引）或 TestPyPI；或直接分发 wheel 文件。
- 注意：根级 `templates_config.yaml` 当前未随包构建（见下文“资源打包注意事项”），需补充清单或迁移到包内。

## 方案B：Windows 可执行文件（PyInstaller）
- 适用：给不想装 Python 的用户；打包所有依赖为一个 `csv2word.exe` 或目录。
- 构建要点：
  - 指定入口脚本：`src\csv_word_converter\cli.py`（主函数 `cli.main()`，src/csv_word_converter/cli.py:428-546,559-560）
  - 添加数据文件：
    - `templates_config.yaml`（根级，当前未进包）
    - `ab_response_formats/*.docx`（根级，MANIFEST.in 已包含，但 PyInstaller 需用 `--add-data` 提供）
    - 包内 `templates/*.docx`、`config/*.yaml/*.json`
  - 示例命令（后续我来写好 .spec 并固化路径）：
    - `pyinstaller --name csv2word --onefile src\csv_word_converter\cli.py --add-data "templates_config.yaml;." --add-data "ab_response_formats;ab_response_formats" --add-data "src/csv_word_converter/templates;csv_word_converter/templates" --add-data "src/csv_word_converter/config;csv_word_converter/config"`
- 分发：提供 `csv2word.exe` 与必要的模板目录（若使用 `--onefolder`，目录更易维护）。
- 限制：若选择 `--format pdf` 且走 `docx2pdf` 路径，Windows 通常需要本机安装 Microsoft Word（src/csv_word_converter/output_formats.py:273-282）。

## 方案C：Web 服务与 Docker
- CLI 开启 Web：提供 `--port` 即进入服务模式（src/csv_word_converter/cli.py:205-209,449-453），内部调用 `start_web_server`（src/csv_word_converter/web_server.py:140-157,159-161）。
- 本地运行示例：`csv2word --port 8501`，浏览器访问 `http://localhost:8501/`
- Docker 镜像（示意）：
  - 基础：`python:3.11-slim`
  - 拷贝代码与模板（同资源注意事项）
  - 安装依赖：`pip install -r requirements.txt && pip install .`
  - 启动命令：`csv2word --port 8000`
  - 运行：`docker run -p 8000:8000 <image>`
- 优点：跨平台、用户仅需浏览器；适合团队共享与云端部署。

## 方案D：便携式 venv 压缩包
- 步骤：在本机创建 `.venv`，安装 `requirements.txt` 与本包；将 `.venv`、`scripts`（Windows 的 `Scripts\csv2word.exe`）和模板打包为 zip。
- 用户侧：解压后执行 `.<解压路径>\.venv\Scripts\csv2word.exe` 或 `python -m csv_word_converter`。
- 优点：不污染系统 Python；缺点：体积较大、与构建机 OS/架构强耦合。

## 资源打包注意事项（当前仓库现状）
- 已包含：`ab_response_formats/*.docx`（MANIFEST.in:18）、包内资源（setup.py:111-118，pyproject.toml:104-110）。
- 未包含：根级 `templates_config.yaml` 未进入包（未出现在 MANIFEST/包数据配置）。
- 建议修复（任选其一）：
  - 在 `MANIFEST.in` 追加：`include templates_config.yaml`
  - 或将 `templates_config.yaml` 迁移到 `src/csv_word_converter/config/` 并更新代码引用，使其被 `package_data` 自动打包。
- 目的：确保 wheel、PyInstaller、Docker、venv 压缩包都能在目标机读取到模板配置。

## 兼容性与外部依赖
- PDF 输出：
  - 首选 `reportlab`（跨平台，src/csv_word_converter/output_formats.py:27-36,191-200）
  - 备选 `docx2pdf`（Windows 上需要 Microsoft Word，src/csv_word_converter/output_formats.py:38-42,273-282）
- 其他依赖：`pandas`、`python-docx`、`docxcompose`、`PyYAML`、`Pillow`、`requests`、`openpyxl`、`lxml`（pyproject.toml:35-45；setup.py:54-67）

## 下一步执行安排（确认后我来落地）
1) 补全资源打包：将 `templates_config.yaml` 纳入构建（MANIFEST 或迁移入包）。
2) 产出三类分发工件：
   - `dist/*.whl`（pip 安装）与发布说明
   - `csv2word.exe`（PyInstaller，含数据文件）与绿色版目录
   - `Dockerfile` 与镜像（启动 `--port` 的 Web 形态）
3) 写一份“安装与使用指南”（pip 安装、EXE 使用、Web/容器部署）。
4) 可选：Streamlit GUI（`streamlit run streamlit_frontend/streamlit_app.py`），适合非技术用户（文件：streamlit_frontend/streamlit_app.py:457）。

— 请确认你更偏好的分发方式（单选或多选）。我将按上述步骤落地构建与验证，并提供可直接交付的工件与文档。