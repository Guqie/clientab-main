# 项目结构说明

本文档整理 `clientab-main` 仓库的主要目录与文件，便于快速了解项目布局及责任划分。

## 顶层目录概览

| 路径 | 说明 |
| --- | --- |
| `src/` | Python 业务代码，包含核心包 `csv_word_converter` 及其工具。 |
| `tests/` | Pytest 测试套件，与核心包结构对应，包含集成测试。 |
| `docs/` | 项目指南、部署说明等文档资源。 |
| `streamlit_frontend/` | Streamlit 前端应用。 |
| `ab_response_formats/` | Word 模板集合，用于文档生成。 |
| `outputs/` | 默认的生成结果目录。 |
| `utils/` | 运行时辅助脚本（如文档处理工具）。 |
| `image_cache/` | 图片缓存目录，通常不纳入版本控制。 |
| `mcp_integration/` | MCP服务器集成相关文件。 |

## 核心代码 (`src/csv_word_converter/`)

- `cli.py`：命令行入口，解析参数并触发转换流程。
- `core.py`：主转换引擎，实现 CSV 到 Word 的数据映射逻辑。
- `async_converter.py`、`batch_processor.py`：异步与批量处理支持，用于并行生成文档。
- `output_formats.py`：定义输出文档格式、样式与结构。
- `web_server.py`：为 Web/Streamlit 集成提供的服务端逻辑。
- `utils/`：包内部工具模块（日志、模板处理等），供核心组件复用。

## 测试与质量保障

- `tests/`：
  - `test_core.py`、`test_async_converter.py` 等单元测试覆盖核心模块。
  - `test_csv_word_integration.py`、`test_integration.py` 等集成测试验证端到端流程。
  - `conftest.py` 提供共享夹具与测试配置。
- 顶层 `pytest.ini`、`pyproject.toml` 配置测试、格式化与类型检查规则。

## 模板、配置与数据资源

- `ab_response_formats/`：各类 Word 模板；模板选择由 `templates_config.yaml` 控制。
- `outputs/`：正式导出的文档目录；提交代码前需清理不必要的生成物。
- `image_cache/`：图片下载缓存目录，提高重复使用图片的加载速度。

## 应用与部署相关

- `streamlit_frontend/`：Streamlit 应用前端。
- `mcp_integration/`：MCP服务器集成配置和脚本。
- `Procfile`、`runtime.txt`、`requirements_streamlit.txt` 等文件支撑 Heroku 等环境部署。

## 辅助脚本与工具

- `utils/`：包含文档处理、图片处理等辅助工具。
- `cli.py`：命令行接口，提供便捷的CSV到Word转换功能。
- `fix_config.py`：配置文件修复工具（如BOM编码问题）。

## 运行说明

- 在 Windows PowerShell 中运行 CLI 时，建议先设置 `PYTHONIOENCODING=utf-8`，例如：
  - `$env:PYTHONIOENCODING = \"utf-8\"`
  - `.\.venv\\Scripts\\python.exe -m csv_word_converter input.csv -t realty`
- 若使用一次性命令，可写成：`$env:PYTHONIOENCODING=\"utf-8\"; python -m csv_word_converter ...`，以避免中文日志或提示的编码异常。

## 目录维护建议

- 避免将生成的文档、缓存或临时文件提交至版本库，保持 `dist/`、`image_cache/`、`test_outputs/` 等目录的清洁。
- 新增模块时同步更新对应测试，并在 `docs/` 中补充文档以保持知识库一致。
- 提交前务必运行 `black`、`isort`、`flake8`、`mypy` 及 `pytest`，确保质量门槛满足团队规范。
