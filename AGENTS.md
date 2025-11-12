# 仓库指南

## 项目结构与模块组织
- `src/csv_word_converter/` 目录包含 CLI (`cli.py`)、转换引擎 (`core.py`)、异步工具以及批处理器。
- 模板文件和 YAML 配置位于 `ab_doc_temps/` 与 `templates_config.yaml`；生成的文档默认写入 `outputs/`，或 `temp-data/` 下的临时目录。
- `utils/` 存放运行辅助脚本，`app-dir/` 负责 Streamlit 打包，`docs/` 提供配套指南；请避免将构建产物提交到 `dist/`、`image_cache/` 或 `test_outputs/`。
- `tests/` 中的测试与包结构对应，公共夹具位于 `conftest.py`；示例 CSV 与媒体资源保存在 `temp-data/`、`uploads/` 以及 `uploaded-files/`。

## 构建、测试与开发命令
- `pip install -e .[dev]` 安装库以及配套的 lint、类型检查和 pytest 工具。
- 使用 `csv2word input.csv --template realty` 体验 CLI；运行 `python -m csv_word_converter --help` 查看运行时选项。
- `pytest` 执行测试套件，覆盖率设置在 `pyproject.toml` 中；添加 `-m "not slow"` 便于快速迭代，或使用 `-m integration` 运行流水线检查。
- `black src tests`、`isort src tests`、`flake8` 与 `mypy src` 必须全部通过后才能推送代码。

## 编码风格与命名约定
- Black（行宽 120 字符，四空格缩进）与 isort 的 Black 配置共同约束格式与导入顺序。
- 模块与函数使用 snake_case，类采用 PascalCase，CLI 标志遵循 kebab-case。
- 提供显式类型注解以满足严格的 mypy 要求；结构化模板负载倾向使用 `dataclass` 或 `TypedDict`。
- 模板与资源文件名保持 kebab-case，以符合打包规则。

## 测试指南
- 为新测试添加 pytest 标记 `unit`、`integration`、`slow`，文件命名为 `test_<feature>.py`。
- 共享夹具或工厂放入 `tests/fixtures/`，临时输出限制在 `temp-data/`。
- `pytest --cov-report=html` 会生成覆盖率报告，可在 `htmlcov/index.html` 中查看；修改核心模块时保持覆盖率水平。
- 模拟网络访问与外部服务，确保测试可重复且离线稳定。

## 提交与拉取请求指南
- 提交信息使用简洁的祈使句主题（例如 `Refine template validation`），正文保持 72 字符左右换行。
- 从 `main` 分支创建分支，引用相关 issue，并记录本地验证步骤（`pytest`、`flake8`、`mypy`）。
- 拉取请求需说明背景，模板或界面改动附上截图或示例文档，并在相关情况下引用 `heroku_deploy_guide.md` 中的部署步骤。
- 仅在自动化检查全部通过后请求评审，并在 PR 备注中列出剩余风险或后续任务。
