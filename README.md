# CSV-Word转换工具（项目整理版）

一个面向中文场景的 CSV → Word 转换工具，支持模板化排版、图片内联、批量处理与 Streamlit 前端。

## 1. 项目概览

- 核心能力：CSV 内容解析、模板驱动的 Word 文档生成、图片下载与居中、返回目录锚点管理。
- 运行形态：命令行（CLI）、Python API、Streamlit 前端。
- 模板配置：由 `templates_config.yaml` 管理，支持科技、房地产、新能源、电力、国资委等模板键。

## 2. 环境准备（Windows + PowerShell + 虚拟环境）

```powershell
Set-Location "D:\桌面\clientab-main";
if (-Not (Test-Path ".\.venv\Scripts\Activate.ps1")) { python -m venv .venv };
. .\.venv\Scripts\Activate.ps1;
$env:PYTHONPATH = "D:\桌面\clientab-main\src";
pip install -r .\requirements.txt;
```

说明：
- 必须在虚拟环境下运行，否则会缺少依赖（如 `python-docx`、`pandas`）。
- 设置 `PYTHONPATH` 以便 tests 与前端正确导入 `src` 包。

## 3. 快速上手

### 3.1 命令行（CLI）

```powershell
# 查看帮助
python -m csv_word_converter --help;

# 基本转换（模板：科技）
python -m csv_word_converter "temp-data/科技与产业排版11.12刊.csv" --template technology --output-dir .\temp-data;

# 交互模式与演示模式
python -m csv_word_converter --interactive;
python -m csv_word_converter "input.csv" -t realty --dry-run;
```

常用选项：
- `-t, --template` 指定模板键，如 `technology`、`realty`。
- `-o, --output-dir` 指定输出目录，默认 `temp-data`。
- `--interactive` 交互式引导选择文件与模板。
- `--dry-run` 仅校验与预览，不生成文档。

### 3.2 Streamlit 前端

```powershell
. .\.venv\Scripts\Activate.ps1; $env:PYTHONPATH = "D:\桌面\clientab-main\src"; streamlit run streamlit_frontend\streamlit_app.py --server.port 8501;
```

功能：文件上传、模板选择、批量转换、ZIP 下载、PDF 二次导出、CSV 必要列校验。

### 3.3 Python API

```python
from csv_word_converter import csv_to_word_universal

docx_path = csv_to_word_universal(
    csv_file="data.csv",            # 输入CSV路径
    template_type="technology",     # 模板键
    config_path="templates_config.yaml"  # 模板配置
)
print(docx_path)
```

## 4. 模板与配置

- 模板键与中文名映射在 `templates_config.yaml`（如 `technology` → “科技与产业”）。
- 支持 `start_template` / `end_template`、样式字典、返回目录超链接等配置。
- 需要批量处理时可使用 `src/csv_word_converter/batch_processor.py`。

## 5. 项目结构（整理后）

```
clientab-main/
├── src/csv_word_converter/         # 核心包
│   ├── cli.py, core.py, async_converter.py, batch_processor.py, output_formats.py
│   └── utils/                      # 工具模块
│       ├── doc_utils.py, image_downloader.py, center_image_utils.py, doc_hyperlink_manager.py
├── streamlit_frontend/             # 前端应用
│   └── streamlit_app.py
├── tests/                          # 测试（已精简）
│   ├── test_core.py, test_async_converter.py, test_csv_word_integration.py, conftest.py
├── docs/                           # 文档
│   ├── project_structure_cn.md, universal_csv_to_word_说明文档.md
├── ab_response_formats/            # Word模板集合
├── templates_config.yaml           # 模板配置
├── requirements.txt, pyproject.toml, pytest.ini, setup.py
├── .streamlit/                     # Streamlit 配置
└── .trae/                          # IDE/协作配置
```

已移除：`image_cache/`、`tools/`、顶层 `utils/`（保留包内 `src/csv_word_converter/utils/`）、`app-dir/`、`.devcontainer/`、示例文档与重复测试。

## 6. 测试与质量保障

```powershell
. .\.venv\Scripts\Activate.ps1; pytest -q;
```

说明：
- 已删除 `tests/test_integration.py`（样式不规范且与现有集成测试重叠）。
- 如遇编码问题，可在 PowerShell 设置：`$env:PYTHONIOENCODING = "utf-8";`。

## 7. 常见错误与排查

- 导入失败：确认已设置并导出 `PYTHONPATH` 指向 `src`。
- 依赖缺失：确认虚拟环境已激活并安装 `requirements.txt`。
- 模板找不到：检查 `templates_config.yaml` 中模板键与文件路径是否一致。
- PDF 导出失败：检查 `output_formats.py` 的格式转换器及依赖版本。

## 8. 推送到 GitHub

```powershell
# 初始化或检查仓库
git init; git status;

# 建议提交（本次整理）
git add -A;
git commit -m "chore: 项目整理（清理冗余目录与测试，完善README）";

# 设置远程（将URL替换为你的仓库地址）
git branch -M main;
git remote add origin https://github.com/<your-username>/clientab-main.git;
git push -u origin main;
```

提示：如需使用 SSH，可替换为 `git@github.com:<your-username>/clientab-main.git`；首次推送需完成本机 Git 凭证配置。

## 9. 许可证与致谢

- 许可证：MIT
- 致谢：感谢所有贡献者与用户的反馈与支持。
