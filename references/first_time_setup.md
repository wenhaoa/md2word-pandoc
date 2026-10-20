# 环境配置

在当前终端检查：

```powershell
pandoc --version
node --version
python --version
python -m pip show python-docx docxcompose lxml
```

缺少Python包时，使用同一个Python解释器安装：

```powershell
python -m pip install python-docx docxcompose lxml
```

Pandoc和Node.js须在PATH中可调用。MathType、npm图片包和Word自动化不是基本转换的必需依赖。
模板为`templates/md2word模板.docx`；必须使用同一份skill中的脚本、过滤器和模板。

```powershell
node scripts/check_markdown.js examples/示例技术报告.md --json
python scripts/test_regressions.py
```

示例报告采用独立题注和章内编号，预检应通过；转换时加`--citations`演示文献上标。
标题、题注、引用规则见[SKILL.md](../SKILL.md)。快捷方式为可选入口，不需要安装即可执行命令行转换。

保留原有安装更新入口`scripts/install_or_update.ps1`及双击、拖拽、右键“发送到”、PowerShell函数入口。封面继续使用现有模板，支持`{{TITLE}}`和`{{DATE_CN}}`，同时兼容字面年月。
