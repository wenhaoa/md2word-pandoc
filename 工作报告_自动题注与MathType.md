# md2word-pandoc 增强：自动题注 + MathType 公式

> 对话日期：2026-03-04 | 对话ID：c4b00d4a-888e-4249-a9b8-25669236f7a0

## 一、功能概述

本次为 md2word-pandoc 技能新增了**图表题注自动编号**功能，已集成到转换流程中。

### 已完成

**自动题注 SEQ 域**：Markdown 中的 `图N-M` / `表N-M` 题注在转换为 Word 后自动变为 SEQ 域，支持自动编号更新。

### 待后续完成

**MathType 公式转换**：OMML → MathType 对象的自动转换脚本框架已搭建，但 MathType 宏接口未能自动化调用（详见下方"已知限制"）。

---

## 二、使用方法

### Markdown 写法

**图片题注**（图片下方，由 `![...]` 语法生成）：
```markdown
![图2-1 智能温室控制系统架构示意图](images/system_architecture.png)
```

**表格题注**（独立一行，表格上方）：
```markdown
表2-1 各层核心组件及功能描述

| 层级 | 核心组件 | 功能描述 |
| ---- | -------- | -------- |
| ...  | ...      | ...      |
```

**正文引用**（不受影响，不会被转换）：
```markdown
系统架构如图2-1所示，各组件如表2-1所示。
```

### 转换命令

```powershell
# 标准转换（含题注处理）
node run_conversion.js "报告.md"

# 跳过题注处理
node run_conversion.js "报告.md" --no-caption

# 跳过 MathType 转换
node run_conversion.js "报告.md" --no-mathtype
```

### Word 中更新域

转换完成后打开 docx，按 **Ctrl+A → F9** 更新所有域即可显示正确编号。

---

## 三、转换流程

```
1    预处理（空格清理+引号转换）
2    Pandoc 转换（Lua 过滤器+模板）
2.5  合并封面与目录
2.6  添加图表题注 SEQ 域        ← V1.3 新增
2.7  转换公式为 MathType         ← V1.3 新增（实验性）
3    重命名输出文件
```

---

## 四、新增/修改的文件

| 文件                             | 类型 | 说明                                 |
| -------------------------------- | ---- | ------------------------------------ |
| `scripts/add_captions.py`        | 新增 | 后处理脚本，注入 SEQ 域              |
| `scripts/convert_mathtype.vbs`   | 新增 | MathType 转换（实验性）              |
| `scripts/run_conversion.js`      | 修改 | 集成 step 2.6/2.7，新增命令行参数    |
| `examples/示例技术报告.md`       | 修改 | 添加图片/表格题注演示                |
| `SKILL.md`                       | 修改 | 新增 §3 题注组件、题注写法规范、V1.3 |
| `references/first_time_setup.md` | 修改 | 更新 Rules D.4 预检清单              |

---

## 五、后续新增功能时需同步更新的检查清单

> 每次为 md2word-pandoc 增加新功能后，按此清单逐项更新：

| 序号 | 更新项                    | 文件路径                                        | 说明                                                   |
| ---- | ------------------------- | ----------------------------------------------- | ------------------------------------------------------ |
| 1    | **SKILL.md**              | `SKILL.md`                                      | 核心组件章节、Markdown 文档要求、AI 操作指令、版本历史 |
| 2    | **Rules Part D**          | `references/first_time_setup.md` → D.4 预检清单 | 新增的 Markdown 写法约定                               |
| 3    | **用户 GEMINI.md**        | `~/.gemini/GEMINI.md` → Part D                  | 与 first_time_setup.md 保持同步                        |
| 4    | **示例文档**              | `examples/示例技术报告.md`                      | 添加新功能的使用示例                                   |
| 5    | **run_conversion.js**     | `scripts/run_conversion.js`                     | 集成新的后处理步骤                                     |
| 6    | **report-check workflow** | `global_workflows/report-check.md`              | 更新预检规则                                           |
| 7    | **工作报告**              | 项目根目录或 brain 目录                         | 记录改动总结                                           |

---

## 六、已知限制

1. **STYLEREF 依赖模板编号**：`STYLEREF 1 \s` 需要 Word 模板中 Heading 1 绑定了多级列表编号，否则章节号显示为空
2. **题注长度限制**：题注行超过 60 字符会被跳过（防误匹配正文）
3. **MathType 自动化未完成**：MathType 的"Convert Equations"功能通过 Ribbon 回调实现，非标准 VBA 宏接口，VBScript 无法直接调用
