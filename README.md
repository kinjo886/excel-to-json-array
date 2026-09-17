# Excel 转二维数组工具（Excel to JSON Array）

[![Python](https://img.shields.io/badge/Python-3.8%2B-blue)](https://www.python.org)
[![License](https://img.shields.io/badge/license-MIT-blue)](#license)

> **开源脱敏版。** 原项目为某政务数据治理平台的「二维数组生成工具」。本仓库仅保留通用 Python 脚本，已移除真实业务数据、内部表名与内网连接配置。

从 Excel（`.xlsx`）读取数据，导出为 **JSON 二维字符串数组**，便于直接粘贴进数据开发平台的批量作业入参。

---

## 功能特性

- 每行占多列时：**按列导出**为二维数组
- 每行只有 1 个单元格、且内容为英文逗号分隔时：**自动按 CSV 规则拆成多列**（常见于整行粘到 A 列的场景）
- 支持指定工作表（`--sheet`）、跳过前 N 行（`--skip-rows`）、关闭 CSV 自动拆分（`--no-split-csv-cell`）
- 输出为 UTF-8 的 JSON 二维数组，可直接被前端 / 脚本消费

## 技术栈

- Python 3.8+，唯一依赖：`openpyxl`
- 纯标准库 + openpyxl，无框架

## 目录结构

```
excel-to-json-array/
├── txt_to_2d_json_array.py          # 主脚本：Excel → 二维 JSON 数组
├── V2-创建CDM作业_调整参数顺序_数源参数.py  # 配套脚本：作业入参顺序调整 / 数源参数对齐（内部连接处已脱敏为占位符）
├── 结果_001.json                    # 示例输出
└── python的执行语句.txt             # 常用命令备忘
```

## 安装

```bash
pip install openpyxl
```

## 用法

```bash
# 最简：读取 示例.xlsx，输出到默认 out.json
python txt_to_2d_json_array.py 示例.xlsx

# 指定输出文件名
python txt_to_2d_json_array.py 示例.xlsx -o out.json

# 指定工作表、跳过表头
python txt_to_2d_json_array.py 示例.xlsx --sheet Sheet1 --skip-rows 1

# 关闭「单行逗号自动拆分多列」行为
python txt_to_2d_json_array.py 示例.xlsx --no-split-csv-cell
```

参数说明：

| 参数 | 默认值 | 说明 |
|---|---|---|
| `xlsx` | 必填 | 输入 Excel 文件路径 |
| `-o, --output` | `out.json` | 输出 JSON 路径 |
| `--sheet` | 第一个工作表 | 指定工作表名 |
| `--skip-rows` | `0` | 跳过前 N 行 |
| `--no-split-csv-cell` | 关闭 | 关闭「单行逗号自动拆分多列」 |

## 输出示例

输入 `示例.xlsx`（A 列每行一条逗号分隔文本）：

```
id,name,type
1,foo,A
2,bar,B
```

输出 `out.json`：

```json
[
  ["id", "name", "type"],
  ["1", "foo", "A"],
  ["2", "bar", "B"]
]
```

## 配套脚本说明

`V2-创建CDM作业_调整参数顺序_数源参数.py` 用于作业入参顺序调整与数源参数对齐。原脚本中含内部 Hive / 大数据平台连接配置，本仓库已将其**替换为占位符**（如 `<INTERNAL_HIVE_HOST>`、`<INTERNAL_B64_TEMPLATE>`），仅保留通用参数处理逻辑，需在你自己的环境中按真实配置替换后使用。

## License



本项目基于 [MIT License](./LICENSE) 开源，可自由使用、修改和分发。欢迎按需二次开发。
