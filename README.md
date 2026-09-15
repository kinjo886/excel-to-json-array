# Excel 转二维数组工具（Excel to JSON Array）

> 开源脱敏版。原项目为某政务数据治理平台的「二维数组生成工具」，本仓库仅保留通用 Python 脚本，已移除真实业务数据与内部表名。

从 Excel（.xlsx）读取数据，导出为 JSON 二维字符串数组：

- 若每行占多列：按列导出
- 若每行只有 1 个单元格且内容是英文逗号分隔：自动按 CSV 规则拆成多列

## 用法

```bash
pip install openpyxl
python txt_to_2d_json_array.py 示例.xlsx -o out.json
```
