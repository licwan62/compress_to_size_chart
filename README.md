# 尺码表压缩工作流

这个目录用于把全量 TSV 或 Excel 尺码表转换成可交付的压缩表，并按需要生成压缩日志、原子事实表和检查表。

当前主入口是 `process_tsv.py`。日常建议优先使用它完成整套流程；`check_atom.py` 只用于单独补跑某一张压缩表的原子事实检查。

## 环境准备

第一次使用先安装依赖：

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

之后每次进入目录，只需要先启用环境：

```powershell
.\.venv\Scripts\Activate.ps1
```

## 推荐工作流

### 1. 放入输入文件

把待处理的 TSV 或 Excel 放到：

```text
data/input/
```

例如：

```text
data/input/0702.xlsx
data/input/full0628.tsv
```

### 2. 生成压缩表

直接传入 TSV 或单工作表 Excel：

```powershell
python .\process_tsv.py .\data\input\0702.xlsx
```

输出会写到：

```text
data/output/0702/
```

主要结果在：

```text
data/output/0702/compress/
```

### 3. 多工作表 Excel

如果 Excel 里有多个工作表，需要用字段 profile 指定要读取的工作表：

```powershell
python .\process_tsv.py --field-profile .\field_profile.default.yaml
```

profile 里的 `input.path` 可以指定输入文件；`input.sheets` 可以指定一个或多个工作表。多个工作表会分别处理，并按 `文件名_工作表名` 生成独立输出目录。

### 4. 生成原子检查结果

最终交付前建议加上 `--check-atom`：

```powershell
python .\process_tsv.py .\data\input\0702.xlsx --field-profile .\field_profile.default.yaml --check-atom
```

它会额外生成：

```text
data/output/0702_工作表名/check/
```

## 常用命令例子

### 换输出目录

```powershell
python .\process_tsv.py .\data\input\0702.xlsx -o .\data\output_test
```

结果会写到：

```text
data/output_test/0702/
```

### 输入不是 UTF-8 BOM 编码

TSV 默认读取编码是 `utf-8-sig`。如果源文件是 GBK：

```powershell
python .\process_tsv.py .\data\input\full0628.tsv --encoding gbk
```

Excel 输入不受 `--encoding` 影响。

### 单独检查原子事实和某张压缩表

例如检查非皮卡高度压缩表：

```powershell
python .\check_atom.py --atom .\data\output\0702\compress\0702_原子事实表.tsv --compress .\data\output\0702\compress\0702_非皮卡高度压缩表.tsv -o .\data\output\0702\check
```

会生成：

```text
data/output/0702/check/0702_原子事实表_check.tsv
```

## 参数说明

### process_tsv.py

```powershell
python .\process_tsv.py [输入TSV或Excel] [参数]
```

常见参数：

```text
-o, --output-dir      输出根目录，默认 data/output
--encoding            输入 TSV 编码，默认 utf-8-sig
--field-profile       字段映射 YAML，也可设置 input.path/input.sheets
--remove-null-size    过滤最终尺码为“无可用尺码”的行
--check-atom          生成非皮卡/皮卡高度压缩表的原子检查结果
--progress-interval   进度输出间隔秒数，默认 10
--no-progress         关闭周期性进度输出
```

### check_atom.py

```powershell
python .\check_atom.py --atom <原子事实表> --compress <压缩表> [参数]
```

常见参数：

```text
--atom              原子事实表 TSV，通常是 *_原子事实表.tsv
--compress          要检查的压缩表 TSV
-o, --output-dir    检查结果输出目录
--encoding          输入 TSV 编码，默认 utf-8-sig
```

## 输入字段

脚本会兼容旧字段名，但推荐全量表使用当前字段。也可以用字段 profile 把 Excel/TSV 的自定义列名映射成标准字段。

非皮卡压缩需要：

```text
品牌
前台车型
结构
版本
年份区间
最终尺码
```

皮卡压缩需要：

```text
品牌
前台车型
版本
年份区间
最终尺码
驾驶室类型
货斗长度_ft
```

兼容规则：

```text
如果没有 前台车型，会尝试使用 车型名 或 车姓名
如果没有 主车型，会用 品牌 + 前台车型 自动生成
如果没有 最终尺码，但有 对应尺码，会使用 对应尺码
开始年 可为空；为空时从 年份区间 左端解析
年份区间 不能为空；为空会直接报错
```

### 使用字段 profile

示例命令：

```powershell
python .\process_tsv.py --field-profile .\field_profile.default.yaml --check-atom
```

profile 示例：

```yaml
input:
  path: data/input/0702.xlsx
  sheets:
    - ALL尺码匹配表
    - TM尺码匹配表
    - HNT尺码匹配表

columns:
  品牌:
    - MAKE
    - 品牌
  前台车型:
    - MODEL
    - 前台车型
  年份区间:
    - YEAR
    - 年份区间
  最终尺码:
    - 确认尺码
    - 最终尺码
    - 对应尺码
  驾驶室类型:
    - CAB
    - 驾驶室类型
  货斗长度_ft:
    - BED
    - 货斗长度_ft

derived:
  主车型:
    join:
      - 品牌
      - 前台车型
    sep: " "

defaults:
  分类: ""
```

左边必须是脚本标准字段名，右边可以写你的输入文件实际列名。某个标准字段已经存在时会直接使用它；不存在时才会按候选列名从上到下查找。`主车型` 可以通过 `derived` 自动由 `品牌 + 前台车型` 生成。

## 输出结果怎么看

`compress` 目录：

```text
*_非皮卡无损压缩表.tsv    最保守的非皮卡压缩表，不应制造原表不存在的原子事实
*_非皮卡高度压缩表.tsv    更高压缩率的非皮卡表，会经过原子事实校验逻辑
*_皮卡无损压缩.tsv        最保守的皮卡压缩表
*_皮卡高度压缩表.tsv      会尝试闭合年份、合并 BED 范围的皮卡压缩表
*_原子事实表.tsv          从输入展开出的原子事实基准
*_压缩log.tsv             每个车型压缩、合并、fallback 的过程日志
```

`check` 目录：

```text
*_非皮卡原子检查.tsv      非皮卡原子事实命中高度压缩表的检查结果
*_皮卡原子检查.tsv        皮卡原子事实命中高度压缩表的检查结果
```

项目输出根目录还会生成一个 `*_output.xlsx`，把主要结果放进同一个工作簿，方便人工查看。

检查表中重点看 `结果` 和 `原因`。如果有未命中、重复命中或尺码不一致，需要回到原始数据或压缩规则排查。

## 压缩逻辑简述

非皮卡会先生成无损压缩表，再尝试在同一品牌车型内继续合并年份、结构和版本。候选合并必须能通过原子事实检查：原表中的一条事实不能被压缩结果匹配到多个尺码，也不能匹配到错误尺码。

皮卡无损压缩按车型、版本、驾驶室、货斗长度和尺码合并连续年份。皮卡高度压缩会继续尝试闭合年份空洞、合并 BED 范围，但如果候选范围内存在不同尺码冲突，会 fallback 到更保守的结果。
