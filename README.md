# auto_fapiao - 发票自动重命名和校验工具

## 这个工具做什么？

帮你批量完成两件事：

1. **自动重命名发票PDF** — 根据合同号索引表，给每个发票文件名前面加上甲方合同号，方便查找和归档。
2. **自动校验发票信息** — 把Excel里填的发票号、日期、金额等信息和PDF里OCR识别出来的结果对比，发现错误自动修复，并告诉你哪些需要手动检查。

## 环境要求

- **Python 3.12**（PaddlePaddle 仅支持 3.12，不支持 3.13+）
- **uv**（Python包管理工具，[安装方法](https://docs.astral.sh/uv/getting-started/installation/)）

### Windows 额外要求

- **Visual C++ Redistributable** — PaddlePaddle 依赖 VC++ 运行时。如果安装后运行报 DLL 缺失错误，请安装 [vc_redist.x64.exe](https://aka.ms/vs/17/release/vc_redist.x64.exe)（微软官方下载，约2MB）

## 安装

```bash
# 进入项目目录
cd auto_fapiao

# 同步所有依赖（包括PaddleOCR、PyMuPDF、openpyxl等）
uv sync

# 如果需要跑测试，额外装开发依赖
uv sync --extra dev
```

> **Windows 备选安装方式：** 如果 `uv sync` 很慢，可以用 `uv pip install -r requirements.txt` 直接安装依赖。

## 使用流程

### 第一步：准备两个Excel文件

你手里应该有两个Excel：

1. **合同号索引表** — 第A列是乙方合同号，第H列是合同名称（里面包含甲方合同号）
2. **发票验证表** — 第AF列是甲方合同号，BZ/CA/CB/CC/CD列分别是开票数量、发票号、发票日期、税金金额、含税金额

把这两个Excel和所有发票PDF放在同一个文件夹里。

### 第二步：重命名发票文件

```bash
uv run main.py rename --dir ./你的发票文件夹 --excel ./你的合同号索引表.xlsx

# 加 -v 可以看到详细处理过程，方便排查问题
uv run main.py -v rename --dir ./你的发票文件夹 --excel ./你的合同号索引表.xlsx
```

工具会：
- 自动识别每个PDF是哪种类型（文件名里有乙方合同号的，还是纯发票编号的）
- 从PDF里提取需要的信息（如果文件名里没有合同号，会用OCR识别备注栏）
- 查表找到甲方合同号
- 将重命名后的文件复制到 `{原文件夹名}_Renamed/` 文件夹中，原文件不动
- 新文件格式为 `{甲方合同号}-{乙方合同号}_{其他部分}.pdf`

### 第三步：校验发票信息

```bash
# --dir 指向重命名后的文件夹（{原文件夹}_Renamed）
uv run main.py verify --dir ./你的发票文件夹_Renamed --excel ./你的发票验证表.xlsx

# 加 -v 可以看到每个字段的比对详情和置信度
uv run main.py -v verify --dir ./你的发票文件夹_Renamed --excel ./你的发票验证表.xlsx
```

工具会：
- 读取Excel里的发票信息
- 用OCR识别每张PDF发票的实际内容
- 逐项对比：发票号、日期、税金、含税金额
- 如果OCR置信度高（>80%），自动修复Excel里的错误
- 如果OCR置信度低，提示你手动检查
- 输出校验报告，告诉你哪些是对的、哪些改了、哪些需要你自己看

> **注意：** 校验前请先运行重命名，校验命令的 `--dir` 指向重命名后的文件夹（`{原文件夹}_Renamed`）。

## 常见问题

**Q: 首次运行PaddleOCR会下载模型吗？**
A: 是的，首次运行会自动下载PaddleOCR的中文识别模型（约100MB），之后就完全离线运行了。

**Q: 支持哪些发票类型？**
A: 支持增值税专用发票、普通发票等。只要发票上有发票号码、开票日期、税额、价税合计、备注栏里的乙方合同号，就能识别。

**Q: 如果OCR识别不准怎么办？**
A: 工具会自动判断置信度。低于80%的不会自动修改Excel，会在报告里标记"需手动核查"，你自己确认后再改。

**Q: 可以只跑重命名不跑校验吗？**
A: 可以，两个命令是独立的，按需使用。

## 测试

```bash
# 跑全部测试
uv run pytest tests/ -v

# 跑单个测试文件
uv run pytest tests/test_contract.py -v

# 跑某个测试
uv run pytest -k test_extract_non_sz -v
```

## 调试与问题排查

如果运行出问题，加 `-v` 参数获取详细日志：

```bash
uv run main.py -v rename --dir ./pdfs --excel ./合同号索引表.xlsx 2>&1 | tee debug.log
```

把 `debug.log` 文件发过来，就能快速定位问题。日志会显示：
- 每个PDF文件被识别为哪种类型
- OCR提取了哪些字段、置信度是多少
- 合同号匹配是否成功
- 哪些字段不一致、是否自动修复
