# 街道分报告生成器

## 用法

直接运行：

```powershell
python app.py
```

命令行模式也可用：

```powershell
python app.py --input "E:\Haidian\input\a.xlsx" "E:\Haidian\input\b.docx" --output "E:\Haidian\out"
```

如需显式指定模板：

```powershell
python app.py --input "E:\Haidian\input\a.xlsx" --output "E:\Haidian\out" `
  --detail-template "E:\Haidian\Haidian-Report-Producer\街道案件明细表模板.xlsx" `
  --report-template "E:\Haidian\Haidian-Report-Producer\街道环境建设管理工作运行情况分析报告模板.docx"
```

## 界面

- 左侧依次是两个模板文件栏、输出目录、已上传文件列表
- “添加文件”支持一次多选
- 右侧按钮竖向排列：添加文件、生成、仅生成表格提取表、打开输出目录、清空
- 进度条为绿色，生成完成后显示用时

## 输入规则

- 支持 `.xlsx`、`.docx`、`.doc`、`.wps`
- 主数据 Excel 文件名必须包含 `现场检查用数据`
- Word/WPS 文件名可以很乱，程序会按表格标题和固定分类关键词识别内容
- 建议 Word/WPS 文件名尽量包含这些关键词之一：
  - 道路清扫保洁
  - 垃圾管理
  - 城市家具
  - 背街小巷
  - 网格化主动治理
  - 综合执法

## 输出

每个街道生成一个子目录，包含：

- `{月份}{街道}案件明细表.xlsx`
- `{月份}{街道}环境建设管理工作运行情况分析报告.docx`

根目录还会生成：

- `专项表格汇总.xlsx`

## 专项表格汇总

- 每个 sheet 自动适应列宽
- sheet 名优先使用源文件中的表格标题
- 没有标题时，回退为 `文件名-序号`

## 模板与逻辑

- 两个模板文件分开选择，互不影响
- 模板文件只读取，不修改
- 报告正文按内容动态生成，空的专题会直接省略
- 被省略的专题后续编号会自动重排
- 动态章节由 `report_sections` 统一控制，城市家具、背街小巷、网格、综合执法、工作建议会按实际内容顺延编号

## 发布目录

- `街道分报告生成器.exe` 放在发布目录根目录
- `templates` 存放随程序发布的模板文件
- `docs` 存放程序使用指南、程序逻辑、README 等说明文件
- `_internal` 存放运行依赖文件，不需要手动打开或修改

## 注意

- 临时文件不写入 `C:`
- 修改代码后请同步检查 `程序使用指南.txt`、`程序逻辑.md` 等说明文件
