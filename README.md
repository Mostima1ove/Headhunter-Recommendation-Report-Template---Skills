# candidate-report

把候选人简历做成一份可以直接发给客户的推荐报告：A4 两页、宋体、带公司 logo 和页脚保密声明，用 python-docx 从零生成，不依赖任何二进制模板文件。

写给猎头 / 咨询公司用：换一家公司只要改一个 JSON，版式不用碰。

## 长什么样

```
客户名称：           XX 基金
推荐职位：           量化研究员
工作地点：           上海
推荐日期：           2026.11.23

推荐理由
■ 四年美股量化研究与实盘经验，独立负责中频多空策略的因子构建与信号融合；
■ 有 NLP 与大模型经验，把每日上万篇新闻做成 alpha 信号；
■ 现已回国常驻上海，可即刻到岗。

基本信息 / 教育背景 / 专业资质与技能 / 工作经验 …（见 examples）
```

## 用

```bash
pip install python-docx

# 1. 配一次公司信息（logo、页脚声明、页眉文字）
cp examples/config.example.json my-config.json
# 编辑 my-config.json

# 2. 写候选人数据（position 和 jd 必填）
cp examples/data.example.json 张三.json
# 编辑 张三.json

# 3. 生成
python scripts/report.py --config my-config.json --data 张三.json --out 张三-XX基金-量化研究员.docx
```

`position` 必填（不知道推什么岗就没法写）；`jd` 有就填，推荐理由会对着职责逐条写，没有就按简历综合总结，脚本会给一句提示但不拦。先试着要齐三样：**职位名、客户公司、岗位介绍原文**。

## 用你自己公司的模板

不想用内置版式的话，拿一份你们已经做完的报告（空白的最好，带 logo 和页脚那种）丢进来：

```bash
python scripts/inspect_template.py --input 你们的模板.docx --logo-dir . --out my-config.json
```

它会扒出页面尺寸、页边距、页眉页脚距离、字体字号、页脚保密声明，并把页眉 logo 导出来，直接生成一份能用的 config。之后 `report.py --config my-config.json` 就按你们的版式出报告。

制表位和条目符号在多数模板里是 Word 默认值、XML 里查不到，会走默认值——抬头或日期列对不齐时手调 config 里 `layout.tab_field_cm / tab_edu_cm / tab_job_cm`。

## 渲染检查

Word 里看不出来的问题（折行、分页、第 2 页只剩两行），渲染成图一眼就能看见：

```bash
soffice --headless --convert-to pdf --outdir ./out 报告.docx
pdftoppm -png -r 144 ./out/报告.pdf ./out/page
```

## 配置项

`config.json`：

| 字段 | 说明 |
| --- | --- |
| `company` | 没有 logo 时显示在页眉右侧的公司名 |
| `logo` | logo 图片路径（png/jpg），设了就优先显示 logo |
| `logo_width_cm` | logo 宽度，默认 3.2 |
| `footer` | 页脚保密声明，带边框 |

版式常量在 `scripts/report.py` 顶部：页面、边距、字体、字号、制表位、条目符号，改一处全局生效。

`data.json` 关键字段：

| 字段 | 说明 |
| --- | --- |
| `position` | **必填**。职位名，照抄客户 JD，别自己改写 |
| `jd` | 岗位介绍原文。填了三条理由对着它写；不填就按简历综合总结 |
| `client` | 客户公司，推多家时可留空（报告里显示「—」） |
| `reasons` | 三条推荐理由，每条压在一行内 |
| `basic` / `education` / `skills` / `experience` | 其余内容，缺字段自动填「—」 |

## 设计取舍

- **不用 docx 当模板**。改模板文件里的占位文本很容易把格式带坏，而且用户一改模板，脚本里的段落索引就全错（这是踩过的坑）。脚本里直接写版式，稳定。
- **条目符号用 `■` 不用 `⚫`**。`⚫`（U+26AB）多数字体没有，渲染会掉成黑块。
- **行距不设固定值**。设成 12pt 会把两页的内容压成一页，看着像缩水。
- **两页是硬约束**。超了砍内容，不缩字号。

## License

MIT
