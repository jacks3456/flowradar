# KRX Money Flow Reporter

从 KRX 官方数据接口抓取韩国市场投资者资金流，并生成 Excel 报表、单文件 HTML 报表和可公开部署的静态网站。

## 功能

- KOSPI 外资、机构、散户每日净买卖金额，单位为亿韩元
- 个股外资、机构、散户每日净买卖股数和金额
- 默认包含三星电子 `005930`、SK 海力士 `000660`
- 自动生成 Excel，多 sheet 保存原始表和汇总表
- 自动生成 HTML 网页仪表盘，可直接用浏览器打开
- 自动生成 `site/data/latest.json`，作为公开网站的数据源
- 自动绘制 KOSPI 外资、机构、散户累计资金流曲线
- 自动计算最近连续净买入/净卖出天数

## 安装

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## KRX 登录

当前 KRX 数据接口会检查登录态。EN/SNS 账号建议使用已登录浏览器的 Cookie：

```bash
export KRX_COOKIE='JSESSIONID=...; __smVisitorID=...; mdc.client_session=true'
python krx_money_flow.py
```

也可以运行时传入：

```bash
python krx_money_flow.py --krx-cookie 'JSESSIONID=...; __smVisitorID=...; mdc.client_session=true'
```

如果你有可用的 KRX 普通账号密码，也可以设置：

```bash
export KRX_ID="你的账号"
export KRX_PW="你的密码"
```

也可以运行时传入：

```bash
python krx_money_flow.py --krx-id 你的账号 --krx-password 你的密码
```

## 使用

默认抓取最近约 90 天：

```bash
python krx_money_flow.py
```

生成日股 JPX 周频投资者资金流数据：

```bash
python jpx_money_flow.py
```

该脚本会抓取 JPX `Trading by Type of Investors` 的周频 `Value` Excel，默认读取最近 52 周，并生成：

```text
site/data/japan/latest.json
```

日股口径为：`Foreigners -> 外资`，`Individuals -> 散户`，`Institutions + Proprietary + Securities Cos. -> 机构`，单位为亿日元。

指定日期和股票：

```bash
python krx_money_flow.py --start 20250101 --end 20250704 --tickers 005930,000660
```

指定输出文件：

```bash
python krx_money_flow.py --output reports/my_report.xlsx
```

指定网页输出文件：

```bash
python krx_money_flow.py --html-output reports/my_dashboard.html
```

指定公开网站目录：

```bash
python krx_money_flow.py --site-dir site
```

本地预览网站：

```bash
python -m http.server 8765
```

然后打开：

```text
http://127.0.0.1:8765/site/
```

## 输出

默认输出到：

```text
reports/krx_money_flow_开始日期_结束日期.xlsx
reports/krx_money_flow_开始日期_结束日期.html
site/data/latest.json
```

Excel 包含：

- `summary`：最近值和连续净买入/净卖出天数
- `kospi_market_eok_krw`：KOSPI 每日净买卖金额，单位亿韩元
- `股票代码_股票名_shares`：个股每日净买卖股数
- `股票代码_股票名_krw`：个股每日净买卖金额，单位韩元
- `charts`：KOSPI 累计资金流曲线

HTML 网页包含：

- 市场外资、机构、散户最近净买卖金额和连续状态
- 外资、机构、散户累计资金流折线图
- 市场和个股最近 30 条明细，可在网页内切换
- 数据和绘图脚本均内嵌在 HTML 中，不依赖外部 CDN

公开网站包含：

- `site/index.html`：可部署到静态托管平台的网页
- `site/data/latest.json`：脚本生成的数据源
- 前端只读取公开 JSON，不需要 KRX 登录信息

## 公开部署

这个项目已经包含 GitHub Pages 工作流：

```text
.github/workflows/update-site.yml
```

部署步骤：

1. 把仓库推到 GitHub。
2. 在仓库 Settings -> Secrets and variables -> Actions 中新增 `KRX_COOKIE`；如果要用账号密码模式，则新增 `KRX_ID` 和 `KRX_PW`。
3. 在仓库 Settings -> Pages 中选择 GitHub Actions 作为来源。
4. 到 Actions 手动运行 `Update FlowRadar site`，之后工作流会在交易日定时更新。

工作流会运行脚本生成最新 `site/data/latest.json`，然后把 `site/` 发布到 GitHub Pages。

## 绑定自己的域名

当前域名：

```text
flowradar.cc
```

推荐绑定方式：

1. 在 GitHub 仓库 Settings -> Pages 中，把 Custom domain 设置为：

```text
flowradar.cc
```

2. 在域名服务商 DNS 里添加 4 条 `A` 记录：

```text
@  A  185.199.108.153
@  A  185.199.109.153
@  A  185.199.110.153
@  A  185.199.111.153
```

3. 如果也想支持 `www.flowradar.cc`，添加一条 `CNAME`：

```text
www  CNAME  你的GitHub用户名.github.io
```

4. 等 DNS 生效后，在 GitHub Pages 里启用 Enforce HTTPS。

注意：

- `site/CNAME` 已经写入 `flowradar.cc`，方便静态站发布时携带域名配置。
- DNS 变更可能需要最长 24 小时生效。
- 不要添加 `*.flowradar.cc` 这种通配符 DNS 记录，容易带来子域名被接管风险。

数据表里的 `合计` 为 `机构 + 散户 + 外资`。`其他法人` 单独保留，不计入该合计列。

## 数据来源

脚本使用 KRX `data.krx.co.kr` 的 JSON 接口：

- `dbms/comm/finder/finder_stkisu`：股票代码和 ISIN 映射
- `dbms/MDC/STAT/standard/MDCSTAT02202`：市场投资者日别交易实绩
- `dbms/MDC/STAT/standard/MDCSTAT02302`：个股投资者日别交易实绩
