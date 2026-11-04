# 新游榜单数据抓取与复现脚本

这套脚本对应「gameres 新游榜（国内、近 30 天）」的数据采集与分析：`fetch_ranking.py` 负责抓取全量榜单并导出 `date.xlsx`；`fetch_game_details.py` 读取榜单前 30 名，逐个抓取详情并生成「游戏名.txt」；`analysis_ranking.py` 读取 `date.xlsx` 与详情 txt，复算品类占比、开发商 CR5/CR10、评价数 Top10 占比和开测月份分布。

## 文件说明

| 文件 | 作用 |
| --- | --- |
| `fetch_ranking.py` | 分页抓取榜单（`p` / `ps=50`、`date_range=30`、`type_range=2`），导出 18 列中文表头的 `date.xlsx`。 |
| `fetch_game_details.py` | 按排名取前 30，抓取详情接口，生成 30 个分区块 txt（基本信息、开测数据、厂商与发行、平台信息）。 |
| `analysis_ranking.py` | 读取 `date.xlsx`，输出品类占比、开发商 CR5/CR10、评价数 Top10 占比；并从详情 txt 解析开测月份分布。 |
| `requirements.txt` | 依赖清单：requests、pandas、openpyxl。 |
| `config.example.json` | 配置文件模板，复制为 `config.json` 后使用。 |

## 安装依赖

需要 Python 3.9 及以上版本。

```powershell
cd c:\Users\21109\Desktop\gameres
python -m pip install -r scripts\requirements.txt
```

## 配置请求头与凭据

脚本不硬编码 Cookie / Token：

1. 把 `scripts\config.example.json` 复制为 `scripts\config.json`；
2. 如需登录态，在 `config.json` 中填写 `cookie` 或 `token`，并把 `require_auth` 设为 `true`；
3. 也可用环境变量临时注入（优先级高于配置文件）：`GAMERES_COOKIE`、`GAMERES_TOKEN`。

```powershell
Copy-Item scripts\config.example.json scripts\config.json
```

当前榜单与详情接口都是公开接口，Cookie / Token 可以留空，脚本会以匿名方式访问。若设置了 `require_auth: true` 但没有提供凭据，脚本会直接报错并提示配置方式。配置文件缺失或 JSON 格式错误时，也会给出明确提示。

## 运行顺序与预计耗时

```powershell
# 1. 抓榜单，生成 date.xlsx
python scripts\fetch_ranking.py

# 2. 抓前 30 名详情，生成 30 个 txt
python scripts\fetch_game_details.py --out-dir top30

# 3. 复算核心指标（只打印，不写文件）
python scripts\analysis_ranking.py --details-dir top30
```

- 第 1 步：11 页、537 条，按 1.5 秒间隔，约 30–50 秒。
- 第 2 步：30 个详情，按 1.5 秒间隔，约 60–90 秒。
- 第 3 步：本地计算，约 1–3 秒。

脚本内路径默认相对脚本位置计算，不依赖固定盘符；也可用 `--out`、`--ranking`、`--out-dir`、`--details-dir` 覆盖。

## 输出位置

- `date.xlsx`：默认输出到项目根目录（即 `scripts` 的上一级），可用 `--out` 指定。
- 30 个「游戏名.txt」：默认输出到项目根目录；建议用 `--out-dir top30` 保持目录整洁。非法文件名字符会替换为全角字符，重名自动加 `_2`。
- `analysis_ranking.py`：只在终端打印指标，便于直接对照报告数字。

## 数据口径与注意

- 榜单固定为国内游戏（`type_range=2`）、近 30 天（`date_range=30`），单页上限实测为 50（`ps=51` 起返回空）。
- 抓取按页码递增，某页返回不足 50 条或为空即停止；第 12 页为空表示已抓全。
- 缺失字段统一写「页面未提供」，不做估算或填充。
- 请求间隔约 1.5 秒，失败最多重试 4 次；请勿把 `sleep_seconds` 调到过低，避免对站点造成压力。
- 品类占比按「游戏标签中第一个可识别的品类词」归类；`analysis_ranking.py` 的默认详情目录是 `top30`，如果 txt 放在别处，请用 `--details-dir` 指定。
