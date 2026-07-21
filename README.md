# reportA

跨境电商多平台利润报表自动化流水线。脚本以 Excel 为中间载体，按模块 **A→M** 依次读写桌面上的报表周期目录。

覆盖平台：Amazon、OTTO、REAL、ManoMano、DLZ（独立站）、TEMU、LeroyMerlin 等。

---

## 项目结构

```
reportA/
├── config/                 # 日期、汇率、路径；db_config.example.json
├── database/               # MySQL 连接与表结构（原 reportPRA 拷贝）
├── common/                 # 公共工具（SKU 映射、店铺映射、runAll 工具等）
├── modules/                # 业务流水线 A→M
│   ├── A_temu_order_amount/
│   ├── B_order_stats_sale_resend/
│   ├── C_refund/  ·  C1_mano_monthly/   # C1 仅月报
│   ├── D_ads/
│   ├── E_amz_seckill_cost/
│   ├── F_review_cost/
│   ├── G_returned_reshelf/  ·  G2_temu_penalty/
│   ├── H_amz_profit_otto_manager_fee/
│   ├── I_son_report_product_id/
│   ├── J_amz_storage_fee/              # daily / monthly
│   ├── K_storage_fee_product_info/
│   ├── L_shop_rent_allocation/
│   └── M_gross_profit_owner_headers/
├── app/                    # 编排入口、HY 派送费试算
├── api/hy_oms/             # 鸿羽 OMS SOAP 对接
├── runtime/local/          # 本机费用覆盖 JSON（不入库）
├── tools/                  # 独立工具（如 OKR 拆分）
├── docs/                   # 模块说明
└── ensure_project_root.py  # 将仓库根加入 sys.path
```

---

## 快速开始

1. 修改 `config/A0_paths.py` 中的 `DESKTOP_ROOT`（换电脑时）。
2. 修改 `config/A0_set_date.py`：`folder_name`（`日报` / `月报`）、汇率等；日期区间按类型自动计算。
3. 确认桌面静态映射表齐全（见下文）。
4. 在项目根目录执行：

```powershell
# 整条流水线（有 runAll 的模块：A B C D F G H K M）
python -m app.run_pipeline
python -m app.run_pipeline --only A,B,C

# 单模块
python modules/A_temu_order_amount/runAll_A.py
python modules/B_order_stats_sale_resend/runAll_B.py
python modules/C_refund/runAll_C.py
python modules/D_ads/runAll_D.py
python modules/F_review_cost/runAll_F.py
python modules/G_returned_reshelf/runAll_G.py
python modules/H_amz_profit_otto_manager_fee/runAll_H.py
python modules/K_storage_fee_product_info/runAll_K.py
python modules/M_gross_profit_owner_headers/run_all_gross_profit.py

# 无 runAll 的步骤需单独跑脚本，例如：
python modules/E_amz_seckill_cost/E1_秒杀_映射_站点识别码.py
python modules/L_shop_rent_allocation/L1_计算月租_合并_分摊_订单统计.py
```

**推荐顺序：** A0 配置 → A → B → C → D → E → F → G → H → I → J → K → L → M  
月报另含 `C1_mano_monthly`、`J_amz_storage_fee/monthly`，以及若干标注「手动操作」的脚本。

---

## 路径配置

| 变量 | 定义位置 | 说明 |
|------|----------|------|
| `DESKTOP_ROOT` | `config/A0_paths.py` | 桌面根目录 |
| `folder_name` / `shared_date` | `config/A0_set_date.py` | `日报` / `月报` 与统计时间段 |
| `REPORT_PERIOD_DIR` | `A0_paths.py` | `{DESKTOP_ROOT}\{folder_name}{shared_date}` |
| `BTH_ALL_SKU_DETAIL_PATH` | `A0_paths.py` | 网络盘 SKU 明细 |
| `MONTH_GOAL_DIR` | `A0_paths.py` | 月目标拆解表目录 |

---

## 模块总览

| 目录 | 说明 | 入口 |
|------|------|------|
| `config/` | 日期、汇率、路径 | `A0_set_date.py`、`A0_paths.py` |
| `A_temu_order_amount` | TEMU 订单金额、日报文件检查 | `runAll_A.py` |
| `B_order_stats_sale_resend` | 订单统计主流程（站点→头程→交易→尾程→采购成本） | `runAll_B.py` |
| `C_refund` | 全平台退款、VAT/平台费 | `runAll_C.py` |
| `C1_mano_monthly` | MANO VAT / 佣金 / 仓租（仅月报） | `C1_*`、`C2_*` |
| `D_ads` | OTTO / REAL / MANO / DLZ 广告 | `runAll_D.py` |
| `E_amz_seckill_cost` | AMZ 秒杀费用 | `E1`、`E2` |
| `F_review_cost` | 测评费用 | `runAll_F.py` |
| `G_returned_reshelf` | 退货二次上架 | `runAll_G.py` |
| `G2_temu_penalty` | TEMU 罚款 | `G2_1`~`G2_3` |
| `H_amz_profit_otto_manager_fee` | SellerSku 利润 / OTTO 客户经理费 / 其他费用分摊 | `runAll_H.py` |
| `I_son_report_product_id` | 儿子报表、商品 ID 映射 | `I2`、`monthly/I1` |
| `J_amz_storage_fee` | FBA 仓租（`daily` / `monthly`） | `J1`~`J3` |
| `K_storage_fee_product_info` | 海外仓仓租、产品信息、分销 | `runAll_K.py` |
| `L_shop_rent_allocation` | 店铺月租按站点摊分 | `L1` |
| `M_gross_profit_owner_headers` | 毛利、销售负责人、表头排序 | `run_all_gross_profit.py` |
| `app/` | 流水线编排、HY 派送费试算 | `run_pipeline`、`delivery_fee_hy` |
| `api/hy_oms/` | 鸿羽 OMS | `python -m api.hy_oms.smoke_test` |
| `tools/` | 独立工具 | `OKR月目标拆分.py` |

---

## common 公共工具

| 模块 | 功能 |
|------|------|
| `style.py` | 终端 ANSI 颜色 |
| `sku_mapping.py` | 通用 SKU 映射 |
| `platform_shop.py` | 店铺→站点/平台/VAT/佣金（MySQL 优先，Excel 兜底） |
| `split_rows_data_SKU.py` | 组合 SKU 拆行，按比例分摊金额 |
| `cang_zu_site.py` | 库存「平台」标签→标准「站点」 |
| `time.py` | `get_month_range()` 等 |
| `runall_utils.py` | runAll：编码、子进程、脚本排序、输出路径提取 |

---

## 桌面静态映射表（`{DESKTOP_ROOT}\*.xlsx`）

| 文件名 | 用途 |
|--------|------|
| `信息-映射.xlsx` | 产品 / SKU 信息 |
| `VAT、平台费-映射.xlsx` | 站点 VAT / 佣金（DB 兜底） |
| `广告-SKU关系对应.xlsx` | 广告 SKU 对应 |
| `MANO-MF 尾程.xlsx` | MANO MF 尾程（也可经 `runtime/local` 覆盖） |
| `仓租-SKU映射.xlsx` | 海外仓 SKU |
| `DLZ-FR_广告分摊sku.xlsx` | DLZ FR 广告分摊 |
| `手动-二次映射.xlsx` | TEMU 无映射单价兜底 |
| `月租总摊分.xlsx` | 月租摊分规则 |

Castorama 类目佣金已迁至 `runtime/local/castorama_commission.json`（F1 / C5），桌面不再需要 `castorama - SKU类目佣金比例.xlsx`。

---

## 报表周期目录（`{REPORT_PERIOD_DIR}\`）

由 A 模块从网络盘复制源文件；后续模块在此读写中间结果（文件名前缀 `(已完成-N)` 递增）。

```
{folder_name}{shared_date}/
├── 订单统计/
├── RMA/
├── transaction交易明细/
├── SellerSku利润报表/
├── 广告/          # OTTO · REAL · MANO · DLZ
├── 测评表/
├── 秒杀/
├── 二次上架/
├── TEMU-罚款/
└── 仓租/          # FBA · 鸿羽 · 4PX · mano · mano-vat
```

---

## 外部依赖

**Python：** `pandas`、`openpyxl`、`pymysql`、`DBUtils`、`numpy`、`chardet`

**数据库（MySQL）：** `sales_order_shipped`、`product_sku`、`temu_order_item`、`mano_mmf_price`、`platform_shop`、`product_sku_mapping` 等（连接经本仓库 `database.db_connection`；配置见 `config/db_config.example.json` → 复制为 `db_config.json`）

**网络盘 / 本机盘：**

| 路径 | 用途 |
|------|------|
| `\\Betohow\数据报表\数据库\BTH全部SKU明细-*.xlsx` | 采购成本、SKU 基础数据 |
| `\\Betohow\数据报表\数据库\产品信息库*.xlsx` | 产品信息映射 |
| `\\Betohow\数据报表\报表自动化下载\` | 日报源文件 |
| `\\Betohow\数据报表\RPA\` | 二次上架等 RPA 结果 |
| `\\Betohow\数据报表\2-定价表\` | 非 MF 尾程定价 |
| `F:\月目标拆解及跟进\` | Amazon 销售负责人映射 |

**鸿羽 OMS：** 凭证写在 `api/hy_oms/config.py`（`APP_TOKEN` / `APP_KEY`）。

---

## 其他

| 路径 | 说明 |
|------|------|
| `docs/海外仓仓租分摊逻辑说明.md` | 海外仓仓租分摊 |
| `docs/F_测评说明.md` | 测评模块 |
| `runtime/local/` | 本机费用 JSON 覆盖（已 gitignore） |
| `tools/OKR月目标拆分.py` | 月度 OKR 目标拆分 |
