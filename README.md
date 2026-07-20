# A_报表

跨境电商多平台利润报表自动化流水线。脚本以 Excel 为中间载体，按模块（A→M）依次读取、加工、写回桌面上的报表周期目录。

覆盖平台：Amazon、OTTO、REAL、ManoMano、DLZ（独立站）、TEMU、LeroyMerlin 等。

---

## 路径配置

| 变量 | 定义位置 | 说明 |
|------|----------|------|
| `DESKTOP_ROOT` | `A0_设置_时间段/A0_paths.py` | 桌面根目录，换电脑或用户名时只改此处 |
| `folder_name` / `shared_date` | `A0_设置_时间段/A0_set_date.py` | 报表类型（`日报` / `月报`）与统计时间段 |
| `REPORT_PERIOD_DIR` | `A0_paths.py` | `{DESKTOP_ROOT}\{folder_name}{shared_date}`，如 `日报7.1-7.13` |
| `BTH_ALL_SKU_DETAIL_PATH` | `A0_paths.py` | 网络盘 SKU 明细 |
| `MONTH_GOAL_DIR` | `A0_paths.py` | 月目标拆解表目录（`F:\月目标拆解及跟进`） |

---

## 模块总览

| 目录 | 说明 | 入口 |
|------|------|------|
| `A0_设置_时间段/` | 全局配置：日期、汇率、路径 | `A0_set_date.py`、`A0_paths.py` |
| `A_TEMU_计算_订单总金额/` | TEMU 订单金额、日报文件检查 | `runAll_A.py` |
| `B_订单统计_sale_resend/` | 订单统计主流程（ERP 映射→站点→头程→交易明细→尾程→采购成本） | `runAll_B.py` |
| `C_退款/` | 全平台退款（RMA + AMZ 退款 + VAT/平台费） | `runAll_C.py` |
| `C1_Mano(月)/` | MANO VAT / 佣金 / 仓租（仅月报） | 各 `C1_*`、`C2_*` |
| `D_广告/` | OTTO / REAL / MANO / DLZ 广告费 | `runAll_D.py` |
| `E_秒杀_只有AMZ有/` | AMZ 秒杀费用 | `E1`、`E2` |
| `F_测评/` | 测评费用 | `runAll_F.py` |
| `G_二次上架/` | 退货二次上架 | `runAll_G.py` |
| `G2_TEMU罚款/` | TEMU 罚款 | `G2_1`~`G2_3` |
| `H_AMZ_利润报表_OTTO_客户经理费/` | AMZ SellerSku 利润 / OTTO 客户经理费 / TEMU 罚款分摊 | `runAll_H.py` |
| `I_输出_报表-儿子_映射_商品ID/` | 儿子报表输出、商品 ID 映射 | `I2`、月报 `I1` |
| `J_AMZ_仓租/` | FBA 仓租（日报/月报分支） | `J1`~`J3` |
| `K_仓租_映射产品信息/` | 海外仓仓租（鸿羽/4PX）、产品信息映射、分销处理 | `runAll_K.py` |
| `L_店铺租金摊分/` | 店铺月租按站点日均摊 | `L1` |
| `M_毛利_销售负责人_表头排序/` | 毛利计算、销售负责人映射、最终表头排序 | `run_all_gross_profit.py` |
| `Z_method/` | 公共工具模块 | — |

**推荐执行顺序：** A0 → A → B → C → D → E → F → G → H → I → J → K → L → M（月报另含 `C1_Mano(月)`、`J/月报` 分支）

---

## Z_method 公共工具

| 模块 | 功能 |
|------|------|
| `style.py` | 终端 ANSI 颜色 |
| `sku_映射.py` | 通用 SKU 映射（读映射表 → 字典 → 写入主表） |
| `platform_shop.py` | 店铺→站点/平台/VAT/佣金映射（MySQL 优先，Excel 兜底） |
| `split_rows_data_SKU.py` | 组合 SKU 拆行（含 `+` 或 `,`），按比例分摊金额 |
| `cang_zu_site.py` | 库存"平台"标签→标准"站点"映射 |
| `time.py` | 日期工具：`get_month_range()` |
| `runall_utils.py` | runAll 脚本通用工具：编码设置、子进程执行、脚本排序、输出路径提取 |

---

## 桌面静态映射表（`{DESKTOP_ROOT}\*.xlsx`）

| 文件名 | 用途 | 引用模块 |
|--------|------|----------|
| `信息-映射.xlsx` | 产品 / SKU 信息映射 | M2、M3、M4、K5、I1 |
| `VAT、平台费-映射.xlsx` | 站点 VAT / 平台佣金（DB 兜底） | `Z_method/platform_shop.py` |
| `广告-SKU关系对应.xlsx` | 广告 SKU 与商品对应 | D1、D2、D3 |
| `castorama - SKU类目佣金比例.xlsx` | Castorama 类目佣金 | C5、F1 |
| `MANO-MF 尾程.xlsx` | MANO MF 尾程费用 | B5 |
| `仓租-SKU映射.xlsx` | 海外仓 SKU 映射 | K1 |
| `DLZ-FR_广告分摊sku.xlsx` | DLZ FR 广告分摊 | D4_4 |
| `手动-二次映射.xlsx` | TEMU 无映射单价兜底 | A2 |
| `月租总摊分.xlsx` | 月租按站点摊分规则 | L1 |

---

## 报表周期目录结构（`{REPORT_PERIOD_DIR}\`）

由 `A1_日报文件检查.py` 从网络盘复制源文件，各模块在此读写中间结果。

```
{folder_name}{shared_date}/
├── 订单统计/          # 主流水线，(已完成-N) 前缀递增
├── RMA/              # 退款中间文件
├── transaction交易明细/  # AMZ 交易明细
├── SellerSku利润报表/    # AMZ 利润报表
├── 广告/
│   ├── OTTO/
│   ├── REAL/
│   ├── MANO/
│   └── DLZ/
├── 测评表/
├── 秒杀/
├── 二次上架/
├── TEMU-罚款/
└── 仓租/
    ├── FBA仓租/
    ├── 鸿羽/
    ├── 4PX/
    ├── mano/
    └── mano-vat/
```

---

## 外部依赖

### 数据库（MySQL）

`sales_order_shipped`、`product_sku`、`temu_order_item`、`mano_mmf_price`、`platform_shop`、`product_sku_mapping`

### 网络盘

| 路径 | 用途 |
|------|------|
| `\\Betohow\数据报表\数据库\BTH全部SKU明细-*.xlsx` | 采购成本、SKU 基础数据 |
| `\\Betohow\数据报表\数据库\产品信息库2025.xlsx` | 产品信息映射 |
| `\\Betohow\数据报表\报表自动化下载\` | 日报源文件 |
| `\\Betohow\数据报表\RPA\` | 二次上架等 RPA 查询结果 |
| `\\Betohow\数据报表\2-定价表\` | 非 MF 尾程定价 |
| `F:\月目标拆解及跟进\` | Amazon 销售负责人映射 |

### Python 包

`pandas`、`openpyxl`、`pymysql`、`numpy`、`chardet`

---

## 快速开始

1. 修改 `A0_设置_时间段/A0_paths.py` 中的 `DESKTOP_ROOT`（换电脑时）。
2. 修改 `A0_设置_时间段/A0_set_date.py`：设置 `folder_name`（日报/月报）、`shared_date`、汇率。
3. 确认桌面静态映射表齐全。
4. 执行：

```powershell
python "A_报表/A_TEMU_计算_订单总金额/runAll_A.py"
python "A_报表/B_订单统计_sale_resend/runAll_B.py"
python "A_报表/C_退款/runAll_C.py"
python "A_报表/D_广告/runAll_D.py"
# E、F、G、H、I、J、K、L、M 同理
python "A_报表/M_毛利_销售负责人_表头排序/run_all_gross_profit.py"
```

---

## 其他文件

| 文件 | 说明 |
|------|------|
| `updateLog.md` | 变更记录 |
| `OKR月目标拆分.py` | 月度 OKR 目标拆分工具 |
| `K_仓租_映射产品信息/海外仓仓租分摊逻辑说明.md` | 海外仓仓租分摊逻辑文档 |
| `G_二次上架/README_G2优化说明.md` | G2 优化说明 |
