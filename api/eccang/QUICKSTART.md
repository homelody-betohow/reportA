# 易仓 ERP API 对接 - 快速入门

## 📦 模块说明

本模块为易仓 ERP 开放平台的 Python API 客户端，提供了完整的接口封装和易用的调用方式。

**位置**：`api/eccang/`

**文档地址**：https://open.eccang.com

## 🚀 快速开始

### 第 1 步：配置凭证

编辑 `/api/eccang/config.py`，填写 API 凭证：

```python
# ─── 易仓 ERP 开放平台凭证 ────
APP_KEY = "你的APP_KEY"
APP_SECRET = "你的APP_SECRET"
SERVICE_ID = "你的SERVICE_ID"
```

**获取凭证**：
1. 访问 https://open.eccang.com 并登录
2. 进入"应用管理" → "新增应用" → 选择"服务商应用"
3. 填写应用信息，设置 IP 白名单
4. 审核通过后获取 `APP_KEY` 和 `APP_SECRET`
5. 联系易仓客户授权获取 `SERVICE_ID`

### 第 2 步：测试连接

```bash
python -m api.eccang.cli test
```

如果配置正确，将显示：
```
✓ 配置加载成功：app_key=xxxxxxxx***, service_id=xxxxxxxx
✓ 请求地址：http://openapi-web.eccang.com/openApi/api/unity
✓ 连接成功！
```

### 第 3 步：开始使用

#### 方式 1：使用封装好的方法（推荐）

```python
from api.eccang import EccangService

# 创建客户端
client = EccangService()

# 获取仓库列表
response = client.get_warehouse_list(page=1, page_size=50)
print(response)

# 获取订单列表
response = client.get_order_list(
    page=1,
    page_size=50,
    start_time="2024-01-01 00:00:00",
    end_time="2024-01-31 23:59:59"
)
print(response)
```

#### 方式 2：调用任意接口

```python
from api.eccang import EccangClient

client = EccangClient()

# 调用任意接口
response = client.call(
    method="getShipAddressBooks",
    body={"page": 1, "page_size": 50}
)
print(response)
```

#### 方式 3：使用命令行工具

```bash
# 获取仓库列表
python -m api.eccang.cli get-warehouses --page 1 --page-size 50

# 获取订单列表
python -m api.eccang.cli get-orders \
    --start-time "2024-01-01 00:00:00" \
    --end-time "2024-01-31 23:59:59"

# 调用任意接口
python -m api.eccang.cli call getShipAddressBooks \
    --body '{"page": 1, "page_size": 50}'
```

## 📚 已封装的常用接口

| Python 方法 | 接口方法 | 说明 |
|------------|----------|------|
| `get_warehouse_list()` | `getWarehouseList` | 获取仓库列表 |
| `get_order_list()` | `getOrderList` | 获取订单列表 |
| `get_order_detail()` | `getOrderDetail` | 获取订单详情 |
| `get_product_list()` | `getProductList` | 获取产品列表 |
| `get_inventory_list()` | `getInventoryList` | 获取库存列表 |
| `get_billing_list()` | `getBillingList` | 获取账单列表 |
| `get_warehouse_rent()` | `getWarehouseRent` | 获取仓租明细 |

更多接口请参考：`methods.py` 中的 `EccangMethods` 类

## 📝 示例代码

运行完整示例：

```bash
python api/eccang/example.py
```

## 📖 完整文档

详细文档请参考：`api/eccang/README.md`

## 🔧 模块结构

```
api/eccang/
├── __init__.py         # 模块导出
├── config.py           # 配置文件（需填写凭证）
├── exceptions.py       # 异常类定义
├── client.py           # HTTP 客户端（签名、请求）
├── methods.py          # 常用接口封装
├── cli.py              # 命令行工具
├── example.py          # 使用示例
├── README.md           # 详细文档
└── QUICKSTART.md       # 本文件（快速入门）
```

## ⚠️ 注意事项

1. **IP 白名单**：务必在易仓开放平台设置 IP 白名单
2. **SERVICE_ID**：需要联系易仓客户授权获取
3. **签名规则**：本模块已自动处理签名，无需手动计算
4. **时间格式**：
   - 日期时间：`YYYY-MM-DD HH:MM:SS`
   - 日期：`YYYY-MM-DD`
5. **字符编码**：统一使用 UTF-8

## 🐛 常见问题

### Q1: 签名错误？

检查 `APP_KEY`、`APP_SECRET`、`SERVICE_ID` 是否正确。

### Q2: IP 不在白名单？

在易仓开放平台的应用管理中添加服务器 IP。

### Q3: 接口返回数据为空？

检查查询条件、时间范围、账号权限。

## 📞 技术支持

- **易仓开放平台**：https://open.eccang.com
- **API 文档**：https://open.eccang.com/#/documentCenter
- **项目 README**：`d:\py-project\README.md`

---

**创建时间**：2026-06-30  
**版本**：1.0.0  
**作者**：报表项目组
