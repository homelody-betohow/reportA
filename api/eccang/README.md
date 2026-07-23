# 易仓 ERP API 对接使用说明

## 一、配置说明

### 1.1 获取 API 凭证

在使用易仓 ERP API 之前，需要先获取以下凭证：

1. **访问易仓开放平台**：https://open.eccang.com
2. **登录并进入应用管理**：
   - 点击左侧导航栏的"应用管理"
   - 选择顶部"新增"按钮
3. **创建服务商应用**：
   - 在创建页面选择"服务商应用"类型
   - 填写应用名称、简介等基础信息
   - **务必设置 IP 白名单**保障访问安全
4. **获取授权凭证**（审核通过后）：
   - **APP_KEY**：应用唯一标识
   - **APP_SECRET**：安全验证密钥
   - **SERVICE_ID**：需联系易仓客户授权获取

### 1.2 配置文件填写

在 `api/eccang/config.py` 中填写获取到的凭证：

```python
# ─── 易仓 ERP 开放平台凭证 ────
APP_KEY = "your_app_key_here"
APP_SECRET = "your_app_secret_here"
SERVICE_ID = "your_service_id_here"
```

## 二、使用方法

### 2.1 基本使用示例

```python
from api.eccang import EccangService

# 创建客户端（自动加载 config.py 中的配置）
client = EccangService()

# 获取仓库列表
response = client.get_warehouse_list(page=1, page_size=50)
print(response)

# 获取订单列表（指定时间范围）
response = client.get_order_list(
    page=1,
    page_size=50,
    start_time="2024-01-01 00:00:00",
    end_time="2024-01-31 23:59:59"
)
print(response)

# 获取产品列表
response = client.get_product_list(page=1, page_size=50)
print(response)

# 获取库存列表（指定仓库）
response = client.get_inventory_list(
    warehouse_code="WH001",
    page=1,
    page_size=50
)
print(response)

# 获取账单列表
response = client.get_billing_list(
    start_date="2024-01-01",
    end_date="2024-01-31",
    page=1,
    page_size=50
)
print(response)
```

### 2.2 调用自定义接口

如果需要调用文档中的其他接口，可以使用 `call()` 方法：

```python
from api.eccang import EccangClient

client = EccangClient()

# 调用任意接口（以获取发货地址簿为例）
response = client.call(
    method="getShipAddressBooks",
    body={
        "page": 1,
        "page_size": 50
    }
)
print(response)
```

### 2.3 自定义签名顺序

某些接口可能有特殊的签名顺序要求，可以通过 `sign_order` 参数指定：

```python
from api.eccang import EccangClient

client = EccangClient()

# 自定义签名字段顺序
custom_order = [
    "app_key",
    "service_id",
    "interface_method",
    "timestamp",
    "nonce_str",
    "biz_content",
    "charset",
    "sign_type",
    "version"
]

response = client.call(
    method="someSpecialMethod",
    body={"key": "value"},
    sign_order=custom_order
)
```

## 三、命令行工具

### 3.1 测试连接

```bash
# 测试 API 连接（获取仓库列表）
python -m api.eccang.cli test
```

### 3.2 获取各类数据

```bash
# 获取仓库列表
python -m api.eccang.cli get-warehouses --page 1 --page-size 50

# 获取订单列表
python -m api.eccang.cli get-orders \
    --page 1 \
    --page-size 50 \
    --start-time "2024-01-01 00:00:00" \
    --end-time "2024-01-31 23:59:59"

# 获取产品列表
python -m api.eccang.cli get-products --page 1 --page-size 50

# 获取库存列表
python -m api.eccang.cli get-inventory \
    --warehouse-code WH001 \
    --page 1 \
    --page-size 50

# 获取账单列表
python -m api.eccang.cli get-billing \
    --start-date "2024-01-01" \
    --end-date "2024-01-31" \
    --page 1 \
    --page-size 50
```

### 3.3 调用任意接口

```bash
# 调用任意接口方法
python -m api.eccang.cli call getShipAddressBooks \
    --body '{"page": 1, "page_size": 50}'
```

## 四、API 接口说明

### 4.1 已封装的常用接口

| 方法名 | 接口方法 | 说明 |
|--------|----------|------|
| `get_ship_address_books()` | `getShipAddressBooks` | 获取发货地址簿 |
| `get_warehouse_list()` | `getWarehouseList` | 获取仓库列表 |
| `get_order_list()` | `getOrderList` | 获取订单列表 |
| `get_order_detail()` | `getOrderDetail` | 获取订单详情 |
| `get_product_list()` | `getProductList` | 获取产品列表 |
| `get_inventory_list()` | `getInventoryList` | 获取库存列表 |
| `get_billing_list()` | `getBillingList` | 获取账单列表 |
| `get_warehouse_rent()` | `getWarehouseRent` | 获取仓租明细 |

### 4.2 添加新接口

如需使用其他接口，可以在 `methods.py` 中添加：

1. 在 `EccangMethods` 类中添加接口方法常量
2. 在 `EccangService` 类中添加封装方法

示例：

```python
# 在 EccangMethods 类中添加
GET_SHIPPING_INFO = "getShippingInfo"  # 获取物流信息

# 在 EccangService 类中添加
def get_shipping_info(self, tracking_no: str) -> dict[str, Any]:
    """获取物流信息。"""
    return self.call(
        EccangMethods.GET_SHIPPING_INFO,
        {"tracking_no": tracking_no}
    )
```

## 五、错误处理

### 5.1 配置错误

```python
from api.eccang import EccangService
from api.eccang.exceptions import EccangConfigError

try:
    client = EccangService()
except EccangConfigError as e:
    print(f"配置错误：{e}")
    # 检查 config.py 中的 APP_KEY、APP_SECRET、SERVICE_ID 是否已填写
```

### 5.2 接口调用错误

```python
from api.eccang import EccangService
from api.eccang.exceptions import EccangApiError

client = EccangService()

try:
    response = client.get_order_list(page=1, page_size=50)
except EccangApiError as e:
    print(f"接口调用失败：{e}")
    print(f"错误代码：{e.code}")
    print(f"接口方法：{e.method}")
```

## 六、签名规则说明

依据官方文档：[签名生成规则](https://open.eccang.com/#/documentCenter?docId=313&catId=0-173-173,0-171)

1. **参与签名的参数**（不含 `sign`）：
   `app_key`、`biz_content`、`charset`、`interface_method`、`nonce_str`、`service_id`、`sign_type`、`timestamp`、`version`
2. **排序**：按参数名 ASCII 码升序
3. **拼接**：`key1=value1&key2=value2&...`
4. **追加密钥**：在拼接串末尾直接追加 `app_secret`（不加 `&`）
5. **计算签名**：对上述字符串做 MD5，取 **32 位小写** 十六进制作为 `sign`

示例：

```text
app_key=xxx&biz_content={"page":1}&charset=UTF-8&interface_method=getOrderList&nonce_str=abc123&service_id=ERPxxx&sign_type=MD5&timestamp=1710000000000&version=1.0.0{app_secret}
sign = md5(上述字符串).lower()
```

**注意**：
- `biz_content` 必须是 JSON 字符串（紧凑格式）
- `timestamp` 为毫秒级时间戳，需在 1 分钟内有效
- 本模块已按上述规则自动签名，一般无需手动计算

## 七、响应格式说明

### 7.1 成功响应

```json
{
  "code": "0",
  "ask": "Success",
  "message": "操作成功",
  "data": {
    "list": [...],
    "total": 100,
    "page": 1,
    "page_size": 50
  }
}
```

### 7.2 失败响应

```json
{
  "code": "1001",
  "ask": "Error",
  "message": "参数错误",
  "data": null
}
```

## 八、参考资料

- **易仓开放平台**：https://open.eccang.com
- **API 文档中心**：https://open.eccang.com/#/documentCenter
- **应用管理入口**：https://open.eccang.com/#/appManage

## 九、注意事项

1. **IP 白名单**：创建应用时务必设置 IP 白名单，否则无法调用接口
2. **SERVICE_ID**：需联系易仓客户授权获取，外包团队需特别注意
3. **签名顺序**：大部分接口使用默认签名顺序，少数特殊接口可能需要自定义
4. **时间格式**：
   - 日期时间：`YYYY-MM-DD HH:MM:SS`
   - 日期：`YYYY-MM-DD`
5. **分页参数**：大部分接口支持 `page` 和 `page_size` 参数
6. **字符编码**：统一使用 UTF-8
7. **超时设置**：默认 60 秒，可在 `config.py` 中修改 `TIMEOUT` 常量

## 十、常见问题

### Q1: 签名错误怎么办？

A: 检查以下几点：
- APP_KEY、APP_SECRET、SERVICE_ID 是否正确
- 签名顺序是否正确（大部分接口使用默认顺序）
- 请求参数是否正确编码（UTF-8）

### Q2: 接口返回 "IP 不在白名单" 错误？

A: 在易仓开放平台的应用管理中，将服务器 IP 添加到白名单。

### Q3: 如何查看具体的接口文档？

A: 访问 https://open.eccang.com/#/documentCenter，选择对应的接口分类查看详细文档。

### Q4: 如何添加新的接口方法？

A: 参考"四、API 接口说明 -> 4.2 添加新接口"部分。

### Q5: 接口返回数据为空？

A: 检查以下几点：
- 查询条件是否正确（时间范围、仓库编码等）
- 账号是否有对应数据的访问权限
- 分页参数是否合理
