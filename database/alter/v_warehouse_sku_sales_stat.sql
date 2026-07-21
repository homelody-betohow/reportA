DROP VIEW IF EXISTS `v_warehouse_sku_sales_stat`;
CREATE VIEW `v_warehouse_sku_sales_stat` AS
SELECT
    warehouse_code,
    warehouse_name,
    warehouse_sku,
    product_name,
    -- 近3天销量
    SUM(CASE WHEN ship_time >= DATE_SUB(CURDATE(), INTERVAL 3 DAY) THEN warehouse_sku_qty ELSE 0 END) AS sale_qty_3d,
    -- 近7天销量
    SUM(CASE WHEN ship_time >= DATE_SUB(CURDATE(), INTERVAL 7 DAY) THEN warehouse_sku_qty ELSE 0 END) AS sale_qty_7d,
    -- 近14天销量
    SUM(CASE WHEN ship_time >= DATE_SUB(CURDATE(), INTERVAL 14 DAY) THEN warehouse_sku_qty ELSE 0 END) AS sale_qty_14d,
    -- 近30天销量
    SUM(CASE WHEN ship_time >= DATE_SUB(CURDATE(), INTERVAL 30 DAY) THEN warehouse_sku_qty ELSE 0 END) AS sale_qty_30d,
    -- 近60天销量
    SUM(CASE WHEN ship_time >= DATE_SUB(CURDATE(), INTERVAL 60 DAY) THEN warehouse_sku_qty ELSE 0 END) AS sale_qty_60d,
    -- 近90天总销量
    SUM(CASE WHEN ship_time >= DATE_SUB(CURDATE(), INTERVAL 90 DAY) THEN warehouse_sku_qty ELSE 0 END) AS sale_qty_90d
FROM sales_order_shipped
WHERE 
    warehouse_code <> '' 
    AND order_type = '销售订单'         
    AND warehouse_sku_qty > 0
    -- 基础数据只取近90天发货单
    AND ship_time >= DATE_SUB(CURDATE(), INTERVAL 90 DAY)
GROUP BY 
    warehouse_code, 
    warehouse_name, 
    warehouse_sku, 
    product_name;