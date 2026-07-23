DROP VIEW IF EXISTS `v_warehouse_sku_sales_stat`;
-- SELECT max(ship_time) FROM `sales_order_shipped`;
-- 因为sales_order_shipped数据导入滞后3天，所以统计数据要延迟3天
CREATE VIEW `v_warehouse_sku_sales_stat` AS
SELECT
		warehouse_code,
    warehouse_name,
    warehouse_sku,
    product_name,
    -- 近3天销量
    SUM(CASE WHEN ship_time >= DATE_SUB(CURDATE(), INTERVAL 6 DAY) THEN warehouse_sku_qty ELSE 0 END) AS sale_qty_3d,
    -- 近7天销量
    SUM(CASE WHEN ship_time >= DATE_SUB(CURDATE(), INTERVAL 10 DAY) THEN warehouse_sku_qty ELSE 0 END) AS sale_qty_7d,
    -- 近14天销量
    SUM(CASE WHEN ship_time >= DATE_SUB(CURDATE(), INTERVAL 17 DAY) THEN warehouse_sku_qty ELSE 0 END) AS sale_qty_14d,
    -- 近30天销量
    SUM(CASE WHEN ship_time >= DATE_SUB(CURDATE(), INTERVAL 33 DAY) THEN warehouse_sku_qty ELSE 0 END) AS sale_qty_30d,
    -- 近60天销量
    SUM(CASE WHEN ship_time >= DATE_SUB(CURDATE(), INTERVAL 63 DAY) THEN warehouse_sku_qty ELSE 0 END) AS sale_qty_60d,
    -- 近90天总销量
    SUM(CASE WHEN ship_time >= DATE_SUB(CURDATE(), INTERVAL 93 DAY) THEN warehouse_sku_qty ELSE 0 END) AS sale_qty_90d
FROM sales_order_shipped
WHERE 
    order_type = '销售订单' and warehouse_name <> '--'         
    AND warehouse_sku_qty > 0
    AND ship_time >= DATE_SUB(CURDATE(), INTERVAL 93 DAY)
		GROUP BY 
    warehouse_code, 
    warehouse_name, 
    warehouse_sku, 
    product_name;