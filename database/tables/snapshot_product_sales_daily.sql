-- ========================================
-- 表：snapshot_product_sales_daily
-- 数据库：rpa-report
-- 说明：产品每日销量统计
-- ========================================
SET NAMES utf8mb4 COLLATE utf8mb4_unicode_ci;
USE rpa-report;

DROP TABLE IF EXISTS snapshot_product_sales_daily;
CREATE TABLE `snapshot_product_sales_daily` (
  `id` bigint unsigned NOT NULL AUTO_INCREMENT COMMENT '自增主键',
  `snap_date` date NOT NULL COMMENT '快照日期',
  `warehouse_code` varchar(128) COLLATE utf8mb4_unicode_ci NOT NULL COMMENT '仓库代码',
  `warehouse_name` varchar(128) COLLATE utf8mb4_unicode_ci NOT NULL DEFAULT '' COMMENT '仓库名称',
  `product_sku` varchar(128) COLLATE utf8mb4_unicode_ci NOT NULL COMMENT '产品SKU',
  `sale_status` tinyint NOT NULL DEFAULT '0' COMMENT '产品销售状态',
  
  `purchase_quantity` int NOT NULL DEFAULT '0' COMMENT '采购数量',
  `pi_purchase_onway_qty` int NOT NULL DEFAULT '0' COMMENT '采购在途',
  `pi_return_onway_qty` int NOT NULL DEFAULT '0' COMMENT '退货在途',
  `pi_pending_qty` int NOT NULL DEFAULT '0' COMMENT '待处理库存',
  `pi_in_used_qty` int NOT NULL DEFAULT '0' COMMENT '可用库存',
  `pi_warning_qty` int NOT NULL DEFAULT '0' COMMENT '预警库存',
  `pi_sellable_qty` int NOT NULL DEFAULT '0' COMMENT '可销售库存',
  `pi_shared_qty` int NOT NULL DEFAULT '0' COMMENT '分销数量',
  `pi_reserved_qty` int NOT NULL DEFAULT '0' COMMENT '预留/待出库',
  `pi_no_stock_qty` int NOT NULL DEFAULT '0' COMMENT '缺货数量',
  `pi_unsellable_qty` int NOT NULL DEFAULT '0' COMMENT '不可销售库存',
  `pi_outbound_qty` int NOT NULL DEFAULT '0' COMMENT '待出库不良',
  `pi_planned_qty` int NOT NULL DEFAULT '0' COMMENT '计划库存',
  `actual_usable_inventory_qty` int NOT NULL DEFAULT '0' COMMENT '实际可用库存',

  `sale_qty_3d` int NOT NULL DEFAULT '0' COMMENT '3天销量',
  `sale_qty_7d` int NOT NULL DEFAULT '0' COMMENT '7天销量',
  `sale_qty_14d` int NOT NULL DEFAULT '0' COMMENT '14天销量',
  `sale_qty_30d` int NOT NULL DEFAULT '0' COMMENT '30天销量',
  `sale_qty_60d` int NOT NULL DEFAULT '0' COMMENT '60天销量',
  `total_qty_60d` int NOT NULL DEFAULT '0' COMMENT '60天总销量',

  `market_code` varchar(64) COLLATE utf8mb4_unicode_ci NOT NULL DEFAULT '' COMMENT '市场编码',
  `market_region` varchar(64) COLLATE utf8mb4_unicode_ci NOT NULL DEFAULT '' COMMENT '市场区域',
  `ops_owner` varchar(128) COLLATE utf8mb4_unicode_ci NOT NULL DEFAULT '' COMMENT '负责人',
  `source` varchar(32) COLLATE utf8mb4_unicode_ci NOT NULL DEFAULT 'eccang' COMMENT '数据来源',
  `created_at` datetime NOT NULL DEFAULT CURRENT_TIMESTAMP COMMENT '写入时间',
  `updated_at` datetime NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP COMMENT '更新时间',
  PRIMARY KEY (`id`),
  UNIQUE KEY `uniq_snap_date_warehouse_code_product_sku` (`snap_date`, `warehouse_code`, `product_sku`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci COMMENT='产品每日销量统计';