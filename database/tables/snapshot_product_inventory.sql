-- ========================================
-- 表：snapshot_product_inventory
-- 数据库：rpa-report
-- 说明：库存快照；同一 snapshot_time（日期）+ 仓库 + SKU 唯一
-- ========================================
SET NAMES utf8mb4 COLLATE utf8mb4_unicode_ci;
USE rpa-report;

DROP TABLE IF EXISTS snapshot_product_inventory;
CREATE TABLE `snapshot_product_inventory` (
  `id` bigint unsigned NOT NULL AUTO_INCREMENT COMMENT '自增主键',
  `snapshot_time` date NOT NULL COMMENT '快照日期（同次拉取写入相同值）',
  `warehouse_id` int NOT NULL COMMENT '易仓仓库ID',
  `warehouse_code` varchar(128) COLLATE utf8mb4_unicode_ci NOT NULL COMMENT '仓库代码',
  `warehouse_name` varchar(128) COLLATE utf8mb4_unicode_ci NOT NULL DEFAULT '' COMMENT '仓库名称',
  `warehouse_sku` varchar(128) COLLATE utf8mb4_unicode_ci NOT NULL COMMENT '仓库SKU',
  `product_sku` varchar(128) COLLATE utf8mb4_unicode_ci NOT NULL COMMENT '产品SKU',
  `product_title` varchar(255) COLLATE utf8mb4_unicode_ci NOT NULL DEFAULT '' COMMENT '产品标题',
  `product_title_en` varchar(255) COLLATE utf8mb4_unicode_ci NOT NULL DEFAULT '' COMMENT '产品标题（英文）',
  `product_weight` decimal(10,3) NOT NULL DEFAULT '0.000' COMMENT '产品重量（kg）',
  `sale_status` tinyint NOT NULL DEFAULT '0' COMMENT '产品销售状态',
  `currency_code` varchar(16) COLLATE utf8mb4_unicode_ci NOT NULL DEFAULT '' COMMENT '币种',
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
  `pi_can_sale_days` int NOT NULL DEFAULT '0' COMMENT '可售天数',
  `pi_no_stock_days` int NOT NULL DEFAULT '0' COMMENT '缺货天数',
  `pending_qc_qty` int NOT NULL DEFAULT '0' COMMENT '待质检数量',
  `actual_usable_inventory_qty` int NOT NULL DEFAULT '0' COMMENT '实际可用库存',
  `pi_update_time` datetime DEFAULT NULL COMMENT '易仓库存更新时间',
  `provider_code` varchar(16) COLLATE utf8mb4_unicode_ci NOT NULL DEFAULT '' COMMENT '服务商：HY、4PX、FBA 等（可人工维护）',
  `market_code` varchar(64) COLLATE utf8mb4_unicode_ci NOT NULL DEFAULT '' COMMENT '市场编码（可人工维护）',
  `market_region` varchar(64) COLLATE utf8mb4_unicode_ci NOT NULL DEFAULT '' COMMENT '市场区域（可人工维护）',
  `is_transfer` tinyint(1) DEFAULT NULL COMMENT '是否为中转仓（回填自 warehouse.is_transfer）',
  `ops_owner` varchar(128) COLLATE utf8mb4_unicode_ci NOT NULL DEFAULT '' COMMENT '负责人（可人工维护）',
  `source` varchar(32) COLLATE utf8mb4_unicode_ci NOT NULL DEFAULT 'eccang' COMMENT '数据来源',
  `created_at` datetime NOT NULL DEFAULT CURRENT_TIMESTAMP COMMENT '写入时间',
  PRIMARY KEY (`id`),
  UNIQUE KEY `uk_spi_snapshot_wh_sku` (`snapshot_time`,`warehouse_id`,`product_sku`),
  KEY `idx_spi_product_sku` (`product_sku`),
  KEY `idx_spi_warehouse_code` (`warehouse_code`),
  KEY `idx_spi_snapshot_time` (`snapshot_time`),
  KEY `idx_spi_pi_update_time` (`pi_update_time`),
  KEY `idx_spi_provider_code` (`provider_code`),
  KEY `idx_spi_sale_status` (`sale_status`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci COMMENT='产品库存快照表';
