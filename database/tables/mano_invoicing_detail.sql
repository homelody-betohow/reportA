-- ========================================
-- 表：mano_invoicing_detail
-- 数据库：rpa-report
-- 用途：MANO MMF 月度账单明细（Invoicing Details）
-- 来源：\\Betohow\数据报表\RPA\MANO\MANO-Invoicing\{YYYY-MM}\*.xlsx
-- Sheet：Master Sheet Invoicing Details（第2行为字段名，第3行起为数据）
-- ========================================
SET NAMES utf8mb4 COLLATE utf8mb4_unicode_ci;
USE rpa-report;

DROP TABLE IF EXISTS mano_invoicing_detail;
CREATE TABLE `mano_invoicing_detail` (
  `id` bigint unsigned NOT NULL AUTO_INCREMENT COMMENT '自增主键',
  `line_hash` char(64) COLLATE utf8mb4_unicode_ci NOT NULL COMMENT '行内容稳定哈希（SHA-256 hex），用于去重与增量 UPSERT',

  `billing_period` char(7) COLLATE utf8mb4_unicode_ci NOT NULL COMMENT '账单月份，如 2026-05',
  `source_file` varchar(255) COLLATE utf8mb4_unicode_ci NOT NULL COMMENT '来源文件名',
  `source_brand` varchar(32) COLLATE utf8mb4_unicode_ci DEFAULT NULL COMMENT '文件名解析品牌：COMFR / OHPA_FR 等',
  `source_sheet` varchar(128) COLLATE utf8mb4_unicode_ci NOT NULL COMMENT '市场区域（按账单文件名 _MARKET_REGION_MAP 映射，如 MANO-DE-COMMF）',
  `import_batch` varchar(50) COLLATE utf8mb4_unicode_ci DEFAULT NULL COMMENT '导入批次标识',
  `source_type` varchar(24) COLLATE utf8mb4_unicode_ci DEFAULT 'Excel' COMMENT '来源类型：Excel/API',

  `seller_account_id` int NOT NULL COMMENT 'SELLER_ACCOUNT_ID',
  `seller_id` bigint NOT NULL COMMENT 'SELLER_ID',
  `seller_name` varchar(64) COLLATE utf8mb4_unicode_ci NOT NULL COMMENT 'SELLER_NAME，如 FR - MF',
  `platform` char(2) COLLATE utf8mb4_unicode_ci NOT NULL COMMENT 'PLATFORM 国家代码：DE/FR/ES/IT',

  `invoicing_item` varchar(32) COLLATE utf8mb4_unicode_ci NOT NULL COMMENT '计费类型：DISPATCH / STORAGE',
  `subcategory` varchar(32) COLLATE utf8mb4_unicode_ci NOT NULL COMMENT '子分类：派送尺寸或库龄段',

  `shipment_number` varchar(64) COLLATE utf8mb4_unicode_ci DEFAULT NULL COMMENT 'SHIPMENT_NUMBER',
  `in_stock_date` date DEFAULT NULL COMMENT 'IN_STOCK_DATE',
  `inbound_wh` varchar(16) COLLATE utf8mb4_unicode_ci DEFAULT NULL COMMENT 'INBOUND_WH，如 CHA/CRE',

  `ean` varchar(20) COLLATE utf8mb4_unicode_ci NOT NULL COMMENT 'EAN 条码',
  `seller_sku` varchar(64) COLLATE utf8mb4_unicode_ci NOT NULL COMMENT 'SELLER_SKU',
  `warehouse_sku` varchar(64) COLLATE utf8mb4_unicode_ci DEFAULT NULL COMMENT 'WAREHOUSE_SKU',
  `mm_id` varchar(20) COLLATE utf8mb4_unicode_ci NOT NULL COMMENT 'MM_ID 平台商品ID',
  `product_price_vat_exc` decimal(18,6) DEFAULT NULL COMMENT 'PRODUCT_PRICE_VAT_EXC（主要 DISPATCH 有值）',
  `product_name` varchar(512) COLLATE utf8mb4_unicode_ci NOT NULL COMMENT 'PRODUCT_NAME',
  `is_multipart` varchar(8) COLLATE utf8mb4_unicode_ci DEFAULT NULL COMMENT 'IS_MULTIPART',

  `length_cm` decimal(10,3) NOT NULL COMMENT 'LENGTH (CM)',
  `width_cm` decimal(10,3) NOT NULL COMMENT 'WIDTH (CM)',
  `height_cm` decimal(10,3) NOT NULL COMMENT 'HEIGHT (CM)',
  `weight_kg` decimal(10,3) NOT NULL COMMENT 'WEIGHT (KG)',
  `volume_m3` decimal(18,9) NOT NULL COMMENT 'VOLUME (M3)',
  `order_code` varchar(64) COLLATE utf8mb4_unicode_ci DEFAULT NULL COMMENT 'ORDER_CODE',
  `order_reference` varchar(32) COLLATE utf8mb4_unicode_ci DEFAULT NULL COMMENT 'ORDER_REFERENCE，如 M2605506023397',
  `quantity_ordered` int DEFAULT NULL COMMENT 'QUANTITY_ORDERED',
  `order_date` date DEFAULT NULL COMMENT 'ORDER_DATE',
  `dispatch_date` date DEFAULT NULL COMMENT 'DISPATCH_DATE',
  `dispatch_warehouse` varchar(16) COLLATE utf8mb4_unicode_ci DEFAULT NULL COMMENT 'DISPATCH_WAREHOUSE',
  `destination_platform` char(2) COLLATE utf8mb4_unicode_ci DEFAULT NULL COMMENT 'DESTINATION_PLATFORM',

  `refund_date` date DEFAULT NULL COMMENT 'REFUND_DATE',
  `return_status` varchar(64) COLLATE utf8mb4_unicode_ci DEFAULT NULL COMMENT 'RETURN_STATUS',
  `return_reason` varchar(255) COLLATE utf8mb4_unicode_ci DEFAULT NULL COMMENT 'RETURN_REASON',

  `customer_refund_amount` decimal(18,6) DEFAULT NULL COMMENT 'CUSTOMER_REFUND_AMOUNT',
  `compensation_reason` varchar(255) COLLATE utf8mb4_unicode_ci DEFAULT NULL COMMENT 'COMPENSATION_REASON',
  `compensation_type` varchar(64) COLLATE utf8mb4_unicode_ci DEFAULT NULL COMMENT 'COMPENSATION_TYPE',
  `product_compensation` decimal(18,6) DEFAULT NULL COMMENT 'PRODUCT_COMPENSATION',
  `applicable_dispatch_fee` decimal(18,6) DEFAULT NULL COMMENT 'APPLICABLE_DISPATCH_FEE',

  `unit_price_vat_exc` decimal(18,9) NOT NULL COMMENT 'UNIT_PRICE_VAT_EXC 单价（不含税 EUR）',
  `quantity` decimal(18,6) NOT NULL COMMENT 'QUANTITY',
  `quantity_type` varchar(64) COLLATE utf8mb4_unicode_ci NOT NULL COMMENT 'QUANTITY_TYPE',
  `gross_amount_vat_exc` decimal(18,9) NOT NULL COMMENT 'GROSS_AMOUNT_VAT_EXC',
  `net_amount_vat_exc` decimal(18,9) NOT NULL COMMENT 'NET_AMOUNT_VAT_EXC',

  `raw_row_json` json DEFAULT NULL COMMENT '原始行 JSON（用于追溯与排错）',
  `created_at` timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP COMMENT '写入时间',
  `updated_at` timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP COMMENT '更新时间',
  PRIMARY KEY (`id`),
  UNIQUE KEY `uk_mano_invoicing_line_hash` (`line_hash`),
  KEY `idx_mano_inv_billing_period` (`billing_period`),
  KEY `idx_mano_inv_seller_period` (`seller_id`,`billing_period`),
  KEY `idx_mano_inv_item_period` (`invoicing_item`,`billing_period`),
  KEY `idx_mano_inv_order_ref` (`order_reference`),
  KEY `idx_mano_inv_seller_sku` (`seller_sku`),
  KEY `idx_mano_inv_mm_id` (`mm_id`),
  KEY `idx_mano_inv_ean` (`ean`),
  KEY `idx_mano_inv_dispatch_date` (`dispatch_date`),
  KEY `idx_mano_inv_source_file` (`source_file`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci COMMENT='MANO MMF 月度账单明细（Invoicing Details）';
