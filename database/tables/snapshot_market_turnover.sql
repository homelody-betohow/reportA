-- ========================================
-- 表：snapshot_market_turnover
-- 数据库：rpa-report
-- 说明：市场库存周转快照（库存指标 + 滚动销量）
--       粒度：snapshot_time + market_code + product_sku
--       用途：替代 Excel「市场库存周转明细」做仓租分摊 / 可售天数 / 周转分析
--       维度字段（provider/market/ops）写入时从 warehouse 冗余，便于按日查询不依赖维表变更
-- ========================================
SET NAMES utf8mb4 COLLATE utf8mb4_unicode_ci;
USE `rpa-report`;

DROP TABLE IF EXISTS `snapshot_market_turnover`;
CREATE TABLE `snapshot_market_turnover` (
  `id` bigint unsigned NOT NULL AUTO_INCREMENT COMMENT '自增主键',
  `snapshot_date` date NOT NULL COMMENT '快照日期',
  -- 市场 / SKU 业务键
  `market_code` varchar(128) COLLATE utf8mb4_unicode_ci NOT NULL COMMENT '市场代码',
  `market_region` varchar(128) COLLATE utf8mb4_unicode_ci NOT NULL COMMENT '市场区域',
  `product_sku` varchar(128) COLLATE utf8mb4_unicode_ci NOT NULL COMMENT '产品SKU',
  `product_uid` varchar(128) COLLATE utf8mb4_unicode_ci NOT NULL COMMENT '商品ID',
  `sku_lifecycle` varchar(128) COLLATE utf8mb4_unicode_ci NOT NULL COMMENT '产品状态',
  `ops_owner` varchar(128) COLLATE utf8mb4_unicode_ci NOT NULL COMMENT '运营负责人',
  `ops_leader` varchar(128) COLLATE utf8mb4_unicode_ci NOT NULL COMMENT '运营经理',
  `dir_planned_qty` int NOT NULL DEFAULT '0' COMMENT '计划库存(禁调拨)',
  `trf_planned_qty` int NOT NULL DEFAULT '0' COMMENT '计划库存(可调拨)',
  `dir_onway_qty` int NOT NULL DEFAULT '0' COMMENT '在途库存(禁调拨)',
  `trf_onway_qty` int NOT NULL DEFAULT '0' COMMENT '在途库存(可调拨)',
  `dir_sellable_qty` int NOT NULL DEFAULT '0' COMMENT '可售库存(禁调拨)',
  `trf_sellable_qty` int NOT NULL DEFAULT '0' COMMENT '可售库存(可调拨)',
  `total_planned_qty` int NOT NULL DEFAULT '0' COMMENT '总计划库存',
  `total_onway_qty` int NOT NULL DEFAULT '0' COMMENT '总在途库存',
  `total_sellable_qty` int NOT NULL DEFAULT '0' COMMENT '总可售库存',
  `ref_month_sales_qty` int NOT NULL DEFAULT '0' COMMENT '参考月销量',
  `supplier_name` varchar(128) COLLATE utf8mb4_unicode_ci NOT NULL COMMENT '供应商名称',
  `cost_price_cny` decimal(18,4) NOT NULL DEFAULT '0.0000' COMMENT '成本价',
  `line_hash` char(64) COLLATE utf8mb4_unicode_ci NOT NULL COMMENT '行内容稳定哈希（SHA-256 hex），用于变更检测与幂等 UPSERT',
  `created_at` datetime NOT NULL DEFAULT CURRENT_TIMESTAMP COMMENT '写入时间',
  `updated_at` datetime NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP COMMENT '更新时间',
  PRIMARY KEY (`id`),
  UNIQUE KEY `uniq_snapshot_date_market_user_product_sku` (`snapshot_date`, `market_code`, `market_region`, `ops_owner`, `product_sku`),
  UNIQUE KEY `uk_smt_line_hash` (`line_hash`),
  KEY `idx_snapshot_date_market_code_product_sku` (`snapshot_date`, `market_code`, `product_sku`),
  KEY `idx_snapshot_date_market_region_ops_owner_product_sku` (`snapshot_date`, `market_region`, `ops_owner`, `product_sku`),
  KEY `idx_product_sku_snapshot_date` (`product_sku`, `snapshot_date`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci COMMENT='市场库存周转快照表';