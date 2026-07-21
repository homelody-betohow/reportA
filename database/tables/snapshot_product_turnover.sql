-- ========================================
-- 表：snapshot_product_turnover
-- 数据库：rpa-report
-- 说明：产品库存周转快照（库存指标 + 滚动销量）
--       粒度：snapshot_time + warehouse_code + product_sku
--       用途：替代 Excel「库存周转明细」做仓租分摊 / 可售天数 / 周转分析
--       维度字段（provider/market/ops）写入时从 warehouse 冗余，便于按日查询不依赖维表变更
-- ========================================
SET NAMES utf8mb4 COLLATE utf8mb4_unicode_ci;
USE `rpa-report`;

DROP TABLE IF EXISTS `snapshot_product_turnover`;
CREATE TABLE `snapshot_product_turnover` (
  `id` bigint unsigned NOT NULL AUTO_INCREMENT COMMENT '自增主键',
  `snapshot_date` date NOT NULL COMMENT '快照日期',

  -- 仓库 / SKU 业务键
  `warehouse_code` varchar(128) COLLATE utf8mb4_unicode_ci NOT NULL COMMENT '仓库代码',
  `warehouse_name` varchar(128) COLLATE utf8mb4_unicode_ci NOT NULL DEFAULT '' COMMENT '仓库名称',
  `product_sku` varchar(128) COLLATE utf8mb4_unicode_ci NOT NULL COMMENT '产品SKU',
  `product_uid` varchar(128) COLLATE utf8mb4_unicode_ci NOT NULL COMMENT '商品ID',
  `product_title` varchar(255) COLLATE utf8mb4_unicode_ci NOT NULL DEFAULT '' COMMENT '产品标题',

  -- 库存指标（与 snapshot_product_inventory 对齐）
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
  `pi_can_sale_days` int NOT NULL DEFAULT '0' COMMENT '可售天数（来源系统）',
  `pi_no_stock_days` int NOT NULL DEFAULT '0' COMMENT '缺货天数',
  `pending_qc_qty` int NOT NULL DEFAULT '0' COMMENT '待质检数量',
  `actual_usable_inventory_qty` int NOT NULL DEFAULT '0' COMMENT '实际可用库存',

  -- 滚动销量（按出库统计窗口，写入时固化）
  `sale_qty_3d` int NOT NULL DEFAULT '0' COMMENT '近3天销量',
  `sale_qty_7d` int NOT NULL DEFAULT '0' COMMENT '近7天销量',
  `sale_qty_14d` int NOT NULL DEFAULT '0' COMMENT '近14天销量',
  `sale_qty_30d` int NOT NULL DEFAULT '0' COMMENT '近30天销量',
  `sale_qty_60d` int NOT NULL DEFAULT '0' COMMENT '近60天销量',
  `sale_qty_90d` int NOT NULL DEFAULT '0' COMMENT '近90天销量',

  -- 冗余维度（来自 warehouse，快照日固化）
  `provider_code` varchar(16) COLLATE utf8mb4_unicode_ci NOT NULL DEFAULT '' COMMENT '服务商：HY、4PX、FBA 等',
  `market_code` varchar(64) COLLATE utf8mb4_unicode_ci NOT NULL DEFAULT '' COMMENT '市场编码',
  `market_region` varchar(64) COLLATE utf8mb4_unicode_ci NOT NULL DEFAULT '' COMMENT '市场区域',
  `is_transfer` tinyint(1) DEFAULT NULL COMMENT '是否为中转仓（回填自 warehouse.is_transfer）',
  `ops_owner` varchar(128) COLLATE utf8mb4_unicode_ci NOT NULL DEFAULT '' COMMENT '负责人',

  `created_at` datetime NOT NULL DEFAULT CURRENT_TIMESTAMP COMMENT '写入时间',
  `updated_at` datetime NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP COMMENT '更新时间',

  PRIMARY KEY (`id`),
  -- 业务唯一：同日同仓同 SKU 只保留一行（幂等重跑）
  UNIQUE KEY `uk_spt_snap_wh_sku` (`snapshot_date`, `warehouse_code`, `product_sku`),
  -- SKU 历史走势：按 SKU 查多日周转
  KEY `idx_spt_sku_snap` (`product_sku`, `snapshot_date`),
  -- 按市场编码 + 日期筛选（仓租/运营常用）
  KEY `idx_spt_market_snap` (`market_code`, `snapshot_date`),
  -- 仓库代码回查
  KEY `idx_spt_warehouse` (`warehouse_code`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci COMMENT='产品库存周转快照';
