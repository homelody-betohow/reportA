-- ========================================
-- 表：warehouse
-- 数据库：rpa-report
-- ========================================
SET NAMES utf8mb4 COLLATE utf8mb4_unicode_ci;
USE rpa-report;

DROP TABLE IF EXISTS warehouse;
CREATE TABLE `warehouse` (
  `id` bigint unsigned NOT NULL AUTO_INCREMENT COMMENT '自增主键',
  `warehouse_id` int NOT NULL COMMENT '易仓仓库ID（getWarehouseList.warehouse_id）',
  `warehouse_code` varchar(128) COLLATE utf8mb4_unicode_ci NOT NULL COMMENT '仓库代码',
  `warehouse_name` varchar(128) COLLATE utf8mb4_unicode_ci NOT NULL COMMENT '仓库名称',
  `warehouse_type` tinyint NOT NULL COMMENT '易仓仓库类型（如 1/2/3）',
  `warehouse_virtual` tinyint NOT NULL COMMENT '运营模式',
  `warehouse_status` tinyint NOT NULL COMMENT '易仓状态：-1=已废弃 0=不可用 1=可用',
  `warehouse_service` varchar(24) COLLATE utf8mb4_unicode_ci NOT NULL DEFAULT '' COMMENT '仓库服务：FBA、FBLM、4PX 等',
  `country_code` varchar(16) COLLATE utf8mb4_unicode_ci NOT NULL DEFAULT '' COMMENT '国家编码',
  `provider_code` varchar(16) COLLATE utf8mb4_unicode_ci NOT NULL DEFAULT '' COMMENT '服务商/来源：HY、4PX、FBA 等（可人工维护）',
  `market_code` varchar(64) COLLATE utf8mb4_unicode_ci NOT NULL DEFAULT '' COMMENT '市场编码',
  `market_region` varchar(64) COLLATE utf8mb4_unicode_ci NOT NULL DEFAULT '' COMMENT '市场区域',
  `ops_owner` varchar(128) COLLATE utf8mb4_unicode_ci NOT NULL DEFAULT '' COMMENT '负责人',
  `source` varchar(32) COLLATE utf8mb4_unicode_ci NOT NULL DEFAULT 'eccang' COMMENT '数据来源',
  `is_transfer` tinyint NOT NULL COMMENT '是否为中转仓',
  `created_at` datetime NOT NULL DEFAULT CURRENT_TIMESTAMP COMMENT '写入时间',
  `updated_at` datetime NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP COMMENT '更新时间',
  PRIMARY KEY (`id`),
  UNIQUE KEY `uk_wh_warehouse_id` (`warehouse_id`),
  UNIQUE KEY `uk_wh_warehouse_code` (`warehouse_code`),
  KEY `idx_wh_status` (`warehouse_status`),
  KEY `idx_wh_type` (`warehouse_type`),
  KEY `idx_wh_provider` (`provider`),
  KEY `idx_wh_name` (`warehouse_name`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci COMMENT='仓库主数据表';
