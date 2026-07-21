-- ========================================
-- 表：product_sku_owners
-- 数据库：rpa-report
-- 说明：产品×市场 运营/开发负责人主数据
--       粒度：product_sku + market_site + market_code + market_region（一行一条归属）
--       用途：回填快照表 ops_owner / ops_leader；按负责人筛市场周转等
-- ========================================
SET NAMES utf8mb4 COLLATE utf8mb4_unicode_ci;
USE `rpa-report`;

DROP TABLE IF EXISTS `product_sku_owners`;
CREATE TABLE `product_sku_owners` (
  `id` bigint unsigned NOT NULL AUTO_INCREMENT COMMENT '自增主键',
  `sku_relation` char(64) COLLATE utf8mb4_unicode_ci NOT NULL COMMENT '行内容稳定哈希（SHA-256 hex），变更检测 / 幂等 UPSERT',

  -- 业务键
  `product_sku` varchar(64) COLLATE utf8mb4_unicode_ci NOT NULL COMMENT '产品SKU',
  `market_code` varchar(64) COLLATE utf8mb4_unicode_ci NOT NULL COMMENT '市场编码（如 AMAZON-EU）',
  `market_region` varchar(64) COLLATE utf8mb4_unicode_ci NOT NULL DEFAULT '' COMMENT '市场区域',
  `market_site` varchar(128) COLLATE utf8mb4_unicode_ci NOT NULL DEFAULT '' COMMENT '市场站点',

  -- 冗余属性（写入时固化，查询不依赖 product_sku 主表）
  `product_uid` varchar(64) COLLATE utf8mb4_unicode_ci NOT NULL DEFAULT '' COMMENT '商品ID/SPU',
  `sku_lifecycle` varchar(32) COLLATE utf8mb4_unicode_ci NOT NULL DEFAULT '' COMMENT '产品状态：新品/保留品/不保留老品等',
  `supplier_abbr` varchar(64) COLLATE utf8mb4_unicode_ci NOT NULL DEFAULT '' COMMENT '供应商简称',
  `supplier_name` varchar(128) COLLATE utf8mb4_unicode_ci NOT NULL DEFAULT '' COMMENT '供应商名称',
  `distribution_lev` tinyint(1) NOT NULL DEFAULT '0' COMMENT '分销等级',

  -- 负责人
  `ops_leader` varchar(64) COLLATE utf8mb4_unicode_ci NOT NULL DEFAULT '' COMMENT '运营经理',
  `ops_owner` varchar(64) COLLATE utf8mb4_unicode_ci NOT NULL DEFAULT '' COMMENT '运营负责人',
  `dev_leader` varchar(64) COLLATE utf8mb4_unicode_ci NOT NULL DEFAULT '' COMMENT '开发经理',
  `dev_owner` varchar(64) COLLATE utf8mb4_unicode_ci NOT NULL DEFAULT '' COMMENT '开发负责人',

  `is_enabled` tinyint(1) NOT NULL DEFAULT '1' COMMENT '是否启用：1启用 0停用',
  `created_at` datetime NOT NULL DEFAULT CURRENT_TIMESTAMP COMMENT '写入时间',
  `updated_at` datetime NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP COMMENT '更新时间',

  PRIMARY KEY (`id`),
  -- 业务唯一：同 SKU + 市场 + 区域只保留一行（负责人变更走 UPDATE，不另起行）
  UNIQUE KEY `uk_pso_sku_market` (`product_sku`,`market_site`, `market_code`, `market_region`),
  -- 变更检测 / 导入去重
  UNIQUE KEY `uk_pso_sku_relation` (`sku_relation`),
  -- 按运营负责人筛市场 SKU
  KEY `idx_pso_ops_owner` (`ops_owner`, `market_code`, `product_sku`),
  -- 按运营经理筛
  KEY `idx_pso_ops_leader` (`ops_leader`, `market_code`, `product_sku`),
  -- 商品ID 回查
  KEY `idx_pso_product_uid` (`product_uid`),
  -- 启用态 + 市场批量拉取
  KEY `idx_pso_enabled_market` (`is_enabled`, `market_code`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci COMMENT='产品×市场运营/开发负责人主数据';
