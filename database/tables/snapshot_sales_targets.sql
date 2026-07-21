-- ========================================
-- 表：snapshot_sales_targets
-- 数据库：rpa-report
-- 说明：月度销售目标拆解快照（来源：月目标拆解及跟进.xlsx → ALL）
--       粒度：target_month + 市场编码 + 账号 + SKU
--       用途：承接平台账号×SKU 月目标（销量/销售额/毛利/费率/费用拆解）及编制时库存
--       金额单位：EUR；占比类字段为小数（0.01 = 1%）
-- 精度规范：
-- 1. 费率/占比：decimal(10,4) 万分位，满足运营目标测算
-- 2. 平均客单/平均销量：decimal(18,4)
-- 3. 财务金额/成本/费用：decimal(18,6) 保留6位，对账分摊不丢精度
-- 4. 库存、退货量等实际件数：int 整数
-- ========================================

SET NAMES utf8mb4 COLLATE utf8mb4_unicode_ci;
SET FOREIGN_KEY_CHECKS = 0;
USE `rpa-report`;

DROP TABLE IF EXISTS `snapshot_sales_targets`;
CREATE TABLE `snapshot_sales_targets` (
  `id` bigint unsigned NOT NULL AUTO_INCREMENT COMMENT '自增主键',
  `target_month` date NOT NULL COMMENT '目标月份（自然月第一天，如 2026-07-01）',

  -- 业务维度（Excel：平台 / 账号 / SKU / 商品ID）
  `market_code` varchar(32) COLLATE utf8mb4_unicode_ci NOT NULL COMMENT '市场/平台编码（如 AMAZON-EU / TEMU-BV）',
  `account_code` varchar(64) COLLATE utf8mb4_unicode_ci NOT NULL COMMENT '账号编码（业务唯一维度）',
  `product_sku` varchar(64) COLLATE utf8mb4_unicode_ci NOT NULL COMMENT 'SKU（业务必填）',
  `product_uid` varchar(64) COLLATE utf8mb4_unicode_ci NOT NULL DEFAULT '' COMMENT '商品ID，例如 25-LYLT-01494',
  `accounting_sku` varchar(64) COLLATE utf8mb4_unicode_ci NOT NULL DEFAULT '' COMMENT '核算SKU',
  `identify_sku` varchar(128) COLLATE utf8mb4_unicode_ci NOT NULL DEFAULT '' COMMENT '识别SKU（账号+SKU，唯一标识）',
  `identify_sku_uid` varchar(128) COLLATE utf8mb4_unicode_ci NOT NULL DEFAULT '' COMMENT '识别商品ID（账号+商品ID）',

  -- 属性维度
  `category` varchar(64) COLLATE utf8mb4_unicode_ci NOT NULL DEFAULT '' COMMENT '品类',
  `product_status` varchar(32) COLLATE utf8mb4_unicode_ci NOT NULL DEFAULT '' COMMENT '产品状态：新品/保留品/不保留老品',
  `ops_owner` varchar(64) COLLATE utf8mb4_unicode_ci NOT NULL DEFAULT '' COMMENT '负责人',
  `dispatch_warehouse` varchar(32) COLLATE utf8mb4_unicode_ci NOT NULL DEFAULT '' COMMENT '发货仓库：HY / 4PX / FBA / MF 等',
  `is_transfer` tinyint(1) NOT NULL DEFAULT '0' COMMENT '是否调拨：0否 1是',

  -- 历史基线（Excel固化均值，费率统一4位小数）
  `hist_avg_aov` decimal(18,4) NOT NULL DEFAULT '0.0000' COMMENT '历史平均客单价（单位：EUR）',
  `hist_avg_sales_qty` decimal(18,4) NOT NULL DEFAULT '0.0000' COMMENT '历史平均销量',
  `hist_avg_gross_margin_rate` decimal(10,4) NOT NULL DEFAULT '0.0000' COMMENT '历史平均毛利率（小数，0.01=1%）',
  `hist_avg_rma_rate` decimal(10,4) NOT NULL DEFAULT '0.0000' COMMENT '历史平均RMA占比（小数，0.01=1%）',
  `hist_avg_ad_rate` decimal(10,4) NOT NULL DEFAULT '0.0000' COMMENT '历史平均广告占比（小数，0.01=1%）',
  `hist_avg_review_rate` decimal(10,4) NOT NULL DEFAULT '0.0000' COMMENT '历史平均测评占比（小数，0.01=1%）',

  -- 目标：费率 / 客单 / 销量 / 销售与毛利
  `target_rma_rate` decimal(10,4) NOT NULL DEFAULT '0.0000' COMMENT 'RMA占比目标（小数，0.01=1%）',
  `target_ad_rate` decimal(10,4) NOT NULL DEFAULT '0.0000' COMMENT '广告占比目标（小数，0.01=1%）',
  `target_review_rate` decimal(10,4) NOT NULL DEFAULT '0.0000' COMMENT '测评占比目标（小数，0.01=1%）',
  `target_aov` decimal(18,4) NOT NULL DEFAULT '0.0000' COMMENT '客单价目标（单位：EUR）',
  `target_sales_qty` decimal(18,4) NOT NULL DEFAULT '0.0000' COMMENT '预估销量',
  `target_platform_sales_amount` decimal(18,6) NOT NULL DEFAULT '0.000000' COMMENT '平台口径销售额目标（单位：EUR）',
  `target_sales_amount` decimal(18,6) NOT NULL DEFAULT '0.000000' COMMENT '公司口径销售额目标（单位：EUR）',
  `target_gross_profit` decimal(18,6) NOT NULL DEFAULT '0.000000' COMMENT '毛利额目标（单位：EUR）',
  `target_operating_gross_profit` decimal(18,6) NOT NULL DEFAULT '0.000000' COMMENT '经营毛利额目标（单位：EUR）',
  `target_gross_margin_rate` decimal(10,4) NOT NULL DEFAULT '0.0000' COMMENT '总目标毛利率（小数，0.01=1%）',

  -- 编制时库存快照（实际件数，int整数）
  `stock_on_hand_qty` int NOT NULL DEFAULT '0' COMMENT '在库库存',
  `stock_in_transit_qty` int NOT NULL DEFAULT '0' COMMENT '在途库存',
  `stock_unfulfilled_qty` int NOT NULL DEFAULT '0' COMMENT '未交库存',
  `stock_total_qty` int NOT NULL DEFAULT '0' COMMENT '总库存（在库+在途+未交）',

  -- 目标：费用拆解（财务金额统一6位小数，单位EUR）
  `target_platform_fee` decimal(18,6) NOT NULL DEFAULT '0.000000' COMMENT '平台费目标',
  `target_sales_tax` decimal(18,6) NOT NULL DEFAULT '0.000000' COMMENT '销售税目标',
  `target_withdrawal_fee` decimal(18,6) NOT NULL DEFAULT '0.000000' COMMENT '提现费目标',
  `target_purchase_cost` decimal(18,6) NOT NULL DEFAULT '0.000000' COMMENT '财务口径采购成本目标',
  `target_purchase_cost_ops` decimal(18,6) NOT NULL DEFAULT '0.000000' COMMENT '经营口径采购成本目标',
  `target_first_leg_tariff` decimal(18,6) NOT NULL DEFAULT '0.000000' COMMENT '头程关税目标',
  `target_last_mile_fee` decimal(18,6) NOT NULL DEFAULT '0.000000' COMMENT '尾程费目标',
  `target_warehouse_rent` decimal(18,6) NOT NULL DEFAULT '0.000000' COMMENT '仓租目标',
  `target_other_allocated_fee` decimal(18,6) NOT NULL DEFAULT '0.000000' COMMENT '其他分摊费用目标',
  `target_seckill_fee` decimal(18,6) NOT NULL DEFAULT '0.000000' COMMENT '秒杀花费目标',
  `target_ad_fee` decimal(18,6) NOT NULL DEFAULT '0.000000' COMMENT '广告花费目标',
  `target_review_fee` decimal(18,6) NOT NULL DEFAULT '0.000000' COMMENT '测评花费目标',
  `target_return_qty` int NOT NULL DEFAULT '0' COMMENT '退货量目标',

  `remark` varchar(512) COLLATE utf8mb4_unicode_ci NOT NULL DEFAULT '' COMMENT '备注',
  `source_sheet` varchar(64) COLLATE utf8mb4_unicode_ci NOT NULL DEFAULT '' COMMENT '来源Sheet（如 AMAZON-EU / TEMU-BV）',

  `created_at` datetime NOT NULL DEFAULT CURRENT_TIMESTAMP COMMENT '写入时间',
  `updated_at` datetime NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP COMMENT '更新时间',
  `deleted_at` datetime NULL DEFAULT NULL COMMENT '删除时间（软删除，NULL=有效）',

  PRIMARY KEY (`id`),
  -- 业务唯一约束：同月+市场+账号+SKU 不可重复，用于导入幂等
  UNIQUE KEY `uk_sst_unique` (`account_code`, `product_sku`, `market_code`, `target_month`),
  -- 核心查询索引：账号+SKU查历史目标
  KEY `idx_sst_account_sku` (`account_code`, `product_sku`, `target_month`),
  -- 按月+市场维度批量查询
  KEY `idx_sst_month_market` (`target_month`, `market_code`),
  -- 商品ID单点查询
  KEY `idx_sst_sku_uid` (`product_uid`),
  -- 运营负责人筛选
  KEY `idx_sst_ops_owner` (`ops_owner`),
  -- 仓库维度筛选
  KEY `idx_sst_warehouse` (`dispatch_warehouse`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci COMMENT='月度销售目标拆解快照';

SET FOREIGN_KEY_CHECKS = 1;