-- ========================================
-- 表：amz_seckill_cost
-- 说明：亚马逊秒杀（Lightning Deal）费用花费明细
-- ========================================
SET NAMES utf8mb4 COLLATE utf8mb4_unicode_ci;

CREATE TABLE `amz_seckill_cost` (
  `id` bigint unsigned NOT NULL AUTO_INCREMENT COMMENT '自增主键',
  `shop_name_en` varchar(128) COLLATE utf8mb4_unicode_ci NOT NULL COMMENT '店铺名称',
  `shop_alias` varchar(128) NOT NULL DEFAULT '' COMMENT '店铺别名',
  `marketplace` varchar(32) COLLATE utf8mb4_unicode_ci NOT NULL COMMENT '站点',
  `promotion_id` varchar(100) COLLATE utf8mb4_unicode_ci NOT NULL COMMENT '关联秒杀活动ID',
  `promotion_name` varchar(255) NOT NULL DEFAULT '' COMMENT '秒杀活动名称',
  `seckill_goods` varchar(255) COLLATE utf8mb4_unicode_ci NOT NULL COMMENT '秒杀商品',
  `seckill_asin` varchar(32) COLLATE utf8mb4_unicode_ci NOT NULL COMMENT '秒杀商品ASIN',
  `seckill_sku` varchar(64) COLLATE utf8mb4_unicode_ci NOT NULL COMMENT '秒杀商品SKU',
  `seckill_fee` varchar(100) NOT NULL DEFAULT '' COMMENT '秒杀费用',
  `seckill_sales` decimal(10,2) NOT NULL DEFAULT 0 COMMENT '销售金额',
  `units_sold` int NOT NULL DEFAULT 0 COMMENT '销售数量',
  `glance_views` int NOT NULL DEFAULT 0 COMMENT '浏览量',
  `conversion_rate` decimal(10,4) DEFAULT NULL COMMENT '转化率（小数形式，0.125 表示 12.5%）',
  `start_date` date NOT NULL COMMENT '秒杀开始日期',
  `end_date` date NOT NULL COMMENT '秒杀结束日期',
  `seckill_status` tinyint NOT NULL COMMENT '秒杀状态：0未开始 1进行中 2已结束',
  `settle_status` tinyint NOT NULL DEFAULT 0 COMMENT '0未结算 1已结算 2对账差异待处理',
  `charge_date` date NOT NULL COMMENT '扣费/产生日期',
  `charge_amount` decimal(10,2) NOT NULL DEFAULT 0 COMMENT '扣费金额',
  `currency_code` varchar(8) COLLATE utf8mb4_unicode_ci NOT NULL COMMENT '币种 USD/EUR/GBP/JPY',
  `settle_batch_no` varchar(64) COLLATE utf8mb4_unicode_ci DEFAULT NULL COMMENT '亚马逊结算批次号',
  `remark` varchar(512) COLLATE utf8mb4_unicode_ci DEFAULT NULL COMMENT '备注、对账说明',
  `created_at` timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP COMMENT '写入时间',
  `updated_at` timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP COMMENT '更新时间',
  PRIMARY KEY (`id`),
  UNIQUE KEY `uk_promo_sku_charge` (`shop_name_en`,`marketplace`,`promotion_id`),
  KEY `idx_shop_market` (`shop_name_en`,`marketplace`),
  KEY `idx_promotion_id` (`promotion_id`),
  KEY `idx_seckill_goods` (`seckill_goods`),
  KEY `idx_charge_date` (`charge_date`),
  KEY `idx_settle_status` (`settle_status`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci COMMENT='亚马逊秒杀费用明细表';
