-- ========================================
-- 表：product_sku
-- 数据库：rpa-report
-- 导出时间：2026-06-12 14:17:41
-- 来源：局域网数据库 172.18.188.18:3309
-- ========================================
SET NAMES utf8mb4 COLLATE utf8mb4_unicode_ci;
USE rpa-report;

DROP TABLE IF EXISTS product_sku;
CREATE TABLE `product_sku` (
  `id` bigint unsigned NOT NULL AUTO_INCREMENT COMMENT '自增主键',
  `line_hash` char(64) DEFAULT NULL COMMENT '行内容稳定哈希（SHA-256 hex），用于外部导入变更检测',
  `product_sku` varchar(64) NOT NULL COMMENT '产品SKU编码，唯一标识',
  `product_uid` varchar(64) DEFAULT NULL COMMENT '商品SPU型号，同系列多变体共用',
  `product_name_cn` varchar(255) NOT NULL DEFAULT '' COMMENT '商品中文名称',
  `product_name_en` varchar(255) NOT NULL DEFAULT '' COMMENT '商品英文名称',
  `warehouse_ref` varchar(64) DEFAULT NULL COMMENT '仓库编码',
  `category_lv1` varchar(64) DEFAULT NULL COMMENT '一级分类',
  `category_lv2` varchar(64) DEFAULT NULL COMMENT '二级分类',
  `category_lv3` varchar(64) DEFAULT NULL COMMENT '三级分类',
  `category_code` varchar(16) DEFAULT NULL COMMENT '产品类别编码（扩充长度防止多层编码溢出）',
  `ean_code` varchar(100) NOT NULL DEFAULT '' COMMENT 'EAN条码',
  `supplier_abbr` varchar(60) NOT NULL DEFAULT '' COMMENT '供应商简称',
  `supplier_name` varchar(128) DEFAULT NULL COMMENT '供应商全称',
  `product_unit` varchar(16) DEFAULT NULL COMMENT '计量单位：套/个/箱',
  `product_color` varchar(32) DEFAULT NULL COMMENT '产品颜色',
  `product_img` varchar(255) NOT NULL DEFAULT '' COMMENT '主图URL',
  `declare_price_usd` decimal(10,2) NOT NULL DEFAULT '0.00' COMMENT '报关申报单价(USD)',
  `declare_name_cn` varchar(200) NOT NULL DEFAULT '' COMMENT '海关中文品名',
  `declare_name_en` varchar(200) NOT NULL DEFAULT '' COMMENT '海关英文品名',
  `hs_code` varchar(32) NOT NULL DEFAULT '' COMMENT 'HS海关编码（常规HS最多10位，放宽预留）',
  `amz_lifecycle` varchar(16) DEFAULT NULL COMMENT '亚马逊生命周期：新品/保留品/清仓品',
  `local_lifecycle` varchar(16) DEFAULT NULL COMMENT '本土平台生命周期',
  `accounting_class` varchar(32) DEFAULT NULL COMMENT '核算品类：全新品/保留品/变体新品等',
  `carton_qty` int DEFAULT NULL COMMENT '外箱装箱数量',
  `purchase_moq` int DEFAULT NULL COMMENT '采购最小起订量',
  `purchase_lead_days` smallint DEFAULT NULL COMMENT '采购交期(天)（改用smallint节约空间）',
  `purchase_price` decimal(12,4) DEFAULT NULL COMMENT '采购单价',
  `cost_price_cny` decimal(12,4) DEFAULT NULL COMMENT '人民币成本价',
  `unit_weight_g` decimal(12,2) DEFAULT NULL COMMENT '单件净重(g)',
  `carton_gross_g` decimal(12,2) DEFAULT NULL COMMENT '外箱毛重(g)',
  `inner_box_l_cm` decimal(10,2) DEFAULT NULL COMMENT '内箱长(cm)',
  `inner_box_w_cm` decimal(10,2) DEFAULT NULL COMMENT '内箱宽(cm)',
  `inner_box_h_cm` decimal(10,2) DEFAULT NULL COMMENT '内箱高(cm)',
  `outer_box_l_cm` decimal(10,2) DEFAULT NULL COMMENT '外箱长(cm)',
  `outer_box_w_cm` decimal(10,2) DEFAULT NULL COMMENT '外箱宽(cm)',
  `outer_box_h_cm` decimal(10,2) DEFAULT NULL COMMENT '外箱高(cm)',
  -- 头程、关税精度下调，跨境运价无需6位小数，4位足够
  `first_leg_eu_au_cny` decimal(12,4) DEFAULT NULL COMMENT 'EU/AU头程运费 RMB/件',
  `first_leg_us_cny` decimal(12,4) DEFAULT NULL COMMENT 'US头程运费 RMB/件',
  `first_leg_uk_cny` decimal(12,4) DEFAULT NULL COMMENT 'UK头程运费 RMB/件',
  `duty_eu_cny` decimal(12,4) DEFAULT NULL COMMENT 'EU单件关税 RMB/件',
  `duty_us_cny` decimal(12,4) DEFAULT NULL COMMENT 'US单件关税 RMB/件',
  `duty_uk_cny` decimal(12,4) DEFAULT NULL COMMENT 'UK单件关税 RMB/件',
  `ops_model` varchar(64) NOT NULL DEFAULT '' COMMENT '运营模式',
  `ops_tax_rate` decimal(6,2) DEFAULT NULL COMMENT '运营税率',
  `distribution_lev` tinyint NOT NULL DEFAULT 0 COMMENT '分销等级',
  `source_type` varchar(24) DEFAULT 'Excel' COMMENT '数据来源：Excel/API',
  `is_deleted` tinyint NOT NULL DEFAULT 0 COMMENT '软删除 0正常 1删除',
  `created_at` datetime NOT NULL DEFAULT CURRENT_TIMESTAMP COMMENT '记录创建时间',
  `updated_at` datetime NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP COMMENT '记录更新时间',

  PRIMARY KEY (`id`),
  UNIQUE KEY `uk_product_sku` (`product_sku`),
  -- 联合索引优先，减少单列索引数量（索引过多影响写入性能）
  KEY `idx_product_uid` (`product_uid`),
  KEY `idx_warehouse_ref` (`warehouse_ref`),
  KEY `idx_category_lv2_lv3` (`category_lv2`, `category_lv3`),
  KEY `idx_supplier_name` (`supplier_name`),
  KEY `idx_amz_lifecycle` (`amz_lifecycle`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci COMMENT='商品SKU基础信息表';