-- 为已有库增加 sales_order_refund.platform_sku（由 sales_order_shipped 回填）
-- 新库请直接用 sales_order_refund.sql 建表，无需执行本脚本。
-- 执行后建议运行：python scripts/dataImport/order_refund.py（导入结束会自动回填）

SET NAMES utf8mb4 COLLATE utf8mb4_unicode_ci;

ALTER TABLE `sales_order_refund`
  ADD COLUMN `platform_sku` VARCHAR(255) NULL
    COMMENT '平台sku（由 sales_order_shipped 按 refund_orig_order_no=order_no 回填）'
    AFTER `rma_product_sku`;
