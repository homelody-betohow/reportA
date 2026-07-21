-- snapshot_product_inventory 增加 is_transfer，供 up_market_data.py 从 warehouse 回填
USE `rpa-report`;

ALTER TABLE `snapshot_product_inventory`
  ADD COLUMN `is_transfer` tinyint(1) DEFAULT NULL
    COMMENT '是否为中转仓（回填自 warehouse.is_transfer）'
    AFTER `market_region`;
