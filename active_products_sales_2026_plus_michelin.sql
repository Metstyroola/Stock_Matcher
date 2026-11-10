-- Active products in the Stock Matcher AU list (suppliers_check_dashboard_table,
-- is_active AND is_active_stock, matcher's 25-supplier set) with AU sales since 2026-01-01,
-- PLUS every Michelin product on that list whether it sold or not.
WITH list AS (
  SELECT spinach_pid AS spinach_id, ANY_VALUE(manufacturer) manufacturer, ANY_VALUE(profile) profile,
    ANY_VALUE(dimensions) dimensions, COUNT(DISTINCT `key`) n_suppliers,
    STRING_AGG(DISTINCT `key`, '|' ORDER BY `key`) suppliers
  FROM `heroic-ruler-198603.Sales_Dashboard_Views.suppliers_check_dashboard_table`
  WHERE is_active AND is_active_stock AND spinach_pid IS NOT NULL AND spinach_pid != ''
    AND `key` IN ('atlastyres','bsatyres','chapelcorner','continental','cypresstyres','freedomtyres','goodride','goodride-wa','logictyres','michelin','neta','newbee','onyx','pirelli','spmotorcycles','stgeorge','superior','tempetyres','toptyres','townsend','tyredepot','tyrenetwork','tyreprofessionals','yhi','yokohama')
  GROUP BY 1),
s AS (  -- canonical billable item filter, all AU stores (B2C, NRMA, SCA, B2B)
  SELECT spinach_id, SUM(pcs_unit) units, COUNT(DISTINCT fk_sales_order) orders,
    ROUND(SUM(sr_fullDS_total_rev),2) revenue, MIN(created_at_aus) first_sale, MAX(created_at_aus) last_sale,
    SUM(IF(channel='B2C',pcs_unit,0)) units_b2c, SUM(IF(channel='B2B',pcs_unit,0)) units_b2b, SUM(IF(channel='SCA',pcs_unit,0)) units_sca
  FROM `heroic-ruler-198603.Sales_Dashboard_Views.sales_report_full_datastudio_v2`
  WHERE billable AND item_billable AND created_at_aus >= '2026-01-01'
    AND store_channel IN ('AU','AU_SCA','AU_NRMA','AU_B2B')
  GROUP BY 1),
out AS (
  SELECT l.*, IFNULL(s.units,0) units_2026, IFNULL(s.orders,0) orders_2026, IFNULL(s.revenue,0) revenue_2026,
    IFNULL(s.units_b2c,0) units_b2c, IFNULL(s.units_b2b,0) units_b2b, IFNULL(s.units_sca,0) units_sca,
    s.first_sale first_sale_2026, s.last_sale last_sale_2026, s.spinach_id IS NOT NULL has_sales_2026,
    l.manufacturer='MICHELIN' is_michelin
  FROM list l LEFT JOIN s USING (spinach_id)
  WHERE s.spinach_id IS NOT NULL OR l.manufacturer='MICHELIN')
SELECT * FROM out
ORDER BY is_michelin DESC, units_2026 DESC, spinach_id;
