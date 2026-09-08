-- Feeds data.js -> months[]: vehicles first seen per month, split by fitment,
-- and the unfitted vehicles broken down by powertrain type.
-- Replace the CTEs with the ones in 01_vehicle_classification.sql.

WITH classified AS ( /* ...see 01_vehicle_classification.sql... */ )
SELECT
  FORMAT_DATE('%Y-%m', first_request) AS ym,
  COUNTIF(fitment_status = 'Has fitment')  AS with_sizes,
  COUNTIF(fitment_status <> 'Has fitment') AS without_sizes,
  COUNTIF(fitment_status <> 'Has fitment' AND ice_class = 'EV')      AS unfitted_ev,
  COUNTIF(fitment_status <> 'Has fitment' AND ice_class = 'Non EV')  AS unfitted_non_ev,
  COUNTIF(fitment_status <> 'Has fitment' AND ice_class = 'Unknown') AS unfitted_unknown
FROM classified
GROUP BY ym
ORDER BY ym;
