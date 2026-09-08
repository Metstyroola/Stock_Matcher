-- Feeds data.js -> types[] and lastRequestedYear[].

-- Type table
WITH classified AS ( /* ...see 01_vehicle_classification.sql... */ )
SELECT
  ice_class AS type,
  COUNT(*) AS vehicles,
  COUNTIF(fitment_status = 'Has fitment')  AS has_sizes,
  COUNTIF(fitment_status = 'No fitment')   AS matched_no_sizes,
  COUNTIF(fitment_status = 'No spec match') AS no_spec_match,
  COUNTIF(fitment_status = 'Not checkable') AS no_make
FROM classified
GROUP BY type;

-- Last-requested year
WITH classified AS ( /* ...see 01_vehicle_classification.sql... */ )
SELECT
  EXTRACT(YEAR FROM last_request) AS year,
  COUNTIF(fitment_status = 'Has fitment')  AS with_sizes,
  COUNTIF(fitment_status <> 'Has fitment') AS without_sizes
FROM classified
GROUP BY year
ORDER BY year;
