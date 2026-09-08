-- One row per vehicle_id, with EV class and tyre fitment state.
-- This is the spine every dashboard figure is derived from.
--
-- Grain note: the source table has 324,350 rows but 322,433 distinct vehicle_ids,
-- because AU and NZ run separate ID sequences and 1,901 ids appear in both.
-- Requests are summed across stores; 102 ids carry conflicting make/model and are
-- genuine collisions between different vehicles, merged here into one row.

WITH ev_lookup AS (
  SELECT UPPER(TRIM(vehicle_make)) AS mk,
         UPPER(TRIM(vehicle_model)) AS md,
         ANY_VALUE(type)       AS ev_type,
         ANY_VALUE(ev_vehicle) AS ev_flag
  FROM `heroic-ruler-198603.RTB.vehicle_catalog_ev_ready`
  GROUP BY 1, 2
),

-- Fitment is resolved on make + model + series_chassis, NOT on vehicle_id.
-- The fitment table's own schema calls its vehicle_id "session-ephemeral; NOT a
-- stable identity key", and testing bears that out: of 50,811 ids that match the
-- log, 99.6% agree on make and model but the year rarely does, so the id lands on
-- the wrong variant. Only 128 of those matches came back fitted (0.25%) against a
-- 34.8% base rate. Do not join these tables on vehicle_id.
fitment AS (
  SELECT UPPER(TRIM(IFNULL(make, '')))  AS mk,
         UPPER(TRIM(IFNULL(model, ''))) AS md,
         REGEXP_REPLACE(UPPER(IFNULL(series_chassis, '')), r'\s+', '') AS sc,
         LOGICAL_OR(has_fitment) AS any_fitment
  FROM `heroic-ruler-198603.tyroola_reporting_prod.vehicle_tyre_fitment`
  GROUP BY 1, 2, 3
),

rows_classified AS (
  SELECT
    v.vehicle_id,
    v.store,
    NULLIF(TRIM(IFNULL(v.make, '')), '')  AS make,
    NULLIF(TRIM(IFNULL(v.model, '')), '') AS model,
    NULLIF(TRIM(IFNULL(v.year, '')), '')  AS year,
    NULLIF(REGEXP_REPLACE(TRIM(IFNULL(v.engine, '')), r'\s+', ' '), '') AS engine_desc,
    UPPER(TRIM(IFNULL(v.make, '')))  AS mk,
    UPPER(TRIM(IFNULL(v.model, ''))) AS md,
    REGEXP_REPLACE(UPPER(IFNULL(v.series_chassis, '')), r'\s+', '') AS sc,
    v.request_count_without_tyre_options AS req,
    -- updated_at sits BEFORE created_at on 313,070 of 324,350 rows, so neither
    -- column alone can be trusted as the later timestamp.
    GREATEST(v.created_at, v.updated_at) AS last_ts,
    LEAST(v.created_at, v.updated_at)    AS first_ts,
    CASE
      WHEN REGEXP_CONTAINS(UPPER(IFNULL(v.engine, '')), r'\bELEC\b|\bBEV\b') THEN 'EV'
      WHEN l.ev_flag IS TRUE  THEN 'EV'
      WHEN l.ev_flag IS FALSE THEN 'Non EV'
      WHEN v.engine IS NULL OR TRIM(v.engine) = '' THEN 'Unknown'
      ELSE 'Non EV'
    END AS ice_class
  FROM `heroic-ruler-198603.popeye_mirror.vehicle_log` v
  LEFT JOIN ev_lookup l
    ON l.mk = UPPER(TRIM(IFNULL(v.make, '')))
   AND l.md = UPPER(TRIM(IFNULL(v.model, '')))
),

per_vehicle AS (
  SELECT
    vehicle_id,
    ANY_VALUE(make) AS make, ANY_VALUE(model) AS model, ANY_VALUE(year) AS year,
    ANY_VALUE(engine_desc) AS engine_desc,
    ANY_VALUE(mk) AS mk, ANY_VALUE(md) AS md, ANY_VALUE(sc) AS sc,
    SUM(req) AS req,
    COUNT(DISTINCT store) AS stores,
    DATE(MAX(last_ts))  AS last_request,
    DATE(MIN(first_ts)) AS first_request,
    CASE WHEN LOGICAL_OR(ice_class = 'EV')     THEN 'EV'
         WHEN LOGICAL_OR(ice_class = 'Non EV') THEN 'Non EV'
         ELSE 'Unknown' END AS ice_class
  FROM rows_classified
  GROUP BY vehicle_id
)

SELECT
  p.*,
  CASE WHEN p.mk = ''         THEN 'Not checkable'
       WHEN f.mk IS NULL      THEN 'No spec match'
       WHEN f.any_fitment     THEN 'Has fitment'
       ELSE 'No fitment' END AS fitment_status
FROM per_vehicle p
LEFT JOIN fitment f ON f.mk = p.mk AND f.md = p.md AND f.sc = p.sc;
