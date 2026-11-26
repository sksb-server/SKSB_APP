-- =========================================================================
-- APP SETTINGS / TELEMETRY (schema 'app'; shared, never mirrored to debug)
-- Idempotent: safe to run on every app start.
-- =========================================================================
CREATE SCHEMA IF NOT EXISTS app;

CREATE TABLE IF NOT EXISTS app.operation_timings (
    operation_key TEXT             NOT NULL,
    machine_name  TEXT             NOT NULL, -- PC hostname or '__global__'
    avg_ms        DOUBLE PRECISION NOT NULL,
    min_ms        DOUBLE PRECISION NOT NULL,
    max_ms        DOUBLE PRECISION NOT NULL,
    last_ms       DOUBLE PRECISION NOT NULL,
    sample_count  INTEGER          NOT NULL DEFAULT 1,
    updated_at    TIMESTAMPTZ      NOT NULL DEFAULT CURRENT_TIMESTAMP,
    PRIMARY KEY (operation_key, machine_name)
);

-- relative_duration_ratio: < 1.0 = faster than server average, > 1.0 = slower.
CREATE OR REPLACE VIEW app.pc_speed_ranking AS
SELECT
    p.machine_name,
    COUNT(*)                                               AS operations_measured,
    SUM(p.sample_count)                                    AS total_samples,
    ROUND(AVG(p.avg_ms / NULLIF(g.avg_ms, 0))::numeric, 2) AS relative_duration_ratio,
    ROUND(AVG(p.avg_ms)::numeric, 0)                       AS avg_ms,
    MAX(p.updated_at)                                      AS last_seen_at
FROM app.operation_timings p
JOIN app.operation_timings g
  ON g.operation_key = p.operation_key
 AND g.machine_name = '__global__'
WHERE p.machine_name <> '__global__'
GROUP BY p.machine_name
ORDER BY relative_duration_ratio ASC;
