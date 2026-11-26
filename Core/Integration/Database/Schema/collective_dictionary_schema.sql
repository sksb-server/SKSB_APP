-- =========================================================================
-- COLLECTIVE DICTIONARY (always lives in schema 'app'; never mirrored to 'debug')
-- Idempotent: safe to run on every app start.
-- =========================================================================
CREATE SCHEMA IF NOT EXISTS app;

CREATE TABLE IF NOT EXISTS app.localization_overrides (
    dict_type            VARCHAR(20)  NOT NULL,
    entry_key            VARCHAR(200) NOT NULL,
    lang_code            VARCHAR(16)  NOT NULL,
    value                TEXT         NOT NULL,
    source_suggestion_id BIGINT       NULL,
    approved_by          VARCHAR(120) NOT NULL DEFAULT '',
    approved_at          TIMESTAMPTZ  NOT NULL DEFAULT CURRENT_TIMESTAMP,
    PRIMARY KEY (dict_type, entry_key, lang_code)
);

CREATE TABLE IF NOT EXISTS app.localization_suggestions (
    suggestion_id   BIGSERIAL    PRIMARY KEY,
    dict_type       VARCHAR(20)  NOT NULL,
    entry_key       VARCHAR(200) NOT NULL,
    lang_code       VARCHAR(16)  NOT NULL,
    suggested_value TEXT         NOT NULL,
    note            TEXT         NOT NULL DEFAULT '',
    suggested_by    VARCHAR(120) NOT NULL DEFAULT '',
    machine_name    VARCHAR(120) NOT NULL DEFAULT '',
    status          VARCHAR(20)  NOT NULL DEFAULT 'pending',
    vote_count      INTEGER      NOT NULL DEFAULT 1,
    created_at      TIMESTAMPTZ  NOT NULL DEFAULT CURRENT_TIMESTAMP,
    reviewed_by     VARCHAR(120) NOT NULL DEFAULT '',
    reviewed_at     TIMESTAMPTZ  NULL
);

CREATE TABLE IF NOT EXISTS app.localization_votes (
    suggestion_id BIGINT       NOT NULL REFERENCES app.localization_suggestions(suggestion_id) ON DELETE CASCADE,
    voter_key     VARCHAR(200) NOT NULL,
    voter_name    VARCHAR(120) NOT NULL DEFAULT '',
    voted_at      TIMESTAMPTZ  NOT NULL DEFAULT CURRENT_TIMESTAMP,
    PRIMARY KEY (suggestion_id, voter_key)
);

CREATE TABLE IF NOT EXISTS app.localization_history (
    history_id     BIGSERIAL    PRIMARY KEY,
    dict_type      VARCHAR(20)  NOT NULL,
    entry_key      VARCHAR(200) NOT NULL,
    lang_code      VARCHAR(16)  NOT NULL,
    previous_value TEXT         NULL,
    new_value      TEXT         NULL,
    action         VARCHAR(30)  NOT NULL,
    changed_by     VARCHAR(120) NOT NULL DEFAULT '',
    changed_at     TIMESTAMPTZ  NOT NULL DEFAULT CURRENT_TIMESTAMP
);

CREATE INDEX IF NOT EXISTS ix_loc_suggestions_status_lang
    ON app.localization_suggestions (status, lang_code, vote_count DESC, created_at DESC);

CREATE INDEX IF NOT EXISTS ix_loc_history_key
    ON app.localization_history (dict_type, entry_key, lang_code, changed_at DESC);
