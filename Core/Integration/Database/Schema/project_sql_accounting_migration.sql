-- =========================================================================
-- SQL ACCOUNTING PROJECT MAPPING (ProjectManager "Other Projects" phase 1)
-- lookup_sql_account_projects mirrors SQL Accounting (Firebird ACC-0014.FDB) PROJECT codes;
-- company_master / project_master gain the mapped SQL Accounting project code.
-- Idempotent: safe to run on every app start.
-- =========================================================================
CREATE TABLE IF NOT EXISTS lookup_sql_account_projects (
    code        VARCHAR(80)  PRIMARY KEY,
    description VARCHAR(320),
    is_active   BOOLEAN      NOT NULL DEFAULT TRUE,
    updated_at  TIMESTAMPTZ  NOT NULL DEFAULT CURRENT_TIMESTAMP
);

ALTER TABLE company_master ADD COLUMN IF NOT EXISTS default_sql_account_project VARCHAR(80);
ALTER TABLE project_master ADD COLUMN IF NOT EXISTS sql_account_project_code    VARCHAR(80);
