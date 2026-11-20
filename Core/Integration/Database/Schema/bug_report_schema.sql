-- =========================================================================
-- BUG REPORTS (always lives in schema 'public'; excluded from debug mirror)
-- Idempotent: safe to run on every app start.
-- =========================================================================
CREATE TABLE IF NOT EXISTS public.bug_reports (
    bug_id            BIGSERIAL PRIMARY KEY,
    title             VARCHAR(200) NOT NULL,
    description       TEXT         NOT NULL DEFAULT '',
    source_module     VARCHAR(60)  NOT NULL DEFAULT '',
    active_sheet      VARCHAR(100) NOT NULL DEFAULT '',
    app_version       VARCHAR(40)  NOT NULL DEFAULT '',
    build_mode        VARCHAR(20)  NOT NULL DEFAULT 'Release',
    is_offline        BOOLEAN      NOT NULL DEFAULT FALSE,
    reported_by       VARCHAR(120) NOT NULL DEFAULT '',
    machine_name      VARCHAR(120) NOT NULL DEFAULT '',
    status            VARCHAR(30)  NOT NULL DEFAULT 'Open',
    dev_notes         TEXT         NOT NULL DEFAULT '',
    runtime_log_tail  TEXT         NOT NULL DEFAULT '',
    created_at        TIMESTAMPTZ  NOT NULL DEFAULT CURRENT_TIMESTAMP,
    updated_at        TIMESTAMPTZ  NOT NULL DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE IF NOT EXISTS public.bug_report_attachments (
    attachment_id  BIGSERIAL PRIMARY KEY,
    bug_id         BIGINT       NOT NULL REFERENCES public.bug_reports(bug_id) ON DELETE CASCADE,
    file_name      VARCHAR(255) NOT NULL,
    content_type   VARCHAR(100) NOT NULL DEFAULT 'application/octet-stream',
    file_size      BIGINT       NOT NULL DEFAULT 0,
    file_data      BYTEA        NOT NULL,
    created_at     TIMESTAMPTZ  NOT NULL DEFAULT CURRENT_TIMESTAMP
);

CREATE INDEX IF NOT EXISTS ix_bug_reports_status_created
    ON public.bug_reports (status, created_at DESC);

CREATE INDEX IF NOT EXISTS ix_bug_report_attachments_bug_id
    ON public.bug_report_attachments (bug_id);
