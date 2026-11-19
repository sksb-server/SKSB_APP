-- Schedule Manager schema (idempotent). Also executed automatically by ServerDispatcher.EnsureScheduleSchema()
-- on the first schedule save / load. No foreign keys: StaffManager full-batch saves delete orphaned staff rows,
-- and schedule history (attendance, premix holes) must survive that.

CREATE TABLE IF NOT EXISTS daily_project_schedules (
    schedule_date DATE         NOT NULL,
    shift         VARCHAR(20)  NOT NULL,            -- Morning (Day Work band) | Night (Night Work band)
    slot_index    INTEGER      NOT NULL CHECK (slot_index >= 0),
    project_id    VARCHAR(50)  NOT NULL DEFAULT '',
    premix_holes  INTEGER      NOT NULL DEFAULT 0,
    workers       TEXT[]       NOT NULL DEFAULT '{}',
    vehicles      TEXT[]       NOT NULL DEFAULT '{}',
    remarks       TEXT         NOT NULL DEFAULT '',
    created_at    TIMESTAMPTZ  NOT NULL DEFAULT CURRENT_TIMESTAMP,
    updated_at    TIMESTAMPTZ  NOT NULL DEFAULT CURRENT_TIMESTAMP,
    PRIMARY KEY (schedule_date, shift, slot_index)
);

CREATE TABLE IF NOT EXISTS attendance_entries (
    entry_id         BIGSERIAL    PRIMARY KEY,
    entry_date       DATE         NOT NULL,
    entry_time       TIME         NOT NULL,
    staff_id         BIGINT       NOT NULL,
    project_id       VARCHAR(50)  NOT NULL DEFAULT '',
    entry_type       VARCHAR(20)  NOT NULL,         -- CheckIn | CheckOut
    source           VARCHAR(20)  NOT NULL,         -- Thumbprint | Manual | SKSB_App | WhatsApp
    override_reason  TEXT         NOT NULL DEFAULT '',
    is_approved      BOOLEAN      NOT NULL DEFAULT TRUE,
    whatsapp_picture TEXT         NOT NULL DEFAULT '',
    created_at       TIMESTAMPTZ  NOT NULL DEFAULT CURRENT_TIMESTAMP,
    updated_at       TIMESTAMPTZ  NOT NULL DEFAULT CURRENT_TIMESTAMP,
    CONSTRAINT uq_attendance_entries_natural UNIQUE (entry_date, entry_time, staff_id, entry_type, source)
);
CREATE INDEX IF NOT EXISTS ix_attendance_entries_date ON attendance_entries (entry_date);

CREATE TABLE IF NOT EXISTS daily_attendances (
    schedule_date         DATE          NOT NULL,
    staff_id              BIGINT        NOT NULL,
    attendance_status     VARCHAR(20)   NOT NULL DEFAULT 'Absence',
    clock_in_time         TIME,
    clock_out_time        TIME,
    project_check_ins     TEXT          NOT NULL DEFAULT '',
    projects_done         TEXT[]        NOT NULL DEFAULT '{}',
    regular_shift         NUMERIC(6,2)  NOT NULL DEFAULT 0,
    additional_shifts     NUMERIC(6,2)  NOT NULL DEFAULT 0,
    ot_hours              NUMERIC(6,2)  NOT NULL DEFAULT 0,
    additional_ot_hours   NUMERIC(6,2)  NOT NULL DEFAULT 0,
    total_flattened_hours NUMERIC(8,2)  NOT NULL DEFAULT 0,
    is_allocated         BOOLEAN       NOT NULL DEFAULT FALSE,
    created_at            TIMESTAMPTZ   NOT NULL DEFAULT CURRENT_TIMESTAMP,
    updated_at            TIMESTAMPTZ   NOT NULL DEFAULT CURRENT_TIMESTAMP,
    PRIMARY KEY (schedule_date, staff_id)
);

-- Per-date attendance parameters (Daily Attendances C2 / D2 / E2). Missing row => defaults 06:00 / 11:00 / 18:00.
CREATE TABLE IF NOT EXISTS daily_schedule_settings (
    schedule_date     DATE         PRIMARY KEY,
    day_start_time    TIME         NOT NULL DEFAULT '06:00:00',
    late_cut_off_time TIME         NOT NULL DEFAULT '11:00:00',
    ot_start_time     TIME         NOT NULL DEFAULT '18:00:00',
    created_at        TIMESTAMPTZ  NOT NULL DEFAULT CURRENT_TIMESTAMP,
    updated_at        TIMESTAMPTZ  NOT NULL DEFAULT CURRENT_TIMESTAMP
);
ALTER TABLE daily_schedule_settings ADD COLUMN IF NOT EXISTS day_start_time TIME NOT NULL DEFAULT '06:00:00';

CREATE TABLE IF NOT EXISTS worker_preferences (
    staff_id                BIGINT       PRIMARY KEY,
    team                    TEXT[]       NOT NULL DEFAULT '{}',
    primary_skill           VARCHAR(100) NOT NULL DEFAULT '',
    other_skills            TEXT[]       NOT NULL DEFAULT '{}',
    blacklisted_skills      TEXT[]       NOT NULL DEFAULT '{}',
    preferred_vehicle_type  VARCHAR(50)  NOT NULL DEFAULT 'None',
    dedicated_vehicle_plate VARCHAR(50)  NOT NULL DEFAULT '',
    preferred_shift         VARCHAR(20)  NOT NULL DEFAULT 'None',
    max_days_per_week       INTEGER      NOT NULL DEFAULT 0,
    rest_day_preference     VARCHAR(20)  NOT NULL DEFAULT 'None',
    created_at              TIMESTAMPTZ  NOT NULL DEFAULT CURRENT_TIMESTAMP,
    updated_at              TIMESTAMPTZ  NOT NULL DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE IF NOT EXISTS lookup_schedule_skills (
    skill_name   VARCHAR(100) PRIMARY KEY,
    display_icon VARCHAR(32)  NOT NULL DEFAULT '',
    updated_at   TIMESTAMPTZ  NOT NULL DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE IF NOT EXISTS teams_master (
    team_name            VARCHAR(100) PRIMARY KEY,
    team_leader_staff_id BIGINT,
    default_vehicles     TEXT[]       NOT NULL DEFAULT '{}',
    created_at           TIMESTAMPTZ  NOT NULL DEFAULT CURRENT_TIMESTAMP,
    updated_at           TIMESTAMPTZ  NOT NULL DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE IF NOT EXISTS project_skill_requirements (
    project_type      VARCHAR(50)  PRIMARY KEY,
    required_skills   TEXT[]       NOT NULL DEFAULT '{}',
    required_vehicles TEXT[]       NOT NULL DEFAULT '{}',
    updated_at        TIMESTAMPTZ  NOT NULL DEFAULT CURRENT_TIMESTAMP
);
ALTER TABLE project_skill_requirements
    ADD COLUMN IF NOT EXISTS required_vehicles TEXT[] NOT NULL DEFAULT '{}';
