BEGIN;

CREATE TABLE IF NOT EXISTS sp_runs (
    run_id uuid PRIMARY KEY,
    run_type text NOT NULL,
    started_at timestamptz NOT NULL DEFAULT now(),
    finished_at timestamptz NULL,
    status text NOT NULL DEFAULT 'running',
    source_site_url text NULL,
    target_site_url text NULL,
    source_tenant_id text NULL,
    target_tenant_id text NULL,
    source_drive_id text NULL,
    target_drive_id text NULL,
    source_library text NULL,
    target_library text NULL,
    total_items bigint NOT NULL DEFAULT 0,
    total_permissions bigint NOT NULL DEFAULT 0,
    total_errors bigint NOT NULL DEFAULT 0,
    notes text NULL
);

CREATE INDEX IF NOT EXISTS idx_sp_runs_type_started_at ON sp_runs(run_type, started_at DESC);

CREATE TABLE IF NOT EXISTS sp_items (
    item_pk bigserial PRIMARY KEY,
    run_id uuid NOT NULL REFERENCES sp_runs(run_id) ON DELETE CASCADE,
    site_url text NOT NULL,
    drive_id text NULL,
    item_id text NULL,
    path_original text NOT NULL,
    path_translated text NULL,
    item_type text NOT NULL,
    is_folder boolean NOT NULL DEFAULT false,
    size_bytes bigint NULL,
    source_version_count integer NOT NULL DEFAULT 0,
    copied_versions boolean NOT NULL DEFAULT false,
    overwrite_enabled boolean NOT NULL DEFAULT false,
    copy_status text NOT NULL DEFAULT 'pending',
    copied_at timestamptz NULL,
    created_at timestamptz NOT NULL DEFAULT now(),
    updated_at timestamptz NOT NULL DEFAULT now()
);

CREATE UNIQUE INDEX IF NOT EXISTS ux_sp_items_run_path ON sp_items(run_id, path_original);
CREATE INDEX IF NOT EXISTS idx_sp_items_run ON sp_items(run_id);
CREATE INDEX IF NOT EXISTS idx_sp_items_status ON sp_items(copy_status);

CREATE TABLE IF NOT EXISTS sp_permissions (
    perm_id bigserial PRIMARY KEY,
    run_id uuid NOT NULL REFERENCES sp_runs(run_id) ON DELETE CASCADE,
    item_pk bigint NULL REFERENCES sp_items(item_pk) ON DELETE SET NULL,
    site_url text NOT NULL,
    drive_id text NULL,
    item_id text NULL,
    path_original text NOT NULL,
    path_translated text NULL,
    item_type text NOT NULL,
    granted_to text NOT NULL,
    granted_to_id text NULL,
    target_type text NOT NULL,
    role_name text NOT NULL,
    is_sharepoint_group boolean NOT NULL DEFAULT false,
    mapped_destination text NULL,
    mapped_destination_id text NULL,
    permission_status text NOT NULL DEFAULT 'detected',
    phase text NULL,
    error_message text NULL,
    created_at timestamptz NOT NULL DEFAULT now()
);

CREATE INDEX IF NOT EXISTS idx_sp_permissions_run ON sp_permissions(run_id);
CREATE INDEX IF NOT EXISTS idx_sp_permissions_path ON sp_permissions(path_original);
CREATE INDEX IF NOT EXISTS idx_sp_permissions_status ON sp_permissions(permission_status);

CREATE TABLE IF NOT EXISTS sp_user_mapping (
    mapping_id bigserial PRIMARY KEY,
    source_id text NULL,
    source_display_name text NOT NULL,
    target_type text NOT NULL,
    destination_display_name text NULL,
    destination_object_id text NULL,
    is_active boolean NOT NULL DEFAULT true,
    created_at timestamptz NOT NULL DEFAULT now(),
    updated_at timestamptz NOT NULL DEFAULT now()
);

CREATE UNIQUE INDEX IF NOT EXISTS ux_sp_user_mapping_source_id ON sp_user_mapping(source_id) WHERE source_id IS NOT NULL;
CREATE UNIQUE INDEX IF NOT EXISTS ux_sp_user_mapping_source_name ON sp_user_mapping(source_display_name);

CREATE TABLE IF NOT EXISTS sp_errors (
    error_id bigserial PRIMARY KEY,
    run_id uuid NOT NULL REFERENCES sp_runs(run_id) ON DELETE CASCADE,
    item_pk bigint NULL REFERENCES sp_items(item_pk) ON DELETE SET NULL,
    path text NULL,
    phase text NULL,
    error_message text NOT NULL,
    error_level text NOT NULL DEFAULT 'ERROR',
    created_at timestamptz NOT NULL DEFAULT now()
);

CREATE INDEX IF NOT EXISTS idx_sp_errors_run ON sp_errors(run_id);
CREATE INDEX IF NOT EXISTS idx_sp_errors_phase ON sp_errors(phase);

CREATE TABLE IF NOT EXISTS sp_logs (
    log_id bigserial PRIMARY KEY,
    run_id uuid NULL REFERENCES sp_runs(run_id) ON DELETE SET NULL,
    run_type text NOT NULL,
    log_level text NOT NULL,
    phase text NULL,
    path text NULL,
    message text NOT NULL,
    context_json text NULL,
    created_at timestamptz NOT NULL DEFAULT now()
);

CREATE INDEX IF NOT EXISTS idx_sp_logs_run ON sp_logs(run_id);
CREATE INDEX IF NOT EXISTS idx_sp_logs_type_level_time ON sp_logs(run_type, log_level, created_at DESC);
CREATE INDEX IF NOT EXISTS idx_sp_logs_created_at ON sp_logs(created_at DESC);

COMMIT;
