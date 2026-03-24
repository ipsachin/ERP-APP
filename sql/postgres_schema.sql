CREATE TABLE IF NOT EXISTS parts (
    part_id TEXT PRIMARY KEY,
    part_number TEXT,
    part_name TEXT,
    description TEXT,
    unit_price DOUBLE PRECISION,
    lead_time_days INTEGER,
    preferred_supplier TEXT,
    stock_on_hand DOUBLE PRECISION,
    manufacturer TEXT,
    category TEXT,
    notes TEXT,
    created_on TEXT,
    updated_on TEXT
);

CREATE TABLE IF NOT EXISTS modules (
    module_code TEXT PRIMARY KEY,
    quote_ref TEXT,
    module_name TEXT,
    description TEXT,
    instruction_text TEXT,
    estimated_hours DOUBLE PRECISION,
    stock_on_hand DOUBLE PRECISION,
    status TEXT,
    created_on TEXT,
    updated_on TEXT
);

CREATE TABLE IF NOT EXISTS tasks (
    task_id TEXT PRIMARY KEY,
    owner_type TEXT,
    owner_code TEXT,
    task_name TEXT,
    department TEXT,
    estimated_hours DOUBLE PRECISION,
    parent_task_id TEXT,
    dependency_task_id TEXT,
    stage TEXT,
    status TEXT,
    notes TEXT,
    created_on TEXT,
    updated_on TEXT
);

CREATE TABLE IF NOT EXISTS components (
    component_id TEXT PRIMARY KEY,
    owner_type TEXT,
    owner_code TEXT,
    component_name TEXT,
    qty DOUBLE PRECISION,
    soh_qty DOUBLE PRECISION,
    preferred_supplier TEXT,
    lead_time_days INTEGER,
    unit_price DOUBLE PRECISION,
    part_number TEXT,
    notes TEXT,
    created_on TEXT,
    updated_on TEXT
);

CREATE TABLE IF NOT EXISTS documents (
    doc_id TEXT PRIMARY KEY,
    owner_type TEXT,
    owner_code TEXT,
    section_name TEXT,
    doc_name TEXT,
    doc_type TEXT,
    file_path TEXT,
    instruction_text TEXT,
    added_on TEXT,
    updated_on TEXT
);

CREATE TABLE IF NOT EXISTS products (
    product_code TEXT PRIMARY KEY,
    quote_ref TEXT,
    product_name TEXT,
    description TEXT,
    revision TEXT,
    status TEXT,
    created_on TEXT,
    updated_on TEXT
);

CREATE TABLE IF NOT EXISTS product_modules (
    link_id TEXT PRIMARY KEY,
    product_code TEXT,
    module_code TEXT,
    module_order INTEGER,
    module_qty DOUBLE PRECISION,
    dependency_module_code TEXT,
    notes TEXT,
    created_on TEXT,
    updated_on TEXT
);

CREATE TABLE IF NOT EXISTS product_documents (
    prod_doc_id TEXT PRIMARY KEY,
    product_code TEXT,
    section_name TEXT,
    doc_name TEXT,
    doc_type TEXT,
    file_path TEXT,
    instruction_text TEXT,
    added_on TEXT,
    updated_on TEXT
);

CREATE TABLE IF NOT EXISTS projects (
    project_code TEXT PRIMARY KEY,
    quote_ref TEXT,
    project_name TEXT,
    client_name TEXT,
    location TEXT,
    description TEXT,
    linked_product_code TEXT,
    status TEXT,
    start_date TEXT,
    due_date TEXT,
    created_on TEXT,
    updated_on TEXT
);

CREATE TABLE IF NOT EXISTS project_modules (
    link_id TEXT PRIMARY KEY,
    project_code TEXT,
    module_code TEXT,
    source_type TEXT,
    source_code TEXT,
    module_order INTEGER,
    module_qty DOUBLE PRECISION,
    stage TEXT,
    status TEXT,
    dependency_module_code TEXT,
    notes TEXT,
    created_on TEXT,
    updated_on TEXT
);

CREATE TABLE IF NOT EXISTS project_tasks (
    project_task_id TEXT PRIMARY KEY,
    project_code TEXT,
    module_code TEXT,
    source_task_id TEXT,
    parent_project_task_id TEXT,
    task_name TEXT,
    department TEXT,
    estimated_hours DOUBLE PRECISION,
    stage TEXT,
    status TEXT,
    dependency_task_id TEXT,
    assigned_to TEXT,
    notes TEXT,
    created_on TEXT,
    updated_on TEXT
);

CREATE TABLE IF NOT EXISTS project_documents (
    project_doc_id TEXT PRIMARY KEY,
    project_code TEXT,
    section_name TEXT,
    doc_name TEXT,
    doc_type TEXT,
    file_path TEXT,
    instruction_text TEXT,
    added_on TEXT,
    updated_on TEXT
);

CREATE TABLE IF NOT EXISTS workorders (
    workorder_id TEXT PRIMARY KEY,
    owner_type TEXT,
    owner_code TEXT,
    workorder_name TEXT,
    stage TEXT,
    owner TEXT,
    due_date TEXT,
    status TEXT,
    notes TEXT,
    created_on TEXT,
    updated_on TEXT
);

CREATE TABLE IF NOT EXISTS completed_jobs (
    snapshot_id TEXT PRIMARY KEY,
    project_code TEXT,
    quote_ref TEXT,
    product_code TEXT,
    product_name TEXT,
    client_name TEXT,
    completed_on TEXT,
    labour_hours DOUBLE PRECISION,
    parts_total DOUBLE PRECISION,
    grand_total DOUBLE PRECISION,
    notes TEXT
);

CREATE TABLE IF NOT EXISTS completed_job_lines (
    snapshot_line_id TEXT PRIMARY KEY,
    snapshot_id TEXT,
    line_type TEXT,
    code TEXT,
    description TEXT,
    part_number TEXT,
    qty DOUBLE PRECISION,
    hours DOUBLE PRECISION,
    unit_price DOUBLE PRECISION,
    line_total DOUBLE PRECISION,
    source TEXT
);

CREATE INDEX IF NOT EXISTS idx_tasks_owner_code ON tasks (owner_code);
CREATE INDEX IF NOT EXISTS idx_components_owner_code ON components (owner_code);
CREATE INDEX IF NOT EXISTS idx_documents_owner_code ON documents (owner_code);
CREATE INDEX IF NOT EXISTS idx_product_modules_product_code ON product_modules (product_code);
CREATE INDEX IF NOT EXISTS idx_product_modules_module_code ON product_modules (module_code);
CREATE INDEX IF NOT EXISTS idx_product_documents_product_code ON product_documents (product_code);
CREATE INDEX IF NOT EXISTS idx_projects_linked_product_code ON projects (linked_product_code);
CREATE INDEX IF NOT EXISTS idx_project_modules_project_code ON project_modules (project_code);
CREATE INDEX IF NOT EXISTS idx_project_modules_module_code ON project_modules (module_code);
CREATE INDEX IF NOT EXISTS idx_project_tasks_project_code ON project_tasks (project_code);
CREATE INDEX IF NOT EXISTS idx_project_tasks_module_code ON project_tasks (module_code);
CREATE INDEX IF NOT EXISTS idx_project_documents_project_code ON project_documents (project_code);
CREATE INDEX IF NOT EXISTS idx_workorders_owner_code ON workorders (owner_code);
CREATE INDEX IF NOT EXISTS idx_completed_jobs_project_code ON completed_jobs (project_code);
CREATE INDEX IF NOT EXISTS idx_completed_job_lines_snapshot_id ON completed_job_lines (snapshot_id);
