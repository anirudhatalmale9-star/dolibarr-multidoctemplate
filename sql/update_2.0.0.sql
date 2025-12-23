-- MultiDocTemplate update to version 2.0.0
-- Adds native Dolibarr category support and fixes direct upload

-- Add fk_category column to templates table
ALTER TABLE llx_multidoctemplate_template ADD COLUMN fk_category INTEGER DEFAULT NULL AFTER tag;
ALTER TABLE llx_multidoctemplate_template ADD INDEX idx_multidoctemplate_template_fk_category (fk_category);

-- Add fk_category column to archives table (if not exists)
ALTER TABLE llx_multidoctemplate_archive ADD COLUMN fk_category INTEGER DEFAULT NULL AFTER filesize;
ALTER TABLE llx_multidoctemplate_archive ADD INDEX idx_multidoctemplate_archive_fk_category (fk_category);

-- Drop foreign key constraint to allow direct uploads (fk_template = NULL)
-- Note: Constraint name may vary, try both possible names
ALTER TABLE llx_multidoctemplate_archive DROP FOREIGN KEY fk_multidoctemplate_archive_template;

-- Allow NULL for fk_template (for direct uploads without template)
ALTER TABLE llx_multidoctemplate_archive MODIFY fk_template INTEGER DEFAULT NULL;

-- Update existing records with fk_template=0 to NULL
UPDATE llx_multidoctemplate_archive SET fk_template = NULL WHERE fk_template = 0;
