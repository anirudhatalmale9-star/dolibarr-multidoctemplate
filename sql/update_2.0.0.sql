-- MultiDocTemplate update to version 2.0.0
-- Adds native Dolibarr category support

-- Add fk_category column to templates table
ALTER TABLE llx_multidoctemplate_template ADD COLUMN fk_category INTEGER DEFAULT NULL AFTER tag;
ALTER TABLE llx_multidoctemplate_template ADD INDEX idx_multidoctemplate_template_fk_category (fk_category);

-- Add fk_category column to archives table (if not exists)
ALTER TABLE llx_multidoctemplate_archive ADD COLUMN fk_category INTEGER DEFAULT NULL AFTER filesize;
ALTER TABLE llx_multidoctemplate_archive ADD INDEX idx_multidoctemplate_archive_fk_category (fk_category);
