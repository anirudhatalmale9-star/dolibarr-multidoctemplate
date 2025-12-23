-- Copyright (C) 2024
-- Keys for llx_multidoctemplate_archive

ALTER TABLE llx_multidoctemplate_archive ADD INDEX idx_multidoctemplate_archive_ref (ref);
ALTER TABLE llx_multidoctemplate_archive ADD INDEX idx_multidoctemplate_archive_fk_template (fk_template);
ALTER TABLE llx_multidoctemplate_archive ADD INDEX idx_multidoctemplate_archive_object (object_type, object_id);
ALTER TABLE llx_multidoctemplate_archive ADD INDEX idx_multidoctemplate_archive_fk_category (fk_category);
ALTER TABLE llx_multidoctemplate_archive ADD INDEX idx_multidoctemplate_archive_entity (entity);
-- Note: No foreign key constraint on fk_template to allow direct uploads (fk_template = NULL)
