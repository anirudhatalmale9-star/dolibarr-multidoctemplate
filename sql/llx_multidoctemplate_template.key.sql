-- Copyright (C) 2024
-- Keys for llx_multidoctemplate_template

ALTER TABLE llx_multidoctemplate_template ADD INDEX idx_multidoctemplate_template_ref (ref);
ALTER TABLE llx_multidoctemplate_template ADD INDEX idx_multidoctemplate_template_fk_usergroup (fk_usergroup);
ALTER TABLE llx_multidoctemplate_template ADD INDEX idx_multidoctemplate_template_active (active);
ALTER TABLE llx_multidoctemplate_template ADD INDEX idx_multidoctemplate_template_entity (entity);
ALTER TABLE llx_multidoctemplate_template ADD CONSTRAINT fk_multidoctemplate_template_usergroup FOREIGN KEY (fk_usergroup) REFERENCES llx_usergroup(rowid);
