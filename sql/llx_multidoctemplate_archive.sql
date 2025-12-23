-- Copyright (C) 2024
-- This program is free software: you can redistribute it and/or modify
-- it under the terms of the GNU General Public License as published by
-- the Free Software Foundation, either version 3 of the License, or
-- (at your option) any later version.

CREATE TABLE llx_multidoctemplate_archive (
    rowid               INTEGER AUTO_INCREMENT PRIMARY KEY,
    ref                 VARCHAR(128) NOT NULL,
    fk_template         INTEGER DEFAULT NULL,
    object_type         VARCHAR(64) NOT NULL,
    object_id           INTEGER NOT NULL,
    filename            VARCHAR(255) NOT NULL,
    filepath            VARCHAR(512) NOT NULL,
    filetype            VARCHAR(32) NOT NULL,
    filesize            INTEGER DEFAULT 0,
    fk_category         INTEGER DEFAULT NULL,
    tag_filter          VARCHAR(255) DEFAULT NULL,
    date_generation     DATETIME NOT NULL,
    fk_user_creat       INTEGER,
    import_key          VARCHAR(14),
    entity              INTEGER DEFAULT 1
) ENGINE=InnoDB;
