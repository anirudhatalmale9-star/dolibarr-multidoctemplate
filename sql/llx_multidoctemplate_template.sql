-- Copyright (C) 2024
-- This program is free software: you can redistribute it and/or modify
-- it under the terms of the GNU General Public License as published by
-- the Free Software Foundation, either version 3 of the License, or
-- (at your option) any later version.

CREATE TABLE llx_multidoctemplate_template (
    rowid           INTEGER AUTO_INCREMENT PRIMARY KEY,
    ref             VARCHAR(128) NOT NULL,
    label           VARCHAR(255) NOT NULL,
    description     TEXT,
    tag             VARCHAR(128) NOT NULL DEFAULT '',
    fk_usergroup    INTEGER NOT NULL,
    filename        VARCHAR(255) NOT NULL,
    filepath        VARCHAR(512) NOT NULL,
    filetype        VARCHAR(32) NOT NULL,
    filesize        INTEGER DEFAULT 0,
    mime_type       VARCHAR(128),
    active          SMALLINT DEFAULT 1,
    date_creation   DATETIME NOT NULL,
    date_modification DATETIME,
    fk_user_creat   INTEGER,
    fk_user_modif   INTEGER,
    import_key      VARCHAR(14),
    entity          INTEGER DEFAULT 1
) ENGINE=InnoDB;
