<?php
/* Copyright (C) 2024
 *
 * This program is free software: you can redistribute it and/or modify
 * it under the terms of the GNU General Public License as published by
 * the Free Software Foundation, either version 3 of the License, or
 * (at your option) any later version.
 */

require_once DOL_DOCUMENT_ROOT.'/core/class/commonobject.class.php';

/**
 * Class Template
 * Manages document templates for user groups
 */
class MultiDocTemplate extends CommonObject
{
    public $element = 'multidoctemplate_template';
    public $table_element = 'multidoctemplate_template';
    public $picto = 'file';

    public $id;
    public $ref;
    public $label;
    public $description;
    public $tag;
    public $fk_usergroup;
    public $filename;
    public $filepath;
    public $filetype;
    public $filesize;
    public $mime_type;
    public $active;
    public $date_creation;
    public $date_modification;
    public $fk_user_creat;
    public $fk_user_modif;
    public $entity;

    // Allowed file extensions
    public static $allowed_extensions = array(
        'odt',   // Mandatory
        'ods',   // Spreadsheet
        'xls',
        'xlsx',
        'doc',
        'docx',
        'pdf',
        'rtf'
    );

    /**
     * Constructor
     *
     * @param DoliDB $db Database handler
     */
    public function __construct($db)
    {
        $this->db = $db;
    }

    /**
     * Create template in database
     *
     * @param User $user User that creates
     * @param int $notrigger 0=launch triggers, 1=disable triggers
     * @return int <0 if KO, Id of created object if OK
     */
    public function create($user, $notrigger = 0)
    {
        global $conf;

        $error = 0;
        $now = dol_now();

        // Clean parameters
        $this->ref = dol_sanitizeFileName($this->ref);
        $this->label = trim($this->label);
        $this->fk_usergroup = (int) $this->fk_usergroup;

        // Check parameters
        if (empty($this->ref)) {
            $this->error = 'ErrorRefRequired';
            return -1;
        }
        if (empty($this->fk_usergroup)) {
            $this->error = 'ErrorUserGroupRequired';
            return -2;
        }

        $this->db->begin();

        $sql = "INSERT INTO ".MAIN_DB_PREFIX.$this->table_element." (";
        $sql .= "ref, label, description, tag, fk_usergroup, filename, filepath, filetype, filesize, mime_type,";
        $sql .= "active, date_creation, fk_user_creat, entity";
        $sql .= ") VALUES (";
        $sql .= "'".$this->db->escape($this->ref)."',";
        $sql .= "'".$this->db->escape($this->label)."',";
        $sql .= "'".$this->db->escape($this->description)."',";
        $sql .= "'".$this->db->escape($this->tag)."',";
        $sql .= $this->fk_usergroup.",";
        $sql .= "'".$this->db->escape($this->filename)."',";
        $sql .= "'".$this->db->escape($this->filepath)."',";
        $sql .= "'".$this->db->escape($this->filetype)."',";
        $sql .= (int) $this->filesize.",";
        $sql .= "'".$this->db->escape($this->mime_type)."',";
        $sql .= "1,";
        $sql .= "'".$this->db->idate($now)."',";
        $sql .= (int) $user->id.",";
        $sql .= (int) $conf->entity;
        $sql .= ")";

        dol_syslog(get_class($this)."::create", LOG_DEBUG);
        $resql = $this->db->query($sql);

        if ($resql) {
            $this->id = $this->db->last_insert_id(MAIN_DB_PREFIX.$this->table_element);
            $this->date_creation = $now;
            $this->fk_user_creat = $user->id;

            if (!$notrigger) {
                // Uncomment to enable triggers
                // $result = $this->call_trigger('MULTIDOCTEMPLATE_CREATE', $user);
                // if ($result < 0) $error++;
            }

            if (!$error) {
                $this->db->commit();
                return $this->id;
            } else {
                $this->db->rollback();
                return -1;
            }
        } else {
            $this->error = $this->db->lasterror();
            $this->db->rollback();
            return -1;
        }
    }

    /**
     * Load template from database
     *
     * @param int $id Id of template to fetch
     * @param string $ref Ref of template to fetch
     * @return int <0 if KO, 0 if not found, >0 if OK
     */
    public function fetch($id, $ref = '')
    {
        global $conf;

        $sql = "SELECT t.rowid, t.ref, t.label, t.description, t.tag, t.fk_usergroup,";
        $sql .= " t.filename, t.filepath, t.filetype, t.filesize, t.mime_type,";
        $sql .= " t.active, t.date_creation, t.date_modification,";
        $sql .= " t.fk_user_creat, t.fk_user_modif, t.entity";
        $sql .= " FROM ".MAIN_DB_PREFIX.$this->table_element." as t";
        $sql .= " WHERE t.entity IN (".getEntity($this->element).")";
        if ($id) {
            $sql .= " AND t.rowid = ".(int) $id;
        } elseif ($ref) {
            $sql .= " AND t.ref = '".$this->db->escape($ref)."'";
        }

        dol_syslog(get_class($this)."::fetch", LOG_DEBUG);
        $resql = $this->db->query($sql);

        if ($resql) {
            if ($this->db->num_rows($resql)) {
                $obj = $this->db->fetch_object($resql);

                $this->id = $obj->rowid;
                $this->ref = $obj->ref;
                $this->label = $obj->label;
                $this->description = $obj->description;
                $this->tag = $obj->tag;
                $this->fk_usergroup = $obj->fk_usergroup;
                $this->filename = $obj->filename;
                $this->filepath = $obj->filepath;
                $this->filetype = $obj->filetype;
                $this->filesize = $obj->filesize;
                $this->mime_type = $obj->mime_type;
                $this->active = $obj->active;
                $this->date_creation = $this->db->jdate($obj->date_creation);
                $this->date_modification = $this->db->jdate($obj->date_modification);
                $this->fk_user_creat = $obj->fk_user_creat;
                $this->fk_user_modif = $obj->fk_user_modif;
                $this->entity = $obj->entity;

                $this->db->free($resql);
                return 1;
            } else {
                $this->db->free($resql);
                return 0;
            }
        } else {
            $this->error = $this->db->lasterror();
            return -1;
        }
    }

    /**
     * Delete template from database
     *
     * @param User $user User that deletes
     * @param int $notrigger 0=launch triggers, 1=disable triggers
     * @return int <0 if KO, >0 if OK
     */
    public function delete($user, $notrigger = 0)
    {
        $error = 0;

        $this->db->begin();

        // Delete file from disk
        if (!empty($this->filepath) && file_exists($this->filepath)) {
            dol_delete_file($this->filepath);
        }

        if (!$notrigger) {
            // Uncomment to enable triggers
            // $result = $this->call_trigger('MULTIDOCTEMPLATE_DELETE', $user);
            // if ($result < 0) $error++;
        }

        // First delete all archives that reference this template
        if (!$error) {
            $sql = "DELETE FROM ".MAIN_DB_PREFIX."multidoctemplate_archive";
            $sql .= " WHERE fk_template = ".(int) $this->id;

            dol_syslog(get_class($this)."::delete archives", LOG_DEBUG);
            $resql = $this->db->query($sql);

            if (!$resql) {
                $this->error = $this->db->lasterror();
                $error++;
            }
        }

        // Then delete the template
        if (!$error) {
            $sql = "DELETE FROM ".MAIN_DB_PREFIX.$this->table_element;
            $sql .= " WHERE rowid = ".(int) $this->id;

            dol_syslog(get_class($this)."::delete", LOG_DEBUG);
            $resql = $this->db->query($sql);

            if (!$resql) {
                $this->error = $this->db->lasterror();
                $error++;
            }
        }

        if (!$error) {
            $this->db->commit();
            return 1;
        } else {
            $this->db->rollback();
            return -1;
        }
    }

    /**
     * Get list of templates for a user group
     *
     * @param int $fk_usergroup User group ID
     * @param int $active Filter by active status (1=active only, 0=inactive only, -1=all)
     * @return array|int Array of templates or -1 on error
     */
    public function fetchAllByUserGroup($fk_usergroup, $active = 1)
    {
        global $conf;

        $templates = array();

        $sql = "SELECT t.rowid, t.ref, t.label, t.description, t.tag, t.fk_usergroup,";
        $sql .= " t.filename, t.filepath, t.filetype, t.filesize, t.mime_type,";
        $sql .= " t.active, t.date_creation";
        $sql .= " FROM ".MAIN_DB_PREFIX.$this->table_element." as t";
        $sql .= " WHERE t.entity IN (".getEntity($this->element).")";
        $sql .= " AND t.fk_usergroup = ".(int) $fk_usergroup;
        if ($active >= 0) {
            $sql .= " AND t.active = ".(int) $active;
        }
        $sql .= " ORDER BY t.tag ASC, t.label ASC";

        dol_syslog(get_class($this)."::fetchAllByUserGroup", LOG_DEBUG);
        $resql = $this->db->query($sql);

        if ($resql) {
            while ($obj = $this->db->fetch_object($resql)) {
                $template = new self($this->db);
                $template->id = $obj->rowid;
                $template->ref = $obj->ref;
                $template->label = $obj->label;
                $template->description = $obj->description;
                $template->tag = $obj->tag;
                $template->fk_usergroup = $obj->fk_usergroup;
                $template->filename = $obj->filename;
                $template->filepath = $obj->filepath;
                $template->filetype = $obj->filetype;
                $template->filesize = $obj->filesize;
                $template->mime_type = $obj->mime_type;
                $template->active = $obj->active;
                $template->date_creation = $this->db->jdate($obj->date_creation);
                $templates[$obj->rowid] = $template;
            }
            $this->db->free($resql);
            return $templates;
        } else {
            $this->error = $this->db->lasterror();
            return -1;
        }
    }

    /**
     * Get all templates accessible to user's groups
     *
     * @param User $user User object
     * @return array|int Array of templates or -1 on error
     */
    public function fetchAllForUser($user)
    {
        global $conf;

        $templates = array();

        // Get user's groups directly from database
        $sql_groups = "SELECT fk_usergroup FROM ".MAIN_DB_PREFIX."usergroup_user";
        $sql_groups .= " WHERE fk_user = ".(int) $user->id;
        $sql_groups .= " AND entity IN (".getEntity('usergroup').")";

        dol_syslog(get_class($this)."::fetchAllForUser get groups", LOG_DEBUG);
        $resql_groups = $this->db->query($sql_groups);

        $groupids = array();
        if ($resql_groups) {
            while ($obj_group = $this->db->fetch_object($resql_groups)) {
                $groupids[] = (int) $obj_group->fk_usergroup;
            }
            $this->db->free($resql_groups);
        }

        // If user has no groups, return empty (or return all templates if user is admin)
        if (empty($groupids)) {
            // If user is admin, show all templates
            if ($user->admin) {
                $sql_all_groups = "SELECT rowid FROM ".MAIN_DB_PREFIX."usergroup";
                $sql_all_groups .= " WHERE entity IN (".getEntity('usergroup').")";
                $resql_all = $this->db->query($sql_all_groups);
                if ($resql_all) {
                    while ($obj_grp = $this->db->fetch_object($resql_all)) {
                        $groupids[] = (int) $obj_grp->rowid;
                    }
                    $this->db->free($resql_all);
                }
            }
            if (empty($groupids)) {
                return $templates;
            }
        }

        $sql = "SELECT t.rowid, t.ref, t.label, t.description, t.tag, t.fk_usergroup,";
        $sql .= " t.filename, t.filepath, t.filetype, t.filesize, t.mime_type,";
        $sql .= " t.active, t.date_creation, ug.nom as usergroup_name";
        $sql .= " FROM ".MAIN_DB_PREFIX.$this->table_element." as t";
        $sql .= " LEFT JOIN ".MAIN_DB_PREFIX."usergroup as ug ON ug.rowid = t.fk_usergroup";
        $sql .= " WHERE t.entity IN (".getEntity($this->element).")";
        $sql .= " AND t.fk_usergroup IN (".implode(',', $groupids).")";
        $sql .= " AND t.active = 1";
        $sql .= " ORDER BY t.tag ASC, ug.nom ASC, t.label ASC";

        dol_syslog(get_class($this)."::fetchAllForUser", LOG_DEBUG);
        $resql = $this->db->query($sql);

        if ($resql) {
            while ($obj = $this->db->fetch_object($resql)) {
                $template = new self($this->db);
                $template->id = $obj->rowid;
                $template->ref = $obj->ref;
                $template->label = $obj->label;
                $template->description = $obj->description;
                $template->tag = $obj->tag;
                $template->fk_usergroup = $obj->fk_usergroup;
                $template->filename = $obj->filename;
                $template->filepath = $obj->filepath;
                $template->filetype = $obj->filetype;
                $template->filesize = $obj->filesize;
                $template->mime_type = $obj->mime_type;
                $template->active = $obj->active;
                $template->date_creation = $this->db->jdate($obj->date_creation);
                $template->usergroup_name = $obj->usergroup_name;
                $templates[$obj->rowid] = $template;
            }
            $this->db->free($resql);
            return $templates;
        } else {
            $this->error = $this->db->lasterror();
            return -1;
        }
    }

    /**
     * Get upload directory path for a user group
     *
     * @param int $fk_usergroup User group ID
     * @return string Directory path
     */
    public static function getUploadDir($fk_usergroup)
    {
        global $conf;
        return DOL_DATA_ROOT.'/multidoctemplate/templates/group_'.(int) $fk_usergroup;
    }

    /**
     * Check if file extension is allowed
     *
     * @param string $filename Filename to check
     * @return bool True if allowed
     */
    public static function isAllowedExtension($filename)
    {
        $ext = strtolower(pathinfo($filename, PATHINFO_EXTENSION));
        return in_array($ext, self::$allowed_extensions);
    }

    /**
     * Check if file is ODT (mandatory format)
     *
     * @param string $filename Filename to check
     * @return bool True if ODT
     */
    public static function isODT($filename)
    {
        $ext = strtolower(pathinfo($filename, PATHINFO_EXTENSION));
        return ($ext === 'odt');
    }

    /**
     * Get MIME type for file
     *
     * @param string $filename Filename
     * @return string MIME type
     */
    public static function getMimeType($filename)
    {
        $ext = strtolower(pathinfo($filename, PATHINFO_EXTENSION));
        $mimes = array(
            'odt' => 'application/vnd.oasis.opendocument.text',
            'ods' => 'application/vnd.oasis.opendocument.spreadsheet',
            'doc' => 'application/msword',
            'docx' => 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            'xls' => 'application/vnd.ms-excel',
            'xlsx' => 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            'pdf' => 'application/pdf',
            'rtf' => 'application/rtf'
        );
        return isset($mimes[$ext]) ? $mimes[$ext] : 'application/octet-stream';
    }
}
