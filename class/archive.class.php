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
 * Class MultiDocArchive
 * Manages generated archive documents
 */
class MultiDocArchive extends CommonObject
{
    public $element = 'multidoctemplate_archive';
    public $table_element = 'multidoctemplate_archive';
    public $picto = 'file';

    public $id;
    public $ref;
    public $fk_template;
    public $object_type;
    public $object_id;
    public $filename;
    public $filepath;
    public $filetype;
    public $filesize;
    public $fk_category;
    public $tag_filter;
    public $date_generation;
    public $fk_user_creat;
    public $entity;

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
     * Create archive record in database
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

        // Check parameters
        if (empty($this->ref)) {
            $this->error = 'ErrorRefRequired';
            return -1;
        }
        if (empty($this->fk_template)) {
            $this->error = 'ErrorTemplateRequired';
            return -2;
        }

        $this->db->begin();

        $sql = "INSERT INTO ".MAIN_DB_PREFIX.$this->table_element." (";
        $sql .= "ref, fk_template, object_type, object_id, filename, filepath, filetype, filesize,";
        $sql .= "fk_category, tag_filter, date_generation, fk_user_creat, entity";
        $sql .= ") VALUES (";
        $sql .= "'".$this->db->escape($this->ref)."',";
        $sql .= (int) $this->fk_template.",";
        $sql .= "'".$this->db->escape($this->object_type)."',";
        $sql .= (int) $this->object_id.",";
        $sql .= "'".$this->db->escape($this->filename)."',";
        $sql .= "'".$this->db->escape($this->filepath)."',";
        $sql .= "'".$this->db->escape($this->filetype)."',";
        $sql .= (int) $this->filesize.",";
        $sql .= ($this->fk_category > 0 ? (int) $this->fk_category : "NULL").",";
        $sql .= (!empty($this->tag_filter) ? "'".$this->db->escape($this->tag_filter)."'" : "NULL").",";
        $sql .= "'".$this->db->idate($now)."',";
        $sql .= (int) $user->id.",";
        $sql .= (int) $conf->entity;
        $sql .= ")";

        dol_syslog(get_class($this)."::create", LOG_DEBUG);
        $resql = $this->db->query($sql);

        if ($resql) {
            $this->id = $this->db->last_insert_id(MAIN_DB_PREFIX.$this->table_element);
            $this->date_generation = $now;
            $this->fk_user_creat = $user->id;

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
     * Load archive from database
     *
     * @param int $id Id of archive to fetch
     * @param string $ref Ref of archive to fetch
     * @return int <0 if KO, 0 if not found, >0 if OK
     */
    public function fetch($id, $ref = '')
    {
        global $conf;

        $sql = "SELECT a.rowid, a.ref, a.fk_template, a.object_type, a.object_id,";
        $sql .= " a.filename, a.filepath, a.filetype, a.filesize,";
        $sql .= " a.fk_category, a.tag_filter, a.date_generation, a.fk_user_creat, a.entity";
        $sql .= " FROM ".MAIN_DB_PREFIX.$this->table_element." as a";
        $sql .= " WHERE a.entity IN (".getEntity($this->element).")";
        if ($id) {
            $sql .= " AND a.rowid = ".(int) $id;
        } elseif ($ref) {
            $sql .= " AND a.ref = '".$this->db->escape($ref)."'";
        }

        dol_syslog(get_class($this)."::fetch", LOG_DEBUG);
        $resql = $this->db->query($sql);

        if ($resql) {
            if ($this->db->num_rows($resql)) {
                $obj = $this->db->fetch_object($resql);

                $this->id = $obj->rowid;
                $this->ref = $obj->ref;
                $this->fk_template = $obj->fk_template;
                $this->object_type = $obj->object_type;
                $this->object_id = $obj->object_id;
                $this->filename = $obj->filename;
                $this->filepath = $obj->filepath;
                $this->filetype = $obj->filetype;
                $this->filesize = $obj->filesize;
                $this->fk_category = $obj->fk_category;
                $this->tag_filter = $obj->tag_filter;
                $this->date_generation = $this->db->jdate($obj->date_generation);
                $this->fk_user_creat = $obj->fk_user_creat;
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
     * Delete archive from database and disk
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
     * Get archives for an object (thirdparty or contact)
     *
     * @param string $object_type Object type (thirdparty, contact)
     * @param int $object_id Object ID
     * @param int $fk_category Category filter (optional)
     * @return array|int Array of archives or -1 on error
     */
    public function fetchAllByObject($object_type, $object_id, $fk_category = 0)
    {
        global $conf;

        $archives = array();

        $sql = "SELECT a.rowid, a.ref, a.fk_template, a.object_type, a.object_id,";
        $sql .= " a.filename, a.filepath, a.filetype, a.filesize,";
        $sql .= " a.fk_category, a.tag_filter, a.date_generation, a.fk_user_creat,";
        $sql .= " t.label as template_label, t.tag as template_tag, t.fk_usergroup";
        $sql .= " FROM ".MAIN_DB_PREFIX.$this->table_element." as a";
        $sql .= " LEFT JOIN ".MAIN_DB_PREFIX."multidoctemplate_template as t ON t.rowid = a.fk_template";
        $sql .= " WHERE a.entity IN (".getEntity($this->element).")";
        $sql .= " AND a.object_type = '".$this->db->escape($object_type)."'";
        $sql .= " AND a.object_id = ".(int) $object_id;
        if ($fk_category > 0) {
            $sql .= " AND a.fk_category = ".(int) $fk_category;
        }
        $sql .= " ORDER BY t.tag ASC, a.date_generation DESC";

        dol_syslog(get_class($this)."::fetchAllByObject", LOG_DEBUG);
        $resql = $this->db->query($sql);

        if ($resql) {
            while ($obj = $this->db->fetch_object($resql)) {
                $archive = new self($this->db);
                $archive->id = $obj->rowid;
                $archive->ref = $obj->ref;
                $archive->fk_template = $obj->fk_template;
                $archive->object_type = $obj->object_type;
                $archive->object_id = $obj->object_id;
                $archive->filename = $obj->filename;
                $archive->filepath = $obj->filepath;
                $archive->filetype = $obj->filetype;
                $archive->filesize = $obj->filesize;
                $archive->fk_category = $obj->fk_category;
                $archive->tag_filter = $obj->tag_filter;
                $archive->date_generation = $this->db->jdate($obj->date_generation);
                $archive->fk_user_creat = $obj->fk_user_creat;
                $archive->template_label = $obj->template_label;
                $archive->template_tag = $obj->template_tag;
                $archive->fk_usergroup = $obj->fk_usergroup;
                $archives[$obj->rowid] = $archive;
            }
            $this->db->free($resql);
            return $archives;
        } else {
            $this->error = $this->db->lasterror();
            return -1;
        }
    }

    /**
     * Get archive storage directory for an object
     *
     * @param string $object_type Object type
     * @param int $object_id Object ID
     * @param string $tag_filter Tag/category filter (optional)
     * @return string Directory path
     */
    public static function getArchiveDir($object_type, $object_id, $tag_filter = '')
    {
        global $conf;

        $basedir = DOL_DATA_ROOT.'/multidoctemplate/archives/'.$object_type.'_'.(int) $object_id;

        if (!empty($tag_filter)) {
            $basedir .= '/'.dol_sanitizeFileName($tag_filter);
        }

        return $basedir;
    }

    /**
     * Generate unique reference for archive
     *
     * @param string $object_type Object type
     * @param int $object_id Object ID
     * @return string Reference
     */
    public static function generateRef($object_type, $object_id)
    {
        return strtoupper(substr($object_type, 0, 3)).'_'.$object_id.'_'.date('YmdHis').'_'.substr(md5(uniqid()), 0, 4);
    }
}
