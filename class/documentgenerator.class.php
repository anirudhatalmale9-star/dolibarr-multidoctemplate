<?php
/* Copyright (C) 2024
 *
 * This program is free software: you can redistribute it and/or modify
 * it under the terms of the GNU General Public License as published by
 * the Free Software Foundation, either version 3 of the License, or
 * (at your option) any later version.
 */

require_once DOL_DOCUMENT_ROOT.'/core/lib/files.lib.php';
require_once DOL_DOCUMENT_ROOT.'/core/lib/functions2.lib.php';

/**
 * Class MultiDocGenerator
 * Generates documents from templates using Dolibarr's variable substitution
 */
class MultiDocGenerator
{
    public $db;
    public $error = '';
    public $errors = array();

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
     * Generate document from template
     *
     * @param MultiDocTemplate $template Template object
     * @param CommonObject $object Object to use for substitution (Societe or Contact)
     * @param string $object_type Object type (thirdparty or contact)
     * @param User $user User generating the document
     * @param string $tag_filter Tag/category filter for folder organization
     * @return int >0 if OK, <0 if KO
     */
    public function generate($template, $object, $object_type, $user, $tag_filter = '')
    {
        global $conf, $langs, $mysoc;

        $error = 0;

        // Check template file exists
        if (!file_exists($template->filepath)) {
            $this->error = $langs->trans('ErrorTemplateFileNotFound');
            return -1;
        }

        // Create archive directory
        $archive_dir = MultiDocArchive::getArchiveDir($object_type, $object->id, $tag_filter);
        if (!is_dir($archive_dir)) {
            if (dol_mkdir($archive_dir) < 0) {
                $this->error = $langs->trans('ErrorCanNotCreateDir', $archive_dir);
                return -2;
            }
        }

        // Generate output filename
        $ext = pathinfo($template->filename, PATHINFO_EXTENSION);
        $output_filename = $this->generateOutputFilename($template, $object, $object_type, $ext);
        $output_filepath = $archive_dir.'/'.$output_filename;

        // Process based on file type
        $result = 0;
        switch (strtolower($ext)) {
            case 'odt':
                $result = $this->processODT($template->filepath, $output_filepath, $object, $object_type);
                break;
            case 'ods':
                $result = $this->processODS($template->filepath, $output_filepath, $object, $object_type);
                break;
            default:
                // For other formats, just copy with basic text substitution if possible
                $result = $this->processCopy($template->filepath, $output_filepath, $object, $object_type);
                break;
        }

        if ($result < 0) {
            return $result;
        }

        // Create archive record
        require_once __DIR__.'/archive.class.php';
        $archive = new MultiDocArchive($this->db);
        $archive->ref = MultiDocArchive::generateRef($object_type, $object->id);
        $archive->fk_template = $template->id;
        $archive->object_type = $object_type;
        $archive->object_id = $object->id;
        $archive->filename = $output_filename;
        $archive->filepath = $output_filepath;
        $archive->filetype = strtolower($ext);
        $archive->filesize = filesize($output_filepath);
        $archive->tag_filter = $tag_filter;

        $result = $archive->create($user);

        if ($result < 0) {
            // Delete generated file if DB insert failed
            dol_delete_file($output_filepath);
            $this->error = $archive->error;
            return -3;
        }

        return $archive->id;
    }

    /**
     * Process ODT template with variable substitution
     *
     * @param string $template_path Path to template
     * @param string $output_path Path for output file
     * @param CommonObject $object Object for substitution
     * @param string $object_type Object type
     * @return int >0 if OK, <0 if KO
     */
    protected function processODT($template_path, $output_path, $object, $object_type)
    {
        global $conf, $langs, $mysoc;

        // Check if ODT processing library exists
        if (!class_exists('Odf')) {
            // Try to include Dolibarr's ODFPhp library
            $odfphp_path = DOL_DOCUMENT_ROOT.'/includes/odtphp/odf.php';
            if (file_exists($odfphp_path)) {
                require_once $odfphp_path;
            } else {
                // Fallback: simple copy
                return $this->processCopy($template_path, $output_path, $object, $object_type);
            }
        }

        try {
            // Load ODT template
            $odfHandler = new Odf(
                $template_path,
                array(
                    'PATH_TO_TMP' => $conf->multidoctemplate->dir_temp,
                    'ZIP_PROXY' => 'PclZipProxy',
                    'DELIMITER_LEFT' => '{',
                    'DELIMITER_RIGHT' => '}'
                )
            );

            // Get substitution array
            $substitutions = $this->getSubstitutionArray($object, $object_type);

            // Apply substitutions
            foreach ($substitutions as $key => $value) {
                try {
                    $odfHandler->setVars($key, $value, true, 'UTF-8');
                } catch (Exception $e) {
                    // Variable not found in template, continue
                }
            }

            // Save output file
            $odfHandler->saveToDisk($output_path);

            return 1;
        } catch (Exception $e) {
            $this->error = $e->getMessage();
            return -10;
        }
    }

    /**
     * Process ODS template with variable substitution
     *
     * @param string $template_path Path to template
     * @param string $output_path Path for output file
     * @param CommonObject $object Object for substitution
     * @param string $object_type Object type
     * @return int >0 if OK, <0 if KO
     */
    protected function processODS($template_path, $output_path, $object, $object_type)
    {
        // For ODS, we use similar approach but with spreadsheet handling
        // Fallback to copy for now
        return $this->processCopy($template_path, $output_path, $object, $object_type);
    }

    /**
     * Simple copy with basic text substitution (for formats we can't process)
     *
     * @param string $template_path Path to template
     * @param string $output_path Path for output file
     * @param CommonObject $object Object for substitution
     * @param string $object_type Object type
     * @return int >0 if OK, <0 if KO
     */
    protected function processCopy($template_path, $output_path, $object, $object_type)
    {
        $result = dol_copy($template_path, $output_path, 0, 1);
        return $result ? 1 : -20;
    }

    /**
     * Generate output filename
     *
     * @param MultiDocTemplate $template Template
     * @param CommonObject $object Object
     * @param string $object_type Object type
     * @param string $ext File extension
     * @return string Filename
     */
    protected function generateOutputFilename($template, $object, $object_type, $ext)
    {
        $name_parts = array();

        // Add object ref/name
        if ($object_type == 'thirdparty' && !empty($object->name)) {
            $name_parts[] = dol_sanitizeFileName($object->name);
        } elseif (!empty($object->lastname)) {
            $name_parts[] = dol_sanitizeFileName($object->lastname);
            if (!empty($object->firstname)) {
                $name_parts[] = dol_sanitizeFileName($object->firstname);
            }
        }

        // Add template label
        $name_parts[] = dol_sanitizeFileName($template->label);

        // Add timestamp
        $name_parts[] = date('Ymd_His');

        return implode('_', $name_parts).'.'.$ext;
    }

    /**
     * Get substitution array for object
     * Uses Dolibarr's native substitution system
     *
     * @param CommonObject $object Object
     * @param string $object_type Object type
     * @return array Substitution array
     */
    public function getSubstitutionArray($object, $object_type)
    {
        global $conf, $langs, $mysoc, $user;

        $substitutions = array();

        // Load complete_substitutions_array function if available
        if (function_exists('complete_substitutions_array')) {
            // Use Dolibarr's native substitution system
            $tmparray = array();

            // Add object-specific substitutions
            if ($object_type == 'thirdparty') {
                complete_substitutions_array($tmparray, $langs, $object, null, 0, 'societe');
            } else {
                complete_substitutions_array($tmparray, $langs, $object, null, 0, 'contact');
            }

            // Convert to simple key => value
            foreach ($tmparray as $key => $val) {
                if (is_string($val)) {
                    // Remove __xxx__ format
                    $clean_key = preg_replace('/^__|__$/', '', $key);
                    $substitutions[$clean_key] = $val;
                    $substitutions[$key] = $val; // Keep original format too
                }
            }
        }

        // Add manual substitutions for common fields
        if ($object_type == 'thirdparty') {
            $substitutions = array_merge($substitutions, $this->getThirdpartySubstitutions($object));
        } else {
            $substitutions = array_merge($substitutions, $this->getContactSubstitutions($object));
        }

        // Add date/time substitutions
        $substitutions['DATE_NOW'] = dol_print_date(dol_now(), 'day');
        $substitutions['DATETIME_NOW'] = dol_print_date(dol_now(), 'dayhour');
        $substitutions['YEAR'] = date('Y');
        $substitutions['MONTH'] = date('m');
        $substitutions['DAY'] = date('d');

        // Add company substitutions
        if (is_object($mysoc)) {
            $substitutions['MYCOMPANY_NAME'] = $mysoc->name;
            $substitutions['MYCOMPANY_ADDRESS'] = $mysoc->address;
            $substitutions['MYCOMPANY_ZIP'] = $mysoc->zip;
            $substitutions['MYCOMPANY_TOWN'] = $mysoc->town;
            $substitutions['MYCOMPANY_COUNTRY'] = $mysoc->country;
            $substitutions['MYCOMPANY_EMAIL'] = $mysoc->email;
            $substitutions['MYCOMPANY_PHONE'] = $mysoc->phone;
            $substitutions['MYCOMPANY_FAX'] = $mysoc->fax;
            $substitutions['MYCOMPANY_WEB'] = $mysoc->url;
            $substitutions['MYCOMPANY_SIRET'] = $mysoc->idprof1;
            $substitutions['MYCOMPANY_SIREN'] = $mysoc->idprof2;
            $substitutions['MYCOMPANY_VAT_INTRA'] = $mysoc->tva_intra;
        }

        // Add user substitutions
        $substitutions['USER_LOGIN'] = $user->login;
        $substitutions['USER_NAME'] = $user->getFullName($langs);
        $substitutions['USER_EMAIL'] = $user->email;

        return $substitutions;
    }

    /**
     * Get thirdparty-specific substitutions
     *
     * @param Societe $object Thirdparty object
     * @return array Substitutions
     */
    protected function getThirdpartySubstitutions($object)
    {
        global $langs;

        $subs = array();

        // Basic info
        $subs['THIRDPARTY_ID'] = $object->id;
        $subs['THIRDPARTY_NAME'] = $object->name;
        $subs['THIRDPARTY_NAME_ALIAS'] = $object->name_alias;
        $subs['THIRDPARTY_ADDRESS'] = $object->address;
        $subs['THIRDPARTY_ZIP'] = $object->zip;
        $subs['THIRDPARTY_TOWN'] = $object->town;
        $subs['THIRDPARTY_STATE'] = $object->state;
        $subs['THIRDPARTY_COUNTRY'] = $object->country;
        $subs['THIRDPARTY_COUNTRY_CODE'] = $object->country_code;

        // Contact info
        $subs['THIRDPARTY_PHONE'] = $object->phone;
        $subs['THIRDPARTY_FAX'] = $object->fax;
        $subs['THIRDPARTY_EMAIL'] = $object->email;
        $subs['THIRDPARTY_WEB'] = $object->url;
        $subs['THIRDPARTY_SKYPE'] = $object->skype;

        // Legal info
        $subs['THIRDPARTY_IDPROF1'] = $object->idprof1;
        $subs['THIRDPARTY_IDPROF2'] = $object->idprof2;
        $subs['THIRDPARTY_IDPROF3'] = $object->idprof3;
        $subs['THIRDPARTY_IDPROF4'] = $object->idprof4;
        $subs['THIRDPARTY_VAT_INTRA'] = $object->tva_intra;
        $subs['THIRDPARTY_CAPITAL'] = $object->capital;

        // Codes
        $subs['THIRDPARTY_CODE_CLIENT'] = $object->code_client;
        $subs['THIRDPARTY_CODE_FOURNISSEUR'] = $object->code_fournisseur;
        $subs['THIRDPARTY_CODE_COMPTA'] = $object->code_compta;

        // Status
        $subs['THIRDPARTY_STATUS'] = $object->status;

        // Notes
        $subs['THIRDPARTY_NOTE_PUBLIC'] = $object->note_public;
        $subs['THIRDPARTY_NOTE_PRIVATE'] = $object->note_private;

        // Also add with simpler names for compatibility
        $subs['COMPANY_NAME'] = $object->name;
        $subs['COMPANY_ADDRESS'] = $object->address;
        $subs['COMPANY_ZIP'] = $object->zip;
        $subs['COMPANY_TOWN'] = $object->town;
        $subs['COMPANY_COUNTRY'] = $object->country;
        $subs['COMPANY_EMAIL'] = $object->email;
        $subs['COMPANY_PHONE'] = $object->phone;

        return $subs;
    }

    /**
     * Get contact-specific substitutions
     *
     * @param Contact $object Contact object
     * @return array Substitutions
     */
    protected function getContactSubstitutions($object)
    {
        global $langs;

        $subs = array();

        // Basic info
        $subs['CONTACT_ID'] = $object->id;
        $subs['CONTACT_CIVILITY'] = $object->civility;
        $subs['CONTACT_FIRSTNAME'] = $object->firstname;
        $subs['CONTACT_LASTNAME'] = $object->lastname;
        $subs['CONTACT_FULLNAME'] = $object->getFullName($langs);
        $subs['CONTACT_POSTE'] = $object->poste;

        // Address
        $subs['CONTACT_ADDRESS'] = $object->address;
        $subs['CONTACT_ZIP'] = $object->zip;
        $subs['CONTACT_TOWN'] = $object->town;
        $subs['CONTACT_STATE'] = $object->state;
        $subs['CONTACT_COUNTRY'] = $object->country;
        $subs['CONTACT_COUNTRY_CODE'] = $object->country_code;

        // Contact info
        $subs['CONTACT_PHONE_PRO'] = $object->phone_pro;
        $subs['CONTACT_PHONE_PERSO'] = $object->phone_perso;
        $subs['CONTACT_PHONE_MOBILE'] = $object->phone_mobile;
        $subs['CONTACT_FAX'] = $object->fax;
        $subs['CONTACT_EMAIL'] = $object->email;
        $subs['CONTACT_SKYPE'] = $object->skype;

        // Notes
        $subs['CONTACT_NOTE_PUBLIC'] = $object->note_public;
        $subs['CONTACT_NOTE_PRIVATE'] = $object->note_private;

        // Birthday
        if (!empty($object->birthday)) {
            $subs['CONTACT_BIRTHDAY'] = dol_print_date($object->birthday, 'day');
        }

        return $subs;
    }
}
