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

        // Use template's tag for folder organization if no tag_filter provided
        $folder_tag = !empty($tag_filter) ? $tag_filter : (!empty($template->tag) ? $template->tag : '');

        // Create archive directory
        $archive_dir = MultiDocArchive::getArchiveDir($object_type, $object->id, $folder_tag);
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
            case 'xlsx':
                $result = $this->processXLSX($template->filepath, $output_filepath, $object, $object_type);
                break;
            case 'docx':
                $result = $this->processDOCX($template->filepath, $output_filepath, $object, $object_type);
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
        $archive->tag_filter = $folder_tag;

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
            // Determine temp directory - use module's temp dir or fallback to main temp
            $temp_dir = '';
            if (!empty($conf->multidoctemplate->dir_temp)) {
                $temp_dir = $conf->multidoctemplate->dir_temp;
            } else {
                // Fallback to Dolibarr's main temp directory
                $temp_dir = DOL_DATA_ROOT.'/multidoctemplate/temp';
            }

            // Create temp directory if it doesn't exist
            if (!is_dir($temp_dir)) {
                if (dol_mkdir($temp_dir) < 0) {
                    $this->error = $langs->trans('ErrorCanNotCreateDir', $temp_dir);
                    return -11;
                }
            }

            // Load ODT template
            $odfHandler = new Odf(
                $template_path,
                array(
                    'PATH_TO_TMP' => $temp_dir,
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
        // ODS files are ZIP archives with XML content
        return $this->processZipXML($template_path, $output_path, $object, $object_type, 'content.xml');
    }

    /**
     * Process XLSX template with variable substitution
     *
     * @param string $template_path Path to template
     * @param string $output_path Path for output file
     * @param CommonObject $object Object for substitution
     * @param string $object_type Object type
     * @return int >0 if OK, <0 if KO
     */
    protected function processXLSX($template_path, $output_path, $object, $object_type)
    {
        // XLSX files are ZIP archives with XML content in xl/sharedStrings.xml and xl/worksheets/*.xml
        return $this->processZipXML($template_path, $output_path, $object, $object_type, array(
            'xl/sharedStrings.xml',
            'xl/worksheets/sheet1.xml',
            'xl/worksheets/sheet2.xml',
            'xl/worksheets/sheet3.xml'
        ));
    }

    /**
     * Process DOCX template with variable substitution
     *
     * @param string $template_path Path to template
     * @param string $output_path Path for output file
     * @param CommonObject $object Object for substitution
     * @param string $object_type Object type
     * @return int >0 if OK, <0 if KO
     */
    protected function processDOCX($template_path, $output_path, $object, $object_type)
    {
        global $langs;

        if (!class_exists('ZipArchive')) {
            $this->error = 'ZipArchive class not available';
            return $this->processCopy($template_path, $output_path, $object, $object_type);
        }

        // Copy template to output location first
        if (!dol_copy($template_path, $output_path, 0, 1)) {
            $this->error = $langs->trans('ErrorFileCopyFailed');
            return -20;
        }

        // Open the copied file for modification
        $zip = new ZipArchive();
        if ($zip->open($output_path) !== true) {
            $this->error = $langs->trans('ErrorCanNotOpenFile', $output_path);
            return -21;
        }

        // Get substitution array
        $substitutions = $this->getSubstitutionArray($object, $object_type);

        // DOCX files to process
        $xml_files = array(
            'word/document.xml',
            'word/header1.xml',
            'word/header2.xml',
            'word/header3.xml',
            'word/footer1.xml',
            'word/footer2.xml',
            'word/footer3.xml'
        );

        foreach ($xml_files as $xml_file) {
            $content = $zip->getFromName($xml_file);
            if ($content !== false) {
                // Clean up split XML runs - Word often splits {variable} into multiple <w:t> tags
                // This regex merges adjacent <w:t> tags within the same paragraph
                $content = $this->cleanDocxXml($content);

                // Replace variables in content
                foreach ($substitutions as $key => $value) {
                    if ($value === null) {
                        $value = '';
                    }
                    $content = str_replace('{'.$key.'}', htmlspecialchars($value, ENT_XML1, 'UTF-8'), $content);
                }

                // Update the file in the ZIP
                $zip->deleteName($xml_file);
                $zip->addFromString($xml_file, $content);
            }
        }

        $zip->close();

        return 1;
    }

    /**
     * Clean DOCX XML by merging split text runs
     * Word often splits text like {company_name} into multiple <w:t> elements
     *
     * @param string $content XML content
     * @return string Cleaned XML content
     */
    protected function cleanDocxXml($content)
    {
        // Pattern to find variables that might be split across runs
        // Look for patterns like </w:t></w:r><w:r><w:t> or </w:t></w:r><w:r...><w:t>
        // and merge them when they contain parts of our {variable} tags

        // First, let's try to find and fix split curly braces
        // Match text between { and } that might span multiple runs
        $pattern = '/\{([^}]*?)(<\/w:t><\/w:r>.*?<w:r[^>]*>.*?<w:t[^>]*>)+([^}]*?)\}/s';

        $content = preg_replace_callback($pattern, function($matches) {
            // Remove all XML tags between the braces, keep only text
            $inner = $matches[1] . $matches[2] . $matches[3];
            $inner = preg_replace('/<[^>]+>/', '', $inner);
            return '{'.$inner.'}';
        }, $content);

        // Also handle simpler cases where just the variable name is split
        // Pattern: {text</w:t></w:r><w:r><w:t>more_text}
        $content = preg_replace_callback(
            '/(\{[^}<]*)<\/w:t><\/w:r>(<w:r[^>]*>)?<w:t[^>]*>([^}]*\})/s',
            function($matches) {
                return $matches[1] . $matches[3];
            },
            $content
        );

        // Repeat a few times to catch nested splits
        for ($i = 0; $i < 5; $i++) {
            $content = preg_replace_callback(
                '/(\{[^}<]*)<\/w:t><\/w:r>(<w:r[^>]*>)?<w:t[^>]*>([^}]*\})/s',
                function($matches) {
                    return $matches[1] . $matches[3];
                },
                $content
            );
        }

        return $content;
    }

    /**
     * Process ZIP-based XML document (ODS, XLSX, DOCX) with variable substitution
     *
     * @param string $template_path Path to template
     * @param string $output_path Path for output file
     * @param CommonObject $object Object for substitution
     * @param string $object_type Object type
     * @param string|array $xml_files XML file(s) inside the ZIP to process
     * @return int >0 if OK, <0 if KO
     */
    protected function processZipXML($template_path, $output_path, $object, $object_type, $xml_files)
    {
        global $langs;

        if (!class_exists('ZipArchive')) {
            $this->error = 'ZipArchive class not available';
            return $this->processCopy($template_path, $output_path, $object, $object_type);
        }

        // Copy template to output location first
        if (!dol_copy($template_path, $output_path, 0, 1)) {
            $this->error = $langs->trans('ErrorFileCopyFailed');
            return -20;
        }

        // Open the copied file for modification
        $zip = new ZipArchive();
        if ($zip->open($output_path) !== true) {
            $this->error = $langs->trans('ErrorCanNotOpenFile', $output_path);
            return -21;
        }

        // Get substitution array
        $substitutions = $this->getSubstitutionArray($object, $object_type);

        // Process each XML file
        if (!is_array($xml_files)) {
            $xml_files = array($xml_files);
        }

        foreach ($xml_files as $xml_file) {
            $content = $zip->getFromName($xml_file);
            if ($content !== false) {
                // Replace variables in content
                foreach ($substitutions as $key => $value) {
                    // Handle null values
                    if ($value === null) {
                        $value = '';
                    }
                    // Replace {variable} format
                    $content = str_replace('{'.$key.'}', htmlspecialchars($value, ENT_XML1, 'UTF-8'), $content);
                }

                // Update the file in the ZIP
                $zip->deleteName($xml_file);
                $zip->addFromString($xml_file, $content);
            }
        }

        $zip->close();

        return 1;
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
     * Uses Dolibarr's standard ODT/ODS substitution tags
     * See: https://wiki.dolibarr.org/index.php/Create_an_ODT_or_ODS_document_template
     *
     * @param CommonObject $object Object
     * @param string $object_type Object type
     * @return array Substitution array
     */
    public function getSubstitutionArray($object, $object_type)
    {
        global $conf, $langs, $mysoc, $user;

        $substitutions = array();

        // Use Dolibarr's standard substitution tags (lowercase with underscores)
        // These match the official Dolibarr ODT template documentation

        if ($object_type == 'thirdparty') {
            $substitutions = $this->getThirdpartySubstitutions($object);
        } else {
            $substitutions = $this->getContactSubstitutions($object);
        }

        // Add mycompany substitutions (own company info)
        if (is_object($mysoc)) {
            $substitutions['mycompany_name'] = $mysoc->name;
            $substitutions['mycompany_address'] = $mysoc->address;
            $substitutions['mycompany_zip'] = $mysoc->zip;
            $substitutions['mycompany_town'] = $mysoc->town;
            $substitutions['mycompany_country'] = $mysoc->country;
            $substitutions['mycompany_country_code'] = $mysoc->country_code;
            $substitutions['mycompany_state'] = $mysoc->state;
            $substitutions['mycompany_phone'] = $mysoc->phone;
            $substitutions['mycompany_fax'] = $mysoc->fax;
            $substitutions['mycompany_email'] = $mysoc->email;
            $substitutions['mycompany_web'] = $mysoc->url;
            $substitutions['mycompany_idprof1'] = $mysoc->idprof1;
            $substitutions['mycompany_idprof2'] = $mysoc->idprof2;
            $substitutions['mycompany_idprof3'] = $mysoc->idprof3;
            $substitutions['mycompany_idprof4'] = $mysoc->idprof4;
            $substitutions['mycompany_idprof5'] = $mysoc->idprof5;
            $substitutions['mycompany_idprof6'] = $mysoc->idprof6;
            $substitutions['mycompany_vatnumber'] = $mysoc->tva_intra;
            $substitutions['mycompany_capital'] = $mysoc->capital;
            $substitutions['mycompany_note_public'] = $mysoc->note_public;
            $substitutions['mycompany_default_bank_iban'] = !empty($mysoc->bank_account) ? $mysoc->bank_account->iban : '';
            $substitutions['mycompany_default_bank_bic'] = !empty($mysoc->bank_account) ? $mysoc->bank_account->bic : '';

            // My company logo URL
            $substitutions['mycompany_logo'] = '';
            if (!empty($mysoc->logo)) {
                $substitutions['mycompany_logo'] = DOL_URL_ROOT.'/viewimage.php?modulepart=mycompany&entity='.$conf->entity.'&file='.urlencode('logos/'.$mysoc->logo);
            }
        }

        // Add logged-in user substitutions (complete user info)
        $substitutions['user_login'] = $user->login;
        $substitutions['user_firstname'] = $user->firstname;
        $substitutions['user_lastname'] = $user->lastname;
        $substitutions['user_fullname'] = $user->getFullName($langs);
        $substitutions['user_email'] = $user->email;
        $substitutions['user_phone'] = $user->office_phone;
        $substitutions['user_phone_mobile'] = $user->user_mobile;
        $substitutions['user_fax'] = $user->office_fax;
        $substitutions['user_address'] = $user->address;
        $substitutions['user_zip'] = $user->zip;
        $substitutions['user_town'] = $user->town;
        $substitutions['user_country'] = $user->country;
        $substitutions['user_signature'] = $user->signature;
        $substitutions['user_job'] = $user->job;
        $substitutions['user_note_public'] = $user->note_public;
        $substitutions['user_note_private'] = $user->note_private;

        // User extra fields: {user_options_xxx}
        if (!empty($user->array_options)) {
            foreach ($user->array_options as $key => $val) {
                $clean_key = preg_replace('/^options_/', '', $key);
                $substitutions['user_options_'.$clean_key] = $val;
            }
        }

        // Add date/time substitutions - Standard Dolibarr format
        $now = dol_now();
        $substitutions['date'] = dol_print_date($now, 'day');
        $substitutions['datehour'] = dol_print_date($now, 'dayhour');
        $substitutions['year'] = date('Y');
        $substitutions['month'] = date('m');
        $substitutions['day'] = date('d');

        // Dolibarr standard date variables (from wiki documentation)
        $substitutions['current_date'] = dol_print_date($now, 'day');
        $substitutions['current_datehour'] = dol_print_date($now, 'dayhour');
        $substitutions['current_server_date'] = dol_print_date($now, 'day');
        $substitutions['current_server_datehour'] = dol_print_date($now, 'dayhour');
        $substitutions['current_date_locale'] = dol_print_date($now, 'daytext');
        $substitutions['current_datehour_locale'] = dol_print_date($now, 'dayhourtext');
        $substitutions['current_server_date_locale'] = dol_print_date($now, 'daytext');
        $substitutions['current_server_datehour_locale'] = dol_print_date($now, 'dayhourtext');

        return $substitutions;
    }

    /**
     * Get thirdparty-specific substitutions
     * Uses Dolibarr's standard {company_xxx} tags
     *
     * @param Societe $object Thirdparty object
     * @return array Substitutions
     */
    protected function getThirdpartySubstitutions($object)
    {
        global $langs;

        $subs = array();

        // Ensure extra fields are loaded
        if (empty($object->array_options) && method_exists($object, 'fetch_optionals')) {
            $object->fetch_optionals();
        }

        // Standard Dolibarr company/thirdparty substitution tags
        // As documented in wiki.dolibarr.org ODT template guide
        $subs['company_name'] = $object->name;
        $subs['company_name_alias'] = $object->name_alias;
        $subs['company_address'] = $object->address;
        $subs['company_zip'] = $object->zip;
        $subs['company_town'] = $object->town;
        $subs['company_country'] = $object->country;
        $subs['company_country_code'] = $object->country_code;
        $subs['company_state'] = $object->state;
        $subs['company_state_code'] = $object->state_code;
        $subs['company_phone'] = $object->phone;
        $subs['company_fax'] = $object->fax;
        $subs['company_email'] = $object->email;
        $subs['company_web'] = $object->url;
        $subs['company_barcode'] = $object->barcode;

        // Codes
        $subs['company_customercode'] = $object->code_client;
        $subs['company_suppliercode'] = $object->code_fournisseur;
        $subs['company_customeraccountancycode'] = $object->code_compta;
        $subs['company_supplieraccountancycode'] = $object->code_compta_fournisseur;

        // Legal/Professional IDs
        $subs['company_idprof1'] = $object->idprof1;
        $subs['company_idprof2'] = $object->idprof2;
        $subs['company_idprof3'] = $object->idprof3;
        $subs['company_idprof4'] = $object->idprof4;
        $subs['company_idprof5'] = $object->idprof5;
        $subs['company_idprof6'] = $object->idprof6;
        $subs['company_vatnumber'] = $object->tva_intra;
        $subs['company_capital'] = $object->capital;
        $subs['company_juridicalstatus'] = $object->forme_juridique;
        $subs['company_outstanding_limit'] = $object->outstanding_limit;

        // Notes
        $subs['company_note_public'] = $object->note_public;
        $subs['company_note_private'] = $object->note_private;

        // Bank info (if loaded)
        if (!empty($object->bank_account)) {
            $subs['company_default_bank_iban'] = $object->bank_account->iban;
            $subs['company_default_bank_bic'] = $object->bank_account->bic;
        } else {
            $subs['company_default_bank_iban'] = '';
            $subs['company_default_bank_bic'] = '';
        }

        // Extra fields support: {company_options_xxx}
        if (!empty($object->array_options)) {
            foreach ($object->array_options as $key => $val) {
                // Remove 'options_' prefix for cleaner tag names
                $clean_key = preg_replace('/^options_/', '', $key);
                $subs['company_options_'.$clean_key] = $val;
            }
        }

        // Company logo URL
        global $conf;
        $subs['company_logo'] = '';
        if (!empty($object->logo)) {
            $subs['company_logo'] = DOL_URL_ROOT.'/viewimage.php?modulepart=societe&entity='.$conf->entity.'&file='.urlencode($object->id.'/logos/'.$object->logo);
        }

        return $subs;
    }

    /**
     * Get contact-specific substitutions
     * Uses Dolibarr's standard {contact_xxx} tags
     * Also includes thirdparty (company) tags and logged-in user tags
     *
     * @param Contact $object Contact object
     * @return array Substitutions
     */
    protected function getContactSubstitutions($object)
    {
        global $langs;

        $subs = array();

        // Ensure extra fields are loaded for contact
        if (empty($object->array_options) && method_exists($object, 'fetch_optionals')) {
            $object->fetch_optionals();
        }

        // Standard Dolibarr contact substitution tags
        $subs['contact_civility'] = $object->civility;
        $subs['contact_firstname'] = $object->firstname;
        $subs['contact_lastname'] = $object->lastname;
        $subs['contact_fullname'] = $object->getFullName($langs);
        $subs['contact_poste'] = $object->poste;

        // Address
        $subs['contact_address'] = $object->address;
        $subs['contact_zip'] = $object->zip;
        $subs['contact_town'] = $object->town;
        $subs['contact_state'] = $object->state;
        $subs['contact_state_code'] = $object->state_code;
        $subs['contact_country'] = $object->country;
        $subs['contact_country_code'] = $object->country_code;

        // Contact info
        $subs['contact_phone'] = $object->phone_pro;
        $subs['contact_phone_pro'] = $object->phone_pro;
        $subs['contact_phone_perso'] = $object->phone_perso;
        $subs['contact_phone_mobile'] = $object->phone_mobile;
        $subs['contact_fax'] = $object->fax;
        $subs['contact_email'] = $object->email;

        // Notes
        $subs['contact_note_public'] = $object->note_public;
        $subs['contact_note_private'] = $object->note_private;

        // Birthday
        if (!empty($object->birthday)) {
            $subs['contact_birthday'] = dol_print_date($object->birthday, 'day');
        } else {
            $subs['contact_birthday'] = '';
        }

        // If contact has a linked thirdparty, also include company tags
        // This allows using {company_xxx} variables in contact templates
        if (!empty($object->socid) && $object->socid > 0) {
            require_once DOL_DOCUMENT_ROOT.'/societe/class/societe.class.php';
            $thirdparty = new Societe($this->db);
            if ($thirdparty->fetch($object->socid) > 0) {
                // Fetch thirdparty extra fields
                $thirdparty->fetch_optionals();
                // Add company tags for the linked thirdparty
                $company_subs = $this->getThirdpartySubstitutions($thirdparty);
                $subs = array_merge($subs, $company_subs);
            }
        }

        // Extra fields support: {contact_options_xxx}
        if (!empty($object->array_options)) {
            foreach ($object->array_options as $key => $val) {
                $clean_key = preg_replace('/^options_/', '', $key);
                $subs['contact_options_'.$clean_key] = $val;
            }
        }

        // Contact photo URL
        global $conf;
        $subs['contact_photo'] = '';
        if (!empty($object->photo)) {
            $subs['contact_photo'] = DOL_URL_ROOT.'/viewimage.php?modulepart=contact&entity='.$conf->entity.'&file='.urlencode($object->id.'/photos/'.$object->photo);
        }

        return $subs;
    }
}
