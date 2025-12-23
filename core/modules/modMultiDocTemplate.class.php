<?php
/* Copyright (C) 2024
 *
 * This program is free software: you can redistribute it and/or modify
 * it under the terms of the GNU General Public License as published by
 * the Free Software Foundation, either version 3 of the License, or
 * (at your option) any later version.
 */

include_once DOL_DOCUMENT_ROOT.'/core/modules/DolibarrModules.class.php';

/**
 * Class modMultiDocTemplate
 * Module for group-based document templates and archive generation
 */
class modMultiDocTemplate extends DolibarrModules
{
    /**
     * Constructor
     *
     * @param DoliDB $db Database handler
     */
    public function __construct($db)
    {
        global $conf, $langs;

        $this->db = $db;

        // Module ID (must be unique)
        $this->numero = 240600;

        // Key text used to identify module
        $this->rights_class = 'multidoctemplate';

        // Family
        $this->family = 'tools';
        $this->module_position = '50';

        // Module label
        $this->name = preg_replace('/^mod/i', '', get_class($this));
        $this->description = 'Group-based document templates and archive generation';
        $this->descriptionlong = 'Upload document templates (ODT, DOCX, XLSX, ODS) by user group and generate archives with Dolibarr variable substitution.<br><br>'.
            '<strong>Supported formats:</strong> ODT, ODS, XLSX, DOCX, PDF, DOC, XLS, RTF<br><br>'.
            '<strong>Available substitution variables:</strong><br>'.
            '<em>Thirdparty:</em> {company_name}, {company_address}, {company_zip}, {company_town}, {company_country}, {company_phone}, {company_email}, {company_web}, {company_vatnumber}, {company_idprof1}-{company_idprof6}, {company_customercode}, {company_note_public}, {company_options_XXX}<br>'.
            '<em>Contact:</em> {contact_firstname}, {contact_lastname}, {contact_fullname}, {contact_email}, {contact_phone}, {contact_address}, {contact_poste}, {contact_options_XXX}<br>'.
            '<em>Logged-in User:</em> {user_login}, {user_firstname}, {user_lastname}, {user_fullname}, {user_email}, {user_phone}, {user_signature}, {user_job}, {user_options_XXX}<br>'.
            '<em>My Company:</em> {mycompany_name}, {mycompany_address}, {mycompany_phone}, {mycompany_email}, {mycompany_vatnumber}, etc.<br>'.
            '<em>Dates:</em> {current_date}, {current_datehour}, {current_date_locale}, {current_datehour_locale}, {date}, {year}, {month}, {day}<br><br>'.
            '<em>Note: Replace XXX with your extra field code for custom fields.</em>';

        // Version
        $this->version = '1.0.0';

        // Module constants
        $this->const_name = 'MAIN_MODULE_MULTIDOCTEMPLATE';
        $this->picto = 'file';

        // Module parts - hooks for extending functionality
        $this->module_parts = array(
            'triggers' => 0,
            'login' => 0,
            'substitutions' => 0,
            'menus' => 0,
            'theme' => 0,
            'tpl' => 0,
            'barcode' => 0,
            'models' => 0,
            'hooks' => array(
                'thirdpartycard',
                'contactcard'
            ),
            'moduleforexternal' => 0
        );

        // Required modules
        $this->depends = array(
            'modSociete',
            'modUser'
        );
        $this->requiredby = array();
        $this->conflictwith = array();

        // Data directories to create when module is enabled
        $this->dirs = array(
            '/multidoctemplate',
            '/multidoctemplate/templates',
            '/multidoctemplate/archives'
        );

        // Config page
        $this->config_page_url = array('setup.php@multidoctemplate');

        // Language files
        $this->langfiles = array('multidoctemplate@multidoctemplate');

        // Initialize conf
        if (!isset($conf->multidoctemplate) || !isset($conf->multidoctemplate->enabled)) {
            $conf->multidoctemplate = new stdClass();
            $conf->multidoctemplate->enabled = 0;
        }

        // TABS - This is the correct way to add tabs in Dolibarr 23
        $this->tabs = array();

        // Tab for user groups - Templates upload
        $this->tabs[] = array(
            'data' => 'group:+templates:Templates:multidoctemplate@multidoctemplate:$user->hasRight("multidoctemplate","template_voir"):/custom/multidoctemplate/templates.php?id=__ID__&object_type=usergroup'
        );

        // Tab for third party - Archives
        $this->tabs[] = array(
            'data' => 'thirdparty:+archives:Archives:multidoctemplate@multidoctemplate:$user->hasRight("societe","lire"):/custom/multidoctemplate/archives.php?id=__ID__&object_type=thirdparty'
        );

        // Tab for contacts - Archives
        $this->tabs[] = array(
            'data' => 'contact:+archives:Archives:multidoctemplate@multidoctemplate:$user->hasRight("societe","lire"):/custom/multidoctemplate/archives.php?id=__ID__&object_type=contact'
        );

        // Constants
        $this->const = array();

        // Boxes/Widgets
        $this->boxes = array();

        // Cronjobs
        $this->cronjobs = array();

        // Permissions
        $this->rights = array();
        $r = 0;

        // Global read permission
        $this->rights[$r][0] = $this->numero + $r;
        $this->rights[$r][1] = 'Read MultiDocTemplate';
        $this->rights[$r][3] = 1;
        $this->rights[$r][4] = 'lire';
        $this->rights[$r][5] = '';
        $r++;

        // RGPD/Archives permissions
        $this->rights[$r][0] = $this->numero + $r;
        $this->rights[$r][1] = 'View archives';
        $this->rights[$r][3] = 1;
        $this->rights[$r][4] = 'archive_voir';
        $this->rights[$r][5] = '';
        $r++;

        $this->rights[$r][0] = $this->numero + $r;
        $this->rights[$r][1] = 'Generate archives';
        $this->rights[$r][3] = 0;
        $this->rights[$r][4] = 'archive_creer';
        $this->rights[$r][5] = '';
        $r++;

        $this->rights[$r][0] = $this->numero + $r;
        $this->rights[$r][1] = 'Delete archives';
        $this->rights[$r][3] = 0;
        $this->rights[$r][4] = 'archive_supprimer';
        $this->rights[$r][5] = '';
        $r++;

        // Template permissions
        $this->rights[$r][0] = $this->numero + $r;
        $this->rights[$r][1] = 'View templates';
        $this->rights[$r][3] = 1;
        $this->rights[$r][4] = 'template_voir';
        $this->rights[$r][5] = '';
        $r++;

        $this->rights[$r][0] = $this->numero + $r;
        $this->rights[$r][1] = 'Upload templates';
        $this->rights[$r][3] = 0;
        $this->rights[$r][4] = 'template_creer';
        $this->rights[$r][5] = '';
        $r++;

        $this->rights[$r][0] = $this->numero + $r;
        $this->rights[$r][1] = 'Delete templates';
        $this->rights[$r][3] = 0;
        $this->rights[$r][4] = 'template_supprimer';
        $this->rights[$r][5] = '';
        $r++;

        // Menus - No left menu, module works only through tabs on User Groups, Third Party and Contact cards
        $this->menu = array();
    }

    /**
     * Function called when module is enabled
     *
     * @param string $options Options when enabling module
     * @return int 1 if OK, 0 if KO
     */
    public function init($options = '')
    {
        $result = $this->_load_tables('/multidoctemplate/sql/');
        if ($result < 0) {
            return -1;
        }

        $sql = array();

        return $this->_init($sql, $options);
    }

    /**
     * Function called when module is disabled
     *
     * @param string $options Options when disabling module
     * @return int 1 if OK, 0 if KO
     */
    public function remove($options = '')
    {
        $sql = array();

        return $this->_remove($sql, $options);
    }
}
