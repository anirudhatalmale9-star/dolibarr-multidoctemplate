<?php
/* Copyright (C) 2024
 *
 * This program is free software: you can redistribute it and/or modify
 * it under the terms of the GNU General Public License as published by
 * the Free Software Foundation, either version 3 of the License, or
 * (at your option) any later version.
 */

/**
 * \file       archives.php
 * \brief      Archive generation and management for thirdparty/contact
 */

// Load Dolibarr environment
require '../main.inc.php';
require_once DOL_DOCUMENT_ROOT.'/societe/class/societe.class.php';
require_once DOL_DOCUMENT_ROOT.'/contact/class/contact.class.php';
require_once DOL_DOCUMENT_ROOT.'/core/lib/company.lib.php';
require_once DOL_DOCUMENT_ROOT.'/core/lib/contact.lib.php';
require_once DOL_DOCUMENT_ROOT.'/core/lib/files.lib.php';
require_once DOL_DOCUMENT_ROOT.'/categories/class/categorie.class.php';
require_once __DIR__.'/class/template.class.php';
require_once __DIR__.'/class/archive.class.php';
require_once __DIR__.'/class/documentgenerator.class.php';

// Load translations
$langs->loadLangs(array('companies', 'multidoctemplate@multidoctemplate'));

// Get parameters
$id = GETPOST('id', 'int');
$object_type = GETPOST('object_type', 'alpha');
$action = GETPOST('action', 'aZ09');
$confirm = GETPOST('confirm', 'alpha');
$template_id = GETPOST('template_id', 'int');
$archive_id = GETPOST('archive_id', 'int');
$category_id = GETPOST('category_id', 'int');

// Default object type
if (empty($object_type)) {
    $object_type = 'thirdparty';
}

// Security check
if (!$user->hasRight('societe', 'lire')) {
    accessforbidden();
}

// Initialize objects
$template = new MultiDocTemplate($db);
$archive = new MultiDocArchive($db);
$generator = new MultiDocGenerator($db);

// Load main object
if ($object_type == 'thirdparty') {
    $object = new Societe($db);
} else {
    $object = new Contact($db);
}

if ($id > 0) {
    $result = $object->fetch($id);
    if ($result < 0) {
        dol_print_error($db);
        exit;
    }
}

if (empty($object->id)) {
    accessforbidden('ErrorRecordNotFound');
}

// Get object's categories/tags for filtering
$object_categories = array();
if ($object_type == 'thirdparty') {
    $c = new Categorie($db);
    $cats = $c->containing($object->id, Categorie::TYPE_CUSTOMER);
    if (is_array($cats)) {
        foreach ($cats as $cat) {
            $object_categories[$cat->id] = $cat->label;
        }
    }
    $cats = $c->containing($object->id, Categorie::TYPE_SUPPLIER);
    if (is_array($cats)) {
        foreach ($cats as $cat) {
            $object_categories[$cat->id] = $cat->label;
        }
    }
}

/*
 * Actions
 */

// Generate archive from template
if ($action == 'generate' && $template_id > 0 && $user->hasRight('multidoctemplate', 'archive_creer')) {
    $template->fetch($template_id);

    if ($template->id > 0) {
        // Determine tag filter from object's categories
        $tag_filter = '';
        if ($category_id > 0 && isset($object_categories[$category_id])) {
            $tag_filter = $object_categories[$category_id];
        }

        // Generate document
        $result = $generator->generate($template, $object, $object_type, $user, $tag_filter);

        if ($result > 0) {
            setEventMessages($langs->trans('ArchiveGeneratedSuccess'), null, 'mesgs');
        } else {
            setEventMessages($generator->error, $generator->errors, 'errors');
        }
    } else {
        setEventMessages($langs->trans('ErrorTemplateNotFound'), null, 'errors');
    }

    header('Location: '.$_SERVER['PHP_SELF'].'?id='.$object->id.'&object_type='.$object_type);
    exit;
}

// Delete archive
if ($action == 'confirm_delete' && $confirm == 'yes' && $user->hasRight('multidoctemplate', 'archive_supprimer')) {
    if ($archive_id > 0) {
        $archive->fetch($archive_id);
        if ($archive->id > 0) {
            $result = $archive->delete($user);
            if ($result > 0) {
                setEventMessages($langs->trans('ArchiveDeleted'), null, 'mesgs');
            } else {
                setEventMessages($archive->error, null, 'errors');
            }
        }
    }
    header('Location: '.$_SERVER['PHP_SELF'].'?id='.$object->id.'&object_type='.$object_type);
    exit;
}

/*
 * View
 */

$title = $langs->trans('Archives').' - '.$object->name;
llxHeader('', $title);

// Prepare tabs
if ($object_type == 'thirdparty') {
    $head = societe_prepare_head($object);
    print dol_get_fiche_head($head, 'archives', $langs->trans('ThirdParty'), -1, 'company');
    $linkback = '<a href="'.DOL_URL_ROOT.'/societe/list.php?restore_lastsearch_values=1">'.$langs->trans('BackToList').'</a>';
    dol_banner_tab($object, 'socid', $linkback, ($user->hasRight('societe', 'lire') ? 1 : 0), 'rowid', 'nom');
} else {
    $head = contact_prepare_head($object);
    print dol_get_fiche_head($head, 'archives', $langs->trans('Contact'), -1, 'contact');
    $linkback = '<a href="'.DOL_URL_ROOT.'/contact/list.php?restore_lastsearch_values=1">'.$langs->trans('BackToList').'</a>';
    dol_banner_tab($object, 'id', $linkback, ($user->hasRight('societe', 'lire') ? 1 : 0), 'rowid', 'name');
}

print '<div class="fichecenter">';
print '<div class="underbanner clearboth"></div>';
print '</div>';

print dol_get_fiche_end();

// Delete confirmation dialog
if ($action == 'delete') {
    $form = new Form($db);
    $formconfirm = $form->formconfirm(
        $_SERVER['PHP_SELF'].'?id='.$object->id.'&object_type='.$object_type.'&archive_id='.$archive_id,
        $langs->trans('DeleteArchive'),
        $langs->trans('ConfirmDeleteArchive'),
        'confirm_delete',
        '',
        0,
        1
    );
    print $formconfirm;
}

// Generate archive section
if ($user->hasRight('multidoctemplate', 'archive_creer')) {
    print '<div class="tabsAction">';
    print load_fiche_titre($langs->trans('GenerateArchive'), '', '');

    // Get available templates for user
    $templates = $template->fetchAllForUser($user);

    if (is_array($templates) && count($templates) > 0) {
        print '<form action="'.$_SERVER['PHP_SELF'].'?id='.$object->id.'&object_type='.$object_type.'" method="POST">';
        print '<input type="hidden" name="token" value="'.newToken().'">';
        print '<input type="hidden" name="action" value="generate">';

        print '<div class="fichecenter">';
        print '<table class="noborder centpercent">';
        print '<tr class="liste_titre">';
        print '<th>'.$langs->trans('SelectTemplate').'</th>';
        if (!empty($object_categories)) {
            print '<th>'.$langs->trans('CategoryFilter').'</th>';
        }
        print '<th></th>';
        print '</tr>';

        print '<tr class="oddeven">';
        print '<td>';
        print '<select name="template_id" class="flat minwidth300">';
        print '<option value="">'.$langs->trans('SelectATemplate').'</option>';
        $current_group = '';
        foreach ($templates as $tpl) {
            if ($current_group != $tpl->usergroup_name) {
                if (!empty($current_group)) {
                    print '</optgroup>';
                }
                print '<optgroup label="'.dol_escape_htmltag($tpl->usergroup_name).'">';
                $current_group = $tpl->usergroup_name;
            }
            print '<option value="'.$tpl->id.'">'.dol_escape_htmltag($tpl->label).' ('.strtoupper($tpl->filetype).')</option>';
        }
        if (!empty($current_group)) {
            print '</optgroup>';
        }
        print '</select>';
        print '</td>';

        // Category filter dropdown
        if (!empty($object_categories)) {
            print '<td>';
            print '<select name="category_id" class="flat minwidth200">';
            print '<option value="">'.$langs->trans('NoFilter').'</option>';
            foreach ($object_categories as $cat_id => $cat_label) {
                print '<option value="'.$cat_id.'">'.dol_escape_htmltag($cat_label).'</option>';
            }
            print '</select>';
            print '</td>';
        }

        print '<td>';
        print '<input type="submit" class="button button-primary" value="'.$langs->trans('Generate').'">';
        print '</td>';
        print '</tr>';

        print '</table>';
        print '</div>';
        print '</form>';
    } else {
        print '<div class="opacitymedium">'.$langs->trans('NoTemplatesAvailable').'</div>';
    }

    print '</div>';
}

// List of archives
print '<br>';
print load_fiche_titre($langs->trans('ArchivesList'), '', '');

$archives = $archive->fetchAllByObject($object_type, $object->id);

print '<div class="div-table-responsive">';
print '<table class="noborder centpercent">';
print '<tr class="liste_titre">';
print '<th>'.$langs->trans('Ref').'</th>';
print '<th>'.$langs->trans('Template').'</th>';
print '<th>'.$langs->trans('Filename').'</th>';
print '<th>'.$langs->trans('Type').'</th>';
print '<th class="right">'.$langs->trans('Size').'</th>';
print '<th>'.$langs->trans('TagFilter').'</th>';
print '<th class="center">'.$langs->trans('DateGeneration').'</th>';
print '<th class="center">'.$langs->trans('Actions').'</th>';
print '</tr>';

if (is_array($archives) && count($archives) > 0) {
    foreach ($archives as $arch) {
        print '<tr class="oddeven">';

        // Ref
        print '<td>'.$arch->ref.'</td>';

        // Template
        print '<td>'.dol_escape_htmltag($arch->template_label).'</td>';

        // Filename
        print '<td>';
        if (file_exists($arch->filepath)) {
            $relative_path = str_replace(DOL_DATA_ROOT.'/multidoctemplate/', '', $arch->filepath);
            print '<a href="'.DOL_URL_ROOT.'/document.php?modulepart=multidoctemplate&file='.urlencode($relative_path).'" target="_blank">';
            print img_picto('', 'file').' '.dol_escape_htmltag($arch->filename);
            print '</a>';
        } else {
            print '<span class="opacitymedium">'.dol_escape_htmltag($arch->filename).' ('.$langs->trans('FileNotFound').')</span>';
        }
        print '</td>';

        // Type
        print '<td>'.strtoupper($arch->filetype).'</td>';

        // Size
        print '<td class="right">'.dol_print_size($arch->filesize).'</td>';

        // Tag filter
        print '<td>';
        if (!empty($arch->tag_filter)) {
            print '<span class="badge badge-secondary">'.dol_escape_htmltag($arch->tag_filter).'</span>';
        } else {
            print '-';
        }
        print '</td>';

        // Date generation
        print '<td class="center">'.dol_print_date($arch->date_generation, 'dayhour').'</td>';

        // Actions
        print '<td class="center nowraponall">';
        // Download
        if (file_exists($arch->filepath)) {
            $relative_path = str_replace(DOL_DATA_ROOT.'/multidoctemplate/', '', $arch->filepath);
            print '<a href="'.DOL_URL_ROOT.'/document.php?modulepart=multidoctemplate&file='.urlencode($relative_path).'" target="_blank">';
            print img_picto($langs->trans('Download'), 'download');
            print '</a> ';
        }
        // Delete
        if ($user->hasRight('multidoctemplate', 'archive_supprimer')) {
            print '<a class="deletefilelink" href="'.$_SERVER['PHP_SELF'].'?id='.$object->id.'&object_type='.$object_type.'&action=delete&archive_id='.$arch->id.'&token='.newToken().'">';
            print img_picto($langs->trans('Delete'), 'delete');
            print '</a>';
        }
        print '</td>';

        print '</tr>';
    }
} else {
    print '<tr class="oddeven"><td colspan="8" class="opacitymedium">'.$langs->trans('NoArchivesYet').'</td></tr>';
}

print '</table>';
print '</div>';

llxFooter();
$db->close();
