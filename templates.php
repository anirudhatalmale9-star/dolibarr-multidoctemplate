<?php
/* Copyright (C) 2024
 *
 * This program is free software: you can redistribute it and/or modify
 * it under the terms of the GNU General Public License as published by
 * the Free Software Foundation, either version 3 of the License, or
 * (at your option) any later version.
 */

/**
 * \file       templates.php
 * \brief      Template management page for user groups
 */

// Load Dolibarr environment
require '../main.inc.php';
require_once DOL_DOCUMENT_ROOT.'/user/class/usergroup.class.php';
require_once DOL_DOCUMENT_ROOT.'/core/lib/usergroups.lib.php';
require_once DOL_DOCUMENT_ROOT.'/core/lib/files.lib.php';
require_once DOL_DOCUMENT_ROOT.'/core/class/html.formfile.class.php';
require_once __DIR__.'/class/template.class.php';

// Load translations
$langs->loadLangs(array('users', 'multidoctemplate@multidoctemplate'));

// Get parameters
$id = GETPOST('id', 'int');
$action = GETPOST('action', 'aZ09');
$confirm = GETPOST('confirm', 'alpha');
$template_id = GETPOST('template_id', 'int');

// Security check
if (!$user->hasRight('multidoctemplate', 'template_voir')) {
    accessforbidden();
}

// Initialize objects
$object = new UserGroup($db);
$template = new MultiDocTemplate($db);

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

// Get upload directory
$upload_dir = MultiDocTemplate::getUploadDir($object->id);

/*
 * Actions
 */

// Upload template file
if ($action == 'upload' && $user->hasRight('multidoctemplate', 'template_creer')) {
    if (!empty($_FILES['templatefile']['name'])) {
        $filename = $_FILES['templatefile']['name'];

        // Check file extension
        if (!MultiDocTemplate::isAllowedExtension($filename)) {
            setEventMessages($langs->trans('ErrorFileExtensionNotAllowed'), null, 'errors');
        } else {
            // Create directory if not exists
            if (!is_dir($upload_dir)) {
                dol_mkdir($upload_dir);
            }

            // Sanitize filename
            $sanitized_filename = dol_sanitizeFileName($filename);
            $filepath = $upload_dir.'/'.$sanitized_filename;

            // Move uploaded file
            if (dol_move_uploaded_file($_FILES['templatefile']['tmp_name'], $filepath, 1, 0, $_FILES['templatefile']['error']) > 0) {
                // Create template record
                $template->ref = 'TPL-'.$object->id.'-'.date('YmdHis');
                $template->label = GETPOST('template_label', 'alphanohtml') ?: pathinfo($filename, PATHINFO_FILENAME);
                $template->description = GETPOST('template_description', 'restricthtml');
                $template->fk_usergroup = $object->id;
                $template->filename = $sanitized_filename;
                $template->filepath = $filepath;
                $template->filetype = strtolower(pathinfo($filename, PATHINFO_EXTENSION));
                $template->filesize = filesize($filepath);
                $template->mime_type = MultiDocTemplate::getMimeType($filename);

                $result = $template->create($user);

                if ($result > 0) {
                    setEventMessages($langs->trans('TemplateUploadSuccess'), null, 'mesgs');
                } else {
                    // Delete file if DB insert failed
                    dol_delete_file($filepath);
                    setEventMessages($template->error, null, 'errors');
                }
            } else {
                setEventMessages($langs->trans('ErrorFileUploadFailed'), null, 'errors');
            }
        }
    } else {
        setEventMessages($langs->trans('ErrorFieldRequired', $langs->transnoentitiesnoconv('File')), null, 'errors');
    }

    header('Location: '.$_SERVER['PHP_SELF'].'?id='.$object->id);
    exit;
}

// Delete template
if ($action == 'confirm_delete' && $confirm == 'yes' && $user->hasRight('multidoctemplate', 'template_supprimer')) {
    if ($template_id > 0) {
        $template->fetch($template_id);
        if ($template->id > 0) {
            $result = $template->delete($user);
            if ($result > 0) {
                setEventMessages($langs->trans('TemplateDeleted'), null, 'mesgs');
            } else {
                setEventMessages($template->error, null, 'errors');
            }
        }
    }
    header('Location: '.$_SERVER['PHP_SELF'].'?id='.$object->id);
    exit;
}

/*
 * View
 */

$title = $langs->trans('Templates').' - '.$object->name;
llxHeader('', $title);

// Prepare tabs
$head = group_prepare_head($object);

print dol_get_fiche_head($head, 'templates', $langs->trans('Group'), -1, 'group');

// Group info
$linkback = '<a href="'.DOL_URL_ROOT.'/user/group/list.php?restore_lastsearch_values=1">'.$langs->trans('BackToList').'</a>';
dol_banner_tab($object, 'id', $linkback, $user->hasRight('user', 'user', 'lire'));

print '<div class="fichecenter">';
print '<div class="underbanner clearboth"></div>';
print '</div>';

print dol_get_fiche_end();

// Delete confirmation dialog
if ($action == 'delete') {
    $formconfirm = $form->formconfirm(
        $_SERVER['PHP_SELF'].'?id='.$object->id.'&template_id='.$template_id,
        $langs->trans('DeleteTemplate'),
        $langs->trans('ConfirmDeleteTemplate'),
        'confirm_delete',
        '',
        0,
        1
    );
    print $formconfirm;
}

// Upload form
if ($user->hasRight('multidoctemplate', 'template_creer')) {
    print '<div class="tabsAction">';
    print '<form enctype="multipart/form-data" action="'.$_SERVER['PHP_SELF'].'?id='.$object->id.'" method="POST">';
    print '<input type="hidden" name="token" value="'.newToken().'">';
    print '<input type="hidden" name="action" value="upload">';

    print '<div class="fichecenter">';
    print '<table class="noborder centpercent">';
    print '<tr class="liste_titre">';
    print '<th colspan="2">'.$langs->trans('UploadNewTemplate').'</th>';
    print '</tr>';

    // Label
    print '<tr class="oddeven">';
    print '<td class="titlefield">'.$langs->trans('Label').'</td>';
    print '<td><input type="text" name="template_label" size="40" class="flat"></td>';
    print '</tr>';

    // Description
    print '<tr class="oddeven">';
    print '<td>'.$langs->trans('Description').'</td>';
    print '<td><textarea name="template_description" rows="3" cols="40" class="flat"></textarea></td>';
    print '</tr>';

    // File
    print '<tr class="oddeven">';
    print '<td>'.$langs->trans('File').' <span class="star">*</span></td>';
    print '<td>';
    print '<input type="file" name="templatefile" class="flat">';
    print '<br><small>'.$langs->trans('AllowedFormats').': '.implode(', ', MultiDocTemplate::$allowed_extensions).'</small>';
    print '</td>';
    print '</tr>';

    print '</table>';
    print '</div>';

    print '<div class="center">';
    print '<input type="submit" class="button button-primary" value="'.$langs->trans('Upload').'">';
    print '</div>';

    print '</form>';
    print '</div>';
}

// List of templates
print '<br>';
print load_fiche_titre($langs->trans('TemplatesList'), '', '');

$templates = $template->fetchAllByUserGroup($object->id, -1);

print '<div class="div-table-responsive">';
print '<table class="noborder centpercent">';
print '<tr class="liste_titre">';
print '<th>'.$langs->trans('Ref').'</th>';
print '<th>'.$langs->trans('Label').'</th>';
print '<th>'.$langs->trans('Filename').'</th>';
print '<th>'.$langs->trans('Type').'</th>';
print '<th class="right">'.$langs->trans('Size').'</th>';
print '<th class="center">'.$langs->trans('DateCreation').'</th>';
print '<th class="center">'.$langs->trans('Status').'</th>';
print '<th class="center">'.$langs->trans('Actions').'</th>';
print '</tr>';

if (is_array($templates) && count($templates) > 0) {
    foreach ($templates as $tpl) {
        print '<tr class="oddeven">';

        // Ref
        print '<td>'.$tpl->ref.'</td>';

        // Label
        print '<td>'.dol_escape_htmltag($tpl->label).'</td>';

        // Filename
        print '<td>';
        if (file_exists($tpl->filepath)) {
            print '<a href="'.DOL_URL_ROOT.'/document.php?modulepart=multidoctemplate&file=templates/group_'.$object->id.'/'.urlencode($tpl->filename).'" target="_blank">';
            print img_picto('', 'file').' '.dol_escape_htmltag($tpl->filename);
            print '</a>';
        } else {
            print '<span class="opacitymedium">'.dol_escape_htmltag($tpl->filename).' ('.$langs->trans('FileNotFound').')</span>';
        }
        print '</td>';

        // Type
        print '<td>'.strtoupper($tpl->filetype).'</td>';

        // Size
        print '<td class="right">'.dol_print_size($tpl->filesize).'</td>';

        // Date creation
        print '<td class="center">'.dol_print_date($tpl->date_creation, 'dayhour').'</td>';

        // Status
        print '<td class="center">';
        if ($tpl->active) {
            print '<span class="badge badge-status4">'.$langs->trans('Enabled').'</span>';
        } else {
            print '<span class="badge badge-status5">'.$langs->trans('Disabled').'</span>';
        }
        print '</td>';

        // Actions
        print '<td class="center nowraponall">';
        if ($user->hasRight('multidoctemplate', 'template_supprimer')) {
            print '<a class="deletefilelink" href="'.$_SERVER['PHP_SELF'].'?id='.$object->id.'&action=delete&template_id='.$tpl->id.'&token='.newToken().'">';
            print img_picto($langs->trans('Delete'), 'delete');
            print '</a>';
        }
        print '</td>';

        print '</tr>';
    }
} else {
    print '<tr class="oddeven"><td colspan="8" class="opacitymedium">'.$langs->trans('NoTemplatesYet').'</td></tr>';
}

print '</table>';
print '</div>';

llxFooter();
$db->close();
