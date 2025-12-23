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
$res = 0;
// Try main.inc.php into web root known defined into CONTEXT_DOCUMENT_ROOT (not always defined)
if (!$res && !empty($_SERVER["CONTEXT_DOCUMENT_ROOT"])) {
    $res = @include $_SERVER["CONTEXT_DOCUMENT_ROOT"]."/main.inc.php";
}
// Try main.inc.php into web root detected using web root calculated from SCRIPT_FILENAME
$tmp = empty($_SERVER['SCRIPT_FILENAME']) ? '' : $_SERVER['SCRIPT_FILENAME'];
$tmp2 = realpath(__FILE__);
$i = strlen($tmp) - 1;
$j = strlen($tmp2) - 1;
while ($i > 0 && $j > 0 && isset($tmp[$i]) && isset($tmp2[$j]) && $tmp[$i] == $tmp2[$j]) {
    $i--;
    $j--;
}
if (!$res && $i > 0 && file_exists(substr($tmp, 0, ($i + 1))."/main.inc.php")) {
    $res = @include substr($tmp, 0, ($i + 1))."/main.inc.php";
}
if (!$res && $i > 0 && file_exists(dirname(substr($tmp, 0, ($i + 1)))."/main.inc.php")) {
    $res = @include dirname(substr($tmp, 0, ($i + 1)))."/main.inc.php";
}
// Try main.inc.php using relative path
if (!$res && file_exists("../main.inc.php")) {
    $res = @include "../main.inc.php";
}
if (!$res && file_exists("../../main.inc.php")) {
    $res = @include "../../main.inc.php";
}
if (!$res) {
    die("Include of main fails");
}
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
                $template->tag = GETPOST('template_tag', 'alphanohtml');
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

    // Tag (folder)
    print '<tr class="oddeven">';
    print '<td class="titlefield">'.$langs->trans('Tag').' <span class="star">*</span></td>';
    print '<td><input type="text" name="template_tag" size="40" class="flat" placeholder="e.g. course, exam, credential" required></td>';
    print '</tr>';

    // Label
    print '<tr class="oddeven">';
    print '<td>'.$langs->trans('Label').'</td>';
    print '<td><input type="text" name="template_label" size="40" class="flat"></td>';
    print '</tr>';

    // Description
    print '<tr class="oddeven">';
    print '<td>'.$langs->trans('Description').'</td>';
    print '<td><textarea name="template_description" rows="2" cols="40" class="flat"></textarea></td>';
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

// Get sorting parameters
$sortfield = GETPOST('sortfield', 'aZ09comma');
$sortorder = GETPOST('sortorder', 'aZ09comma');
if (empty($sortfield)) $sortfield = 'tag';
if (empty($sortorder)) $sortorder = 'ASC';

// Expand/Collapse controls
print '<div style="margin-bottom: 10px;">';
print load_fiche_titre($langs->trans('TemplatesList'), '<a href="javascript:expandAllFolders();">'.$langs->trans('ExpandAll').'</a> | <a href="javascript:collapseAllFolders();">'.$langs->trans('CollapseAll').'</a>', '');
print '</div>';

$templates = $template->fetchAllByUserGroup($object->id, -1);

// Group templates by tag
$templates_by_tag = array();
if (is_array($templates) && count($templates) > 0) {
    foreach ($templates as $tpl) {
        $tag_key = !empty($tpl->tag) ? $tpl->tag : $langs->trans('Uncategorized');
        if (!isset($templates_by_tag[$tag_key])) {
            $templates_by_tag[$tag_key] = array();
        }
        $templates_by_tag[$tag_key][] = $tpl;
    }
    // Sort tags
    if ($sortfield == 'tag') {
        if ($sortorder == 'ASC') {
            ksort($templates_by_tag);
        } else {
            krsort($templates_by_tag);
        }
    }
}

print '<div class="div-table-responsive">';

if (count($templates_by_tag) > 0) {
    $tag_index = 0;
    foreach ($templates_by_tag as $tag_label => $tag_templates) {
        $tag_id = 'tag_'.md5($tag_label);
        $count = count($tag_templates);
        $tag_index++;

        // Tag folder header (collapsible) - Dolibarr style
        print '<div class="template-tag-folder" data-tag="'.dol_escape_htmltag(strtolower($tag_label)).'" style="margin-bottom: 10px;">';
        print '<div class="folder-header" onclick="toggleTagFolder(\''.$tag_id.'\')" style="cursor: pointer; padding: 8px; background: #e8e8e8; margin-bottom: 0; border-radius: 3px 3px 0 0;">';
        print '<span id="'.$tag_id.'_icon">'.img_picto('', '1downarrow', 'class="imgdownforline"').'</span> ';
        print '<strong>'.img_picto('', 'folder', 'style="vertical-align: middle;"').' '.dol_escape_htmltag($tag_label).'</strong>';
        print ' <span class="opacitymedium">('.$count.' '.$langs->trans('Files').')</span>';
        print '</div>';

        // Tag folder content
        print '<div id="'.$tag_id.'_content" class="folder-content">';
        print '<table class="noborder centpercent" style="margin-top: 0;">';

        // Table header with sortable columns
        print '<tr class="liste_titre">';
        print_liste_field_titre($langs->trans('Label'), $_SERVER['PHP_SELF'], 'label', '', '&id='.$object->id, '', $sortfield, $sortorder);
        print_liste_field_titre($langs->trans('Filename'), $_SERVER['PHP_SELF'], 'filename', '', '&id='.$object->id, '', $sortfield, $sortorder);
        print_liste_field_titre($langs->trans('Type'), $_SERVER['PHP_SELF'], 'filetype', '', '&id='.$object->id, '', $sortfield, $sortorder);
        print '<th class="center">'.$langs->trans('Actions').'</th>';
        print '</tr>';

        // Sort templates within tag if needed
        if ($sortfield == 'label') {
            usort($tag_templates, function($a, $b) use ($sortorder) {
                return $sortorder == 'ASC' ? strcmp($a->label, $b->label) : strcmp($b->label, $a->label);
            });
        } elseif ($sortfield == 'filename') {
            usort($tag_templates, function($a, $b) use ($sortorder) {
                return $sortorder == 'ASC' ? strcmp($a->filename, $b->filename) : strcmp($b->filename, $a->filename);
            });
        } elseif ($sortfield == 'filetype') {
            usort($tag_templates, function($a, $b) use ($sortorder) {
                return $sortorder == 'ASC' ? strcmp($a->filetype, $b->filetype) : strcmp($b->filetype, $a->filetype);
            });
        }

        foreach ($tag_templates as $tpl) {
            print '<tr class="oddeven">';

            // Label with info icon
            print '<td>'.dol_escape_htmltag($tpl->label);
            if (!empty($tpl->description)) {
                print ' '.img_picto(dol_escape_htmltag($tpl->description), 'info', 'class="classfortooltip" title="'.dol_escape_htmltag($tpl->description).'"');
            }
            print '</td>';

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

        print '</table>';
        print '</div>'; // folder-content
        print '</div>'; // template-tag-folder
    }
} else {
    print '<table class="noborder centpercent">';
    print '<tr class="liste_titre">';
    print '<th>'.$langs->trans('Label').'</th>';
    print '<th>'.$langs->trans('Filename').'</th>';
    print '<th>'.$langs->trans('Type').'</th>';
    print '<th class="center">'.$langs->trans('Actions').'</th>';
    print '</tr>';
    print '<tr class="oddeven"><td colspan="4" class="opacitymedium">'.$langs->trans('NoTemplatesYet').'</td></tr>';
    print '</table>';
}

print '</div>';

// JavaScript for expand/collapse - Dolibarr style icons
$downArrowJs = addslashes(img_picto('', '1downarrow', 'class="imgdownforline"'));
$rightArrowJs = addslashes(img_picto('', '1rightarrow', 'class="imgdownforline"'));

print '<script type="text/javascript">
var downArrowHtml = \''.$downArrowJs.'\';
var rightArrowHtml = \''.$rightArrowJs.'\';

function toggleTagFolder(tagId) {
    var content = document.getElementById(tagId + "_content");
    var icon = document.getElementById(tagId + "_icon");
    if (content.style.display === "none") {
        content.style.display = "block";
        icon.innerHTML = downArrowHtml;
    } else {
        content.style.display = "none";
        icon.innerHTML = rightArrowHtml;
    }
}

function expandAllFolders() {
    var contents = document.querySelectorAll(".folder-content");
    var icons = document.querySelectorAll("[id$=\'_icon\']");
    contents.forEach(function(el) { el.style.display = "block"; });
    icons.forEach(function(el) {
        if (el.id.startsWith("tag_")) {
            el.innerHTML = downArrowHtml;
        }
    });
}

function collapseAllFolders() {
    var contents = document.querySelectorAll(".folder-content");
    var icons = document.querySelectorAll("[id$=\'_icon\']");
    contents.forEach(function(el) { el.style.display = "none"; });
    icons.forEach(function(el) {
        if (el.id.startsWith("tag_")) {
            el.innerHTML = rightArrowHtml;
        }
    });
}
</script>';

llxFooter();
$db->close();
