<?php
/* Copyright (C) 2024
 *
 * This program is free software: you can redistribute it and/or modify
 * it under the terms of the GNU General Public License as published by
 * the Free Software Foundation, either version 3 of the License, or
 * (at your option) any later version.
 */

/**
 * \file       index.php
 * \brief      MultiDocTemplate module home page
 */

require '../main.inc.php';
require_once __DIR__.'/class/template.class.php';
require_once __DIR__.'/class/archive.class.php';

$langs->loadLangs(array('multidoctemplate@multidoctemplate'));

// Security check
if (!$user->hasRight('multidoctemplate', 'lire')) {
    accessforbidden();
}

$title = $langs->trans('MultiDocTemplate');
llxHeader('', $title);

print load_fiche_titre($langs->trans('MultiDocTemplate'), '', 'file');

print '<div class="fichecenter">';

// Module description
print '<div class="info">';
print $langs->trans('ModuleMultiDocTemplateDesc');
print '</div>';

print '<br>';

// Quick stats
$template = new MultiDocTemplate($db);
$archive = new MultiDocArchive($db);

// Count templates
$sql = "SELECT COUNT(*) as nb FROM ".MAIN_DB_PREFIX."multidoctemplate_template WHERE entity IN (".getEntity('multidoctemplate_template').")";
$resql = $db->query($sql);
$num_templates = 0;
if ($resql) {
    $obj = $db->fetch_object($resql);
    $num_templates = $obj->nb;
}

// Count archives
$sql = "SELECT COUNT(*) as nb FROM ".MAIN_DB_PREFIX."multidoctemplate_archive WHERE entity IN (".getEntity('multidoctemplate_archive').")";
$resql = $db->query($sql);
$num_archives = 0;
if ($resql) {
    $obj = $db->fetch_object($resql);
    $num_archives = $obj->nb;
}

print '<table class="noborder centpercent">';
print '<tr class="liste_titre">';
print '<th colspan="2">'.$langs->trans('Statistics').'</th>';
print '</tr>';

print '<tr class="oddeven">';
print '<td>'.$langs->trans('Templates').'</td>';
print '<td class="right">'.$num_templates.'</td>';
print '</tr>';

print '<tr class="oddeven">';
print '<td>'.$langs->trans('Archives').'</td>';
print '<td class="right">'.$num_archives.'</td>';
print '</tr>';

print '</table>';

print '</div>';

llxFooter();
$db->close();
