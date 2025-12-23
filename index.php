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
