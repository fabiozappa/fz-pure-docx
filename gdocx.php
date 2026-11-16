<?php
/*
 * gdocx — Compila un modello DOCX e, se richiesto, lo converte in PDF.
 *
 * POST: apikey, nomehost, modello_docx (base64), vars, headers_vars, footers_vars,
 *       remove_vars, html_vars, table_vars, checkbox_vars, converti_pdf (si/no)
 * Risposta JSON: esito, errore, tool, data (DOCX base64), data_pdf (PDF base64, se richiesto)
 */
$tool = 'gdocx';
require __DIR__ . '/_versione.php';
require __DIR__ . '/_start.php';
require __DIR__ . '/lib/DocTemplate.php';
require __DIR__ . '/lib/GDocx.php';

$modello = post_base64('modello_docx');
if ($modello === '') {
    json_error('Parametri POST non validi!');
}
if (!str_starts_with($modello, "PK")) {
    json_error('Il modello non è un file DOCX valido');
}

$converti_pdf = in_array($_POST['converti_pdf'] ?? '', ['si', 's', 'yes', '1', 'true'], true);

$docx = GDocx::fill($modello, [
    'vars' => post_array('vars'),
    'headers_vars' => post_array('headers_vars'),
    'footers_vars' => post_array('footers_vars'),
    'remove_vars' => post_array('remove_vars'),
    'html_vars' => post_array('html_vars'),
    'table_vars' => post_array('table_vars'),
    'checkbox_vars' => post_array('checkbox_vars'),
]);

if ($converti_pdf) {
    $tmp = temp_docx($docx);
    try {
        $A_return['data_pdf'] = base64_encode(docx_to_pdf($tmp));
    } finally {
        temp_cleanup($tmp);
    }
    $A_return['pdf_engine'] = DocTemplate::$pdfEngine;
}

$A_return['esito'] = 'OK';
$A_return['errore'] = '';
$A_return['data'] = base64_encode($docx);
if (GDocx::$warnings || DocTemplate::$pdfWarnings) {
    $A_return['avvisi'] = array_values(array_merge(GDocx::$warnings, DocTemplate::$pdfWarnings));
}
json_response($A_return);
