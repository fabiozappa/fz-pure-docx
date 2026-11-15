<?php
/*
 * gdocx2pdf — Converte un DOCX in PDF.
 *
 * POST: apikey, nomehost, file_docx (base64)
 * Risposta JSON: esito, errore, tool, data_pdf (PDF base64)
 */
$tool = 'gdocx2pdf';
require __DIR__ . '/_start.php';
require __DIR__ . '/lib/DocTemplate.php';

$file_docx = post_base64('file_docx');
if ($file_docx === '') {
    json_error('Parametri POST non validi!');
}
if (!str_starts_with($file_docx, "PK")) {
    json_error('Il file non è un DOCX valido');
}

$tmp = temp_docx($file_docx);
try {
    $A_return['data_pdf'] = base64_encode(docx_to_pdf($tmp));
} finally {
    temp_cleanup($tmp);
}

$A_return['esito'] = 'OK';
$A_return['errore'] = '';
$A_return['pdf_engine'] = DocTemplate::$pdfEngine;
if (DocTemplate::$pdfWarnings) {
    $A_return['avvisi'] = DocTemplate::$pdfWarnings;
}
json_response($A_return);
