<?php
/*
 * Controllo della versione di PHP, incluso per primo da ogni strumento.
 * Questo file deve restare leggibile anche da PHP 5: niente sintassi moderna.
 * Senza questo controllo un PHP vecchio si ferma con una pagina bianca.
 */
if (!defined('PHP_VERSION_ID') || PHP_VERSION_ID < 80100) {
    $fzpd_msg = 'FZPureDOCX richiede PHP 8.1 o successivo: questo spazio web usa PHP ' . PHP_VERSION
        . '. Cambia la versione di PHP dal pannello di controllo dell\'hosting.';
    header('Cache-Control: no-store');
    if (isset($tool) && $tool === 'verifica') {
        header('Content-Type: text/html; charset=utf-8');
        echo '<!doctype html><html lang="it"><head><meta charset="utf-8">'
            . '<meta name="viewport" content="width=device-width, initial-scale=1">'
            . '<meta name="robots" content="noindex"><title>Verifica FZPureDOCX</title></head>'
            . '<body style="font:15px/1.5 sans-serif;margin:0;padding:24px 16px;background:#f5f6f8;color:#1d2330">'
            . '<div style="max-width:860px;margin:0 auto"><h1 style="font-size:22px;margin:0 0 4px">Verifica di FZPureDOCX</h1>'
            . '<div style="padding:12px 16px;border-radius:8px;margin:16px 0;font-weight:600;background:#fbe4e4;color:#8a1c1c">'
            . htmlspecialchars($fzpd_msg, ENT_QUOTES, 'UTF-8') . '</div></div></body></html>';
    } else {
        header('Content-Type: application/json; charset=utf-8');
        echo json_encode(array('esito' => 'KO', 'errore' => $fzpd_msg, 'tool' => isset($tool) ? $tool : ''));
    }
    exit();
}
