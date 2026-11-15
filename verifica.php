<?php
/*
 * verifica — Controllo dell'installazione dal browser.
 *
 *   https://indirizzo-del-server/verifica.php?apikey=LA_TUA_CHIAVE
 *
 * Controlla requisiti e configurazione, poi genera un documento di prova
 * (DOCX compilato e convertito in PDF). Con &pdf=1 scarica il PDF di prova.
 * Richiede una chiave valida; non è soggetta al controllo sugli IP, perché
 * si apre dal proprio browser e non dal server dell'applicazione client.
 */
$tool = 'verifica';
require __DIR__ . '/_start.php';
require __DIR__ . '/lib/DocTemplate.php';
require __DIR__ . '/lib/GDocx.php';

$checks = [];
/** Aggiunge un controllo: $stato 'ok', 'avviso' o 'errore'. */
$add = static function (string $nome, string $stato, string $dettaglio) use (&$checks): void {
    $checks[] = [$nome, $stato, $dettaglio];
};
$bytes = static function (string $v): int {
    $v = trim($v);
    $n = (int) $v;
    return match (strtolower(substr($v, -1))) { 'g' => $n << 30, 'm' => $n << 20, 'k' => $n << 10, default => $n };
};

// PHP ed estensioni
$add('Versione PHP', version_compare(PHP_VERSION, '8.1.0', '>=') ? 'ok' : 'errore', PHP_VERSION . ' (serve 8.1 o successiva)');
foreach (['zip' => 'lettura e scrittura dei DOCX', 'dom' => 'elaborazione XML', 'zlib' => 'compressione del PDF', 'mbstring' => 'testo UTF-8'] as $ext => $perche) {
    $add("Estensione $ext", extension_loaded($ext) ? 'ok' : 'errore', extension_loaded($ext) ? 'presente' : "mancante: serve per $perche");
}
foreach (['curl' => 'immagini HTML indicate con un indirizzo http/https', 'gd' => 'immagini GIF/WebP/BMP nell\'HTML'] as $ext => $perche) {
    $add("Estensione $ext", extension_loaded($ext) ? 'ok' : 'avviso', extension_loaded($ext) ? 'presente' : "mancante (facoltativa): serve per $perche");
}

// Limiti
$post = (string) ini_get('post_max_size');
$add('post_max_size', $bytes($post) >= 32 << 20 ? 'ok' : 'avviso', $post . ' (consigliato almeno 32M: i documenti viaggiano in base64)');
$vars = (int) ini_get('max_input_vars');
$add('max_input_vars', $vars >= 5000 ? 'ok' : 'avviso', $vars . ' (consigliato almeno 5000: tabelle con molte righe)');
$add('memory_limit', 'ok', (string) ini_get('memory_limit'));

// Cartelle
$tmp = sys_get_temp_dir();
$add('Cartella temporanea', is_writable($tmp) ? 'ok' : 'errore', $tmp . (is_writable($tmp) ? ' scrivibile' : ' NON scrivibile'));
$log = $CONFIG['log_file'] ?? null;
if ($log) {
    $w = is_dir(dirname($log)) && is_writable(dirname($log));
    $add('Registro chiamate', $w ? 'ok' : 'avviso', $w ? 'log/ scrivibile' : 'la cartella log/ non è scrivibile: le chiamate non vengono registrate');
} else {
    $add('Registro chiamate', 'ok', 'disattivato in _config.php');
}

// Chiavi e IP
$add('Chiavi API', 'ok', count($A_apikeys) . ' chiave/i configurata/e; quella usata è valida');
$tools = $A_apikeys[$apikey]['tools'] ?? [];
$add('Strumenti della chiave', in_array('gdocx', $tools, true) ? 'ok' : 'avviso', $tools ? implode(', ', $tools) : 'nessuno');
$add('IP consentiti', 'ok', $A_ip_consentiti
    ? count($A_ip_consentiti) . ' regole (il tuo IP attuale è ' . ($_SERVER['REMOTE_ADDR'] ?? '?') . ': serve quello del server che chiama il servizio)'
    : 'nessun limite: basta una chiave valida');

// Font per il convertitore PHP
$fonts = glob(__DIR__ . '/lib/fonts/*.{ttf,TTF,ttc,TTC}', GLOB_BRACE) ?: [];
$add('Font in lib/fonts', $fonts ? 'ok' : 'avviso', $fonts
    ? count($fonts) . ' file: ' . implode(', ', array_map(static fn ($f) => pathinfo($f, PATHINFO_FILENAME), $fonts))
    : 'nessun font: il convertitore PHP userà solo quelli di sistema, se ci sono (vedi README, sezione Font)');

// Prova completa: modello → DOCX compilato → PDF
$pdf = null;
$prova = null;
try {
    $W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
    $p = static fn (string $t) => '<w:p><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/></w:rPr><w:t xml:space="preserve">' . $t . '</w:t></w:r></w:p>';
    $doc = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="' . $W . '"><w:body>'
        . $p('Prova di installazione: $titolo$') . $p('$corpo$') . $p('Data: $data$')
        . '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1417" w:right="1134" w:bottom="1134" w:left="1134" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:body></w:document>';
    $zipPath = temp_docx('');
    $z = new ZipArchive();
    if ($z->open($zipPath, ZipArchive::OVERWRITE) !== true) {
        throw new RuntimeException('impossibile creare il modello di prova');
    }
    $z->addFromString('[Content_Types].xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
    $z->addFromString('_rels/.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
    $z->addFromString('word/document.xml', $doc);
    $z->close();
    $modello = (string) file_get_contents($zipPath);
    temp_cleanup($zipPath);

    $t0 = microtime(true);
    $docx = GDocx::fill($modello, [
        'vars' => ['titolo' => 'FZPureDOCX', 'data' => date('d/m/Y H:i')],
        'html_vars' => ['corpo' => '<p>Se leggi questo testo con la tabella qui sotto, la generazione funziona.</p>'
            . '<table style="width:100%;border-collapse:collapse"><tr><th style="background:#3170A7;color:#fff;border:1px solid #000;padding:4px">Controllo</th>'
            . '<th style="background:#3170A7;color:#fff;border:1px solid #000;padding:4px">Esito</th></tr>'
            . '<tr><td style="border:1px solid #000;padding:4px">Sostituzione segnaposto</td><td style="border:1px solid #000;padding:4px"><b>OK</b></td></tr>'
            . '<tr><td style="border:1px solid #000;padding:4px">HTML e tabelle</td><td style="border:1px solid #000;padding:4px"><b>OK</b></td></tr></table>'],
    ]);
    $tDocx = microtime(true) - $t0;
    $add('Generazione DOCX', 'ok', sprintf('riuscita in %.2f s', $tDocx));

    $path = temp_docx($docx);
    try {
        $t0 = microtime(true);
        $pdf = docx_to_pdf($path);
        $tPdf = microtime(true) - $t0;
    } finally {
        temp_cleanup($path);
    }
    $motore = DocTemplate::$pdfEngine === 'libreoffice' ? 'LibreOffice' : 'convertitore PHP';
    $nota = '';
    if (($CONFIG['pdf_engine'] ?? 'auto') === 'auto' && DocTemplate::$pdfEngine === 'php') {
        $nota = ' (LibreOffice non trovato: si usa il convertitore PHP)';
    }
    $add('Conversione PDF', 'ok', sprintf('riuscita in %.2f s con %s%s', $tPdf, $motore, $nota));
    foreach (DocTemplate::$pdfWarnings as $w) {
        $add('Avviso del convertitore', 'avviso', $w);
    }
} catch (Throwable $e) {
    $add('Prova di generazione', 'errore', $e->getMessage());
}

if (isset($_GET['pdf']) && $pdf !== null) {
    while (ob_get_level() > 0) {
        ob_end_clean();
    }
    header('Content-Type: application/pdf');
    header('Content-Disposition: inline; filename="prova_fzpuredocx.pdf"');
    print $pdf;
    exit();
}

// Pagina di esito
while (ob_get_level() > 0) {
    ob_end_clean();
}
header('Content-Type: text/html; charset=utf-8');
header('Cache-Control: no-store');
$errori = count(array_filter($checks, static fn ($c) => $c[1] === 'errore'));
$h = static fn (string $s) => htmlspecialchars($s, ENT_QUOTES, 'UTF-8');
$pdfLink = '?apikey=' . rawurlencode($apikey) . '&pdf=1';
?>
<!doctype html>
<html lang="it">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="robots" content="noindex">
<title>Verifica FZPureDOCX</title>
<style>
    body { font: 15px/1.5 system-ui, sans-serif; margin: 0; padding: 24px 16px; background: #f5f6f8; color: #1d2330; }
    main { max-width: 860px; margin: 0 auto; }
    h1 { font-size: 22px; margin: 0 0 4px; }
    .esito { padding: 12px 16px; border-radius: 8px; margin: 16px 0; font-weight: 600; }
    .esito.ok { background: #e3f4e8; color: #17602f; }
    .esito.ko { background: #fbe4e4; color: #8a1c1c; }
    table { width: 100%; border-collapse: collapse; background: #fff; border-radius: 8px; overflow: hidden; }
    td { padding: 8px 12px; border-bottom: 1px solid #e6e8ec; vertical-align: top; overflow-wrap: anywhere; }
    td:first-child { font-weight: 600; }
    .st { white-space: nowrap; }
    @media (max-width: 640px) {
        tr { display: block; border-bottom: 1px solid #e6e8ec; padding: 6px 0; }
        td { display: inline-block; border: 0; padding: 2px 12px; }
        td:last-child { display: block; }
    }
    .ok .st { color: #17602f; } .avviso .st { color: #9a6200; } .errore .st { color: #b42318; }
    a.btn { display: inline-block; margin-top: 16px; padding: 8px 14px; background: #3170A7; color: #fff; border-radius: 6px; text-decoration: none; }
</style>
</head>
<body>
<main>
    <h1>Verifica di FZPureDOCX</h1>
    <div>Pagina di controllo dell'installazione</div>
    <div class="esito <?= $errori ? 'ko' : 'ok' ?>">
        <?= $errori ? $errori . ' problema/i da risolvere (righe in rosso)' : 'Tutto pronto: il server può generare documenti DOCX e PDF.' ?>
    </div>
    <table>
        <?php foreach ($checks as [$nome, $stato, $dettaglio]): ?>
            <tr class="<?= $h($stato) ?>">
                <td><?= $h($nome) ?></td>
                <td class="st"><?= ['ok' => '✔ OK', 'avviso' => '⚠ Avviso', 'errore' => '✖ Errore'][$stato] ?></td>
                <td><?= $h($dettaglio) ?></td>
            </tr>
        <?php endforeach; ?>
    </table>
    <?php if ($pdf !== null): ?>
        <a class="btn" href="<?= $h($pdfLink) ?>">Apri il PDF di prova</a>
    <?php endif; ?>
</main>
</body>
</html>
<?php
exit();
