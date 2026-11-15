<?php
/*
 * Avvio comune degli strumenti: configurazione, verifica della chiave API,
 * risposta JSON. Ogni strumento imposta $tool prima di includere questo file.
 */
declare(strict_types=1);

date_default_timezone_set('Europe/Rome');

$CONFIG = require __DIR__ . '/_config.php';
$A_privato = is_file(__DIR__ . '/_apikeys.php') ? require __DIR__ . '/_apikeys.php' : null;
// Formato: ['ip_consentiti' => [...], 'chiavi' => [...]]; accettato anche il solo elenco delle chiavi
$A_apikeys = is_array($A_privato) ? ($A_privato['chiavi'] ?? $A_privato) : [];
$A_ip_consentiti = is_array($A_privato) ? ($A_privato['ip_consentiti'] ?? []) : [];

define('FZPD_DEBUG', ($CONFIG['debug'] ?? 'no') === 'si');
define('FZPD_T0', microtime(true));

// Niente messaggi PHP nell'output: rovinerebbero il JSON
ini_set('display_errors', '0');
error_reporting(FZPD_DEBUG ? E_ALL : E_ALL & ~E_DEPRECATED);
if (!empty($CONFIG['memory_limit'])) {
    ini_set('memory_limit', (string) $CONFIG['memory_limit']);
}
@set_time_limit(300);

$A_return = ['esito' => 'NN', 'errore' => 'Dati non presenti!', 'tool' => $tool];

/** Invia la risposta JSON e termina. */
function json_response(array $A_return): never
{
    global $tool;
    fzpd_log($A_return);
    while (ob_get_level() > 0) {
        ob_end_clean();
    }
    if (!headers_sent()) {
        header('Content-Type: application/json; charset=utf-8');
        header('Cache-Control: no-store');
    }
    $json = json_encode($A_return, JSON_UNESCAPED_SLASHES | JSON_INVALID_UTF8_SUBSTITUTE);
    print $json === false ? '{"esito":"KO","errore":"Errore di codifica JSON","tool":"' . $tool . '"}' : $json;
    exit();
}

function json_error(string $msg): never
{
    global $A_return;
    $A_return['esito'] = 'KO';
    $A_return['errore'] = $msg;
    unset($A_return['data'], $A_return['data_pdf']);
    json_response($A_return);
}

/** Una riga di registro: data, IP, strumento, host, esito, durata, errore. */
function fzpd_log(array $r): void
{
    global $CONFIG, $tool;
    $file = $CONFIG['log_file'] ?? null;
    if (!$file || !is_dir(dirname($file)) || !is_writable(dirname($file))) {
        return;
    }
    $host = preg_replace('/[^A-Za-z0-9._-]/', '_', substr((string) ($_POST['nomehost'] ?? ''), 0, 80));
    $line = sprintf("%s\t%s\t%s\t%s\t%s\t%.2fs\t%s\n",
        date('Y-m-d H:i:s'), $_SERVER['REMOTE_ADDR'] ?? '-', $tool, $host ?: '-', $r['esito'] ?? '-',
        microtime(true) - FZPD_T0, str_replace(["\n", "\t"], ' ', (string) ($r['errore'] ?? '')));
    @file_put_contents($file, $line, FILE_APPEND | LOCK_EX);
}

// Eccezioni ed errori fatali diventano risposte JSON con esito KO
set_exception_handler(static function (Throwable $e): void {
    // Niente percorsi del server nei messaggi: resta solo il nome del file
    $msg = preg_replace('#(?:[A-Za-z]:)?[\\\\/](?:[^\s\\\\/:]+[\\\\/])+#', '', $e->getMessage()) ?? $e->getMessage();
    if (FZPD_DEBUG) {
        $msg .= ' [' . basename($e->getFile()) . ':' . $e->getLine() . ']';
    }
    json_error($msg !== '' ? $msg : 'Errore interno');
});
register_shutdown_function(static function (): void {
    $e = error_get_last();
    if ($e !== null && in_array($e['type'], [E_ERROR, E_PARSE, E_CORE_ERROR, E_COMPILE_ERROR], true)) {
        json_error('Errore interno' . (FZPD_DEBUG ? ': ' . $e['message'] : ''));
    }
});

/** true se l'IP rientra nella regola: indirizzo singolo o rete CIDR (IPv4 o IPv6). */
function ip_in_regola(string $ip, string $regola): bool
{
    $regola = trim($regola);
    [$rete, $bit] = str_contains($regola, '/') ? explode('/', $regola, 2) : [$regola, null];
    $a = @inet_pton($ip);
    $b = @inet_pton($rete);
    if ($a === false || $b === false || strlen($a) !== strlen($b)) {
        return false;
    }
    $bit = $bit === null ? strlen($a) * 8 : max(0, min(strlen($a) * 8, (int) $bit));
    $byte = intdiv($bit, 8);
    if (substr($a, 0, $byte) !== substr($b, 0, $byte)) {
        return false;
    }
    if ($bit % 8 === 0) {
        return true;
    }
    $mask = (0xFF << (8 - $bit % 8)) & 0xFF;
    return (ord($a[$byte]) & $mask) === (ord($b[$byte]) & $mask);
}

if ($A_privato === null) {
    json_error('Server non configurato: copiare _apikeys.example.php come _apikeys.php e inserire le chiavi');
}

// Controllo IP (la pagina di verifica si apre dal browser, quindi ne è esclusa)
if ($A_ip_consentiti && $tool !== 'verifica') {
    $ip = (string) ($_SERVER['REMOTE_ADDR'] ?? '');
    $ok = false;
    foreach ($A_ip_consentiti as $regola) {
        if (is_string($regola) && ip_in_regola($ip, $regola)) {
            $ok = true;
            break;
        }
    }
    if (!$ok) {
        http_response_code(403);
        json_error('Indirizzo IP non autorizzato: ' . $ip);
    }
}

// Validazione chiave API (POST o GET)
$apikey = strtolower(trim((string) ($_POST['apikey'] ?? $_GET['apikey'] ?? '')));
if (!preg_match('/^[0-9a-f]{32}$/', $apikey)) {
    json_error('ApiKey errata!');
}
if (!isset($A_apikeys[$apikey])) {
    json_error('ApiKey non valida!');
}
if ($tool !== 'verifica' && !in_array($tool, $A_apikeys[$apikey]['tools'] ?? [], true)) {
    json_error('Tool non abilitato!');
}

/** nomehost ripulito (solo per registro e nomi file). */
function post_nomehost(): string
{
    return preg_replace('/[^A-Za-z0-9._-]/', '_', substr(trim((string) ($_POST['nomehost'] ?? '')), 0, 80)) ?? '';
}

/** Campo POST base64 decodificato, o stringa vuota. */
function post_base64(string $name): string
{
    $v = $_POST[$name] ?? '';
    if (!is_string($v) || $v === '') {
        return '';
    }
    // http_build_query codifica i "+" come %2B, ma alcuni client mandano spazi al loro posto
    $bin = base64_decode(strtr($v, ' ', '+'), true);
    return $bin === false ? '' : $bin;
}

/** Campo POST che deve essere un array (vars, html_vars…). */
function post_array(string $name): array
{
    $v = $_POST[$name] ?? [];
    return is_array($v) ? $v : [];
}

/** Converte un DOCX (file) in PDF con il motore configurato; restituisce il PDF. */
function docx_to_pdf(string $docxPath): string
{
    global $CONFIG;
    $opt = ['output' => 'string', 'engine' => $CONFIG['pdf_engine'] ?? 'auto', 'filename' => 'documento.pdf'];
    if (!empty($CONFIG['soffice'])) {
        $opt['soffice'] = $CONFIG['soffice'];
    }
    if (!empty($CONFIG['fonts'])) {
        $opt['fonts'] = $CONFIG['fonts'];
    }
    $pdf = DocTemplate::pdf($docxPath, [], $opt);
    if ($pdf === null || !str_starts_with($pdf, '%PDF')) {
        throw new RuntimeException('Conversione PDF fallita');
    }
    return $pdf;
}

/** Scrive dei byte in un file temporaneo .docx; restituisce il percorso. */
function temp_docx(string $bytes): string
{
    $dir = sys_get_temp_dir() . '/fzpd_' . bin2hex(random_bytes(8));
    if (!mkdir($dir, 0700)) {
        throw new RuntimeException('Impossibile creare la cartella temporanea');
    }
    $path = $dir . '/documento.docx';
    if (file_put_contents($path, $bytes) === false) {
        throw new RuntimeException('Impossibile scrivere il file temporaneo');
    }
    return $path;
}

function temp_cleanup(string $path): void
{
    @unlink($path);
    @rmdir(dirname($path));
}
