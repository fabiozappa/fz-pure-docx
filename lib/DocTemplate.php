<?php

declare(strict_types=1);

/**
 * DocTemplate — sostituzione di segnaposto $chiave$ in file DOCX.
 *
 * Requisiti: PHP 8.1+, estensioni zip e dom (più zlib per il PDF in puro PHP).
 * Nessuna dipendenza esterna.
 *
 * Uso:
 *   require 'DocTemplate.php';
 *
 *   DocTemplate::docx('modello.docx', ['nome' => 'Mario', 'data' => '19/10/2026'], [
 *       'output'   => 'save',            // 'save' | 'download' | 'string'
 *       'path'     => __DIR__ . '/out',  // cartella di destinazione (solo 'save')
 *       'filename' => 'contratto.docx',  // opzionale
 *   ]);
 *
 *   // PDF: usa LibreOffice se è installato, altrimenti il convertitore in puro PHP (DocxPdf.php)
 *   DocTemplate::pdf('modello.docx', $vars, ['output' => 'download']);
 *   echo DocTemplate::$pdfEngine;        // 'libreoffice' oppure 'php'
 *   print_r(DocTemplate::$pdfWarnings);  // avvisi del convertitore PHP (font sostituiti, ecc.)
 *
 * I segnaposto vengono sostituiti nel corpo, in tutti gli header e footer,
 * nelle note a piè di pagina e nelle note di chiusura, anche dentro tabelle
 * e caselle di testo. Il valore eredita la formattazione del primo carattere
 * del segnaposto. Un "\n" nel valore diventa un a capo (<w:br/>).
 */
final class DocTemplate
{
    private const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

    /** Caratteri ammessi nel nome della chiave tra i due $ */
    private const PLACEHOLDER_RE = '/\$([\p{L}\p{N}_.\-]+)\$/u';

    /**
     * Sostituisce i segnaposto e produce un DOCX.
     *
     * Opzioni:
     *  - output    'save' (default) | 'download' | 'string'
     *  - path      cartella di destinazione per 'save'
     *  - filename  nome del file risultante (default: nome modello + '_compilato.docx')
     *  - onMissing 'keep' (default, lascia $chiave$) | 'empty' (lo cancella) | 'throw'
     *
     * @return string|null percorso del file ('save'), contenuto binario ('string'), nulla ('download')
     */
    public static function docx(string $templatePath, array $vars, array $options = []): ?string
    {
        $tmp = self::fillToTempFile($templatePath, $vars, $options);
        try {
            $filename = self::outputFilename($templatePath, $options, 'docx');
            return self::emit(
                $tmp,
                $filename,
                'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                $options
            );
        } finally {
            @unlink($tmp);
        }
    }

    /** Motore usato dall'ultima chiamata a pdf(): 'libreoffice' oppure 'php' */
    public static ?string $pdfEngine = null;

    /** Avvisi dell'ultima conversione fatta dal motore PHP (font sostituiti, elementi ignorati) */
    public static array $pdfWarnings = [];

    /**
     * Come docx(), ma converte il risultato in PDF.
     *
     * Se sul server c'è LibreOffice lo usa (resa più completa); se il comando
     * non c'è, o la conversione fallisce, ripiega sul convertitore in puro PHP
     * di DocxPdf.php, che deve trovarsi nella stessa cartella di questo file.
     *
     * Opzioni aggiuntive:
     *  - engine   'auto' (default) | 'libreoffice' (solo LibreOffice, errore se manca) | 'php' (solo motore PHP)
     *  - soffice  percorso o nome dell'eseguibile di LibreOffice (default: ricerca automatica)
     *  - fonts    cartella, o elenco di cartelle, con i font .ttf per il motore PHP
     */
    public static function pdf(string $templatePath, array $vars, array $options = []): ?string
    {
        $engine = $options['engine'] ?? 'auto';
        if (!in_array($engine, ['auto', 'libreoffice', 'php'], true)) {
            throw new DocTemplateException("Valore non valido per engine: $engine");
        }
        self::$pdfEngine = null;
        self::$pdfWarnings = [];

        $tmpDocx = self::fillToTempFile($templatePath, $vars, $options);
        $workDir = self::makeTempDir();
        try {
            // Nome fisso in una cartella dedicata: evita collisioni tra conversioni simultanee
            $src = $workDir . '/documento.docx';
            rename($tmpDocx, $src);
            $pdf = $workDir . '/documento.pdf';

            $soffice = $engine === 'php' ? null : self::findSoffice($options['soffice'] ?? null);
            if ($soffice === null && $engine === 'libreoffice') {
                throw new DocTemplateException('LibreOffice non trovato: installalo o indica il percorso con l\'opzione "soffice"');
            }
            if ($soffice !== null) {
                $log = self::convertWithSoffice($soffice, $src, $workDir);
                if (is_file($pdf)) {
                    self::$pdfEngine = 'libreoffice';
                } elseif ($engine === 'libreoffice') {
                    throw new DocTemplateException("Conversione PDF fallita:\n" . $log);
                }
            }
            if (self::$pdfEngine === null) {
                self::convertWithPhp($src, $pdf, $templatePath, $options);
                self::$pdfEngine = 'php';
            }

            $filename = self::outputFilename($templatePath, $options, 'pdf');
            return self::emit($pdf, $filename, 'application/pdf', $options);
        } finally {
            @unlink($tmpDocx);
            self::removeDir($workDir);
        }
    }

    // ------------------------------------------------------------------
    // Conversione in PDF
    // ------------------------------------------------------------------

    /**
     * Cerca l'eseguibile di LibreOffice. Restituisce null se non c'è o se
     * PHP non può lanciare comandi (exec disabilitata).
     */
    private static function findSoffice(?string $custom): ?string
    {
        if (!function_exists('exec')) {
            return null;
        }
        $candidates = $custom !== null && $custom !== '' ? [$custom] : [
            'soffice',
            'libreoffice',
            '/Applications/LibreOffice.app/Contents/MacOS/soffice',
            '/opt/homebrew/bin/soffice',
            '/usr/local/bin/soffice',
            '/usr/bin/soffice',
            '/usr/bin/libreoffice',
            '/snap/bin/libreoffice',
            '/opt/libreoffice/program/soffice',
            'C:\\Program Files\\LibreOffice\\program\\soffice.exe',
        ];
        $windows = DIRECTORY_SEPARATOR === '\\';
        $pathDirs = array_filter(explode(PATH_SEPARATOR, (string) getenv('PATH')));
        foreach ($candidates as $cand) {
            if (str_contains($cand, '/') || str_contains($cand, '\\')) {
                if (is_file($cand) && is_executable($cand)) {
                    return $cand;
                }
                continue;
            }
            // Nome semplice: va cercato nelle cartelle del PATH
            foreach ($pathDirs as $dir) {
                foreach ($windows ? [$cand . '.exe', $cand . '.com', $cand] : [$cand] as $name) {
                    $full = rtrim($dir, '/\\') . DIRECTORY_SEPARATOR . $name;
                    if (is_file($full) && is_executable($full)) {
                        return $full;
                    }
                }
            }
        }
        return null;
    }

    /** Lancia LibreOffice headless; restituisce l'output del comando. */
    private static function convertWithSoffice(string $soffice, string $src, string $workDir): string
    {
        // Profilo LibreOffice separato per ogni chiamata, altrimenti due
        // conversioni in contemporanea si bloccano a vicenda
        $profile = 'file://' . $workDir . '/lo_profile';
        $cmd = sprintf(
            '%s%s -env:UserInstallation=%s --headless --convert-to pdf --outdir %s %s 2>&1',
            DIRECTORY_SEPARATOR === '\\' ? '' : 'HOME=' . escapeshellarg($workDir) . ' ',
            escapeshellarg($soffice),
            escapeshellarg($profile),
            escapeshellarg($workDir),
            escapeshellarg($src)
        );
        $out = [];
        exec($cmd, $out, $code);
        return implode("\n", $out);
    }

    /** Converte con il motore in puro PHP (DocxPdf.php, nella stessa cartella). */
    private static function convertWithPhp(string $src, string $pdf, string $templatePath, array $options): void
    {
        if (!class_exists('DocxPdf', false)) {
            $lib = __DIR__ . '/DocxPdf.php';
            if (!is_file($lib)) {
                throw new DocTemplateException('LibreOffice non è disponibile e manca DocxPdf.php accanto a DocTemplate.php');
            }
            require_once $lib;
        }
        $opt = ['title' => pathinfo(self::outputFilename($templatePath, $options, 'pdf'), PATHINFO_FILENAME)];
        foreach (['fonts', 'decimal'] as $k) {
            if (isset($options[$k])) {
                $opt[$k] = $options[$k];
            }
        }
        try {
            DocxPdf::convertToFile($src, $pdf, $opt);
        } catch (DocxPdfException $e) {
            throw new DocTemplateException('Conversione PDF fallita: ' . $e->getMessage(), 0, $e);
        }
        self::$pdfWarnings = DocxPdf::$warnings;
    }

    // ------------------------------------------------------------------
    // Sostituzione
    // ------------------------------------------------------------------

    /** Copia il modello in un file temporaneo e vi applica le sostituzioni. */
    private static function fillToTempFile(string $templatePath, array $vars, array $options): string
    {
        foreach (['zip' => 'ZipArchive', 'dom' => 'DOMDocument'] as $ext => $class) {
            if (!class_exists($class)) {
                throw new DocTemplateException("Estensione PHP '$ext' mancante");
            }
        }
        if (!is_file($templatePath) || !is_readable($templatePath)) {
            throw new DocTemplateException("Modello non trovato o non leggibile: $templatePath");
        }

        $onMissing = $options['onMissing'] ?? 'keep';
        if (!in_array($onMissing, ['keep', 'empty', 'throw'], true)) {
            throw new DocTemplateException("Valore non valido per onMissing: $onMissing");
        }

        $tmp = tempnam(sys_get_temp_dir(), 'doct_');
        if ($tmp === false || !copy($templatePath, $tmp)) {
            throw new DocTemplateException('Impossibile creare il file temporaneo');
        }

        $zip = new ZipArchive();
        if ($zip->open($tmp) !== true) {
            @unlink($tmp);
            throw new DocTemplateException("Il file non è un DOCX valido: $templatePath");
        }

        try {
            // Le parti che possono contenere testo visibile
            $parts = [];
            for ($i = 0; $i < $zip->numFiles; $i++) {
                $name = $zip->getNameIndex($i);
                if (preg_match('#^word/(document|header\d*|footer\d*|footnotes|endnotes)\.xml$#', $name)) {
                    $parts[] = $name;
                }
            }
            if (!in_array('word/document.xml', $parts, true)) {
                throw new DocTemplateException("Il file non è un DOCX valido: $templatePath");
            }

            foreach ($parts as $name) {
                $xml = $zip->getFromName($name);
                if ($xml === false || !str_contains($xml, '$')) {
                    continue; // nessun segnaposto possibile in questa parte
                }
                $new = self::replaceInXml($xml, $vars, $onMissing);
                if ($new !== $xml) {
                    // Sostituisce solo questa voce; il resto dell'archivio resta intatto
                    $zip->addFromString($name, $new);
                }
            }
            $zip->close();
        } catch (Throwable $e) {
            $zip->close();
            @unlink($tmp);
            throw $e;
        }

        return $tmp;
    }

    /**
     * Applica le sostituzioni a una parte XML (document, header, footer…).
     * Pubblico per chi deve lavorare su singole parti del DOCX.
     */
    public static function replaceInXml(string $xml, array $vars, string $onMissing = 'keep'): string
    {
        $dom = new DOMDocument();
        $dom->preserveWhiteSpace = true;
        if (!$dom->loadXML($xml, LIBXML_NONET | LIBXML_PARSEHUGE)) {
            throw new DocTemplateException('XML non valido nel DOCX');
        }

        return self::replaceInNode($dom->documentElement, $vars, $onMissing) ? $dom->saveXML() : $xml;
    }

    /**
     * Applica le sostituzioni ai paragrafi contenuti in un nodo DOM già
     * caricato (un'intera parte, una riga di tabella, una cella…).
     *
     * @return bool true se è stato modificato qualcosa
     */
    public static function replaceInNode(DOMNode $root, array $vars, string $onMissing = 'keep'): bool
    {
        if (!in_array($onMissing, ['keep', 'empty', 'throw'], true)) {
            throw new DocTemplateException("Valore non valido per onMissing: $onMissing");
        }
        $doc = $root instanceof DOMDocument ? $root : $root->ownerDocument;
        $xp = new DOMXPath($doc);
        $xp->registerNamespace('w', self::NS_W);

        // Raggruppa i nodi <w:t> per paragrafo di appartenenza (il <w:p> più vicino).
        // Così i paragrafi annidati (caselle di testo) vengono trattati a parte.
        $groups = [];
        foreach ($xp->query('.//w:t', $root) as $t) {
            $p = $t->parentNode;
            while ($p !== null && !($p->namespaceURI === self::NS_W && $p->localName === 'p')) {
                $p = $p->parentNode;
            }
            if ($p !== null) {
                $groups[spl_object_id($p)][] = $t;
            }
        }

        $changed = false;
        foreach ($groups as $nodes) {
            if (self::replaceInParagraph($nodes, $vars, $onMissing)) {
                $changed = true;
            }
        }
        return $changed;
    }

    /**
     * Cuore della libreria: Word spezza spesso il testo in più run
     * (<w:r><w:t>$no</w:t></w:r><w:r><w:t>me$</w:t></w:r>), quindi un
     * str_replace sull'XML non trova il segnaposto.
     *
     * Qui si ricostruisce il testo completo del paragrafo, si cercano i
     * segnaposto con i relativi offset e poi si riscrivono i singoli <w:t>:
     * il valore va nel nodo dove inizia il segnaposto (ereditandone la
     * formattazione), mentre i pezzi rimanenti negli altri nodi vengono tolti.
     *
     * @param DOMElement[] $nodes nodi <w:t> del paragrafo, in ordine di documento
     */
    private static function replaceInParagraph(array $nodes, array $vars, string $onMissing): bool
    {
        // Segmenti: testo di ogni nodo e posizione (in byte) nel testo complessivo
        $texts = [];
        $starts = [];
        $lengths = [];
        $full = '';
        foreach ($nodes as $i => $node) {
            $texts[$i] = $node->textContent;
            $starts[$i] = strlen($full);
            $lengths[$i] = strlen($texts[$i]);
            $full .= $texts[$i];
        }

        if (!str_contains($full, '$')
            || !preg_match_all(self::PLACEHOLDER_RE, $full, $matches, PREG_SET_ORDER | PREG_OFFSET_CAPTURE)) {
            return false;
        }

        $touched = [];

        // Dall'ultimo al primo: gli offset dei segnaposto precedenti restano validi
        foreach (array_reverse($matches) as $m) {
            [$token, $start] = $m[0];
            $key = $m[1][0];
            $end = $start + strlen($token); // esclusivo

            if (array_key_exists($key, $vars)) {
                $value = self::stringify($vars[$key], $key);
            } elseif ($onMissing === 'throw') {
                throw new DocTemplateException("Valore mancante per il segnaposto \$$key\$");
            } elseif ($onMissing === 'empty') {
                $value = '';
            } else {
                continue; // 'keep': lascia il segnaposto com'è
            }

            $first = self::segmentAt($starts, $lengths, $start);
            $last = self::segmentAt($starts, $lengths, $end - 1);

            $localStart = $start - $starts[$first];
            $localEnd = $end - $starts[$last];

            if ($first === $last) {
                $texts[$first] = substr($texts[$first], 0, $localStart)
                    . $value
                    . substr($texts[$first], $localEnd);
            } else {
                $texts[$first] = substr($texts[$first], 0, $localStart) . $value;
                for ($i = $first + 1; $i < $last; $i++) {
                    $texts[$i] = '';
                    $touched[$i] = true;
                }
                $texts[$last] = substr($texts[$last], $localEnd);
                $touched[$last] = true;
            }
            $touched[$first] = true;
        }

        if (!$touched) {
            return false;
        }

        foreach (array_keys($touched) as $i) {
            self::writeText($nodes[$i], $texts[$i]);
        }
        return true;
    }

    /** Indice del segmento che contiene il byte $pos (lunghezze originali). */
    private static function segmentAt(array $starts, array $lengths, int $pos): int
    {
        foreach ($starts as $i => $s) {
            if ($pos >= $s && $pos < $s + $lengths[$i]) {
                return $i;
            }
        }
        throw new DocTemplateException('Offset del segnaposto non valido');
    }

    /**
     * Scrive il testo nel nodo <w:t>. Se contiene a capo, divide il nodo in
     * più <w:t> separati da <w:br/> all'interno dello stesso run.
     * L'escape XML (&, <, >, ") lo fa DOM in automatico.
     */
    private static function writeText(DOMElement $t, string $text): void
    {
        $text = str_replace(["\r\n", "\r"], "\n", $text);
        $lines = explode("\n", $text);
        $doc = $t->ownerDocument;

        $t->textContent = $lines[0];
        $t->setAttributeNS('http://www.w3.org/XML/1998/namespace', 'xml:space', 'preserve');

        $after = $t;
        for ($i = 1, $n = count($lines); $i < $n; $i++) {
            $br = $doc->createElementNS(self::NS_W, 'w:br');
            $nt = $doc->createElementNS(self::NS_W, 'w:t');
            $nt->setAttributeNS('http://www.w3.org/XML/1998/namespace', 'xml:space', 'preserve');
            $nt->textContent = $lines[$i];

            $next = $after->nextSibling;
            $t->parentNode->insertBefore($br, $next);
            $t->parentNode->insertBefore($nt, $next);
            $after = $nt;
        }
    }

    private static function stringify(mixed $value, string $key): string
    {
        if ($value === null) {
            return '';
        }
        if (is_bool($value)) {
            return $value ? 'Sì' : 'No';
        }
        if (is_scalar($value) || $value instanceof Stringable) {
            $s = (string) $value;
            // Toglie i caratteri di controllo non ammessi in XML (tranne tab e a capo)
            return preg_replace('/[\x00-\x08\x0B\x0C\x0E-\x1F]/', '', $s) ?? '';
        }
        throw new DocTemplateException("Il valore per \$$key\$ non è convertibile in testo");
    }

    // ------------------------------------------------------------------
    // Output
    // ------------------------------------------------------------------

    private static function outputFilename(string $templatePath, array $options, string $ext): string
    {
        $name = $options['filename'] ?? pathinfo($templatePath, PATHINFO_FILENAME) . '_compilato.' . $ext;
        // Niente percorsi né caratteri problematici nel nome
        $name = basename(str_replace('\\', '/', (string) $name));
        $name = preg_replace('/[^\p{L}\p{N} ._\-()]/u', '_', $name) ?? 'documento';
        $name = ltrim($name, '.');
        if ($name === '') {
            $name = 'documento';
        }
        if (strtolower(pathinfo($name, PATHINFO_EXTENSION)) !== $ext) {
            $name .= '.' . $ext;
        }
        return $name;
    }

    private static function emit(string $file, string $filename, string $mime, array $options): ?string
    {
        $mode = $options['output'] ?? 'save';

        switch ($mode) {
            case 'string':
                $data = file_get_contents($file);
                if ($data === false) {
                    throw new DocTemplateException('Lettura del file generato fallita');
                }
                return $data;

            case 'save':
                $dir = $options['path'] ?? null;
                if ($dir === null || $dir === '') {
                    throw new DocTemplateException("Con output 'save' serve l'opzione 'path'");
                }
                if (!is_dir($dir) && !mkdir($dir, 0775, true) && !is_dir($dir)) {
                    throw new DocTemplateException("Impossibile creare la cartella: $dir");
                }
                if (!is_writable($dir)) {
                    throw new DocTemplateException("Cartella non scrivibile: $dir");
                }
                $dest = rtrim(realpath($dir), DIRECTORY_SEPARATOR) . DIRECTORY_SEPARATOR . $filename;
                if (!copy($file, $dest)) {
                    throw new DocTemplateException("Scrittura fallita: $dest");
                }
                return $dest;

            case 'download':
                if (headers_sent($f, $l)) {
                    throw new DocTemplateException("Output già inviato ($f:$l), impossibile scaricare il file");
                }
                while (ob_get_level() > 0) {
                    ob_end_clean();
                }
                $ascii = preg_replace('/[^A-Za-z0-9 ._\-()]/', '_', $filename);
                header('Content-Type: ' . $mime);
                header('Content-Disposition: attachment; filename="' . $ascii . '"; filename*=UTF-8\'\'' . rawurlencode($filename));
                header('Content-Length: ' . filesize($file));
                header('Cache-Control: private, no-cache, must-revalidate');
                readfile($file);
                return null;

            default:
                throw new DocTemplateException("Valore non valido per output: $mode");
        }
    }

    // ------------------------------------------------------------------
    // Utilità per i file temporanei
    // ------------------------------------------------------------------

    private static function makeTempDir(): string
    {
        $dir = sys_get_temp_dir() . '/doct_' . bin2hex(random_bytes(8));
        if (!mkdir($dir, 0700)) {
            throw new DocTemplateException('Impossibile creare la cartella temporanea');
        }
        return $dir;
    }

    private static function removeDir(string $dir): void
    {
        if (!is_dir($dir)) {
            return;
        }
        $it = new RecursiveIteratorIterator(
            new RecursiveDirectoryIterator($dir, FilesystemIterator::SKIP_DOTS),
            RecursiveIteratorIterator::CHILD_FIRST
        );
        foreach ($it as $f) {
            $f->isDir() && !$f->isLink() ? @rmdir($f->getPathname()) : @unlink($f->getPathname());
        }
        @rmdir($dir);
    }
}

final class DocTemplateException extends RuntimeException
{
}
