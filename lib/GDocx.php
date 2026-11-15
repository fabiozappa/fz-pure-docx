<?php

declare(strict_types=1);

/**
 * GDocx — compilazione di un modello DOCX: testo, HTML, tabelle ripetute,
 * caselle di controllo e rimozioni, costruita sopra DocTemplate.
 *
 *   $docx = GDocx::fill($bytesModello, [
 *       'vars'          => ['nome' => 'Mario'],                  // testo semplice, "\n" = a capo
 *       'headers_vars'  => [...],                                 // solo intestazioni
 *       'footers_vars'  => [...],                                 // solo piè di pagina
 *       'remove_vars'   => ['blocco_note'],                       // toglie il paragrafo del segnaposto
 *       'html_vars'     => ['corpo_pagina' => '<p>…</p>'],        // il paragrafo diventa l'HTML convertito
 *       'table_vars'    => ['prodotti' => [['codice' => 'A1', 'nome' => '…'], …]],  // $prodotti_codice$
 *       'checkbox_vars' => ['privacy' => 'si'],                   // caselle con tag $privacy$
 *   ]);
 *
 * Restituisce il contenuto binario del DOCX compilato. GDocx::$warnings
 * raccoglie gli avvisi non bloccanti (immagini non scaricabili, ecc.).
 */
final class GDocx
{
    public const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
    public const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
    public const NS_W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
    private const NS_REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
    private const NS_CT = 'http://schemas.openxmlformats.org/package/2006/content-types';

    /** Avvisi dell'ultima chiamata a fill() */
    public static array $warnings = [];

    private ZipArchive $zip;
    private DOMDocument $doc;
    private DOMXPath $xp;
    private ?DOMDocument $rels = null;
    private ?DOMDocument $types = null;
    private int $relSeq = 0;
    private int $mediaSeq = 0;
    private int $docPrSeq = 0;
    /** @var array<string, array> stili di paragrafo e carattere letti da styles.xml */
    private array $styles = [];
    private array $docDefaults = [];
    private ?string $defaultParaStyle = null;
    private array $themeFonts = [];

    private function __construct()
    {
    }

    public static function fill(string $templateBytes, array $data): string
    {
        self::$warnings = [];
        foreach (['zip' => 'ZipArchive', 'dom' => 'DOMDocument'] as $ext => $class) {
            if (!class_exists($class)) {
                throw new GDocxException("Estensione PHP '$ext' mancante");
            }
        }
        $tmp = tempnam(sys_get_temp_dir(), 'gdocx_');
        if ($tmp === false || file_put_contents($tmp, $templateBytes) === false) {
            throw new GDocxException('Impossibile creare il file temporaneo');
        }
        try {
            $g = new self();
            $g->run($tmp, $data);
            $out = file_get_contents($tmp);
            if ($out === false) {
                throw new GDocxException('Lettura del documento generato fallita');
            }
            return $out;
        } finally {
            @unlink($tmp);
        }
    }

    private function run(string $file, array $data): void
    {
        $this->zip = new ZipArchive();
        if ($this->zip->open($file) !== true || $this->zip->locateName('word/document.xml') === false) {
            throw new GDocxException('Il modello non è un file DOCX valido');
        }
        try {
            $vars = self::textVars($data['vars'] ?? []);
            $hVars = self::textVars($data['headers_vars'] ?? []) + $vars;
            $fVars = self::textVars($data['footers_vars'] ?? []) + $vars;

            // Intestazioni e piè di pagina: solo testo
            for ($i = 0; $i < $this->zip->numFiles; $i++) {
                $name = (string) $this->zip->getNameIndex($i);
                if (preg_match('#^word/(header|footer)\d*\.xml$#', $name, $m)) {
                    $xml = $this->zip->getFromName($name);
                    $use = $m[1] === 'header' ? $hVars : $fVars;
                    if ($xml !== false && $use && str_contains($xml, '$')) {
                        $new = DocTemplate::replaceInXml($xml, $use);
                        if ($new !== $xml) {
                            $this->zip->addFromString($name, $new);
                        }
                    }
                }
            }

            $this->doc = self::loadXml((string) $this->zip->getFromName('word/document.xml'));
            $this->xp = new DOMXPath($this->doc);
            $this->xp->registerNamespace('w', self::NS_W);
            $this->xp->registerNamespace('w14', self::NS_W14);
            $this->xp->registerNamespace('wp', 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing');

            // Ordine: testo, rimozioni, HTML, tabelle, checkbox
            if ($vars) {
                DocTemplate::replaceInNode($this->doc->documentElement, $vars);
            }
            foreach ((array) ($data['remove_vars'] ?? []) as $name) {
                if (is_scalar($name) && (string) $name !== '') {
                    foreach ($this->paragraphsWith((string) $name) as $p) {
                        $this->replaceParagraph($p, []);
                    }
                }
            }
            $html = $data['html_vars'] ?? [];
            if (is_array($html) && $html) {
                $this->htmlVars($html);
            }
            $tables = $data['table_vars'] ?? [];
            if (is_array($tables) && $tables) {
                $this->tableVars($tables);
            }
            $checks = $data['checkbox_vars'] ?? [];
            if (is_array($checks) && $checks) {
                $this->checkboxVars($checks);
            }

            $this->zip->addFromString('word/document.xml', $this->doc->saveXML());
            if ($this->rels !== null) {
                $this->zip->addFromString('word/_rels/document.xml.rels', $this->rels->saveXML());
            }
            if ($this->types !== null) {
                $this->zip->addFromString('[Content_Types].xml', $this->types->saveXML());
            }
            if (!$this->zip->close()) {
                throw new GDocxException('Scrittura del documento generato fallita');
            }
        } catch (Throwable $e) {
            try {
                $this->zip->close();
            } catch (Throwable) {
            }
            throw $e;
        }
    }

    /** Solo valori scalari, come stringhe; il resto viene scartato. */
    private static function textVars(mixed $vars): array
    {
        $out = [];
        if (is_array($vars)) {
            foreach ($vars as $k => $v) {
                if ($v === null || is_scalar($v)) {
                    $out[(string) $k] = (string) $v;
                }
            }
        }
        return $out;
    }

    public static function loadXml(string $xml): DOMDocument
    {
        $dom = new DOMDocument();
        $dom->preserveWhiteSpace = true;
        if ($xml === '' || !$dom->loadXML($xml, LIBXML_NONET | LIBXML_PARSEHUGE)) {
            throw new GDocxException('XML non valido nel DOCX');
        }
        return $dom;
    }

    // ------------------------------------------------------------------
    // Paragrafi e segnaposto
    // ------------------------------------------------------------------

    /** Testo di un paragrafo (esclusi i paragrafi annidati nelle caselle di testo). */
    private function paragraphText(DOMElement $p): string
    {
        $s = '';
        foreach ($this->xp->query('.//w:t', $p) as $t) {
            $q = $t->parentNode;
            while ($q !== null && !($q->namespaceURI === self::NS_W && $q->localName === 'p')) {
                $q = $q->parentNode;
            }
            if ($q === $p) {
                $s .= $t->textContent;
            }
        }
        return $s;
    }

    /** @return DOMElement[] paragraphs del corpo che contengono $name$ */
    private function paragraphsWith(string $name): array
    {
        $token = '$' . $name . '$';
        $out = [];
        foreach ($this->xp->query('//w:body//w:p') as $p) {
            if (str_contains($this->paragraphText($p), $token)) {
                $out[] = $p;
            }
        }
        return $out;
    }

    /**
     * Sostituisce un paragrafo con dei nodi a livello di blocco (paragrafi,
     * tabelle). Se il paragrafo porta un'interruzione di sezione o è l'unico
     * di una cella, ne resta una copia vuota per non rovinare il documento.
     *
     * @param DOMNode[] $nodes
     */
    private function replaceParagraph(DOMElement $p, array $nodes): void
    {
        $parent = $p->parentNode;
        foreach ($nodes as $n) {
            $parent->insertBefore($n, $p);
        }
        $pPr = null;
        foreach ($p->childNodes as $c) {
            if ($c instanceof DOMElement && $c->namespaceURI === self::NS_W && $c->localName === 'pPr') {
                $pPr = $c;
            }
        }
        $hasSect = $pPr !== null && $this->xp->query('w:sectPr', $pPr)->length > 0;
        $last = $nodes ? end($nodes) : null;
        $inCell = $parent instanceof DOMElement && $parent->localName === 'tc';
        $cellNeedsP = $inCell && ($last === null || $last->localName !== 'p');
        if ($hasSect || $cellNeedsP) {
            // Svuota il paragrafo tenendo solo le proprietà
            foreach (iterator_to_array($p->childNodes) as $c) {
                if ($c !== $pPr) {
                    $p->removeChild($c);
                }
            }
            return;
        }
        $parent->removeChild($p);
    }

    // ------------------------------------------------------------------
    // Variabili HTML
    // ------------------------------------------------------------------

    private function htmlVars(array $vars): void
    {
        $this->loadStyles();
        foreach ($vars as $name => $html) {
            if (!is_scalar($html) && $html !== null) {
                continue;
            }
            $html = trim((string) $html);
            foreach ($this->paragraphsWith((string) $name) as $p) {
                $nodes = [];
                if ($html !== '') {
                    $conv = new GDocxHtml($this, $this->baseStyle($p), $this->availableWidth($p));
                    $xml = $conv->convert($html);
                    $nodes = $this->fragment($xml);
                }
                $this->replaceParagraph($p, $nodes);
            }
        }
    }

    /** Converte una stringa di elementi WordprocessingML in nodi del documento. */
    private function fragment(string $xml): array
    {
        if ($xml === '') {
            return [];
        }
        $wrap = '<w:body xmlns:w="' . self::NS_W . '" xmlns:r="' . self::NS_R . '"'
            . ' xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"'
            . ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
            . ' xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">' . $xml . '</w:body>';
        $frag = self::loadXml($wrap);
        $out = [];
        foreach ($frag->documentElement->childNodes as $c) {
            if ($c instanceof DOMElement) {
                $out[] = $this->doc->importNode($c, true);
            }
        }
        return $out;
    }

    /** Larghezza utile (twip) nel punto del segnaposto: cella o pagina. */
    private function availableWidth(DOMElement $p): int
    {
        $tc = $this->xp->query('ancestor::w:tc[1]', $p)->item(0);
        if ($tc instanceof DOMElement) {
            $w = $this->xp->query('w:tcPr/w:tcW', $tc)->item(0);
            if ($w instanceof DOMElement && ($w->getAttributeNS(self::NS_W, 'type') ?: 'dxa') === 'dxa') {
                $tw = (int) $w->getAttributeNS(self::NS_W, 'w');
                if ($tw > 300) {
                    return $tw - 216;
                }
            }
        }
        $sect = $this->xp->query('following::w:sectPr[1]', $p)->item(0)
            ?? $this->xp->query('//w:body/w:sectPr')->item(0);
        $w = 11906;
        $l = $r = 1134;
        if ($sect instanceof DOMElement) {
            $sz = $this->xp->query('w:pgSz', $sect)->item(0);
            $mar = $this->xp->query('w:pgMar', $sect)->item(0);
            if ($sz instanceof DOMElement && $sz->hasAttributeNS(self::NS_W, 'w')) {
                $w = (int) $sz->getAttributeNS(self::NS_W, 'w');
            }
            if ($mar instanceof DOMElement) {
                $l = (int) $mar->getAttributeNS(self::NS_W, 'left') + (int) $mar->getAttributeNS(self::NS_W, 'gutter');
                $r = (int) $mar->getAttributeNS(self::NS_W, 'right');
            }
        }
        return max(1440, $w - $l - $r);
    }

    // ------------------------------------------------------------------
    // Stile di partenza per l'HTML: quello del segnaposto
    // ------------------------------------------------------------------

    private function loadStyles(): void
    {
        if ($this->styles || $this->docDefaults) {
            return;
        }
        $theme = $this->zip->getFromName('word/theme/theme1.xml');
        if ($theme !== false) {
            if (preg_match('#<a:majorFont>\s*<a:latin typeface="([^"]*)"#', $theme, $m)) {
                $this->themeFonts['major'] = $m[1];
            }
            if (preg_match('#<a:minorFont>\s*<a:latin typeface="([^"]*)"#', $theme, $m)) {
                $this->themeFonts['minor'] = $m[1];
            }
        }
        $xml = $this->zip->getFromName('word/styles.xml');
        if ($xml === false) {
            return;
        }
        try {
            $st = self::loadXml($xml);
        } catch (GDocxException) {
            return;
        }
        $xp = new DOMXPath($st);
        $xp->registerNamespace('w', self::NS_W);
        $def = $xp->query('/w:styles/w:docDefaults/w:rPrDefault/w:rPr')->item(0);
        $this->docDefaults = $def instanceof DOMElement ? $this->readRPr($def) : [];
        foreach ($xp->query('/w:styles/w:style') as $s) {
            /** @var DOMElement $s */
            $id = $s->getAttributeNS(self::NS_W, 'styleId');
            $based = $xp->query('w:basedOn', $s)->item(0);
            $rpr = $xp->query('w:rPr', $s)->item(0);
            $this->styles[$id] = [
                'based' => $based instanceof DOMElement ? $based->getAttributeNS(self::NS_W, 'val') : null,
                'r' => $rpr instanceof DOMElement ? $this->readRPr($rpr) : [],
            ];
            if ($s->getAttributeNS(self::NS_W, 'type') === 'paragraph'
                && in_array($s->getAttributeNS(self::NS_W, 'default'), ['1', 'true'], true)) {
                $this->defaultParaStyle = $id;
            }
        }
    }

    /** Proprietà di carattere utili all'HTML: font, dimensione, colore, grassetto, corsivo. */
    private function readRPr(DOMElement $rpr): array
    {
        $o = [];
        foreach ($rpr->childNodes as $c) {
            if (!$c instanceof DOMElement || $c->namespaceURI !== self::NS_W) {
                continue;
            }
            $val = $c->getAttributeNS(self::NS_W, 'val');
            switch ($c->localName) {
                case 'rFonts':
                    $f = $c->getAttributeNS(self::NS_W, 'ascii');
                    if ($f === '') {
                        $th = $c->getAttributeNS(self::NS_W, 'asciiTheme');
                        if ($th !== '') {
                            $f = $this->themeFonts[str_starts_with($th, 'major') ? 'major' : 'minor'] ?? '';
                        }
                    }
                    if ($f !== '') {
                        $o['ff'] = $f;
                    }
                    break;
                case 'sz':
                    if (is_numeric($val)) {
                        $o['sz'] = (float) $val / 2;
                    }
                    break;
                case 'color':
                    if (preg_match('/^[0-9A-Fa-f]{6}$/', $val)) {
                        $o['color'] = strtoupper($val);
                    }
                    break;
                case 'b':
                case 'i':
                    $o[$c->localName] = !in_array($val, ['0', 'false', 'off'], true);
                    break;
            }
        }
        return $o;
    }

    private function styleChain(?string $id, int $depth = 0): array
    {
        if ($id === null || $id === '' || $depth > 20 || !isset($this->styles[$id])) {
            return [];
        }
        return $this->styles[$id]['r'] + $this->styleChain($this->styles[$id]['based'], $depth + 1);
    }

    /** Stile ereditato dall'HTML: font del segnaposto, allineamento del suo paragrafo. */
    private function baseStyle(DOMElement $p): array
    {
        $pStyle = $this->xp->query('w:pPr/w:pStyle', $p)->item(0);
        $pid = $pStyle instanceof DOMElement ? $pStyle->getAttributeNS(self::NS_W, 'val') : $this->defaultParaStyle;
        $r = [];
        foreach ($this->xp->query('.//w:r[w:t]', $p) as $run) {
            if (str_contains($run->textContent, '$')) {
                $rpr = $this->xp->query('w:rPr', $run)->item(0);
                if ($rpr instanceof DOMElement) {
                    $rs = $this->xp->query('w:rStyle', $rpr)->item(0);
                    $r = $this->readRPr($rpr)
                        + ($rs instanceof DOMElement ? $this->styleChain($rs->getAttributeNS(self::NS_W, 'val')) : []);
                }
                break;
            }
        }
        $r += $this->styleChain($pid) + $this->docDefaults;
        $jc = $this->xp->query('w:pPr/w:jc', $p)->item(0);
        $align = 'left';
        if ($jc instanceof DOMElement) {
            $align = match ($jc->getAttributeNS(self::NS_W, 'val')) {
                'center' => 'center', 'right', 'end' => 'right', 'both', 'distribute' => 'justify', default => 'left',
            };
        }
        return [
            'ff' => $r['ff'] ?? null,
            'sz' => $r['sz'] ?? 10.0,
            'color' => $r['color'] ?? null,
            'b' => $r['b'] ?? false,
            'i' => $r['i'] ?? false,
            'align' => $align,
        ];
    }

    // ------------------------------------------------------------------
    // Immagini e collegamenti (usati dal convertitore HTML)
    // ------------------------------------------------------------------

    private function relsDoc(): DOMDocument
    {
        if ($this->rels === null) {
            $xml = $this->zip->getFromName('word/_rels/document.xml.rels');
            $this->rels = $xml !== false ? self::loadXml($xml)
                : self::loadXml('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="' . self::NS_REL . '"/>');
        }
        return $this->rels;
    }

    private function addRel(string $type, string $target, bool $external): string
    {
        $rels = $this->relsDoc();
        $ids = [];
        foreach ($rels->documentElement->childNodes as $c) {
            if ($c instanceof DOMElement) {
                $ids[$c->getAttribute('Id')] = true;
            }
        }
        do {
            $id = 'rIdGd' . (++$this->relSeq);
        } while (isset($ids[$id]));
        $el = $rels->createElementNS(self::NS_REL, 'Relationship');
        $el->setAttribute('Id', $id);
        $el->setAttribute('Type', $type);
        $el->setAttribute('Target', $target);
        if ($external) {
            $el->setAttribute('TargetMode', 'External');
        }
        $rels->documentElement->appendChild($el);
        return $id;
    }

    public function addHyperlink(string $url): string
    {
        return $this->addRel('http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink', $url, true);
    }

    /** Aggiunge un'immagine PNG/JPEG al pacchetto; restituisce l'id della relazione. */
    public function addImage(string $bytes, string $ext): string
    {
        do {
            $name = 'gdocx_img' . (++$this->mediaSeq) . '.' . $ext;
        } while ($this->zip->locateName('word/media/' . $name) !== false);
        $this->zip->addFromString('word/media/' . $name, $bytes);
        $this->ensureContentType($ext, $ext === 'png' ? 'image/png' : 'image/jpeg');
        return $this->addRel('http://schemas.openxmlformats.org/officeDocument/2006/relationships/image', 'media/' . $name, false);
    }

    private function ensureContentType(string $ext, string $mime): void
    {
        if ($this->types === null) {
            $this->types = self::loadXml((string) $this->zip->getFromName('[Content_Types].xml'));
        }
        foreach ($this->types->documentElement->childNodes as $c) {
            if ($c instanceof DOMElement && $c->localName === 'Default' && strtolower($c->getAttribute('Extension')) === $ext) {
                return;
            }
        }
        $el = $this->types->createElementNS(self::NS_CT, 'Default');
        $el->setAttribute('Extension', $ext);
        $el->setAttribute('ContentType', $mime);
        $this->types->documentElement->insertBefore($el, $this->types->documentElement->firstChild);
    }

    /** Id univoco per wp:docPr. */
    public function nextDocPrId(): int
    {
        if ($this->docPrSeq === 0) {
            $max = 0;
            foreach ($this->xp->query('//wp:docPr') as $d) {
                $max = max($max, (int) $d->getAttribute('id'));
            }
            $this->docPrSeq = max($max, 1000);
        }
        return ++$this->docPrSeq;
    }

    public static function warn(string $msg): void
    {
        if (!in_array($msg, self::$warnings, true)) {
            self::$warnings[] = $msg;
        }
    }

    // ------------------------------------------------------------------
    // Tabelle con righe ripetute
    // ------------------------------------------------------------------

    /**
     * table_vars[nome] = [ [col => valore, …], … ]
     * Nel modello i segnaposto sono $nome_col$: la riga (o le righe) che li
     * contiene viene ripetuta per ogni elemento dell'elenco.
     */
    private function tableVars(array $tables): void
    {
        foreach ($tables as $name => $rows) {
            if (!is_array($rows) || !$rows) {
                continue;
            }
            $keys = [];
            $data = [];
            foreach ($rows as $row) {
                if (!is_array($row)) {
                    continue;
                }
                $r = [];
                foreach ($row as $col => $val) {
                    if ($val === null || is_scalar($val)) {
                        $key = $name . '_' . $col;
                        $r[$key] = (string) $val;
                        $keys[$key] = '';
                    }
                }
                $data[] = $r;
            }
            if (!$keys) {
                continue;
            }
            // Righe del modello che contengono almeno un segnaposto della tabella
            $groups = [];
            foreach ($this->xp->query('//w:body//w:tr') as $tr) {
                $text = '';
                foreach ($this->xp->query('.//w:t', $tr) as $t) {
                    $text .= $t->textContent;
                }
                foreach (array_keys($keys) as $k) {
                    if (str_contains($text, '$' . $k . '$')) {
                        // Solo le righe della tabella più interna che contiene il segnaposto
                        $nested = false;
                        foreach ($this->xp->query('.//w:tr', $tr) as $inner) {
                            if (str_contains($inner->textContent, '$' . $k . '$')) {
                                $nested = true;
                                break;
                            }
                        }
                        if (!$nested) {
                            $groups[spl_object_id($tr->parentNode)][] = $tr;
                        }
                        break;
                    }
                }
            }
            foreach ($groups as $tpl) {
                $first = $tpl[0];
                $parent = $first->parentNode;
                foreach ($data as $r) {
                    foreach ($tpl as $tr) {
                        $clone = $tr->cloneNode(true);
                        DocTemplate::replaceInNode($clone, $r + $keys);
                        $parent->insertBefore($clone, $first);
                    }
                }
                foreach ($tpl as $tr) {
                    $parent->removeChild($tr);
                }
            }
        }
    }

    // ------------------------------------------------------------------
    // Checkbox
    // ------------------------------------------------------------------

    /**
     * Caselle di controllo di Word con tag (o testo di stato, per quelle
     * legacy) uguale a $nome$. Spuntate con 1, si, s, yes, y, true.
     */
    private function checkboxVars(array $vars): void
    {
        foreach ($vars as $name => $value) {
            $v = is_bool($value) ? $value : in_array(strtolower(trim((string) (is_scalar($value) ? $value : ''))),
                ['1', 'si', 'sì', 's', 'yes', 'y', 'true', 'on', 'x'], true);
            $token = '$' . $name . '$';
            $lit = self::xpathLiteral($token);
            // Caselle legacy (campi modulo)
            foreach ($this->xp->query('//w:ffData[w:statusText[@w:val=' . $lit . ']]') as $ff) {
                foreach ($this->xp->query('w:checkBox', $ff) as $cb) {
                    $def = $this->xp->query('w:default', $cb)->item(0);
                    if ($def instanceof DOMElement) {
                        $def->setAttributeNS(self::NS_W, 'w:val', $v ? '1' : '0');
                    }
                    $chk = $this->xp->query('w:checked', $cb)->item(0);
                    if ($chk instanceof DOMElement) {
                        $chk->setAttributeNS(self::NS_W, 'w:val', $v ? '1' : '0');
                    }
                }
            }
            // Caselle Word 2010+ (controlli contenuto)
            foreach ($this->xp->query('//w:sdtPr[w:tag[@w:val=' . $lit . ']]') as $pr) {
                $chk = $this->xp->query('w14:checkbox/w14:checked', $pr)->item(0);
                if (!$chk instanceof DOMElement) {
                    continue;
                }
                $chk->setAttributeNS(self::NS_W14, 'w14:val', $v ? '1' : '0');
                $t = $this->xp->query('../w:sdtContent//w:t', $pr)->item(0);
                if ($t instanceof DOMElement) {
                    $t->nodeValue = $v ? '☒' : '☐';
                }
            }
        }
    }

    private static function xpathLiteral(string $s): string
    {
        if (!str_contains($s, '"')) {
            return '"' . $s . '"';
        }
        if (!str_contains($s, "'")) {
            return "'" . $s . "'";
        }
        return 'concat("' . str_replace('"', '", \'"\', "', $s) . '")';
    }
}

final class GDocxException extends RuntimeException
{
}

/**
 * Fogli di stile CSS ridotti: selettori di tipo, classe e id, combinatori
 * discendente e figlio. Le regole con pseudo-classi o attributi vengono
 * ignorate. Le proprietà abbreviate (margin, border, background…) sono
 * espanse in fase di lettura.
 */
final class GDocxCss
{
    /** @var array<int, array{0: array, 1: int, 2: int, 3: array}> parti, specificità, ordine, dichiarazioni */
    private array $rules = [];
    private int $seq = 0;

    public function add(string $css): void
    {
        $css = preg_replace('#/\*.*?\*/#s', '', $css) ?? '';
        $css = str_replace(['<!--', '-->'], '', $css);
        $css = self::stripAtRules($css);
        if (!preg_match_all('/([^{}]+)\{([^{}]*)\}/', $css, $mm, PREG_SET_ORDER)) {
            return;
        }
        foreach ($mm as $m) {
            $decl = self::parseDecl($m[2]);
            if (!$decl) {
                continue;
            }
            $order = ++$this->seq;
            foreach (explode(',', $m[1]) as $sel) {
                $p = self::parseSelector($sel);
                if ($p !== null) {
                    $this->rules[] = [$p[0], $p[1], $order, $decl];
                }
            }
        }
    }

    /** Dichiarazioni che si applicano all'elemento, già in ordine di cascata. */
    public function match(DOMElement $el): array
    {
        $hit = [];
        foreach ($this->rules as $r) {
            if (self::matches($r[0], $el)) {
                $hit[] = $r;
            }
        }
        usort($hit, static fn ($a, $b) => [$a[1], $a[2]] <=> [$b[1], $b[2]]);
        return array_map(static fn ($r) => $r[3], $hit);
    }

    private static function stripAtRules(string $css): string
    {
        $out = '';
        $n = strlen($css);
        for ($i = 0; $i < $n; $i++) {
            if ($css[$i] !== '@') {
                $out .= $css[$i];
                continue;
            }
            $semi = strpos($css, ';', $i);
            $brace = strpos($css, '{', $i);
            if ($brace === false || ($semi !== false && $semi < $brace)) {
                $i = $semi === false ? $n : $semi;
                continue;
            }
            $depth = 0;
            for ($j = $brace; $j < $n; $j++) {
                if ($css[$j] === '{') {
                    $depth++;
                } elseif ($css[$j] === '}' && --$depth === 0) {
                    break;
                }
            }
            $i = $j;
        }
        return $out;
    }

    private static function parseSelector(string $s): ?array
    {
        $s = trim(preg_replace('/\s*>\s*/', ' > ', trim($s)) ?? '');
        if ($s === '') {
            return null;
        }
        $parts = [];
        $spec = 0;
        $comb = ' ';
        foreach (preg_split('/\s+/', $s) as $tok) {
            if ($tok === '>') {
                $comb = '>';
                continue;
            }
            if (!preg_match('/^(\*|[a-z][a-z0-9-]*)?((?:[.#][\w-]+)*)$/i', $tok, $m) || ($m[1] === '' && $m[2] === '')) {
                return null; // pseudo-classi, attributi, combinatori + e ~: non supportati
            }
            $tag = strtolower($m[1]);
            $id = null;
            $cls = [];
            preg_match_all('/([.#])([\w-]+)/', $m[2], $sm, PREG_SET_ORDER);
            foreach ($sm as $x) {
                if ($x[1] === '#') {
                    $id = $x[2];
                    $spec += 100;
                } else {
                    $cls[] = $x[2];
                    $spec += 10;
                }
            }
            if ($tag !== '' && $tag !== '*') {
                $spec += 1;
            }
            $parts[] = ['tag' => $tag === '*' ? '' : $tag, 'id' => $id, 'cls' => $cls, 'comb' => $comb];
            $comb = ' ';
        }
        return $parts ? [$parts, $spec] : null;
    }

    private static function simple(array $p, DOMElement $el): bool
    {
        if ($p['tag'] !== '' && strtolower($el->nodeName) !== $p['tag']) {
            return false;
        }
        if ($p['id'] !== null && $el->getAttribute('id') !== $p['id']) {
            return false;
        }
        if ($p['cls']) {
            $have = preg_split('/\s+/', trim($el->getAttribute('class'))) ?: [];
            foreach ($p['cls'] as $c) {
                if (!in_array($c, $have, true)) {
                    return false;
                }
            }
        }
        return true;
    }

    private static function matches(array $parts, DOMElement $el): bool
    {
        $i = count($parts) - 1;
        return self::simple($parts[$i], $el) && self::ancestors($parts, $i, $el);
    }

    private static function ancestors(array $parts, int $i, DOMElement $el): bool
    {
        if ($i === 0) {
            return true;
        }
        $child = $parts[$i]['comb'] === '>';
        for ($p = $el->parentNode; $p instanceof DOMElement; $p = $p->parentNode) {
            if (self::simple($parts[$i - 1], $p) && self::ancestors($parts, $i - 1, $p)) {
                return true;
            }
            if ($child) {
                return false;
            }
        }
        return false;
    }

    /**
     * Legge un blocco di dichiarazioni ("color: red; margin: 0 4px").
     * @return array<string, array{0: string, 1: bool}> proprietà => [valore, !important]
     */
    public static function parseDecl(string $s): array
    {
        $out = [];
        foreach (preg_split('/;(?![^(]*\))/', $s) as $d) {
            $c = strpos($d, ':');
            if ($c === false) {
                continue;
            }
            $prop = strtolower(trim(substr($d, 0, $c)));
            $val = trim(substr($d, $c + 1));
            $imp = false;
            if (preg_match('/\s*!\s*important\s*$/i', $val, $m)) {
                $imp = true;
                $val = trim(substr($val, 0, -strlen($m[0])));
            }
            if ($prop === '' || $val === '') {
                continue;
            }
            // Niente spazi dentro le parentesi: rgb(1, 2, 3) resta un solo valore
            $val = preg_replace_callback('/\([^)]*\)/', static fn ($m) => preg_replace('/\s+/', '', $m[0]), $val) ?? $val;
            foreach (self::expand($prop, $val) as $p => $v) {
                $out[$p] = [$v, $imp];
            }
        }
        return $out;
    }

    private const SIDES = ['top', 'right', 'bottom', 'left'];
    private const BORDER_STYLES = ['none', 'hidden', 'dotted', 'dashed', 'solid', 'double', 'groove', 'ridge', 'inset', 'outset'];

    private static function box(string $v): array
    {
        $t = preg_split('/\s+/', trim($v));
        $t = match (count($t)) {
            1 => [$t[0], $t[0], $t[0], $t[0]],
            2 => [$t[0], $t[1], $t[0], $t[1]],
            3 => [$t[0], $t[1], $t[2], $t[1]],
            default => [$t[0], $t[1], $t[2], $t[3]],
        };
        return array_combine(self::SIDES, $t);
    }

    private static function expand(string $prop, string $val): array
    {
        $lv = strtolower($val);
        switch ($prop) {
            case 'margin':
            case 'padding':
                $o = [];
                foreach (self::box($lv) as $side => $x) {
                    $o["$prop-$side"] = $x;
                }
                return $o;
            case 'border-width':
            case 'border-style':
            case 'border-color':
                $what = substr($prop, 7);
                $o = [];
                foreach (self::box($lv) as $side => $x) {
                    $o["border-$side-$what"] = $x;
                }
                return $o;
            case 'border':
            case 'border-top':
            case 'border-right':
            case 'border-bottom':
            case 'border-left':
                $w = 'medium';
                $st = 'none';
                $col = 'currentcolor';
                foreach (preg_split('/\s+/', $lv) as $t) {
                    if (in_array($t, self::BORDER_STYLES, true)) {
                        $st = $t;
                    } elseif (in_array($t, ['thin', 'medium', 'thick'], true) || preg_match('/^[\d.]+[a-z%]*$/', $t)) {
                        $w = $t;
                    } else {
                        $col = $t;
                    }
                }
                $sides = $prop === 'border' ? self::SIDES : [substr($prop, 7)];
                $o = [];
                foreach ($sides as $s) {
                    $o["border-$s-width"] = $w;
                    $o["border-$s-style"] = $st;
                    $o["border-$s-color"] = $col;
                }
                return $o;
            case 'background':
                foreach (preg_split('/\s+/', $lv) as $t) {
                    if (GDocxHtml::color($t) !== null || $t === 'transparent' || $t === 'none') {
                        return ['background-color' => $t];
                    }
                }
                return [];
            case 'list-style':
                foreach (preg_split('/\s+/', $lv) as $t) {
                    if (!in_array($t, ['inside', 'outside'], true) && !str_starts_with($t, 'url(')) {
                        return ['list-style-type' => $t];
                    }
                }
                return [];
            case 'font':
                $o = [];
                if (preg_match('/^(.*?)([\d.]+(?:px|pt|em|rem|%)|xx-small|x-small|small|medium|large|x-large|xx-large|smaller|larger)(?:\/(\S+))?\s+(.+)$/', $val, $m)) {
                    foreach (preg_split('/\s+/', strtolower(trim($m[1]))) as $t) {
                        if (in_array($t, ['italic', 'oblique'], true)) {
                            $o['font-style'] = 'italic';
                        } elseif ($t === 'bold' || $t === 'bolder' || (is_numeric($t) && (int) $t >= 600)) {
                            $o['font-weight'] = 'bold';
                        } elseif ($t === 'small-caps') {
                            $o['font-variant'] = 'small-caps';
                        }
                    }
                    $o['font-size'] = strtolower($m[2]);
                    if (($m[3] ?? '') !== '') {
                        $o['line-height'] = strtolower($m[3]);
                    }
                    $o['font-family'] = $m[4];
                }
                return $o;
            case 'text-decoration-line':
                return ['text-decoration' => $lv];
            case 'font-family':
            case 'content':
                return [$prop => $val];
            default:
                return [$prop => $lv];
        }
    }
}

/**
 * Convertitore HTML → WordprocessingML per le variabili html_vars.
 *
 * Gestisce paragrafi e titoli, formattazione del testo (grassetto, corsivo,
 * sottolineato, barrato, colore, sfondo, dimensione, font, apice/pedice),
 * a capo, elenchi puntati e numerati (anche annidati), tabelle (larghezze,
 * bordi, sfondi, celle unite, righe di intestazione), immagini (data URI o
 * URL http/https), collegamenti, interruzioni di pagina (page-break-before).
 * Gli stili arrivano da <style> e dall'attributo style.
 */
final class GDocxHtml
{
    private const BLOCK = [
        'address', 'article', 'aside', 'blockquote', 'body', 'center', 'dd', 'div', 'dl', 'dt', 'fieldset',
        'figcaption', 'figure', 'footer', 'form', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'header', 'hr', 'html',
        'li', 'main', 'nav', 'ol', 'p', 'pre', 'section', 'table', 'ul', 'caption',
    ];
    private const SKIP = ['head', 'script', 'style', 'title', 'meta', 'link', 'noscript', 'template', 'object', 'iframe', 'svg', 'colgroup', 'col'];

    private const UA_CSS = <<<'CSS'
        b, strong { font-weight: bold }
        i, em, cite, var, dfn, address { font-style: italic }
        u, ins { text-decoration: underline }
        s, strike, del { text-decoration: line-through }
        sub { vertical-align: sub; font-size: smaller }
        sup { vertical-align: super; font-size: smaller }
        small { font-size: smaller }
        big { font-size: larger }
        code, kbd, samp, tt, pre { font-family: Courier New }
        pre { white-space: pre; margin: 1em 0 }
        mark { background-color: yellow }
        a { color: #0563C1; text-decoration: underline }
        p, dl { margin: 1em 0 }
        blockquote { margin: 1em 40px }
        ul, ol { margin: 1em 0; padding-left: 40px }
        ul ul, ul ol, ol ul, ol ol { margin: 0 }
        ol { list-style-type: decimal }
        ul { list-style-type: disc }
        ul ul { list-style-type: circle }
        ul ul ul { list-style-type: square }
        dd { margin-left: 40px }
        h1 { font-size: 2em; margin: .67em 0; font-weight: bold }
        h2 { font-size: 1.5em; margin: .83em 0; font-weight: bold }
        h3 { font-size: 1.17em; margin: 1em 0; font-weight: bold }
        h4 { font-size: 1em; margin: 1.33em 0; font-weight: bold }
        h5 { font-size: .83em; margin: 1.67em 0; font-weight: bold }
        h6 { font-size: .67em; margin: 2.33em 0; font-weight: bold }
        th { font-weight: bold; text-align: center }
        tr, thead, tbody, tfoot { vertical-align: middle }
        td, th { padding: 1px }
        center { text-align: center }
        hr { margin: .5em 0 }
        CSS;

    private const COLORS = [
        'black' => '000000', 'white' => 'FFFFFF', 'red' => 'FF0000', 'green' => '008000', 'blue' => '0000FF',
        'yellow' => 'FFFF00', 'orange' => 'FFA500', 'purple' => '800080', 'gray' => '808080', 'grey' => '808080',
        'silver' => 'C0C0C0', 'maroon' => '800000', 'olive' => '808000', 'lime' => '00FF00', 'aqua' => '00FFFF',
        'cyan' => '00FFFF', 'teal' => '008080', 'navy' => '000080', 'fuchsia' => 'FF00FF', 'magenta' => 'FF00FF',
        'pink' => 'FFC0CB', 'brown' => 'A52A2A', 'gold' => 'FFD700', 'darkgray' => 'A9A9A9', 'darkgrey' => 'A9A9A9',
        'lightgray' => 'D3D3D3', 'lightgrey' => 'D3D3D3', 'darkblue' => '00008B', 'darkred' => '8B0000',
        'darkgreen' => '006400', 'lightblue' => 'ADD8E6', 'lightgreen' => '90EE90', 'whitesmoke' => 'F5F5F5',
        'gainsboro' => 'DCDCDC', 'dimgray' => '696969', 'dimgrey' => '696969', 'steelblue' => '4682B4',
        'royalblue' => '4169E1', 'crimson' => 'DC143C', 'coral' => 'FF7F50', 'tomato' => 'FF6347',
        'beige' => 'F5F5DC', 'ivory' => 'FFFFF0', 'khaki' => 'F0E68C', 'lavender' => 'E6E6FA', 'salmon' => 'FA8072',
        'indigo' => '4B0082', 'violet' => 'EE82EE', 'skyblue' => '87CEEB', 'darkorange' => 'FF8C00',
    ];

    private GDocxCss $ua;
    private GDocxCss $css;
    private float $rootPt;
    /** Segno dell'elemento di elenco in attesa del suo primo paragrafo */
    private ?array $marker = null;
    private bool $pendingBreak = false;

    public function __construct(private GDocx $pkg, private array $base, private int $width)
    {
        $this->ua = new GDocxCss();
        $this->ua->add(self::UA_CSS);
        $this->css = new GDocxCss();
        $this->rootPt = (float) $base['sz'];
    }

    /** @return string elementi w:p / w:tbl */
    public function convert(string $html): string
    {
        $dom = new DOMDocument();
        $prev = libxml_use_internal_errors(true);
        $dom->loadHTML('<?xml encoding="UTF-8"><html><body>' . $html . '</body></html>',
            LIBXML_NONET | LIBXML_HTML_NODEFDTD | LIBXML_PARSEHUGE | LIBXML_COMPACT);
        libxml_clear_errors();
        libxml_use_internal_errors($prev);

        foreach ($dom->getElementsByTagName('style') as $st) {
            $this->css->add($st->textContent);
        }
        $body = $dom->getElementsByTagName('body')->item(0);
        if (!$body instanceof DOMElement) {
            return '';
        }
        $root = [
            'ff' => $this->base['ff'], 'sz' => (float) $this->base['sz'], 'color' => $this->base['color'],
            'b' => (bool) $this->base['b'], 'i' => (bool) $this->base['i'], 'u' => false, 'strike' => false,
            'bg' => null, 'va' => 'baseline', 'caps' => false, 'tt' => 'none', 'align' => $this->base['align'],
            'lh' => null, 'ws' => 'normal', 'ls' => 0, 'indent' => 0, 'lst' => 'disc', 'link' => null, 'd' => [],
            'author' => [],
        ];
        $html = $body->parentNode instanceof DOMElement ? $this->compute($body->parentNode, $root) : $root;
        $cs = $this->compute($body, $html);
        $bc = ['indL' => 0, 'indR' => 0, 'width' => $this->width, 'shd' => null, 'bdr' => null, 'list' => []];
        $items = $this->blockContent($body, $cs, $bc);
        $xml = $this->serialize($items);
        // Una tabella a contatto con un'altra tabella del modello verrebbe fusa con essa
        if ($items && $items[0]['k'] === 'raw' && !$items[0]['break']) {
            $xml = self::SEPARATOR . $xml;
        }
        if ($items && end($items)['k'] === 'raw') {
            $xml .= self::SEPARATOR;
        }
        return $xml;
    }

    /** Paragrafo quasi invisibile (1 pt): separa tabelle che Word e LibreOffice fonderebbero. */
    private const SEPARATOR = '<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="20" w:lineRule="exact"/>'
        . '<w:rPr><w:sz w:val="2"/><w:szCs w:val="2"/></w:rPr></w:pPr></w:p>';

    // ------------------------------------------------------------------
    // Stile calcolato
    // ------------------------------------------------------------------

    private function compute(DOMElement $el, array $parent): array
    {
        $tag = strtolower($el->nodeName);
        $d = [];
        $author = [];
        foreach ($this->ua->match($el) as $decl) {
            foreach ($decl as $p => $v) {
                $d[$p] = $v[0];
            }
        }
        foreach ($this->presentational($el, $tag) as $p => $v) {
            $d[$p] = $v;
        }
        $important = [];
        $layers = $this->css->match($el);
        if ($el->hasAttribute('style')) {
            $layers[] = GDocxCss::parseDecl($el->getAttribute('style'));
        }
        foreach ($layers as $decl) {
            foreach ($decl as $p => $v) {
                if ($v[1]) {
                    $important[$p] = $v[0];
                } else {
                    $d[$p] = $v[0];
                }
                $author[$p] = true;
            }
        }
        foreach ($important as $p => $v) {
            $d[$p] = $v;
        }

        $cs = $parent;
        $cs['d'] = $d;
        $cs['author'] = $author;
        // Lo sfondo del testo vale solo dentro gli elementi in linea
        $boxed = $this->isBlock($el, $d) || in_array($tag, ['td', 'th', 'tr', 'thead', 'tbody', 'tfoot'], true);
        if ($boxed) {
            $cs['bg'] = null;
        }
        $pt = $parent['sz'];
        if (isset($d['font-size'])) {
            $cs['sz'] = $this->fontSize($d['font-size'], $pt);
        }
        if (isset($d['font-family'])) {
            $cs['ff'] = self::fontFamily($d['font-family']) ?? $cs['ff'];
        }
        if (isset($d['color'])) {
            $c = self::color($d['color']);
            if ($c !== null) {
                $cs['color'] = $c;
            }
        }
        if (isset($d['font-weight'])) {
            $w = $d['font-weight'];
            $cs['b'] = $w === 'bold' || $w === 'bolder' || (is_numeric($w) && (int) $w >= 600);
        }
        if (isset($d['font-style'])) {
            $cs['i'] = in_array($d['font-style'], ['italic', 'oblique'], true);
        }
        if (isset($d['text-decoration'])) {
            $td = $d['text-decoration'];
            if (str_contains($td, 'underline')) {
                $cs['u'] = true;
            }
            if (str_contains($td, 'line-through')) {
                $cs['strike'] = true;
            }
            if ($td === 'none' && isset($author['text-decoration'])) {
                // Uso comune (a { text-decoration: none }): toglie la sottolineatura del collegamento
                $cs['u'] = $cs['strike'] = false;
            }
        }
        if (isset($d['background-color']) && !$boxed) {
            $c = self::color($d['background-color']);
            if ($c !== null) {
                $cs['bg'] = $c;
            }
        }
        if (isset($d['vertical-align']) && in_array($d['vertical-align'], ['super', 'sub'], true)) {
            $cs['va'] = $d['vertical-align'];
        }
        if (isset($d['font-variant'])) {
            $cs['caps'] = $d['font-variant'] === 'small-caps';
        }
        if (isset($d['text-transform'])) {
            $cs['tt'] = $d['text-transform'];
        }
        if (isset($d['text-align'])) {
            $a = $d['text-align'];
            $cs['align'] = match ($a) {
                'center', '-webkit-center', 'middle' => 'center', 'right', 'end' => 'right', 'justify' => 'justify', default => 'left',
            };
        }
        if (isset($d['line-height'])) {
            $lh = $d['line-height'];
            if ($lh === 'normal') {
                $cs['lh'] = null;
            } elseif (is_numeric($lh)) {
                $cs['lh'] = ['auto', (int) round((float) $lh * 240)];
            } elseif (str_ends_with($lh, '%')) {
                $cs['lh'] = ['auto', (int) round((float) $lh * 2.4)];
            } else {
                $tw = $this->len($lh, $cs['sz']);
                $cs['lh'] = $tw !== null && $tw > 0 ? ['atLeast', (int) round($tw)] : null;
            }
        }
        if (isset($d['white-space'])) {
            $cs['ws'] = $d['white-space'];
        }
        if (isset($d['letter-spacing'])) {
            $cs['ls'] = $d['letter-spacing'] === 'normal' ? 0 : (int) round($this->len($d['letter-spacing'], $cs['sz']) ?? 0);
        }
        if (isset($d['text-indent'])) {
            $cs['indent'] = (int) round($this->len($d['text-indent'], $cs['sz'], $this->width) ?? 0);
        }
        if (isset($d['list-style-type'])) {
            $cs['lst'] = $d['list-style-type'];
        }
        if ($tag === 'a' && $el->hasAttribute('href')) {
            $cs['link'] = trim($el->getAttribute('href'));
        }
        return $cs;
    }

    /** Attributi HTML di presentazione (align, bgcolor, width, <font>…). */
    private function presentational(DOMElement $el, string $tag): array
    {
        $d = [];
        $a = static fn (string $n) => $el->hasAttribute($n) ? trim($el->getAttribute($n)) : null;
        if (($v = $a('align')) !== null && !in_array($tag, ['table', 'img'], true)) {
            $d['text-align'] = strtolower($v);
        }
        if (($v = $a('valign')) !== null) {
            $d['vertical-align'] = strtolower($v);
        }
        if (($v = $a('bgcolor')) !== null) {
            $d['background-color'] = strtolower($v);
        }
        foreach (['width', 'height'] as $n) {
            if (($v = $a($n)) !== null && in_array($tag, ['table', 'td', 'th', 'img', 'tr'], true)) {
                $d[$n] = is_numeric($v) ? $v . 'px' : strtolower($v);
            }
        }
        if ($tag === 'font') {
            if (($v = $a('color')) !== null) {
                $d['color'] = strtolower($v);
            }
            if (($v = $a('face')) !== null) {
                $d['font-family'] = $v;
            }
            if (($v = $a('size')) !== null && preg_match('/^([+-]?)(\d)$/', $v, $m)) {
                $n = $m[1] === '' ? (int) $m[2] : 3 + (int) ($m[1] . $m[2]);
                $d['font-size'] = [1 => '7.5pt', 2 => '10pt', 3 => '12pt', 4 => '13.5pt', 5 => '18pt', 6 => '24pt', 7 => '36pt'][max(1, min(7, $n))];
            }
        }
        if (in_array($tag, ['td', 'th'], true)) {
            // cellpadding della tabella: vale per tutte le sue celle
            $t = $el->parentNode;
            while ($t instanceof DOMElement && strtolower($t->nodeName) !== 'table') {
                $t = $t->parentNode;
            }
            if ($t instanceof DOMElement && is_numeric($t->getAttribute('cellpadding'))) {
                foreach (['top', 'right', 'bottom', 'left'] as $s) {
                    $d["padding-$s"] = $t->getAttribute('cellpadding') . 'px';
                }
            }
        }
        return $d;
    }

    private function isBlock(DOMElement $el, array $d): bool
    {
        $disp = $d['display'] ?? null;
        if ($disp !== null) {
            return in_array($disp, ['block', 'list-item', 'table', 'flex', 'grid'], true);
        }
        return in_array(strtolower($el->nodeName), self::BLOCK, true);
    }

    private function fontSize(string $v, float $parent): float
    {
        $kw = ['xx-small' => 7.0, 'x-small' => 7.5, 'small' => 10.0, 'medium' => 12.0, 'large' => 13.5, 'x-large' => 18.0, 'xx-large' => 24.0];
        if (isset($kw[$v])) {
            return $kw[$v];
        }
        if ($v === 'smaller') {
            return max(1.0, $parent / 1.2);
        }
        if ($v === 'larger') {
            return $parent * 1.2;
        }
        if (preg_match('/^([\d.]+)(px|pt|em|rem|%)?$/', $v, $m)) {
            $n = (float) $m[1];
            $r = match ($m[2] ?? '') {
                'pt' => $n, 'em' => $n * $parent, 'rem' => $n * $this->rootPt, '%' => $n * $parent / 100, default => $n * 0.75,
            };
            return max(1.0, min(400.0, $r));
        }
        return $parent;
    }

    private static function fontFamily(string $v): ?string
    {
        foreach (explode(',', $v) as $f) {
            $f = trim($f, " \t\"'");
            $l = strtolower($f);
            if ($f === '' || $l === 'inherit' || $l === 'initial') {
                continue;
            }
            return match ($l) {
                'serif' => 'Times New Roman',
                'sans-serif', 'system-ui', '-apple-system', 'helvetica', 'helvetica neue' => 'Arial',
                'monospace' => 'Courier New',
                default => $f,
            };
        }
        return null;
    }

    /** Lunghezza CSS in twip (1/20 di punto). */
    private function len(?string $v, float $fontPt, float $pctBase = 0.0): ?float
    {
        if ($v === null) {
            return null;
        }
        $v = strtolower(trim($v));
        if ($v === '0') {
            return 0.0;
        }
        if (!preg_match('/^(-?[\d.]+)(px|pt|em|rem|%|cm|mm|in|pc|ex)?$/', $v, $m)) {
            return null;
        }
        $n = (float) $m[1];
        return match ($m[2] ?? '') {
            'pt' => $n * 20, 'em' => $n * $fontPt * 20, 'rem' => $n * $this->rootPt * 20, '%' => $n * $pctBase / 100,
            'cm' => $n * 567, 'mm' => $n * 56.7, 'in' => $n * 1440, 'pc' => $n * 240, 'ex' => $n * $fontPt * 10,
            default => $n * 15,
        };
    }

    public static function color(string $v): ?string
    {
        $v = strtolower(trim($v));
        if (isset(self::COLORS[$v])) {
            return self::COLORS[$v];
        }
        if (preg_match('/^#([0-9a-f]{3})$/', $v, $m)) {
            return strtoupper($m[1][0] . $m[1][0] . $m[1][1] . $m[1][1] . $m[1][2] . $m[1][2]);
        }
        if (preg_match('/^#([0-9a-f]{6})([0-9a-f]{2})?$/', $v, $m)) {
            return (($m[2] ?? '') === '00') ? null : strtoupper($m[1]);
        }
        if (preg_match('/^rgba?\((\d+)%?,(\d+)%?,(\d+)%?(?:,([\d.]+))?\)$/', $v, $m)) {
            if (isset($m[4]) && $m[4] !== '' && (float) $m[4] === 0.0) {
                return null;
            }
            return sprintf('%02X%02X%02X', min(255, (int) $m[1]), min(255, (int) $m[2]), min(255, (int) $m[3]));
        }
        return null;
    }

    /** Bordi dei quattro lati: [lato => [spessore in ottavi di punto, stile Word, colore]] */
    private function borders(array $cs): array
    {
        $d = $cs['d'];
        $out = [];
        foreach (['top', 'right', 'bottom', 'left'] as $s) {
            $st = $d["border-$s-style"] ?? 'none';
            if ($st === 'none' || $st === 'hidden') {
                continue;
            }
            $w = $d["border-$s-width"] ?? 'medium';
            $tw = match ($w) {
                'thin' => 15.0, 'medium' => 45.0, 'thick' => 75.0, default => $this->len($w, $cs['sz']) ?? 15.0,
            };
            if ($tw <= 0) {
                continue;
            }
            $c = $d["border-$s-color"] ?? 'currentcolor';
            $col = $c === 'currentcolor' ? ($cs['color'] ?? '000000') : (self::color($c) ?? '000000');
            $out[$s] = [
                max(2, min(96, (int) round($tw / 20 * 8))),
                match ($st) { 'dotted' => 'dotted', 'dashed' => 'dashed', 'double' => 'double', default => 'single' },
                $col,
            ];
        }
        return $out;
    }

    // ------------------------------------------------------------------
    // Struttura a blocchi
    // ------------------------------------------------------------------

    private function newPara(array $cs, array $bc): array
    {
        return ['k' => 'p', 'cs' => $cs, 'bc' => $bc, 'runs' => [], 'before' => 0, 'after' => 0,
            'break' => false, 'marker' => null, 'br' => false];
    }

    /**
     * Contenuto di un elemento a blocco: il testo in linea diventa paragrafi,
     * i figli a blocco producono i propri elementi.
     */
    private function blockContent(DOMElement $el, array $cs, array $bc): array
    {
        $items = [];
        $p = $this->newPara($cs, $bc);
        foreach (iterator_to_array($el->childNodes) as $c) {
            $this->inline($c, $cs, $bc, $p, $items);
        }
        $this->flush($p, $items);
        return $items;
    }

    private function inline(DOMNode $n, array $pcs, array $bc, array &$p, array &$items): void
    {
        if ($n instanceof DOMText) {
            $this->text($n->data, $pcs, $p);
            return;
        }
        if (!$n instanceof DOMElement) {
            return;
        }
        $tag = strtolower($n->nodeName);
        if (in_array($tag, self::SKIP, true)) {
            return;
        }
        $cs = $this->compute($n, $pcs);
        if (($cs['d']['display'] ?? '') === 'none' || ($cs['d']['visibility'] ?? '') === 'hidden') {
            return;
        }
        if ($this->isBlock($n, $cs['d'])) {
            $this->flush($p, $items);
            foreach ($this->block($n, $cs, $bc) as $it) {
                $items[] = $it;
            }
            $p = $this->newPara($p['cs'], $bc);
            return;
        }
        switch ($tag) {
            case 'br':
                $p['runs'][] = ['x' => '<w:br/>', 'cs' => $cs];
                $p['br'] = true;
                return;
            case 'img':
                $run = $this->image($n, $cs, $bc);
                if ($run !== null) {
                    $p['runs'][] = ['x' => $run, 'cs' => $cs];
                }
                return;
        }
        foreach (iterator_to_array($n->childNodes) as $c) {
            $this->inline($c, $cs, $bc, $p, $items);
        }
    }

    private function text(string $s, array $cs, array &$p): void
    {
        $ws = $cs['ws'];
        if ($ws === 'pre' || $ws === 'pre-wrap' || $ws === 'pre-line') {
            $s = str_replace(["\r\n", "\r"], "\n", $s);
            if ($ws === 'pre-line') {
                $s = preg_replace('/[ \t]+/', ' ', $s) ?? $s;
            }
            $lines = explode("\n", $s);
            foreach ($lines as $i => $line) {
                if ($i > 0) {
                    $p['runs'][] = ['x' => '<w:br/>', 'cs' => $cs];
                }
                foreach (explode("\t", $line) as $j => $part) {
                    if ($j > 0) {
                        $p['runs'][] = ['x' => '<w:tab/>', 'cs' => $cs];
                    }
                    if ($part !== '') {
                        $p['runs'][] = ['t' => $this->transform($part, $cs), 'cs' => $cs];
                    }
                }
            }
            return;
        }
        $s = preg_replace('/[ \t\n\r\f]+/', ' ', $s) ?? $s;
        if ($s === '') {
            return;
        }
        // Spazio iniziale: si toglie a inizio paragrafo o dopo un altro spazio
        if ($s[0] === ' ' && $this->endsWithSpace($p)) {
            $s = substr($s, 1);
        }
        if ($s !== '') {
            $p['runs'][] = ['t' => $this->transform($s, $cs), 'cs' => $cs];
        }
    }

    private function endsWithSpace(array $p): bool
    {
        for ($i = count($p['runs']) - 1; $i >= 0; $i--) {
            $r = $p['runs'][$i];
            if (isset($r['t'])) {
                return str_ends_with($r['t'], ' ');
            }
            return $r['x'] === '<w:br/>' || $r['x'] === '<w:tab/>';
        }
        return true;
    }

    private function transform(string $s, array $cs): string
    {
        return match ($cs['tt']) {
            'uppercase' => mb_strtoupper($s, 'UTF-8'),
            'lowercase' => mb_strtolower($s, 'UTF-8'),
            'capitalize' => mb_convert_case($s, MB_CASE_TITLE, 'UTF-8'),
            default => $s,
        };
    }

    /** Chiude il paragrafo corrente aggiungendolo agli elementi, se ha contenuto. */
    private function flush(array &$p, array &$items): void
    {
        $runs = $p['runs'];
        // Spazi finali
        for ($i = count($runs) - 1; $i >= 0; $i--) {
            if (isset($runs[$i]['t'])) {
                $runs[$i]['t'] = rtrim($runs[$i]['t'], ' ');
                if ($runs[$i]['t'] === '') {
                    array_splice($runs, $i, 1);
                    continue;
                }
            }
            break;
        }
        // Un <br> finale non crea una riga in più (come nel browser)
        if ($runs && ($runs[count($runs) - 1]['x'] ?? '') === '<w:br/>') {
            array_pop($runs);
        }
        $content = false;
        foreach ($runs as $r) {
            if (isset($r['t']) ? $r['t'] !== '' : true) {
                $content = true;
                break;
            }
        }
        if ($content || $p['br']) {
            $p['runs'] = $runs;
            $this->takePending($p);
            $items[] = $p;
        }
        $p = $this->newPara($p['cs'], $p['bc']);
    }

    /** Applica segno di elenco e interruzione di pagina in sospeso. */
    private function takePending(array &$it): void
    {
        if ($this->pendingBreak) {
            $it['break'] = true;
            $this->pendingBreak = false;
        }
        if ($it['k'] === 'p' && $this->marker !== null) {
            $it['marker'] = $this->marker;
            $this->marker = null;
        }
    }

    private function block(DOMElement $el, array $cs, array $bc): array
    {
        $tag = strtolower($el->nodeName);
        $d = $cs['d'];
        if (($d['page-break-before'] ?? $d['break-before'] ?? '') === 'always' || ($d['break-before'] ?? '') === 'page') {
            $this->pendingBreak = true;
        }

        if ($tag === 'table' || ($d['display'] ?? '') === 'table') {
            $items = [$this->table($el, $cs, $bc)];
            $this->takePending($items[0]);
        } elseif ($tag === 'hr') {
            $p = $this->newPara($cs, $bc);
            $p['hr'] = true;
            $p['br'] = true;
            $items = [$p];
            $this->takePending($items[0]);
        } else {
            $ml = $this->len($d['margin-left'] ?? null, $cs['sz'], $bc['width']) ?? 0;
            $mr = $this->len($d['margin-right'] ?? null, $cs['sz'], $bc['width']) ?? 0;
            $pl = $this->len($d['padding-left'] ?? null, $cs['sz'], $bc['width']) ?? 0;
            $pr = $this->len($d['padding-right'] ?? null, $cs['sz'], $bc['width']) ?? 0;
            $nbc = $bc;
            $nbc['indL'] = (int) round($bc['indL'] + $ml + $pl);
            $nbc['indR'] = (int) round($bc['indR'] + $mr + $pr);
            $nbc['width'] = (int) max(720, $bc['width'] - $ml - $pl - $mr - $pr);
            $bg = isset($d['background-color']) ? self::color($d['background-color']) : null;
            if ($bg !== null) {
                $nbc['shd'] = $bg;
            }
            $bdr = $this->borders($cs);
            if ($bdr) {
                $nbc['bdr'] = $bdr;
            }

            if ($tag === 'ul' || $tag === 'ol') {
                $nbc['list'][] = ['tag' => $tag, 'n' => is_numeric($el->getAttribute('start')) ? (int) $el->getAttribute('start') : 1,
                    'type' => $el->getAttribute('type')];
            }
            if ($tag === 'li' || ($d['display'] ?? '') === 'list-item') {
                $this->marker = $this->listMarker($el, $cs, $bc);
            }
            $items = $this->blockContent($el, $cs, $nbc);
            if ($tag === 'li') {
                $this->marker = null;
            }
            if ($items) {
                $mt = (int) round($this->len($d['margin-top'] ?? null, $cs['sz'], $bc['width']) ?? 0)
                    + (int) round($this->len($d['padding-top'] ?? null, $cs['sz'], $bc['width']) ?? 0);
                $mb = (int) round($this->len($d['margin-bottom'] ?? null, $cs['sz'], $bc['width']) ?? 0)
                    + (int) round($this->len($d['padding-bottom'] ?? null, $cs['sz'], $bc['width']) ?? 0);
                $first = 0;
                $last = count($items) - 1;
                if ($items[$first]['k'] === 'p') {
                    $items[$first]['before'] = max($items[$first]['before'], $mt);
                }
                if ($items[$last]['k'] === 'p') {
                    $items[$last]['after'] = max($items[$last]['after'], $mb);
                }
            }
        }
        if (($d['page-break-after'] ?? '') === 'always' || ($d['break-after'] ?? '') === 'page') {
            $this->pendingBreak = true;
        }
        return $items;
    }

    /** Testo del segno per un <li>: punto, cerchio, numero, lettera… */
    private function listMarker(DOMElement $li, array $cs, array $bc): ?array
    {
        $parent = $li->parentNode;
        $type = $cs['lst'];
        if ($parent instanceof DOMElement && $parent->hasAttribute('type')) {
            $type = match ($parent->getAttribute('type')) {
                'a' => 'lower-alpha', 'A' => 'upper-alpha', 'i' => 'lower-roman', 'I' => 'upper-roman',
                '1' => 'decimal', 'circle' => 'circle', 'square' => 'square', default => $type,
            };
        }
        if ($type === 'none') {
            return null;
        }
        $n = 1;
        if ($parent instanceof DOMElement && strtolower($parent->nodeName) === 'ol') {
            $n = is_numeric($parent->getAttribute('start')) ? (int) $parent->getAttribute('start') : 1;
            for ($s = $parent->firstChild; $s !== null && $s !== $li; $s = $s->nextSibling) {
                if ($s instanceof DOMElement && strtolower($s->nodeName) === 'li') {
                    $n = is_numeric($s->getAttribute('value')) ? (int) $s->getAttribute('value') + 1 : $n + 1;
                }
            }
            if (is_numeric($li->getAttribute('value'))) {
                $n = (int) $li->getAttribute('value');
            }
        }
        $txt = match ($type) {
            'disc' => '•', 'circle' => '◦', 'square' => '▪',
            'decimal' => $n . '.', 'decimal-leading-zero' => sprintf('%02d.', $n),
            'lower-alpha', 'lower-latin' => self::alpha($n) . '.', 'upper-alpha', 'upper-latin' => strtoupper(self::alpha($n)) . '.',
            'lower-roman' => strtolower(self::roman($n)) . '.', 'upper-roman' => self::roman($n) . '.',
            default => '•',
        };
        return ['t' => $txt, 'cs' => $cs];
    }

    private static function alpha(int $n): string
    {
        $s = '';
        for ($n = max(1, $n); $n > 0; $n = intdiv($n - 1, 26)) {
            $s = chr(97 + ($n - 1) % 26) . $s;
        }
        return $s;
    }

    private static function roman(int $n): string
    {
        $map = ['M' => 1000, 'CM' => 900, 'D' => 500, 'CD' => 400, 'C' => 100, 'XC' => 90, 'L' => 50, 'XL' => 40, 'X' => 10, 'IX' => 9, 'V' => 5, 'IV' => 4, 'I' => 1];
        $s = '';
        $n = max(1, min(3999, $n));
        foreach ($map as $r => $v) {
            while ($n >= $v) {
                $s .= $r;
                $n -= $v;
            }
        }
        return $s;
    }

    // ------------------------------------------------------------------
    // Tabelle
    // ------------------------------------------------------------------

    private function table(DOMElement $t, array $cs, array $bc): array
    {
        $d = $cs['d'];
        $avail = max(720, $bc['width']);
        $tw = $this->len($d['width'] ?? null, $cs['sz'], $avail);
        $tw = (int) round($tw === null || $tw <= 0 ? $avail : min($tw, $avail));

        // Righe (dirette o dentro thead/tbody/tfoot), con il loro stile
        $rows = [];
        $collect = function (DOMElement $parent, array $pcs, bool $head) use (&$rows, &$collect): void {
            foreach ($parent->childNodes as $c) {
                if (!$c instanceof DOMElement) {
                    continue;
                }
                $tag = strtolower($c->nodeName);
                if (in_array($tag, ['thead', 'tbody', 'tfoot'], true)) {
                    $collect($c, $this->compute($c, $pcs), $head || $tag === 'thead');
                } elseif ($tag === 'tr') {
                    $rcs = $this->compute($c, $pcs);
                    if (($rcs['d']['display'] ?? '') === 'none') {
                        continue;
                    }
                    $cells = [];
                    foreach ($c->childNodes as $td) {
                        if ($td instanceof DOMElement && in_array(strtolower($td->nodeName), ['td', 'th'], true)) {
                            $cells[] = [$td, $this->compute($td, $rcs)];
                        }
                    }
                    $rows[] = ['el' => $c, 'cs' => $rcs, 'head' => $head, 'cells' => $cells];
                }
            }
        };
        $collect($t, $cs, false);
        if (!$rows) {
            return ['k' => 'raw', 'xml' => '', 'break' => false];
        }

        // Griglia: posizione di ogni cella tenendo conto di colspan e rowspan
        $occupied = [];
        $placed = [];
        $ncols = 0;
        foreach ($rows as $ri => $row) {
            $col = 0;
            foreach ($row['cells'] as $ci => [$td]) {
                while (isset($occupied[$ri][$col])) {
                    $col++;
                }
                $span = max(1, min(50, (int) ($td->getAttribute('colspan') ?: 1)));
                $rspan = max(1, min(count($rows) - $ri, (int) ($td->getAttribute('rowspan') ?: 1)));
                for ($r = 0; $r < $rspan; $r++) {
                    for ($k = 0; $k < $span; $k++) {
                        $occupied[$ri + $r][$col + $k] = [$ri, $ci];
                    }
                }
                $placed[$ri][$col] = ['ci' => $ci, 'span' => $span, 'rspan' => $rspan];
                $col += $span;
            }
            $ncols = max($ncols, $col, isset($occupied[$ri]) ? max(array_keys($occupied[$ri])) + 1 : 0);
        }
        $ncols = max(1, $ncols);

        // Larghezze delle colonne
        $widths = array_fill(0, $ncols, null);
        foreach ($placed as $ri => $cols) {
            foreach ($cols as $col => $info) {
                if ($info['span'] !== 1) {
                    continue;
                }
                $ccs = $rows[$ri]['cells'][$info['ci']][1];
                $w = $this->len($ccs['d']['width'] ?? null, $ccs['sz'], $tw);
                if ($w !== null && $w > 0 && $widths[$col] === null) {
                    $widths[$col] = $w;
                }
            }
        }
        $known = array_sum(array_filter($widths, static fn ($w) => $w !== null));
        $unknown = count(array_filter($widths, static fn ($w) => $w === null));
        if ($unknown > 0) {
            $rest = max($tw - $known, 0);
            $each = $rest > $unknown * 300 ? $rest / $unknown : max(300, $tw / $ncols);
            foreach ($widths as $i => $w) {
                if ($w === null) {
                    $widths[$i] = $each;
                }
            }
        }
        $sum = array_sum($widths);
        $grid = [];
        foreach ($widths as $w) {
            $grid[] = max(60, (int) round($w * $tw / $sum));
        }
        $tw = array_sum($grid);

        // Bordi: attributo border della tabella e CSS
        $attrBorder = is_numeric($t->getAttribute('border')) && (int) $t->getAttribute('border') > 0;
        $tblBdr = $this->borders($cs);
        if ($attrBorder && !$tblBdr) {
            $b = [max(2, min(96, (int) $t->getAttribute('border') * 6)), 'single', '000000'];
            $tblBdr = ['top' => $b, 'right' => $b, 'bottom' => $b, 'left' => $b];
        }
        $tblBg = isset($d['background-color']) ? self::color($d['background-color']) : null;

        $x = '<w:tbl><w:tblPr><w:tblW w:w="' . $tw . '" w:type="dxa"/>';
        $align = strtolower($t->getAttribute('align'));
        if ($align === 'center' || (($d['margin-left'] ?? '') === 'auto' && ($d['margin-right'] ?? '') === 'auto')) {
            $x .= '<w:jc w:val="center"/>';
        } elseif ($align === 'right') {
            $x .= '<w:jc w:val="right"/>';
        }
        if ($bc['indL'] > 0) {
            $x .= '<w:tblInd w:w="' . $bc['indL'] . '" w:type="dxa"/>';
        }
        if ($tblBdr) {
            $x .= '<w:tblBorders>' . self::bdrXml($tblBdr, ['top', 'left', 'bottom', 'right']) . '</w:tblBorders>';
        }
        $x .= '<w:tblLayout w:type="fixed"/>'
            . '<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="0" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar>'
            . '<w:tblLook w:val="0000" w:firstRow="0" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" w:noHBand="1" w:noVBand="1"/>'
            . '</w:tblPr><w:tblGrid>';
        foreach ($grid as $g) {
            $x .= '<w:gridCol w:w="' . $g . '"/>';
        }
        $x .= '</w:tblGrid>';

        $cellBdr = $attrBorder ? ['top' => [4, 'single', '000000'], 'right' => [4, 'single', '000000'], 'bottom' => [4, 'single', '000000'], 'left' => [4, 'single', '000000']] : [];
        $cellXml = [];
        foreach ($rows as $ri => $row) {
            $x .= '<w:tr><w:trPr><w:cantSplit/>';
            $h = $this->len($row['cs']['d']['height'] ?? null, $row['cs']['sz']);
            if ($h !== null && $h > 0) {
                $x .= '<w:trHeight w:val="' . (int) round($h) . '"/>';
            }
            if ($row['head']) {
                $x .= '<w:tblHeader/>';
            }
            $x .= '</w:trPr>';
            $rowBg = isset($row['cs']['d']['background-color']) ? self::color($row['cs']['d']['background-color']) : null;
            for ($col = 0; $col < $ncols;) {
                if (isset($placed[$ri][$col])) {
                    $info = $placed[$ri][$col];
                    [$td, $ccs] = $row['cells'][$info['ci']];
                    $w = array_sum(array_slice($grid, $col, $info['span']));
                    $bg = isset($ccs['d']['background-color']) ? self::color($ccs['d']['background-color']) : null;
                    $bg ??= $rowBg ?? $tblBg;
                    $bdr = $this->borders($ccs) + $cellBdr;
                    $mar = $this->cellMargins($ccs, $w);
                    $va = match ($ccs['d']['vertical-align'] ?? $row['cs']['d']['vertical-align'] ?? 'middle') {
                        'top', 'baseline', 'text-top' => 'top', 'bottom', 'text-bottom' => 'bottom', default => 'center',
                    };
                    $pr = '<w:tcW w:w="' . $w . '" w:type="dxa"/>'
                        . ($info['span'] > 1 ? '<w:gridSpan w:val="' . $info['span'] . '"/>' : '')
                        . ($info['rspan'] > 1 ? '<w:vMerge w:val="restart"/>' : '');
                    $tail = ($bdr ? '<w:tcBorders>' . self::bdrXml($bdr, ['top', 'left', 'bottom', 'right']) . '</w:tcBorders>' : '')
                        . ($bg !== null ? '<w:shd w:val="clear" w:color="auto" w:fill="' . $bg . '"/>' : '')
                        . '<w:tcMar><w:top w:w="' . $mar[0] . '" w:type="dxa"/><w:left w:w="' . $mar[3] . '" w:type="dxa"/>'
                        . '<w:bottom w:w="' . $mar[2] . '" w:type="dxa"/><w:right w:w="' . $mar[1] . '" w:type="dxa"/></w:tcMar>'
                        . '<w:vAlign w:val="' . $va . '"/>';
                    $inner = max(300, $w - $mar[1] - $mar[3]);
                    $saveMarker = $this->marker;
                    $saveBreak = $this->pendingBreak;
                    $this->marker = null;
                    $this->pendingBreak = false;
                    $content = $this->serialize($this->blockContent($td, $ccs,
                        ['indL' => 0, 'indR' => 0, 'width' => $inner, 'shd' => null, 'bdr' => null, 'list' => []]), true, $ccs);
                    $this->marker = $saveMarker;
                    $this->pendingBreak = $saveBreak;
                    $x .= '<w:tc><w:tcPr>' . $pr . $tail . '</w:tcPr>' . $content . '</w:tc>';
                    if ($info['rspan'] > 1) {
                        // Le righe sotto ricevono una cella di continuazione con le stesse proprietà
                        for ($r = 1; $r < $info['rspan']; $r++) {
                            $cellXml[$ri + $r][$col] = ['span' => $info['span'], 'xml' => '<w:tc><w:tcPr><w:tcW w:w="' . $w . '" w:type="dxa"/>'
                                . ($info['span'] > 1 ? '<w:gridSpan w:val="' . $info['span'] . '"/>' : '')
                                . '<w:vMerge/>' . $tail . '</w:tcPr><w:p/></w:tc>'];
                        }
                    }
                    $col += $info['span'];
                } elseif (isset($cellXml[$ri][$col])) {
                    $x .= $cellXml[$ri][$col]['xml'];
                    $col += $cellXml[$ri][$col]['span'];
                } else {
                    // Riga più corta delle altre: cella vuota
                    $x .= '<w:tc><w:tcPr><w:tcW w:w="' . $grid[$col] . '" w:type="dxa"/></w:tcPr><w:p/></w:tc>';
                    $col++;
                }
            }
            $x .= '</w:tr>';
        }
        $x .= '</w:tbl>';
        return ['k' => 'raw', 'xml' => $x, 'break' => false];
    }

    /** Margini interni della cella in twip: alto, destra, basso, sinistra. */
    private function cellMargins(array $cs, int $w): array
    {
        $d = $cs['d'];
        $author = $cs['author'];
        $prop = 'padding';
        // Diversi modelli usano "margin" sulle celle come se fosse padding: se il
        // padding non è stato indicato dall'autore, lo si prende da lì
        $hasPad = isset($author['padding-top']) || isset($author['padding-left']) || isset($author['padding-right']) || isset($author['padding-bottom']);
        $hasMar = isset($author['margin-top']) || isset($author['margin-left']) || isset($author['margin-right']) || isset($author['margin-bottom']);
        if (!$hasPad && $hasMar) {
            $prop = 'margin';
        }
        $out = [];
        foreach (['top', 'right', 'bottom', 'left'] as $s) {
            $v = $this->len($d["$prop-$s"] ?? null, $cs['sz'], $w);
            $out[] = (int) max(0, min(2000, round($v ?? 15)));
        }
        return $out;
    }

    private static function bdrXml(array $b, array $order): string
    {
        $x = '';
        foreach ($order as $s) {
            if (isset($b[$s])) {
                [$sz, $st, $col] = $b[$s];
                $x .= '<w:' . $s . ' w:val="' . $st . '" w:sz="' . $sz . '" w:space="0" w:color="' . $col . '"/>';
            }
        }
        return $x;
    }

    // ------------------------------------------------------------------
    // Immagini
    // ------------------------------------------------------------------

    private function image(DOMElement $img, array $cs, array $bc): ?string
    {
        $src = trim($img->getAttribute('src'));
        $bytes = null;
        if (preg_match('#^data:image/[a-z+.-]+;base64,(.*)$#is', $src, $m)) {
            $bytes = base64_decode(preg_replace('/\s+/', '', $m[1]) ?? '', true) ?: null;
        } elseif (preg_match('#^https?://#i', $src)) {
            $bytes = self::download($src);
        }
        if ($bytes === null) {
            GDocx::warn('Immagine non caricata: ' . (strlen($src) > 80 ? substr($src, 0, 80) . '…' : $src));
            return null;
        }
        $info = @getimagesizefromstring($bytes);
        if ($info === false) {
            GDocx::warn('Formato immagine non riconosciuto');
            return null;
        }
        [$pw, $ph, $type] = $info;
        if ($type === IMAGETYPE_JPEG) {
            $ext = 'jpeg';
        } elseif ($type === IMAGETYPE_PNG) {
            $ext = 'png';
        } elseif (function_exists('imagecreatefromstring') && ($gd = @imagecreatefromstring($bytes)) !== false) {
            // GIF, WebP, BMP…: convertite in PNG
            imagesavealpha($gd, true);
            ob_start();
            imagepng($gd);
            $bytes = (string) ob_get_clean();
            $ext = 'png';
        } else {
            GDocx::warn('Formato immagine non supportato (servono PNG o JPEG, oppure l\'estensione GD)');
            return null;
        }
        if ($pw <= 0 || $ph <= 0) {
            return null;
        }
        // Dimensioni in twip: attributi/CSS, altrimenti quelle naturali a 96 dpi
        $w = $this->len($cs['d']['width'] ?? null, $cs['sz'], $bc['width']);
        $h = $this->len($cs['d']['height'] ?? null, $cs['sz']);
        if ($w !== null && $w > 0 && ($h === null || $h <= 0)) {
            $h = $w * $ph / $pw;
        } elseif ($h !== null && $h > 0 && ($w === null || $w <= 0)) {
            $w = $h * $pw / $ph;
        } elseif ($w === null || $w <= 0) {
            $w = $pw * 15;
            $h = $ph * 15;
        }
        $max = max(300, $bc['width']);
        if ($w > $max) {
            $h = $h * $max / $w;
            $w = $max;
        }
        $cx = (int) round($w * 635);
        $cy = (int) round($h * 635);
        $rid = $this->pkg->addImage($bytes, $ext);
        $id = $this->pkg->nextDocPrId();
        $alt = htmlspecialchars($img->getAttribute('alt'), ENT_XML1 | ENT_QUOTES, 'UTF-8');
        return '<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="' . $cx . '" cy="' . $cy . '"/>'
            . '<wp:docPr id="' . $id . '" name="Immagine ' . $id . '" descr="' . $alt . '"/>'
            . '<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>'
            . '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic>'
            . '<pic:nvPicPr><pic:cNvPr id="' . $id . '" name="img' . $id . '.' . $ext . '"/><pic:cNvPicPr/></pic:nvPicPr>'
            . '<pic:blipFill><a:blip r:embed="' . $rid . '"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
            . '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="' . $cx . '" cy="' . $cy . '"/></a:xfrm>'
            . '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic>'
            . '</wp:inline></w:drawing>';
    }

    private static function download(string $url): ?string
    {
        if (!function_exists('curl_init')) {
            return null;
        }
        $ch = curl_init($url);
        $data = '';
        curl_setopt_array($ch, [
            CURLOPT_FOLLOWLOCATION => true,
            CURLOPT_MAXREDIRS => 3,
            CURLOPT_PROTOCOLS => CURLPROTO_HTTP | CURLPROTO_HTTPS,
            CURLOPT_REDIR_PROTOCOLS => CURLPROTO_HTTP | CURLPROTO_HTTPS,
            CURLOPT_CONNECTTIMEOUT => 5,
            CURLOPT_TIMEOUT => 15,
            CURLOPT_USERAGENT => 'FZPureDOCX',
            // Interrompe i download oltre 10 MB
            CURLOPT_WRITEFUNCTION => static function ($ch, string $chunk) use (&$data): int {
                $data .= $chunk;
                return strlen($data) > 10 * 1024 * 1024 ? 0 : strlen($chunk);
            },
        ]);
        $ok = curl_exec($ch);
        $code = (int) curl_getinfo($ch, CURLINFO_RESPONSE_CODE);
        return $ok !== false && $code >= 200 && $code < 300 && $data !== '' ? $data : null;
    }

    // ------------------------------------------------------------------
    // Scrittura XML
    // ------------------------------------------------------------------

    /**
     * @param bool $cell contenuto di una cella: serve almeno un paragrafo, e uno finale dopo una tabella
     */
    private function serialize(array $items, bool $cell = false, ?array $cellCs = null): string
    {
        $x = '';
        $prevAfter = 0;
        $lastKind = null;
        foreach ($items as $it) {
            if ($it['k'] === 'raw') {
                if ($it['xml'] === '') {
                    continue;
                }
                if ($it['break']) {
                    $x .= '<w:p><w:pPr><w:pageBreakBefore/><w:spacing w:before="0" w:after="0" w:line="20" w:lineRule="exact"/>'
                        . '<w:rPr><w:sz w:val="2"/><w:szCs w:val="2"/></w:rPr></w:pPr></w:p>';
                } elseif ($lastKind === 't') {
                    // Due tabelle consecutive verrebbero fuse in una sola
                    $x .= self::SEPARATOR;
                }
                $x .= $it['xml'];
                $prevAfter = 0;
                $lastKind = 't';
                continue;
            }
            // Margini adiacenti: in CSS si fondono, in Word si sommano
            $before = max(0, $it['before'] - $prevAfter);
            $x .= $this->paraXml($it, $before);
            $prevAfter = $it['after'];
            $lastKind = 'p';
        }
        if ($cell && $lastKind !== 'p') {
            $sz = (int) round(($cellCs['sz'] ?? 10) * 2);
            $x .= '<w:p><w:pPr><w:spacing w:before="0" w:after="0"/><w:rPr><w:sz w:val="' . $sz . '"/><w:szCs w:val="' . $sz . '"/></w:rPr></w:pPr></w:p>';
        }
        return $x;
    }

    private function paraXml(array $p, int $before): string
    {
        $cs = $p['cs'];
        $bc = $p['bc'];
        $ppr = '';
        if ($p['break']) {
            $ppr .= '<w:pageBreakBefore/>';
        }
        if (!empty($p['hr'])) {
            $c = isset($cs['d']['border-top-color']) ? (self::color($cs['d']['border-top-color']) ?? '808080') : '808080';
            $ppr .= '<w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="' . $c . '"/></w:pBdr>';
        } elseif ($bc['bdr']) {
            $ppr .= '<w:pBdr>' . self::bdrXml(array_map(static fn ($b) => [$b[0], $b[1], $b[2]], $bc['bdr']), ['top', 'left', 'bottom', 'right']) . '</w:pBdr>';
        }
        if ($bc['shd'] !== null) {
            $ppr .= '<w:shd w:val="clear" w:color="auto" w:fill="' . $bc['shd'] . '"/>';
        }
        $indL = $bc['indL'];
        $marker = $p['marker'];
        if ($marker !== null) {
            $ppr .= '<w:tabs><w:tab w:val="left" w:pos="' . $indL . '"/></w:tabs>';
        }
        $line = $cs['lh'];
        $ppr .= '<w:spacing w:before="' . $before . '" w:after="' . $p['after'] . '"'
            . ($line !== null ? ' w:line="' . $line[1] . '" w:lineRule="' . $line[0] . '"' : ' w:line="240" w:lineRule="auto"') . '/>';
        $ind = '';
        if ($indL > 0) {
            $ind .= ' w:left="' . $indL . '"';
        }
        if ($bc['indR'] > 0) {
            $ind .= ' w:right="' . $bc['indR'] . '"';
        }
        if ($marker !== null) {
            $ind .= ' w:hanging="' . min(360, max(0, $indL)) . '"';
        } elseif ($cs['indent'] > 0) {
            $ind .= ' w:firstLine="' . $cs['indent'] . '"';
        } elseif ($cs['indent'] < 0) {
            $ind .= ' w:hanging="' . (-$cs['indent']) . '"';
        }
        if ($ind !== '') {
            $ppr .= '<w:ind' . $ind . '/>';
        }
        $jc = ['left' => null, 'center' => 'center', 'right' => 'right', 'justify' => 'both'][$cs['align']] ?? null;
        if ($jc !== null) {
            $ppr .= '<w:jc w:val="' . $jc . '"/>';
        }
        $ppr .= '<w:rPr>' . $this->rPr($cs, true) . '</w:rPr>';

        $body = '';
        if ($marker !== null) {
            $mcs = $marker['cs'];
            $mcs['u'] = $mcs['strike'] = false;
            $mcs['bg'] = null;
            $body .= '<w:r><w:rPr>' . $this->rPr($mcs) . '</w:rPr><w:t xml:space="preserve">' . self::esc($marker['t']) . '</w:t><w:tab/></w:r>';
        }
        // Raggruppa le sequenze con lo stesso collegamento
        $link = null;
        $open = false;
        foreach ($p['runs'] as $r) {
            $l = $r['cs']['link'];
            if ($l !== $link) {
                if ($open) {
                    $body .= '</w:hyperlink>';
                    $open = false;
                }
                $link = $l;
                if ($l !== null && $l !== '') {
                    if ($l[0] === '#') {
                        $body .= '<w:hyperlink w:anchor="' . self::esc(substr($l, 1)) . '">';
                        $open = true;
                    } elseif (preg_match('#^(https?|mailto|tel|ftp):#i', $l)) {
                        $body .= '<w:hyperlink r:id="' . $this->pkg->addHyperlink($l) . '">';
                        $open = true;
                    }
                }
            }
            $rpr = '<w:rPr>' . $this->rPr($r['cs']) . '</w:rPr>';
            if (isset($r['t'])) {
                $body .= '<w:r>' . $rpr . '<w:t xml:space="preserve">' . self::esc($r['t']) . '</w:t></w:r>';
            } else {
                $body .= '<w:r>' . $rpr . $r['x'] . '</w:r>';
            }
        }
        if ($open) {
            $body .= '</w:hyperlink>';
        }
        return '<w:p><w:pPr>' . $ppr . '</w:pPr>' . $body . '</w:p>';
    }

    private function rPr(array $cs, bool $mark = false): string
    {
        $x = '';
        if ($cs['ff'] !== null) {
            $f = self::esc($cs['ff']);
            $x .= '<w:rFonts w:ascii="' . $f . '" w:hAnsi="' . $f . '" w:cs="' . $f . '" w:eastAsia="' . $f . '"/>';
        }
        if ($cs['b']) {
            $x .= '<w:b/><w:bCs/>';
        }
        if ($cs['i']) {
            $x .= '<w:i/><w:iCs/>';
        }
        if ($mark) {
            $sz = (int) round($cs['sz'] * 2);
            return $x . '<w:sz w:val="' . $sz . '"/><w:szCs w:val="' . $sz . '"/>';
        }
        if ($cs['caps']) {
            $x .= '<w:smallCaps/>';
        }
        if ($cs['strike']) {
            $x .= '<w:strike/>';
        }
        if ($cs['color'] !== null) {
            $x .= '<w:color w:val="' . $cs['color'] . '"/>';
        }
        if ($cs['ls'] !== 0) {
            $x .= '<w:spacing w:val="' . $cs['ls'] . '"/>';
        }
        $sz = (int) round($cs['sz'] * 2);
        $x .= '<w:sz w:val="' . $sz . '"/><w:szCs w:val="' . $sz . '"/>';
        if ($cs['u']) {
            $x .= '<w:u w:val="single"/>';
        }
        if ($cs['bg'] !== null) {
            $x .= '<w:shd w:val="clear" w:color="auto" w:fill="' . $cs['bg'] . '"/>';
        }
        if ($cs['va'] !== 'baseline') {
            $x .= '<w:vertAlign w:val="' . ($cs['va'] === 'super' ? 'superscript' : 'subscript') . '"/>';
        }
        return $x;
    }

    private static function esc(string $s): string
    {
        $s = preg_replace('/[\x00-\x08\x0B\x0C\x0E-\x1F]/', '', $s) ?? '';
        return htmlspecialchars($s, ENT_XML1 | ENT_QUOTES | ENT_SUBSTITUTE, 'UTF-8');
    }
}
