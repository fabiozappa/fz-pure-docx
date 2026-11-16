<?php

declare(strict_types=1);

/**
 * DocxPdf — conversione DOCX → PDF in puro PHP, senza programmi esterni.
 *
 * Requisiti: PHP 8.1+, estensioni zip, dom e zlib. Nessuna dipendenza esterna.
 *
 * Uso diretto:
 *   require 'DocxPdf.php';
 *   $pdf = DocxPdf::convert('documento.docx');              // contenuto binario del PDF
 *   DocxPdf::convertToFile('documento.docx', 'out.pdf');
 *   print_r(DocxPdf::$warnings);                            // font sostituiti, elementi ignorati
 *
 * Opzioni:
 *   'fonts'   => cartella (o elenco di cartelle) con i file .ttf/.ttc da usare per primi
 *   'decimal' => separatore per le tabulazioni decimali (default: quello del documento, altrimenti ',')
 *   'title'   => titolo scritto nelle proprietà del PDF (default: nome del file)
 *
 * Font: il motore legge i TrueType veri, ne usa le metriche per calcolare gli
 * a capo e li incorpora nel PDF (solo i glifi usati). Cerca nell'ordine: le
 * cartelle dell'opzione 'fonts', la cartella "fonts" accanto a questo file, i
 * font di sistema (macOS, Windows, Linux). Se un font manca usa un sostituto
 * e lo segnala in DocxPdf::$warnings: in quel caso gli a capo possono
 * differire da quelli di Word.
 *
 * Supportato: stili ed ereditarietà, paragrafi (allineamento, rientri,
 * interlinea, spaziature, bordi, sfondi), caratteri (grassetto, corsivo,
 * sottolineato, barrato, colore, evidenziatore, apice/pedice, spaziatura),
 * tabulazioni con riempimento, elenchi puntati e numerati, tabelle (celle
 * unite, bordi, sfondi, righe di intestazione ripetute, tabelle annidate),
 * immagini PNG/JPEG in linea e flottanti, forme semplici e caselle di testo,
 * intestazioni e piè di pagina (prima pagina, pari/dispari), numeri di
 * pagina, sezioni con formati di pagina diversi, interruzioni di pagina.
 *
 * Non supportato (ignorato con un avviso): testo che scorre attorno agli
 * oggetti flottanti, colonne multiple, note a piè di pagina, grafici,
 * SmartArt, forme raggruppate, immagini EMF/WMF, font OpenType CFF (.otf).
 */
final class DocxPdf
{
    /** Avvisi dell'ultima conversione (font sostituiti, elementi non supportati). */
    public static array $warnings = [];

    /** Converte un DOCX e restituisce il contenuto binario del PDF. */
    public static function convert(string $docxPath, array $options = []): string
    {
        foreach (['zip' => 'ZipArchive', 'dom' => 'DOMDocument'] as $ext => $class) {
            if (!class_exists($class)) {
                throw new DocxPdfException("Estensione PHP '$ext' mancante");
            }
        }
        if (!function_exists('gzcompress')) {
            throw new DocxPdfException("Estensione PHP 'zlib' mancante");
        }
        if (!is_file($docxPath) || !is_readable($docxPath)) {
            throw new DocxPdfException("File non trovato o non leggibile: $docxPath");
        }
        $engine = new DocxPdfEngine($docxPath, $options);
        $pdf = $engine->render();
        self::$warnings = array_values(array_unique($engine->warn));
        return $pdf;
    }

    /** Converte un DOCX e scrive il PDF nel percorso indicato. */
    public static function convertToFile(string $docxPath, string $pdfPath, array $options = []): string
    {
        $pdf = self::convert($docxPath, $options);
        if (file_put_contents($pdfPath, $pdf) === false) {
            throw new DocxPdfException("Scrittura fallita: $pdfPath");
        }
        return $pdfPath;
    }
}

final class DocxPdfException extends RuntimeException
{
}

// ======================================================================
// Font TrueType: metriche, mappa caratteri, sottoinsieme da incorporare
// ======================================================================

final class DocxPdfFont
{
    public string $family = '';
    public string $psName = 'Font';
    public bool $bold = false;
    public bool $italic = false;
    public int $upm = 1000;
    /** Ascendente, discendente e interlinea esterna, in frazioni di em */
    public float $asc = 0.9;
    public float $desc = 0.2;
    public float $lead = 0.0;
    public float $ulPos = -0.1;
    public float $ulThick = 0.05;
    public float $italicAngle = 0.0;
    public float $capH = 0.7;
    public array $bbox = [0, -200, 1000, 900];
    /** Nome della risorsa nel PDF (assegnato al primo uso) */
    public string $res = '';
    /** Glifi usati: gid => codepoint */
    public array $used = [0 => 0];

    private string $data;
    private int $base;
    private array $tab = [];
    private array $cmap = [];
    private string $hmtx = '';
    private int $nHM = 0;
    private int $numGlyphs = 0;
    private int $locFmt = 0;
    private array $adv = [];
    private array $wcache = [];
    private static array $files = [];

    public function __construct(string $file, int $offset = 0)
    {
        if (!isset(self::$files[$file])) {
            $d = @file_get_contents($file);
            if ($d === false) {
                throw new DocxPdfException("Font non leggibile: $file");
            }
            self::$files[$file] = $d;
        }
        $this->data = self::$files[$file];
        $this->base = $offset;
        $d = $this->data;
        $n = self::u16($d, $offset + 4);
        for ($i = 0; $i < $n; $i++) {
            $p = $offset + 12 + $i * 16;
            $this->tab[substr($d, $p, 4)] = [self::u32($d, $p + 8), self::u32($d, $p + 12)];
        }
        foreach (['head', 'hhea', 'maxp', 'hmtx', 'cmap', 'glyf', 'loca'] as $t) {
            if (!isset($this->tab[$t])) {
                throw new DocxPdfException("Font senza tabella $t (non TrueType): $file");
            }
        }
        $h = $this->tab['head'][0];
        $this->upm = max(16, self::u16($d, $h + 18));
        $u = $this->upm;
        $this->bbox = [self::s16($d, $h + 36), self::s16($d, $h + 38), self::s16($d, $h + 40), self::s16($d, $h + 42)];
        $mac = self::u16($d, $h + 44);
        $this->bold = (bool) ($mac & 1);
        $this->italic = (bool) ($mac & 2);
        $this->locFmt = self::s16($d, $h + 50);

        $hh = $this->tab['hhea'][0];
        $hAsc = self::s16($d, $hh + 4);
        $hDesc = -self::s16($d, $hh + 6);
        $hGap = self::s16($d, $hh + 8);
        $this->nHM = self::u16($d, $hh + 34);
        $this->numGlyphs = self::u16($d, $this->tab['maxp'][0] + 4);
        $this->hmtx = substr($d, $this->tab['hmtx'][0], $this->tab['hmtx'][1]);

        $asc = $hAsc;
        $desc = $hDesc;
        $lead = $hGap;
        if (isset($this->tab['OS/2']) && $this->tab['OS/2'][1] >= 78) {
            $o = $this->tab['OS/2'][0];
            $fs = self::u16($d, $o + 62);
            $winA = self::u16($d, $o + 74);
            $winD = self::u16($d, $o + 76);
            if ($fs & 0x80) {
                $asc = self::s16($d, $o + 68);
                $desc = -self::s16($d, $o + 70);
                $lead = self::s16($d, $o + 72);
            } elseif ($winA + $winD > 0) {
                // Regola usata da Word su Windows
                $asc = $winA;
                $desc = $winD;
                $lead = max(0, $hGap - (($winA + $winD) - ($hAsc + $hDesc)));
            }
            if (self::u16($d, $o) >= 2 && $this->tab['OS/2'][1] >= 90) {
                $this->capH = self::s16($d, $o + 88) / $u;
            }
            if ($fs & 0x20) {
                $this->bold = true;
            }
            if ($fs & 0x01) {
                $this->italic = true;
            }
        }
        $this->asc = $asc / $u;
        $this->desc = $desc / $u;
        $this->lead = $lead / $u;

        if (isset($this->tab['post'])) {
            $p = $this->tab['post'][0];
            $this->italicAngle = self::s16($d, $p + 4) + self::u16($d, $p + 6) / 65536;
            $this->ulPos = self::s16($d, $p + 8) / $u;
            $this->ulThick = max(0.02, self::s16($d, $p + 10) / $u);
        }
        $names = self::names($d, $this->tab['name'] ?? null);
        $this->family = $names[1] ?? basename($file);
        $this->psName = preg_replace('/[^A-Za-z0-9\-]/', '', $names[6] ?? $this->family) ?: 'Font';
        $this->loadCmap();
    }

    public static function u16(string $d, int $p): int
    {
        return (ord($d[$p]) << 8) | ord($d[$p + 1]);
    }

    public static function s16(string $d, int $p): int
    {
        $v = (ord($d[$p]) << 8) | ord($d[$p + 1]);
        return $v >= 0x8000 ? $v - 0x10000 : $v;
    }

    public static function u32(string $d, int $p): int
    {
        return (ord($d[$p]) << 24) | (ord($d[$p + 1]) << 16) | (ord($d[$p + 2]) << 8) | ord($d[$p + 3]);
    }

    /** Legge la tabella "name": id => stringa (1 famiglia, 2 stile, 6 nome PostScript). */
    public static function names(string $d, ?array $tab, int $shift = 0): array
    {
        $out = [];
        if ($tab === null) {
            return $out;
        }
        $t = $tab[0] - $shift;
        if ($t < 0 || $t + 6 > strlen($d)) {
            return $out;
        }
        $cnt = self::u16($d, $t + 2);
        $str = $t + self::u16($d, $t + 4);
        $rank = [];
        for ($i = 0; $i < $cnt; $i++) {
            $p = $t + 6 + $i * 12;
            if ($p + 12 > strlen($d)) {
                break;
            }
            $pid = self::u16($d, $p);
            $lang = self::u16($d, $p + 4);
            $id = self::u16($d, $p + 6);
            if (!in_array($id, [1, 2, 4, 6, 16, 17], true)) {
                continue;
            }
            $raw = substr($d, $str + self::u16($d, $p + 10), self::u16($d, $p + 8));
            if ($pid === 3 || $pid === 0) {
                $s = @mb_convert_encoding($raw, 'UTF-8', 'UTF-16BE');
                $r = ($pid === 3 && $lang === 0x409) ? 3 : 2;
            } else {
                $s = @mb_convert_encoding($raw, 'UTF-8', 'ISO-8859-1');
                $r = 1;
            }
            if ($s !== '' && $s !== false && $r > ($rank[$id] ?? 0)) {
                $out[$id] = $s;
                $rank[$id] = $r;
            }
        }
        return $out;
    }

    /**
     * Descrive i font contenuti in un file senza caricarlo tutto in memoria.
     * @return array<int, array{0:int,1:string,2:bool,3:bool}> offset, famiglia, grassetto, corsivo
     */
    public static function describe(string $file): array
    {
        $out = [];
        $fh = @fopen($file, 'rb');
        if (!$fh) {
            return $out;
        }
        $read = static function (int $pos, int $len) use ($fh): string {
            if ($len <= 0 || $len > 4000000) {
                return '';
            }
            fseek($fh, $pos);
            $s = fread($fh, $len);
            return $s === false ? '' : $s;
        };
        $head = $read(0, 12);
        $offsets = [0];
        if (strlen($head) === 12 && substr($head, 0, 4) === 'ttcf') {
            $n = min(64, self::u32($head, 8));
            $tbl = $read(12, $n * 4);
            $offsets = [];
            for ($i = 0; $i < $n && ($i + 1) * 4 <= strlen($tbl); $i++) {
                $offsets[] = self::u32($tbl, $i * 4);
            }
        }
        foreach ($offsets as $off) {
            $h = $read($off, 12);
            if (strlen($h) < 12) {
                continue;
            }
            $ver = substr($h, 0, 4);
            if ($ver !== "\x00\x01\x00\x00" && $ver !== 'true') {
                continue; // CFF/OpenType non supportato
            }
            $n = self::u16($h, 4);
            $dir = $read($off + 12, $n * 16);
            $tabs = [];
            for ($i = 0; $i < $n && ($i + 1) * 16 <= strlen($dir); $i++) {
                $tabs[substr($dir, $i * 16, 4)] = [self::u32($dir, $i * 16 + 8), self::u32($dir, $i * 16 + 12)];
            }
            if (!isset($tabs['name'], $tabs['head'], $tabs['glyf'])) {
                continue;
            }
            $nameData = $read($tabs['name'][0], $tabs['name'][1]);
            $names = self::names($nameData, [$tabs['name'][0], $tabs['name'][1]], $tabs['name'][0]);
            $family = $names[1] ?? '';
            if ($family === '') {
                continue;
            }
            $hd = $read($tabs['head'][0], 54);
            $mac = strlen($hd) >= 46 ? self::u16($hd, 44) : 0;
            $bold = (bool) ($mac & 1);
            $italic = (bool) ($mac & 2);
            if (isset($tabs['OS/2'])) {
                $os = $read($tabs['OS/2'][0], 64);
                if (strlen($os) >= 64) {
                    $fs = self::u16($os, 62);
                    $bold = $bold || ($fs & 0x20);
                    $italic = $italic || ($fs & 0x01);
                }
            }
            $out[] = [$off, $family, (bool) $bold, (bool) $italic];
        }
        fclose($fh);
        return $out;
    }

    private function loadCmap(): void
    {
        $d = $this->data;
        $c = $this->tab['cmap'][0];
        $n = self::u16($d, $c + 2);
        $best = -1;
        $bestOff = 0;
        $symbol = false;
        for ($i = 0; $i < $n; $i++) {
            $p = $c + 4 + $i * 8;
            $pid = self::u16($d, $p);
            $eid = self::u16($d, $p + 2);
            $score = match (true) {
                $pid === 3 && $eid === 10 => 5,
                $pid === 0 && ($eid === 4 || $eid === 6) => 4,
                $pid === 3 && $eid === 1 => 3,
                $pid === 0 => 2,
                $pid === 3 && $eid === 0 => 1,
                default => 0,
            };
            $fmt = self::u16($d, $c + self::u32($d, $p + 4));
            if (!in_array($fmt, [0, 4, 6, 12], true)) {
                continue;
            }
            if ($score > $best) {
                $best = $score;
                $bestOff = $c + self::u32($d, $p + 4);
                $symbol = ($pid === 3 && $eid === 0);
            }
        }
        if ($best < 0) {
            return;
        }
        $o = $bestOff;
        $fmt = self::u16($d, $o);
        $map = [];
        if ($fmt === 4) {
            $segs = intdiv(self::u16($d, $o + 6), 2);
            $end = array_values(unpack('n*', substr($d, $o + 14, $segs * 2)));
            $start = array_values(unpack('n*', substr($d, $o + 16 + $segs * 2, $segs * 2)));
            $delta = array_values(unpack('n*', substr($d, $o + 16 + $segs * 4, $segs * 2)));
            $roPos = $o + 16 + $segs * 6;
            $ro = array_values(unpack('n*', substr($d, $roPos, $segs * 2)));
            $len = strlen($d);
            for ($s = 0; $s < $segs; $s++) {
                $e = $end[$s];
                for ($ch = $start[$s]; $ch <= $e && $ch < 0xFFFF; $ch++) {
                    if ($ro[$s] === 0) {
                        $g = ($ch + $delta[$s]) & 0xFFFF;
                    } else {
                        $p = $roPos + $s * 2 + $ro[$s] + 2 * ($ch - $start[$s]);
                        if ($p + 2 > $len) {
                            continue;
                        }
                        $g = self::u16($d, $p);
                        if ($g !== 0) {
                            $g = ($g + $delta[$s]) & 0xFFFF;
                        }
                    }
                    if ($g !== 0) {
                        $map[$ch] = $g;
                    }
                }
            }
        } elseif ($fmt === 12) {
            $groups = self::u32($d, $o + 12);
            for ($g = 0; $g < $groups; $g++) {
                $p = $o + 16 + $g * 12;
                $a = self::u32($d, $p);
                $b = self::u32($d, $p + 4);
                $gid = self::u32($d, $p + 8);
                if ($b - $a > 70000) {
                    continue;
                }
                for ($ch = $a; $ch <= $b; $ch++) {
                    $map[$ch] = $gid + $ch - $a;
                }
            }
        } elseif ($fmt === 6) {
            $first = self::u16($d, $o + 6);
            $cnt = self::u16($d, $o + 8);
            for ($i = 0; $i < $cnt; $i++) {
                $g = self::u16($d, $o + 10 + $i * 2);
                if ($g) {
                    $map[$first + $i] = $g;
                }
            }
        } else {
            for ($i = 0; $i < 256; $i++) {
                $g = ord($d[$o + 6 + $i]);
                if ($g) {
                    $map[$i] = $g;
                }
            }
        }
        if ($symbol) {
            // Font simbolo (Wingdings, Symbol): i caratteri stanno in U+F0xx
            foreach ($map as $ch => $g) {
                if ($ch >= 0xF000 && $ch <= 0xF0FF && !isset($map[$ch - 0xF000])) {
                    $map[$ch - 0xF000] = $g;
                }
            }
        }
        $this->cmap = $map;
    }

    public function has(int $cp): bool
    {
        return isset($this->cmap[$cp]);
    }

    private function advance(int $gid): int
    {
        if (!isset($this->adv[$gid])) {
            $i = $gid < $this->nHM ? $gid : $this->nHM - 1;
            $this->adv[$gid] = $i >= 0 && $i * 4 + 2 <= strlen($this->hmtx) ? self::u16($this->hmtx, $i * 4) : 0;
        }
        return $this->adv[$gid];
    }

    /** Larghezza del testo in unità del font e numero di caratteri. */
    public function measure(string $s): array
    {
        if (isset($this->wcache[$s])) {
            return $this->wcache[$s];
        }
        $w = 0;
        $n = 0;
        foreach (self::codepoints($s) as $cp) {
            $w += $this->advance($this->cmap[$cp] ?? 0);
            $n++;
        }
        if (count($this->wcache) < 20000) {
            $this->wcache[$s] = [$w, $n];
        }
        return [$w, $n];
    }

    /** Stringa esadecimale dei glifi (codifica Identity-H) e registrazione dei glifi usati. */
    public function encode(string $s): string
    {
        $hex = '';
        foreach (self::codepoints($s) as $cp) {
            $g = $this->cmap[$cp] ?? 0;
            if (!isset($this->used[$g])) {
                $this->used[$g] = $cp;
            }
            $hex .= sprintf('%04X', $g);
        }
        return $hex;
    }

    public static function codepoints(string $s): array
    {
        if ($s === '') {
            return [];
        }
        if (!preg_match('/[\x80-\xFF]/', $s)) {
            return array_values(unpack('C*', $s));
        }
        $u = @mb_convert_encoding($s, 'UTF-32BE', 'UTF-8');
        return $u === false || $u === '' ? [] : array_values(unpack('N*', $u));
    }

    /** Larghezze per il PDF (1000 unità per em) dei glifi usati. */
    public function pdfWidths(): array
    {
        $out = [];
        $k = 1000 / $this->upm;
        foreach ($this->used as $g => $_) {
            $out[$g] = (int) round($this->advance($g) * $k);
        }
        ksort($out);
        return $out;
    }

    /** Costruisce un file TrueType ridotto ai soli glifi usati (stessi indici). */
    public function subset(): string
    {
        $d = $this->data;
        [$lo, $ll] = $this->tab['loca'];
        [$go, $gl] = $this->tab['glyf'];
        $ng = $this->numGlyphs;
        if ($this->locFmt === 0) {
            $loca = array_values(unpack('n*', substr($d, $lo, ($ng + 1) * 2)));
            foreach ($loca as $i => $v) {
                $loca[$i] = $v * 2;
            }
        } else {
            $loca = array_values(unpack('N*', substr($d, $lo, ($ng + 1) * 4)));
        }
        if (count($loca) < $ng + 1) {
            return substr($d, 0); // tabella anomala: incorpora il file intero
        }
        $need = [];
        $queue = array_keys($this->used);
        while ($queue) {
            $g = array_pop($queue);
            if (isset($need[$g]) || $g >= $ng) {
                continue;
            }
            $need[$g] = true;
            $len = $loca[$g + 1] - $loca[$g];
            if ($len < 10) {
                continue;
            }
            $p = $go + $loca[$g];
            if (self::s16($d, $p) >= 0) {
                continue;
            }
            // Glifo composto: servono anche i componenti
            $q = $p + 10;
            $endp = $p + $len;
            do {
                if ($q + 4 > $endp) {
                    break;
                }
                $flags = self::u16($d, $q);
                $queue[] = self::u16($d, $q + 2);
                $q += 4 + (($flags & 1) ? 4 : 2);
                if ($flags & 8) {
                    $q += 2;
                } elseif ($flags & 0x40) {
                    $q += 4;
                } elseif ($flags & 0x80) {
                    $q += 8;
                }
            } while ($flags & 0x20);
        }
        $glyf = '';
        $newLoca = [];
        for ($g = 0; $g < $ng; $g++) {
            $newLoca[] = strlen($glyf);
            if (isset($need[$g])) {
                $len = $loca[$g + 1] - $loca[$g];
                if ($len > 0) {
                    $glyf .= substr($d, $go + $loca[$g], $len);
                    if ($len % 4) {
                        $glyf .= str_repeat("\0", 4 - $len % 4);
                    }
                }
            }
        }
        $newLoca[] = strlen($glyf);
        if ($glyf === '') {
            $glyf = "\0\0\0\0";
        }
        $tables = ['glyf' => $glyf, 'loca' => pack('N*', ...$newLoca)];
        foreach (['head', 'hhea', 'maxp', 'hmtx', 'cvt ', 'fpgm', 'prep'] as $t) {
            if (isset($this->tab[$t])) {
                $tables[$t] = substr($d, $this->tab[$t][0], $this->tab[$t][1]);
            }
        }
        $tables['head'] = substr_replace($tables['head'], "\0\0\0\0", 8, 4);
        $tables['head'] = substr_replace($tables['head'], "\0\1", 50, 2);
        ksort($tables);

        $n = count($tables);
        $pow = 1;
        $sel = 0;
        while ($pow * 2 <= $n) {
            $pow *= 2;
            $sel++;
        }
        $out = pack('Nnnnn', 0x00010000, $n, $pow * 16, $sel, $n * 16 - $pow * 16);
        $offset = 12 + $n * 16;
        $body = '';
        foreach ($tables as $tag => $data) {
            $pad = (4 - strlen($data) % 4) % 4;
            $padded = $data . str_repeat("\0", $pad);
            $out .= $tag . pack('NNN', self::checksum($padded), $offset, strlen($data));
            $body .= $padded;
            $offset += strlen($padded);
        }
        return $out . $body;
    }

    private static function checksum(string $s): int
    {
        $sum = 0;
        foreach (unpack('N*', $s) as $v) {
            $sum = ($sum + $v) & 0xFFFFFFFF;
        }
        return $sum;
    }
}

// ======================================================================
// Catalogo dei font disponibili e regole di sostituzione
// ======================================================================

final class DocxPdfFontBook
{
    /** Sostituti con metriche uguali o vicine, in ordine di preferenza */
    private const SUBST = [
        'arial' => ['Liberation Sans', 'Arimo', 'Helvetica'],
        'helvetica' => ['Arial', 'Liberation Sans', 'Arimo'],
        'helvetica neue' => ['Arial', 'Liberation Sans'],
        'times new roman' => ['Liberation Serif', 'Tinos', 'Times'],
        'times' => ['Times New Roman', 'Liberation Serif'],
        'courier new' => ['Liberation Mono', 'Cousine', 'Courier'],
        'calibri' => ['Carlito'],
        'cambria' => ['Caladea'],
        'verdana' => ['DejaVu Sans'],
        'tahoma' => ['Liberation Sans', 'Arial'],
        'segoe ui' => ['Liberation Sans', 'Arial'],
        'georgia' => ['DejaVu Serif', 'Liberation Serif'],
        'symbol' => ['DejaVu Sans'],
        'wingdings' => ['DejaVu Sans'],
    ];

    private array $index = [];
    private array $faces = [];
    private array $fonts = [];
    /** Sostituzioni avvenute: famiglia richiesta => famiglia usata */
    public array $subst = [];

    public function __construct(array $dirs)
    {
        $dirs = array_values(array_unique(array_filter($dirs, static fn ($d) => is_string($d) && $d !== '' && is_dir($d))));
        $sig = '';
        foreach ($dirs as $d) {
            $sig .= $d . ':' . @filemtime($d) . ';';
        }
        $cache = rtrim(sys_get_temp_dir(), '/\\') . '/docxpdf_fonts_' . md5($sig . 'v2') . '.json';
        if (is_file($cache)) {
            $idx = json_decode((string) @file_get_contents($cache), true);
            if (is_array($idx)) {
                $this->index = $idx;
                return;
            }
        }
        foreach ($dirs as $d) {
            $this->scan($d, 0);
        }
        @file_put_contents($cache, json_encode($this->index), LOCK_EX);
    }

    private function scan(string $dir, int $depth): void
    {
        $list = @scandir($dir);
        if ($list === false) {
            return;
        }
        foreach ($list as $name) {
            if ($name[0] === '.') {
                continue;
            }
            $path = $dir . '/' . $name;
            if (is_dir($path)) {
                if ($depth < 4) {
                    $this->scan($path, $depth + 1);
                }
                continue;
            }
            $ext = strtolower(pathinfo($name, PATHINFO_EXTENSION));
            if ($ext !== 'ttf' && $ext !== 'ttc' && $ext !== 'otf') {
                continue;
            }
            foreach (DocxPdfFont::describe($path) as [$off, $family, $bold, $italic]) {
                $key = self::key($family, $bold, $italic);
                if (!isset($this->index[$key])) {
                    $this->index[$key] = [$path, $off];
                }
            }
        }
    }

    private static function key(string $family, bool $b, bool $i): string
    {
        return mb_strtolower(trim($family)) . '|' . ($b ? 'b' : '') . ($i ? 'i' : '');
    }

    private function load(string $key): ?DocxPdfFont
    {
        if (!array_key_exists($key, $this->fonts)) {
            $this->fonts[$key] = null;
            if (isset($this->index[$key])) {
                try {
                    $this->fonts[$key] = new DocxPdfFont($this->index[$key][0], $this->index[$key][1]);
                } catch (Throwable $e) {
                    unset($this->index[$key]);
                }
            }
        }
        return $this->fonts[$key];
    }

    /**
     * Trova il font per famiglia e stile.
     * @return array{0:DocxPdfFont,1:bool,2:bool} font, grassetto simulato, corsivo simulato
     */
    public function face(string $family, bool $b, bool $i): array
    {
        $ck = self::key($family, $b, $i);
        if (isset($this->faces[$ck])) {
            return $this->faces[$ck];
        }
        $low = mb_strtolower(trim($family));
        $cands = array_merge([$family], self::SUBST[$low] ?? []);
        if (preg_match('/times|serif|garamond|georgia|cambria|antiqua|palatino|roman|bodoni|century school/i', $family)
            && !preg_match('/sans/i', $family)) {
            array_push($cands, 'Liberation Serif', 'Times New Roman', 'DejaVu Serif');
        } elseif (preg_match('/courier|consolas|mono|menlo|typewriter/i', $family)) {
            array_push($cands, 'Liberation Mono', 'Courier New', 'DejaVu Sans Mono');
        }
        array_push($cands, 'Liberation Sans', 'Arial', 'Helvetica', 'DejaVu Sans');

        foreach ($cands as $n => $cand) {
            foreach ([[$b, $i], [$b, false], [false, $i], [false, false]] as [$tb, $ti]) {
                $font = $this->load(self::key($cand, $tb, $ti));
                if ($font !== null) {
                    if ($n > 0) {
                        $this->subst[$family] = $font->family;
                    }
                    return $this->faces[$ck] = [$font, $b && !$tb, $i && !$ti];
                }
            }
        }
        // Ultima risorsa: il primo font utilizzabile dell'indice
        foreach (array_keys($this->index) as $key) {
            $font = $this->load($key);
            if ($font !== null) {
                $this->subst[$family] = $font->family;
                return $this->faces[$ck] = [$font, $b && !$font->bold, $i && !$font->italic];
            }
        }
        throw new DocxPdfException(
            'Nessun font TrueType disponibile: metti dei file .ttf nella cartella "fonts" accanto a DocxPdf.php'
        );
    }

    /** Font alternativo che contiene il carattere (simboli, caselle, frecce…). */
    public function fallback(int $cp, bool $b, bool $i): ?DocxPdfFont
    {
        foreach (['DejaVu Sans', 'Arial Unicode MS', 'Segoe UI Symbol', 'Apple Symbols', 'Liberation Sans', 'Arial'] as $fam) {
            foreach ([[$b, $i], [false, false]] as [$tb, $ti]) {
                $font = $this->load(self::key($fam, $tb, $ti));
                if ($font !== null && $font->has($cp)) {
                    return $font;
                }
            }
        }
        return null;
    }
}

// ======================================================================
// Scrittura del PDF
// ======================================================================

final class DocxPdfWriter
{
    private array $objs = [];
    private array $pages = [];
    private array $fonts = [];
    private array $images = [];
    public array $warn = [];

    public static function n(float $v): string
    {
        $s = number_format($v, 3, '.', '');
        return rtrim(rtrim($s, '0'), '.') ?: '0';
    }

    private function obj(string $body): int
    {
        $this->objs[] = $body;
        return count($this->objs);
    }

    private function alloc(): int
    {
        $this->objs[] = '';
        return count($this->objs);
    }

    private function stream(string $dict, string $data, bool $compress = true): int
    {
        if ($compress) {
            $data = gzcompress($data, 6);
            $dict .= ' /Filter /FlateDecode';
        }
        return $this->obj('<<' . $dict . ' /Length ' . strlen($data) . ">>\nstream\n" . $data . "\nendstream");
    }

    public function fontRes(DocxPdfFont $f): string
    {
        if ($f->res === '') {
            $f->res = 'F' . (count($this->fonts) + 1);
            $this->fonts[$f->res] = $f;
        }
        return $f->res;
    }

    public function addPage(float $w, float $h, string $content): void
    {
        $this->pages[] = [$w, $h, $content];
    }

    public function resetPages(): void
    {
        $this->pages = [];
    }

    /**
     * Registra un'immagine (PNG o JPEG).
     * @return array{0:string,1:int,2:int}|null nome risorsa, larghezza e altezza in pixel
     */
    public function image(string $key, string $bytes): ?array
    {
        if (array_key_exists($key, $this->images)) {
            return $this->images[$key];
        }
        $this->images[$key] = null;
        $res = 'I' . count($this->images);
        if (strncmp($bytes, "\x89PNG\r\n\x1a\n", 8) === 0) {
            $img = $this->png($bytes);
        } elseif (strncmp($bytes, "\xFF\xD8", 2) === 0) {
            $img = $this->jpeg($bytes);
        } else {
            $img = $this->viaGd($bytes);
        }
        if ($img === null) {
            return null;
        }
        return $this->images[$key] = [$res, $img[1], $img[2], $img[0]];
    }

    private function jpeg(string $b): ?array
    {
        $p = 2;
        $len = strlen($b);
        while ($p + 9 < $len) {
            if ($b[$p] !== "\xFF") {
                $p++;
                continue;
            }
            $m = ord($b[$p + 1]);
            if ($m === 0xFF) {
                $p++;
                continue;
            }
            if ($m >= 0xC0 && $m <= 0xCF && $m !== 0xC4 && $m !== 0xC8 && $m !== 0xCC) {
                $h = DocxPdfFont::u16($b, $p + 5);
                $w = DocxPdfFont::u16($b, $p + 7);
                $c = ord($b[$p + 9]);
                $cs = $c === 1 ? '/DeviceGray' : ($c === 4 ? '/DeviceCMYK /Decode [1 0 1 0 1 0 1 0]' : '/DeviceRGB');
                $id = $this->stream(
                    " /Type /XObject /Subtype /Image /Width $w /Height $h /ColorSpace $cs /BitsPerComponent 8 /Filter /DCTDecode",
                    $b,
                    false
                );
                return [$id, $w, $h];
            }
            if ($m === 0xD8 || ($m >= 0xD0 && $m <= 0xD7) || $m === 0x01) {
                $p += 2;
                continue;
            }
            $p += 2 + DocxPdfFont::u16($b, $p + 2);
        }
        return null;
    }

    private function png(string $b, bool $retry = true): ?array
    {
        $p = 8;
        $len = strlen($b);
        $ihdr = null;
        $idat = '';
        $plte = '';
        $trns = null;
        while ($p + 8 <= $len) {
            $l = DocxPdfFont::u32($b, $p);
            $type = substr($b, $p + 4, 4);
            $data = substr($b, $p + 8, $l);
            if ($type === 'IHDR') {
                $ihdr = $data;
            } elseif ($type === 'PLTE') {
                $plte = $data;
            } elseif ($type === 'tRNS') {
                $trns = $data;
            } elseif ($type === 'IDAT') {
                $idat .= $data;
            } elseif ($type === 'IEND') {
                break;
            }
            $p += 12 + $l;
        }
        if ($ihdr === null || strlen($ihdr) < 13 || $idat === '') {
            return null;
        }
        $w = DocxPdfFont::u32($ihdr, 0);
        $h = DocxPdfFont::u32($ihdr, 4);
        $bd = ord($ihdr[8]);
        $ct = ord($ihdr[9]);
        if (ord($ihdr[12]) !== 0) {
            return $retry ? $this->viaGd($b) : null; // interlacciato
        }
        $colors = [0 => 1, 2 => 3, 3 => 1, 4 => 1, 6 => 3][$ct] ?? null;
        if ($colors === null) {
            return null;
        }
        $cs = match ($ct) {
            0, 4 => '/DeviceGray',
            2, 6 => '/DeviceRGB',
            default => '[/Indexed /DeviceRGB ' . (intdiv(strlen($plte), 3) - 1) . ' <' . bin2hex($plte) . '>]',
        };
        $hasAlpha = $ct === 4 || $ct === 6;
        $palAlpha = $ct === 3 && $trns !== null && $bd === 8;
        if (!$hasAlpha && !$palAlpha) {
            // Nessuna trasparenza: i dati compressi del PNG vanno bene così come sono
            $id = $this->obj(
                "<< /Type /XObject /Subtype /Image /Width $w /Height $h /ColorSpace $cs /BitsPerComponent $bd"
                . " /Filter /FlateDecode /DecodeParms << /Predictor 15 /Colors $colors /BitsPerComponent $bd /Columns $w >>"
                . ' /Length ' . strlen($idat) . " >>\nstream\n" . $idat . "\nendstream"
            );
            return [$id, $w, $h];
        }
        $raw = @gzuncompress($idat);
        if ($raw === false) {
            return null;
        }
        $bytes = intdiv($bd, 8);
        $bpp = ($colors + ($hasAlpha ? 1 : 0)) * $bytes;
        $stride = $w * $bpp;
        if (strlen($raw) < ($stride + 1) * $h) {
            return null;
        }
        $color = '';
        $alpha = '';
        $prev = array_fill(0, $stride, 0);
        $cb = $colors * $bytes;
        $reC = '/(.{' . $cb . '}).{' . $bytes . '}/s';
        $reA = '/.{' . $cb . '}(.{' . $bytes . '})/s';
        $tr = $palAlpha ? str_pad($trns, 256, "\xFF") : '';
        for ($y = 0, $pos = 0; $y < $h; $y++, $pos += $stride + 1) {
            $ft = ord($raw[$pos]);
            $line = substr($raw, $pos + 1, $stride);
            if ($ft !== 0) {
                $row = array_values(unpack('C*', $line));
                switch ($ft) {
                    case 1:
                        for ($i = $bpp; $i < $stride; $i++) {
                            $row[$i] = ($row[$i] + $row[$i - $bpp]) & 255;
                        }
                        break;
                    case 2:
                        for ($i = 0; $i < $stride; $i++) {
                            $row[$i] = ($row[$i] + $prev[$i]) & 255;
                        }
                        break;
                    case 3:
                        for ($i = 0; $i < $stride; $i++) {
                            $row[$i] = ($row[$i] + ((($i >= $bpp ? $row[$i - $bpp] : 0) + $prev[$i]) >> 1)) & 255;
                        }
                        break;
                    case 4:
                        for ($i = 0; $i < $stride; $i++) {
                            $a = $i >= $bpp ? $row[$i - $bpp] : 0;
                            $bb = $prev[$i];
                            $c = $i >= $bpp ? $prev[$i - $bpp] : 0;
                            $pa = abs($bb - $c);
                            $pb = abs($a - $c);
                            $pc = abs($a + $bb - 2 * $c);
                            $pr = ($pa <= $pb && $pa <= $pc) ? $a : ($pb <= $pc ? $bb : $c);
                            $row[$i] = ($row[$i] + $pr) & 255;
                        }
                        break;
                }
                $prev = $row;
                $line = pack('C*', ...$row);
            } else {
                $prev = array_values(unpack('C*', $line));
            }
            if ($hasAlpha) {
                $color .= preg_replace($reC, '$1', $line);
                $alpha .= preg_replace($reA, '$1', $line);
            } else {
                $color .= $line;
                $alpha .= strtr($line, implode('', array_map('chr', range(0, 255))), $tr);
            }
        }
        $mask = $this->stream(
            " /Type /XObject /Subtype /Image /Width $w /Height $h /ColorSpace /DeviceGray /BitsPerComponent $bd",
            $alpha
        );
        $id = $this->stream(
            " /Type /XObject /Subtype /Image /Width $w /Height $h /ColorSpace $cs /BitsPerComponent $bd /SMask $mask 0 R",
            $color
        );
        return [$id, $w, $h];
    }

    /** Formati non gestiti direttamente (GIF, BMP, PNG interlacciati): conversione con GD, se presente. */
    private function viaGd(string $bytes): ?array
    {
        if (!function_exists('imagecreatefromstring')) {
            return null;
        }
        $im = @imagecreatefromstring($bytes);
        if (!$im) {
            return null;
        }
        imagepalettetotruecolor($im);
        imagesavealpha($im, true);
        imageinterlace($im, false);
        ob_start();
        imagepng($im, null, 6);
        $png = (string) ob_get_clean();
        return $this->png($png, false);
    }

    public function output(string $title = ''): string
    {
        // Font
        $fontDict = '';
        foreach ($this->fonts as $res => $f) {
            $tag = '';
            $seed = crc32($f->psName . count($f->used));
            for ($i = 0; $i < 6; $i++) {
                $tag .= chr(65 + ($seed % 26));
                $seed = intdiv($seed, 26) + 7;
            }
            $name = $tag . '+' . $f->psName;
            $file = $f->subset();
            $fileId = $this->stream(' /Length1 ' . strlen($file), $file);
            $k = 1000 / $f->upm;
            $bbox = implode(' ', array_map(static fn ($v) => (int) round($v * $k), $f->bbox));
            $desc = $this->obj(
                "<< /Type /FontDescriptor /FontName /$name /Flags " . ($f->italic ? 96 : 32)
                . " /FontBBox [$bbox] /ItalicAngle " . self::n($f->italicAngle)
                . ' /Ascent ' . (int) round($f->asc * 1000) . ' /Descent ' . (int) round(-$f->desc * 1000)
                . ' /CapHeight ' . (int) round($f->capH * 1000) . ' /StemV ' . ($f->bold ? 120 : 80)
                . " /FontFile2 $fileId 0 R >>"
            );
            $w = '';
            foreach ($f->pdfWidths() as $g => $adv) {
                $w .= "$g [$adv] ";
            }
            $cid = $this->obj(
                "<< /Type /Font /Subtype /CIDFontType2 /BaseFont /$name"
                . ' /CIDSystemInfo << /Registry (Adobe) /Ordering (Identity) /Supplement 0 >>'
                . " /FontDescriptor $desc 0 R /DW 0 /W [ $w] /CIDToGIDMap /Identity >>"
            );
            $map = [];
            foreach ($f->used as $g => $cp) {
                if ($g === 0 || $cp === 0) {
                    continue;
                }
                $u = bin2hex(mb_convert_encoding(mb_chr($cp, 'UTF-8'), 'UTF-16BE', 'UTF-8'));
                $map[] = sprintf('<%04X> <%s>', $g, strtoupper($u));
            }
            $cmap = "/CIDInit /ProcSet findresource begin\n12 dict begin\nbegincmap\n"
                . "/CIDSystemInfo << /Registry (Adobe) /Ordering (UCS) /Supplement 0 >> def\n"
                . "/CMapName /Adobe-Identity-UCS def\n/CMapType 2 def\n"
                . "1 begincodespacerange\n<0000> <FFFF>\nendcodespacerange\n";
            foreach (array_chunk($map, 100) as $chunk) {
                $cmap .= count($chunk) . " beginbfchar\n" . implode("\n", $chunk) . "\nendbfchar\n";
            }
            $cmap .= "endcmap\nCMapName currentdict /CMap defineresource pop\nend\nend";
            $tu = $this->stream('', $cmap);
            $fid = $this->obj(
                "<< /Type /Font /Subtype /Type0 /BaseFont /$name /Encoding /Identity-H"
                . " /DescendantFonts [$cid 0 R] /ToUnicode $tu 0 R >>"
            );
            $fontDict .= "/$res $fid 0 R ";
        }
        $xo = '';
        foreach ($this->images as $img) {
            if ($img !== null) {
                $xo .= "/{$img[0]} {$img[3]} 0 R ";
            }
        }
        $resId = $this->obj("<< /Font << $fontDict>> /XObject << $xo>> /ProcSet [/PDF /Text /ImageB /ImageC /ImageI] >>");

        $pagesId = $this->alloc();
        $kids = '';
        foreach ($this->pages as [$w, $h, $content]) {
            $cid = $this->stream('', $content);
            $pid = $this->obj(
                "<< /Type /Page /Parent $pagesId 0 R /MediaBox [0 0 " . self::n($w) . ' ' . self::n($h)
                . "] /Resources $resId 0 R /Contents $cid 0 R >>"
            );
            $kids .= "$pid 0 R ";
        }
        $this->objs[$pagesId - 1] = "<< /Type /Pages /Kids [ $kids] /Count " . count($this->pages) . ' >>';
        $catalog = $this->obj("<< /Type /Catalog /Pages $pagesId 0 R >>");
        $t = '';
        if ($title !== '') {
            $t = ' /Title <FEFF' . strtoupper(bin2hex(mb_convert_encoding($title, 'UTF-16BE', 'UTF-8'))) . '>';
        }
        $info = $this->obj('<< /Producer (DocxPdf PHP) /CreationDate (D:' . date('YmdHis') . ')' . $t . ' >>');

        $out = "%PDF-1.5\n%\xE2\xE3\xCF\xD3\n";
        $xref = [];
        foreach ($this->objs as $i => $body) {
            $xref[] = strlen($out);
            $out .= ($i + 1) . " 0 obj\n" . $body . "\nendobj\n";
        }
        $start = strlen($out);
        $n = count($this->objs) + 1;
        $out .= "xref\n0 $n\n0000000000 65535 f \n";
        foreach ($xref as $off) {
            $out .= sprintf("%010d 00000 n \n", $off);
        }
        $out .= "trailer\n<< /Size $n /Root $catalog 0 R /Info $info 0 R >>\nstartxref\n$start\n%%EOF\n";
        return $out;
    }
}

// ======================================================================
// Formato di carattere risolto (font reale, misure in punti)
// ======================================================================

final class DocxPdfFmt
{
    public DocxPdfFont $font;
    public float $size = 10.0;
    public bool $fakeB = false;
    public bool $fakeI = false;
    public bool $u = false;
    public bool $strike = false;
    public bool $caps = false;
    public bool $small = false;
    public string $color = '000000';
    public ?string $hl = null;
    public float $sp = 0.0;
    public float $rise = 0.0;
    public float $scale = 1.0;
    public float $asc = 0.0;
    public float $desc = 0.0;
    public float $lead = 0.0;
    /** Proprietà di partenza, per derivare varianti */
    public array $src = [];

    public function width(string $s): float
    {
        [$w, $n] = $this->font->measure($s);
        return ($w * $this->size / $this->font->upm + $n * $this->sp) * $this->scale;
    }
}

// ======================================================================
// Motore: lettura del DOCX, impaginazione, disegno
// ======================================================================

final class DocxPdfEngine
{
    private const NS = [
        'http://schemas.openxmlformats.org/wordprocessingml/2006/main' => 'w',
        'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing' => 'wp',
        'http://schemas.openxmlformats.org/drawingml/2006/main' => 'a',
        'http://schemas.openxmlformats.org/drawingml/2006/picture' => 'pic',
        'http://schemas.microsoft.com/office/word/2010/wordprocessingShape' => 'wps',
        'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup' => 'wpg',
        'http://schemas.openxmlformats.org/markup-compatibility/2006' => 'mc',
        'urn:schemas-microsoft-com:vml' => 'v',
    ];
    private const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
    private const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
    private const HL = [
        'yellow' => 'FFFF00', 'green' => '00FF00', 'cyan' => '00FFFF', 'magenta' => 'FF00FF',
        'blue' => '0000FF', 'red' => 'FF0000', 'darkBlue' => '000080', 'darkCyan' => '008080',
        'darkGreen' => '008000', 'darkMagenta' => '800080', 'darkRed' => '800000', 'darkYellow' => '808000',
        'darkGray' => '808080', 'lightGray' => 'C0C0C0', 'black' => '000000', 'white' => 'FFFFFF',
    ];
    /** Caratteri dei font simbolo (Symbol, Wingdings) → equivalenti Unicode */
    private const SYM = [
        0xF0B7 => 0x2022, 0xF0A7 => 0x25AA, 0xF0D8 => 0x27A2, 0xF0FC => 0x2713, 0xF0FE => 0x2612,
        0xF0A8 => 0x2610, 0xF06F => 0x2610, 0xF071 => 0x2751, 0xF076 => 0x2756, 0xF0E0 => 0x2192,
        0xF02D => 0x2013, 0xF06E => 0x25A0, 0xF0A1 => 0x25CB, 0xF06C => 0x25CF, 0xF0FB => 0x2717,
    ];

    public array $warn = [];

    private ZipArchive $zip;
    private DocxPdfFontBook $book;
    private DocxPdfWriter $pdf;
    private string $decimal;
    private string $title = '';

    private array $styles = [];
    private array $styleCache = [];
    private ?string $defPStyle = null;
    private array $defR = ['font' => 'Times New Roman', 'size' => 10.0];
    private array $defP = [];
    private array $nums = [];
    private array $absNums = [];
    private array $counters = [];
    private array $themeFonts = ['major' => 'Calibri Light', 'minor' => 'Calibri'];
    private array $themeColors = [];
    private float $defTab = 36.0;
    private int $compat = 12;
    private bool $evenOdd = false;
    private bool $expandBr = true;
    /** Tra due paragrafi vale lo spazio maggiore tra "dopo" e "prima" (regola predefinita di Word) */
    private bool $collapse = true;
    private array $fmts = [];
    private array $rels = [];
    private string $part = 'word/document.xml';
    private array $hfCache = [];
    private array $hfH = [];
    private array $sections = [];

    // Stato dell'impaginazione
    private array $pages = [];
    private int $pi = -1;
    private float $ph = 842.0;
    private string $layer = 'main';
    private float $y = 0.0;
    private float $top = 0.0;
    private float $bot = 0.0;
    private bool $pgEmpty = true;
    private bool $autoBreak = false;
    private array $sec = [];
    private int $secPages = 0;
    private array $fv = [];
    private array $dfv = [];
    private bool $needTotal = false;
    private ?int $total = null;
    /** Il segno di paragrafo contribuisce all'altezza dell'ultima riga (come in Word) */
    private bool $markInLine = true;

    public function __construct(string $path, array $opt)
    {
        $this->zip = new ZipArchive();
        if ($this->zip->open($path) !== true || $this->zip->locateName('word/document.xml') === false) {
            throw new DocxPdfException("Il file non è un DOCX valido: $path");
        }
        $this->decimal = (string) ($opt['decimal'] ?? '');
        $this->markInLine = (bool) ($opt['markInLine'] ?? true);
        $this->title = (string) ($opt['title'] ?? pathinfo($path, PATHINFO_FILENAME));
        $home = (string) (getenv('HOME') ?: '');
        $dirs = array_merge((array) ($opt['fonts'] ?? []), [
            __DIR__ . '/fonts',
            '/System/Library/Fonts/Supplemental', '/Library/Fonts', '/System/Library/Fonts',
            $home . '/Library/Fonts',
            '/Applications/Microsoft Word.app/Contents/Resources/DFonts',
            'C:/Windows/Fonts',
            '/usr/share/fonts', '/usr/local/share/fonts', $home . '/.fonts', $home . '/.local/share/fonts',
        ]);
        $this->book = new DocxPdfFontBook($dirs);
        $this->pdf = new DocxPdfWriter();
        $this->loadTheme();
        $this->loadSettings();
        if ($this->decimal === '') {
            $this->decimal = ',';
        }
        $this->loadStyles();
        $this->loadNumbering();
    }

    // ------------------------------------------------------------------
    // Utilità XML
    // ------------------------------------------------------------------

    private function xml(string $name): ?DOMElement
    {
        $s = $this->zip->getFromName($name);
        if ($s === false) {
            return null;
        }
        $d = new DOMDocument();
        if (!@$d->loadXML($s, LIBXML_NONET | LIBXML_PARSEHUGE | LIBXML_COMPACT)) {
            return null;
        }
        return $d->documentElement;
    }

    private static function nm(DOMNode $n): string
    {
        return (self::NS[$n->namespaceURI] ?? '?') . ':' . $n->localName;
    }

    private static function kid(?DOMNode $e, string $name): ?DOMElement
    {
        if ($e === null) {
            return null;
        }
        foreach ($e->childNodes as $c) {
            if ($c->nodeType === XML_ELEMENT_NODE && self::nm($c) === $name) {
                return $c;
            }
        }
        return null;
    }

    /** @return DOMElement[] */
    private static function kids(?DOMNode $e, ?string $name = null): array
    {
        $out = [];
        if ($e !== null) {
            foreach ($e->childNodes as $c) {
                if ($c->nodeType === XML_ELEMENT_NODE && ($name === null || self::nm($c) === $name)) {
                    $out[] = $c;
                }
            }
        }
        return $out;
    }

    private static function wa(?DOMElement $e, string $attr): ?string
    {
        if ($e === null || !$e->hasAttributeNS(self::NS_W, $attr)) {
            return null;
        }
        return $e->getAttributeNS(self::NS_W, $attr);
    }

    private static function val(?DOMNode $parent, string $child, string $attr = 'val'): ?string
    {
        return self::wa(self::kid($parent, $child), $attr);
    }

    private static function on(DOMElement $e): bool
    {
        $v = self::wa($e, 'val');
        return !($v === '0' || $v === 'false' || $v === 'off' || $v === 'none');
    }

    private function relsOf(string $part): array
    {
        if (!isset($this->rels[$part])) {
            $map = [];
            $root = $this->xml(dirname($part) . '/_rels/' . basename($part) . '.rels');
            if ($root !== null) {
                foreach ($root->childNodes as $r) {
                    if ($r->nodeType !== XML_ELEMENT_NODE || $r->getAttribute('TargetMode') === 'External') {
                        continue;
                    }
                    $t = $r->getAttribute('Target');
                    $t = $t !== '' && $t[0] === '/' ? ltrim($t, '/') : dirname($part) . '/' . $t;
                    // Risolve eventuali "../"
                    $segs = [];
                    foreach (explode('/', $t) as $seg) {
                        if ($seg === '..') {
                            array_pop($segs);
                        } elseif ($seg !== '.' && $seg !== '') {
                            $segs[] = $seg;
                        }
                    }
                    $map[$r->getAttribute('Id')] = implode('/', $segs);
                }
            }
            $this->rels[$part] = $map;
        }
        return $this->rels[$part];
    }

    private function note(string $msg): void
    {
        $this->warn[] = $msg;
    }

    // ------------------------------------------------------------------
    // Tema, impostazioni, stili, numerazioni
    // ------------------------------------------------------------------

    private function loadTheme(): void
    {
        $root = $this->xml('word/theme/theme1.xml');
        if ($root === null) {
            return;
        }
        $el = self::kid($root, 'a:themeElements');
        $fs = self::kid($el, 'a:fontScheme');
        foreach (['major' => 'a:majorFont', 'minor' => 'a:minorFont'] as $k => $tag) {
            $latin = self::kid(self::kid($fs, $tag), 'a:latin');
            if ($latin !== null && $latin->getAttribute('typeface') !== '') {
                $this->themeFonts[$k] = $latin->getAttribute('typeface');
            }
        }
        foreach (self::kids(self::kid($el, 'a:clrScheme')) as $c) {
            foreach (self::kids($c) as $v) {
                // sysClr ha val="window"/"windowText" (nomi, non colori): il colore è in lastClr
                $hex = $v->localName === 'sysClr' ? $v->getAttribute('lastClr') : $v->getAttribute('val');
                if (preg_match('/^[0-9A-Fa-f]{6}$/', $hex) === 1) {
                    $this->themeColors[$c->localName] = strtoupper($hex);
                }
            }
        }
    }

    private function loadSettings(): void
    {
        $root = $this->xml('word/settings.xml');
        if ($root === null) {
            return;
        }
        $v = self::val($root, 'w:defaultTabStop');
        if ($v !== null && (float) $v > 0) {
            $this->defTab = (float) $v / 20;
        }
        if ($this->decimal === '') {
            $this->decimal = (string) (self::val($root, 'w:decimalSymbol') ?? ',');
        }
        $eo = self::kid($root, 'w:evenAndOddHeaders');
        $this->evenOdd = $eo !== null && self::on($eo);
        $compat = self::kid($root, 'w:compat');
        foreach (self::kids($compat) as $c) {
            if ($c->localName === 'compatSetting' && self::wa($c, 'name') === 'compatibilityMode') {
                $this->compat = (int) self::wa($c, 'val');
            } elseif ($c->localName === 'doNotExpandShiftReturn' && self::on($c)) {
                $this->expandBr = false;
            } elseif ($c->localName === 'doNotUseHTMLParagraphAutoSpacing' && self::on($c)) {
                $this->collapse = false;
            }
        }
    }

    private function loadStyles(): void
    {
        $root = $this->xml('word/styles.xml');
        if ($root === null) {
            return;
        }
        $dd = self::kid($root, 'w:docDefaults');
        $this->defR = $this->rpr(self::kid(self::kid($dd, 'w:rPrDefault'), 'w:rPr')) + $this->defR;
        $this->defP = $this->ppr(self::kid(self::kid($dd, 'w:pPrDefault'), 'w:pPr'));
        foreach (self::kids($root, 'w:style') as $s) {
            $id = self::wa($s, 'styleId');
            if ($id === null) {
                continue;
            }
            $this->styles[$id] = $s;
            $def = self::wa($s, 'default');
            if (self::wa($s, 'type') === 'paragraph' && ($def === '1' || $def === 'true')) {
                $this->defPStyle = $id;
            }
        }
    }

    private function style(?string $id): array
    {
        $empty = ['p' => [], 'r' => [], 'tbl' => ['bdr' => [], 'mar' => [], 'cond' => []]];
        if ($id === null || !isset($this->styles[$id])) {
            return $empty;
        }
        if (isset($this->styleCache[$id])) {
            return $this->styleCache[$id];
        }
        $this->styleCache[$id] = $empty; // protezione dai cicli
        $el = $this->styles[$id];
        $base = $this->style(self::val($el, 'w:basedOn'));
        $p = $this->ppr(self::kid($el, 'w:pPr'));
        $tabs = array_merge($base['p']['tabs'] ?? [], $p['tabs'] ?? []);
        $pm = $p + $base['p'];
        unset($pm['markR'], $pm['sect'], $pm['style']);
        if ($tabs) {
            $pm['tabs'] = $tabs;
        }
        $tp = self::kid($el, 'w:tblPr');
        // Formattazione condizionale delle tabelle (prima riga, righe alterne, prima colonna…)
        $cond = $base['tbl']['cond'] ?? [];
        foreach (self::kids($el, 'w:tblStylePr') as $sp) {
            $type = (string) self::wa($sp, 'type');
            $tc = self::kid($sp, 'w:tcPr');
            $shd = (string) self::val($tc, 'w:shd', 'fill');
            $own = [
                'p' => array_diff_key($this->ppr(self::kid($sp, 'w:pPr')), ['markR' => 1, 'sect' => 1, 'style' => 1]),
                'r' => $this->rpr(self::kid($sp, 'w:rPr')),
                'bdr' => $this->bdrs(self::kid($tc, 'w:tcBorders')),
            ];
            if (preg_match('/^[0-9A-Fa-f]{6}$/', $shd)) {
                $own['shd'] = strtoupper($shd);
            }
            $old = $cond[$type] ?? ['p' => [], 'r' => [], 'bdr' => []];
            $cond[$type] = ['p' => $own['p'] + $old['p'], 'r' => $own['r'] + $old['r'], 'bdr' => $own['bdr'] + $old['bdr']]
                + (isset($own['shd']) ? ['shd' => $own['shd']] : (isset($old['shd']) ? ['shd' => $old['shd']] : []));
        }
        return $this->styleCache[$id] = [
            'p' => $pm,
            'r' => $this->rpr(self::kid($el, 'w:rPr')) + $base['r'],
            'tbl' => [
                'bdr' => $this->bdrs(self::kid($tp, 'w:tblBorders')) + $base['tbl']['bdr'],
                'mar' => $this->cellMar(self::kid($tp, 'w:tblCellMar')) + $base['tbl']['mar'],
                'cond' => $cond,
            ],
        ];
    }

    private function loadNumbering(): void
    {
        $root = $this->xml('word/numbering.xml');
        if ($root === null) {
            return;
        }
        foreach (self::kids($root, 'w:abstractNum') as $a) {
            $lv = [];
            foreach (self::kids($a, 'w:lvl') as $l) {
                $lv[(int) self::wa($l, 'ilvl')] = $this->numLevel($l);
            }
            $this->absNums[(string) self::wa($a, 'abstractNumId')] = $lv;
        }
        foreach (self::kids($root, 'w:num') as $n) {
            $ov = [];
            $own = false;
            foreach (self::kids($n, 'w:lvlOverride') as $o) {
                $il = (int) self::wa($o, 'ilvl');
                $l = self::kid($o, 'w:lvl');
                if ($l !== null) {
                    $ov[$il] = $this->numLevel($l);
                }
                $so = self::val($o, 'w:startOverride');
                if ($so !== null) {
                    $ov[$il]['start'] = (int) $so;
                    $own = true;
                }
            }
            $this->nums[(string) self::wa($n, 'numId')] = [
                'abs' => (string) self::val($n, 'w:abstractNumId'),
                'ov' => $ov,
                'own' => $own,
            ];
        }
    }

    private function numLevel(DOMElement $l): array
    {
        $p = $this->ppr(self::kid($l, 'w:pPr'));
        return [
            'start' => (int) (self::val($l, 'w:start') ?? 1),
            'fmt' => self::val($l, 'w:numFmt') ?? 'decimal',
            'text' => self::val($l, 'w:lvlText') ?? '',
            'suff' => self::val($l, 'w:suff') ?? 'tab',
            'p' => array_intersect_key($p, ['indL' => 1, 'indR' => 1, 'indF' => 1]),
            'r' => $this->rpr(self::kid($l, 'w:rPr')),
        ];
    }

    private static function numFormat(int $n, string $fmt): string
    {
        switch ($fmt) {
            case 'lowerLetter':
            case 'upperLetter':
                $s = '';
                $n = max(1, $n);
                $s = str_repeat(chr(97 + ($n - 1) % 26), intdiv($n - 1, 26) + 1);
                return $fmt === 'upperLetter' ? strtoupper($s) : $s;
            case 'lowerRoman':
            case 'upperRoman':
                $map = ['M' => 1000, 'CM' => 900, 'D' => 500, 'CD' => 400, 'C' => 100, 'XC' => 90,
                    'L' => 50, 'XL' => 40, 'X' => 10, 'IX' => 9, 'V' => 5, 'IV' => 4, 'I' => 1];
                $s = '';
                $n = max(1, min(3999, $n));
                foreach ($map as $r => $v) {
                    while ($n >= $v) {
                        $s .= $r;
                        $n -= $v;
                    }
                }
                return $fmt === 'lowerRoman' ? strtolower($s) : $s;
            case 'decimalZero':
                return sprintf('%02d', $n);
            case 'none':
                return '';
            default:
                return (string) $n;
        }
    }

    // ------------------------------------------------------------------
    // Proprietà di carattere e di paragrafo
    // ------------------------------------------------------------------

    private function rpr(?DOMElement $e): array
    {
        $o = [];
        if ($e === null) {
            return $o;
        }
        foreach ($e->childNodes as $c) {
            if ($c->nodeType !== XML_ELEMENT_NODE || $c->namespaceURI !== self::NS_W) {
                continue;
            }
            switch ($c->localName) {
                case 'rFonts':
                    $t = self::wa($c, 'asciiTheme') ?? self::wa($c, 'hAnsiTheme');
                    $f = $t !== null
                        ? $this->themeFonts[str_starts_with($t, 'major') ? 'major' : 'minor']
                        : (self::wa($c, 'ascii') ?? self::wa($c, 'hAnsi'));
                    if ($f !== null && $f !== '') {
                        $o['font'] = $f;
                    }
                    break;
                case 'sz':
                    $o['size'] = max(1.0, (float) self::wa($c, 'val') / 2);
                    break;
                case 'b':
                case 'i':
                case 'caps':
                case 'smallCaps':
                case 'vanish':
                    $o[$c->localName] = self::on($c);
                    break;
                case 'strike':
                case 'dstrike':
                    $o['strike'] = self::on($c);
                    break;
                case 'u':
                    $o['u'] = self::on($c);
                    break;
                case 'color':
                    $v = (string) self::wa($c, 'val');
                    $o['color'] = preg_match('/^[0-9A-Fa-f]{6}$/', $v) ? strtoupper($v) : null;
                    break;
                case 'highlight':
                    $o['hl'] = self::HL[(string) self::wa($c, 'val')] ?? null;
                    break;
                case 'shd':
                    $v = (string) self::wa($c, 'fill');
                    if (preg_match('/^[0-9A-Fa-f]{6}$/', $v)) {
                        $o['shd'] = strtoupper($v);
                    }
                    break;
                case 'vertAlign':
                    $o['va'] = self::wa($c, 'val');
                    break;
                case 'spacing':
                    $o['sp'] = (float) self::wa($c, 'val') / 20;
                    break;
                case 'position':
                    $o['pos'] = (float) self::wa($c, 'val') / 2;
                    break;
                case 'w':
                    $v = (float) self::wa($c, 'val');
                    if ($v > 0) {
                        $o['scale'] = $v;
                    }
                    break;
                case 'rStyle':
                    $o['rStyle'] = self::wa($c, 'val');
                    break;
            }
        }
        return $o;
    }

    private function ppr(?DOMElement $e): array
    {
        $o = [];
        if ($e === null) {
            return $o;
        }
        foreach ($e->childNodes as $c) {
            if ($c->nodeType !== XML_ELEMENT_NODE || $c->namespaceURI !== self::NS_W) {
                continue;
            }
            switch ($c->localName) {
                case 'pStyle':
                    $o['style'] = self::wa($c, 'val');
                    break;
                case 'jc':
                    $v = self::wa($c, 'val');
                    $o['jc'] = ['start' => 'left', 'end' => 'right', 'distribute' => 'both'][$v] ?? $v;
                    break;
                case 'ind':
                    $l = self::wa($c, 'left') ?? self::wa($c, 'start');
                    $r = self::wa($c, 'right') ?? self::wa($c, 'end');
                    if ($l !== null) {
                        $o['indL'] = (float) $l / 20;
                    }
                    if ($r !== null) {
                        $o['indR'] = (float) $r / 20;
                    }
                    if (($h = self::wa($c, 'hanging')) !== null) {
                        $o['indF'] = -(float) $h / 20;
                    } elseif (($f = self::wa($c, 'firstLine')) !== null) {
                        $o['indF'] = (float) $f / 20;
                    }
                    break;
                case 'spacing':
                    if (($v = self::wa($c, 'before')) !== null) {
                        $o['sb'] = (float) $v / 20;
                    }
                    if (($v = self::wa($c, 'after')) !== null) {
                        $o['sa'] = (float) $v / 20;
                    }
                    if (in_array(self::wa($c, 'beforeAutospacing'), ['1', 'true', 'on'], true)) {
                        $o['sb'] = 14.0;
                    }
                    if (in_array(self::wa($c, 'afterAutospacing'), ['1', 'true', 'on'], true)) {
                        $o['sa'] = 14.0;
                    }
                    if (($v = self::wa($c, 'line')) !== null) {
                        $o['line'] = (float) $v;
                        $o['lineRule'] = self::wa($c, 'lineRule') ?? 'auto';
                    }
                    break;
                case 'tabs':
                    $tabs = [];
                    foreach (self::kids($c, 'w:tab') as $t) {
                        $type = (string) self::wa($t, 'val');
                        $type = ['start' => 'left', 'end' => 'right'][$type] ?? $type;
                        if ($type === 'bar' || $type === 'num') {
                            continue;
                        }
                        $tabs[] = [(float) self::wa($t, 'pos') / 20, $type, self::wa($t, 'leader')];
                    }
                    $o['tabs'] = $tabs;
                    break;
                case 'keepNext':
                case 'keepLines':
                    $o[$c->localName] = self::on($c);
                    break;
                case 'pageBreakBefore':
                    $o['pbb'] = self::on($c);
                    break;
                case 'contextualSpacing':
                    $o['ctx'] = self::on($c);
                    break;
                case 'widowControl':
                    $o['widow'] = self::on($c);
                    break;
                case 'numPr':
                    $o['numId'] = (string) (self::val($c, 'w:numId') ?? '0');
                    $il = self::val($c, 'w:ilvl');
                    if ($il !== null) {
                        $o['ilvl'] = (int) $il;
                    }
                    break;
                case 'pBdr':
                    $o['bdr'] = $this->bdrs($c);
                    break;
                case 'shd':
                    $v = (string) self::wa($c, 'fill');
                    $o['shd'] = preg_match('/^[0-9A-Fa-f]{6}$/', $v) ? strtoupper($v) : null;
                    break;
                case 'rPr':
                    $o['markR'] = $this->rpr($c);
                    break;
                case 'sectPr':
                    $o['sect'] = $c;
                    break;
            }
        }
        return $o;
    }

    /** Bordi: lato => ['v' stile, 'w' spessore pt, 'sp' distanza pt, 'c' colore] */
    private function bdrs(?DOMElement $e): array
    {
        $o = [];
        foreach (self::kids($e) as $c) {
            $side = ['start' => 'left', 'end' => 'right'][$c->localName] ?? $c->localName;
            $v = (string) self::wa($c, 'val');
            $col = (string) self::wa($c, 'color');
            $o[$side] = [
                'v' => ($v === 'nil' || $v === '') ? 'none' : $v,
                'w' => max(0.25, (float) (self::wa($c, 'sz') ?? 4) / 8),
                'sp' => (float) (self::wa($c, 'space') ?? 0),
                'c' => preg_match('/^[0-9A-Fa-f]{6}$/', $col) ? strtoupper($col) : '000000',
            ];
        }
        return $o;
    }

    private function cellMar(?DOMElement $e): array
    {
        $o = [];
        foreach (self::kids($e) as $c) {
            $k = ['top' => 't', 'bottom' => 'b', 'left' => 'l', 'start' => 'l', 'right' => 'r', 'end' => 'r'][$c->localName] ?? null;
            if ($k !== null && self::wa($c, 'w') !== null) {
                $o[$k] = (float) self::wa($c, 'w') / 20;
            }
        }
        return $o;
    }

    private static function bdrOn(?array $b): bool
    {
        return $b !== null && $b['v'] !== 'none';
    }

    private static function bdrW(?array $b): float
    {
        if (!self::bdrOn($b)) {
            return 0.0;
        }
        return $b['v'] === 'double' ? $b['w'] * 3 : $b['w'];
    }

    /** Unisce più livelli di proprietà di paragrafo (il primo ha la precedenza). */
    private static function mergeP(array ...$layers): array
    {
        $out = [];
        $tabs = [];
        foreach (array_reverse($layers) as $l) {
            foreach ($l['tabs'] ?? [] as $t) {
                $k = sprintf('%.2f', $t[0]);
                if ($t[1] === 'clear') {
                    unset($tabs[$k]);
                } else {
                    $tabs[$k] = $t;
                }
            }
            $out = $l + $out;
        }
        usort($tabs, static fn ($a, $b) => $a[0] <=> $b[0]);
        $out['tabs'] = $tabs;
        return $out;
    }

    private function fmt(array $r): DocxPdfFmt
    {
        $hl = $r['hl'] ?? ($r['shd'] ?? null);
        $key = implode('|', [
            $r['font'] ?? '', $r['size'] ?? 10, (int) !empty($r['b']), (int) !empty($r['i']),
            (int) !empty($r['u']), (int) !empty($r['strike']), $r['color'] ?? '', $hl ?? '',
            $r['va'] ?? '', $r['sp'] ?? 0, (int) !empty($r['caps']), (int) !empty($r['smallCaps']),
            $r['pos'] ?? 0, $r['scale'] ?? 100,
        ]);
        if (isset($this->fmts[$key])) {
            return $this->fmts[$key];
        }
        $f = new DocxPdfFmt();
        $f->src = $r;
        [$f->font, $f->fakeB, $f->fakeI] = $this->book->face((string) ($r['font'] ?? 'Times New Roman'), !empty($r['b']), !empty($r['i']));
        $size = (float) ($r['size'] ?? 10);
        $f->asc = $f->font->asc * $size;
        $f->desc = $f->font->desc * $size;
        $f->lead = $f->font->lead * $size;
        $f->size = $size;
        $f->rise = (float) ($r['pos'] ?? 0);
        $va = $r['va'] ?? '';
        if ($va === 'superscript') {
            $f->size = $size * 0.58;
            $f->rise += $size * 0.33;
        } elseif ($va === 'subscript') {
            $f->size = $size * 0.58;
            $f->rise -= $size * 0.14;
        }
        $f->u = !empty($r['u']);
        $f->strike = !empty($r['strike']);
        $f->caps = !empty($r['caps']);
        $f->small = !empty($r['smallCaps']) && !$f->caps;
        $f->color = $r['color'] ?? '000000';
        $f->hl = $hl;
        $f->sp = (float) ($r['sp'] ?? 0);
        $f->scale = (float) ($r['scale'] ?? 100) / 100;
        return $this->fmts[$key] = $f;
    }

    private function variant(DocxPdfFmt $f, array $over): DocxPdfFmt
    {
        return $this->fmt($over + $f->src);
    }

    // ------------------------------------------------------------------
    // Lettura dei blocchi (paragrafi e tabelle)
    // ------------------------------------------------------------------

    private function blocks(DOMElement $container, array $ctx = []): array
    {
        $out = [];
        foreach ($container->childNodes as $c) {
            if ($c->nodeType !== XML_ELEMENT_NODE) {
                continue;
            }
            switch (self::nm($c)) {
                case 'w:p':
                    $out[] = $this->para($c, $ctx);
                    break;
                case 'w:tbl':
                    $out[] = $this->table($c, $ctx);
                    break;
                case 'w:sdt':
                    $k = self::kid($c, 'w:sdtContent');
                    if ($k !== null) {
                        array_push($out, ...$this->blocks($k, $ctx));
                    }
                    break;
                case 'w:customXml':
                case 'w:ins':
                    array_push($out, ...$this->blocks($c, $ctx));
                    break;
            }
        }
        return $out;
    }

    private function para(DOMElement $p, array $ctx): array
    {
        $direct = $this->ppr(self::kid($p, 'w:pPr'));
        $styleId = $direct['style'] ?? $this->defPStyle;
        $st = $this->style($styleId);
        $tb = isset($ctx['tblStyle']) ? $this->style($ctx['tblStyle']) : ['p' => [], 'r' => []];
        if (isset($ctx['cond'])) {
            $tb = ['p' => $ctx['cond']['p'] + $tb['p'], 'r' => $ctx['cond']['r'] + $tb['r']];
        }
        $rBase = $st['r'] + $tb['r'] + $this->defR;

        // Numerazione
        $numId = $direct['numId'] ?? ($st['p']['numId'] ?? '0');
        $ilvl = $direct['ilvl'] ?? ($st['p']['ilvl'] ?? 0);
        $lvl = null;
        if ($numId !== '0' && isset($this->nums[$numId])) {
            $num = $this->nums[$numId];
            $abs = $this->absNums[$num['abs']] ?? [];
            $lvl = ($num['ov'][$ilvl] ?? []) + ($abs[$ilvl] ?? []);
            if (!isset($lvl['fmt'])) {
                $lvl = null;
            }
        }
        $np = $lvl['p'] ?? [];
        $pp = isset($direct['numId'])
            ? self::mergeP($direct, $np, $st['p'], $tb['p'], $this->defP)
            : self::mergeP($direct, $st['p'], $np, $tb['p'], $this->defP);
        $pp += [
            'jc' => 'left', 'indL' => 0.0, 'indR' => 0.0, 'indF' => 0.0, 'sb' => 0.0, 'sa' => 0.0,
            'line' => 240.0, 'lineRule' => 'auto', 'keepNext' => false, 'keepLines' => false,
            'pbb' => false, 'ctx' => false, 'widow' => true, 'bdr' => null, 'shd' => null,
        ];
        $markR = $direct['markR'] ?? [];
        if (isset($markR['rStyle'])) {
            $markR += $this->style($markR['rStyle'])['r'];
        }
        $mark = $this->fmt($markR + $rBase);

        $A = [];
        $st2 = ['glue' => false, 'fld' => null, 'depth' => 0];
        if ($lvl !== null) {
            $label = $this->numLabel($numId, $ilvl, $lvl);
            if ($label !== '') {
                $lf = $this->fmt($lvl['r'] + array_diff_key($markR, ['u' => 1]) + $rBase);
                $this->word($label, $lf, $A, false);
                if ($lvl['suff'] === 'space') {
                    $A[] = ['t' => 'sp', 's' => ' ', 'f' => $lf, 'w' => $lf->width(' ')];
                } elseif ($lvl['suff'] !== 'nothing') {
                    $A[] = ['t' => 'tab', 'f' => $lf];
                }
            }
        }
        $this->inl($p, $rBase, $A, $st2);

        return ['k' => 'p', 'pp' => $pp, 'mark' => $mark, 'A' => $A, 'sect' => $direct['sect'] ?? null, 'style' => $styleId];
    }

    private function numLabel(string $numId, int $ilvl, array $lvl): string
    {
        $num = $this->nums[$numId];
        $key = $num['own'] ? 'n' . $numId : 'a' . $num['abs'];
        $abs = $this->absNums[$num['abs']] ?? [];
        $cnt = $this->counters[$key] ?? [];
        for ($l = 0; $l < $ilvl; $l++) {
            if (!isset($cnt[$l])) {
                $cnt[$l] = (($num['ov'][$l] ?? []) + ($abs[$l] ?? []))['start'] ?? 1;
            }
        }
        $cnt[$ilvl] = isset($cnt[$ilvl]) ? $cnt[$ilvl] + 1 : $lvl['start'];
        foreach (array_keys($cnt) as $l) {
            if ($l > $ilvl) {
                unset($cnt[$l]);
            }
        }
        $this->counters[$key] = $cnt;
        if ($lvl['fmt'] === 'none') {
            return '';
        }
        if ($lvl['fmt'] === 'bullet') {
            return $lvl['text'];
        }
        return (string) preg_replace_callback('/%(\d)/', function ($m) use ($cnt, $num, $abs) {
            $l = (int) $m[1] - 1;
            $f = ((($num['ov'][$l] ?? []) + ($abs[$l] ?? []))['fmt']) ?? 'decimal';
            return self::numFormat($cnt[$l] ?? 1, $f === 'bullet' ? 'decimal' : $f);
        }, $lvl['text']);
    }

    private function runProps(?DOMElement $r, array $base): array
    {
        $d = $this->rpr(self::kid($r, 'w:rPr'));
        if (isset($d['rStyle'])) {
            $d += $this->style($d['rStyle'])['r'];
        }
        return $d + $base;
    }

    private function inl(DOMElement $parent, array $base, array &$A, array &$st): void
    {
        foreach ($parent->childNodes as $c) {
            if ($c->nodeType !== XML_ELEMENT_NODE) {
                continue;
            }
            switch (self::nm($c)) {
                case 'w:r':
                    $props = $this->runProps($c, $base);
                    if (empty($props['vanish'])) {
                        $f = null;
                        foreach ($c->childNodes as $k) {
                            if ($k->nodeType === XML_ELEMENT_NODE) {
                                $this->runChild($k, $props, $f, $A, $st);
                            }
                        }
                    }
                    break;
                case 'w:hyperlink':
                case 'w:smartTag':
                case 'w:ins':
                case 'w:customXml':
                case 'w:sdtContent':
                case 'w:dir':
                case 'w:bdo':
                    $this->inl($c, $base, $A, $st);
                    break;
                case 'w:sdt':
                    $k = self::kid($c, 'w:sdtContent');
                    if ($k !== null) {
                        $this->inl($k, $base, $A, $st);
                    }
                    break;
                case 'w:fldSimple':
                    $code = strtoupper((string) strtok(trim((string) self::wa($c, 'instr')), " \\"));
                    if (in_array($code, ['PAGE', 'NUMPAGES', 'SECTIONPAGES'], true)) {
                        $r = self::kid($c, 'w:r');
                        $f = $this->fmt($this->runProps($r, $base));
                        $s = $r !== null ? trim($r->textContent) : '1';
                        $A[] = ['t' => 'fld', 'code' => $code, 's' => $s, 'f' => $f, 'w' => $f->width($s), 'bk' => !$st['glue']];
                        $st['glue'] = true;
                    } else {
                        $this->inl($c, $base, $A, $st);
                    }
                    break;
            }
        }
    }

    private function runChild(DOMElement $c, array $props, ?DocxPdfFmt &$f, array &$A, array &$st): void
    {
        $name = self::nm($c);
        if ($name === 'w:fldChar') {
            $type = self::wa($c, 'fldCharType');
            if ($type === 'begin') {
                if ($st['fld'] === null) {
                    $st['fld'] = ['instr' => '', 'phase' => 'instr', 'known' => false, 'code' => ''];
                }
                $st['depth']++;
            } elseif ($type === 'separate' && $st['depth'] === 1 && $st['fld'] !== null) {
                $st['fld']['phase'] = 'result';
                $this->fieldStart($props, $A, $st);
            } elseif ($type === 'end' && $st['depth'] > 0) {
                $st['depth']--;
                if ($st['depth'] === 0 && $st['fld'] !== null) {
                    if ($st['fld']['phase'] === 'instr') {
                        $this->fieldStart($props, $A, $st);
                    }
                    $st['fld'] = null;
                }
            }
            return;
        }
        if ($name === 'w:instrText') {
            if ($st['fld'] !== null && $st['depth'] === 1 && $st['fld']['phase'] === 'instr') {
                $st['fld']['instr'] .= $c->textContent;
            }
            return;
        }
        if ($st['fld'] !== null && ($st['fld']['phase'] === 'instr' || $st['fld']['known'])) {
            return;
        }
        $f ??= $this->fmt($props);
        switch ($name) {
            case 'w:t':
                $this->text($c->textContent, $f, $A, $st);
                break;
            case 'w:tab':
            case 'w:ptab':
                $A[] = ['t' => 'tab', 'f' => $f];
                $st['glue'] = false;
                break;
            case 'w:br':
            case 'w:cr':
                $type = self::wa($c, 'type') ?? 'line';
                $A[] = ['t' => 'br', 'kind' => ($type === 'page' || $type === 'column') ? 'page' : 'line', 'f' => $f];
                $st['glue'] = false;
                break;
            case 'w:noBreakHyphen':
                $this->word('-', $f, $A, !$st['glue']);
                $st['glue'] = true;
                break;
            case 'w:sym':
                $cp = (int) hexdec((string) self::wa($c, 'char'));
                $font = self::wa($c, 'font');
                $sf = $font !== null ? $this->variant($f, ['font' => $font]) : $f;
                if ($cp > 0) {
                    $this->word(mb_chr($cp, 'UTF-8'), $sf, $A, !$st['glue']);
                    $st['glue'] = true;
                }
                break;
            case 'w:drawing':
                $this->drawing($c, $f, $A, $st);
                break;
            case 'mc:AlternateContent':
                $choice = self::kid($c, 'mc:Choice');
                $use = ($choice !== null && in_array($choice->getAttribute('Requires'), ['wps', 'wpg'], true))
                    ? $choice
                    : (self::kid($c, 'mc:Fallback') ?? $choice);
                foreach (self::kids($use) as $k) {
                    $this->runChild($k, $props, $f, $A, $st);
                }
                break;
            case 'w:pict':
                $this->vml($c, $f, $A, $st);
                break;
            case 'w:footnoteReference':
            case 'w:endnoteReference':
                $this->note('Note a piè di pagina e di chiusura non supportate: riferimento ignorato');
                break;
        }
    }

    private function fieldStart(array $props, array &$A, array &$st): void
    {
        $code = strtoupper((string) strtok(trim($st['fld']['instr']), " \\"));
        $st['fld']['code'] = $code;
        if (in_array($code, ['PAGE', 'NUMPAGES', 'SECTIONPAGES'], true)) {
            $st['fld']['known'] = true;
            $f = $this->fmt($props);
            $A[] = ['t' => 'fld', 'code' => $code, 's' => '1', 'f' => $f, 'w' => $f->width('1'), 'bk' => !$st['glue']];
            $st['glue'] = true;
        }
    }

    private function text(string $s, DocxPdfFmt $f, array &$A, array &$st): void
    {
        if ($s === '') {
            return;
        }
        $s = str_replace(["\u{00AD}", "\u{200B}", "\u{FEFF}", "\r"], '', $s);
        $s = strtr($s, "\n", ' ');
        if ($f->caps) {
            $s = mb_strtoupper($s, 'UTF-8');
        }
        $hy = '\-\x{2010}\x{2013}\x{2014}';
        foreach (preg_split('/( +|\t)/u', $s, -1, PREG_SPLIT_DELIM_CAPTURE | PREG_SPLIT_NO_EMPTY) ?: [] as $piece) {
            if ($piece[0] === ' ') {
                $A[] = ['t' => 'sp', 's' => $piece, 'f' => $f, 'w' => $f->width($piece)];
                $st['glue'] = false;
                continue;
            }
            if ($piece === "\t") {
                $A[] = ['t' => 'tab', 'f' => $f];
                $st['glue'] = false;
                continue;
            }
            foreach (preg_split('/(?<=[' . $hy . '])(?=[^' . $hy . '])/u', $piece) ?: [] as $w) {
                if ($w === '') {
                    continue;
                }
                $this->word($w, $f, $A, !$st['glue']);
                $st['glue'] = !preg_match('/[' . $hy . ']$/u', $w);
            }
        }
    }

    /** Aggiunge una parola, cambiando font per i caratteri che quello richiesto non contiene. */
    private function word(string $w, DocxPdfFmt $f, array &$A, bool $bk): void
    {
        $cps = DocxPdfFont::codepoints($w);
        $plain = !$f->small;
        if ($plain) {
            foreach ($cps as $cp) {
                if (!$f->font->has($cp)) {
                    $plain = false;
                    break;
                }
            }
        }
        if ($plain) {
            $A[] = ['t' => 'x', 's' => $w, 'f' => $f, 'w' => $f->width($w), 'bk' => $bk];
            return;
        }
        $cur = '';
        $curF = $f;
        $flush = function () use (&$cur, &$curF, &$A, &$bk): void {
            if ($cur !== '') {
                $A[] = ['t' => 'x', 's' => $cur, 'f' => $curF, 'w' => $curF->width($cur), 'bk' => $bk];
                $bk = false;
                $cur = '';
            }
        };
        foreach ($cps as $cp) {
            $ff = $f;
            if ($f->small) {
                $ch = mb_chr($cp, 'UTF-8');
                $up = mb_strtoupper($ch, 'UTF-8');
                if ($up !== $ch && mb_strlen($up, 'UTF-8') === 1) {
                    $cp = mb_ord($up, 'UTF-8');
                    $ff = $this->variant($f, ['size' => ($f->src['size'] ?? 10) * 0.8, 'smallCaps' => false]);
                }
            }
            if (!$ff->font->has($cp)) {
                if (isset(self::SYM[$cp])) {
                    $cp = self::SYM[$cp];
                }
                if (!$ff->font->has($cp)) {
                    $fb = $this->book->fallback($cp, !empty($f->src['b']), !empty($f->src['i']));
                    if ($fb !== null) {
                        $ff = $this->variant($ff, ['font' => $fb->family]);
                    }
                }
            }
            if ($ff !== $curF) {
                $flush();
                $curF = $ff;
            }
            $cur .= mb_chr($cp, 'UTF-8');
        }
        $flush();
    }

    // ------------------------------------------------------------------
    // Immagini e forme
    // ------------------------------------------------------------------

    private static function emu(?string $v): float
    {
        return (float) $v / 12700;
    }

    private function drawing(DOMElement $dr, DocxPdfFmt $f, array &$A, array &$st): void
    {
        $box = self::kid($dr, 'wp:inline') ?? self::kid($dr, 'wp:anchor');
        if ($box === null) {
            return;
        }
        $inline = $box->localName === 'inline';
        $ext = self::kid($box, 'wp:extent');
        $d = [
            'w' => self::emu($ext?->getAttribute('cx')),
            'h' => self::emu($ext?->getAttribute('cy')),
            'behind' => $box->getAttribute('behindDoc') === '1',
            'kind' => null,
        ];
        if (!$inline) {
            foreach (['H' => 'wp:positionH', 'V' => 'wp:positionV'] as $ax => $tag) {
                $p = self::kid($box, $tag);
                $d['rel' . $ax] = $p?->getAttribute('relativeFrom') ?: ($ax === 'H' ? 'column' : 'paragraph');
                $al = self::kid($p, 'wp:align');
                $off = self::kid($p, 'wp:posOffset');
                $d['align' . $ax] = $al !== null ? trim($al->textContent) : null;
                $d['off' . $ax] = $off !== null ? self::emu(trim($off->textContent)) : 0.0;
            }
            if ($box->getAttribute('simplePos') === '1') {
                $sp = self::kid($box, 'wp:simplePos');
                $d['relH'] = $d['relV'] = 'page';
                $d['alignH'] = $d['alignV'] = null;
                $d['offH'] = self::emu($sp?->getAttribute('x'));
                $d['offV'] = self::emu($sp?->getAttribute('y'));
            }
            foreach (self::kids($box) as $k) {
                if (in_array($k->localName, ['wrapSquare', 'wrapTight', 'wrapThrough', 'wrapTopAndBottom'], true)) {
                    $this->note('Testo che scorre attorno a un oggetto flottante non supportato: l\'oggetto viene sovrapposto');
                }
            }
        }
        $gd = self::kid(self::kid($box, 'a:graphic'), 'a:graphicData');
        foreach (self::kids($gd) as $g) {
            $n = self::nm($g);
            if ($n === 'pic:pic') {
                $bf = self::kid($g, 'pic:blipFill');
                $d['kind'] = 'pic';
                $d['img'] = $this->blip($bf);
                $d['crop'] = $this->crop($bf);
                $d['line'] = $this->lineOf(self::kid(self::kid($g, 'pic:spPr'), 'a:ln'), null);
            } elseif ($n === 'wps:wsp') {
                $this->shape($g, $d);
            } else {
                $this->note('Oggetto grafico non supportato (' . $g->localName . '): ignorato');
            }
        }
        if ($d['kind'] === null) {
            return;
        }
        if ($inline) {
            $A[] = ['t' => 'img', 'd' => $d, 'w' => $d['w'], 'h' => $d['h'], 'f' => $f, 'bk' => true];
            $st['glue'] = false;
        } else {
            $A[] = ['t' => 'anc', 'd' => $d, 'f' => $f, 'w' => 0.0];
        }
    }

    private function blip(?DOMElement $blipFill): ?string
    {
        $blip = self::kid($blipFill, 'a:blip');
        if ($blip === null) {
            return null;
        }
        $id = $blip->getAttributeNS(self::NS_R, 'embed');
        return $this->relsOf($this->part)[$id] ?? null;
    }

    private function crop(?DOMElement $blipFill): array
    {
        $s = self::kid($blipFill, 'a:srcRect');
        $c = [];
        foreach (['l', 't', 'r', 'b'] as $k) {
            $c[$k] = $s !== null ? max(0.0, (float) $s->getAttribute($k) / 100000) : 0.0;
        }
        return $c;
    }

    /** Colore DrawingML (srgbClr, schemeClr, sysClr…) con le modifiche di luminosità. */
    private function clr(?DOMElement $parent): ?string
    {
        foreach (self::kids($parent) as $c) {
            $hex = null;
            switch ($c->localName) {
                case 'srgbClr':
                    $hex = $c->getAttribute('val');
                    break;
                case 'sysClr':
                    $hex = $c->getAttribute('lastClr') ?: '000000';
                    break;
                case 'scrgbClr':
                    $hex = sprintf(
                        '%02X%02X%02X',
                        (int) round(min(100000, (float) $c->getAttribute('r')) * 255 / 100000),
                        (int) round(min(100000, (float) $c->getAttribute('g')) * 255 / 100000),
                        (int) round(min(100000, (float) $c->getAttribute('b')) * 255 / 100000)
                    );
                    break;
                case 'schemeClr':
                    $v = $c->getAttribute('val');
                    $v = ['tx1' => 'dk1', 'tx2' => 'dk2', 'bg1' => 'lt1', 'bg2' => 'lt2'][$v] ?? $v;
                    $hex = $this->themeColors[$v] ?? '000000';
                    break;
                case 'prstClr':
                    $hex = ['white' => 'FFFFFF', 'red' => 'FF0000', 'blue' => '0000FF', 'green' => '008000'][$c->getAttribute('val')] ?? '000000';
                    break;
            }
            if ($hex === null || strlen($hex) !== 6) {
                continue;
            }
            $rgb = array_map('hexdec', str_split($hex, 2));
            foreach (self::kids($c) as $m) {
                $v = (float) $m->getAttribute('val') / 100000;
                foreach ($rgb as $i => $ch) {
                    $rgb[$i] = match ($m->localName) {
                        'lumMod', 'shade' => $ch * $v,
                        'lumOff' => $ch + 255 * $v,
                        'tint' => $ch * $v + 255 * (1 - $v),
                        default => $ch,
                    };
                }
            }
            return sprintf('%02X%02X%02X', ...array_map(static fn ($x) => (int) max(0, min(255, round($x))), $rgb));
        }
        return null;
    }

    private function lineOf(?DOMElement $ln, ?DOMElement $styleRef): ?array
    {
        $w = null;
        $color = null;
        $dash = null;
        if ($ln !== null) {
            if (self::kid($ln, 'a:noFill') !== null) {
                return null;
            }
            if ($ln->getAttribute('w') !== '') {
                $w = self::emu($ln->getAttribute('w'));
            }
            $color = $this->clr(self::kid($ln, 'a:solidFill'));
            $pd = self::kid($ln, 'a:prstDash');
            if ($pd !== null && $pd->getAttribute('val') !== 'solid') {
                $dash = str_contains($pd->getAttribute('val'), 'ot') ? 'dot' : 'dash';
            }
        }
        if ($color === null && $styleRef !== null && (int) $styleRef->getAttribute('idx') > 0) {
            $color = $this->clr($styleRef);
            $w ??= [1 => 0.5, 2 => 1.0, 3 => 1.5][(int) $styleRef->getAttribute('idx')] ?? 0.75;
        }
        if ($color === null) {
            return null;
        }
        return ['w' => max(0.1, $w ?? 0.75), 'c' => $color, 'dash' => $dash];
    }

    private function shape(DOMElement $wsp, array &$d): void
    {
        $sp = self::kid($wsp, 'wps:spPr');
        $style = self::kid($wsp, 'wps:style');
        $xf = self::kid($sp, 'a:xfrm');
        $d['kind'] = 'shape';
        $d['geom'] = self::kid($sp, 'a:prstGeom')?->getAttribute('prst') ?: 'rect';
        $d['flipV'] = $xf !== null && $xf->getAttribute('flipV') === '1';
        $d['flipH'] = $xf !== null && $xf->getAttribute('flipH') === '1';
        $d['fill'] = null;
        $d['img'] = null;
        if (self::kid($sp, 'a:noFill') === null) {
            $bf = self::kid($sp, 'a:blipFill');
            if ($bf !== null) {
                $d['img'] = $this->blip($bf);
                $d['crop'] = $this->crop($bf);
            } elseif (($sf = self::kid($sp, 'a:solidFill')) !== null) {
                $d['fill'] = $this->clr($sf);
            } elseif (($gf = self::kid($sp, 'a:gradFill')) !== null) {
                $d['fill'] = $this->clr(self::kid(self::kid($gf, 'a:gsLst'), 'a:gs'));
            } elseif (($fr = self::kid($style, 'a:fillRef')) !== null && (int) $fr->getAttribute('idx') > 0) {
                $d['fill'] = $this->clr($fr);
            }
        }
        $d['line'] = $this->lineOf(self::kid($sp, 'a:ln'), self::kid($style, 'a:lnRef'));
        if (in_array($d['geom'], ['line', 'straightConnector1'], true)) {
            $d['fill'] = null;
            $d['img'] = null;
        }
        $tx = self::kid(self::kid($wsp, 'wps:txbx'), 'w:txbxContent');
        if ($tx !== null) {
            $bp = self::kid($wsp, 'wps:bodyPr');
            $ins = [];
            foreach (['lIns' => 7.2, 'tIns' => 3.6, 'rIns' => 7.2, 'bIns' => 3.6] as $k => $def) {
                $ins[] = ($bp !== null && $bp->getAttribute($k) !== '') ? self::emu($bp->getAttribute($k)) : $def;
            }
            $d['ins'] = $ins;
            $d['anchor'] = $bp?->getAttribute('anchor') ?: 't';
            $d['txbx'] = $this->blocks($tx);
        }
    }

    /** Immagini nel vecchio formato VML (w:pict): solo quelle in linea. */
    private function vml(DOMElement $pict, DocxPdfFmt $f, array &$A, array &$st): void
    {
        foreach ($pict->getElementsByTagNameNS('urn:schemas-microsoft-com:vml', 'imagedata') as $im) {
            $shape = $im->parentNode;
            $style = $shape instanceof DOMElement ? $shape->getAttribute('style') : '';
            $path = $this->relsOf($this->part)[$im->getAttributeNS(self::NS_R, 'id')] ?? null;
            $dim = [];
            foreach (['width', 'height'] as $k) {
                if (preg_match('/(?:^|;)\s*' . $k . ':\s*([\d.]+)(pt|in|cm|mm|px)?/', $style, $m)) {
                    $dim[$k] = (float) $m[1] * (['in' => 72, 'cm' => 28.3465, 'mm' => 2.83465, 'px' => 0.75][$m[2] ?? 'pt'] ?? 1);
                }
            }
            if ($path === null || count($dim) < 2 || str_contains($style, 'position:absolute')) {
                $this->note('Oggetto VML non supportato: ignorato');
                continue;
            }
            $d = ['w' => $dim['width'], 'h' => $dim['height'], 'kind' => 'pic', 'img' => $path,
                'crop' => ['l' => 0, 't' => 0, 'r' => 0, 'b' => 0], 'line' => null, 'behind' => false];
            $A[] = ['t' => 'img', 'd' => $d, 'w' => $d['w'], 'h' => $d['h'], 'f' => $f, 'bk' => true];
            $st['glue'] = false;
        }
    }

    // ------------------------------------------------------------------
    // Tabelle
    // ------------------------------------------------------------------

    private function table(DOMElement $t, array $ctx): array
    {
        $pr = self::kid($t, 'w:tblPr');
        $styleId = self::val($pr, 'w:tblStyle');
        $ts = $this->style($styleId)['tbl'];
        $tbdr = $this->bdrs(self::kid($pr, 'w:tblBorders')) + $ts['bdr'];
        $mar = $this->cellMar(self::kid($pr, 'w:tblCellMar')) + $ts['mar'] + ['l' => 5.4, 'r' => 5.4, 't' => 0.0, 'b' => 0.0];
        $grid = [];
        foreach (self::kids(self::kid($t, 'w:tblGrid'), 'w:gridCol') as $g) {
            $grid[] = max(1.0, (float) self::wa($g, 'w') / 20);
        }
        $ind = self::kid($pr, 'w:tblInd');
        $sub = ['tblStyle' => $styleId];
        $cond = $ts['cond'];
        $look = self::kid($pr, 'w:tblLook');
        $mask = $look !== null ? (int) hexdec((string) (self::wa($look, 'val') ?? '04A0')) : 0x04A0;
        $flag = static function (string $attr, int $bit) use ($look, $mask): bool {
            $v = self::wa($look, $attr);
            return $v !== null ? ($v === '1' || $v === 'true') : (bool) ($mask & $bit);
        };
        $lk = $cond ? [
            'firstRow' => $flag('firstRow', 0x20), 'lastRow' => $flag('lastRow', 0x40),
            'firstCol' => $flag('firstColumn', 0x80), 'lastCol' => $flag('lastColumn', 0x100),
            'hBand' => !$flag('noHBand', 0x200), 'vBand' => !$flag('noVBand', 0x400),
        ] : null;
        $rows = [];
        $rowEls = [];
        foreach (self::kids($t) as $c) {
            if (self::nm($c) === 'w:tr') {
                $rowEls[] = $c;
            } elseif (self::nm($c) === 'w:sdt') {
                array_push($rowEls, ...self::kids(self::kid($c, 'w:sdtContent'), 'w:tr'));
            }
        }
        $maxCols = 0;
        $nRows = count($rowEls);
        $nGrid = count($grid);
        foreach ($rowEls as $ri => $tr) {
            $trPr = self::kid($tr, 'w:trPr');
            $hEl = self::kid($trPr, 'w:trHeight');
            $hdr = self::kid($trPr, 'w:tblHeader');
            $cells = [];
            $cols = (int) (self::val($trPr, 'w:gridBefore') ?? 0);
            $before = $cols;
            $cellEls = [];
            foreach (self::kids($tr) as $c) {
                if (self::nm($c) === 'w:tc') {
                    $cellEls[] = $c;
                } elseif (self::nm($c) === 'w:sdt') {
                    array_push($cellEls, ...self::kids(self::kid($c, 'w:sdtContent'), 'w:tc'));
                }
            }
            foreach ($cellEls as $tc) {
                $tcPr = self::kid($tc, 'w:tcPr');
                $vm = self::kid($tcPr, 'w:vMerge');
                $shd = (string) self::val($tcPr, 'w:shd', 'fill');
                $span = max(1, (int) (self::val($tcPr, 'w:gridSpan') ?? 1));
                $cc = ['p' => [], 'r' => [], 'bdr' => []];
                if ($lk !== null) {
                    $types = [];
                    if ($lk['vBand']) {
                        $ci = $cols - ($lk['firstCol'] ? 1 : 0);
                        if ($ci >= 0) {
                            $types[] = $ci % 2 === 0 ? 'band1Vert' : 'band2Vert';
                        }
                    }
                    if ($lk['hBand']) {
                        $bi = $ri - ($lk['firstRow'] ? 1 : 0);
                        if ($bi >= 0) {
                            $types[] = $bi % 2 === 0 ? 'band1Horz' : 'band2Horz';
                        }
                    }
                    if ($lk['firstCol'] && $cols === 0) {
                        $types[] = 'firstCol';
                    }
                    if ($lk['lastCol'] && $nGrid > 1 && $cols + $span >= $nGrid) {
                        $types[] = 'lastCol';
                    }
                    if ($lk['lastRow'] && $nRows > 1 && $ri === $nRows - 1) {
                        $types[] = 'lastRow';
                    }
                    if ($lk['firstRow'] && $ri === 0) {
                        $types[] = 'firstRow';
                    }
                    foreach ($types as $type) {
                        if (isset($cond[$type])) {
                            $c2 = $cond[$type];
                            $cc = ['p' => $c2['p'] + $cc['p'], 'r' => $c2['r'] + $cc['r'], 'bdr' => $c2['bdr'] + $cc['bdr']]
                                + (isset($c2['shd']) ? ['shd' => $c2['shd']] : array_intersect_key($cc, ['shd' => 1]));
                        }
                    }
                }
                $cells[] = [
                    'span' => $span,
                    'vm' => $vm === null ? null : (self::wa($vm, 'val') === 'restart' ? 'r' : 'c'),
                    'bdr' => $this->bdrs(self::kid($tcPr, 'w:tcBorders')) + array_intersect_key($cc['bdr'], ['top' => 1, 'bottom' => 1, 'left' => 1, 'right' => 1]),
                    'shd' => preg_match('/^[0-9A-Fa-f]{6}$/', $shd) ? strtoupper($shd) : ($cc['shd'] ?? null),
                    'mar' => $this->cellMar(self::kid($tcPr, 'w:tcMar')) + $mar,
                    'va' => self::val($tcPr, 'w:vAlign') ?? 'top',
                    'tcW' => (self::val($tcPr, 'w:tcW', 'type') ?? 'dxa') === 'dxa' ? (float) self::val($tcPr, 'w:tcW', 'w') / 20 : 0.0,
                    'blocks' => $this->blocks($tc, $sub + ['cond' => $cc]),
                ];
                $cols += $span;
            }
            $maxCols = max($maxCols, $cols);
            $rows[] = [
                'h' => $hEl !== null ? (float) self::wa($hEl, 'val') / 20 : null,
                'hRule' => self::wa($hEl, 'hRule') ?? 'atLeast',
                'hdr' => $hdr !== null && self::on($hdr),
                'before' => $before,
                'cells' => $cells,
            ];
        }
        if (count($grid) < $maxCols) {
            // Griglia assente o incompleta: larghezze dalle celle della prima riga, o uguali
            $grid = [];
            foreach ($rows[0]['cells'] ?? [] as $c) {
                for ($i = 0; $i < $c['span']; $i++) {
                    $grid[] = $c['tcW'] > 0 ? $c['tcW'] / $c['span'] : 0.0;
                }
            }
            $grid = array_pad($grid, $maxCols, 0.0);
        }
        return [
            'k' => 't', 'grid' => $grid, 'rows' => $rows, 'bdr' => $tbdr, 'mar' => $mar,
            'ind' => ($ind !== null && (self::wa($ind, 'type') ?? 'dxa') === 'dxa') ? (float) self::wa($ind, 'w') / 20 : 0.0,
            'jc' => self::val($pr, 'w:jc') ?? 'left',
        ];
    }

    // ------------------------------------------------------------------
    // Sezioni, intestazioni e piè di pagina
    // ------------------------------------------------------------------

    private function sectProps(?DOMElement $sp, ?array $prev): array
    {
        $s = [
            'pw' => 595.3, 'ph' => 841.9, 'ml' => 72.0, 'mr' => 72.0, 'mt' => 72.0, 'mb' => 72.0,
            'hd' => 36.0, 'fd' => 36.0, 'titlePg' => false, 'type' => 'nextPage',
            'h' => $prev['h'] ?? [], 'f' => $prev['f'] ?? [],
        ];
        if ($sp === null) {
            return $s;
        }
        $sz = self::kid($sp, 'w:pgSz');
        if ($sz !== null) {
            $s['pw'] = (float) (self::wa($sz, 'w') ?? 11906) / 20;
            $s['ph'] = (float) (self::wa($sz, 'h') ?? 16838) / 20;
        }
        $m = self::kid($sp, 'w:pgMar');
        if ($m !== null) {
            foreach (['mt' => 'top', 'mr' => 'right', 'mb' => 'bottom', 'ml' => 'left', 'hd' => 'header', 'fd' => 'footer'] as $k => $a) {
                if (self::wa($m, $a) !== null) {
                    $s[$k] = (float) self::wa($m, $a) / 20;
                }
            }
            $s['ml'] += (float) (self::wa($m, 'gutter') ?? 0) / 20;
        }
        $rels = $this->relsOf('word/document.xml');
        foreach (self::kids($sp) as $c) {
            $k = ['headerReference' => 'h', 'footerReference' => 'f'][$c->localName] ?? null;
            if ($k !== null) {
                $id = $c->getAttributeNS(self::NS_R, 'id');
                if (isset($rels[$id])) {
                    $s[$k][self::wa($c, 'type') ?? 'default'] = $rels[$id];
                }
            }
        }
        $tp = self::kid($sp, 'w:titlePg');
        $s['titlePg'] = $tp !== null && self::on($tp);
        $s['type'] = self::val($sp, 'w:type') ?? 'nextPage';
        if ((int) (self::val($sp, 'w:cols', 'num') ?? 1) > 1) {
            $this->note('Colonne multiple non supportate: il testo viene disposto su una colonna');
        }
        return $s;
    }

    private function hfBlocks(array $s, string $kind, string $variant): array
    {
        $m = $s[$kind];
        $target = match ($variant) {
            'first' => $m['first'] ?? null,
            'even' => $m['even'] ?? ($m['default'] ?? null),
            default => $m['default'] ?? null,
        };
        if ($target === null) {
            return [];
        }
        if (!isset($this->hfCache[$target])) {
            $old = $this->part;
            $this->part = $target;
            $root = $this->xml($target);
            $this->hfCache[$target] = $root !== null ? $this->blocks($root) : [];
            $this->part = $old;
        }
        return $this->hfCache[$target];
    }

    private function hfHeight(array $s, string $kind, string $variant): float
    {
        $key = $kind . $variant . json_encode([$s[$kind], $s['pw'], $s['ml'], $s['mr']]);
        if (!isset($this->hfH[$key])) {
            $blocks = $this->hfBlocks($s, $kind, $variant);
            $save = $this->fv;
            $this->fv = ['PAGE' => 1, 'NUMPAGES' => 1, 'SECTIONPAGES' => 1];
            $this->hfH[$key] = $blocks ? self::stackH($this->flow($blocks, $s['pw'] - $s['ml'] - $s['mr'])) : 0.0;
            $this->fv = $save;
        }
        return $this->hfH[$key];
    }

    // ------------------------------------------------------------------
    // Conversione completa
    // ------------------------------------------------------------------

    public function render(): string
    {
        $root = $this->xml('word/document.xml');
        $body = self::kid($root, 'w:body');
        if ($body === null) {
            throw new DocxPdfException('Il documento non ha un corpo leggibile');
        }
        $cur = [];
        $prev = null;
        foreach ($this->blocks($body) as $b) {
            $cur[] = $b;
            if ($b['k'] === 'p' && $b['sect'] !== null) {
                $prev = $this->sectProps($b['sect'], $prev);
                $this->sections[] = ['pr' => $prev, 'blocks' => $cur];
                $cur = [];
            }
        }
        $this->sections[] = ['pr' => $this->sectProps(self::kid($body, 'w:sectPr'), $prev), 'blocks' => $cur];

        $this->bodyPass();
        if ($this->needTotal) {
            // Il corpo contiene il numero totale di pagine: seconda passata con il valore reale
            $this->total = count($this->pages);
            $this->bodyPass();
        }

        $total = count($this->pages);
        foreach ($this->pages as $i => $pg) {
            $s = $pg['sec'];
            $this->pi = $i;
            $this->ph = $pg['h'];
            $this->layer = 'hf';
            $this->fv = $this->dfv = ['PAGE' => $pg['num'], 'NUMPAGES' => $total, 'SECTIONPAGES' => $total];
            $W = $s['pw'] - $s['ml'] - $s['mr'];
            $hb = $this->hfBlocks($s, 'h', $pg['variant']);
            if ($hb) {
                $this->stack($this->flow($hb, $W), $s['ml'], $s['hd']);
            }
            $fb = $this->hfBlocks($s, 'f', $pg['variant']);
            if ($fb) {
                $items = $this->flow($fb, $W);
                $this->stack($items, $s['ml'], $s['ph'] - $s['fd'] - self::stackH($items));
            }
        }
        foreach ($this->pages as $pg) {
            $this->pdf->addPage($pg['w'], $pg['h'], $pg['back'] . $pg['hf'] . $pg['main'] . $pg['front']);
        }
        foreach ($this->book->subst as $from => $to) {
            $this->note("Font \"$from\" non trovato: usato \"$to\" (gli a capo possono differire da Word)");
        }
        return $this->pdf->output($this->title);
    }

    private function bodyPass(): void
    {
        $this->pages = [];
        $this->pi = -1;
        $this->fv = $this->total !== null ? ['NUMPAGES' => $this->total, 'SECTIONPAGES' => $this->total] : [];
        $last = count($this->sections) - 1;
        $prevS = null;
        foreach ($this->sections as $si => $sec) {
            $s = $sec['pr'];
            $cont = $prevS !== null && $s['type'] === 'continuous'
                && [$s['pw'], $s['ph'], $s['ml'], $s['mr']] === [$prevS['pw'], $prevS['ph'], $prevS['ml'], $prevS['mr']];
            $this->sec = $s;
            $items = $this->flow($sec['blocks'], $s['pw'] - $s['ml'] - $s['mr']);
            if (!$cont) {
                $this->secPages = 0;
                $this->newPage(true);
            }
            $this->paginate($items, $s['ml'], $si < $last);
            $prevS = $s;
        }
    }

    private function newPage(bool $explicit): void
    {
        $s = $this->sec;
        $num = count($this->pages) + 1;
        $variant = ($s['titlePg'] && $this->secPages === 0)
            ? 'first'
            : (($this->evenOdd && $num % 2 === 0) ? 'even' : 'default');
        $hh = $this->hfHeight($s, 'h', $variant);
        $fh = $this->hfHeight($s, 'f', $variant);
        $top = $s['mt'] < 0 ? -$s['mt'] : max($s['mt'], $s['hd'] + $hh);
        $bot = $s['ph'] - ($s['mb'] < 0 ? -$s['mb'] : max($s['mb'], $s['fd'] + $fh));
        $this->pages[] = [
            'w' => $s['pw'], 'h' => $s['ph'], 'back' => '', 'hf' => '', 'main' => '', 'front' => '',
            'sec' => $s, 'variant' => $variant, 'num' => $num,
        ];
        $this->pi = count($this->pages) - 1;
        $this->ph = $s['ph'];
        $this->layer = 'main';
        $this->y = $this->top = $top;
        $this->bot = max($top + 20, $bot);
        $this->pgEmpty = true;
        $this->autoBreak = !$explicit;
        $this->secPages++;
        $this->dfv = ['PAGE' => $num] + $this->fv;
    }

    private function put(array $it, float $X): void
    {
        $this->y += ($this->pgEmpty && $this->autoBreak) ? 0.0 : $it['sb'];
        $this->drawItem($it, $X, $this->y);
        $this->y += $it['h'] + $it['sa'];
        $this->pgEmpty = false;
    }

    /** Distribuisce righe e righe di tabella sulle pagine. */
    private function paginate(array $items, float $X, bool $moreAfter): void
    {
        $n = count($items);
        $i = 0;
        $eps = 0.01;
        while ($i < $n) {
            $it = $items[$i];
            if ($it['k'] === 'r') {
                if ($this->y + $it['h'] > $this->bot + $eps && !$this->pgEmpty) {
                    $this->newPage(false);
                    if (!$it['row']['hdr']) {
                        foreach ($it['hdrRows'] as $hr) {
                            $this->put($hr, $X);
                        }
                    }
                }
                if ($it['h'] > $this->bot - $this->top + $eps) {
                    $this->note('Una riga di tabella è più alta della pagina: il contenuto in eccesso esce dal bordo');
                }
                $this->put($it, $X);
                $i++;
                continue;
            }

            $cnt = $it['n'];
            $pp = $it['pp'];
            if ($pp['pbb'] && !$this->pgEmpty) {
                $this->newPage(true);
            }
            if ($pp['keepNext'] && !$this->pgEmpty) {
                // "Mantieni con il successivo": serve posto anche per l'inizio del blocco seguente
                $need = 0.0;
                $k = $i;
                for ($guard = 0; $k < $n && $guard < 10; $guard++) {
                    $q = $items[$k];
                    if ($q['k'] === 'r') {
                        $need += $q['h'];
                        break;
                    }
                    $m = $q['n'];
                    $keep = $q['pp']['keepNext'];
                    $take = $keep ? $m : ($m <= 3 ? $m : 2);
                    for ($j = 0; $j < $take; $j++) {
                        $need += $items[$k + $j]['sb'] + $items[$k + $j]['h'] + ($keep ? $items[$k + $j]['sa'] : 0);
                    }
                    if (!$keep) {
                        break;
                    }
                    $k += $m;
                }
                if ($this->y + $need > $this->bot + $eps && $need <= $this->bot - $this->top) {
                    $this->newPage(false);
                }
            }

            $k = 0;
            while ($k < $cnt) {
                // Righe fino alla prossima interruzione di pagina forzata (compresa)
                $seg = $cnt - $k;
                for ($m = $k; $m < $cnt; $m++) {
                    if ($items[$i + $m]['ln']['pb']) {
                        $seg = $m - $k + 1;
                        break;
                    }
                }
                $fit = 0;
                $yy = $this->y;
                for ($m = 0; $m < $seg; $m++) {
                    $q = $items[$i + $k + $m];
                    $sb = ($this->pgEmpty && $this->autoBreak && $m === 0) ? 0.0 : $q['sb'];
                    if ($yy + $sb + $q['h'] > $this->bot + $eps) {
                        break;
                    }
                    $yy += $sb + $q['h'];
                    $fit++;
                }
                if ($fit >= $seg) {
                    for ($m = 0; $m < $seg; $m++) {
                        $this->put($items[$i + $k + $m], $X);
                    }
                    $k += $seg;
                    if ($items[$i + $k - 1]['ln']['pb'] && ($k < $cnt || $i + $cnt < $n || $moreAfter)) {
                        $this->newPage(true);
                    }
                    continue;
                }
                $raw = $fit;
                if ($pp['keepLines']) {
                    $fit = 0;
                } elseif ($pp['widow']) {
                    if ($seg - $fit === 1) {
                        $fit--; // niente ultima riga isolata in cima alla pagina seguente
                    }
                    if ($k === 0 && $fit === 1) {
                        $fit = 0; // niente prima riga isolata in fondo alla pagina
                    }
                }
                if ($fit <= 0) {
                    if (!$this->pgEmpty) {
                        $this->newPage(false);
                        continue;
                    }
                    $fit = max(1, $raw);
                }
                for ($m = 0; $m < $fit; $m++) {
                    $this->put($items[$i + $k + $m], $X);
                }
                $k += $fit;
                $this->newPage(false);
            }
            $i += $cnt;
        }
    }

    // ------------------------------------------------------------------
    // Disposizione: paragrafi in righe, tabelle in righe di tabella
    // ------------------------------------------------------------------

    private static function stackH(array $items): float
    {
        $h = 0.0;
        foreach ($items as $it) {
            $h += $it['sb'] + $it['h'] + $it['sa'];
        }
        return $h;
    }

    private function stack(array $items, float $X, float $Y): void
    {
        foreach ($items as $it) {
            $Y += $it['sb'];
            $this->drawItem($it, $X, $Y);
            $Y += $it['h'] + $it['sa'];
        }
    }

    private function flow(array $blocks, float $W): array
    {
        $items = [];
        $prevSa = 0.0;
        foreach ($blocks as $bi => $b) {
            if ($b['k'] === 't') {
                array_push($items, ...$this->layT($b, $W));
                continue;
            }
            $pp = $b['pp'];
            $prev = $blocks[$bi - 1] ?? null;
            $next = $blocks[$bi + 1] ?? null;
            $sb = $pp['sb'];
            $sa = $pp['sa'];
            if ($pp['ctx']) {
                if ($prev !== null && $prev['k'] === 'p' && $prev['style'] === $b['style']) {
                    $sb = 0.0;
                }
                if ($next !== null && $next['k'] === 'p' && $next['style'] === $b['style']) {
                    $sa = 0.0;
                }
            }
            if ($this->collapse && $prev !== null && $prev['k'] === 'p') {
                // Lo spazio "dopo" del paragrafo precedente e lo spazio "prima" di questo non si sommano
                $sb = max(0.0, $sb - $prevSa);
            }
            $prevSa = $sa;
            $bd = $pp['bdr'];
            $same = static fn (?array $o): bool => $o !== null && $o['k'] === 'p' && $o['pp']['bdr'] == $bd
                && abs($o['pp']['indL'] - $pp['indL']) < 0.01 && abs($o['pp']['indR'] - $pp['indR']) < 0.01;
            $bt = $bd && self::bdrOn($bd['top'] ?? null) && !$same($prev);
            $bb = $bd && self::bdrOn($bd['bottom'] ?? null) && !$same($next);
            $lines = $this->layP($b, $W);
            $cnt = count($lines);
            $off = 0.0;
            foreach ($lines as $li => $ln) {
                $padT = ($li === 0 && $bt) ? $bd['top']['sp'] + self::bdrW($bd['top']) : 0.0;
                $padB = ($li === $cnt - 1 && $bb) ? $bd['bottom']['sp'] + self::bdrW($bd['bottom']) : 0.0;
                $items[] = [
                    'k' => 'l', 'h' => $ln['h'] + $padT + $padB,
                    'sb' => $li === 0 ? $sb : 0.0, 'sa' => $li === $cnt - 1 ? $sa : 0.0,
                    'ln' => $ln, 'pp' => $pp, 'n' => $cnt, 'W' => $W, 'ptop' => $off,
                    'padT' => $padT, 'bt' => $li === 0 && $bt, 'bb' => $li === $cnt - 1 && $bb,
                ];
                $off += $ln['h'] + $padT + $padB;
            }
        }
        return $items;
    }

    /** Prossima tabulazione dopo la posizione $abs (dal bordo sinistro dell'area di testo). */
    private function tabStop(float $abs, array $pp): array
    {
        $cand = null;
        foreach ($pp['tabs'] as $t) {
            if ($t[0] > $abs + 0.05) {
                $cand = $t;
                break;
            }
        }
        if ($pp['indF'] < 0 && $abs < $pp['indL'] - 0.05 && ($cand === null || $pp['indL'] < $cand[0])) {
            return [$pp['indL'], 'left', null, false];
        }
        if ($cand !== null) {
            return [$cand[0], $cand[1], $cand[2], true];
        }
        $d = $this->defTab;
        return [(floor(($abs + 0.05) / $d) + 1) * $d, 'left', null, false];
    }

    /** Posiziona gli atomi da $i a $j su una riga che inizia a $org. */
    private function place(array $A, int $i, int $j, float $org, array $pp): array
    {
        $out = [];
        $x = 0.0;
        $lim = 0.0;
        for ($k = $i; $k <= $j; $k++) {
            $a = $A[$k];
            $leader = null;
            if ($a['t'] === 'tab') {
                [$pos, $type, $leader, $custom] = $this->tabStop($org + $x, $pp);
                $w = $pos - ($org + $x);
                if ($custom && $type !== 'left') {
                    // Una tabulazione impostata oltre il margine destro resta valida: il testo sporge
                    $lim = max($lim, $pos - $org);
                }
                if ($type !== 'left') {
                    // Tabulazione destra, centrata o decimale: dipende dal testo che segue
                    $fw = 0.0;
                    $dec = null;
                    $trail = 0.0;
                    for ($m = $k + 1; $m <= $j && $A[$m]['t'] !== 'tab' && $A[$m]['t'] !== 'br'; $m++) {
                        if ($type === 'decimal' && $dec === null && $A[$m]['t'] === 'x'
                            && ($p = mb_strpos($A[$m]['s'], $this->decimal)) !== false) {
                            $dec = $fw + $A[$m]['f']->width(mb_substr($A[$m]['s'], 0, $p));
                        }
                        $fw += $A[$m]['w'];
                        $trail = $A[$m]['t'] === 'sp' ? $trail + $A[$m]['w'] : 0.0;
                    }
                    $fw -= $trail;
                    $need = $type === 'center' ? $fw / 2 : (($type === 'decimal' && $dec !== null) ? $dec : $fw);
                    $w = max(0.0, $w - $need);
                }
            } else {
                $w = $a['w'] ?? 0.0;
            }
            $out[] = [$a, $x, $w, $leader];
            $x += $w;
        }
        $cw = $x;
        for ($m = count($out) - 1; $m >= 0; $m--) {
            $t = $out[$m][0]['t'];
            if ($t === 'sp') {
                $cw -= $out[$m][2];
            } elseif ($t !== 'anc') {
                break;
            }
        }
        // Larghezza degli spazi interni dopo l'ultima tabulazione (quelli che la giustificazione può modificare)
        $sw = 0.0;
        for ($q = $m - 1; $q >= 0 && $out[$q][0]['t'] !== 'tab'; $q--) {
            if ($out[$q][0]['t'] === 'sp') {
                $sw += $out[$q][2];
            }
        }
        return ['p' => $out, 'w' => $cw, 'sw' => $sw, 'lim' => $lim];
    }

    private function layP(array $b, float $W): array
    {
        $pp = $b['pp'];
        $A = $b['A'];
        foreach ($A as $k => $a) {
            if ($a['t'] === 'fld') {
                if (!isset($this->fv[$a['code']]) && $a['code'] !== 'PAGE') {
                    $this->needTotal = true;
                }
                $s = (string) ($this->fv[$a['code']] ?? $a['s']);
                $A[$k]['s'] = $s;
                $A[$k]['w'] = $a['f']->width($s);
            }
        }
        $n = count($A);
        $lines = [];
        $i = 0;
        $extra = false;
        $shrink = $pp['jc'] === 'both' && $this->compat >= 15;
        while (true) {
            $org = $pp['indL'] + ($lines ? 0.0 : $pp['indF']);
            $avail = max(1.0, $W - $pp['indR'] - $org);
            if ($i >= $n) {
                if (!$lines || $extra) {
                    $lines[] = $this->mkLine([], 0.0, $org, $avail, $b, true, null);
                }
                break;
            }
            $extra = false;
            $forced = null;
            $ok = ['p' => [], 'w' => 0.0, 'sw' => 0.0, 'lim' => 0.0];
            $j = $i;
            $next = null;
            while ($j < $n) {
                $a = $A[$j];
                if ($a['t'] === 'br') {
                    $forced = $a['kind'];
                    $next = $j + 1;
                    break;
                }
                $pl = $this->place($A, $i, $j, $org, $pp);
                // Word 2013+ nei paragrafi giustificati stringe gli spazi (fino al 20%) per far entrare una parola in più
                $slack = $shrink ? $pl['sw'] * 0.2 + 0.1 : 0.0;
                if ($pl['w'] <= max($avail, $pl['lim']) + 0.02 + $slack || $a['t'] === 'sp' || $a['t'] === 'anc') {
                    $ok = $pl;
                    $j++;
                    continue;
                }
                // L'atomo $j non entra: cerca l'ultimo punto in cui si può andare a capo
                $bkAt = null;
                for ($k = $j; $k > $i; $k--) {
                    $t = $A[$k]['t'];
                    if ($t === 'img' || $t === 'tab' || (($t === 'x' || $t === 'fld') && $A[$k]['bk'])) {
                        $bkAt = $k;
                        break;
                    }
                }
                if ($bkAt !== null) {
                    $ok = $this->place($A, $i, $bkAt - 1, $org, $pp);
                    $next = $bkAt;
                } elseif ($j > $i) {
                    $next = $j;
                } elseif ($a['t'] === 'x' && mb_strlen($a['s'], 'UTF-8') > 1) {
                    // Parola più larga della riga: si spezza tra i caratteri
                    $len = mb_strlen($a['s'], 'UTF-8');
                    $cut = 1;
                    for ($c = 2; $c < $len; $c++) {
                        if ($a['f']->width(mb_substr($a['s'], 0, $c, 'UTF-8')) > $avail) {
                            break;
                        }
                        $cut = $c;
                    }
                    $head = mb_substr($a['s'], 0, $cut, 'UTF-8');
                    $tail = mb_substr($a['s'], $cut, null, 'UTF-8');
                    $A[$j] = ['s' => $head, 'w' => $a['f']->width($head)] + $a;
                    array_splice($A, $j + 1, 0, [['s' => $tail, 'w' => $a['f']->width($tail), 'bk' => true] + $a]);
                    $n++;
                    $ok = $this->place($A, $i, $j, $org, $pp);
                    $next = $j + 1;
                } else {
                    $ok = $pl;
                    $next = $j + 1;
                }
                break;
            }
            $i = $next ?? $n;
            $isLast = $i >= $n;
            if ($isLast && $forced === 'line') {
                // Dopo un a capo manuale finale resta la riga vuota con il segno di paragrafo
                // (dopo un'interruzione di pagina, invece, il segno resta sulla pagina dell'interruzione)
                $isLast = false;
                $extra = true;
            }
            $lines[] = $this->mkLine($ok['p'], $ok['w'], $org, max($avail, $ok['lim']), $b, $isLast, $forced);
        }
        return $lines;
    }

    private function mkLine(array $placed, float $cw, float $org, float $avail, array $b, bool $isLast, ?string $forced): array
    {
        $pp = $b['pp'];
        $asc = $desc = $lead = 0.0;
        $has = false;
        foreach ($placed as [$a]) {
            if ($a['t'] === 'anc') {
                continue;
            }
            $f = $a['f'];
            $asc = max($asc, $a['t'] === 'img' ? $a['h'] : $f->asc);
            $desc = max($desc, $f->desc);
            $lead = max($lead, $f->lead);
            $has = true;
        }
        if (!$has || ($isLast && $this->markInLine)) {
            $m = $b['mark'];
            $asc = max($asc, $m->asc);
            $desc = max($desc, $m->desc);
            $lead = max($lead, $m->lead);
        }
        $nat = $asc + $desc + $lead;
        if ($pp['lineRule'] === 'exact') {
            $h = max(1.0, $pp['line'] / 20);
            $base = $h - $desc;
        } elseif ($pp['lineRule'] === 'atLeast') {
            $h = max($nat, $pp['line'] / 20);
            $base = $h - $desc - $lead;
        } else {
            $h = $nat * $pp['line'] / 240;
            $base = $asc - max(0.0, $nat - $h);
        }
        $free = $avail - $cw;
        $x0 = $org;
        if ($pp['jc'] === 'center') {
            $x0 += $free / 2;
        } elseif ($pp['jc'] === 'right') {
            $x0 += $free;
        } elseif ($pp['jc'] === 'both' && ($free < -0.01
            || ($free > 0.01 && !$isLast && ($forced === null || ($forced === 'line' && $this->expandBr))))) {
            // Giustificato: lo spazio libero (o in eccesso) va negli spazi dopo l'ultima tabulazione
            $from = 0;
            $to = -1;
            foreach ($placed as $k => [$a]) {
                if ($a['t'] === 'tab') {
                    $from = $k + 1;
                }
                if ($a['t'] !== 'sp' && $a['t'] !== 'anc') {
                    $to = $k;
                }
            }
            $spaces = 0;
            for ($k = $from; $k < $to; $k++) {
                if ($placed[$k][0]['t'] === 'sp') {
                    $spaces += strlen($placed[$k][0]['s']);
                }
            }
            if ($spaces > 0) {
                $add = $free / $spaces;
                $shift = 0.0;
                foreach ($placed as $k => $pl) {
                    $placed[$k][1] += $shift;
                    if ($k >= $from && $k < $to && $pl[0]['t'] === 'sp') {
                        $e = $add * strlen($pl[0]['s']);
                        $placed[$k][2] += $e;
                        $shift += $e;
                    }
                }
            }
        }
        return ['p' => $placed, 'h' => $h, 'base' => $base, 'x0' => $x0, 'pb' => $forced === 'page'];
    }

    private function layT(array $t, float $W): array
    {
        $grid = $t['grid'];
        $nc = count($grid);
        if ($nc === 0) {
            return [];
        }
        $zero = count(array_filter($grid, static fn ($g) => $g <= 0));
        if ($zero > 0) {
            $rest = max(10.0, $W - array_sum($grid));
            foreach ($grid as $k => $g) {
                if ($g <= 0) {
                    $grid[$k] = $rest / $zero;
                }
            }
        }
        $tw = array_sum($grid);
        $x0 = $t['ind'] - ($this->compat < 15 ? $t['mar']['l'] : 0.0);
        if ($t['jc'] === 'center') {
            $x0 = ($W - $tw) / 2;
        } elseif ($t['jc'] === 'right' || $t['jc'] === 'end') {
            $x0 = $W - $tw;
        }
        $colX = [0.0];
        foreach ($grid as $g) {
            $colX[] = end($colX) + $g;
        }
        $tb = $t['bdr'];
        $rows = [];
        $open = [];
        $nr = count($t['rows']);
        foreach ($t['rows'] as $ri => $row) {
            $cells = [];
            $col = $row['before'];
            $maxH = 0.0;
            foreach ($row['cells'] as $cell) {
                if ($col >= $nc) {
                    break;
                }
                $span = min($cell['span'], $nc - $col);
                $cw = $colX[$col + $span] - $colX[$col];
                $m = $cell['mar'];
                $cb = $cell['bdr'];
                $lc = [
                    'x' => $colX[$col], 'w' => $cw, 'col' => $col, 'mar' => $m, 'va' => $cell['va'],
                    'shd' => $cell['shd'], 'vm' => $cell['vm'], 'items' => [], 'ch' => 0.0, 'rows' => 1,
                    'edges' => [
                        'top' => $cell['vm'] === 'c' ? null : ($cb['top'] ?? ($ri === 0 ? ($tb['top'] ?? null) : ($tb['insideH'] ?? null))),
                        'bottom' => $cb['bottom'] ?? ($ri === $nr - 1 ? ($tb['bottom'] ?? null) : ($tb['insideH'] ?? null)),
                        'left' => $cb['left'] ?? ($col === 0 ? ($tb['left'] ?? null) : ($tb['insideV'] ?? null)),
                        'right' => $cb['right'] ?? ($col + $span >= $nc ? ($tb['right'] ?? null) : ($tb['insideV'] ?? null)),
                    ],
                ];
                if ($cell['vm'] !== 'c') {
                    $lc['items'] = $this->flow($cell['blocks'], max(1.0, $cw - $m['l'] - $m['r']));
                    $lc['ch'] = self::stackH($lc['items']) + $m['t'] + $m['b'];
                }
                if ($cell['vm'] === 'r') {
                    $open[$col] = [$ri, count($cells)];
                } elseif ($cell['vm'] === 'c') {
                    if (isset($open[$col])) {
                        [$r0, $c0] = $open[$col];
                        $rows[$r0]['cells'][$c0]['rows']++;
                        $rows[$ri - 1]['noBottom'][$col] = true;
                    }
                } else {
                    unset($open[$col]);
                    $maxH = max($maxH, $lc['ch']);
                }
                $cells[] = $lc;
                $col += $span;
            }
            $h = $maxH;
            if ($row['h'] !== null) {
                $h = $row['hRule'] === 'exact' ? $row['h'] : max($h, $row['h']);
            }
            // Come in Word, il bordo superiore occupa spazio: il contenuto parte sotto il bordo
            $bt = 0.0;
            foreach ($cells as $c) {
                $bt = max($bt, self::bdrW($c['edges']['top']));
            }
            $rows[] = ['h' => $h + $bt, 'bt' => $bt, 'cells' => $cells, 'hdr' => $row['hdr'], 'noBottom' => []];
        }
        // Celle unite in verticale: l'altezza complessiva deve contenere il contenuto
        foreach ($rows as $ri => $row) {
            foreach ($row['cells'] as $ci => $c) {
                if (!empty($row['noBottom'][$c['col']])) {
                    $rows[$ri]['cells'][$ci]['edges']['bottom'] = null;
                }
                if ($c['vm'] !== 'r') {
                    continue;
                }
                $last = min(count($rows) - 1, $ri + $c['rows'] - 1);
                $sum = 0.0;
                for ($k = $ri; $k <= $last; $k++) {
                    $sum += $rows[$k]['h'];
                }
                if ($c['ch'] > $sum) {
                    $rows[$last]['h'] += $c['ch'] - $sum;
                    $sum = $c['ch'];
                }
                $rows[$ri]['cells'][$ci]['spanH'] = $sum;
            }
        }
        foreach ($rows as $ri => $row) {
            foreach ($row['cells'] as $ci => $c) {
                if ($c['vm'] === 'r') {
                    $sum = 0.0;
                    for ($k = $ri; $k < $ri + $c['rows'] && $k < count($rows); $k++) {
                        $sum += $rows[$k]['h'];
                    }
                    $rows[$ri]['cells'][$ci]['spanH'] = $sum;
                }
            }
        }
        $items = [];
        $hdrRows = [];
        foreach ($rows as $row) {
            $it = ['k' => 'r', 'h' => $row['h'], 'sb' => 0.0, 'sa' => 0.0, 'row' => $row, 'x0' => $x0, 'hdrRows' => []];
            if ($row['hdr'] && count($hdrRows) === count($items)) {
                $hdrRows[] = $it;
            }
            $it['hdrRows'] = $row['hdr'] ? [] : $hdrRows;
            $items[] = $it;
        }
        return $items;
    }

    // ------------------------------------------------------------------
    // Disegno
    // ------------------------------------------------------------------

    private function out(string $s): void
    {
        $this->pages[$this->pi][$this->layer] .= $s;
    }

    private static function rgb(string $hex): string
    {
        return DocxPdfWriter::n(hexdec(substr($hex, 0, 2)) / 255) . ' '
            . DocxPdfWriter::n(hexdec(substr($hex, 2, 2)) / 255) . ' '
            . DocxPdfWriter::n(hexdec(substr($hex, 4, 2)) / 255);
    }

    private function rect(float $x, float $y, float $w, float $h, string $hex): void
    {
        $n = [DocxPdfWriter::class, 'n'];
        $this->out(self::rgb($hex) . ' rg ' . $n($x) . ' ' . $n($this->ph - $y - $h) . ' ' . $n($w) . ' ' . $n($h) . " re f\n");
    }

    private function stroke(float $x1, float $y1, float $x2, float $y2, float $w, string $hex, ?string $dash = null, int $cap = 0): void
    {
        $n = [DocxPdfWriter::class, 'n'];
        $d = $dash === 'dot' ? '[' . $n($w) . ' ' . $n($w * 2) . '] 0 d ' : ($dash === 'dash' ? '[' . $n($w * 4) . ' ' . $n($w * 3) . '] 0 d ' : '');
        $this->out('q ' . $n($w) . ' w ' . $cap . ' J ' . self::rgb($hex) . ' RG ' . $d
            . $n($x1) . ' ' . $n($this->ph - $y1) . ' m ' . $n($x2) . ' ' . $n($this->ph - $y2) . " l S Q\n");
    }

    /** Linea di bordo (orizzontale o verticale) nello stile indicato. */
    private function bline(float $x1, float $y1, float $x2, float $y2, ?array $b, int $cap = 0): void
    {
        if (!self::bdrOn($b)) {
            return;
        }
        $w = $b['w'];
        $dash = match (true) {
            str_contains($b['v'], 'dot') => 'dot',
            str_contains($b['v'], 'dash') => 'dash',
            default => null,
        };
        if ($b['v'] === 'double') {
            $dx = abs($y2 - $y1) > abs($x2 - $x1) ? $w : 0.0;
            $dy = $dx > 0 ? 0.0 : $w;
            $this->stroke($x1 - $dx, $y1 - $dy, $x2 - $dx, $y2 - $dy, $w, $b['c'], null, $cap);
            $this->stroke($x1 + $dx, $y1 + $dy, $x2 + $dx, $y2 + $dy, $w, $b['c'], null, $cap);
            return;
        }
        $this->stroke($x1, $y1, $x2, $y2, $w, $b['c'], $dash, $cap);
    }

    private function txt(float $x, float $base, string $s, DocxPdfFmt $f): void
    {
        if ($s === '') {
            return;
        }
        $n = [DocxPdfWriter::class, 'n'];
        $res = $this->pdf->fontRes($f->font);
        $hex = $f->font->encode($s);
        $y = $this->ph - $base + $f->rise;
        $col = self::rgb($f->color);
        $special = $f->sp != 0.0 || $f->scale != 1.0 || $f->fakeB;
        $o = ($special ? 'q ' : '') . "BT /$res " . $n($f->size) . " Tf $col rg ";
        if ($f->sp != 0.0) {
            $o .= $n($f->sp) . ' Tc ';
        }
        if ($f->scale != 1.0) {
            $o .= $n($f->scale * 100) . ' Tz ';
        }
        if ($f->fakeB) {
            $o .= '2 Tr ' . $n($f->size * 0.03) . " w $col RG ";
        }
        $o .= $f->fakeI ? '1 0 0.2126 1 ' . $n($x) . ' ' . $n($y) . ' Tm ' : $n($x) . ' ' . $n($y) . ' Td ';
        $this->out($o . "<$hex> Tj ET" . ($special ? ' Q' : '') . "\n");
    }

    private function drawItem(array $it, float $X, float $Y): void
    {
        if ($it['k'] === 'r') {
            $this->drawRow($it, $X, $Y);
        } else {
            $this->drawLine($it, $X, $Y);
        }
    }

    private function drawLine(array $it, float $X, float $Y): void
    {
        $ln = $it['ln'];
        $pp = $it['pp'];
        $h = $it['h'];
        $l = $X + $pp['indL'];
        $r = $X + $it['W'] - $pp['indR'];
        if ($pp['shd'] !== null) {
            $this->rect($l, $Y, $r - $l, $h, $pp['shd']);
        }
        $bd = $pp['bdr'];
        if ($bd) {
            $bl = $l - ($bd['left']['sp'] ?? 0);
            $br = $r + ($bd['right']['sp'] ?? 0);
            if ($it['bt']) {
                $this->bline($bl, $Y + self::bdrW($bd['top']) / 2, $br, $Y + self::bdrW($bd['top']) / 2, $bd['top']);
            }
            if ($it['bb']) {
                $this->bline($bl, $Y + $h - self::bdrW($bd['bottom']) / 2, $br, $Y + $h - self::bdrW($bd['bottom']) / 2, $bd['bottom']);
            }
            $this->bline($bl, $Y, $bl, $Y + $h, $bd['left'] ?? null);
            $this->bline($br, $Y, $br, $Y + $h, $bd['right'] ?? null);
        }
        $top = $Y + $it['padT'];
        $base = $top + $ln['base'];
        foreach ($ln['p'] as [$a, $x, $w, $leader]) {
            $ax = $X + $ln['x0'] + $x;
            $f = $a['f'];
            switch ($a['t']) {
                case 'anc':
                    $this->anchor($a['d'], $ax, $top, $Y - $it['ptop'], $X, $it['W']);
                    continue 2;
                case 'img':
                    $this->drawShape($a['d'], $ax, $base - $a['h']);
                    continue 2;
            }
            if ($w <= 0 && $a['t'] !== 'x') {
                continue;
            }
            if ($f->hl !== null) {
                $this->rect($ax, $base - $f->asc, $w, $f->asc + $f->desc, $f->hl);
            }
            if ($a['t'] === 'tab') {
                $ch = ['dot' => '.', 'hyphen' => '-', 'underscore' => '_', 'heavy' => '_', 'middleDot' => "\u{00B7}"][$leader ?? ''] ?? null;
                if ($ch !== null) {
                    [$u] = $f->font->measure($ch);
                    $cwd = $u * $f->size / $f->font->upm;
                    $cnt = $cwd > 0 ? (int) floor(($w + 0.01) / $cwd) : 0;
                    if ($cnt > 0) {
                        // Riempimento senza spaziatura caratteri, allineato alla fine della tabulazione
                        $lf = ($f->sp != 0.0 || $f->scale != 1.0) ? $this->variant($f, ['sp' => 0, 'scale' => 100]) : $f;
                        $this->txt($ax + $w - $cnt * $cwd, $base, str_repeat($ch, $cnt), $lf);
                    }
                }
            } elseif ($a['t'] === 'fld') {
                $this->txt($ax, $base, (string) ($this->dfv[$a['code']] ?? $a['s']), $f);
            } else {
                $this->txt($ax, $base, $a['s'], $f);
            }
            if ($f->u) {
                $t = $f->font->ulThick * $f->size;
                $this->rect($ax, $base - $f->font->ulPos * $f->size - $t / 2, $w, $t, $f->color);
            }
            if ($f->strike) {
                $t = $f->font->ulThick * $f->size;
                $this->rect($ax, $base - $f->size * 0.27, $w, $t, $f->color);
            }
        }
    }

    /** Oggetto flottante: calcola la posizione dai riferimenti di pagina, margine, colonna o paragrafo. */
    private function anchor(array $d, float $ax, float $lineTop, float $paraTop, float $colX, float $colW): void
    {
        $s = $this->pages[$this->pi]['sec'];
        $pw = $s['pw'];
        $ph = $s['ph'];
        $mt = abs($s['mt']);
        $mb = abs($s['mb']);
        [$rx, $rw] = match ($d['relH']) {
            'page' => [0.0, $pw],
            'margin' => [$s['ml'], $pw - $s['ml'] - $s['mr']],
            'leftMargin', 'insideMargin' => [0.0, $s['ml']],
            'rightMargin', 'outsideMargin' => [$pw - $s['mr'], $s['mr']],
            'character' => [$ax, 0.0],
            default => [$colX, $colW],
        };
        $x = match ($d['alignH']) {
            null => $rx + $d['offH'],
            'center' => $rx + ($rw - $d['w']) / 2,
            'right', 'outside' => $rx + $rw - $d['w'],
            default => $rx,
        };
        [$ry, $rh] = match ($d['relV']) {
            'page' => [0.0, $ph],
            'margin' => [$mt, $ph - $mt - $mb],
            'topMargin' => [0.0, $mt],
            'bottomMargin' => [$ph - $mb, $mb],
            'line' => [$lineTop, 0.0],
            default => [$paraTop, 0.0],
        };
        $y = match ($d['alignV']) {
            null => $ry + $d['offV'],
            'center' => $ry + ($rh - $d['h']) / 2,
            'bottom', 'outside' => $ry + $rh - $d['h'],
            default => $ry,
        };
        $save = $this->layer;
        $this->layer = $d['behind'] ? 'back' : 'front';
        $this->drawShape($d, $x, $y);
        $this->layer = $save;
    }

    private function drawImg(?string $path, array $crop, float $x, float $y, float $w, float $h): void
    {
        if ($path === null || $w <= 0 || $h <= 0) {
            return;
        }
        $bytes = $this->zip->getFromName($path);
        $img = $bytes === false ? null : $this->pdf->image(md5($bytes), $bytes);
        if ($img === null) {
            $this->note('Immagine in formato non supportato (' . basename($path) . '): ignorata');
            return;
        }
        $n = [DocxPdfWriter::class, 'n'];
        $fw = $w / max(0.01, 1 - $crop['l'] - $crop['r']);
        $fh = $h / max(0.01, 1 - $crop['t'] - $crop['b']);
        $fx = $x - $crop['l'] * $fw;
        $fy = $y - $crop['t'] * $fh;
        $this->out('q ' . $n($x) . ' ' . $n($this->ph - $y - $h) . ' ' . $n($w) . ' ' . $n($h) . ' re W n '
            . $n($fw) . ' 0 0 ' . $n($fh) . ' ' . $n($fx) . ' ' . $n($this->ph - $fy - $fh) . " cm /{$img[0]} Do Q\n");
    }

    private function drawShape(array $d, float $x, float $y): void
    {
        $n = [DocxPdfWriter::class, 'n'];
        $w = $d['w'];
        $h = $d['h'];
        $ln = $d['line'] ?? null;
        if ($d['kind'] === 'pic') {
            $this->drawImg($d['img'], $d['crop'], $x, $y, $w, $h);
            if ($ln !== null) {
                $this->out('q ' . $n($ln['w']) . ' w ' . self::rgb($ln['c']) . ' RG ' . $n($x) . ' ' . $n($this->ph - $y - $h) . ' ' . $n($w) . ' ' . $n($h) . " re S Q\n");
            }
            return;
        }
        if (in_array($d['geom'], ['line', 'straightConnector1'], true)) {
            if ($ln !== null) {
                [$y1, $y2] = $d['flipV'] ? [$y + $h, $y] : [$y, $y + $h];
                [$x1, $x2] = $d['flipH'] ? [$x + $w, $x] : [$x, $x + $w];
                $this->stroke($x1, $y1, $x2, $y2, $ln['w'], $ln['c'], $ln['dash']);
            }
            return;
        }
        $py = $this->ph - $y - $h;
        if ($d['geom'] === 'ellipse') {
            $k = 0.5523;
            $rx = $w / 2;
            $ry = $h / 2;
            $cx = $x + $rx;
            $cy = $py + $ry;
            $path = $n($cx + $rx) . ' ' . $n($cy) . ' m '
                . $n($cx + $rx) . ' ' . $n($cy + $k * $ry) . ' ' . $n($cx + $k * $rx) . ' ' . $n($cy + $ry) . ' ' . $n($cx) . ' ' . $n($cy + $ry) . ' c '
                . $n($cx - $k * $rx) . ' ' . $n($cy + $ry) . ' ' . $n($cx - $rx) . ' ' . $n($cy + $k * $ry) . ' ' . $n($cx - $rx) . ' ' . $n($cy) . ' c '
                . $n($cx - $rx) . ' ' . $n($cy - $k * $ry) . ' ' . $n($cx - $k * $rx) . ' ' . $n($cy - $ry) . ' ' . $n($cx) . ' ' . $n($cy - $ry) . ' c '
                . $n($cx + $k * $rx) . ' ' . $n($cy - $ry) . ' ' . $n($cx + $rx) . ' ' . $n($cy - $k * $ry) . ' ' . $n($cx + $rx) . ' ' . $n($cy) . ' c h';
        } else {
            $path = $n($x) . ' ' . $n($py) . ' ' . $n($w) . ' ' . $n($h) . ' re';
        }
        if (!empty($d['fill'])) {
            $this->out(self::rgb($d['fill']) . " rg $path f\n");
        }
        if (!empty($d['img'])) {
            $this->drawImg($d['img'], $d['crop'], $x, $y, $w, $h);
        }
        if ($ln !== null) {
            $dash = $ln['dash'] === 'dot' ? '[' . $n($ln['w']) . ' ' . $n($ln['w'] * 2) . '] 0 d ' : ($ln['dash'] === 'dash' ? '[' . $n($ln['w'] * 4) . ' ' . $n($ln['w'] * 3) . '] 0 d ' : '');
            $this->out('q ' . $n($ln['w']) . ' w ' . self::rgb($ln['c']) . " RG $dash$path S Q\n");
        }
        if (!empty($d['txbx'])) {
            [$il, $it, $ir, $ib] = $d['ins'];
            $items = $this->flow($d['txbx'], max(1.0, $w - $il - $ir));
            $ch = self::stackH($items);
            $off = match ($d['anchor']) {
                'ctr' => max(0.0, ($h - $it - $ib - $ch) / 2),
                'b' => max(0.0, $h - $it - $ib - $ch),
                default => 0.0,
            };
            $this->stack($items, $x + $il, $y + $it + $off);
        }
    }

    private function drawRow(array $it, float $X, float $Y): void
    {
        $x0 = $X + $it['x0'];
        $h = $it['h'];
        foreach ($it['row']['cells'] as $c) {
            if ($c['vm'] === 'c') {
                continue;
            }
            $cx = $x0 + $c['x'];
            $H = $c['spanH'] ?? $h;
            if ($c['shd'] !== null) {
                $this->rect($cx, $Y, $c['w'], $H, $c['shd']);
            }
            $m = $c['mar'];
            $off = match ($c['va']) {
                'center' => max(0.0, ($H - $c['ch']) / 2),
                'bottom' => max(0.0, $H - $c['ch']),
                default => 0.0,
            };
            $bt = $it['row']['bt'];
            $off = $c['va'] === 'top' ? $bt : max($bt, $off);
            $this->stack($c['items'], $cx + $m['l'], $Y + $m['t'] + $off);
        }
        foreach ($it['row']['cells'] as $c) {
            $cx = $x0 + $c['x'];
            $e = $c['edges'];
            $this->bline($cx, $Y, $cx + $c['w'], $Y, $e['top'], 2);
            $this->bline($cx, $Y + $h, $cx + $c['w'], $Y + $h, $e['bottom'], 2);
            $this->bline($cx, $Y, $cx, $Y + $h, $e['left'], 2);
            $this->bline($cx + $c['w'], $Y, $cx + $c['w'], $Y + $h, $e['right'], 2);
        }
    }
}
