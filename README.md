# FZPureDOCX — DOCX and PDF documents from Word templates

**English** · [Italiano](README.it.md)

A small PHP server that fills Word templates (`.docx`) by replacing
placeholders with text, HTML, tables and checkboxes, and returns the
document as **DOCX** or **PDF**.

Applications call it with a plain HTTP POST request and get a JSON response.

- No external libraries, no licences to buy.
- PDF through **LibreOffice**, when it is installed on the server, or through
  the bundled **pure-PHP converter**, which needs no external programs.
- No database. No files are left on the server after a request.

> The project comes from an Italian codebase: API field names (`modello_docx`,
> `converti_pdf`, `esito`, `avvisi`…), configuration keys and error messages
> are in Italian. Each one is explained below.

## Contents

- [How it works](#how-it-works)
- [Requirements](#requirements)
- [Installation](#installation)
- [Fonts](#fonts)
- [Configuration](#configuration)
- [Checking the installation](#checking-the-installation)
- [Calling the server](#calling-the-server)
- [Template placeholders](#template-placeholders)
- [API reference](#api-reference)
- [PDF conversion: LibreOffice or PHP](#pdf-conversion-libreoffice-or-php)
- [Security](#security)
- [Project layout](#project-layout)
- [Support the project](#support-the-project)
- [Licence](#licence)

## How it works

```
Application                            FZPureDOCX server
       ── POST (base64 template + variables) ──►  gdocx.php
       ◄──────── JSON (base64 DOCX and/or PDF) ──  │
                                                   ├─ GDocx       fills the template
                                                   ├─ DocTemplate replaces the placeholders
                                                   └─ DocxPdf     DOCX → PDF in PHP
                                                      (or LibreOffice, if present)
```

## Requirements

- PHP **8.1 or later**
- PHP extensions: `zip`, `dom`, `zlib`, `mbstring` (required);
  `curl` (for HTML images given by URL) and `gd` (for GIF/WebP/BMP)
  optional
- Apache with `.htaccess` enabled (for nginx see [Security](#security))
- Some TrueType fonts (see [Fonts](#fonts))
- Optional: LibreOffice (`soffice`) for PDF conversion

Ordinary shared hosting works too.

## Installation

1. Download the project (*Code → Download ZIP*) and extract it.
2. Download the fonts and copy them into `lib/fonts/` (see [Fonts](#fonts)).
3. Copy `_apikeys.example.php` to **`_apikeys.php`** and add your keys (see
   [Configuration](#configuration)).
4. Upload everything to a web-served folder, for example
   `public_html/fzpuredocx/`. The hidden files `.htaccess` and `.user.ini`
   must be uploaded too: FTP clients usually need *show hidden files*
   turned on.
5. Make the `log/` folder writable by the web server user (or turn logging
   off in `_config.php`).
6. Open `https://your-server/fzpuredocx/verifica.php?apikey=KEY` and check
   that every line is green.

## Fonts

Font files are **not included** in the repository: download them from the
official sites and copy them into `lib/fonts/`.

The PHP converter **embeds** the fonts used by the document in the PDF and
uses their metrics to decide where lines break. It looks for them, in this
order: in the folders set in `_config.php`, in `lib/fonts/`, and among the
server's system fonts. If a font used by the document is missing, it uses a
similar one and reports it in `avvisi` (warnings); line breaks may then
differ from Word. **If it finds no fonts at all, PDF conversion fails.**

### Base fonts (recommended)

Free fonts that stand in for the most common ones.

| Font | Replaces | Download | Files to copy into `lib/fonts/` |
|------|----------|----------|---------------------------------|
| Liberation Sans and Serif | Arial, Helvetica, Times New Roman | <https://github.com/liberationfonts/liberation-fonts/releases> → `liberation-fonts-ttf-…tar.gz` | `LiberationSans-*.ttf`, `LiberationSerif-*.ttf` |
| DejaVu Sans | Verdana; symbols (☐ ☒ • ◦ ▪ …) | <https://github.com/dejavu-fonts/dejavu-fonts/releases> → `dejavu-fonts-ttf-…zip` (`ttf/` folder) | `DejaVuSans.ttf`, `DejaVuSans-Bold.ttf` |

Alternatively, on a Debian/Ubuntu server you manage, install them as system
fonts:

```bash
apt install fonts-liberation fonts-dejavu-core
```

### Fonts used in your templates

For a PDF that **matches Word exactly**, the fonts chosen in your templates
are needed too. Copy their `.ttf` or `.ttc` files into `lib/fonts/`;
`.otf` fonts are not supported.

- **Free fonts**, for example Roboto: download them from
  [Google Fonts](https://fonts.google.com/specimen/Roboto) (*Get font →
  Download*). Use the files in the `static/` folder, such as
  `Roboto-Regular.ttf`, `Roboto-Bold.ttf`, `Roboto-Light.ttf`.
- **Commercial fonts**, such as Microsoft's Calibri, Arial, Verdana and
  Tahoma: copy them from a Windows PC where they are installed
  (`C:\Windows\Fonts`), if their licence allows use on the server.

The [check page](#checking-the-installation) lists the fonts it finds and,
after the test conversion, the ones that were substituted.

LibreOffice needs the fonts as well, installed in the system: when they are
missing it substitutes them, sometimes inconsistently (serif bold text, very
wide line spacing).

## Configuration

### `_apikeys.php` — who may use the server (private, excluded from git)

```php
return [
    // IPs of the servers allowed to call the service (networks too: '198.51.100.0/24').
    // Empty list = no IP check, a valid key is enough.
    'ip_consentiti' => [
        '203.0.113.10',
    ],

    // API keys: 32 characters 0-9 a-f (lower case) and the tools each may use
    'chiavi' => [
        '0123456789abcdef0123456789abcdef' => ['tools' => ['gdocx', 'gdocx2pdf']],
    ],
];
```

`ip_consentiti` means *allowed IPs*, `chiavi` means *keys*. To create a key:

```bash
php -r 'echo bin2hex(random_bytes(16)), "\n";'
```

### `_config.php` — general settings

| Key            | Default            | Meaning |
|----------------|--------------------|---------|
| `debug`        | `'no'`             | `'si'` (yes) adds file and line to error messages (testing only) |
| `pdf_engine`   | `'auto'`           | `'auto'` LibreOffice if present, otherwise PHP · `'libreoffice'` · `'php'` |
| `soffice`      | `null`             | Path to LibreOffice; `null` = search automatically |
| `fonts`        | `[]`               | Extra `.ttf` font folders for the PHP converter |
| `log_file`     | `log/fzpuredocx.log` | Request log; `null` to turn it off |
| `memory_limit` | `'1024M'`          | Memory for processing |

### PHP limits

Documents travel base64-encoded inside the POST body, and tables can have
many rows. `.user.ini` (PHP-FPM/CGI) and `.htaccess` (mod_php) raise
`post_max_size` to 128M and `max_input_vars` to 20000. The check page shows
the values actually in effect.

## Checking the installation

```
https://your-server/fzpuredocx/verifica.php?apikey=YOUR_KEY
```

The page checks the PHP version, extensions, limits, folders, keys and
fonts, then **actually generates a test document** (a DOCX filled with HTML
and a table, converted to PDF) and reports the engine used and the time
taken. The *Apri il PDF di prova* button (open the test PDF) shows the
result. A valid key is required; the page is exempt from the IP check
because you open it from your own browser.

## Calling the server

PHP example with cURL. List fields (`vars`, `html_vars`, …) are sent as
arrays, encoding the body with `http_build_query`.

```php
$post = [
    'apikey'       => 'YOUR_KEY',
    'nomehost'     => 'my-site',                        // free label, only for the log
    'modello_docx' => base64_encode(file_get_contents('template.docx')),
    'vars'         => ['surname' => 'Smith', 'notes' => "Line 1\nLine 2"],
    'html_vars'    => ['body' => '<style>td{border:1px solid #000}</style><table><tr><td>…</td></tr></table>'],
    'headers_vars' => ['heading' => 'Quote no. 12'],
    'table_vars'   => ['products' => [
        ['code' => 'A1', 'name' => 'Chair', 'price' => '120.00'],
        ['code' => 'B2', 'name' => 'Desk',  'price' => '350.00'],
    ]],
    'converti_pdf' => 'si',                              // also return the PDF
];

$ch = curl_init('https://your-server/fzpuredocx/gdocx.php');
curl_setopt_array($ch, [
    CURLOPT_POST           => true,
    CURLOPT_POSTFIELDS     => http_build_query($post),
    CURLOPT_RETURNTRANSFER => true,
    CURLOPT_TIMEOUT        => 120,
]);
$response = json_decode((string) curl_exec($ch), true);

if (($response['esito'] ?? '') === 'OK') {
    file_put_contents('document.docx', base64_decode($response['data']));
    file_put_contents('document.pdf', base64_decode($response['data_pdf']));
} else {
    echo 'Error: ', $response['errore'] ?? 'invalid response';
}
```

## Template placeholders

In the Word template, placeholders are written as **`$name$`**. They work
even when Word internally splits them into several pieces, for example
after part of the text was corrected or reformatted.

| Parameter | In the template | Effect |
|-----------|-----------------|--------|
| `vars[name]` | `$name$` | Text; `\n` starts a new line. Applies to the body and, unless redefined, to headers and footers. The text keeps the placeholder's formatting. |
| `headers_vars[name]` | `$name$` in a header | Text in headers only |
| `footers_vars[name]` | `$name$` in a footer | Text in footers only |
| `html_vars[name]` | `$name$` alone in a paragraph | The paragraph is replaced by the converted HTML; if the HTML is empty the paragraph is removed |
| `table_vars[name]` | `$name_column$` in a table row | The row is repeated for each item: `[['column' => 'value', …], …]` |
| `checkbox_vars[name]` | checkbox content control tagged `$name$` | Ticked with `1`, `si`, `s`, `yes`, `y`, `true`; unticked with any other value |
| `remove_vars[]` | `$name$` | Removes the paragraph containing the placeholder |

### Supported HTML (`html_vars`)

HTML text starts from the font, size, colour and alignment of the
placeholder in the template.

- **Blocks**: `p`, `div`, `h1`–`h6`, `blockquote`, `pre`, `hr`, `br`,
  `ul`/`ol`/`li` lists, nested too (`type`, `start`, `list-style-type`)
- **Text**: `b`, `strong`, `i`, `em`, `u`, `s`, `del`, `sub`, `sup`,
  `small`, `big`, `mark`, `code`, `font`, `span`, `a href` links
- **Tables**: widths in % or px, `colspan`, `rowspan`, borders,
  backgrounds, `padding`, `cellpadding`, alignment, `<thead>` repeated on
  every page
- **Images**: `<img>` with a data URI (`data:image/png;base64,…`) or an
  http/https URL; PNG and JPEG, GIF/WebP/BMP with the `gd` extension
- **CSS**: `<style>` blocks (type, `.class`, `#id`, descendant and child `>`
  selectors) and the `style` attribute. Properties: `color`,
  `background(-color)`, `font-family`, `font-size`, `font-weight`,
  `font-style`, `text-decoration`, `text-align`, `text-indent`,
  `text-transform`, `line-height`, `letter-spacing`, `margin`, `padding`,
  `border`, `width`, `height`, `vertical-align`, `white-space`,
  `display:none`, `page-break-before`/`page-break-after: always`

Not supported: `float`, absolute positioning, pseudo-classes (`:hover`,
`:first-child`) and attribute selectors.

## API reference

All calls are `POST` (form-urlencoded) and answer in JSON with `esito`
(outcome: `OK`/`KO`) and `errore` (error message). The key may also be sent
in the query string.

### `gdocx.php` — fill a template

| Field | Description |
|-------|-------------|
| `apikey` | API key |
| `nomehost` | Free label for the calling application (only for the log) |
| `modello_docx` | `.docx` template, base64-encoded |
| `vars`, `headers_vars`, `footers_vars`, `html_vars`, `table_vars`, `checkbox_vars`, `remove_vars` | See [Template placeholders](#template-placeholders) |
| `converti_pdf` | `si` (yes) to also get the PDF |

Response: `data` (base64 DOCX), `data_pdf` (base64 PDF, if requested),
`pdf_engine` (`libreoffice` or `php`), `avvisi` (warnings: substituted fonts,
images that could not be loaded, if any).

### `gdocx2pdf.php` — convert a DOCX to PDF

| Field | Description |
|-------|-------------|
| `apikey`, `nomehost` | As above |
| `file_docx` | `.docx` document, base64-encoded |

Response: `data_pdf` (base64 PDF), `pdf_engine`, `avvisi`.

### Common errors

| `errore` | Cause |
|----------|-------|
| `ApiKey errata!` / `ApiKey non valida!` | Key missing or not listed in `_apikeys.php` |
| `Tool non abilitato!` | The key does not include the tool in `tools` |
| `Indirizzo IP non autorizzato: …` | The caller's IP is not in `ip_consentiti` |
| `Server non configurato…` | `_apikeys.php` is missing |
| `Parametri POST non validi!` | Template missing or invalid base64 (often: `post_max_size` too low) |
| `Nessun font TrueType disponibile…` | No fonts found: see [Fonts](#fonts) |

## PDF conversion: LibreOffice or PHP

| | LibreOffice | PHP converter |
|---|---|---|
| Installation | external program on the server | bundled, no requirements |
| Speed | 2–5 s per document | hundredths of a second |
| Fidelity | more complete | very good for text documents, tables, images, headers and footers |
| Fonts | those installed in the system | `lib/fonts/` + system |

The PHP converter does not support, and so skips with a warning: text
wrapping around floating objects, multiple columns, footnotes, charts,
SmartArt, grouped shapes, EMF/WMF images.

With `pdf_engine => 'auto'`, LibreOffice is used when present, otherwise the
PHP converter.

## Security

- `_apikeys.php` is private: it is excluded from git (`.gitignore`) and
  cannot be reached from the web.
- `.htaccess` blocks access to `_*.php` files, `lib/` and `log/`.
  **On nginx** `.htaccess` files are ignored: deny access to those paths in
  the site configuration.
- HTML images given by URL are only fetched over http/https, with a 10 MB
  and 15-second limit.
- Error messages do not reveal server paths.

## Project layout

```
├── gdocx.php              fills a template (DOCX and optionally PDF)
├── gdocx2pdf.php          converts a DOCX to PDF
├── verifica.php           installation check page
├── _start.php             shared bootstrap: keys, IPs, JSON responses
├── _config.php            general settings
├── _apikeys.example.php   template for _apikeys.php (private)
├── .htaccess, .user.ini   file protection and PHP limits
├── index.php              answers 404 at the root
├── lib/
│   ├── GDocx.php          HTML, tables, checkboxes, removals
│   ├── DocTemplate.php    replaces $name$ placeholders
│   ├── DocxPdf.php        pure-PHP DOCX → PDF converter
│   └── fonts/             fonts to download (see LEGGIMI.txt)
└── log/                   request log
```

`DocTemplate.php` and `DocxPdf.php` can also be used on their own, without
the server: usage notes are in the comments at the top of each file.

## Support the project

If you find this project useful, you can buy me a coffee: a simple way to
support its development.

[![Buy Me a Coffee](https://img.shields.io/badge/Buy%20me%20a%20coffee-FFDD00?logo=buymeacoffee&logoColor=000)](https://buymeacoffee.com/fabiozappa)

## Licence

The code is released under the [MIT licence](LICENSE). Fonts are not part of
the project: each one is subject to its own licence (see [Fonts](#fonts)).
