# FZPureDOCX — documenti DOCX e PDF da modelli Word

[English](README.md) · **Italiano**

Piccolo server PHP che compila modelli Word (`.docx`) sostituendo dei
segnaposto con testi, HTML, tabelle e caselle di controllo, e restituisce il
documento in **DOCX** o in **PDF**.

Le applicazioni lo chiamano con una semplice richiesta HTTP POST e ricevono
una risposta JSON.

- Nessuna libreria esterna e nessuna licenza da acquistare.
- PDF con **LibreOffice**, se installato sul server, oppure con un
  **convertitore in puro PHP** incluso, che non richiede programmi esterni.
- Nessun database. Nessun file resta sul server dopo la richiesta.

## Indice

- [Come funziona](#come-funziona)
- [Requisiti](#requisiti)
- [Installazione](#installazione)
- [Font](#font)
- [Configurazione](#configurazione)
- [Verifica dell'installazione](#verifica-dellinstallazione)
- [Chiamare il server](#chiamare-il-server)
- [Segnaposto nei modelli](#segnaposto-nei-modelli)
- [Riferimento API](#riferimento-api)
- [Conversione PDF: LibreOffice o PHP](#conversione-pdf-libreoffice-o-php)
- [Sicurezza](#sicurezza)
- [Struttura del progetto](#struttura-del-progetto)
- [Sostieni il progetto](#sostieni-il-progetto)
- [Licenza](#licenza)

## Come funziona

```
Applicazione                           Server FZPureDOCX
       ── POST (modello in base64 + variabili) ──►  gdocx.php
       ◄──────── JSON (DOCX e/o PDF in base64) ───  │
                                                    ├─ GDocx       compila il modello
                                                    ├─ DocTemplate sostituisce i segnaposto
                                                    └─ DocxPdf     DOCX → PDF in PHP
                                                       (o LibreOffice, se presente)
```

## Requisiti

- PHP **8.1 o successivo**
- Estensioni PHP: `zip`, `dom`, `zlib`, `mbstring` (obbligatorie);
  `curl` (per le immagini HTML indicate con un URL) e `gd` (per GIF/WebP/BMP)
  facoltative
- Apache con `.htaccess` abilitato (su nginx vedi [Sicurezza](#sicurezza))
- Alcuni font TrueType (vedi [Font](#font))
- Facoltativo: LibreOffice (`soffice`) per la conversione PDF

Va bene anche un normale hosting condiviso.

## Installazione

1. Scaricare il progetto (*Code → Download ZIP*) ed estrarlo.
2. Scaricare i font e copiarli in `lib/fonts/` (vedi [Font](#font)).
3. Copiare `_apikeys.example.php` come **`_apikeys.php`** e inserire le
   chiavi (vedi [Configurazione](#configurazione)).
4. Caricare il tutto in una cartella pubblicata dal web, per esempio
   `public_html/fzpuredocx/`. I file nascosti `.htaccess` e `.user.ini`
   vanno caricati anch'essi: nei client FTP di solito occorre attivare la
   visualizzazione dei file nascosti.
5. Rendere scrivibile la cartella `log/` dall'utente del server web (oppure
   disattivare il registro in `_config.php`).
6. Aprire `https://indirizzo/fzpuredocx/verifica.php?apikey=CHIAVE` e
   controllare che sia tutto verde.

## Font

I file dei font **non sono inclusi** nel repository: vanno scaricati dai
siti ufficiali e copiati in `lib/fonts/`.

Il convertitore PHP **incorpora nel PDF** i font usati dal documento e ne
usa le misure per decidere dove andare a capo. Li cerca, in quest'ordine:
nelle cartelle indicate in `_config.php`, in `lib/fonts/` e tra i font di
sistema del server. Se un font del documento manca, ne usa uno simile e lo
segnala negli `avvisi`; in quel caso gli a capo possono essere diversi da
quelli di Word. **Se non trova nessun font, la conversione PDF fallisce.**

### Font di base (consigliati)

Sono liberi e fanno da sostituti per i font più comuni.

| Font | Sostituisce | Dove scaricarlo | File da copiare in `lib/fonts/` |
|------|-------------|-----------------|---------------------------------|
| Liberation Sans e Serif | Arial, Helvetica, Times New Roman | <https://github.com/liberationfonts/liberation-fonts/releases> → `liberation-fonts-ttf-…tar.gz` | `LiberationSans-*.ttf`, `LiberationSerif-*.ttf` |
| DejaVu Sans | Verdana; simboli (☐ ☒ • ◦ ▪ …) | <https://github.com/dejavu-fonts/dejavu-fonts/releases> → `dejavu-fonts-ttf-…zip` (cartella `ttf/`) | `DejaVuSans.ttf`, `DejaVuSans-Bold.ttf` |

In alternativa, su un server Debian/Ubuntu gestito da voi, si possono
installare come font di sistema:

```bash
apt install fonts-liberation fonts-dejavu-core
```

### Font usati nei propri modelli

Per avere un PDF **identico a Word** servono anche i font scelti nei modelli.
Si copiano in `lib/fonts/` i file `.ttf` o `.ttc`; i font `.otf` non sono
supportati.

- **Font liberi**, per esempio Roboto: si scaricano da
  [Google Fonts](https://fonts.google.com/specimen/Roboto) (*Get font →
  Download*). Usare i file della cartella `static/`, come
  `Roboto-Regular.ttf`, `Roboto-Bold.ttf`, `Roboto-Light.ttf`.
- **Font commerciali**, come Calibri, Arial, Verdana e Tahoma di Microsoft:
  si possono copiare da un PC Windows su cui sono installati
  (`C:\Windows\Fonts`), se la licenza ne consente l'uso sul server.

La pagina di [verifica](#verifica-dellinstallazione) elenca i font trovati
e, dopo la prova di conversione, quelli sostituiti.

Anche LibreOffice ha bisogno dei font, installati nel sistema: se mancano li
sostituisce, a volte in modo incoerente (grassetto con le grazie, interlinea
molto larga).

## Configurazione

### `_apikeys.php` — chi può usare il server (privato, escluso da git)

```php
return [
    // IP dei server autorizzati a chiamare il servizio (anche reti: '198.51.100.0/24').
    // Elenco vuoto = nessun controllo sull'IP, basta una chiave valida.
    'ip_consentiti' => [
        '203.0.113.10',
    ],

    // Chiavi API: 32 caratteri 0-9 a-f (minuscoli) e strumenti consentiti
    'chiavi' => [
        '0123456789abcdef0123456789abcdef' => ['tools' => ['gdocx', 'gdocx2pdf']],
    ],
];
```

Per creare una chiave:

```bash
php -r 'echo bin2hex(random_bytes(16)), "\n";'
```

### `_config.php` — impostazioni generali

| Voce           | Valore predefinito | Significato |
|----------------|--------------------|-------------|
| `debug`        | `'no'`             | `'si'` aggiunge file e riga agli errori (solo per prove) |
| `pdf_engine`   | `'auto'`           | `'auto'` LibreOffice se c'è, altrimenti PHP · `'libreoffice'` · `'php'` |
| `soffice`      | `null`             | Percorso di LibreOffice; `null` = ricerca automatica |
| `fonts`        | `[]`               | Cartelle aggiuntive di font `.ttf` per il convertitore PHP |
| `log_file`     | `log/fzpuredocx.log` | Registro delle chiamate; `null` per disattivarlo |
| `memory_limit` | `'1024M'`          | Memoria per l'elaborazione |

### Limiti PHP

I documenti viaggiano in base64 dentro il POST e le tabelle possono avere
molte righe. `.user.ini` (PHP-FPM/CGI) e `.htaccess` (mod_php) alzano
`post_max_size` a 128M e `max_input_vars` a 20000. La pagina di verifica
mostra i valori effettivi.

## Verifica dell'installazione

```
https://indirizzo/fzpuredocx/verifica.php?apikey=LA_TUA_CHIAVE
```

La pagina controlla versione di PHP, estensioni, limiti, cartelle, chiavi e
font, poi **genera davvero un documento di prova** (DOCX compilato con HTML
e tabella, convertito in PDF) e indica il motore usato e il tempo impiegato.
Il pulsante *Apri il PDF di prova* mostra il risultato. Serve una chiave
valida; la pagina non è soggetta al controllo sugli IP, perché si apre dal
proprio browser.

## Chiamare il server

Esempio in PHP con cURL. I campi che sono elenchi (`vars`, `html_vars`, …)
si inviano come array, codificando il corpo con `http_build_query`.

```php
$post = [
    'apikey'       => 'LA_TUA_CHIAVE',
    'nomehost'     => 'mio-sito',                       // solo per il registro
    'modello_docx' => base64_encode(file_get_contents('modello.docx')),
    'vars'         => ['cognome' => 'Rossi', 'note' => "Riga 1\nRiga 2"],
    'html_vars'    => ['corpo' => '<style>td{border:1px solid #000}</style><table><tr><td>…</td></tr></table>'],
    'headers_vars' => ['intestazione' => 'Preventivo n. 12'],
    'table_vars'   => ['prodotti' => [
        ['codice' => 'A1', 'nome' => 'Sedia',     'prezzo' => '120,00'],
        ['codice' => 'B2', 'nome' => 'Scrivania', 'prezzo' => '350,00'],
    ]],
    'converti_pdf' => 'si',                              // anche il PDF
];

$ch = curl_init('https://indirizzo/fzpuredocx/gdocx.php');
curl_setopt_array($ch, [
    CURLOPT_POST           => true,
    CURLOPT_POSTFIELDS     => http_build_query($post),
    CURLOPT_RETURNTRANSFER => true,
    CURLOPT_TIMEOUT        => 120,
]);
$risposta = json_decode((string) curl_exec($ch), true);

if (($risposta['esito'] ?? '') === 'OK') {
    file_put_contents('documento.docx', base64_decode($risposta['data']));
    file_put_contents('documento.pdf', base64_decode($risposta['data_pdf']));
} else {
    echo 'Errore: ', $risposta['errore'] ?? 'risposta non valida';
}
```

## Segnaposto nei modelli

Nel modello Word i segnaposto si scrivono **`$nome$`**. Funzionano anche se
Word li spezza internamente in più parti, per esempio quando una parte è stata
corretta o formattata dopo.

| Parametro | Nel modello | Effetto |
|-----------|-------------|---------|
| `vars[nome]` | `$nome$` | Testo; `\n` va a capo. Vale nel corpo e, se non ridefinito, in intestazioni e piè di pagina. Il testo prende la formattazione del segnaposto. |
| `headers_vars[nome]` | `$nome$` in un'intestazione | Testo solo nelle intestazioni |
| `footers_vars[nome]` | `$nome$` in un piè di pagina | Testo solo nei piè di pagina |
| `html_vars[nome]` | `$nome$` da solo in un paragrafo | Il paragrafo viene sostituito dall'HTML convertito; se l'HTML è vuoto il paragrafo viene tolto |
| `table_vars[nome]` | `$nome_colonna$` in una riga di tabella | La riga si ripete per ogni elemento: `[['colonna' => 'valore', …], …]` |
| `checkbox_vars[nome]` | casella di controllo con tag `$nome$` | Spuntata con `1`, `si`, `s`, `yes`, `y`, `true`; vuota con gli altri valori |
| `remove_vars[]` | `$nome$` | Toglie il paragrafo che contiene il segnaposto |

### HTML supportato (`html_vars`)

Il testo HTML parte dal font, dalla dimensione, dal colore e
dall'allineamento del segnaposto nel modello.

- **Blocchi**: `p`, `div`, `h1`–`h6`, `blockquote`, `pre`, `hr`, `br`,
  elenchi `ul`/`ol`/`li` anche annidati (`type`, `start`, `list-style-type`)
- **Testo**: `b`, `strong`, `i`, `em`, `u`, `s`, `del`, `sub`, `sup`,
  `small`, `big`, `mark`, `code`, `font`, `span`, link `a href`
- **Tabelle**: larghezze in % o px, `colspan`, `rowspan`, bordi, sfondi,
  `padding`, `cellpadding`, allineamenti, `<thead>` ripetuto a ogni pagina
- **Immagini**: `<img>` con data URI (`data:image/png;base64,…`) o URL
  http/https; PNG e JPEG, GIF/WebP/BMP con l'estensione `gd`
- **CSS**: blocchi `<style>` (selettori di tipo, `.classe`, `#id`,
  discendente e figlio `>`) e attributo `style`. Proprietà: `color`,
  `background(-color)`, `font-family`, `font-size`, `font-weight`,
  `font-style`, `text-decoration`, `text-align`, `text-indent`,
  `text-transform`, `line-height`, `letter-spacing`, `margin`, `padding`,
  `border`, `width`, `height`, `vertical-align`, `white-space`, `display:none`,
  `page-break-before`/`page-break-after: always`

Non supportati: `float`, posizionamento assoluto, pseudo-classi (`:hover`,
`:first-child`) e selettori di attributo.

## Riferimento API

Tutte le chiamate sono `POST` (form-urlencoded) e rispondono in JSON con
`esito` (`OK`/`KO`) ed `errore`. La chiave può arrivare anche in GET.

### `gdocx.php` — compila un modello

| Campo | Descrizione |
|-------|-------------|
| `apikey` | Chiave API |
| `nomehost` | Nome libero dell'applicazione chiamante (solo per il registro) |
| `modello_docx` | Modello `.docx` in base64 |
| `vars`, `headers_vars`, `footers_vars`, `html_vars`, `table_vars`, `checkbox_vars`, `remove_vars` | Vedi [Segnaposto](#segnaposto-nei-modelli) |
| `converti_pdf` | `si` per avere anche il PDF |

Risposta: `data` (DOCX in base64), `data_pdf` (PDF in base64, se richiesto),
`pdf_engine` (`libreoffice` o `php`), `avvisi` (font sostituiti, immagini non
caricate, se presenti).

### `gdocx2pdf.php` — converte un DOCX in PDF

| Campo | Descrizione |
|-------|-------------|
| `apikey`, `nomehost` | Come sopra |
| `file_docx` | Documento `.docx` in base64 |

Risposta: `data_pdf` (PDF in base64), `pdf_engine`, `avvisi`.

### Errori comuni

| `errore` | Causa |
|----------|-------|
| `ApiKey errata!` / `ApiKey non valida!` | Chiave mancante o non presente in `_apikeys.php` |
| `Tool non abilitato!` | La chiave non ha lo strumento nella voce `tools` |
| `Indirizzo IP non autorizzato: …` | L'IP del chiamante non è in `ip_consentiti` |
| `Server non configurato…` | Manca `_apikeys.php` |
| `Parametri POST non validi!` | Modello assente o base64 non valido (spesso: `post_max_size` troppo basso) |
| `Nessun font TrueType disponibile…` | Mancano i font: vedi [Font](#font) |

## Conversione PDF: LibreOffice o PHP

| | LibreOffice | Convertitore PHP |
|---|---|---|
| Installazione | programma esterno sul server | incluso, nessun requisito |
| Velocità | 2–5 s a documento | centesimi di secondo |
| Fedeltà | più completa | ottima per documenti di testo, tabelle, immagini, intestazioni e piè di pagina |
| Font | quelli installati nel sistema | `lib/fonts/` + sistema |

Il convertitore PHP non supporta, e quindi ignora segnalandolo: testo che
scorre attorno a oggetti flottanti, colonne multiple, note a piè di pagina,
grafici, SmartArt, forme raggruppate, immagini EMF/WMF.

Con `pdf_engine => 'auto'` si usa LibreOffice quando è presente, altrimenti
il convertitore PHP.

## Sicurezza

- `_apikeys.php` è privato: è escluso da git (`.gitignore`) e non è
  raggiungibile dal web.
- Il `.htaccess` blocca l'accesso ai file `_*.php`, a `lib/` e a `log/`.
  **Su nginx** i `.htaccess` non valgono: negare l'accesso a quei percorsi
  nella configurazione del sito.
- Le immagini HTML da URL vengono scaricate solo via http/https, con un
  limite di 10 MB e 15 secondi.
- I messaggi d'errore non riportano i percorsi del server.

## Struttura del progetto

```
├── gdocx.php              compila un modello (DOCX ed eventualmente PDF)
├── gdocx2pdf.php          converte un DOCX in PDF
├── verifica.php           pagina di controllo dell'installazione
├── _start.php             avvio comune: chiavi, IP, risposte JSON
├── _config.php            impostazioni generali
├── _apikeys.example.php   modello per _apikeys.php (privato)
├── .htaccess, .user.ini   protezione dei file e limiti PHP
├── index.php              risponde 404 alla radice
├── lib/
│   ├── GDocx.php          HTML, tabelle, checkbox, rimozioni
│   ├── DocTemplate.php    sostituzione dei segnaposto $nome$
│   ├── DocxPdf.php        convertitore DOCX → PDF in puro PHP
│   └── fonts/             font da scaricare (vedi LEGGIMI.txt)
└── log/                   registro delle chiamate
```

`DocTemplate.php` e `DocxPdf.php` si possono usare anche da soli, senza
server: le istruzioni sono nei commenti all'inizio di ciascun file.

## Sostieni il progetto

Se questo progetto ti è utile, puoi offrirmi un caffè: è un modo semplice
per sostenerne lo sviluppo.

[![Offrimi un caffè](https://img.shields.io/badge/Buy%20me%20a%20coffee-FFDD00?logo=buymeacoffee&logoColor=000)](https://buymeacoffee.com/fabiozappa)

## Licenza

Il codice è distribuito con licenza [MIT](LICENSE). I font non fanno parte
del progetto: ciascuno resta soggetto alla propria licenza (vedi [Font](#font)).
