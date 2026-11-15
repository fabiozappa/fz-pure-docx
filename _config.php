<?php
/*
 * Configurazione generale del server FZPureDOCX.
 */
return [
    // 'si' mostra i dettagli degli errori PHP nella risposta JSON (solo per prove)
    'debug' => 'no',

    // Motore per la conversione PDF:
    //   'auto'        LibreOffice se installato, altrimenti il convertitore in puro PHP
    //   'libreoffice' solo LibreOffice (errore se manca)
    //   'php'         solo il convertitore in puro PHP (DocxPdf.php)
    'pdf_engine' => 'auto',

    // Percorso di LibreOffice; null = ricerca automatica (soffice nel PATH, /usr/bin/soffice, …)
    'soffice' => null,

    // Cartelle aggiuntive con font .ttf per il convertitore PHP (oltre a lib/fonts)
    'fonts' => [],

    // Registro delle chiamate (una riga per richiesta); null per disattivarlo
    'log_file' => __DIR__ . '/log/fzpuredocx.log',

    // Memoria per gli strumenti che elaborano documenti
    'memory_limit' => '1024M',
];
