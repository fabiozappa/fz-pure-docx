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
// Questo file resta leggibile anche da PHP vecchi: il controllo vero è in _verifica.php
$tool = 'verifica';
require __DIR__ . '/_versione.php';
require __DIR__ . '/_verifica.php';
