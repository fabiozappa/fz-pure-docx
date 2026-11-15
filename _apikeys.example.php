<?php
/*
 * CHI PUÒ USARE IL SERVER
 *
 * Copiare questo file come  _apikeys.php  (stessa cartella) e modificarlo.
 * _apikeys.php è privato: non va pubblicato né condiviso.
 *
 * ip_consentiti: indirizzi IP dei server autorizzati a chiamare il servizio.
 *   Si possono indicare anche reti intere (es. '198.51.100.0/24').
 *   Se l'elenco è vuoto non si controlla l'IP e basta una chiave valida.
 *
 * chiavi: le chiavi API abilitate. Ogni chiave è di 32 caratteri tra 0-9 e
 *   a-f (minuscoli). Per crearne una:
 *       php -r 'echo bin2hex(random_bytes(16)), "\n";'
 *   'tools' indica gli strumenti consentiti: gdocx, gdocx2pdf.
 */
return [
    'ip_consentiti' => [
        // '203.0.113.10',
    ],

    'chiavi' => [
        '0123456789abcdef0123456789abcdef' => ['tools' => ['gdocx', 'gdocx2pdf']],
    ],
];
