# VALIDATION_REPORT — Controllo qualità cumulativo prospect Dupuy Malesia

**Data del controllo:** 2026-09-26 (esecuzione automatica, routine di verifica quinquennale)

## Esito: nessun dato da verificare

La cartella `output/` non contiene alcun file `.xlsx` (solo il segnaposto `.gitkeep`). La
routine di raccolta (esecuzione oraria) non ha ancora prodotto un solo file di output da
quando è stata avviata il 2026-09-24 15:05 UTC.

- **Aziende raccolte finora:** 0
- **Aziende verificate come genuine:** 0 (0%)
- **Aziende sospette/da scartare:** 0 (0%)
- **Voci sospette rilevate:** nessuna (nessun dato su cui applicare i controlli anti-allucinazione)
- **Combinazioni stato/settore completate:** 0 su 91 (`queue.json`: tutte le voci hanno `done: false`)
- **progress.json:** `index = 0` (nessun avanzamento)

## Causa

Secondo `progress_log.md`, la routine di raccolta ha eseguito **30 tentativi orari
consecutivi** (dal 2026-09-24 15:10 UTC al 2026-09-26 06:10 UTC, ~39 ore) sulla stessa
voce di coda (Johor / Cleaning), sempre bloccata dalla policy di rete dell'ambiente cloud:
`WebFetch` restituisce sistematicamente `EGRESS_BLOCKED` verso qualsiasi dominio esterno
(inclusi domini di test come www.google.com). `WebSearch` funziona e produce candidati
(es. gcsequipment.com.my, ftcleaning.com.my, wlt-my.com, ascleaningequipment.com,
alliancesuppliesonline.com.my, grexsupply.com, ultrahygiene.com.my), ma senza `WebFetch`
la fase di verifica sito/dominio non è mai stata eseguibile, quindi nessuna azienda è mai
stata inserita in un file Excel.

Il log indica che l'utente è già stato notificato più volte (al 4°, 21° e 30° tentativo)
della necessità di ampliare il "Network access" dell'ambiente cloud (menu ambiente → Edit
→ Network access). Questo problema è di configurazione dell'ambiente, non risolvibile da
questa sessione di verifica.

## Voci sospette

Nessuna — non essendoci alcun dato raccolto, non è stato possibile applicare i controlli
di coerenza (nomi aziendali generici/templated, domini riciclati o duplicati tra file,
contenuti del sito non pertinenti al settore dichiarato).

## Raccomandazioni operative

1. **Priorità immediata:** un umano deve ampliare l'accesso di rete in uscita
   dell'ambiente cloud della routine di raccolta (menu ambiente → Edit → Network access),
   consentendo almeno i domini `.com.my` e i domini aziendali generici necessari alla
   verifica dei siti. Finché questo non viene fatto, la routine di raccolta continuerà a
   produrre 0 risultati ogni ora.
2. Una volta sbloccato l'accesso di rete, questa routine di verifica (che gira ogni 5
   giorni) riprenderà il controllo a campione dei file `.xlsx` prodotti man mano che
   vengono generati.
3. Nessuna azione di scarto è necessaria in questo momento: non ci sono righe, file o dati
   da scartare perché non ne esistono ancora.

## Nota metodologica

Questo report non modifica né cancella alcun file in `output/` (attualmente vuota a parte
`.gitkeep`). Le decisioni di scarto di eventuali voci sospette, quando i dati inizieranno
ad arrivare, restano a un operatore umano.
