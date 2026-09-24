# Log di avanzamento — Ricerca prospect B2B Dupuy Malesia

Bootstrap eseguito il: 2026-09-24 15:05 UTC

Coda totale: 91 combinazioni (13 stati/territori × 7 settori Dupuy).
Ogni esecuzione oraria della routine processa una sola voce della coda e aggiunge una riga qui sotto.

| Data/ora (UTC) | Stato | Settore | Aziende trovate | Aziende verificate | Note |
|---|---|---|---|---|---|
| 2026-09-24 15:10 UTC | Johor | Cleaning | 6 (candidate via WebSearch) | 0 | ESECUZIONE BLOCCATA: la policy di rete dell'ambiente cloud impedisce l'accesso in uscita (WebFetch) a qualsiasi dominio esterno (bloccato anche google.com), quindi non è stato possibile eseguire la fase di verifica (fase 3) sui siti candidati. Nessun file Excel prodotto, voce NON segnata come completata, progress.json non avanzato. Candidati trovati via ricerca web (non verificati, da NON considerare attendibili): wlt-my.com, gcsequipment.com.my, ascleaningequipment.com, alliancesuppliesonline.com.my, ftcleaning.com.my, grexsupply.com. |
| 2026-09-24 16:15 UTC | Johor | Cleaning | 6 (stessi candidate via WebSearch) | 0 | ESECUZIONE ANCORA BLOCCATA (2° tentativo consecutivo): WebSearch funziona, ma WebFetch verso qualsiasi dominio esterno (wlt-my.com, gcsequipment.com.my, ftcleaning.com.my, ascleaningequipment.com, alliancesuppliesonline.com.my, grexsupply.com) restituisce EGRESS_BLOCKED dal proxy di rete dell'ambiente. Impossibile eseguire la fase di verifica (fase 3). Nessun file Excel prodotto, voce NON segnata come completata, progress.json non avanzato. Serve intervento umano: ampliare l'accesso di rete dell'ambiente cloud (impostazioni ambiente → Network access) o aggiungere i domini necessari alla allowlist. |
