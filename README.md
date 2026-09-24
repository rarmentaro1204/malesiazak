# malesiazak — Ricerca prospect B2B Dupuy Malesia

Repository di supporto per una routine automatica (agente cloud, esecuzione oraria) che
effettua ricerca prospect B2B per il brand industriale **Dupuy** nei 13 stati/territori
della Malesia, nei 7 settori di prodotto Dupuy:

- Cleaning
- Utensileria
- Depolverazione
- ATEX
- Sabbiatura
- Welding
- Meccanica

## Scopo

Per ciascuna combinazione stato × settore (91 combinazioni totali), la routine cerca sul
web potenziali aziende distributrici/rivenditrici/rappresentanti che potrebbero essere
interessate a distribuire prodotti industriali Dupuy in quello stato/settore, verifica che
i domini/siti trovati esistano davvero e siano pertinenti, e salva i risultati grezzi in un
foglio Excel.

## Struttura del repository

- `queue.json` — coda delle 91 combinazioni stato/settore da processare, con flag `done`.
- `progress.json` — segnalibro con l'indice della prossima voce di `queue.json` da processare.
- `progress_log.md` — log testuale di ogni esecuzione (data/ora, voce processata, esito).
- `output/` — file Excel `.xlsx` prodotti, uno per ogni combinazione stato/settore
  effettivamente processata, con nome
  `Distributori_Dupuy_Malesia_<Stato>_<Settore>_<YYYY-MM-DD>.xlsx`.

## ⚠️ Avvertenza importante: dati grezzi, da verificare umanamente

**Tutti i dati contenuti nei file in `output/` sono grezzi (raw) e prodotti da un agente
automatico.** Anche dopo il controllo automatico di esistenza/pertinenza del dominio, queste
informazioni:

- possono contenere errori, aziende non più attive, o dati di contatto non aggiornati;
- NON sono mai state validate da un operatore umano;
- NON devono essere usate per contatti commerciali, invii massivi o decisioni di business
  senza prima una verifica manuale puntuale (controllo del sito, del settore reale
  dell'azienda, e dei recapiti).

Ogni riga nei file Excel riporta l'etichetta **"DA VERIFICARE UMANAMENTE"** proprio per
ricordare che nessun dato è da considerarsi definitivo.

## Come funziona la routine

Ad ogni esecuzione (circa una volta all'ora):

1. legge `progress.json` per sapere quale voce di `queue.json` processare;
2. cerca sul web aziende pertinenti per quello stato/settore, senza mai inventare nomi,
   domini o contatti (se non trova nulla di verificabile, lo scrive esplicitamente nel log);
3. verifica che i domini trovati esistano e siano coerenti con l'azienda indicata, scartando
   le voci sospette o non verificabili;
4. salva un file Excel in `output/` con le aziende verificate per quella combinazione;
5. aggiunge una riga a `progress_log.md`;
6. aggiorna `progress.json` passando alla voce successiva (tornando a 0 dopo la 91esima,
   segnando un giro completo);
7. effettua commit e push delle modifiche.
