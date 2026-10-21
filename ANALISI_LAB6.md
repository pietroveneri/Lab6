# Analisi Lab 6

Apri **risultati_analisi/report_analisi_lab6.html** per leggere risultati, grafici,
differenze individuali, letture escluse e confronto con il piano sperimentale.
Il file contiene le immagini e funziona anche senza connessione.

Per rigenerare tutto dalla cartella del progetto:

```powershell
powershell -ExecutionPolicy Bypass -File .\Avvia_analisi.ps1
```

Oppure esegui `analisi_lab6.py` dal tuo interprete Python. Su un altro computer,
installa prima `requirements_analisi.txt` nello stesso interprete.
Lo script accetta anche `--input-dir CARTELLA` e `--output-dir CARTELLA`.
Le librerie aggiunte per questa esecuzione si trovano in `.analysis_packages`.
L'avvio PowerShell prova il venv del progetto e poi il runtime locale di Codex:
il venv attuale punta a un'installazione Python 3.12 non più presente.

Si usano esclusivamente:

- `misure_artificial_30.xlsx`, gruppo **Artificiali 30 bpm**;
- `misure_artificial_150.xlsx`, gruppo **Artificiali 150 bpm**;
- `misure_pietro.xlsx`, gruppo **Pietro**;
- `misure_manfred.xlsx`, gruppo **Manfred**.

Gli artificiali a 30 e 150 bpm hanno statistiche, sezioni e grafici separati.
Le differenze sono PPG meno riferimento nominale. Il riferimento
dei due file artificiali viene dal nome: la sua verifica indipendente non è
documentata. I risultati umani descrivono l'accordo con Garmin.

Il report calcola bias, SD e varianza campionaria, MAE, RMSE, quota entro ±5 bpm,
bande descrittive bias ±1,96 SD e percentili empirici. Include grafici temporali,
Bland-Altman, istogrammi e Q-Q plot. Non elimina letture con errori grandi.
Le bande non hanno una copertura del 95% dimostrata: controllare le cautele
sulla distribuzione e sulla varianza riportate nel report.

L'IC del bias previsto dal PDF resta non disponibile: mancano prove ripetute
indipendenti della stessa dimensione, con identificazione del PCB. Non si
creano repliche dividendo una serie temporale né usando i due livelli artificiali.

Gli Excel originali restano intatti. Nella cartella `risultati_analisi`:

- `riepilogo_categorie.csv`: le quattro categorie separate;
- `dettaglio_artificiali.csv`: le due categorie artificiali separate;
- `dati_analizzati.csv`: tutte le righe, differenze e motivi di esclusione;
- `risultati.json`: statistiche con precisione completa;
- `manifest.json`: fonti, hash SHA-256, impostazioni e versioni;
- `artificiali_30_bpm/`, `artificiali_150_bpm/`, `pietro/`, `manfred/`: una cartella
  per esperimento, con quattro grafici separati e testi in inglese.

Ogni cartella dell'esperimento contiene `time_series`, `bland_altman`,
`histogram` e `qq_plot`, ciascuno in PNG e PDF vettoriale. Il report li mostra
separatamente e include i collegamenti ai singoli file.

I CSV usano `;` come separatore e il punto decimale; possono essere importati
in Excel specificando questi separatori.
