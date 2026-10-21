# Registrazione micro:bit e Garmin per Lab 6

Lo script salva le informazioni del micro:bit in un vero file Excel `.xlsx`,
una riga ogni **10 secondi**. Dopo la prova scarica dall'API Garmin Connect il
file originale dell'attività e aggiunge i BPM riferiti agli stessi orari.
Non esegue analisi statistiche, medie, interpolazioni, grafici o riempimento
dei dati mancanti.

## Cosa significa sincronizzazione in questa soluzione

I due dispositivi devono misurare **durante la stessa prova**. Il download
Garmin avviene dopo, ma vengono usati gli orari originali dei campioni, non
l'ora del download. L'attività Garmin deve coprire tutta la registrazione.

La Health API cloud ufficiale rende disponibili i dati dopo la sincronizzazione
del dispositivo: non garantisce letture in diretta ogni 10 secondi.
Questa implementazione usa l'API di **Garmin Connect tramite il client Python
non ufficiale `garminconnect`**, per un account personale. Non è un'integrazione
con la Garmin Health API per sviluppatori. Può richiedere aggiornamenti se
Garmin modifica il proprio servizio.

Il gateway fornito non contiene un timestamp della misura del micro:bit:
lo script assegna l'orario UTC di **ricezione sul PC**. Garmin conserva invece
il timestamp del record FIT. L'abbinamento è quindi approssimato, con una
tolleranza dichiarata di **1 secondo**, e non certifica la simultaneità fisica
delle due misure o delle finestre di elaborazione dei sensori.

Per ciascuna riga vengono richiesti un campione micro:bit vicino all'istante
previsto e un campione Garmin vicino sia all'istante previsto sia al timestamp
micro:bit effettivo. Lo scarto rimane visibile in Excel. Se manca un campione
entro la tolleranza, la cella BPM corrispondente resta vuota. Nessun campione
viene riutilizzato per una seconda riga.

Se servono BPM Garmin **in diretta** tramite un'interfaccia Garmin ufficiale,
occorrono accesso al Health Companion SDK per Android/iOS e un'applicazione
ponte. Se è obbligatoria la Health API ufficiale, servono l'accesso sviluppatore
e la relativa specifica: il modulo Garmin presente qui va sostituito.

## Configurazione per il tuo Garmin vívoactive 3 Music

Il modello indicato supporta la registrazione delle attività **Ogni secondo**.
Il flusso già implementato (attività Garmin, sincronizzazione, download API e
abbinamento dei record FIT) non richiede modifiche al codice per questo modello.
Il funzionamento con il tuo orologio resta da verificare durante la prova.

Prima della registrazione:

1. Tieni premuto il touchscreen e seleziona **Impostazioni > Sistema >
   Registrazione dei dati > Ogni secondo** (in inglese: Settings > System >
   Data Recording > Every Second).
2. Sincronizza l'orologio con lo smartphone per aggiornare l'ora. In alternativa,
   usa **Impostazioni > Sistema > Ora > Imposta tempo con GPS** all'aperto.
   Controlla anche la sincronizzazione dell'orologio del PC.
3. Per una prova da fermi, disattiva **Auto Pause**, se presente: premi il tasto,
   seleziona l'attività e apri **Impostazioni > Auto Pause**. È una precauzione
   per evitare che la registrazione si sospenda quando non ti muovi.
4. Avvia manualmente il timer di un'attività con la frequenza cardiaca attiva,
   prima di avviare `logger.py record`. Visualizzare soltanto il widget del
   battito o avviare il cronometro generico non sostituisce questa registrazione.
5. Lascia l'attività in esecuzione per tutta la prova micro:bit. Alla fine,
   arrestala, salvala e sincronizzala con Garmin Connect; esegui poi
   `logger.py merge` con l'ID di quell'attività, come descritto sotto.

La funzione di trasmissione cardio documentata per il vívoactive 3 Music usa
**ANT+**. Un eventuale flusso in diretta verso il PC tramite quella funzione
richiederebbe un ricevitore ANT+ compatibile e un'integrazione dedicata.
La versione presente usa invece il download API dopo la prova; non richiede
un ricevitore ANT+ e non riceve il battito Garmin tramite Bluetooth UART.

Fonti del produttore:
[registrazione Ogni secondo](https://www8.garmin.com/manuals/webhelp/vivoactive3music/IT-IT/GUID-CE4F91AB-1F4F-4AE3-8937-CBB98C4C6442.html),
[orario](https://www8.garmin.com/manuals/webhelp/vivoactive3music/EN-US/GUID-9E140B2A-B378-4C77-86C0-FC3D2F9B2C62.html),
[Auto Pause](https://www8.garmin.com/manuals/webhelp/vivoactive3music/EN-US/GUID-6D3747A8-EB3E-4C32-9C06-C8AC081D0C84.html),
[trasmissione ANT+](https://www8.garmin.com/manuals/webhelp/vivoactive3music/EN-US/GUID-8EC8B7FE-2AC2-46E9-94FC-06416FF1E2ED.html).

## Preparazione

1. Installa Python **3.12 o successivo** e abilita il Bluetooth del PC.
2. Sul Garmin sincronizza prima ora e data tramite telefono/GPS. Sul PC verifica
   che data, ora e sincronizzazione automatica siano corrette **prima** della prova.
3. Sul Garmin scegli un'attività temporizzata che registri la frequenza cardiaca.
   Se il modello lo permette, imposta la registrazione **ogni secondo** e
   disabilita Auto Pause. L'attività può essere registrata anche restando fermi.
   Lo script legge il BPM contenuto nei record: verifica che il Garmin usi il
   sensore da confrontare, ad esempio quello al polso e non una fascia collegata.
4. Il micro:bit deve avere attivo Bluetooth UART e inviare BPM. I formati gestiti,
   ripresi dal gateway fornito, sono `BPM 73` oppure `512 498 73`
   (segnale, segnale filtrato, BPM). I messaggi aggiuntivi restano nel log originale.
5. Chiudi o disconnetti il gateway BLEGW nel browser prima di collegare lo script
   al micro:bit. Un'altra connessione può occupare il dispositivo.

Apri PowerShell nella cartella che contiene questi file:

```powershell
py -3.12 -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

Se usi una versione successiva di Python, sostituisci `py -3.12` con il comando
di quella versione. Non serve installare Microsoft Excel per creare il file.

## 1. Registrare il micro:bit

Avvia l'attività sul Garmin **prima** della registrazione del micro:bit, poi:

```powershell
.\.venv\Scripts\python.exe logger.py record --output misure.xlsx --samples 40
```

Lo script seleziona il dispositivo con nome contenente `BBC micro:bit`, attende
20 secondi di riscaldamento, poi il primo BPM valido. Quel campione definisce
l'inizio della griglia: `t0`, `t0 + 10 s`, `t0 + 20 s`, ecc. Per 40 righe
l'ultimo istante è `t0 + 390 s`. Il salvataggio attende un secondo aggiuntivo
per includere eventuali campioni appena successivi all'istante previsto.

Se ci sono più micro:bit o il nome è differente:

```powershell
.\.venv\Scripts\python.exe logger.py scan
.\.venv\Scripts\python.exe logger.py record --address "AA:BB:CC:DD:EE:FF" --output misure.xlsx --samples 40
```

Sostituisci l'indirizzo con quello realmente mostrato da `scan`. Usa `--pair`
se il firmware richiede l'associazione Bluetooth. Per interrompere premi Ctrl+C:
le righe già raccolte vengono conservate. Una nuova prova richiede un nuovo
nome di file, così non sovrascrive la precedente.

Il formato predefinito `--framing packet` segue il gateway fornito: ogni notifica
deve contenere un messaggio completo, o più righe complete. Se il firmware
termina ogni messaggio con LF (`\n`), usa invece **`--framing line`**, che gestisce
anche messaggi suddivisi fra più notifiche. Se un messaggio viene frammentato
senza un terminatore, va corretto il firmware; il confine non è deducibile
in modo affidabile. Il limite di pacchetto del Bluetooth va rispettato.

`--samples` sceglie il numero di righe; `--warmup` i secondi di riscaldamento.
L'intervallo fra le righe resta 10 secondi. Il micro:bit deve trasmettere almeno
un campione vicino a ogni istante; se la sua periodicità deriva nel tempo,
lo script segnala i campioni mancanti e non ricopia BPM vecchi.

## 2. Aggiungere i BPM dall'API Garmin

Alla fine della registrazione salva l'attività sul Garmin e sincronizzala con
Garmin Connect. Apri quell'attività nel sito Garmin Connect: il numero finale
nell'indirizzo `https://connect.garmin.com/modern/activity/123456789` è l'ID.

```powershell
.\.venv\Scripts\python.exe logger.py merge --session misure.session.json --activity-id 123456789
```

Sostituisci `123456789` con l'ID della tua attività. Al primo accesso vengono
richiesti email, password e, se previsto, codice MFA. Password e codice non
vengono mostrati. La sessione viene salvata nella cartella locale `.garmin_tokens`
per i successivi accessi; questa cartella non va condivisa. Non inserire la
password nei comandi o nei file dello script.

Il file `misure.xlsx` viene aggiornato con i BPM Garmin abbinati e il FIT
originale viene conservato come `garmin_activity_ID.fit`. Non vengono usati
i dati cardiaci quotidiani aggregati, che non sono adatti a una griglia di 10 s.

Un FIT originale già disponibile può essere importato anche senza login:

```powershell
.\.venv\Scripts\python.exe logger.py merge --session misure.session.json --fit attivita.fit
```

`--tolerance 1` indica il massimo scarto accettato in secondi. Può essere
ridotto fino a zero per richiedere una corrispondenza temporale esatta; se i
timestamp non coincidono la coppia resta mancante. Deve essere inferiore a
5 secondi, per evitare finestre sovrapposte. Non basta aumentarlo per rendere
sincroni due orologi errati.

`--garmin-offset` è disponibile solo per una correzione dell'orologio **nota
indipendentemente**: il valore in secondi viene aggiunto all'orario Garmin.
Il valore predefinito è zero; non viene stimato automaticamente dai BPM.
Excel conserva sia l'orario originale sia quello corretto.

## File prodotti e colonne

- `misure.xlsx`: foglio **Misure**, una riga ogni 10 s, e **Impostazioni** con
  fonti, tolleranze, riscaldamento e stato della registrazione.
- `misure.session.json`: copia recuperabile delle righe, salvata a ogni passo.
- `misure.microbit.jsonl`: tutte le notifiche micro:bit originali, con testo,
  byte e orario di ricezione, anche quelle del riscaldamento o non riconosciute.
- `garmin_activity_ID.fit`: attività originale scaricata dall'API.

Il foglio Misure contiene numero della riga, istante previsto UTC, ricezione
micro:bit UTC, BPM micro:bit, segnale e segnale filtrato se disponibili,
timestamp Garmin originale e corretto UTC, BPM Garmin, scarto temporale in
secondi, stato della coppia e messaggio originale micro:bit.

Gli stati distinguono `GARMIN_DA_SCARICARE`, `MICROBIT_MANCANTE`,
`GARMIN_MANCANTE_ENTRO_TOLLERANZA` e `ABBINATO_ENTRO_TOLLERANZA`.
Una riga con stato mancante non è una coppia valida. Le 40 righe predefinite
non garantiscono 40 coppie complete se un sensore perde campioni.

Tieni chiuso Excel durante il salvataggio. Se il file era aperto, i dati
rimangono nel JSON; dopo averlo chiuso puoi ricreare il file:

```powershell
.\.venv\Scripts\python.exe logger.py export --session misure.session.json --output misure.xlsx
```

In caso di disconnessione micro:bit la prova si interrompe e i dati raccolti
vengono salvati. Lo script interrompe inoltre la prova se rileva un salto
dell'orologio PC superiore a 0,5 s. Non può correggere da solo la latenza
Bluetooth, l'orologio del Garmin o le finestre interne dei due sensori.

## Verifica e riferimenti

I controlli automatici verificano parsing UART, messaggi frammentati,
campioni mancanti, scarti e mancato riuso, griglia temporale e file Excel.
La prova con i dispositivi e con il tuo account Garmin richiede la connessione
e le credenziali, che non sono state fornite.

- [Garmin Health API: disponibilità dopo la sincronizzazione](https://developer.garmin.com/gc-developer-program/health-api/)
- [Garmin Health SDK: opzione ufficiale per dati in tempo reale](https://developer.garmin.com/health-sdk/)
- [Client non ufficiale Garmin Connect](https://github.com/cyberjunky/python-garminconnect)
- [SDK FIT ufficiale Garmin](https://github.com/garmin/fit-python-sdk)
- [Bleak: notifiche Bluetooth](https://bleak.readthedocs.io/en/latest/api/client.html)

Le istruzioni Lab 6 sono state usate come riferimento per le 40 letture
predefinite. Le richieste di analisi statistica presenti nel PDF non fanno
parte di questo script, seguendo la richiesta esplicita dell'utente.
