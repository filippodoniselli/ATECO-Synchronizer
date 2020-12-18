# ATECO Synchronizer
## Abstract
Questo esercizio richiede la creazione di un'applicazione console senza parametri in ingresso. 

**Sarà richiesto l'uso specifico di alcune librerie software.** 

**Sarà richiesta la creazione di un progetto Unit Test** che verifichi il corretto svolgimento dell'esercizio.

**Sarà richiesto l'utilizzo di Git per il versioning del codice e di GitKraken Board per il tracking delle attività svolte.** 

Al fine dello svolgimento dell'esercizio sono forniti:

1. File di input
2. Soluzione VS suddivisa nei due progetti da svolgere
3. Board GitKraken
4. Repository Git
## Spiegazione dei file di input
All'interno del progetto console è presente una cartella XLS. In essa sono presenti file Excel nei due formati classici standard (_.xls_/_.xlsx_). Del file *Cod_Att.xls* sono di vostro interesse le colonne **A** e **C**.
Dei file presenti nella sottocartella **_IRP_** sono di interesse le colonne **A** e **B**. Dei file nella sottocartella **_AIRAP_** sono di interesse le colonne **A** e **C**.
### Considerazioni utili
1. **La prima riga di ogni file è da considerarsi d'intestazione.**
2. **I file delle cartelle AIRAP e IRP hanno una correlazione desumibile dal nome. Questa correlazione deve essere gestita in maniera dinamica, tramite file di configurazione.**
3. **Il file _Cod_Att.xls_ è necessario per svolgere adeguatamente l'analisi correlata.**
## Risultato atteso
1. I file di input devono essere importati nel progetto console (pertanto visibili da VS) mantenendo la struttura in cartelle e sottocartelle.
2. L'applicativo deve eseguire l'analisi dei file di input in automatico, senza necessità di alcuna operazione dell'utente.
3. L'applicativo deve generare una sottocartella **_Result_** contenente i file generati dopo la sua analisi. La sottocartella deve trovarsi nella stessa directory dell'applicativo.
   * All'interno della sottocartella  **_Result_** deve esser presente un file per ogni file presente in sottocartella **_IRP_**, con pari nome.
   * Ogni file prodotto deve contenere, correttamente suddivisi in colonne, i seguenti dati presi dai vari file: **Colonna A file IRP**, **_Colonna C file Cod_Att.xls_**, **Colonna B file IRP**, **_Colonna C file AIRAP_**.
   * Le celle, dato il valore delle corrispettive colonne, dovranno essere del seguente tipo: **stringa**, **_stringa_**, **stringa**, **_numerico_**.
   * Dovrà essere presente intestazione come nei file originali.
4. Dovrà essere presente una sottocartella **_logs_** nella stessa directory dell'applicativo. Al suo interno dovrà trovarsi un file di log che tracci le attività svolte. Il file di log non dovrà mai superare dimensione di 1024 Kb ma dovrà essere mantenuto uno storico fino ad un massimo di 10 file precedenti.
   * Il file di log dovrà essere chiaro, conciso, ben strutturato. Dovranno essere presenti tutte le informazioni necessarie a comprendere la storia di un ciclo applicativo.
5. **Fornire in input uno dei file di output non dovrà essere causa di eccezione, l'applicativo dovrà poter usare i file _IRP_ che genera come file di input, eseguendo gli stessi controlli e rigenerando il file in output.**

## Librerie
    Le librerie DEVONO esser aggiunte al progetto tramite NuGet.
Ci si attende che usiate le seguenti librerie obbligatorie e facoltative:
### Obbligatorie
* **NPOI** per la gestione dei file di input e output
* **Log4Net** per la gestione del file di log
### Facoltative
* **System.Linq** perché vista più volte in passato e fondamentale in C#

## Ottimizzazione
Non è richiesta ottimizzazione in termini di tempo speso e qualità del codice.

## Test
Per il testing del progetto è considerato fondamentale l'utilizzo di un progetto Unit Test. All'interno della soluzione è già presente. In un progetto di test ben strutturato, teoricamente ogni metodo scritto dovrebbe essere testato. A voi stabilire come meglio creare un test che validi il vostro applicativo.

## Documentazione
    Assioma:    "Un buon progetto è tale se e solo se è ben documentato."
    Corollario: "Nessun progetto è buono se non è documentato."
GitLab ha a disposizione una sezione Wiki per ogni progetto. 
A voi popolare in maniera adeguata la vostra.
**La documentazione delle configurazioni non è opzionale.**
Sarebbe bene avere una visione chiara e precisa dell'architettura in generale e di ogni singolo metodo in particolare. Avete visto come sono documentate le API dei framework Microsoft e vedrete come sono documentate **Log4Net** e **NPOI**. Sfruttate questi esempi. 

# Tempi e Qualità Attesi
## Professionali
1. **Tempo:** 4.5 ore
2. **Qualità:** uso di tutte le librerie di cui sopra, non oltre 150 righe.
## Richiesti
1. **Tempo:** entro il 10/01/2021
2. **Qualità:** uso delle sole librerie obbligatorie, nessun limite.