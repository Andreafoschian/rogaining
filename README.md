# rogaining
Come estrerre da oribos i dati per generare classifiche di cornoscalata per una gara di rogaining

Lo script serve per generare la classifica delle cronoscalata per una gara di orienteering / rogaining a sequenza obbligata a coppie.
Impostare un file di oribos come se fosse una staffetta con i frazionisti che hanno lo stessto minuto di partenza. Partono assieme. 
Assegnare i pettorali ai concorrenti con il seguente criterio: Pettorale > 999 => "IND" (Individuale).
Pettorale <= 999 => Formattazione a 3 cifre, dove le prime due sono il Codice Staffetta (es. 021 => 02).
ES: Staffetta 30 avrà il primo concorrente con il pettorale 31 ed il secondo con il 32

Creare da oribos il file export.xml  (esporta / formato WinSplit / IOF Xml 3.0)
Pre requisito installazione di python sul pc -   Scarica l'installer da https://www.python.org/downloads/ ed iserisci python nel path
installa le librerie pandas, openpyxl e reportlab con i comandi da shell: 
pip install pandas 
pip install openpyxl
pip install reportlab

rampe.py chiede dove sta l'export xml di orbos
Trova in automatico le categorie presenti nel file xml
chiede quante rampe creare per ogni categoria
chiede i codici di partenza e di arrivo delle varie rampe
salva i dati della configurazione in configurazione.cfg
genera un excel con la classifica dei singoli concorrenti divisi per categoria
individua le coppie in base al numero di pettorale:
La colonna Codice_Staffetta sarà inclusa nell'output Excel individuale.
Crea un excel con la classifica delle coppie individuate con il criterio di cui  sopra
Infine crea 2 file pdf pronti per le stampe sia individuali che di squadra.
