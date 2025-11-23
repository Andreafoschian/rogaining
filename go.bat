@echo off
cls

echo -----------------------------------------------------------------
echo  Avvio lo spettacolare script Python del Fox 
echo  per le classifiche rampe skyrun rogaining
echo -----------------------------------------------------------------
echo.
echo Pre requisito installazione di python sul pc
echo Scarica l'installer da https://www.python.org/downloads/
echo installa le librerie pandas, openpyxl e reportlab con i comandi da shell: 
echo pip install pandas 
echo pip install openpyxl
echo pip install reportlab
echo.
echo rampe.py chiede dove sta l'export xml di orbos (esporta / formato WinSplit / IOF Xml 3.0)
echo Trova in automatico le categorie presenti nel file xml
echo chiede quante rampe creare per ogni categoria
echo chiede i codici di partenza e di arrivo delle varie rampe
echo salva i dati della configurazione in configurazione.cfg
echo.
echo genera un excel con la classifica dei singoli concorrenti divisi per categoria
echo individua le coppie in base al numero di pettorale:
echo Pettorale > 999 => "IND" (Individuale).
echo Pettorale <= 999 => Formattazione a 3 cifre, dove le prime due sono il Codice Staffetta (es. 021 => 02).
echo La colonna Codice_Staffetta sarà inclusa nell'output Excel individuale.

echo .....
echo Crea un excel con la classifica delle coppie individuate con il criterio sopra
echo.
echo.Infine crea 2 file pdf pronti per le stampe sia individuali che di squadra.
pause

python rampe.py

