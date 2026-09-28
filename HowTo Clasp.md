### **Beim Erstellen einer neuen Datei**

:: 1. In Ihren Wunschordner wechseln (Anführungszeichen sind wegen des Leerzeichens 'py Files' nötig)

cd "C:\\Users\\andreas.steinert\\Documents\\Python\\py Files\\Google\_Scripts"



:: 2. Bei Google anmelden (öffnet das Browser-Anmeldefenster)

clasp login



:: 3. Skripte herunterladen (Skript-ID aus Schritt 1 einfügen)

clasp clone "DEINE\_SKRIPT\_ID"









1\. Wie finde ich die richtige Skript-ID?⚠️ Wichtig: Verwenden Sie nicht die ID aus der Webadresse Ihrer Google-Tabelle (117gwY1sK...). Das ist die Spreadsheet-ID.



Sie benötigen die Skript-ID des dazugehörigen Code-Projekts.Öffnen Sie Ihre Google-Tabelle im Browser.Klicken Sie oben im Menü auf Erweiterungen $\\rightarrow$ Apps Script.Sie finden die Skript-ID auf zwei Wegen:Weg A (Projekteinstellungen): Klicken Sie in der linken Seitenleiste auf das Zahnrad-Symbol (⚙️ Projekteinstellungen). Kopieren Sie den Code unter Skript-ID.Weg B (Browser-URL): Schauen Sie oben in die Adresszeile des Browsers. Die ID steht direkt zwischen /projects/ und /edit:\[https://script.google.com/u/0/home/projects/](https://script.google.com/u/0/home/projects/)DEINE\_SKRIPT\_ID/edit



### Beim Update einer Datei

**In Eingabeauffoderung**



clasp pull

