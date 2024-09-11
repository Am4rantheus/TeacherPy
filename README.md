# TeacherPy

TeacherPy ist ein pythonbasiertes Werkzeug zur Erstellung und Verwaltung von Stundenverlaufsplänen, der automatisierten PDF-Erstellung von Arbeitsblättern mit und ohne Erwartungsbildern und vielfältigen Möglichkeiten zur einfachen Dateiverwaltung für Lehrerinnen und Lehrer!

## Installation

1. Python im Windows Store installieren (empfohlen wird Python 3.12) ![python installieren](https://github.com/user-attachments/assets/26583b5c-436b-466e-b2b9-20afd606d339)
2. Den Ordner '/TeacherPy' öffnen
3. Im Verzeichnis `setup.py` ausführen. ![setup py](https://github.com/user-attachments/assets/4258b843-856f-4afe-8515-eaecf0e558ea)

4. setup.py installiert alle benötigten Pakete und fragt nach Pfaden zu Basisordner, USB-Ordner und ob ein Notiz-Ordner gewünscht ist, außerdem nach den gewünschten Optionen für die zukünftig zu planenden Stunden. ![Fragen beim ersten Ausführen](https://github.com/user-attachments/assets/47b52f5f-3097-4198-a88a-6083dda3bf03)
   > **Hinweis:** TeacherPy ist in der Lage, mit Clouddatein umzugehen, sofern Sie diese als virtuelles Laufwerk eingebunden haben. Verwenden Sie hierfür einfach den entsprechenden Ordnerpfad z.B. C:\Users\Name\OneDrive. Der USB-Ordner sollte zu einem USB-Laufwerk führen, welches für die reibungslose Nutzung mit dem PC verbunden sein muss, wenn Sie TeacherPy nutzen.
   
5. Ist die Installation erfolgreich, wird eine Verknüpfung "TeacherPy" im Verzeichnis von TeacherPy erstellt, die auf den Desktop verschoben werden kann. ![Verknüpfung](https://github.com/user-attachments/assets/bd0a9693-1d53-4885-b9d7-b07acc986185)

6. Beim Doppelklick auf die Verknüpfung öffnet sich das Programm.

## TeacherPy - Funktionen

TeacherPy bietet drei Optionen:
1. Das Erstellen einer neuen Stunde im Basisordner.
2. Das Finalisieren einer Stunde aus dem Basisordner. 
3. Das Archivieren von gehaltenen Stunden.

### Eine neue Stunde erstellen

1. Starten Sie TeacherPy über die Verknüpfung "TeacherPy".
2. Wählen Sie die Option "1. Eine neue Stunde erstellen"
3. Vergeben Sie einen Namen für die neu zu erstellende Stundeneinheit ![Bsp neue Stunde 1](https://github.com/user-attachments/assets/d5747bb3-57a0-40a5-ad12-c6641650fe1a)

4. Geben Sie den gewünschten Zeitslot, Klasse, Raumnummer, Datum, den Lernbereich und das Stundenthema ein. Der Name der Lehrperson ist in der Regel Ihr eigener und muss nicht noch einmal angepasst werden. ![Bsp neue Stunde mit Optionen](https://github.com/user-attachments/assets/47aea1c0-d777-45c7-9d86-5627563061f3)

   
   TeacherPy erstellt im Basisordner einen Unterordner mit dem Namen der Stundeneinheit, darin einen weiteren Unterordner "Ressources" für Bilder, Dokumente etc, die Sie später nicht mehr benötigen und einen Stundenverlaufsplan (SVP_Name der Stunde.docx) ![SVP mit ausgefüllten Tabellenzellen](https://github.com/user-attachments/assets/3ada0614-442b-4226-b365-143bd76e506d)

   
5. In der geöffneten SVP_Datei können Sie nun wie gewohnt Ihren Unterricht planen. Achten Sie darauf, dass Sie eventuell erstellte Arbeitsblätter nach dem Schema AB_Name_des_Blattes im gleichen Verzeichnis speichern, in dem sich die SVP_Datei befindet. ![SVP mit ABs](https://github.com/user-attachments/assets/5fbc83ce-d66a-4f85-980d-2b3b0af8b1f1)


   > **Hinweis:** Für eine optimale Verwendung von TeacherPy wird empfohlen, die Namen der Arbeitsblätter 1zu1 in die Tabellenspalte Materialien/Medien einzutragen. Das Programm kann diese Informationen auslesen und später verwenden. Dazu ist es unerlässlich, die Dateien mit dem Kürzel AB_ zu versehen. 

### Eine Stunde finalisieren

1. Wählen Sie die Option "2. Eine Stunde finalisieren"
2. Wählen Sie aus den verfügbaren Unterordnern Ihres Basisordners denjenigen aus, den Sie finalisieren möchten ![finalisieren](https://github.com/user-attachments/assets/cd2c1ac5-e488-436f-9386-a83908815866)

3. Das Programm sucht den Ordner nun nach sämtlichen Office-Dateien ab, welche Arbeitsblätter (AB_), Leistungsbewertung (LB_) sowie Stundenverlaufsplan (SVP_) im Dateinamen enthalten
4. Sie werden einzeln gefragt, ob Sie ein Erwartungsbild mit Kommentaren erstellen lassen wollen. Diese Funktion erstellt eine zusätzliche PDF mit sämtlichen Kommentaren, die als Erwartungsbilder erstellt worden sind.  ![AB mit Erwartungsbild](https://github.com/user-attachments/assets/a0bd674e-d866-4d5f-b251-3e6014f3cfbf)
![Finalisieren Erwartungsbild](https://github.com/user-attachments/assets/1a989b12-9432-4b4f-a4ff-417dd38998fc)

5. Lehnen Sie diese Option ab, wird lediglich die Ausgangsdatei ohne Kommentare in eine PDF umgewandelt. 
6. Nachdem alle ABs sowie LBs im Ordner konvertiert wurden, wird die SVP_Datei ebenfalls konvertiert. 
7. Sie werden gefragt, ob Sie die konvertierten Dateien zusammenführen möchten. Dieser Schritt bietet sich vor allem an, wenn Sie mit einem E-Ink-Tablet arbeiten und eine Gesamtdatei samt Erwartungsbildern und ABs benötigen. Die Dateien werden hierbei in der Reihenfolge angehängt, in der sie in der Tabelle genannt werden 
8. INBOX Funktion ist in der Testphase, bitte lehnen Sie diese ab ![Inbox ablehnen (testphase)](https://github.com/user-attachments/assets/8969a055-7605-4220-bdf9-676a57e14d1b)

9. Sie werden gefragt, ob Sie die Dateien direkt auf den USB-Stick kopieren möchten. Wählen Sie die gewünschte Option. ![kopieren auf USB](https://github.com/user-attachments/assets/08a87a75-347a-41d8-b1b8-bba1e7d02a08)

10. Ihre Dateien werden auf den USB-Stick kopiert.

### Eine Stunde archivieren

1. Wählen Sie die Option "3. Eine abgeschlossene Stunde archivieren" 
2. TeacherPy durchsucht Ihren USB-Ordner nach Unterordnern, wählen Sie den gewünschten Ordner aus ![Archiv 1](https://github.com/user-attachments/assets/1c1cf6c0-fdcc-4bfa-bb99-f5110b1a0e49)
3. Die im Ordner befindlichen Dateien werden aufgelistet ![Archiv 2](https://github.com/user-attachments/assets/d73f312f-5a6c-4666-85fe-2fef9eb5a164)

4. Sie haben nun die Möglichkeiten:
   a) die auf dem USB-Stick befindlichen Dateien zurück in den Basisordner zu verschieben (dabei werden die vorhandenen Dateien überschrieben)
   b) nur neu erstellte Dateien auf dem Stick in den Basisordner zu verschieben
   c) den gesamten Ordner (Bilder, Präsentationen, Videos, Dokumente, Unterordner) in den Basisordner zu verschieben

5. Sie werden gefragt, ob Sie die Dateien auf dem Stick behalten möchten oder nicht. Verneinen Sie dies, werden die Dateien gelöscht! ![Archiv gelöscht](https://github.com/user-attachments/assets/895530f5-5b6f-4206-9959-8cd40222f636)


## Allgemeine Nutzungshinweise

Abkürzungen, mit denen TeacherPy aktuell arbeitet sind:
- SVP_ für Stundenverlaufspläne
- AB_ für Arbeitsblätter
- LB_ für Leistungsbewertungen

In der project-config.json lassen sich weitere Funktionen aktivieren und deaktivieren.

## Fehlerbehebung

TeacherPy ist in der Entwicklung, es kann nicht für einen fehlerfreien Ablauf garantiert werden. Sollten Sie Fehler vorfinden, informieren Sie mich bitte und ich versuche, die Fehler zu beheben. 
Sollten Sie falsche oder irrtümliche Eingaben durchführen und das Skript deswegen Fehlermeldungen ausgeben, wiederholen Sie den Schritt, welchen Sie durchführen wollten und prüfen Sie, ob dies den Fehler bereits behebt.

Der Ersteller dieses Programms übernimmt keinerlei Verantwortung für den Verlust von sensiblen Daten. Bitte gehen Sie sicher, dass Sie Ihre Daten jederzeit gesichert haben, bevor Sie sie verschieben, löschen oder verändern, um einen Datenverlust zu vermeiden. 

## Lizenz

©Am4ranth/Am4rantheus

Dieses Tool ist zur freien Nutzung angedacht. Jede/r Benutzer/-in hat das Recht, das Programm beliebig oft zu vervielfältigen und zu verändern.
