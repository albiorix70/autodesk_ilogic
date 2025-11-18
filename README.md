# Autodesk Inventor iLogic Modules

Ein paar kleine Code-Schnipsel für Autodesk Inventor iLogic.

## Installation

Entweder man kopiert die Inhalte der Dateien in jeweils eine eigene Regel, um diese dann in den Dokumenten dann aufrufen zu können.

**oder**

Die iLogic-Konfiguration unter Extras öffen

![Screenshot of the Autodesk Inventor Extras tab showing various options including Anwendungsoptionen, Dokumenteinstellungen, Einstellungen migrieren, Autodesk App Manager, Neue markieren, Anpassen, Makros, VBA-Editor, Layout der Benutzeroberfläche zurücksetzen, and iLogic-Konfiguration. The interface is neutral and professional, with a white background and blue and orange icons. The focus is on accessing the iLogic-Konfiguration option under Extras.](docs/pic1.png)

In diesem Optionsdialog das Verzeichnis, in dem sich diese einzelnen Module gefinden, als externes Regel-Verzeichnis anlegen

![Ausicht der Einstellungen in diesem Optionsdialog, unter anderem die Verzeichnisse, die Standard-Erweiterungen und die Protokoll-Ebene](docs/pic2.png)

Zusätzlich sollte die Standard-Erweiterung der RegelDateien von **.iLogicVB** auf **.vb** geändert werden. Diese Änderung veranlasst die iLogic-Engine dazu, den erweiterten .NET-Befehlssatz für diese Dateien zu verwenden. Ohne diese Einstellungen können die Spalten in der Parameter-Tabelle nicht automatisch angelegt werden.