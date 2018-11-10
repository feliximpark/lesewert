# Lesewert Python-Modul #

Lesewert.py ist ein Modul für die automatisierte Auswertung von Lesewert-Messungen. Das Modul übernimmt Dataframes, berechnet die Auswertungen und fertigt Powerpoint-Folien an.
---

## Kernfunktionen

### Umfrage-Daten auswerten

#### gestapelte Balken / "Ihre Zeitung ist..."
##### zeitungsattribute_berechnung(prs, df, title="")

Die Funktion erstellt mehrfarbige, gestapelte Bar-Charts für die Darstellung
der Attribut auf den Satz "Ihre Zeitung ist...".

*Paramenter:*
*prs*: das pptx-Objekt, in der die Powerpoint-Präsentation abgelegt ist.
*df*: das Pandas-Dataframe mit den Umfragedaten
*title*: voreingestellt ist "Zeitungsnutzung". Mit der Angabe des Ausgabe-Kürzels wird der Ausgabename mit in den Titel geschrieben

Beispiel:
`for elem in ausgaben_liste:
    df_ = df_umfrage[df_umfrage["ZTG"]==elem]
    zeitungsattribute_berechnung(prs, df_, title = elem) `

#### Übersicht Umfrage zu Themen
##### zeitung_themen(prs, df_, title="")

*Parameter:*

*prs*: pptx-Objekt mit der Powerpoint-Präsentation
*df*: Pandas-Dataframe mit den Umfrageergebnissen
*title*: voreingestellt ist "Zeitungsnutzung". Mit der Angabe des Ausgabe-Kürzels wird der Ausgabename mit in den Titel geschrieben

Beispiel:
`for elem in ausgaben_liste:
    df_ = df_umfrage[df_umfrage["ZTG"]==elem]
    zeitung_themen(prs, df_, title=elem)`

#### Kleine Tortengrafiken für Umfragethemen
##### umfrage_pie(prs, df)
Die Funktion erstellt Powerpoint-Folien mit jeweils zwei kleinen Tortengrafiken.
Die Reihenfolge der Pie-Chrats wird über die Liste umfrage_piecharts festgelegt.

*Parameter:*

*prs*: pptx-Objekt mit der Powerpoint-Präsentation
*df*: Pandas-Dataframe mit den Umfrageergebnissen
*title*: voreingestellt ist "Zeitungsnutzung". Mit der Angabe des Ausgabe-Kürzels wird der Ausgabename mit in den Titel geschrieben


