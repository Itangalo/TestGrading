TestGrading
===========

Scripts for setting up Google spreadsheets to collect and sum up scores from
(school) tests. Usable only inside Google spreadsheets.


Documentation in Swedish
========================

Det här skriptet kan hjälpa dig att sätta upp kalkylblad för att sammanställa
poängresultat på prov. Arbetsflödet ser ut så här:

* Använd fliken 'Maxpoäng' för att ange hur många poäng man kan få på varje
  fråga. Det går att ange olika typer av poäng, exempelvis E-, C- och A-poäng.
* Menyvalet 'Lägg till blad för poänginmatning' skapar ett kalkylblad där du
  kan föra in elevernas poängresultat. Du får automatiskt summor för varje del
  på provet, och totalsummor för hela provet (även uppdelat på olika
  poängtyper.)
* Menyvalet 'Lägg till blad för poänggränser' skapar ett blad där du kan ange
  vilka poäng som krävs för olika provbetyg. Gränserna kan anges för varje typ
  av poäng, och det går även att slå samman poäng av olika typer (exempelvis för
  att kräva minst 5 poäng på C- och A-nivå tillsammans).
* Om poänggränser finns satta kan du använda menyvalet 'Sätt provbetyg' för att
  beräkna provbetyg. Provbetyg syns då för varje elev, och du får även en
  sammanräkning av hur många elever som fått varje betyg.

Här följer en närmare beskrivning av dessa steg.


Steg 1: Ange maxpoäng
=====================

Om du inte redan har ett blad 'Maxpoäng' kan du använda menyvalet 'Lägg till
exempelblad för maxpoäng' för att skapa det. Bladet för maxpoäng fyller du i på
det här viset:

* På första raden anger du vilka typer av poäng man kan få på provet. Det är
  vanligtvis E-, C- och A-poäng, men du kan använda vilka namn du vill och du
  kan använda hur många typer av poäng du vill.
* På raderna under anger du namn på varje fråga, och hur många poäng av de olika
  typerna den frågan kan ge som mest.
* Om ditt prov har olika delar kan du ange namn på en del istället för en fråga.
  I så fall lämnar du alla poängkolumner tomma.

För att göra det lättare att läsa poängen senare kan det vara bra att använda
bakgrundsfärger för att skilja frågorna åt.

Det exempelblad som följer med skriptet innehåller ett exempel på hur dessa
saker kan användas.


Steg 2: Mata in provresultat
============================

När du har angett maxpoängen är det dags att skapa ett blad där du för in de
faktiska poängresultaten. Det gör du genom menyvalet 'Lägg till blad för
poänginmatning'. Då skapas en ny flik i arbetsboken, med en rad för varje elev.
(Du får ange hur många elever som skrev provet innan bladet skapas.)

I bladet finns varje fråga med i en egen kolumn, och om en fråga har mer än en
typ av poäng är dessa uppdelade på flera kolumner. (Om du använt bakgrundsfärger
för frågorna i bladet för maxpoäng kopieras färgerna över till det här bladet.
Det kan göra det rätt mycket lättare att läsa.)

Bladet för poänginmatning innehåller formler som summerar poäng för varje elev,
dels per avsnitt i provet, dels för provet totalt. Summorna är dessutom
uppdelade per poängtyp. Du kan dessutom se hur många procent av maxpoäng som
varje fråga gett i snitt.


Steg 3: Ange poänggränser
=========================

Om ditt prov har poänggränser kan du ange dessa i ett blad som du skapar genom
menyvalet 'Lägg till blad för poänggränser'. Detta blad innehåller ett exempel
på hur poänggränser kan se ut, men du vill troligtvis ändra de gränserna.

På varje rad (med start i cell B2) anger du namn på ett provbetyg, och hur många
poäng som krävs för att få det provbetyget. Du kan dels ange totala poäng, dels
hur många poäng man måste tagit av varje typ. Du kan alltså ange att provbetyg
"B" kräver 25 poäng totalt och 7 poäng på A-nivå.

Om du har poänggränser där flera typer av poäng ska räknas tillsammans kan du
slå använda funktionerna för att slå samman celler. Om ett provbetyg exempelvis
kräver 5 poäng på C- eller A-nivå kan du slå samman cellerna för dessa poäng
och ange '5' i cellen.

Två är bra att känna till när du anger poänggränser:

* Ange poänggränser uppifrån och ner. När skripten sätter ut provbetyg letar
  det efter första matchande provbetyget för en elev.
* Det är bra att skriva en rad för provbetyget 'noll', som inte kräver några
  poäng alls. På det viset kommer alla elever att få ett provbetyg utskrivet,
  även om de inte når upp till lägsta godkända provbetyget.


Steg 4: Sätt ut provbetyg
=========================

När poänggränser är satta, och alla poäng är införda, kan du använda menyvalet
'Sätt provbetyg' för att låta skriptet sätta ut provbetyg för alla elever. I
bladet 'poäng' kommer det då att skrivas ut vilket provbetyg varje elev fick,
och i bladet 'poänggränser' syns en sammanställning av hur många elever som fått
varje provbetyg.
