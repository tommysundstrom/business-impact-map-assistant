business-impact-map-assistant
=============================

Transforms business impact maps in Excel to different views.



Title: Effektkarta – ett nytt sätt beskriva den
Author: Tommy Sundström
Base Header Level: 1

Facebook:
Till alla vänner som är effektkartläggare: Jag behöver hjälp med att testa ett nytt sätt att göra effektkartor, som gör det enklare att se kartan ur olika synvinklar och se vilka samband där finns. Om du har lust att testa, skicka ett meddelande. Du måste ha en mac. 
Effektkartor – ett annat sätt att rita dem
Jag gillar effektkartor (vad är det?). Men jag kan vara skeptisk till sättet de ritas på. 
[Bild Örebro]
Missförstå mig rätt. Mindmaps är utmärkta när man börjar spåna på vad som ska vara i kartan, och de kan också vara bra när man presenterar slutresultatet för en del intressenter. 
<p class='in'>Men de ger inte något särskilt bra stöd när man ska analysera och bygga upp sin effektkarta. </p>
Ta Örebros effektkarta som exempel (jag har valt den eftersom den finns publikt tillgänglig och används som exempel på hur en effektkarta kan se ut). Här ligger syftet med projektet, alla verksamhetsmål, som en absurd klump i mitten. Kopplingen mellan verksamhetsmål och åtgärd har gått förlorad – kartan ger ingen vägvisning utan man får gissa vilken åtgärd som motiveras av vilket mål. 
Att det blir så tror jag har två orsaker.
<p class='in'>Delvis ligger problemet i mindmappens natur. Den förutsätter ett centralt ämne i mitten och sedan allt mer detaljer ju längre ut man kommer. En effektkarta däremot är mer jämntjock, med undantag att mängden åtgärder ofta är ganska stor. Framförallt är det sällan bara ett syfte som ska uppnås, utan de brukar vara ett litet gäng. Resultatet blir överlastade mittnoder. </p>
<p class='in'>Men jag tror också att det finns en djupare orsak. Det centrala i en effektkarta är sammankopplingen av verksamhetsmål och användningsmål i kontext av en målgrupp; endast om det finns användare som själva har ett mål som stödjer det har verksamheten någon chans att nå sitt mål. </p>
<p class='in'>En modell som på detta sätt rör sig över tre olika dimensioner är alltid svår att återge på ett överskådligt sätt på ett platt papper. 	</p>
Lyckligtvis kan vi slänga datorkraft på problemet och lösa det genom att ur samma modell lyfta fram de två dimensioner man för ögonblicket är mest intresserad av. 
För att illustrera, har jag skrivit in delar av Örebros effektkarta i ett vanligt Excel-blad. 
[Bild]
(Excel-bladet finns att ladda ner senare i artikeln)
Den första och kanske största skillnaden är att verksamhetsmålen inte klumpas samman, utan varje användningsmål, och åtgärderna som följer av detta, ses i perspektiv av ett (eller flera) verksamhetsmål. 
<p class='in'>Eftersom denna koppling inte finns i den ursprungliga effektkartan har jag tvingats gissa hur det är tänkt, men det var relativt lätt att se vad som hörde vart.</p>
Redan i arbetet med att göra effektkartan på detta sätt, uppstår en del insikter. 
<p class='in'>Till exempel att verksamhetsmålen ”Bättre stöd i vardagssituationer” och ”Enklare att veta vad/hur man ska göra” kraftigt överlappar. Om en åtgärd är tänkt att höra till den ena eller den andra av dessa, är nästan helt omöjligt att gissa.</p>
<p class='in'>Inget hindrar att man behåller och kopierar in samma användare och användningsmål under båda. Men det skulle göra kartan större utan att den bidrar med mer insikt, så jag förenklar hellre genom att bara behålla ”Enklare att veta vad/hur man ska göra”, och låta den ta över mätpunkterna från ”Bättre stöd i vardagssituationer”.</p>
<p class='in'>(Som du ser skriver jag inte in mätpunkterna i effektkartan, eftersom jag tycker det komplicerar den utan att ge något mervärde. De kan hållas på ett papper för sig.)</p>
En annan och kanske intressantare sak som blir uppenbar under det här arbetet är att ett av verksamhetsmålen inte har något stöd i verksamheten.
<p class='in'>”Ökad användning av intranätet” motsvaras inte av någon användares mål. Det är ett mål utan konkret nytta för någon. </p>
<p class='in'>Vilket är uppenbart när man tänker efter. En ökad användning av intranätet kan vara en bieffekt på att det blivit bättre och användbarare – men som verksamhetsmål är det ett självändamål. </p>
Med denna förenklade målbild, kommer jag fram till kartan som fanns att [ladda ner] ovan.
[Bild + Länk]
Denna och de andra kan laddas ner. Gjord med Excel 2011, på mac, men torde kunna öppnas med de flesta kalkylprogram.
I sig är detta inte upphetsande. Det är lika roligt som att läsa en databastabell. 
<p class='in'>Men det intressanta händer som sagt när man tillsätter lite datakraft och kör ett script på effektkartan.</p>
<p class='in'></p>
Nedladdning Viktigt: För att kunna köra det här scriptet måste du högerklicka på det och välja öppna, samt godkänna att det körs på din maskin. 
View Raw
https://github.com/tommysundstrom/business-impact-map-assistant/blob/master/The%20application/business-impact-map-assistant.zip?raw=true
En applikation som heter ”business-impact-map-assistant” laddas ner. Högerklicka på den och välj ”Öppna”. Klicka sedan på ”Öppna”-knappen. 
<p class='in'>Nu ska applikationen vara igång. Den är superenkel – består bara av en knapp ”Generera vyer”. Se till att effektkartan är det aktiva dokumentet i Excel, klicka på den och gå och hämta lite kaffe medan den arbetar. När du kommer tillbaka bör du se ett antal nya kalkylblad som flikar i sidans underkant. </p>
Tom effektkarta: View Raw 	https://github.com/tommysundstrom/business-impact-map-assistant/blob/master/business-impact-map-assistant/Resources/empty-impact-map.xlsx?raw=true
Exempel-karta: https://github.com/tommysundstrom/business-impact-map-assistant/blob/master/business-impact-map-assistant/Resources/impact-map-exemple.xlsx?raw=true
---
## Olika perspektiv
När man kört scriptet gör det olika vyer på effektkartan, som läggs som egna blad – man navigerar mellan dem med flikarna i underkant av sidan. Dessa vyer gör det enkelt att svara på frågor som:
 
* Vilka åtgärder föreslår vi för att hjälpa administratörerna – och varför?
* Vad det är för motiv som driver användarna – och hur bidrar de till att uppnå verksamhetsmålen?
* Vilka är verksamhetens användningsmål och vilka användare är viktiga för att de nås (alltså den fråga som effektkartor traditionellt är bra på att svara på). 
Och, till sist, en vy ordnad utifrån de konkreta åtgärderna. På många sätt den intressantaste, i vart fall om man jobbar agilt, eftersom man här har en färdig backlog – med bonuset att varje åtgärd är kopplad till det verksamhetsmål som motiverat den, så att det blir lättare att bedöma vad som är angelägnast att göra just nu. 
## När användaren är ett beteende
När jag vänder och vrider på kartan så här, dyker det upp en fråga jag funderat mycket över. Om effektkartans målgrupp – användaren – är ett personifierat beteende, här till exempel ”Spridaren”, ”Den undrande” och ”Samarbetaren”, så blir den egentligen bara en omforumlering av det som redan står under användningsmålet. Spridaren vill sprida information, den undrande vill hitta den och samarbetaren vill samarbeta. 
<p class='in'>Jag vet att långt skickligare effektkartemakare än jag gör på detta sätt, så det kan vara något jag missat, men personligen föredrar jag att antingen använda de personas man tagit fram (om effektkartans huvudsyfte är att bygga förståelse och empati) eller yrkesgrupper/traditionella demografiska målgrupper (när det är viktigt att kunna mäta effekterna). </p>
<p class='in'>Men, det är en annan diskussion.</p>
## Nya möjligheter
Att ha effektkartan i ett kalkylblad öppnar en mängd möjligheter. 
* Mindmappens möjlighet att öppna och stänga grenar är elegant, men de tenderar att bli väldigt yviga. I många situationer är det praktiskt att kunna skriva ut effektkartan på A4or. 
* I en traditionell effektkarta är det (förmodligen som en följd av hur mindmaps fungerar) bara åtgärderna som detaljeras. Men i Excelbladet skulle det gå bra att göra motsvarande sak med syftena, visa hur de hänger samman med mer övergripande mål.
<p class='in'>På så sätt kan man visa kopplingen mellan projektets effektmål och de mer övergripande affärsstrategier som organisationen har, vilket kan vara en stor fördel i förankringsarbetet.</p>
* Det är enkelt att lägga till hjälpkolumner för att hålla reda på beräknad tidsåtgång och prioriteringarna av backloggen.
* Den som behöver stöd att prioritera i stora effektkartor kan lägga in värderingar av de olika kolumnerna. Ange hur stor andel av projektets totala effekt de olika syftena står för, hur viktiga de olika målgrupperna är, vilket bidrag olika användningsmål ger till uppfyllandet av effektmålen, och så vidare. Med hjälp av detta kan man sedan lyfta fram de åtgärder som kommer att ge störst effekt.
* Ökat utrymme i kartans mitt öppnar också för att arbeta med fler verksamhetsmål. För min del är jag framförallt intresserad av att använda effektkartan även för att arbeta med varumärket. 
## Men, jag gillar ju mindmaps…
Mindmaps har definitivt sina styrkor. De är trevliga att se på och ofta betydligt roligare att arbete med än kalkylblad. 
<p class='in'>Lyckligtvis behöver det inte vara någon motsättning. Under ytan arbetar scriptet med apple events, vilket betyder att det är lika enkelt att göra en vy som en mindmap som det är att göra den som ett kalkylblad. Och på samma sätt som med kalkylbladen kan man då välja ur vilket perspektiv man vill se den, vad som ska vara i mitten av mindmappen. </p>
<p class='in'>Men hur man gör det får bli ämnet för en annan artikel. </p>
---
TEKNIK
Behöver Excell 2011 för mac. Har du det inte, men vill prova hur det här fungerar kan du ta en provprenumeration på Office-paketet här: http://office.microsoft.com/sv-se/kop-microsoft-office-365-home-premium-FX102853961.aspx så får du den. (Tips: Vill du bara prova kan du prenumerera och sedan säga upp prenumerationen direkt, programmet fortsätter ändå att fungera första månaden ut.) 
---
RUTA 1
Några ord om Excel-bladet.
* Celler kopieras automatiskt nedåt (om värdet till vänster om dem är detsamma). Det betyder att du inte behöver upprepa samma sak, bara lämna tomt, så förstår scriptet. 
* Lite extra-kolumner om man vill vara mer detaljerade om åtgärderna.
Några ord om scriptet som gör det här. 
<p class='in'>Det är gjort för mac och Excel 2011 eller Office 365. Det är alldeles möjligt att det fungerar även på andra versioner av Excel, men det har jag inte testat.</p>
<p class='in'>Att det bara fungerar med mac är dock helt klart, eftersom det arbetar via apple events. För den som kan VBA är det förmodligen en trivial sak att omvandla det till ett Excel-makro och därigenom få det att fungera även på en PC, men tyvärr räcker inte mina kunskaper till för det. Om man arbetar med Excel-makron kan man inte heller på samma enkelt sätt visa effektkartan i form av en mindmap.</p>
Tyvärr finns det för tillfället inte många scriptspråk att välja mellan på macen, så jag har skrivit det i applescript – som måste vara ett av världens värsta språk att arbeta med. 
Scriptet är gjort i applescript, som är ett scriptspråk med en del quirks. Så har jag till exempel inte hittat något sätt att ta bort gamla vyer utan att godkänna borttagningen för var och en. 
RUTA 2
## Vad är en effektkarta?
Effektstyrning är en projektstyrningsmetod som lägger stort fokus vid att säkerställa att rätt effekter kommer ur projektet, att projektet åstadkommer verksamhetsnytta. En viktig del i metoden är effektkartan, som visar sambandet mellan verksamhetsmål och de åtgärder man gör. 
Tillbaka till artikeln

