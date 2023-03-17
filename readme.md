

# Návod na používání programu VIES VAT Checker.

## Po spuštění
Pokud spouštíte program poprvé, je potřeba vybrat excel soubor
s daty, které chcete zkontrolovat. V prvním sloupci musí být
čísla DIČ, která chcete zkontrolovat.
Druhý sloupec musí být prázný, do něj se uloží výsledek.

Pokud jste již v minulosti program použili, poslední použitý
soubor se automaticky vybere.

## Výběr souboru
Pokud máte již vybraný soubor, stačí kliknout na tlačítko
Start a program začne zpracovávat data.
Pokud jste zadali špatný soubor, můžete jej změnit kliknutím
na tlačítko Vybrat soubor.
Pokud soubor není čitelný nebo není správného formátu,
zobrazí se hláška: **Chyba při čtení souboru.**
Po výběru souboru a jeho načtení se zobrazí počet řádků pro zpracování ve formátu:
**Zpracování: 0/330.**
Pokud v souboru jsou již zpracované řádky, tak tyto řádky
nebudou zpracovány znovu. Číslo značící počet DIČ pro zpracování se tím sníží o již zpracované.

## Zpracování dat
Po kliknutí na tlačítko Start se začnou zpracovávat data.
DIČ se zpracovávají po jednom, aktuálně zpracovávané DIČ se
zobrazí v prvním řádku. Po úspěšném zpracování se zobrazí 
v pravém sloupci výsledek. Počet zpracovaných řádků se zobrazí
v stavovém řádku ve formátu: **Zpracování: 1/330.**
Po každém úspěšném zpracování se data uloží do souboru.

    ###Chyby při zpracování
    ###Na každý DIČ se pošle dotaz na server VIES. 
    ###Pokud server odpověděl správně bude DIČ zobrazeno v pravém sloupci.
    ###Pokud server odpověděl chybou, program se bude ještě 9 krát pokoušet znovu. 
    ###Pokud se nepodaří získat odpověď, během 10 pokusů, 
    ###nebo během 60 vteřin bude DIČ označeno jako chybné a bude zobrazeno v levém sloupci.
    ###V souboru bude v druhém sloupci zapsáno **ERROR**.

Po kompletním zpracování se zobrazí hláška: **Status dokončeno.**
Pokud nekteré DIČ nebylo možné zpracovat, je možné spustit program znovu a zpracovat je znovu.
Program načte pouze ty DIČ ze souboru které mají v druhém sloupci prázdnou buňku a nebo 
slovo **ERROR**.
