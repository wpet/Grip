comm # Sessiesettings bepalen
DO .SessionSettings

comm # Opslaan informatie over modules die voor deze client actief zijn in array variabele
v_CountModules = 0
DO .ActiveModules

comm # Procedure MainMenuStr bouwt een string op met daarin de geactiveerde modules en een radiobutton om de module te draaien of niet
comm # De gebruiker klikt op de radiobuttons van de modules die hij wil uitvoeren, die clicks worden opgeslagen in de variabelen RADIO_mnu_main1-x (x = aantal modules client)
comm # In deze procedure wordt ook de info van de actieve modules in de array variabele v_ModuleInfoArr opgeslagen
v_MainMenuStr   = ""
v_ModuleCounter = 1
DO .MainMenuStr WHILE v_ModuleCounter <= v_CountModules AND LENGTH(ini_AutoRunModules) = 0
DO .ShowMainMenuDialog IF LENGTH(ini_AutoRunModules) = 0

DELETE v_Height      OK
DELETE v_HeightStr   OK
DELETE v_ModuleIdStr OK
DELETE v_MainMenuStr OK
DELETE v_ModuleId    OK
DELETE v_ModuleName  OK
DELETE v_ModuleType  OK
DELETE v_ModuleTempl OK

comm # Aflopen en executeren van de modules die voor deze klant zijn geactiveerd
v_ModuleCounter = 1
DO .ExecuteModules WHILE v_ModuleCounter <= v_CountModules

comm # Opruimen bij beeindigen schript
DO .CleanUp

comm ----------------------------------------------------------------------------------------------------------------

comm # Bepalen specifieke settings voor deze sessie van Grip
PROCEDURE SessionSettings
    comm # Procedure Logging verwijderen
    DELETE FORMAT ProcedureLog_ OK
    comm # ClientID opzoeken en variabele package vullen
    OPEN clients_
    SET FILTER TO SUBSTRING(clientid,1,15) = ini_ClientID + BLANKS(15-LENGTH(ini_ClientID))
    COUNT
    DO .ErrorClientIdNotFound IF COUNT1 = 0
    LOCATE IF SUBSTRING(clientid,1,15) = ini_ClientID + BLANKS(15-LENGTH(ini_ClientID))
    comm # Vullen variabelen voor de gebruikte packages, package1=finpakket
    v_Package1 = ALLTRIM(package1)
    v_Package2 = ALLTRIM(package2)
    v_Package3 = ALLTRIM(package3)
    v_Package4 = ALLTRIM(package4)
    v_Package5 = ALLTRIM(package5)
    comm # Veldnamen standaard output in een ini_variabele opslaan
    DO .SetStandardOutputFields_%v_Package1%
    comm DO .SetStandardOutputFields_%v_Package2%
    comm DO .SetStandardOutputFields_%v_Package3%
    comm DO .SetStandardOutputFields_%v_Package4%
    comm DO .SetStandardOutputFields_%v_Package5%
    CLOSE PRIMARY
    DELETE FORMAT clients_ OK
RETURN

PROCEDURE ErrorClientIdNotFound
    PAUSE 'Klantcode niet gevonden.' WAIT 2
    DO Exit%sec%
RETURN

comm # Op basis van ClientID nagaan welke modules actief zijn en aantal actieve modules bepalen en array variabele v_ModuleInfoArr initialiseren
PROCEDURE ActiveModules
    OPEN clientmodulemap_
    v_FilterModules    = 'clientid = ' + '"' + ini_ClientID + '"'
    v_FilterModules    = 'clientid = ' + '"' + ini_ClientID + '"' + ' AND MATCH(moduleid,' + '"' + REPLACE(ini_AutoRunModules,',','","') + '"' + ')' IF LENGTH(ini_AutoRunModules) > 0
    SET FILTER TO %v_FilterModules%
    SORT ON moduleseq TO "client_module_temp_" OPEN
    DELETE v_FilterModules OK
    COUNT
    DO .ErrorNoModulesFound IF COUNT1 = 0
    v_CountModules  = COUNT1
    ASSIGN v_ModuleInfoArr[v_CountModules] = BLANKS(1000)
    v_ModuleCounter = 1
    DO .LoadActiveModuleInfo WHILE v_ModuleCounter <= v_CountModules
    DELETE FilterModules OK
RETURN

PROCEDURE ErrorNoModulesFound
    PAUSE 'Geen modules voor deze klant gevonden.' WAIT 2
    DO Exit%sec%
RETURN

comm # Zet variabelen voor de modules die unattended moet draaien
PROCEDURE LoadActiveModuleInfo
    DO .LoadModuleInfo
    RADIO_%v_ModuleId% = 1
    v_ModuleCounter    = v_ModuleCounter + 1
RETURN

PROCEDURE LoadModuleInfo
    comm # Opslaan moduleinfo in array v_ModuleInfoArr
    LOCATE IF RECNO()  = v_ModuleCounter
    v_ModuleId         = moduleid
    v_ModuleName       = ALLTRIM(modulename)
    v_ModuleType       = ALLTRIM(moduletype)
    v_ModuleTempl      = ALLTRIM(moduletemplate)
    v_ModuleInfoArr[v_ModuleCounter] = v_ModuleId + '|' + v_ModuleName + '|' + v_ModuleType + '|' + v_ModuleTempl
RETURN

comm # Opbouwen menu op basis van gegevens uit clientmodulemap voor deze client
PROCEDURE MainMenuStr
    comm # Opbouwen string met menu-informatie van af te beelden modules in het hoofdmenu
    v_Height        = 70 IF v_ModuleCounter = 1
    v_Height        = v_Height + 30 IF v_ModuleCounter >= 1
    v_HeightStr     = ALLTRIM(STRING(v_Height,12))
    v_ModuleId      = ALLTRIM(SPLIT(v_ModuleInfoArr[v_ModuleCounter],'|',1))
    v_ModuleName    = ALLTRIM(SPLIT(v_ModuleInfoArr[v_ModuleCounter],'|',2))
    v_ModuleIdStr   = ""
    v_ModuleIdStr   = ' [' + v_ModuleId + ']' IF v_Realm = "dev"
    v_RadioButton   = 'RADIO_' + v_ModuleId
    v_MainMenuStr   = v_MainMenuStr + '(TEXT TITLE "' + ALLTRIM(STRING(v_ModuleCounter,3)) + ') ' + v_ModuleName + v_ModuleIdStr + REPEAT(".",200-LENGTH(v_ModuleName)) + '" AT 20 ' + v_HeightStr + ') (RADIOBUTTON TITLE "ja;nee" TO "' + v_RadioButton + '" AT 500 ' + v_HeightStr + ' DEFAULT 2 HORZ) (TEXT TITLE "          " AT 570 ' + v_HeightStr + ') '
    DELETE %v_RadioButton% OK
    v_ModuleCounter = v_ModuleCounter + 1
RETURN

comm # Tonen hoofdmenu op basis van de modules die voor deze client geactiveerd zijn
PROCEDURE ShowMainMenuDialog
    DIALOG (DIALOG TITLE "Hoofdmenu" WIDTH 600 HEIGHT 500 ) (BUTTONSET TITLE "&OK;&Annuleren" AT 500 440 DEFAULT 1 ) (TEXT TITLE "GRIP :: Generiek risicoanalyse en informatie platform" AT 20 10 ) (TEXT TITLE "Versie 1.1 01-12-2014" AT 20 30 ) (TEXT TITLE "Overzicht modules:" AT 20 70 ) (TEXT TITLE "Uitvoeren?" AT 500 70 ) %v_MainMenuStr%
RETURN

comm # Uitvoeren modules die geselecteerd zijn
PROCEDURE ExecuteModules
    v_ModuleId       = ALLTRIM(SPLIT(v_ModuleInfoArr[v_ModuleCounter],'|',1))
    v_ModuleName     = ALLTRIM(SPLIT(v_ModuleInfoArr[v_ModuleCounter],'|',2))
    DO .ExecuteScripts IF RADIO_%v_ModuleId% = 1
    PAUSE 'Scripts van module ' + v_ModuleId + '_' + v_ModuleName + ' afgerond.' WAIT 2 IF RADIO_%v_ModuleId% = 1
    v_ModuleCounter  = v_ModuleCounter + 1
RETURN

comm # Ophalen scripts die voor de module moeten worden uitgevoerd
PROCEDURE ExecuteScripts
    OPEN modulesscriptmap_
    v_Filter         = "moduleid = v_ModuleInfoArr[v_ModuleCounter] AND active = 'j'"
    v_ModuleType     = ALLTRIM(SPLIT(v_ModuleInfoArr[v_ModuleCounter],'|',3))
    v_ModuleTemplate = ALLTRIM(SPLIT(v_ModuleInfoArr[v_ModuleCounter],'|',4))
    comm # Kopieer template van de module vanaf de server naar de lokale machine en overschrijf de lokale versie
    DO .CopyTemplateFromServer IF v_Realm = "prod" AND LENGTH(ALLTRIM(v_ModuleTemplate)) > 0 AND v_ModuleTemplate <> '0'
    DO .CopyTemplateFromLocal  IF v_Realm = "dev"  AND LENGTH(ALLTRIM(v_ModuleTemplate)) > 0 AND v_ModuleTemplate <> '0'
    comm # Bepalen naam van het output bestand
    DO .SetOutputName IF LENGTH(ALLTRIM(v_ModuleTemplate)) > 0 AND v_ModuleTemplate <> '0'
    OPEN modulesscriptmap_
    SET FILTER TO %v_Filter%
    EXTRACT ALL FIELDS TO "scripts_temp_" OPEN
    COUNT
    v_CountScripts = COUNT1
    v_ScriptCounter = 1
    DO .RunScript WHILE v_ScriptCounter <= v_CountScripts
RETURN

comm # Ophalen naam van het script en uitvoeren
PROCEDURE RunScript
    OPEN scripts_temp_
    LOCATE IF RECNO() = v_ScriptCounter
    v_ScriptName = ALLTRIM(scriptname)
    EXTRACT FIELDS SUBSTRING(v_ModuleName,1,50) AS "ModuleName" SUBSTRING(v_ScriptName,1,50) AS "ScriptName" TO ProcedureLog_ FIRST 1 APPEND
    CLOSE PRIMARY
    CLOSE SECONDARY
    SET ECHO ON
    comm ----------------------------------------------------------------------------------------------------------------------------------------------------
    comm Script: %v_ScriptName%%Sec%
    comm ----------------------------------------------------------------------------------------------------------------------------------------------------
    DO %v_ScriptName%%Sec%
    SET ECHO OFF
    v_ScriptCounter = v_ScriptCounter + 1
RETURN

comm # Kopieren template van server naar lokale machine
PROCEDURE CopyTemplateFromServer
    v_CopyFile = "ServerFrom=SERVER Gripserver&From=%ini_ServerFolder%\_data_\%v_Package1%\%v_ModuleTemplate%.xlsx&ServerTo=LOCAL&To=%ini_TemplateFolder%\%v_ModuleTemplate%.xlsx"
    DO CopyFile%Sec%
RETURN

comm # Kopieren template van lokale servermap naar templatefolder in werkmap
PROCEDURE CopyTemplateFromLocal
    v_CopyFile = "From=%ini_ServerFolder%\_data_\%v_Package1%\%v_ModuleTemplate%.xlsx&To=%ini_TemplateFolder%\%v_ModuleTemplate%.xlsx"
    DO CopyFile%Sec%
RETURN

comm # Naam van outputbestand definieren
PROCEDURE SetOutputName
    v_TimeStamp  = ""
    DO TimeStamp%Sec%
    ini_OutputFolder = ini_OutputFolder + "\" IF SUBSTRING(ini_OutputFolder,LENGTH(ini_OutputFolder),1) <> "\"
    v_OutputName = ini_OutputFolder + v_ModuleType + "_" + v_TimeStamp + ".xlsx"
    v_CopyFile   = "From=%ini_TemplateFolder%\%v_ModuleTemplate%.xlsx&To=%v_outputName%"
    DO CopyFile%Sec%
RETURN

comm # Opruimen
PROCEDURE CleanUp
    DELETE FORMAT client_module_temp_ OK
    DELETE FORMAT scripts_temp_ OK
RETURN

comm # Zetten van de velden van het bestand FinMuta, afhankelijk van het pakket
PROCEDURE SetStandardOutputFields_civ
   ini_FinMutaFields = "RecNo BdNr Positie DocNr DocSoort DocStatus DocDat DatInvoer MutDat OmrekDat FaktOntvDat BoekDat BoekJaar BoekJaarMaand Periode BoekRegId RefCode GebrNaam RegDoor TransCode InkomstenUitgavenBBV BBV OmsBBV Functie OmsFunctie Referentie DocKopTekst Valuta RekSoort GrootboekRek OmsGrootboekRek OmsGrootboekRek2 KostenDrager OmsKostenDrager Tekst DebCred Bedrag Debet Credit BelBedrag BedrBTW BelCode Omsbelcode Factor GRSD Perc PercAfdracht PercBTW PercBCF PercKVH BTWBerTotaal AfdrachtBedrBer BTWBedrBer BCFBedrBer KVHBedrBer OpCode InkDoc Toewijzing  DatVerval DebNr CredNr DebCredNr NaamDebCred VereffDat DatInvVereff VereffDoc IngDatBetTerm BetTerm1 BetTerm2 BoekingsSleutel BankRek IBAN GeldigVanIBAN WbsElement OmsWbsElement KostenPlaats OmsKostenPlaats Order OmsOrder CO CN BTW_GRSD BCF_GRSD KVH_GRSD Omschrijving_GRSD BtwVanNaar BtwJuist GrVlak_key"
RETURN

PROCEDURE SetStandardOutputFields_key2
   ini_FinMutaFields = "RecNo VolgNummer Bedrijf DocumentNummer BoekingsDatum ValutaDatum DagboekNummer TypeDagboek OmschrijvingDagboek DienstJaar PeriodeGrootboek DecimaleRubriek OmschrijvingDecimaleRubriek GrootboekNummer OmschrijvingGrootboekNummer Taak OmschrijvingTaak KostenSoort OmschrijvingKostenSoort CategorieBBV InkomstenUitgavenBBV OmschrijvingCategorieBBV DebiteurNummer DebiteurNaam CrediteurNummer CrediteurNaam OmschrijvingBoeking BedragDebet BedragCredit BedragSaldo BedragBTW BedragBTW_FISC BedragBTW_BCF BedragBTW_KVH BedragBTW_AFDR CO CN BTW_GRSD BCF_GRSD KVH_GRSD Omschrijving_GRSD BtwVanNaar BtwJuist Functie FunctieOmschrijving GrVlak_key"
RETURN
