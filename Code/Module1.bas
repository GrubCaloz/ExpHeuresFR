Attribute VB_Name = "Module1"
Sub Export_Word()

ActiveWorkbook.Save

'Verification de la version Excel
Dim Version
Version = AppVersion()
    'Version = 2016 'pour test
If Not (Version = 365 Or Version >= 2021) Then
'Version incompatible
    MsgBox "Les fonctions de ce classeur ne sont pas compatibles avec votre version d'Excel" & vbCrLf & "Utilisiez Office 365, Office 2021 ou sup�rieur" & vbCrLf & vbCrLf & "Votre version actuelle: " & Version, vbCritical, "Version incompatible"
    GoTo SortieRapide
End If

'D�finition du document word
Dim wordApp As New Word.Application
Dim wDoc As Word.Document
Dim PathWB, docPath As String


'Cr�ation du dossier pour stocker les fichiers export�s
Dim PathOutput, CopyName, FacturOK, AnnexeOK As String
Dim fsoObj As Object
Set fsoObj = CreateObject("Scripting.FileSystemObject")
PathWB = fsoObj.GetAbsolutePathName(Application.ActiveWorkbook.path)

'Contr�le si le dossier de travail est OneDrive
If InStr(PathWB, "http") <> 0 Then
MsgBox "Le classeur est stock� dans un dossier OneDrive" & vbCrLf & "Les fichiers seront export�s sur votre bureau", vbInformation, "OneDrive"
PathWB = CreateObject("WScript.Shell").SpecialFolders("Desktop")
End If

PathOutput = fsoObj.BuildPath(PathWB, "Export_HExp")

If Not fsoObj.FolderExists(PathOutput) Then
    fsoObj.CreateFolder (PathOutput)
End If


'Va chercher le fichier sur Github
docPath = "https://github.com/GrubCaloz/ExpHeuresFR/raw/main/Fichiers/02_Formulaire_Org.docx"
Set wordApp = CreateObject("word.application")

'wordApp.Visible = True 'Permet le debug

Set wDoc = wordApp.Documents.Open(docPath, , True)


'D�termine le b�n�ficiaire du paiement
Dim Benf As Boolean
If MsgBox("Je suis le b�n�ficiare", vbYesNo, "B�n�ficiaire") = vbYes Then
    Benf = True
Else
    Benf = False
End If

    
'Lecture des valeurs du tableau pour facture
Dim PrepaH, TpH, SurvH, CorrH, DeplKM, DeplTP, NbrRepas, TotDivers As Single
PrepaH = Worksheets("Annexe").Range("PrepaH")
TpH = Worksheets("Annexe").Range("TotPrepaT")
SurvH = Worksheets("Annexe").Range("SurvH")
CorrH = Worksheets("Annexe").Range("CorrH")
DeplKM = Worksheets("Annexe").Range("DeplKM")
DeplTP = Worksheets("Annexe").Range("DeplTP")
NbrRepas = Worksheets("Annexe").Range("NbrRepas")
TotDivers = Worksheets("Annexe").Range("TotDivers")

'Attribution des tarifs
Dim TarifPrepa, TarifTP, TarifCorr, TarifSurv, TarifKM, TarifRepas As Single
TarifPrepa = Worksheets("Param�tres").Range("TarifPrepa").Value
TarifTP = Worksheets("Param�tres").Range("TarifTP").Value
TarifSurv = Worksheets("Param�tres").Range("TarifSurv").Value
TarifCorr = Worksheets("Param�tres").Range("TarifCorr").Value
TarifKM = Worksheets("Param�tres").Range("TarifKM").Value
TarifRepas = Worksheets("Param�tres").Range("TarifRepas").Value



'Info Profession
Dim ProfFact As String
ProfFact = Worksheets("Annexe").Range("ProfFact").Value
Dim Profs As Word.Range
Set rng_Prof = wDoc.Bookmarks("Prof").Range
rng_Prof.Text = ProfFact

'Type Examen
Dim ExaType As String
Dim Final, Intermediaire, Partiel As Word.Range
Set rng_Final = wDoc.Bookmarks("Final").Range
Set rng_Intem = wDoc.Bookmarks("Intermediaire").Range
Set rng_Partiel = wDoc.Bookmarks("Partiel").Range
ExaType = Worksheets("Annexe").Range("ExaTypeFact").Value

If ExaType = "Final" Then
wDoc.FormFields("Final").CheckBox.Value = True
End If

If ExaType = "Intermediaire" Then
wDoc.FormFields("Intermediaire").CheckBox.Value = True
End If

If ExaType = "Partiel" Then
wDoc.FormFields("Partiel").CheckBox.Value = True
End If

  
'Donn�es expert
Dim ExpNom As String
ExpNom = Worksheets("Param�tres").Range("ExpNom").Value

    wDoc.FormFields("BenefMoi").CheckBox.Value = True
    Call WordWrite(wDoc, "ExpNom", ExpNom)
    Call WordWrite(wDoc, "Adre", Worksheets("Param�tres").Range("Adre").Value)
    Call WordWrite(wDoc, "ComplExp", Worksheets("Param�tres").Range("ComplExp").Value)
    Call WordWrite(wDoc, "NpaExp", Worksheets("Param�tres").Range("NpaExp").Value)
    Call WordWrite(wDoc, "TelExp", Worksheets("Param�tres").Range("TelExp").Value)
    Call WordWrite(wDoc, "BanqueExp", Worksheets("Param�tres").Range("BanqueExp").Value)
    Call WordWrite(wDoc, "IbanExp", Worksheets("Param�tres").Range("IbanExp").Value)

'Donn�es entreprises si non b�n�ficiaire
If Not Benf Then
    wDoc.FormFields("BenefEmpl").CheckBox.Value = True
    wDoc.FormFields("BenefMoi").CheckBox.Value = False
    Call WordWrite(wDoc, "EmplNom", Worksheets("Param�tres").Range("EmplNom").Value)
    Call WordWrite(wDoc, "AdreEntre", Worksheets("Param�tres").Range("AdreEntre").Value)
    Call WordWrite(wDoc, "ComplEntre", Worksheets("Param�tres").Range("ComplEntre").Value)
    Call WordWrite(wDoc, "NpaEntre", Worksheets("Param�tres").Range("NpaEntre").Value)
    Call WordWrite(wDoc, "TelEntre", Worksheets("Param�tres").Range("TelEntre").Value)
    Call WordWrite(wDoc, "BanqueEntre", Worksheets("Param�tres").Range("BanqueEntre").Value)
    Call WordWrite(wDoc, "IbanEntre", Worksheets("Param�tres").Range("IbanEntre").Value)
End If

'Salari� oui/non
Dim SalarieStat As String
SalarieStat = Worksheets("Param�tres").Range("SalarieStat").Value

If SalarieStat = "Salari�" Then
    wDoc.FormFields("Sal").CheckBox.Value = True
Else
    wDoc.FormFields("Indep").CheckBox.Value = True
End If


'donn�e en-t�te de la facture
Dim dateMin, dateMax, NumFinanace
NumFinance = Application.WorksheetFunction.XLookup(ProfFact, Range("Tbl_Prof[Professions]"), Range("Tbl_Prof[N� Finances]"))

Call WordWrite(wDoc, "NumFinance", CStr(NumFinance))
Call WordWrite(wDoc, "NumCollab", Worksheets("Param�tres").Range("NumCollab").Value)
Call WordWrite(wDoc, "DateNaiss", Worksheets("Param�tres").Range("DateNaiss").Value)
Call WordWrite(wDoc, "NumAvs", Worksheets("Param�tres").Range("NumAvs").Value)
Call WordWrite(wDoc, "AdMail", Worksheets("Param�tres").Range("AdMail").Value)


'Calcul des dates Min Max
'toujours des soucis avec cette fonction
Dim TblTache As ListObject
Set TblTache = Worksheets("T�ches").ListObjects("Tbl_tache")

dateMin = WorksheetFunction.Min(TblTache.ListColumns("Date").Range)
dateMax = WorksheetFunction.Max(TblTache.ListColumns("Date").Range)
'dateMin = Application.WorksheetFunction.MinIfs(Range("Tbl_tache[Date]"), Range("Tbl_tache[Type examen]"), ExaType, Range("Tbl_tache[Profession]"), ProfFact)
'dateMax = Application.WorksheetFunction.MaxIfs(Range("Tbl_tache[Date]"), Range("Tbl_tache[Type examen]"), ExaType, Range("Tbl_tache[Profession]"), ProfFact)

Call WordWrite(wDoc, "Dates", "Du " & Format(dateMin, "dd.mm.yyyy") & " au " & Format(dateMax, "dd.mm.yyyy"))

'Calcul des valeures de facturation et report dans le Word
    'Valeurs arrondies au 1/4 d'heure
    Call WordWrite(wDoc, "PrepaHeure", WorksheetFunction.MRound(PrepaH, 0.25))
    Call WordWrite(wDoc, "TPHeure", WorksheetFunction.MRound(TpH, 0.25))
    Call WordWrite(wDoc, "SurvHeure", WorksheetFunction.MRound(SurvH, 0.25))
    Call WordWrite(wDoc, "CorrHeure", WorksheetFunction.MRound(CorrH, 0.25))
    Call WordWrite(wDoc, "DeplKMs", Round(DeplKM, 2))
    Call WordWrite(wDoc, "NbrRepass", CStr(NbrRepas))
    'Call WordWrite(wDoc, "DeplTPs", CStr(DeplTP))
    
    Dim TotPrepa, TotTp, TotSurv, TotCorr, TotDepl, TotRepas As Single
    TotPrepa = WorksheetFunction.MRound(PrepaH * TarifPrepa, 0.05)
    TotTp = WorksheetFunction.MRound(TpH * TarifTP, 0.05)
    TotSurv = WorksheetFunction.MRound(SurvH * TarifSurv, 0.05)
    TotCorr = WorksheetFunction.MRound(CorrH * TarifCorr, 0.05)
    TotDepl = WorksheetFunction.MRound(DeplKM * TarifKM, 0.05)
    TotRepas = WorksheetFunction.MRound(NbrRepas * TarifRepas, 0.05)
    
    Call WordWrite(wDoc, "PrepaCHF", CStr(TotPrepa), "CHF")
    Call WordWrite(wDoc, "TPCHF", CStr(TotTp), "CHF")
    Call WordWrite(wDoc, "SurvCHF", CStr(TotSurv), "CHF")
    Call WordWrite(wDoc, "CorrCHF", CStr(TotCorr), "CHF")
    Call WordWrite(wDoc, "DeplKMCHF", CStr(TotDepl), "CHF")
    'Call WordWrite(wDoc, "DeplTPCHF", WorksheetFunction.MRound(DeplTP, 0.05), "CHF")
    Call WordWrite(wDoc, "NbrRepasCHF", CStr(TotRepas), "CHF")

    'Totaux
    Dim Tot1_5, Tot6_9 As Single
    Tot1_5 = TotPrepa + TotTp + TotSurv + TotCorr
    Tot6_9 = TotDepl + TotRepas + DeplTP + TotDivers

    Call WordWrite(wDoc, "Tot1_5", WorksheetFunction.MRound(Tot1_5, 0.05), "CHF")
    Call WordWrite(wDoc, "Tot6_9", WorksheetFunction.MRound(Tot6_9, 0.05), "CHF")



'Export
'Cr�aton des noms de fichiers
PathOutput = fsoObj.BuildPath(PathOutput, ExpNom & "_" & ProfFact & "_" & ExaType)

'Sauvegarde du fichier Word
wDoc.SaveAs2 FileName:=PathOutput & "_Facture.docx"

'Exports des PDF si les fichiers ne sont pas ouverts
If IsFileOpen(PathOutput & "_Facture.pdf") = False Then
    'Exports PDF
    wDoc.ExportAsFixedFormat PathOutput & "_Facture.pdf", wdExportFormatPDF, True
    FactureOK = "OK"
Else
    'Fichier ouvert
    MsgBox "XXX_Facture.pdf est ouvert!" & vbCrLf & "Fermer le fichier et recommencez", vbCritical
    FactureOK = "non g�n�r�e"
End If

wordApp.Quit (False)

If IsFileOpen(PathOutput & "_Annexe.pdf") = False Then
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, FileName:=PathOutput & "_Annexe.pdf", OpenAfterPublish:=True
    AnnexeOK = "OK"
Else
    MsgBox "XXX_Annexe.pdf est ouvert!" & vbCrLf & "Fermer le fichier et recommencez", vbCritical
    AnnexeOK = "non g�n�r�e"
End If


MsgBox " - Facture PDF " & FactureOK & vbCrLf & " - Annexe PDF " & AnnexeOK & vbCrLf & vbCrLf & "N'oubliez pas de signer la facture!", vbInformation, "Termin�"

'Sortie rapide si pas la bonne version voir ligne 14
SortieRapide:

End Sub


'Verification de l'�tat d'ouverture du fichier
Function IsFileOpen(FileName As String)

Dim fileNum As Integer
Dim errNum As Integer
Dim strFichierExiste As String
strFichierExiste = Dir(FileName)

'Allow all errors to happen
On Error Resume Next
fileNum = FreeFile()

'Try to open and close the file for input.
'Errors mean the file is already open
If strFichierExiste <> "" Then
    Open FileName For Input Lock Read As #fileNum
End If
Close fileNum

'Get the error number
errNum = Err

'Do not allow errors to happen
On Error GoTo 0

'Check the Error Number
Select Case errNum

    'errNum = 0 means no errors, therefore file closed
    Case 0
    IsFileOpen = False
 
    'errNum = 70 means the file is already open
    Case 70
    IsFileOpen = True

    'Something else went wrong
    Case Else
    IsFileOpen = errNum

End Select

End Function


'Test the Office application version
'Written by Ken Puls (www.excelguru.ca)
Function AppVersion() As Long

Dim registryObject As Object
Dim rootDirectory As String
Dim keyPath As String
Dim arrEntryNames As Variant
Dim arrValueTypes As Variant
Dim x As Long

Select Case Val(Application.Version)
Case Is = 16
'Check for existence of Licensing key
keyPath = "Software\Microsoft\Office\" & CStr(Application.Version) & "\Common\Licensing\LicensingNext"
rootDirectory = "."
Set registryObject = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & rootDirectory & "\root\default:StdRegProv")
registryObject.EnumValues &H80000001, keyPath, arrEntryNames, arrValueTypes

On Error GoTo ErrorExit
For x = 0 To UBound(arrEntryNames)
If InStr(arrEntryNames(x), "365") > 0 Then
AppVersion = 365
Exit Function
End If
If InStr(arrEntryNames(x), "2019") > 0 Then
AppVersion = 2019
Exit Function
End If
If InStr(arrEntryNames(x), "2021") > 0 Then
AppVersion = 2021
Exit Function
End If
Next x
Case Is = 15
AppVersion = 2013
Case Is = 14
AppVersion = 2010
Case Is = 12
AppVersion = 2007
Case Else
'Too old to bother with
AppVersion = 0
End Select

Exit Function

ErrorExit:
'Version 16, but no licensing key. Must be Office 2016
AppVersion = 2016

End Function

'Ecriture dans le word
Function WordWrite(Document As Word.Document, Zone As String, Valeur As String, Optional Unit As String)

If Valeur <> "0" Then
    Dim rng_ToWrite As Word.Range
    Set rng_ToWrite = Document.Bookmarks(Zone).Range
    rng_ToWrite.Text = Valeur & " " & Unit
End If

End Function




