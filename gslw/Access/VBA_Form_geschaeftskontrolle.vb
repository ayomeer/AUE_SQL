
' Hi



'╔══════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════╗
'║ VBA Modul: Form_geschaeftskontrolle																																			║
'╠══════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════╣
'║ 																																												║	
'║ Author: 				Andreas Kuhn (Zivildienstleistender August - Januar 2016)																								║
'║ email privat:		andreas.kuhn@bluewin.ch																																	║
'║ Archiviert unter:	M:\abt_umwelt_energie\gewaesserschutz\jauchegrube\2_DB																									║
'║ 																																												║
'║ Beschreibung:																																								║
'║ Beinhaltet alle Event Handler für die Controls im  Form 'geschaeftskontrolle'. Hilfsunktionen sind in dem separaten Modul "tools" definiert.									║
'║ 																																												║
'║ Good to know:																																								║
'║ - Beim Ändern der Datenquellen von oder zu zustelladresse ist zu beachten, dass .zustelladresse die einzige Tabelle ist, die ortschaft als "bort" bezeichnet. 				║
'║   Entsprechend müssen die Filterprozeduren angepasst werden.																													║
'║ - .ListCount (die ListBox Property, welche die Anzahl Reihen in der ListBox wiedergibt) zählt die Header Row mit, wenn diese aktiviert ist. Allerding nur wenn 				║
'║   die ListBox nicht leer ist.																																				║
'║ - SVA steht im Zusammenhang mit diesem Formular für Standortvoranfrage.																										║
'║ - Wenn eine Control und ihr Event Handler umbenennt werden, müssen die Event Handler neu verknüpft werden, in dem man im Eigenschaftenblatt der Control nochmals				║
'║   beim Entsprechenden Event "code" wählt																																		║
'║ - VBA unterstützt kein line wrapping, dh ein Statement kann nicht ohne weiteres auf mehrere Zeilen aufgeteilt werden. Es steht aber ein Syntax zur verfügung um 				║
'║   dies zu erreichen "&_"  (Siehe Konstanten Definitionen)																													║
'║																																												║
'║ Verbesserungspotential:																																						║
'║ - Code Redudanz: Mehr in Subroutienen/Funktionen abpacken, wobei bei vielen Prozeduren genug anders ist, um dies in Frage zu stellen.										║
'║ - Wenn man einen Weg findet um ListeJauchegruben mit anderen Daten zu füllen, wenn sich die Selektion in ListeBewirtschafter ändert, wäre das besseres Design und			║
'║   man könnte sich dann auch einen Event Handler pro Tab sparen.																												║
'║ - Views besser benennen																																						║
'║																																												║
'║ Links Microsoft Dokumentation:																																				║
'║ - For...Next Statement: 	https://msdn.microsoft.com/en-us/library/5z06z1kb.aspx																								║
'║ - ListBox Object:		https://msdn.microsoft.com/en-us/library/office/ff195480.aspx																						║
'║																																												║
'╚══════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════╝

Option Compare Database

'---- Query Konstanten ----'
' Da diese Querydefinitionen durch das ganze Modul verwendet werden sind sie hier, in "Bausteine" aufgeteilt, als Konstanten definiert.

'--Union Query "Abklärungsphase"--'
' Dies sind die beiden Teilqueries, mit denen in den Prozeduren weiter unten das Union Query zwischen view_bewinfo_geb und view_bew_ohne_geb gebaut wird.
' Das Union Query erlaubt das dynamische bestimmen der Ortschaft der Bewirtschafter anhand von Gebäudedaten, ohne die Bewirtschafter ohne Gebäude auszuschliessen.
Private Const SELECT_UNION_ABKLAERUNGSPHASE_GEB As String = "SELECT zustellid, MIN(ablagenr), MIN(name), MIN(vorname), MIN(ortschaft), MIN(versand_voranfrage), " & _
                                                            "MIN(antwort_voranfrage), MIN(versand_aufgebot), MIN(frist_aufgebot), MIN(id_zustelladresse)"
Private Const FROM_UNION_GEB As String = "FROM dbu_aue_gslw_view_bewinfo_geb "

Private Const SELECT_UNION_ABKLAERUNGSPHASE_BEW As String = "SELECT zustellid, ablagenr, name, vorname, ortschaft, versand_voranfrage As v_voranfrage, " & _
                                                            "antwort_voranfrage As a_voranfrage, versand_aufgebot As v_aufgebot, frist_aufgebot As f_aufgebot, id_zustelladresse "
Private Const FROM_UNION_BEW As String = "FROM dbu_aue_gslw_view_bew_ohne_geb "

'--view_geschaeftskontrolle_jauchegruben--'
' View basierend auf view_bewinfo_jg welche die Gruben nach Bewirtschafter gruppiert und die Anzahl Gruben in den wichtigsten Kontrollstati zählt.
Private Const SELECT_BEWINFO_JG_JOINED As String = "SELECT dbu_aue_gslw_view_bewinfo_jg_grouped.zustellid, " & _
                                            "dbu_aue_gslw_view_bewinfo_jg_grouped.ablagenr, " & _
                                            "dbu_aue_gslw_view_bewinfo_jg_grouped.name, " & _
                                            "dbu_aue_gslw_view_bewinfo_jg_grouped.vorname, " & _
                                            "dbu_aue_gslw_view_bewinfo_jg_grouped.ortschaft, " & _
                                            "dbu_aue_gslw_view_bewinfo_jg_grouped.versand_voranfrage As v_voranfrage, " & _
                                            "dbu_aue_gslw_view_bewinfo_jg_grouped.antwort_voranfrage As a_voranfrage, " & _
                                            "dbu_aue_gslw_view_bewinfo_jg_grouped.versand_aufgebot As v_aufgebot, " & _
                                            "dbu_aue_gslw_view_bewinfo_jg_grouped.frist_aufgebot As f_aufgebot, " & _
                                            "dbu_aue_gslw_view_bewinfo_jg_grouped.tage_bis_frist, " & _
                                            "dbu_aue_gslw_view_bewinfo_jg_grouped.anz_io, " & _
                                            "dbu_aue_gslw_view_bewinfo_jg_grouped.anz_kontrolliert_mbr, " & _
                                            "dbu_aue_gslw_view_bewinfo_jg_grouped.anz_stillgelegt, " & _
                                            "dbu_aue_gslw_view_bewinfo_jg_grouped.anz_gruben, " & _
                                            "dbu_aue_gslw_view_bewinfo_jg_grouped.id_zustelladresse "

Private Const FROM_VIEW_GK_JG As String = "FROM dbu_aue_gslw_view_bewinfo_jg_grouped "

'--gebaeude--'
' RowSource für Sekundärlistenfeld ListeGebaeude_SVAgesendet
Private Const SELECT_GEBAEUDE As String = "SELECT id_geb, bezeichnung, ext_lbnr As lbnr, ext_grundstck As parz, ext_ortschaft As ortschaft, ext_flurname As flurbez, alpname, bemerkungen "
Private Const FROM_GEBAEUDE As String = "FROM dbu_aue_gslw_gebaeude "

'--jauchegruben--'
Private Const SELECT_JAUCHEGRUBE As String = "SELECT jgstatus As status, ext_ortschaft As ortschaft, flurbez, ext_lbnr As lbnr, ext_grundstck As parz, nutzvolumen As vol, " & _
                                             "baujahr, ext_gschzone As SZ, letztekontr, kontr_status_frist As frist, kontr_grund As kontrollgrund, kontr_status As kontrollstatus, bemerkung "
Private Const FROM_JAUCHEGRUBE As String = "FROM dbu_aue_gslw_jauchegrube "

'--WHERE Filter (Definieren welche Records im jeweiligen Tab gelistet werden)--'
Private Const WHERE_SVA_GESENDET As String = "WHERE ((versand_voranfrage >#1111-11-11#) AND (antwort_voranfrage IS NULL) AND (versand_aufgebot IS NULL)) "
Private Const WHERE_SVA_ANTWORT As String = "WHERE ((antwort_voranfrage >#1111-11-11#) AND (versand_aufgebot IS NULL) AND anz_io < (anz_gruben - anz_stillgelegt)) "
Private Const WHERE_AUFGEBOT As String = "WHERE anz_aufgeboten > 0 "
Private Const WHERE_ALPEN As String = "WHERE (zustellid LIKE '%/50/%') "
Private Const WHERE_KONTROLLIERT As String = "WHERE (anz_io < anz_gruben-anz_stillgelegt) AND (anz_kontrolliert_mbr > 0) "
Private Const WHERE_ABGESCHLOSSEN As String = "WHERE (anz_io >= anz_gruben-anz_stillgelegt) AND (anz_kontrolliert_mbr > 0) "
'--ORDER BY Klauseln--'
Private Const ORDER_BY_DATES As String = "ORDER BY versand_aufgebot, versand_voranfrage "
Private Const ORDER_ALIAS_DATES As String = "ORDER BY v_aufgebot DESC, v_voranfrage DESC "
Private Const ORDER_GROUPED As String = "ORDER BY zustellid "
Private Const ORDER_AUFGEBOT As String = "ORDER BY tage_bis_frist ASC "
Private Const ORDER_PRZ_MBR As String = "ORDER BY prz_kontrolliert_mbr DESC "

Private Const GROUP As String = "GROUP BY zustellid "


Private Sub Form_Open(Cancel As Integer)
    'Initialisieren der Controls beim Oeffnen des Forms
    
    '---- Set ListBox RowSource to default values ----'
    Me.ListeBewirtschafter_Alle.RowSource = SELECT_UNION_ABKLAERUNGSPHASE_BEW & FROM_UNION_BEW & "UNION " & SELECT_UNION_ABKLAERUNGSPHASE_GEB & FROM_UNION_GEB & GROUP & ORDER_ALIAS_DATES
    Me.ListeBewirtschafter_SVAgesendet.RowSource = SELECT_UNION_ABKLAERUNGSPHASE_BEW & FROM_UNION_BEW & WHERE_SVA_GESENDET & "UNION " & SELECT_UNION_ABKLAERUNGSPHASE_GEB & FROM_UNION_GEB & WHERE_SVA_GESENDET & GROUP & ORDER_ALIAS_DATES
    Me.ListeBewirtschafter_SVAantwort.RowSource = SELECT_BEWINFO_JG_JOINED & FROM_VIEW_GK_JG & WHERE_SVA_ANTWORT & ORDER_AUFGEBOT
    'SELECT_UNION_ABKLAERUNGSPHASE_BEW & FROM_UNION_BEW & WHERE_SVA_ANTWORT & "UNION " & SELECT_UNION_ABKLAERUNGSPHASE_GEB & FROM_UNION_GEB & WHERE_SVA_ANTWORT & GROUP & ORDER_ALIAS_DATES
    Me.ListeBewirtschafter_Aufgebot.RowSource = SELECT_BEWINFO_JG_JOINED & FROM_VIEW_GK_JG & WHERE_AUFGEBOT & ORDER_AUFGEBOT
    Me.ListeBewirtschafter_Kontrolliert.RowSource = SELECT_BEWINFO_JG_JOINED & FROM_VIEW_GK_JG & WHERE_KONTROLLIERT & ORDER_GROUPED
    Me.ListeBewirtschafter_Abgeschlossen.RowSource = SELECT_BEWINFO_JG_JOINED & FROM_VIEW_GK_JG & WHERE_ABGESCHLOSSEN & ORDER_BY_DATES
    
    '---- Clear jg list ----'
    Me.ListeJauchegruben_Alle.RowSource = ""
    Me.ListeGebaeude_SVAgesendet.RowSource = ""
    Me.ListeJauchegruben_SVAantwort.RowSource = ""
    Me.ListeJauchegruben_Aufgebot.RowSource = ""
    Me.ListeJauchegruben_Kontrolliert.RowSource = ""
    Me.ListeJauchegruben_Abgeschlossen.RowSource = ""

    '---- Update listcount values ----'
    'ListCount zaehlt die Header Row mit(ausser wenn ListBox leer), deshalb mit "-1" angepasst.
    TabAll.Caption = "Alle Bewirtschafter (" & ListeBewirtschafter_Alle.ListCount - 1 & ")"
    TabSVAgesendet.Caption = "Voranfrage offen (" & ListeBewirtschafter_SVAgesendet.ListCount - 1 & ")"
    TabSVAantwort.Caption = "Voranfrage beantwortet (" & ListeBewirtschafter_SVAantwort.ListCount - 1 & ")"
    TabAufgebot.Caption = "Aufgebot offen (" & ListeBewirtschafter_Aufgebot.ListCount - 1 & ")"
    TabKontrolliert.Caption = "Kontrolliert, weitere Massnahmen (" & ListeBewirtschafter_Kontrolliert.ListCount - 1 & ")"
    TabAbgeschlossen.Caption = "Abgeschlossen (" & ListeBewirtschafter_Abgeschlossen.ListCount - 1 & ")"
    
    '---- Clear dynamic list count labels ----'
    lblListCount_All.Caption = ""
    lblListCount_SVAgesendet.Caption = ""
    lblListCount_SVAantwort.Caption = ""
    lblListCount_Aufgebot.Caption = ""
    lblListCount_Kontrolliert.Caption = ""
    lblListCount_Abgeschlossen.Caption = ""
    
    Me.ListeBewirtschafter_Alle.SetFocus
    Me.ListeBewirtschafter_Alle.ListIndex = 0
    
End Sub

'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Excel Export Button Event Handlers
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Beschreibung und Definition der Funktion in 'tools' Modul

Private Sub ButtonExportToExcel_Alle_Click()
    Call tools.ExportToExcel(ListeBewirtschafter_Alle)
End Sub

Private Sub ButtonExportToExcel_SVAgesendet_Click()
    Call tools.ExportToExcel(ListeBewirtschafter_SVAgesendet)
End Sub

Private Sub ButtonExportToExcel_SVAantwort_Click()
    Call tools.ExportToExcel(ListeBewirtschafter_SVAantwort)
End Sub

Private Sub ButtonExportToExcel_Aufgebot_Click()
    Call tools.ExportToExcel(ListeBewirtschafter_Aufgebot)
End Sub

Private Sub ButtonExportToExcel_Kontrolliert_Click()
    Call tools.ExportToExcel(ListeBewirtschafter_Kontrolliert)
End Sub

Private Sub ButtonExportToExcel_Abgeschlossen_Click()
    Call tools.ExportToExcel(ListeBewirtschafter_Abgeschlossen)
End Sub

'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Bewirtschafter Info Button Event Handlers
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Beschreibung:
' Listboxen haben keine Property, die direkt den Index der selektierten Row zurück gibt. Mit ListBox.Selected(lRow) kann man aber fragen, ob
' die row mit dem mitgegebenen Index lRow selektiert ist oder nicht. Mann muss dies also für jede Row machen, bis man die Selektierte
' gefunden hat. Dann kann mit ListBox.Column(Column, Row) die zustellid gelesen werden, damit Sie beim Öffnen des bewirtschafter_info Formulars
' als Filterargument mitgegeben werden kann.
'
' Caveats:
' - Die Index der Rows beginnen bei 0, deshalb muss in der 'To' Definition der For-Loops ListCount - 1 gerechnet werden.
' - Beim zusammensetzen von RowSource Queries muss man sich immer bewusst sein, dass dieses in Form eines Strings in die Property geschrieben wird
  ' und dass nur der Inhalt einer Variable oder Property genutzt wird (mit ' umschliessen), da der Kontext für die Variable verloren geht.

  
Private Sub ButtonBewInfo_Alle_Click()
    Dim i As Integer
    Dim strWHERE As String
    
    For i = 0 To ListeBewirtschafter_Alle.ListCount - 1
       If ListeBewirtschafter_Alle.Selected(i) Then
            strWHERE = "zustellid =" & "'" & ListeBewirtschafter_Alle.Column(0, i) & "'" 
            Exit For
       End If
    Next i
    
    If CurrentProject.AllForms("bewirtschafter_info").IsLoaded Then DoCmd.Close acForm, "bewirtschafter_info", acSaveYes
    DoCmd.OpenForm "bewirtschafter_info", acViewNormal, , strWHERE
End Sub

Private Sub ButtonBewInfo_SVAgesendet_Click()
    Dim i As Integer
    Dim strWHERE As String
    
    For i = 0 To ListeBewirtschafter_SVAgesendet.ListCount - 1
       If ListeBewirtschafter_SVAgesendet.Selected(i) Then
            strWHERE = "zustellid =" & "'" & ListeBewirtschafter_SVAgesendet.Column(0, i) & "'" 
            Exit For
       End If
    Next i
    If CurrentProject.AllForms("bewirtschafter_info").IsLoaded Then DoCmd.Close acForm, "bewirtschafter_info", acSaveYes
    DoCmd.OpenForm "bewirtschafter_info", acViewNormal, , strWHERE
End Sub

Private Sub ButtonBewInfo_SVAantwort_Click()
    Dim i As Integer
    Dim strWHERE As String
    
    For i = 0 To ListeBewirtschafter_SVAantwort.ListCount - 1
       If ListeBewirtschafter_SVAantwort.Selected(i) Then
            strWHERE = "zustellid =" & "'" & ListeBewirtschafter_SVAantwort.Column(0, i) & "'" 
            Exit For
       End If
    Next i
    If CurrentProject.AllForms("bewirtschafter_info").IsLoaded Then DoCmd.Close acForm, "bewirtschafter_info", acSaveYes
    DoCmd.OpenForm "bewirtschafter_info", acViewNormal, , strWHERE
End Sub

Private Sub ButtonBewInfo_Aufgebot_Click()
    Dim i As Integer
    Dim strWHERE As String
    
    For i = 0 To ListeBewirtschafter_Aufgebot.ListCount - 1
       If ListeBewirtschafter_Aufgebot.Selected(i) Then
            strWHERE = "zustellid =" & "'" & ListeBewirtschafter_Aufgebot.Column(0, i) & "'" 
            Exit For
       End If
    Next i
    If CurrentProject.AllForms("bewirtschafter_info").IsLoaded Then DoCmd.Close acForm, "bewirtschafter_info", acSaveYes
    DoCmd.OpenForm "bewirtschafter_info", acViewNormal, , strWHERE
End Sub

Private Sub ButtonBewInfo_Kontrolliert_Click()
    Dim i As Integer
    Dim strWHERE As String
    
    For i = 0 To ListeBewirtschafter_Kontrolliert.ListCount - 1
       If ListeBewirtschafter_Kontrolliert.Selected(i) Then
            strWHERE = "zustellid =" & "'" & ListeBewirtschafter_Kontrolliert.Column(0, i) & "'" 
            Exit For
       End If
    Next i
    If CurrentProject.AllForms("bewirtschafter_info").IsLoaded Then DoCmd.Close acForm, "bewirtschafter_info", acSaveYes
    DoCmd.OpenForm "bewirtschafter_info", acViewNormal, , strWHERE
End Sub

Private Sub ButtonBewInfo_Abgeschlossen_Click()
    Dim i As Integer
    Dim strWHERE As String
    
    For i = 0 To ListeBewirtschafter_Abgeschlossen.ListCount - 1
       If ListeBewirtschafter_Abgeschlossen.Selected(i) Then
            strWHERE = "zustellid =" & "'" & ListeBewirtschafter_Abgeschlossen.Column(0, i) & "'"
            Exit For
       End If
    Next i
    If CurrentProject.AllForms("bewirtschafter_info").IsLoaded Then DoCmd.Close acForm, "bewirtschafter_info", acSaveYes 'Zum verhindern von Fehlern
    DoCmd.OpenForm "bewirtschafter_info", acViewNormal, , strWHERE
End Sub

'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' ListBox Double Click Event Handlers
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Beschreibung:
' Doppelklick auf Bewirtschafter record ändert RowSource von ListeJauchegruben_Alle um die dem Bewirtschafter zugehörigen Jauchegruben anzuzeigen.
' Um die zustellid des geklickten Bewirtschafters zu finden geht die Prozedur jeden Eintrag der Listbox durch, und prüft ob dieser "Selected" ist.
' Da nach nur eine Selektion zulaessig ist, wird die For... Next Schlaufe nach dem Finden des Eintrags verlassen.
 
' ListCount beginnt bei 1 und der ListBox index bei 0, deshalb wird Endkondition mit "-1" angepasst.
            
Private Sub ListeBewirtschafter_Alle_DblClick(Cancel As Integer)
    Dim i As Integer
    Dim strRowSource As String
    strRowSource = SELECT_JAUCHEGRUBE & FROM_JAUCHEGRUBE & "WHERE help_kt_id = "
    
    For i = 0 To ListeBewirtschafter_Alle.ListCount - 1
       If ListeBewirtschafter_Alle.Selected(i) Then
            strRowSource = strRowSource & "'" & ListeBewirtschafter_Alle.Column(0, i) & "'"
            Exit For
       End If
    Next i
    ListeJauchegruben_Alle.RowSource = strRowSource
End Sub

Private Sub ListeBewirtschafter_SVAgesendet_DblClick(Cancel As Integer)
    Dim i As Integer
    Dim strRowSource As String
    
    strRowSource = SELECT_GEBAEUDE & FROM_GEBAEUDE & "WHERE help_bew_id_gebaeude = "
    
    For i = 0 To ListeBewirtschafter_SVAgesendet.ListCount - 1
       If ListeBewirtschafter_SVAgesendet.Selected(i) Then
            strRowSource = strRowSource & "'" & ListeBewirtschafter_SVAgesendet.Column(0, i) & "'"
            Exit For
       End If
    Next i
    ListeGebaeude_SVAgesendet.RowSource = strRowSource
End Sub

Private Sub ListeBewirtschafter_SVAantwort_DblClick(Cancel As Integer)
    Dim i As Integer
    Dim strRowSource As String
    
    strRowSource = SELECT_JAUCHEGRUBE & FROM_JAUCHEGRUBE & "WHERE help_kt_id = "
    
    For i = 0 To ListeBewirtschafter_SVAantwort.ListCount - 1
       If ListeBewirtschafter_SVAantwort.Selected(i) Then
            strRowSource = strRowSource & "'" & ListeBewirtschafter_SVAantwort.Column(0, i) & "'"
            Exit For
       End If
    Next i
    ListeJauchegruben_SVAantwort.RowSource = strRowSource
End Sub

Private Sub ListeBewirtschafter_Aufgebot_DblClick(Cancel As Integer)
    Dim i As Integer
    Dim strRowSource As String
    
    strRowSource = SELECT_JAUCHEGRUBE & FROM_JAUCHEGRUBE & "WHERE help_kt_id = "
    
    For i = 0 To ListeBewirtschafter_Aufgebot.ListCount - 1
       If ListeBewirtschafter_Aufgebot.Selected(i) Then
            strRowSource = strRowSource & "'" & ListeBewirtschafter_Aufgebot.Column(0, i) & "'"
            Exit For
       End If
    Next i
    ListeJauchegruben_Aufgebot.RowSource = strRowSource
End Sub

Private Sub ListeBewirtschafter_Kontrolliert_DblClick(Cancel As Integer)
    Dim i As Integer
    Dim strRowSource As String
    
    strRowSource = SELECT_JAUCHEGRUBE & FROM_JAUCHEGRUBE & "WHERE help_kt_id = "
    
    For i = 0 To ListeBewirtschafter_Kontrolliert.ListCount - 1
       If ListeBewirtschafter_Kontrolliert.Selected(i) Then
            strRowSource = strRowSource & "'" & ListeBewirtschafter_Kontrolliert.Column(0, i) & "'"
            Exit For
       End If
    Next i
    ListeJauchegruben_Kontrolliert.RowSource = strRowSource
End Sub

Private Sub ListeBewirtschafter_Abgeschlossen_DblClick(Cancel As Integer)
    Dim i As Integer
    Dim strRowSource As String
    
    strRowSource = SELECT_JAUCHEGRUBE & FROM_JAUCHEGRUBE & "WHERE help_kt_id = "
    
    For i = 0 To ListeBewirtschafter_Abgeschlossen.ListCount - 1
       If ListeBewirtschafter_Abgeschlossen.Selected(i) Then
            strRowSource = strRowSource & "'" & ListeBewirtschafter_Abgeschlossen.Column(0, i) & "'"
            Exit For
       End If
    Next i
    ListeJauchegruben_Abgeschlossen.RowSource = strRowSource
End Sub

'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' ListBox KeyUp Event Handlers
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Beschreibung:
' Aktualisiert ListeJauchegruben_Alle wenn selektion mit Cursor Tasten geändert wird.
' Nutzt selbe Prozedur wie ListeBewirtschafter_Alle_DblClick.

' Caveats:
' Es wurde KeyUp anstelle der intuitiveren KeyDown oder KeyPress Events verwendet, da KeyPress die Cursor Keys nicht unterstuetzt und KeyDown
' ausgefuehrt wird, bevor die Selektion aktualisiert wird.

Private Sub ListeBewirtschafter_Alle_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = vbKeyDown Or KeyCode = vbKeyUp) Then ListeBewirtschafter_Alle_DblClick (False)
End Sub

Private Sub ListeBewirtschafter_SVAgesendet_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = vbKeyDown Or KeyCode = vbKeyUp) Then ListeBewirtschafter_SVAgesendet_DblClick (False)
End Sub

Private Sub ListeBewirtschafter_SVAantwort_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = vbKeyDown Or KeyCode = vbKeyUp) Then ListeBewirtschafter_SVAantwort_DblClick (False)
End Sub

Private Sub ListeBewirtschafter_Aufgebot_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = vbKeyDown Or KeyCode = vbKeyUp) Then ListeBewirtschafter_Aufgebot_DblClick (False)
End Sub

Private Sub ListeBewirtschafter_Kontrolliert_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = vbKeyDown Or KeyCode = vbKeyUp) Then ListeBewirtschafter_Kontrolliert_DblClick (False)
End Sub

Private Sub ListeBewirtschafter_Abgeschlossen_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = vbKeyDown Or KeyCode = vbKeyUp) Then ListeBewirtschafter_Abgeschlossen_DblClick (False)
End Sub

'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Filter TextBox Change Event Handlers
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Beschreibung:
' Textbox zum Filtern der gelisteten Eintraege.
' Filterkriterien von ListeBewirtschafter werden mit jeder Änderung des Textbox Inhalts aktualisiert Damit schon bevor der ganze Suchbegriff eingegeben wurde
' Ergebnisse erscheinen wird das Change Event und das LIKE Statement zusammen mit Wildcard Charaktern (% in SQL) genutzt.

' Die Textboxeingabe filtert sowohl in zustellid als auch in ortschaft/bort, allerdings nicht beides gleichzeitig (OR verknüpft).

' Caveats:
' - Die .Text Property einer TextBox Control kann nur verwendet werden, wenn diese derzeit den Fokus hat.
'   Wenn dies nicht garantiert werden kann sollte .Value genutzt werden.

Private Sub tbFilter_Alle_Change()
    Dim strFilterArgs As String
    Dim strTBconv As String
    strTBconv = LCase(tbFilter_Alle.Text) 'Sucheingabe in lowercase konvertieren und mit dem resultierenden String SQL-Filterargumente bilden (nächste Linie)
    strFilterArgs = "WHERE lcase(ortschaft) LIKE '%" & strTBconv & "%' OR lcase(zustellid) LIKE '%" & strTBconv & "%' OR lcase(name) LIKE '%" & strTBconv & "%'"
    
    If (tbFilter_Alle.Text = "") Then
        Me.ListeBewirtschafter_Alle.RowSource = SELECT_UNION_ABKLAERUNGSPHASE_BEW & FROM_UNION_BEW & "UNION " & SELECT_UNION_ABKLAERUNGSPHASE_GEB & FROM_UNION_GEB & GROUP & ORDER_ALIAS_DATES
        lblListCount_All.Caption = ""
    Else
        Me.ListeBewirtschafter_Alle.RowSource = SELECT_UNION_ABKLAERUNGSPHASE_BEW & FROM_UNION_BEW & strFilterArgs & "UNION " & SELECT_UNION_ABKLAERUNGSPHASE_GEB & FROM_UNION_GEB & strFilterArgs & GROUP & ORDER_ALIAS_DATES
        lblListCount_All.Caption = "(" & ListeBewirtschafter_Alle.ListCount - 1 & ")"
    End If
End Sub

Private Sub tbFilter_SVAgesendet_Change()
    Dim strFilterArgs As String
    Dim strTBconv As String
    strTBconv = LCase(tbFilter_SVAgesendet.Text) 'Sucheingabe in lowercase konvertieren und mit dem resultierenden String SQL-Filterargumente bilden (nächste Linie)
    strFilterArgs = "AND ( lcase (ortschaft) LIKE '%" & strTBconv & "%' OR lcase(zustellid) LIKE '%" & strTBconv & "%' OR lcase(name) LIKE '%" & strTBconv & "%')"
    
    If (tbFilter_SVAgesendet.Text = "") Then
        Me.ListeBewirtschafter_SVAgesendet.RowSource = SELECT_UNION_ABKLAERUNGSPHASE_BEW & FROM_UNION_BEW & WHERE_SVA_GESENDET & "UNION " & SELECT_UNION_ABKLAERUNGSPHASE_GEB & FROM_UNION_GEB & WHERE_SVA_GESENDET & GROUP & ORDER_ALIAS_DATES
        lblListCount_SVAgesendet.Caption = ""
    Else
        Me.ListeBewirtschafter_SVAgesendet.RowSource = SELECT_UNION_ABKLAERUNGSPHASE_BEW & FROM_UNION_BEW & WHERE_SVA_GESENDET & strFilterArgs & "UNION " & SELECT_UNION_ABKLAERUNGSPHASE_GEB & FROM_UNION_GEB & WHERE_SVA_GESENDET & strFilterArgs & GROUP & ORDER_ALIAS_DATES
        lblListCount_SVAgesendet.Caption = "(" & ListeBewirtschafter_SVAgesendet.ListCount - 1 & ")"
    End If
End Sub

Private Sub tbFilter_SVAantwort_Change()
    Dim strFilterArgs As String
    Dim strTBconv As String
    strTBconv = LCase(tbFilter_SVAantwort.Text) 'Sucheingabe in lowercase konvertieren und mit dem resultierenden String SQL-Filterargumente bilden (nächste Linie)
    strFilterArgs = "AND ( lcase (ortschaft) LIKE '%" & strTBconv & "%' OR lcase(zustellid) LIKE '%" & strTBconv & "%' OR lcase(name) LIKE '%" & strTBconv & "%')"

    If (tbFilter_SVAantwort.Text = "") Then
        Me.ListeBewirtschafter_SVAantwort.RowSource = SELECT_BEWINFO_JG_JOINED & FROM_VIEW_GK_JG & WHERE_SVA_ANTWORT & ORDER_AUFGEBOT
        'SELECT_UNION_ABKLAERUNGSPHASE_BEW & FROM_UNION_BEW & WHERE_SVA_ANTWORT & "UNION " & SELECT_UNION_ABKLAERUNGSPHASE_GEB & FROM_UNION_GEB & WHERE_SVA_ANTWORT & GROUP & ORDER_ALIAS_DATES
        lblListCount_SVAantwort.Caption = ""
    Else
        Me.ListeBewirtschafter_SVAantwort.RowSource = SELECT_BEWINFO_JG_JOINED & FROM_VIEW_GK_JG & WHERE_SVA_ANTWORT & strFilterArgs & ORDER_AUFGEBOT
        'SELECT_UNION_ABKLAERUNGSPHASE_BEW & FROM_UNION_BEW & WHERE_SVA_ANTWORT & strFilterArgs & "UNION " & SELECT_UNION_ABKLAERUNGSPHASE_GEB & FROM_UNION_GEB & WHERE_SVA_ANTWORT & strFilterArgs & GROUP & ORDER_ALIAS_DATES
        lblListCount_SVAantwort.Caption = "(" & ListeBewirtschafter_SVAantwort.ListCount - 1 & ")"
    End If
End Sub

Private Sub tbFilter_Aufgebot_Change()
    Dim FilterArg_Name_Aufgebot As String
    Dim strTBconv As String
    strTBconv = LCase(tbFilter_Aufgebot.Text) 'Sucheingabe in lowercase konvertieren und mit dem resultierenden String SQL-Filterargumente bilden (nächste Linie)
    strFilterArgs = "AND ( lcase (ortschaft) LIKE '%" & strTBconv & "%' OR lcase(zustellid) LIKE '%" & strTBconv & "%' OR lcase(name) LIKE '%" & strTBconv & "%')"

    If (tbFilter_Aufgebot.Text = "") Then
        Me.ListeBewirtschafter_Aufgebot.RowSource = SELECT_BEWINFO_JG_JOINED & FROM_VIEW_GK_JG & WHERE_AUFGEBOT & ORDER_AUFGEBOT
        lblListCount_Aufgebot.Caption = ""
    Else
        Me.ListeBewirtschafter_Aufgebot.RowSource = SELECT_BEWINFO_JG_JOINED & FROM_VIEW_GK_JG & WHERE_AUFGEBOT & strFilterArgs & ORDER_AUFGEBOT
        lblListCount_Aufgebot.Caption = "(" & ListeBewirtschafter_Aufgebot.ListCount - 1 & ")"
    End If
    
End Sub

Private Sub tbFilter_Kontrolliert_Change()
    Dim FilterArg_Name_Kontrolliert As String
    Dim strTBconv As String
    strTBconv = LCase(tbFilter_Kontrolliert.Text) 'Sucheingabe in lowercase konvertieren und mit dem resultierenden String SQL-Filterargumente bilden (nächste Linie)
    strFilterArgs = "AND ( lcase (ortschaft) LIKE '%" & strTBconv & "%' OR lcase(zustellid) LIKE '%" & strTBconv & "%' OR lcase(name) LIKE '%" & strTBconv & "%')"

    If (tbFilter_Kontrolliert.Text = "") Then
        Me.ListeBewirtschafter_Kontrolliert.RowSource = SELECT_BEWINFO_JG_JOINED & FROM_VIEW_GK_JG & WHERE_KONTROLLIERT & ORDER_GROUPED
        lblListCount_Kontrolliert.Caption = ""
    Else
        Me.ListeBewirtschafter_Kontrolliert.RowSource = SELECT_BEWINFO_JG_JOINED & FROM_VIEW_GK_JG & WHERE_KONTROLLIERT & strFilterArgs & ORDER_GROUPED
        lblListCount_Kontrolliert.Caption = "(" & ListeBewirtschafter_Kontrolliert.ListCount - 1 & ")"
    End If
    
End Sub

Private Sub tbFilter_Abgeschlossen_Change()
    Dim FilterArg_Name_Abgeschlossen As String
    Dim strTBconv As String
    strTBconv = LCase(tbFilter_Abgeschlossen.Text) 'Sucheingabe in lowercase konvertieren und mit dem resultierenden String SQL-Filterargumente bilden (nächste Linie)
    strFilterArgs = "AND ( lcase (ortschaft) LIKE '%" & strTBconv & "%' OR lcase(zustellid) LIKE '%" & strTBconv & "%' OR lcase(name) LIKE '%" & strTBconv & "%')"
    
    If (tbFilter_Abgeschlossen.Text = "") Then
        Me.ListeBewirtschafter_Abgeschlossen.RowSource = SELECT_BEWINFO_JG_JOINED & FROM_VIEW_GK_JG & WHERE_ABGESCHLOSSEN & ORDER_BY_DATES
        lblListCount_Abgeschlossen.Caption = ""
    Else
        Me.ListeBewirtschafter_Abgeschlossen.RowSource = SELECT_BEWINFO_JG_JOINED & FROM_VIEW_GK_JG & WHERE_ABGESCHLOSSEN & strFilterArgs & ORDER_BY_DATES
        lblListCount_Abgeschlossen.Caption = "(" & ListeBewirtschafter_Abgeschlossen.ListCount - 1 & ")"
    End If
End Sub

'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Event Handlers für inviduelle Controls
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Private Sub ButtonAnsichtVoranfrage1_Click()
    Dim i As Integer
    Dim strWHERE As String
    
    For i = 0 To ListeBewirtschafter_Alle.ListCount - 1
       If ListeBewirtschafter_Alle.Selected(i) Then
            strWHERE = "zustellid = " & "'" & ListeBewirtschafter_Alle.Column(0, i) & "'"
            Exit For
       End If
    Next i
    DoCmd.OpenReport "erste_voranfrage", acLayout, , strWHERE 'Um direkt zur Druckansicht zu gelangen: acPreview, acNormal druckt direkt aus
End Sub

Private Sub ButtonAnsichtVoranfrage2_Click()
    Dim i As Integer
    Dim strWHERE As String
    
    For i = 0 To ListeBewirtschafter_SVAgesendet.ListCount - 1
       If ListeBewirtschafter_SVAgesendet.Selected(i) Then
            strWHERE = "zustellid = " & "'" & ListeBewirtschafter_SVAgesendet.Column(0, i) & "'"
            Exit For
       End If
    Next i
    DoCmd.OpenReport "zweite_voranfrage", acLayout, , strWHERE 'Um direkt zur Druckansicht zu gelangen: acPreview, acNormal druckt direkt aus
End Sub

Private Sub ButtonAnsichtAufgebot_Click()
    Dim i As Integer
    Dim strWHERE As String
    
    For i = 0 To ListeBewirtschafter_SVAantwort.ListCount - 1
       If ListeBewirtschafter_SVAantwort.Selected(i) Then
            strWHERE = "zustellid = " & "'" & ListeBewirtschafter_SVAantwort.Column(0, i) & "'"
            Exit For
       End If
    Next i
    
    If (InStr(ListeBewirtschafter_SVAantwort.Column(0, i), "/50/")) > 0 Then 'InStr() gibt Position des zweiten Strings im Ersten zurueck, 0 wenn nicht gefunden
        DoCmd.OpenReport "schreiben_aufgebot_alpbetrieb", acLayout, , strWHERE 'Um direkt zur Druckansicht zu gelangen: acPreview, acNormal druckt direkt aus
    Else
        DoCmd.OpenReport "schreiben_aufgebot_talbetrieb", acPreview, , strWHERE 'Um direkt zur Druckansicht zu gelangen: acPreview, acNormal druckt direkt aus
    End If
    
End Sub

'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' PDF-Export Briefe ohne Dialog
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

' Aufgebote
Private Sub buttonPDF_Aufgebot_Click()
    
    Dim i As Integer
    Dim strWHERE As String
    Dim strPath As String
    
    'Gewählten Eintrag in ListBox identifizieren und zustellid entnehmen
    For i = 0 To ListeBewirtschafter_SVAantwort.ListCount - 1
       If ListeBewirtschafter_SVAantwort.Selected(i) Then
            strWHERE = "zustellid = " & "'" & ListeBewirtschafter_SVAantwort.Column(0, i) & "'"
            Exit For
       End If
    Next i
    
    'Nach Tal- oder Alpbetrieb unterschieden entsprechenden Report öffnen, als PDF exportieren und wieder schliessen.
    If (InStr(ListeBewirtschafter_SVAantwort.Column(0, i), "/50/")) > 0 Then 'InStr() gibt Position des zweiten Strings im Ersten zurueck, 0 wenn nicht gefunden
        On Error GoTo Err_Cancel 'Bei Error Sub verlassen, wird ausgelöst wenn manuelle Eingabe vom Alpnamen abgebrochen wird
        DoCmd.OpenReport "schreiben_aufgebot_alpbetrieb", acLayout, , strWHERE, acHidden
        
        strPath = [Forms]![geschaeftskontrolle]![tbPath_Aufgebot].Value & "\" & [Reports]![schreiben_aufgebot_alpbetrieb].Caption & ".pdf"
        On Error GoTo Err_InvalidPath 'Falls beim Ausführen der nächsten Linie ein Fehler auftritt zur "Err_InvalidPath"-Sprungmarke springen. Passiert wenn kein Pfad angegeben.
        DoCmd.OutputTo acOutputReport, "schreiben_aufgebot_alpbetrieb", acFormatPDF, strPath
        
        DoEvents
        DoCmd.Close acReport, "schreiben_aufgebot_alpbetrieb", acSaveNo
    Else
        DoCmd.OpenReport "schreiben_aufgebot_talbetrieb", acLayout, , strWHERE, acHidden
          
        strPath = [Forms]![geschaeftskontrolle]![tbPath_Aufgebot].Value & "\" & [Reports]![schreiben_aufgebot_talbetrieb].Caption & ".pdf"
        On Error GoTo Err_InvalidPath 'Falls beim Ausführen der nächsten Linie ein Fehler auftritt zur "Err_InvalidPath"-Sprungmarke springen. Passiert wenn kein Pfad angegeben.
        DoCmd.OutputTo acOutputReport, "schreiben_aufgebot_talbetrieb", acFormatPDF, strPath
        
        DoEvents
        DoCmd.Close acReport, "schreiben_aufgebot_talbetrieb", acSaveNo
    End If

Exit Sub ' Ausstiegspunkt wenn keine Errors
Err_InvalidPath:
        MsgBox "Error! Speicherpfad prüfen"
Err_Cancel:
End Sub

' Erste Voranfrage
Private Sub buttonPDF_All_Click()
    
    Dim i As Integer
    Dim strWHERE As String
    Dim strPath As String
    
    'Gewählten Eintrag in ListBox identifizieren und zustellid entnehmen
    For i = 0 To ListeBewirtschafter_Alle.ListCount - 1
       If ListeBewirtschafter_Alle.Selected(i) Then
            strWHERE = "zustellid = " & "'" & ListeBewirtschafter_Alle.Column(0, i) & "'"
            Exit For
       End If
    Next i
    
    DoCmd.OpenReport "erste_voranfrage", acLayout, , strWHERE, acHidden
   
    strPath = [Forms]![geschaeftskontrolle]![tbPath_all].Value & "\" & [Reports]![erste_voranfrage].Caption & ".pdf"
    On Error GoTo Err_InvalidPath 'Falls beim Ausführen der nächsten Linie ein Fehler auftritt zur "Err_InvalidPath"-Sprungmarke springen. Passiert wenn kein Pfad angegeben.
    DoCmd.OutputTo acOutputReport, "erste_voranfrage", acFormatPDF, strPath
   
    DoEvents
    DoCmd.Close acReport, "erste_voranfrage", acSaveNo

Exit Sub ' Ausstiegspunkt wenn keine Errors
Err_InvalidPath:
        MsgBox "Error! Speicherpfad prüfen"
Err_Cancel:
End Sub


'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' PDF-Export Pfadauswahl Buttons
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'
Private Sub buttonChoosePathPDF_Aufgebot_Click()
    tbPath_Aufgebot.Value = tools.GetFolderPath()
End Sub

Private Sub buttonChoosePathPDF_All_Click()
    tbPath_all.Value = tools.GetFolderPath()
End Sub



