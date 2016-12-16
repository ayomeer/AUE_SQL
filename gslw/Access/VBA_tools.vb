'╔══════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════╗
'║ VBA Modul: tools																																								║
'╠══════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════╣
'║ 																																												║
'║ Author: 				Andreas Kuhn (Zivildienstleistender August - Januar 2016)																								║
'║ email privat:		andreas.kuhn@bluewin.ch																																	║
'║ Archiviert unter:	M:\abt_umwelt_energie\gewaesserschutz\jauchegrube\2_DB																									║
'║ 																																												║
'║ Beschreibung:																																								║
'║ Dieses Modul bietet Hilfsfunktionalitäten für Access Anwendungen wie Forms. Bisher sind nur Funktionalitäten, die sich in Form_geschaeftskontrolle als nötig erwiesen 		║
'║ hatten.																																										║
'║ 																																												║
'║ Good to know:
'║
'╚══════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════╝

' test change
Option Compare Database

'Debug/Testing
Public Function Filter(tbValue As String, FieldName As String) As String

    Filter = "tbValue = " & tbValue & " FieldName = " & FieldName

End Function

Public Sub ExportToExcel(myListBox As ListBox)
    
    Dim myExApp As Object       'variable for Excel App
    Set myExApp = CreateObject("Excel.Application")
    
    Dim myExSheet As Object     'variable for Excel Sheet
    Set myExSheet = CreateObject("Excel.Sheet")
    
    Dim i As Long               'variable für das iterieren durch die Spalten
    Dim j As Long               'variable für das iterieren durch die Reihen
    
    myExApp.Visible = False
    myExApp.Workbooks.Add
    Set myExSheet = myExApp.Workbooks(1).Worksheets(1)

    For i = 1 To myListBox.ColumnCount
        For j = 1 To myListBox.ListCount
            myExSheet.Cells(j, i) = myListBox.ItemData(j - 1)   'Daten in Excel-Zelle schreiben
        Next j  'Iterieren durch ListCount
        myListBox.BoundColumn = myListBox.BoundColumn + 1
    Next i  'Iterieren durch ColumnCount
    myListBox.BoundColumn = 1    'Zurücksetzen von BoundColumn auf die ursprüngliche 1
    myExApp.Visible = True

End Sub

Public Function GetFolderPath(Optional OpenAt As String) As String
    'Oeffnet Dialog zum waehlen von Speicherpfaeden
    Dim lCount As Long
    GetFolderPath = vbNullString ' Variable mit selbem Name wie Funktion funktionieren als Rueckgabeparameter
    
    With Application.FileDialog(4) '4 = FileDialog: FolderPicker, mso Konstante nicht erkannt
        .InitialFileName = OpenAt
        .Show
        For lCount = 1 To .SelectedItems.Count
            GetFolderPath = .SelectedItems(lCount)
        Next lCount
    End With

End Function

