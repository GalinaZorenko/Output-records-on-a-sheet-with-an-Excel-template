Attribute VB_Name = "PrintBlank"
Option Explicit

Sub PrintBlank()


'  ������� ����������
'  Dim cnn As Object
'  Dim rst As Object

' ������ ����������
  Dim cnn As ADODB.Connection
  Dim rst As ADODB.Recordset
  
  Dim strSQL As String
  Dim strExcelDataSource As String

  Dim i As Integer                '   ������� ��� ������ ������
  Dim n As Integer                '   ����� ���� ��� ������ ������� ������
  Dim intLastRow As Integer       '   ����� ���������� ���� � �������
  Dim intLastColumn As Integer    '   ����� ���������� ������� � �������
  Dim kolPageForPrint As Integer  '   ����� ������ ��� ������
  
  Dim l As Long                   ' ������ ��� HPageBreaks
  Dim HPBcount As Long            ' ����� �������� ��������, ����������� Excel
  Dim hp() As Long                ' ������ ������� ����� ������� ������� ������� ��������
  
  intLastColumn = 9

'����������� � ����� Excel
  strExcelDataSource = ThisWorkbook.Path & "\" & ThisWorkbook.Name

'  Set cnn = CreateObject("ADODB.Connection")  ' ������� ����������
  Set cnn = New ADODB.Connection               ' ������ ����������
  cnn.ConnectionString = _
    "Provider=Microsoft.ACE.OLEDB.12.0;" & _
    "Data Source= " & strExcelDataSource & _
    ";Extended Properties='Excel 12.0;HDR=YES'"
  cnn.Open

' ������� ����� ������ ���� Recordset
'  Set rst = CreateObject("ADODB.Recordset") ' ������� ����������
  Set rst = New ADODB.Recordset              ' ������ ����������

'���������� ������ SQL �������
  strSQL = "SELECT `����� ������`, `���� ����������`, `���� ����������`, `��� ����������`, `��� �������`, `��� ������`, ������  FROM `tbZakaz$` `tbZakaz$`"
 
 
 '���������� ������ �������� �� ��������� �������� � rst
  rst.Open strSQL, cnn, adOpenStatic, adLockReadOnly ' ������ � rst �������� ������ �� ����� �������

    If rst.RecordCount = 0 Then
      MsgBox "������� ���"
      rst.Close
      Exit Sub
    End If

'  ��������� Recordset
  
  rst.MoveLast
  rst.MoveFirst

'  �������� �����

  Application.Goto Reference:="Blank"
  Selection.Copy
  
'  ��������� �� ����� ����
  Sheets.Add After:=Sheets(Sheets.Count)
  
    ' ������� ���-�� ������� ��� ������
  kolPageForPrint = Fix(rst.RecordCount / 3) + Sgn(rst.RecordCount / 3 - Fix(rst.RecordCount / 3))
  ReDim hp(kolPageForPrint - 1)

  n = 1                            '   ����� ���� ��� ������ ������� ������
  i = 1                            '   ������� ��� ������ ������
  l = 0                            '   ������ �������
  
  Do Until rst.EOF
  
'               ������� ������
     
    Range("A" & n).Select
    Selection.PasteSpecial Paste:=xlPasteColumnWidths, _
                           Operation:=xlNone, _
                           SkipBlanks:=False, _
                           Transpose:=False
    ActiveSheet.Paste
    
    ' ������� �������� �����
    Range("A" & n).Offset(2, 1) = rst.Fields("����� ������").Value
    Range("A" & n).Offset(1, 4) = rst.Fields("���� ����������").Value
    Range("A" & n).Offset(1, 7) = rst.Fields("���� ����������").Value
    Range("A" & n).Offset(4, 2) = rst.Fields("��� ����������").Value
    Range("A" & n).Offset(7, 1) = rst.Fields("��� �������").Value
    Range("A" & n).Offset(7, 6) = rst.Fields("����� ������").Value
    Range("A" & n).Offset(9, 2) = rst.Fields("������").Value
        
    If i = rst.RecordCount Then intLastRow = n + 10
    
    n = n + 13              ' ��������� ������� ��� �������
                                        
' �� 3 �� ��������. ����� ��������� ��������� ��������� ����� ������ ������ ������� �������
                                        
    If (i Mod 3) = 0 Then
      hp(l) = CLng((n - 2))
      l = l + 1
    End If
    i = i + 1

    rst.MoveNext
  
  Loop
    
'��������� ����� �������
  rst.Close
  
'������� ������
  Set rst = Nothing
  
'��������� ���������� � �����
  Set cnn = Nothing
    
  ActiveSheet.Name = "ForPrint"
  
' **************************************************************
' ************* ����������� ������� ������ *********************
' **************************************************************

' ������������� HPageBreaks ��� VPageBreaks.Location
' https://support.microsoft.com/en-us/help/210663/you-receive-a-subscript-out-of-range-error-message-when-you-use-hpageb
' **************************************************************
  
  Sheets("ForPrint").Select
  
  Range(Cells(1, 1), Cells(intLastRow + 1, intLastColumn)).Select

  Application.PrintCommunication = False
  With ActiveSheet.PageSetup
  
      .PrintArea = Range(Cells(1, 1), Cells(intLastRow + 1, intLastColumn)).Address
      .Orientation = xlPortrait
      .LeftMargin = Application.CentimetersToPoints(1.5)
      .RightMargin = Application.CentimetersToPoints(0.5)
      .TopMargin = Application.CentimetersToPoints(1)
      .BottomMargin = Application.CentimetersToPoints(1)
      .HeaderMargin = Application.CentimetersToPoints(0)
      .FooterMargin = Application.CentimetersToPoints(0)
      .PaperSize = xlPaperA4
      .Order = xlDownThenOver
      .Zoom = False
      .FitToPagesWide = 1
      .FitToPagesTall = kolPageForPrint
  
  End With
  Application.PrintCommunication = True

' ************  ������� ������� �����  HPageBreaks  ************
  Cells(intLastRow + 1, intLastColumn).Select
  
  ActiveWindow.View = xlPageBreakPreview
  
  HPBcount = ActiveSheet.HPageBreaks.Count
  
  For l = 1 To kolPageForPrint - 1
    If kolPageForPrint = 1 Then Exit For

    Set ActiveSheet.HPageBreaks(l).Location = ActiveSheet.Range("A" & hp(l - 1))
    HPBcount = ActiveSheet.HPageBreaks.Count
  Next l
  
  ActiveWindow.View = xlNormalView
  Range("A1").Select
  ActiveSheet.PrintPreview (True)
  
End Sub

Sub ButtonPreviewPrint()
  If MsgBox(" ����������� ��� ������?", vbYesNo + vbDefaultButton2, _
            "��������������� �������� ����� �������") = vbYes Then
    PrintBlank
  Else
    Exit Sub
  End If
End Sub
