Attribute VB_Name = "PrintBlank"
Option Explicit

Sub PrintBlank()


'  Позднее связывание
'  Dim cnn As Object
'  Dim rst As Object

' Раннее связывание
  Dim cnn As ADODB.Connection
  Dim rst As ADODB.Recordset
  
  Dim strSQL As String
  Dim strExcelDataSource As String

  Dim i As Integer                '   Счетчик для номера записи
  Dim n As Integer                '   Номер ряда для ячейки вставки бланка
  Dim intLastRow As Integer       '   Номер последнего ряда с данными
  Dim intLastColumn As Integer    '   Номер последнего столбца с данными
  Dim kolPageForPrint As Integer  '   Число листов для печати
  
  Dim l As Long                   ' Индекс для HPageBreaks
  Dim HPBcount As Long            ' Число разрывов страницы, создаваемое Excel
  Dim hp() As Long                ' массив номеров строк адресов вставки разрыва страницы
  
  intLastColumn = 9

'подключение к файлу Excel
  strExcelDataSource = ThisWorkbook.Path & "\" & ThisWorkbook.Name

'  Set cnn = CreateObject("ADODB.Connection")  ' позднее связывание
  Set cnn = New ADODB.Connection               ' раннее связывание
  cnn.ConnectionString = _
    "Provider=Microsoft.ACE.OLEDB.12.0;" & _
    "Data Source= " & strExcelDataSource & _
    ";Extended Properties='Excel 12.0;HDR=YES'"
  cnn.Open

' Создаем новый объект типа Recordset
'  Set rst = CreateObject("ADODB.Recordset") ' позднее связывание
  Set rst = New ADODB.Recordset              ' раннее связывание

'составляем строку SQL запроса
  strSQL = "SELECT `Номер заказа`, `Дата размещения`, `Дата назначения`, `Код сотрудника`, `Код клиента`, `Код заказа`, Модель  FROM `tbZakaz$` `tbZakaz$`"
 
 
 'отправляем запрос открытой БД результат сохранен в rst
  rst.Open strSQL, cnn, adOpenStatic, adLockReadOnly ' Теперь В rst хранится ссылка на набор записей

    If rst.RecordCount = 0 Then
      MsgBox "Записей нет"
      rst.Close
      Exit Sub
    End If

'  Заполняем Recordset
  
  rst.MoveLast
  rst.MoveFirst

'  Копируем бланк

  Application.Goto Reference:="Blank"
  Selection.Copy
  
'  Вставляем на Новый лист
  Sheets.Add After:=Sheets(Sheets.Count)
  
    ' Считаем кол-во страниц для печати
  kolPageForPrint = Fix(rst.RecordCount / 3) + Sgn(rst.RecordCount / 3 - Fix(rst.RecordCount / 3))
  ReDim hp(kolPageForPrint - 1)

  n = 1                            '   Номер ряда для ячейки вставки бланка
  i = 1                            '   Счетчик для номера записи
  l = 0                            '   Индекс разрыва
  
  Do Until rst.EOF
  
'               Вставка бланка
     
    Range("A" & n).Select
    Selection.PasteSpecial Paste:=xlPasteColumnWidths, _
                           Operation:=xlNone, _
                           SkipBlanks:=False, _
                           Transpose:=False
    ActiveSheet.Paste
    
    ' Вставка значений полей
    Range("A" & n).Offset(2, 1) = rst.Fields("Номер заказа").Value
    Range("A" & n).Offset(1, 4) = rst.Fields("Дата размещения").Value
    Range("A" & n).Offset(1, 7) = rst.Fields("Дата назначения").Value
    Range("A" & n).Offset(4, 2) = rst.Fields("Код сотрудника").Value
    Range("A" & n).Offset(7, 1) = rst.Fields("Код клиента").Value
    Range("A" & n).Offset(7, 6) = rst.Fields("Номер заказа").Value
    Range("A" & n).Offset(9, 2) = rst.Fields("Модель").Value
        
    If i = rst.RecordCount Then intLastRow = n + 10
    
    n = n + 13              ' Следующая позиция для вставки
                                        
' По 3 на странице. Перед следующей страницей сохраняем номер строки ячейки вставки разрыва
                                        
    If (i Mod 3) = 0 Then
      hp(l) = CLng((n - 2))
      l = l + 1
    End If
    i = i + 1

    rst.MoveNext
  
  Loop
    
'Закрываем набор записей
  rst.Close
  
'Очищаем память
  Set rst = Nothing
  
'Закрываем соединение с базой
  Set cnn = Nothing
    
  ActiveSheet.Name = "ForPrint"
  
' **************************************************************
' ************* Форматируем область печати *********************
' **************************************************************

' Использование HPageBreaks или VPageBreaks.Location
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

' ************  Вставка разрыва листа  HPageBreaks  ************
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
  If MsgBox(" Подготовить для печати?", vbYesNo + vbDefaultButton2, _
            "Предварительный просмотр перед печатью") = vbYes Then
    PrintBlank
  Else
    Exit Sub
  End If
End Sub
