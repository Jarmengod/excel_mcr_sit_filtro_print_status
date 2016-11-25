
 
Sub Macro0_renombrar()
' -------------------------------------
'  Macro0 to save this macro excel file with another name
' -------------------------------------
' Variable
' Declaration
'
 Dim nameMacro As String
 Dim nameSheetConfig As String
 Dim defaultFileName As String
 Dim fileSaveName As String
 Dim currentFile As String
 
 Dim varInitCell As Range
 
 

' Assignation
'
'  defaultFileName = "filtro_mpts_wXX"
'  Puntero a area de variables
  nameMacro = "Macro0"
  nameSheetConfig = Sheets(Sheets.Count).Range("A1").Value
  Set varInitCell = Sheets(nameSheetConfig).Range("A1:bZ255").Find(nameMacro)

' inicializacion variables

  defaultFileName = Sheets(nameSheetConfig).Cells(varInitCell.Row + 1, varInitCell.Column + 1).Value
  



 
 ChDir ActiveWorkbook.Path
 'currentFile = ActiveWorkbook.FullName
 currentFile = ActiveWorkbook.Name
 
 ' MsgBox currentFile

 fileSaveName = Application.GetSaveAsFilename( _
    fileFilter:="Excel Macro-Enabled Workbook (*.xlsm),*.xlsm", _
    InitialFileName:=defaultFileName, _
    Title:="Save As File Name")
 If fileSaveName <> "" Then
      ' MsgBox "Save as " & fileSaveName    End If
      ActiveWorkbook.SaveAs Filename:=fileSaveName, FileFormat:=52
  End If
   
    Workbooks.Open Filename:=fileSaveName
         
    Workbooks(currentFile).Close False
  
  


End Sub
Sub Macro1_all()

t = Timer

MacroA_crea_hoja_status
MacroB_borrado_columnas_no_info
MacroC_borrado_filas_en_blanco
MacroD_borrado_filas_no_info
MacroE_borrado_layer
MacroF_fw_code
MacroG_column_status
MacroH_dejar_name_device
MacroI_formato
MacroJ_limpieza_final

MsgBox Timer - t

End Sub
Sub MacroA_crea_hoja_status()
' ----------------------------------
'  Cambia el nombre en la hoja en blanco y copia en hoja nueva para trabajo
' ---------------------------------
'------------------------------------------
'  Variable
'   ----- Declaracion --------------------
Dim nameSheetConfig, sheetRawData, sheetStatusUnit As String
Dim nameMacro As String
'     ----- Asignacion ---------------------------
'     ----------  Busqueda zona de variables macro
nameMacro = "MacroA"
nameSheetConfig = Sheets(Sheets.Count).Range("A1").Value
  Set varInitCell = Sheets(nameSheetConfig).Range("A1:BZ255").Find(nameMacro)
  
'    ------ Asignacion de valores a variables
  
sheetRawData = Sheets(nameSheetConfig).Cells(varInitCell.Row + 1, varInitCell.Column + 1).Value
sheetStatusUnit = Sheets(nameSheetConfig).Cells(varInitCell.Row + 2, varInitCell.Column + 1).Value
textMensaje = Sheets(nameSheetConfig).Cells(varInitCell.Row + 3, varInitCell.Column + 1).Value
titleMensaje = Sheets(nameSheetConfig).Cells(varInitCell.Row + 4, varInitCell.Column + 1).Value


'  -------- Cuerpo macro  ---------------
'    ----   Mensaje de recordatorio

resultado = MsgBox(textMensaje, vbYesNo + vbExclamation, titleMensaje)


If resultado = 6 Then

'  ------ Cambio de nombre en hoja de trabajo y copia para filtrar informacion
    Sheets(1).Name = sheetRawData
'   ------ Copio la hoja con los datos copiados y le pongo nombre
    Sheets(sheetRawData).Copy Before:=Sheets(1)
    Sheets(1).Name = sheetStatusUnit
    
Else

    MsgBox "COPIALO"
    
End If

End Sub


Sub MacroB_borrado_columnas_no_info()
' ----------------------------------
'  De la hoja de trabajo borramos las columnas con informacion que no utilizar
' ---------------------------------
'------------------------------------------
'  Variable
'   ----- Declaracion --------------------
Dim nameSheetConfig, sheetStatusUnit As String
Dim nameMacro As String
'     ----- Asignacion ---------------------------
'     ----------  Busqueda zona de variables macro
nameMacro = "MacroB"
nameSheetConfig = Sheets(Sheets.Count).Range("A1").Value
  Set varInitCell = Sheets(nameSheetConfig).Range("A1:BZ255").Find(nameMacro)
  
'    ------ Asignacion de valores a variables
sheetStatusUnit = Sheets(nameSheetConfig).Cells(varInitCell.Row + 1, varInitCell.Column + 1).Value
finalColumn = Sheets(nameSheetConfig).Cells(varInitCell.Row + 2, varInitCell.Column + 1).Value
 


'  -------- Cuerpo macro  ---------------
'    ----   Borrado columnas de derecha a izquierda
Sheets(sheetStatusUnit).Select

For i = finalColumn To 2 Step -1
    Columns(i).EntireColumn.Delete
Next

 

End Sub

Sub MacroC_borrado_filas_en_blanco()
' ----------------------------------
'  Buscamos las filas que esten en blanco - miramos columna A (1)
' ---------------------------------
'------------------------------------------
'  Variable
'   ----- Declaracion --------------------
Dim nameSheetConfig, sheetStatusUnit As String
Dim nameMacro As String
Dim EndRow, rowComp, colDades As Integer
'     ----- Asignacion ---------------------------
'     ----------  Busqueda zona de variables macro
nameMacro = "MacroC"
nameSheetConfig = Sheets(Sheets.Count).Range("A1").Value
  Set varInitCell = Sheets(nameSheetConfig).Range("A1:BZ255").Find(nameMacro)
  
'    ------ Asignacion de valores a variables
sheetStatusUnit = Sheets(nameSheetConfig).Cells(varInitCell.Row + 1, varInitCell.Column + 1).Value
initialCell = Sheets(nameSheetConfig).Cells(varInitCell.Row + 2, varInitCell.Column + 1).Value
EndRow = Sheets(nameSheetConfig).Cells(varInitCell.Row + 3, varInitCell.Column + 1).Value
 


'  -------- Cuerpo macro  ---------------
'    ----   Select columna y fila inicial
  Sheets(sheetStatusUnit).Select
    Range(initialCell).Select

    colDades = Range(initialCell).Column
    rowComp = Range(initialCell).Row
    

' Borrado de lineas en blanco

For i = EndRow To rowComp Step -1
    If IsEmpty(Cells(i, colDades)) Then Cells(i, 1).EntireRow.Delete
Next i



End Sub

Sub MacroD_borrado_filas_no_info()
' ----------------------------------
'  Borramos las filas que no empiecen por la info de una lista (nombre de las impresoras y printbuckets
' ---------------------------------
'------------------------------------------
'  Variable
'   ----- Declaracion --------------------
Dim nameSheetConfig, sheetStatusUnit As String
Dim nameMacro As String
Dim rowComp, colDades As Integer
Dim listfilter1Col, listfilter1Row As Integer
'     ----- Asignacion ---------------------------
'     ----------  Busqueda zona de variables macro
nameMacro = "MacroD"
nameSheetConfig = Sheets(Sheets.Count).Range("A1").Value
  Set varInitCell = Sheets(nameSheetConfig).Range("A1:BZ255").Find(nameMacro)
  
'    ------ Asignacion de valores a variables
sheetStatusUnit = Sheets(nameSheetConfig).Cells(varInitCell.Row + 1, varInitCell.Column + 1).Value
initialCell = Sheets(nameSheetConfig).Cells(varInitCell.Row + 2, varInitCell.Column + 1).Value
listFilter1Position = Sheets(nameSheetConfig).Cells(varInitCell.Row + 3, varInitCell.Column + 1).Value
 



'  -------- Cuerpo macro  ---------------
'    ----   Select columna y fila inicial
  Sheets(sheetStatusUnit).Select
    Range(initialCell).Select

    colDades = Range(initialCell).Column
    rowComp = Range(initialCell).Row
    

' Borrado de lineas con info diferente de la lista de si acceptadas

Do While Not IsEmpty(Cells(rowComp, colDades))

    borrado = True
    listfilter1Col = Sheets(nameSheetConfig).Range(listFilter1Position).Column
    listfilter1Row = Sheets(nameSheetConfig).Range(listFilter1Position).Row
           
    Do While Not IsEmpty(Sheets(nameSheetConfig).Cells(listfilter1Row, listfilter1Col)) And borrado
              If Sheets(nameSheetConfig).Cells(listfilter1Row, listfilter1Col).Value = Left(Cells(rowComp, colDades).Value, 3) Then
                borrado = False
              Else
                listfilter1Row = listfilter1Row + 1
              End If
    Loop
    
    If borrado Then
        Cells(rowComp, colDades).EntireRow.Delete
       Else
         rowComp = rowComp + 1
     End If
                
Loop



End Sub

Sub MacroE_borrado_layer()
' ----------------------------------
'  Borramos las filas que contienen el numero de layers de los device.
'  SE realiza buscando una palabra clave de inicio de borrado y otra de final de borrado
' ---------------------------------
'------------------------------------------
'  Variable
'   ----- Declaracion --------------------
Dim nameSheetConfig, sheetStatusUnit As String
Dim nameMacro As String
Dim secondFilterStart, secondFilterEnd As String
Dim rowComp, colDades As Integer
Dim listfilter1Col, listfilter1Row As Integer
Dim borrado As Boolean

'     ----- Asignacion ---------------------------
'     ----------  Busqueda zona de variables macro
nameMacro = "MacroE"
nameSheetConfig = Sheets(Sheets.Count).Range("A1").Value
  Set varInitCell = Sheets(nameSheetConfig).Range("A1:BZ255").Find(nameMacro)
  
'    ------ Asignacion de valores a variables
sheetStatusUnit = Sheets(nameSheetConfig).Cells(varInitCell.Row + 1, varInitCell.Column + 1).Value
initialCell = Sheets(nameSheetConfig).Cells(varInitCell.Row + 2, varInitCell.Column + 1).Value
secondFilterStart = Sheets(nameSheetConfig).Cells(varInitCell.Row + 3, varInitCell.Column + 1).Value
secondFilterEnd = Sheets(nameSheetConfig).Cells(varInitCell.Row + 4, varInitCell.Column + 1).Value
borrado = False



'  -------- Cuerpo macro  ---------------
'    ----   Select columna y fila inicial
  Sheets(sheetStatusUnit).Select
    Range(initialCell).Select

    colDades = Range(initialCell).Column
    rowComp = Range(initialCell).Row
    

'  ---  Borrado de lineas con info de layer

'    -----  Busqueda de fila con la info de inicio de borrado

    Do While Not IsEmpty(Cells(rowComp, colDades)) And Not borrado
        If Cells(rowComp, colDades).Value = secondFilterStart Then
            borrado = True
        Else
                  rowComp = rowComp + 1
        End If
    Loop

'    ------    Borrar filas hasta llegar a la fila de final de borrado

Do While Left(Cells(rowComp, colDades).Value, 12) <> secondFilterEnd And borrado
              Cells(rowComp, colDades).EntireRow.Delete
Loop
                


End Sub

Sub MacroF_fw_code()
' ----------------------------------
'  Buscamos las filas cn la info del fw release y
'  la ponemos al lado de la maquina
' ---------------------------------
'------------------------------------------
'  Variable
'   ----- Declaracion --------------------
Dim nameSheetConfig, sheetStatusUnit As String
Dim nameMacro As String
Dim fwMoveWord, textToSubs As String
Dim rowComp, colDades As Integer
Dim offsetColumn As Integer
'     ----- Asignacion ---------------------------
'     ----------  Busqueda zona de variables macro
nameMacro = "MacroF"
nameSheetConfig = Sheets(Sheets.Count).Range("A1").Value
  Set varInitCell = Sheets(nameSheetConfig).Range("A1:BZ255").Find(nameMacro)
  
'    ------ Asignacion de valores a variables
sheetStatusUnit = Sheets(nameSheetConfig).Cells(varInitCell.Row + 1, varInitCell.Column + 1).Value
initialCell = Sheets(nameSheetConfig).Cells(varInitCell.Row + 2, varInitCell.Column + 1).Value
fwMoveWord = Sheets(nameSheetConfig).Cells(varInitCell.Row + 3, varInitCell.Column + 1).Value
offsetColumn = Sheets(nameSheetConfig).Cells(varInitCell.Row + 4, varInitCell.Column + 1).Value
textToSubs = Sheets(nameSheetConfig).Cells(varInitCell.Row + 5, varInitCell.Column + 1).Value


'  -------- Cuerpo macro  ---------------
'    ----   Select columna y fila inicial
  Sheets(sheetStatusUnit).Select
    Range(initialCell).Select

    colDades = Range(initialCell).Column
    rowComp = Range(initialCell).Row
    

'  ---  Buscar revision de fw y colocarlo en columna al lado printer

Do While Not IsEmpty(Cells(rowComp, colDades))

    If Left(Cells(rowComp, colDades).Value, 2) = fwMoveWord Then
        Cells(rowComp - 1, colDades + offsetColumn).Value = WorksheetFunction.Substitute(Cells(rowComp, colDades).Value, textToSubs, "")
        Cells(rowComp, colDades).EntireRow.Delete
    Else
       rowComp = rowComp + 1
    End If
   
                
Loop
                


End Sub

Sub MacroG_column_status()
' ----------------------------------
'  Buscamos el status del device y lo copiamos en columna del lado.
'  SE realiza buscando una palabra clave de inicio de borrado y otra de final de borrado
' ---------------------------------
'------------------------------------------
'  Variable
'   ----- Declaracion --------------------
Dim nameSheetConfig, sheetStatusUnit As String
Dim nameMacro As String
Dim filter1, filter2, filter3 As String
Dim text1, text2, text3, text4 As String
Dim rowComp, colDades As Integer
Dim offsetColumn As Integer
'     ----- Asignacion ---------------------------
'     ----------  Busqueda zona de variables macro
nameMacro = "MacroG"
nameSheetConfig = Sheets(Sheets.Count).Range("A1").Value
  Set varInitCell = Sheets(nameSheetConfig).Range("A1:BZ255").Find(nameMacro)
  
'    ------ Asignacion de valores a variables
sheetStatusUnit = Sheets(nameSheetConfig).Cells(varInitCell.Row + 1, varInitCell.Column + 1).Value
initialCell = Sheets(nameSheetConfig).Cells(varInitCell.Row + 2, varInitCell.Column + 1).Value
offsetColumn = Sheets(nameSheetConfig).Cells(varInitCell.Row + 3, varInitCell.Column + 1).Value
filter1 = Sheets(nameSheetConfig).Cells(varInitCell.Row + 4, varInitCell.Column + 1).Value
filter2 = Sheets(nameSheetConfig).Cells(varInitCell.Row + 5, varInitCell.Column + 1).Value
filter3 = Sheets(nameSheetConfig).Cells(varInitCell.Row + 6, varInitCell.Column + 1).Value
text1 = Sheets(nameSheetConfig).Cells(varInitCell.Row + 7, varInitCell.Column + 1).Value
text2 = Sheets(nameSheetConfig).Cells(varInitCell.Row + 8, varInitCell.Column + 1).Value
text3 = Sheets(nameSheetConfig).Cells(varInitCell.Row + 9, varInitCell.Column + 1).Value
text4 = Sheets(nameSheetConfig).Cells(varInitCell.Row + 10, varInitCell.Column + 1).Value



'  -------- Cuerpo macro  ---------------
'    ----   Select columna y fila inicial
  Sheets(sheetStatusUnit).Select
    Range(initialCell).Select

    colDades = Range(initialCell).Column
    rowComp = Range(initialCell).Row
    

'  ---  Buscar revision de fw y colocarlo en columna al lado printer

Do While Not IsEmpty(Cells(rowComp, colDades))
    
 ' Poner operativa, no operativa o repa...
 
    If InStr(LCase(Cells(rowComp, colDades).Value), filter1) <> 0 Then
        Cells(rowComp, colDades + offsetColumn).Value = text1
         
    Else
       If InStr(LCase(Cells(rowComp, colDades).Value), filter2) <> 0 Then
            Cells(rowComp, colDades + offsetColumn).Value = text2
          Else
             If InStr(LCase(Cells(rowComp, colDades).Value), filter3) <> 0 Then
                Cells(rowComp, colDades + offsetColumn).Value = text3
               Else
                 Cells(rowComp, colDades + offsetColumn).Value = text4
                End If
           End If
      End If
  rowComp = rowComp + 1
Loop
            


End Sub

Sub MacroH_dejar_name_device()
' ----------------------------------
'  De la celda dejamos solo el nombre de la maquina
'
' ---------------------------------
'------------------------------------------
'  Variable
'   ----- Declaracion --------------------
Dim nameSheetConfig, sheetStatusUnit As String
Dim nameMacro As String
Dim numberletter As Integer
Dim filter1, filter2, filter3, filter4 As String
Dim offset1, offset2, offset3, offset4 As Integer

Dim rowComp, colDades As Integer
Dim temp As String
'     ----- Asignacion ---------------------------
'     ----------  Busqueda zona de variables macro
nameMacro = "MacroH"
nameSheetConfig = Sheets(Sheets.Count).Range("A1").Value
  Set varInitCell = Sheets(nameSheetConfig).Range("A1:BZ255").Find(nameMacro)
  
'    ------ Asignacion de valores a variables
sheetStatusUnit = Sheets(nameSheetConfig).Cells(varInitCell.Row + 1, varInitCell.Column + 1).Value
initialCell = Sheets(nameSheetConfig).Cells(varInitCell.Row + 2, varInitCell.Column + 1).Value
numberletter = Sheets(nameSheetConfig).Cells(varInitCell.Row + 3, varInitCell.Column + 1).Value
filter1 = Sheets(nameSheetConfig).Cells(varInitCell.Row + 4, varInitCell.Column + 1).Value
filter2 = Sheets(nameSheetConfig).Cells(varInitCell.Row + 5, varInitCell.Column + 1).Value
filter3 = Sheets(nameSheetConfig).Cells(varInitCell.Row + 6, varInitCell.Column + 1).Value
filter4 = Sheets(nameSheetConfig).Cells(varInitCell.Row + 7, varInitCell.Column + 1).Value
offset1 = Sheets(nameSheetConfig).Cells(varInitCell.Row + 8, varInitCell.Column + 1).Value
offset2 = Sheets(nameSheetConfig).Cells(varInitCell.Row + 9, varInitCell.Column + 1).Value
offset3 = Sheets(nameSheetConfig).Cells(varInitCell.Row + 10, varInitCell.Column + 1).Value
offset4 = Sheets(nameSheetConfig).Cells(varInitCell.Row + 11, varInitCell.Column + 1).Value


'  -------- Cuerpo macro  ---------------
'    ----   Select columna y fila inicial
  Sheets(sheetStatusUnit).Select
    Range(initialCell).Select

    colDades = Range(initialCell).Column
    rowComp = Range(initialCell).Row
    
'  ---  Buscar revision nombre device y borrar resto info

Do While Not IsEmpty(Cells(rowComp, colDades))
    
 ' Poner operativa, no operativa o repa...
        temp = Left(Cells(rowComp, colDades).Value, 3)
        
   Select Case temp
     Case filter1
         Cells(rowComp, colDades).Value = Left(Cells(rowComp, colDades).Value, offset1)
     Case filter2
         Cells(rowComp, colDades).Value = Left(Cells(rowComp, colDades).Value, offset2)
     Case filter3
         Cells(rowComp, colDades).Value = Left(Cells(rowComp, colDades).Value, offset3)
     Case filter4
         Cells(rowComp, colDades).Value = Left(Cells(rowComp, colDades).Value, offset4)
     Case Else
              Cells(rowComp, colDades).Value = Cells(rowComp, colDades).Value
  End Select
      
   rowComp = rowComp + 1
Loop
            


End Sub

Sub MacroI_formato()
' ----------------------------------
'  Borramos info fila 1 y 2, ponemos cabecera, autofit y formato condicional
'
' ---------------------------------
'------------------------------------------
'  Variable
'   ----- Declaracion --------------------
Dim nameSheetConfig, sheetStatusUnit As String
Dim nameMacro As String
'Dim listHeader As Range
Dim Nooperativa, reparacion As String

Dim listHeaderCol, listHeaderRow As Integer
Dim rowComp, colDades As Integer
Dim i, temp As String
'     ----- Asignacion ---------------------------
'     ----------  Busqueda zona de variables macro
nameMacro = "MacroI"
nameSheetConfig = Sheets(Sheets.Count).Range("A1").Value
  Set varInitCell = Sheets(nameSheetConfig).Range("A1:BZ255").Find(nameMacro)
  
'    ------ Asignacion de valores a variables
sheetStatusUnit = Sheets(nameSheetConfig).Cells(varInitCell.Row + 1, varInitCell.Column + 1).Value
initialCell = Sheets(nameSheetConfig).Cells(varInitCell.Row + 2, varInitCell.Column + 1).Value
listHeader = Sheets(nameSheetConfig).Cells(varInitCell.Row + 3, varInitCell.Column + 1).Value
Nooperativa = Sheets(nameSheetConfig).Cells(varInitCell.Row + 4, varInitCell.Column + 1).Value
reparacion = Sheets(nameSheetConfig).Cells(varInitCell.Row + 5, varInitCell.Column + 1).Value

'  -------- Cuerpo macro  ---------------
'    ----   Select columna y fila inicial
  Sheets(sheetStatusUnit).Select
    Range(initialCell).Select

    colDades = Range(initialCell).Column
    rowComp = Range(initialCell).Row
    
'   ------  Borrado casillas filas anteriores
For i = rowComp - 1 To 1 Step -1
    Cells(i, colDades).Value = ""
Next i

'   ------   Poner cabecera
 colDades = Range(initialCell).Column
 rowComp = Range(initialCell).Row - 1
 listHeaderCol = Sheets(nameSheetConfig).Range(listHeader).Column
 listHeaderRow = Sheets(nameSheetConfig).Range(listHeader).Row
           
    Do While Not IsEmpty(Sheets(nameSheetConfig).Cells(listHeaderRow, listHeaderCol))
              Sheets(sheetStatusUnit).Cells(rowComp, colDades).Value = Sheets(nameSheetConfig).Cells(listHeaderRow, listHeaderCol).Value
              colDades = colDades + 1
              listHeaderRow = listHeaderRow + 1
                           
    Loop

'  Autofit de columnas

colDades = Range(initialCell).Column
rowComp = Range(initialCell).Row - 1

        
    Do While Not IsEmpty(Sheets(sheetStatusUnit).Cells(rowComp, colDades))
             Columns(colDades).AutoFit
             colDades = colDades + 1
    Loop

    
'    ----   Formato Condicional



    Range(initialCell).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    colDades = Range(initialCell).Column
    
    
    Selection.FormatConditions.Add Type:=xlTextString, String:=Nooperativa, _
        TextOperator:=xlContains
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 255
        .TintAndShade = 0
    End With
     With Selection.FormatConditions(1).Font
        .Bold = True
        .Italic = False
        .Color = -16711681
        .TintAndShade = 0
    End With
     Selection.FormatConditions.Add Type:=xlTextString, String:=reparacion, _
        TextOperator:=xlContains
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
     With Selection.FormatConditions(1).Font
        .Bold = True
        .Italic = False
        .Color = -16711681
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 49407
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
            


End Sub

Sub MacroJ_limpieza_final()
' ----------------------------------
'  Borramos del resultado final lineas que se nos cuelen (trolleyNvm, ...)
' ---------------------------------
'------------------------------------------
'  Variable
'   ----- Declaracion --------------------
Dim nameSheetConfig, sheetStatusUnit As String
Dim nameMacro As String
Dim rowComp, colDades As Integer
Dim longLimpieza, listfilter1Col, listfilter1Row As Integer
'     ----- Asignacion ---------------------------
'     ----------  Busqueda zona de variables macro
nameMacro = "MacroJ"
nameSheetConfig = Sheets(Sheets.Count).Range("A1").Value
  Set varInitCell = Sheets(nameSheetConfig).Range("A1:BZ255").Find(nameMacro)
  
'    ------ Asignacion de valores a variables
sheetStatusUnit = Sheets(nameSheetConfig).Cells(varInitCell.Row + 1, varInitCell.Column + 1).Value
initialCell = Sheets(nameSheetConfig).Cells(varInitCell.Row + 2, varInitCell.Column + 1).Value
listLimpieza = Sheets(nameSheetConfig).Cells(varInitCell.Row + 3, varInitCell.Column + 1).Value
 



'  -------- Cuerpo macro  ---------------
'    ----   Select columna y fila inicial
  Sheets(sheetStatusUnit).Select
    Range(initialCell).Select

    colDades = Range(initialCell).Column
    rowComp = Range(initialCell).Row
    

' Borrado de lineas con info diferente de la lista de si acceptadas

Do While Not IsEmpty(Cells(rowComp, colDades))

    borrado = False
    listfilter1Col = Sheets(nameSheetConfig).Range(listLimpieza).Column
    listfilter1Row = Sheets(nameSheetConfig).Range(listLimpieza).Row
    longLimpieza = Len(Sheets(nameSheetConfig).Cells(listfilter1Row, listfilter1Col).Value)
    
    Do While Not IsEmpty(Sheets(nameSheetConfig).Cells(listfilter1Row, listfilter1Col)) And Not borrado
              If Sheets(nameSheetConfig).Cells(listfilter1Row, listfilter1Col).Value = Left(Cells(rowComp, colDades).Value, longLimpieza) Then
                borrado = True
              Else
                listfilter1Row = listfilter1Row + 1
                longLimpieza = Len(Sheets(nameSheetConfig).Cells(listfilter1Row, listfilter1Col).Value)
              End If
    Loop
    
    If borrado Then
        Cells(rowComp, colDades).EntireRow.Delete
       Else
         rowComp = rowComp + 1
     End If
                
Loop



End Sub

Sub Macroy_Borrar_no_info()
'
' Macro2_reparaciones Macro
'
Dim nameSheetConfig, sheetRawData, sheetStatusUnit As String
Dim rgFound, rgToFind As Range
Dim borrado As Boolean

'Dim key1, key2, key3, key4 As Range


'        Initialization
t = Timer
 nameMacro = "MacroA"
 
  nameSheetConfig = Sheets(Sheets.Count).Range("A1").Value
  Set varInitCell = Sheets(nameSheetConfig).Range("A1:BZ255").Find(nameMacro)

sheetRawData = Sheets(nameSheetConfig).Cells(varInitCell.Row + 1, varInitCell.Column + 1).Value
sheetStatusUnit = Sheets(nameSheetConfig).Cells(varInitCell.Row + 2, varInitCell.Column + 1).Value
initialCell = Sheets(nameSheetConfig).Cells(varInitCell.Row + 3, varInitCell.Column + 1).Value
columnfilter1 = Sheets(nameSheetConfig).Cells(varInitCell.Row + 4, varInitCell.Column + 1).Value
listFilter1 = Sheets(nameSheetConfig).Cells(varInitCell.Row + 5, varInitCell.Column + 1).Value
EndRow = Sheets(nameSheetConfig).Cells(varInitCell.Row + 6, varInitCell.Column + 1).Value
secondFilterStart = Sheets(nameSheetConfig).Cells(varInitCell.Row + 7, varInitCell.Column + 1).Value
secondFilterEnd = Sheets(nameSheetConfig).Cells(varInitCell.Row + 8, varInitCell.Column + 1).Value
fwMoveWord = Sheets(nameSheetConfig).Cells(varInitCell.Row + 9, varInitCell.Column + 1).Value
Nooperativa = Sheets(nameSheetConfig).Cells(varInitCell.Row + 10, varInitCell.Column + 1).Value

        

  
'  Cambio de nombre en hoja de trabajo y copia para filtrar informacion
    Sheets(1).Name = sheetRawData
          
    Sheets(sheetRawData).Copy Before:=Sheets(1)
    Sheets(1).Name = sheetStatusUnit

' Proceso de limpieza de linea cuya texto no empieza por un listado que hay en la hoja config_var , columna to find
    Sheets(sheetStatusUnit).Select
    Range(initialCell).Select

    colDades = Range(initialCell).Column
    rowComp = Range(initialCell).Row

'Borrado de filas en blanco

'Borrado columna D, C i B
Columns(4).EntireColumn.Delete
Columns(3).EntireColumn.Delete
Columns(2).EntireColumn.Delete


' Borrado de lineas en blanco

For i = EndRow To rowComp Step -1
    If IsEmpty(Cells(i, colDades)) Then Cells(i, 1).EntireRow.Delete
Next i


'Borrado de lineas que no empiezan con la info contenida en config_Var-to find

Do While Not IsEmpty(Cells(rowComp, colDades))

    borrado = True
    listfilter1Col = Sheets(nameSheetConfig).Range(listFilter1).Column
    listfilter1Row = Sheets(nameSheetConfig).Range(listFilter1).Row
           
    Do While Not IsEmpty(Sheets(nameSheetConfig).Cells(listfilter1Row, listfilter1Col)) And borrado
              If Sheets(nameSheetConfig).Cells(listfilter1Row, listfilter1Col).Value = Left(Cells(rowComp, colDades).Value, 3) Then
                borrado = False
              Else
                listfilter1Row = listfilter1Row + 1
              End If
    Loop
    
    If borrado Then
        Cells(rowComp, colDades).EntireRow.Delete
       Else
         rowComp = rowComp + 1
     End If
                
Loop
   
' Segunda limpieza, informacion de numero de layer que no me interesa
  Sheets(sheetStatusUnit).Select
    Range(initialCell).Select

    colDades = Range(initialCell).Column
    rowComp = Range(initialCell).Row

Do While Cells(rowComp, colDades).Value <> secondFilterStart
              rowComp = rowComp + 1
    Loop

Do While Cells(rowComp, colDades).Value <> secondFilterEnd
              Cells(rowComp, colDades).EntireRow.Delete
Loop
    
'poner fw al lado de su maquina.

    Sheets(sheetStatusUnit).Select
    Range(initialCell).Select

    colDades = Range(initialCell).Column
    rowComp = Range(initialCell).Row

Do While Not IsEmpty(Cells(rowComp, colDades))

    If Left(Cells(rowComp, colDades).Value, 2) = fwMoveWord Then
        Cells(rowComp - 1, colDades + 2).Value = WorksheetFunction.Substitute(Cells(rowComp, colDades).Value, "FW REVISION:", "")
        Cells(rowComp, colDades).EntireRow.Delete
    Else
       rowComp = rowComp + 1
    End If
   
                
Loop

'poner operativa o no al lado de su maquina.

    Sheets(sheetStatusUnit).Select
    Range(initialCell).Select

    colDades = Range(initialCell).Column
    rowComp = Range(initialCell).Row

Do While Not IsEmpty(Cells(rowComp, colDades))
    
 ' Poner operativa, no operativa o repa...
 
    If InStr(LCase(Cells(rowComp, colDades).Value), "repa") <> 0 Then
        Cells(rowComp, colDades + 1).Value = "REPAIR"
         
    Else
       If InStr(LCase(Cells(rowComp, colDades).Value), "no opera") <> 0 Then
            Cells(rowComp, colDades + 1).Value = "NO OPERATIVA"
          Else
             If InStr(LCase(Cells(rowComp, colDades).Value), "opera") <> 0 Then
                Cells(rowComp, colDades + 1).Value = "OPERATIVA"
               Else
                 Cells(rowComp, colDades + 1).Value = "N.A."
                End If
           End If
      End If
   ' Poner solo el numero de maquina
   temp = Left(Cells(rowComp, colDades).Value, 3)
   Select Case temp
     Case "R2#"
         Cells(rowComp, colDades).Value = Left(Cells(rowComp, colDades).Value, 5)
     Case "MB0"
         Cells(rowComp, colDades).Value = Left(Cells(rowComp, colDades).Value, 7)
     Case "MB1"
         Cells(rowComp, colDades).Value = Left(Cells(rowComp, colDades).Value, 6)
     Case "PB2"
         Cells(rowComp, colDades).Value = Left(Cells(rowComp, colDades).Value, 6)
     Case Else
              Cells(rowComp, colDades).Value = Cells(rowComp, colDades).Value
  End Select
      
      rowComp = rowComp + 1
    
   
                
Loop




    Range(initialCell).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.FormatConditions.Add Type:=xlTextString, String:=Nooperativa, _
        TextOperator:=xlContains
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 255
        .TintAndShade = 0
    End With
     With Selection.FormatConditions(1).Font
        .Bold = True
        .Italic = False
        .Color = -16711681
        .TintAndShade = 0
    End With
     Selection.FormatConditions.Add Type:=xlTextString, String:="repa", _
        TextOperator:=xlContains
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
     With Selection.FormatConditions(1).Font
        .Bold = True
        .Italic = False
        .Color = -16711681
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 49407
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
' Escalar tamaño columna 1
Columns(1).AutoFit
Columns(2).AutoFit
Columns(3).AutoFit
MsgBox Timer - t

End Sub

Sub Macroz_formateo_casillas()
'
' Macroz_formateo_casillas Macro
'

'
    Range("A3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.FormatConditions.Add Type:=xlTextString, String:="No Opera", _
        TextOperator:=xlContains
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 255
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
End Sub

