Option Explicit

Const wdFormatPDF As Integer = 17

Sub GenerarContratosProfesionales()
    
    Dim wdApp As Object, wdDoc As Object, wdRng As Object
    Dim rutaExcel As String, rutaCarpetaSalida As String, rutaPlantilla As String
    Dim nombreArchivo As String, nombreCliente As String, rut As String
    Dim direccion As String, cargo As String, estado As String, ciudad As String
    Dim lastRow As Long
    Dim i As Long
    Dim WsMain As Worksheet
    
    rutaExcel = ThisWorkbook.Path & "\"
    rutaPlantilla = rutaExcel & "Plantilla_Contrato1.docx"
    rutaCarpetaSalida = rutaExcel & "Contratos_Generados\"
    
    Set WsMain = ThisWorkbook.Sheets("Main")
    
    ' --- CONFIGURACIÓN DE COLUMNAS ---
    Dim colNombre As Integer: colNombre = 1
    Dim colDireccion As Integer: colDireccion = 3
    Dim colRUT As Integer: colRUT = 2
    Dim colCargo As Integer: colCargo = 4
    Dim colCiudad As Integer: colCiudad = 5
    Dim colEstado As Integer: colEstado = 6
    
    ' Crear carpeta de salida si no existe
    If Dir(rutaCarpetaSalida, vbDirectory) = "" Then
        MkDir rutaCarpetaSalida
    End If
    
    ' Verificar plantilla
    If Dir(rutaPlantilla) = "" Then
        MsgBox "? No se encontró la plantilla en: " & rutaPlantilla, vbCritical
        Exit Sub
    End If
    
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    
    ' Última fila de clientes (Columna A)
    lastRow = Cells(Rows.Count, colNombre).End(xlUp).Row
    
    ' Iniciar Word
    Set wdApp = CreateObject("Word.Application")
    wdApp.Visible = False
        
        
    For i = 10 To lastRow
        
        nombreCliente = Trim(WsMain.Cells(i, colNombre).Value)
        direccion = Trim(WsMain.Cells(i, colDireccion).Value)
        rut = Trim(WsMain.Cells(i, colRUT).Value)
        cargo = Trim(WsMain.Cells(i, colCargo).Value)
        ciudad = Trim(WsMain.Cells(i, colCiudad).Value)
        
        If nombreCliente <> "" Then
            
            ' Limpiar nombre para archivo
            nombreArchivo = LimpiarNombreArchivo(nombreCliente)
            
            ' --- VALIDAR SI EL ARCHIVO YA EXISTE ---
                    If Dir(rutaCarpetaSalida & nombreArchivo & ".docx") <> "" Or Dir(rutaCarpetaSalida & nombreArchivo & ".pdf") <> "" Then
                        
                        ' El archivo ya existe
                        Cells(i, colEstado).Value = "?? No generado: Archivo ya existe"
                        Cells(i, colEstado).Font.Color = RGB(255, 165, 0) ' Naranja
                        Cells(i, colEstado).Font.Bold = True
                        
                        ' Opcional: Saltar a siguiente cliente sin hacer nada
                        GoTo SiguienteCliente
                        
                    End If
                    
            ' --- EL ARCHIVO NO EXISTE, PROCEDER A GENERAR ---
                    On Error Resume Next
                    Set wdDoc = wdApp.Documents.Open(rutaPlantilla)
                    On Error GoTo 0
                    
                    If wdDoc Is Nothing Then
                        Cells(i, colEstado).Value = "? Error: No se pudo abrir la plantilla"
                        Cells(i, colEstado).Font.Color = RGB(255, 0, 0) ' Rojo
                        GoTo SiguienteCliente
                    End If
            
            ' 1. Reemplazar Textos --------------------------------------------------------------------<<<<
            ReemplazarTexto wdDoc, "<<NOMBRE_CLIENTE>>", nombreCliente
            ReemplazarTexto wdDoc, "<<DIRECCION>>", direccion
            ReemplazarTexto wdDoc, "<<RUT_CLIENTE>>", rut
            ReemplazarTexto wdDoc, "<<CARGO>>", cargo
            ReemplazarTexto wdDoc, "<<FECHA>>", Format(Date, "DD-MM-YYYY")
            ReemplazarTexto wdDoc, "<<CIUDAD>>", ciudad ' Cambiar si es necesario
            
            
            ' 2. Pegar la Tabla en el Marcador --------------------------------------------------------<<<<
            'Copiar Tabla
            WsMain.Range("I9:M11").Copy

                On Error Resume Next
                Set wdRng = wdDoc.Bookmarks("TablaContrato").Range
                On Error GoTo 0
                
                If Not wdRng Is Nothing Then
                    ' Pegar la tabla manteniendo formato de Excel
                    ' wdRng.PasteExcelTable False, True, False
                    ' Primer Parámetro: LinkToExcel (Vincular a Excel)
                    ' Segundo Parámetro: OverwriteFormat (Sobrescribir Formato)
                    ' Tercer Parámetro: AutoFitBehavior (Ajuste Automático)
                    
                    wdRng.PasteExcelTable False, False, False
                Else
                    WsMain.Cells(i, colEstado).Value = "? Error: No se encontró el marcador 'TablaContrato'"
                    WsMain.Cells(i, colEstado).Font.Color = RGB(255, 0, 0)
                    wdDoc.Close SaveChanges:=False
                    Set wdDoc = Nothing
                    GoTo SiguienteCliente
                End If
            
            ' 3. Guardar Word
            On Error Resume Next
            wdDoc.SaveAs2 Filename:=rutaCarpetaSalida & nombreArchivo & ".docx"
            If Err.Number <> 0 Then
                WsMain.Cells(i, colEstado).Value = "? Error al guardar Word"
                WsMain.Cells(i, colEstado).Font.Color = RGB(255, 0, 0)
                wdDoc.Close SaveChanges:=False
                Set wdDoc = Nothing
                Err.Clear
                GoTo SiguienteCliente
            End If
            On Error GoTo 0
            
            ' 4. Guardar PDF
            On Error Resume Next
            wdDoc.ExportAsFixedFormat OutputFileName:=rutaCarpetaSalida & nombreArchivo & ".pdf", ExportFormat:=wdFormatPDF
            On Error GoTo 0
            
            ' 5. Cerrar documento y marcar éxito
            wdDoc.Close SaveChanges:=False
            Set wdDoc = Nothing
            
            WsMain.Cells(i, colEstado).Value = "? Contrato Creado"
            WsMain.Cells(i, colEstado).Font.Color = RGB(0, 128, 0) ' Verde
            WsMain.Cells(i, colEstado).Font.Bold = True
            
SiguienteCliente:
            Set wdRng = Nothing
            Set wdDoc = Nothing
            
        Else
            WsMain.Cells(i, colEstado).Value = "? Sin nombre"
            WsMain.Cells(i, colEstado).Font.Color = RGB(128, 128, 128)
        End If
        
    Next i
    
    ' Cerrar Word
    wdApp.Quit
    
    ' Limpiar objetos
    Set wdDoc = Nothing
    Set wdApp = Nothing
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    MsgBox "?? Proceso finalizado. Revisa la columna 'Estado Generación' para ver el resultado.", vbInformation

End Sub

' Función auxiliar para reemplazar texto en Word
Sub ReemplazarTexto(doc As Object, marcador As String, valor As String)
    On Error Resume Next
    With doc.Content.Find
        .Text = marcador
        .Replacement.Text = valor
        .Forward = True
        .Wrap = 1 ' wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .Execute Replace:=2 ' wdReplaceAll
    End With
    On Error GoTo 0
End Sub

' Función para limpiar nombres de archivo (Caracteres prohibidos en Windows)
Function LimpiarNombreArchivo(nombre As String) As String
    
    Dim caracteresInvalidos As Variant
    Dim i As Integer
    caracteresInvalidos = Array("/", "\", ":", "*", "?", """", "<", ">", "|")
    LimpiarNombreArchivo = nombre
    For i = LBound(caracteresInvalidos) To UBound(caracteresInvalidos)
        LimpiarNombreArchivo = Replace(LimpiarNombreArchivo, caracteresInvalidos(i), "_")
    Next i
    ' Recortar si es muy largo (límite seguro 50 caracteres)
    If Len(LimpiarNombreArchivo) > 50 Then
        LimpiarNombreArchivo = Left(LimpiarNombreArchivo, 50)
    End If
    ' Eliminar espacios al inicio y final
    LimpiarNombreArchivo = Trim(LimpiarNombreArchivo)
    
End Function

