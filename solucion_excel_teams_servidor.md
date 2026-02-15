# Solución exacta en VBA para tu proceso (Teams + servidor)

Perfecto, con la aclaración que diste, la mejor opción es **VBA** en un libro “controlador” (por ejemplo `X1.xlsm`).

## Qué hará la macro exactamente
Desde el libro donde tengas el programa:
1. Abre estos 3 ficheros (el CSV se abre delimitado por `;` para separar columnas A:K):
   - `CW_ES_Extract_Test.xlsm` (Teams/SharePoint)
   - `PW_ES_Extract_Test.xlsm` (Teams/SharePoint)
   - `BOARD reporting extract_Iberia - Copy.csv` (servidor `X:`)
2. Borra en `PW` los datos del rango **A:K** excepto cabecera (fila 1).
3. Copia datos de `CW` a `PW` **sin cabecera** (desde fila 2), solo columnas **A:K**.
4. Borra en `CW` los datos del rango **A:K** excepto cabecera (fila 1).
5. Copia datos del `CSV` a `CW` **sin cabecera** (desde fila 2), solo columnas **A:K**.
6. Guarda `CW` y `PW`, cierra `CW`, `PW` y `CSV`, y deja mensaje final.

---


## Hojas utilizadas (según tu indicación)
- Primer fichero (`CW_ES_Extract_Test.xlsm`): **"BOARD reporting extract"**
- Segundo fichero (`PW_ES_Extract_Test.xlsm`): **"BOARD reporting extract"**
- Tercer fichero (`BOARD reporting extract_Iberia - Copy.csv`): **"BOARD reporting extract_Iberia"**

## Macro VBA lista para pegar
En tu Excel:
- `ALT + F11` → **Insertar > Módulo** → pega este código.
- Guarda el archivo como `.xlsm`.

```vb
Option Explicit

Public Sub Proceso_Semanal_Iberia()
    On Error GoTo ControlError

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False

    Dim wbCW As Workbook
    Dim wbPW As Workbook
    Dim wbCSV As Workbook

    Dim wsCW As Worksheet
    Dim wsPW As Worksheet
    Dim wsCSV As Worksheet

    Const RUTA_CW As String = "https://tdworldwide.sharepoint.com/sites/CG_BEXIBITCountries-IBERIA/Shared%20Documents/IBERIA/00.CoE%20Ctry%20Structure/111.Weekly%20Board/00.Datasets/CW_ES_Extract_Test.xlsm"
    Const RUTA_PW As String = "https://tdworldwide.sharepoint.com/sites/CG_BEXIBITCountries-IBERIA/Shared%20Documents/IBERIA/00.CoE%20Ctry%20Structure/111.Weekly%20Board/00.Datasets/PW_ES_Extract_Test.xlsm"
    Const RUTA_CSV As String = "X:\DE\EUROPESHARE\EU_FP&A_CoE\03_BEXIBIT\02.COUNTRIES\02.IBERIA\111.Weekly Board\00.Datasets\MOVEIT_files\BOARD reporting extract_Iberia - Copy.csv"

    ' 1) Abrir ficheros
    Set wbCW = Workbooks.Open(Filename:=RUTA_CW)
    Set wbPW = Workbooks.Open(Filename:=RUTA_PW)

    ' Abrimos el CSV forzando delimitador ; para que no quede todo en columna A
    Workbooks.OpenText Filename:=RUTA_CSV, DataType:=xlDelimited, Semicolon:=True, Comma:=False, Local:=True
    Set wbCSV = ActiveWorkbook

    Set wsCW = ObtenerHoja(wbCW, "BOARD reporting extract", True)
    Set wsPW = ObtenerHoja(wbPW, "BOARD reporting extract", True)
    Set wsCSV = ObtenerHoja(wbCSV, "BOARD reporting extract_Iberia", False) ' CSV: fallback a hoja 1

    ' 2) Borrar datos en PW excepto cabecera (fila 1)
    LimpiarDatosExceptoCabecera wsPW

    ' 3) Copiar datos de CW -> PW sin cabecera
    CopiarDatosSinCabecera wsCW, wsPW

    ' 4) Borrar datos en CW excepto cabecera (fila 1)
    LimpiarDatosExceptoCabecera wsCW

    ' 5) Copiar datos de CSV -> CW sin cabecera
    CopiarDatosSinCabecera wsCSV, wsCW

    ' 6) Guardar CW/PW y cerrar CW, PW y CSV
    GuardarLibro wbCW, "CW_ES_Extract_Test.xlsm"
    GuardarLibro wbPW, "PW_ES_Extract_Test.xlsm"

    wbCW.Close SaveChanges:=False
    wbPW.Close SaveChanges:=False
    wbCSV.Close SaveChanges:=False

SalidaOK:
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    MsgBox "Proceso completado correctamente.", vbInformation
    Exit Sub

ControlError:
    Dim msg As String
    msg = "Error " & Err.Number & ": " & Err.Description

    On Error Resume Next
    If Not wbCSV Is Nothing Then wbCSV.Close SaveChanges:=False

    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    MsgBox msg, vbCritical, "Proceso_Semanal_Iberia"
End Sub

Private Function ObtenerHoja(ByVal wb As Workbook, ByVal nombreHoja As String, ByVal obligatorio As Boolean) As Worksheet
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = wb.Worksheets(nombreHoja)
    On Error GoTo 0

    If Not ws Is Nothing Then
        Set ObtenerHoja = ws
        Exit Function
    End If

    If obligatorio Then
        Err.Raise vbObjectError + 1001, "ObtenerHoja", _
                  "No se encontró la hoja '" & nombreHoja & "' en el archivo: " & wb.Name
    Else
        ' Para CSV, algunos archivos abren con otro nombre de hoja. Usamos la primera hoja como fallback.
        Set ObtenerHoja = wb.Worksheets(1)
    End If
End Function

Private Sub GuardarLibro(ByVal wb As Workbook, ByVal etiqueta As String)
    Dim i As Long
    Dim guardadoOK As Boolean
    Dim ultimoError As String

    On Error Resume Next

    ' 1) Intentar pasar a lectura/escritura si vino en solo lectura
    If wb.ReadOnly Then wb.ChangeFileAccess xlReadWrite

    If wb.ReadOnly Then
        On Error GoTo 0
        Err.Raise vbObjectError + 1002, "GuardarLibro", _
                  "No se puede guardar " & etiqueta & " porque está en SOLO LECTURA. " & _
                  "Revisa bloqueo/check-out/permisos en Teams-SharePoint."
    End If

    ' 2) Reintentos de Save (fallos transitorios de sincronización)
    For i = 1 To 3
        Err.Clear
        wb.Save
        If Err.Number = 0 Then
            guardadoOK = True
            Exit For
        Else
            ultimoError = Err.Description
            DoEvents
            Application.Wait Now + TimeValue("00:00:02")
        End If
    Next i

    ' 3) Fallback para ubicaciones web: SaveAs sobre la misma ruta
    If Not guardadoOK Then
        Err.Clear
        wb.SaveAs Filename:=wb.FullName, FileFormat:=wb.FileFormat, ConflictResolution:=xlLocalSessionChanges, AddToMru:=False
        If Err.Number = 0 Then
            guardadoOK = True
        Else
            If Len(ultimoError) = 0 Then ultimoError = Err.Description
        End If
    End If

    On Error GoTo 0

    If Not guardadoOK Then
        Err.Raise vbObjectError + 1003, "GuardarLibro", _
                  "Falló el guardado de " & etiqueta & " (" & wb.Name & "). " & _
                  "Último detalle: " & ultimoError & ". " & _
                  "Sugerencia: abrir manualmente el archivo en Excel, confirmar edición y guardar una vez."
    End If
End Sub

Private Sub LimpiarDatosExceptoCabecera(ByVal ws As Worksheet)
    Dim ultimaFila As Long

    ultimaFila = UltimaFilaEnAK(ws)
    If ultimaFila < 2 Then Exit Sub

    ' Limpia solo A:K desde fila 2 (cabecera intacta)
    ws.Range("A2:K" & ultimaFila).ClearContents
End Sub

Private Sub CopiarDatosSinCabecera(ByVal wsOrigen As Worksheet, ByVal wsDestino As Worksheet)
    Dim ufOri As Long

    ufOri = UltimaFilaEnAK(wsOrigen)

    ' Si origen solo tiene cabecera o está vacío, no hay nada que copiar
    If ufOri < 2 Then Exit Sub

    ' Pegamos en destino desde A2, solo columnas A:K (11 columnas)
    wsDestino.Range("A2").Resize(ufOri - 1, 11).Value = _
        wsOrigen.Range("A2").Resize(ufOri - 1, 11).Value
End Sub

Private Function UltimaFilaEnAK(ByVal ws As Worksheet) As Long
    Dim c As Range

    Set c = ws.Range("A:K").Find(What:="*", After:=ws.Range("A1"), LookIn:=xlFormulas, _
                                  LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    If c Is Nothing Then
        UltimaFilaEnAK = 1
    Else
        UltimaFilaEnAK = c.Row
    End If
End Function
```

---

## Notas importantes para que funcione a la primera
- Debes estar autenticado en Microsoft 365 para abrir URLs de SharePoint/Teams.
- La unidad `X:` debe estar mapeada y accesible en tu equipo (VPN si aplica).
- La macro usa estas hojas específicas:
  - Primer fichero (CW): `BOARD reporting extract`
  - Segundo fichero (PW): `BOARD reporting extract`
  - Tercer fichero (CSV): `BOARD reporting extract_Iberia`
- Todas las operaciones de copia/pegado y limpieza trabajan exclusivamente con columnas **A:K** y número de filas variable en cada ejecución.
- Si aparece `Error 9: Subscript out of range`, normalmente es porque el nombre de hoja no coincide. Esta versión valida las hojas y en el CSV usa fallback a la primera hoja.
- El CSV se abre con `OpenText` y delimitador `;` para que los datos queden separados en columnas (A:K) y no concatenados en la columna A.
- El guardado de CW y PW intenta: (1) pasar a lectura/escritura, (2) reintentar `Save` 3 veces y (3) fallback con `SaveAs` sobre la misma ruta; si falla, devuelve diagnóstico ampliado.
- Al finalizar correctamente, la macro guarda y cierra automáticamente `CW`, `PW` y `CSV`.
- Si quieres, en un siguiente paso te puedo dar una versión que además registre en un log: fecha/hora, filas copiadas y posibles errores.
