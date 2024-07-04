Imports Microsoft.Office.Interop.Excel
Imports DocumentFormat.OpenXml.Spreadsheet
Imports DocumentFormat.OpenXml.Packaging
Imports System.Linq
Imports System.Data
Imports System.Collections.Generic
Imports Autodesk.AutoCAD.ApplicationServices
Imports System.IO
Imports System.Data.SqlClient
Imports System.Windows.Forms
Imports System.ComponentModel.ListSortDirection
Imports Newtonsoft.Json
Imports System.Configuration
Imports CreaRQI.CapaDatos
Imports System.Runtime.InteropServices

Public Class Frm_CargaListados

    Dim Conn As ConexionBD = New ConexionBD 'Instanciamos la clase de conexion
    Dim Orden As Orden = New Orden 'Instanciamos la clase Orden
    Public Shared grills() As Object
    Public Shared datagridviewuno As System.Windows.Forms.DataGridView
    Public Shared datagridviewdos As System.Windows.Forms.DataGridView
    Public Shared datagridviewtres As System.Windows.Forms.DataGridView
    Public Shared tabcontroluno As System.Windows.Forms.TabControl
    Public Shared libroExcel As String
    Public Shared mensa As Microsoft.VisualBasic.MsgBoxResult
    Public ruta As String
    Public mensajes As Boolean = True
    Dim ancho As Integer = 43
    Dim impidereemplazar As Boolean
    Public Shared CIA As String
    Public Shared status As String
    Dim entregada, rayalibre, estatuserplibre As Boolean
    Dim rayavacia As Integer = 0
    Dim rayasantesdevacia As Integer = 0
    Dim rayaimparvacia As Integer = 0
    Dim rayaparvacia As Integer = 0
    Dim semaforo As String 'vacio abierto modificable cerrado
    Public loginvalidez As Boolean
    Public darusuario, darpassw As String
    Dim muestrafiltro As Boolean = True
    Public unidadNegocio As Integer
    Dim itemssincodigo As Boolean
    Dim campoPlanoConFormato As Boolean 'verifica si el campo plano tiene P o tiene otro texto extraño
    Dim mensajesfinales As String()
    Dim PlantaId As Integer = 0 'Planta seleccionada por el usuario
    Public Str_Con_Bd As String = ConexionBD.getStringConexion()


    Sub LlenaGrillas()

        TextBox4.Text = 0
        'Do While DataGridView1.RowCount > 1
        'For i = 0 To DataGridView1.RowCount - 1
        'DataGridView1.Rows.Remove(DataGridView1.Rows.Item(DataGridView1.RowCount - 1))
        'Next
        'Loop
        'Do While DataGridView2.RowCount > 1
        'For i = 0 To DataGridView2.RowCount - 1
        'DataGridView2.Rows.Remove(DataGridView2.Rows.Item(DataGridView2.RowCount - 1))
        'Next
        'Loop
        'Do While DataGridView3.RowCount > 1
        'For i = 0 To DataGridView2.RowCount - 1
        'DataGridView3.Rows.Remove(DataGridView3.Rows.Item(DataGridView3.RowCount - 1))
        'Next- nk  n
        'Loop

        DataGridView1.Rows.Clear()
        If DataGridView1.RowCount < 1 Then
            DataGridView1.Rows.Add()
        End If
        DataGridView2.Rows.Clear()
        If DataGridView2.RowCount < 1 Then
            DataGridView2.Rows.Add()
        End If
        DataGridView3.Rows.Clear()
        If DataGridView3.RowCount < 1 Then
            DataGridView3.Rows.Add()
        End If
        DataGridView5.Rows.Clear()
        If DataGridView5.RowCount < 1 Then
            DataGridView5.Rows.Add()
        End If
        'acc
        DataGridView4.Rows.Clear()
        If DataGridView4.RowCount < 1 Then
            DataGridView4.Rows.Add()
        End If
        DataGridView7.Rows.Clear()
        If DataGridView7.RowCount < 1 Then
            DataGridView7.Rows.Add()
        End If

        Button12.Enabled = True
        campoPlanoConFormato = True

        ancho = 890
        DataGridView1.Width = ancho
        DataGridView2.Width = ancho
        DataGridView3.Width = ancho
        DataGridView5.Width = ancho
        DataGridView4.Width = ancho 'acc
        DataGridView7.Width = ancho 'acc

        DataGridView6.RowCount = 5
        DataGridView6.Rows.Item(0).Cells(0).Value = "MUROS"
        DataGridView6.Rows.Item(1).Cells(0).Value = "UNION"
        DataGridView6.Rows.Item(2).Cells(0).Value = "LOSAS"
        DataGridView6.Rows.Item(3).Cells(0).Value = "CULAT"
        DataGridView6.Rows.Item(4).Cells(0).Value = "TOTAL"

        DataGridView6.Rows.Item(0).DefaultCellStyle.BackColor = System.Drawing.Color.LightGray
        DataGridView6.Rows.Item(1).DefaultCellStyle.BackColor = System.Drawing.Color.LightGray
        DataGridView6.Rows.Item(2).DefaultCellStyle.BackColor = System.Drawing.Color.LightGray
        DataGridView6.Rows.Item(3).DefaultCellStyle.BackColor = System.Drawing.Color.LightGray

        DataGridView8.Rows.Add()
        DataGridView8.Rows.Add()
        DataGridView8.Rows.Item(0).Cells(0).Value = "Almacen"
        DataGridView8.Rows.Item(1).Cells(0).Value = "Planta"
        DataGridView8.Rows.Item(2).Cells(0).Value = "Total"
        DataGridView8.Rows.Item(0).Cells(1).Value = 0
        DataGridView8.Rows.Item(1).Cells(1).Value = 0
        DataGridView8.Rows.Item(2).Cells(1).Value = 0

        'Dim vecsheets As String() 'crea vector de hojas de excel
        Dim usarxml As Boolean
        Dim xlWorkBook As SpreadsheetDocument
        If ruta = Nothing Then
            Exit Sub
        Else
            itemssincodigo = False
            Try
                xlWorkBook = SpreadsheetDocument.Open(ruta, False)
                usarxml = True
                'End Using
            Catch ex As Exception
                usarxml = False
            End Try

            If usarxml = True Then
                Dim sheet As Sheet = xlWorkBook.WorkbookPart.Workbook.Sheets.GetFirstChild(Of Sheet)()
                Dim nombrearchivosplit = Split(ruta, "\")
                Dim nombrearchivo = nombrearchivosplit(UBound(nombrearchivosplit))
                Label1.Text = nombrearchivo
                Label2.Text = nombrearchivo
                Label9.Text = nombrearchivo

                Dim workbookPart As WorkbookPart = xlWorkBook.WorkbookPart
                Dim worksheetPart As WorksheetPart


                'Dim xlWorkSheet As DocumentFormat.OpenXml.Spreadsheet.Worksheet
                For Each esheet As Sheet In xlWorkBook.WorkbookPart.Workbook.Sheets
                    If esheet.Name = "LISTADO_INICIAL" Then
                        'xlWorkSheet = TryCast(xlWorkBook.WorkbookPart.GetPartById(esheet.Id.Value), WorksheetPart).Worksheet
                        worksheetPart = xlWorkBook.WorkbookPart.GetPartById(esheet.Id.Value)
                        Exit For
                    Else
                        'xlWorkSheet = TryCast(xlWorkBook.WorkbookPart.GetPartById(esheet.Id.Value), WorksheetPart).Worksheet 'xlWorkBook.Worksheets("ROBLES_FUNDICION 1")
                        worksheetPart = xlWorkBook.WorkbookPart.GetPartById(esheet.Id.Value)
                    End If
                Next
                'Dim rows As IEnumerable(Of DocumentFormat.OpenXml.Spreadsheet.Row) = xlWorkSheet.Elements(Of DocumentFormat.OpenXml.Spreadsheet.Row)() 'xlWorkSheet.GetFirstChild(Of SheetData)().Descendants(Of DocumentFormat.OpenXml.Spreadsheet.Row)()
                'Dim workbookPart As WorkbookPart = xlWorkBook.WorkbookPart
                'Dim worksheetPart As WorksheetPart = workbookPart.WorksheetParts.First()
                Dim sheetData As SheetData = worksheetPart.Worksheet.Elements(Of SheetData)().First()

                Dim i As Integer = sheetData.Descendants(Of DocumentFormat.OpenXml.Spreadsheet.Row)().LastOrDefault().RowIndex.Value - 2
                While i > 1
                    Try
                        Dim RowCnt As DocumentFormat.OpenXml.Spreadsheet.Row = sheetData.Elements(Of DocumentFormat.OpenXml.Spreadsheet.Row)().ElementAt(i)
                        'Dim CellCnt As DocumentFormat.OpenXml.Spreadsheet.Cell = RowCnt.Elements(Of DocumentFormat.OpenXml.Spreadsheet.Cell)().ElementAt(0)
                        Dim texto = RowCnt.InnerText
                        If texto = "" Then
                            i = i - 1
                        Else
                            Exit While
                        End If
                    Catch ex As Exception
                        Exit While
                    End Try
                End While
                Dim progreso As New Frm_BarraCarga
                progreso.ProgressBar1.Minimum = 0
                progreso.ProgressBar1.Maximum = i + 1
                progreso.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
                Me.Hide()
                progreso.Show()
                progreso.TopMost = True
                progreso.Text = "cargando lista de items"


                For Each RowCnt As DocumentFormat.OpenXml.Spreadsheet.Row In sheetData.Elements(Of DocumentFormat.OpenXml.Spreadsheet.Row)()
                    Dim numfila = RowCnt.RowIndex.Value
                    If RowCnt.InnerText = "" Then
                        Continue For
                    End If
                    Dim datosfila() As String

                    ReDim datosfila(10)
                    'datosfila(0) = cant
                    'datosfila(1) = nomenclatura
                    'datosfila(2) = fm ancho1   acc desaux
                    'datosfila(3) = fm alto1    acc dim1
                    'datosfila(4) = fm alto2    acc dim2
                    'datosfila(5) = fm ancho2   acc dim3
                    'datosfila(6) = fm plano    acc dim4
                    'datosfila(7) = fm desaux   acc dim5
                    'datosfila(8) = fm familia  acc null
                    'datosfila(9) = fm areauni  acc null
                    'datosfila(10) = fm aretot  acc null

                    For Each CellCnt As DocumentFormat.OpenXml.Spreadsheet.Cell In RowCnt.Elements(Of DocumentFormat.OpenXml.Spreadsheet.Cell)()
                        Dim letra = CellCnt.CellReference.ToString.Substring(0, 1)
                        Select Case letra

                            Case "A"
                                datosfila(0) = xmlGetValue(xlWorkBook, CellCnt)
                            Case "B"
                                datosfila(1) = xmlGetValue(xlWorkBook, CellCnt)
                            Case "C"
                                datosfila(2) = xmlGetValue(xlWorkBook, CellCnt)
                            Case "D"
                                datosfila(3) = xmlGetValue(xlWorkBook, CellCnt)
                            Case "E"
                                datosfila(4) = xmlGetValue(xlWorkBook, CellCnt)
                            Case "F"
                                datosfila(5) = xmlGetValue(xlWorkBook, CellCnt)
                            Case "G"
                                datosfila(6) = xmlGetValue(xlWorkBook, CellCnt)
                            Case "H"
                                datosfila(7) = xmlGetValue(xlWorkBook, CellCnt)
                            Case "I"
                                datosfila(8) = xmlGetValue(xlWorkBook, CellCnt)
                            Case "J"
                                datosfila(9) = xmlGetValue(xlWorkBook, CellCnt)
                            Case "K"
                                datosfila(10) = xmlGetValue(xlWorkBook, CellCnt)
                                Exit For
                        End Select
                        'If CellCnt.CellReference.ToString.Contains("A") = True Then
                        '    datosfila(0) = xmlGetValue(xlWorkBook, CellCnt)
                        'ElseIf CellCnt.CellReference.ToString.Contains("B") = True Then
                        '    datosfila(1) = xmlGetValue(xlWorkBook, CellCnt)
                        'ElseIf CellCnt.CellReference.ToString.Contains("C") = True Then
                        '    datosfila(2) = xmlGetValue(xlWorkBook, CellCnt)
                        'ElseIf CellCnt.CellReference.ToString.Contains("D") = True Then
                        '    datosfila(3) = xmlGetValue(xlWorkBook, CellCnt)
                        'ElseIf CellCnt.CellReference.ToString.Contains("E") = True Then
                        '    datosfila(4) = xmlGetValue(xlWorkBook, CellCnt)
                        'ElseIf CellCnt.CellReference.ToString.Contains("F") = True Then
                        '    datosfila(5) = xmlGetValue(xlWorkBook, CellCnt)
                        'ElseIf CellCnt.CellReference.ToString.Contains("G") = True Then
                        '    datosfila(6) = xmlGetValue(xlWorkBook, CellCnt)
                        'ElseIf CellCnt.CellReference.ToString.Contains("H") = True Then
                        '    datosfila(7) = xmlGetValue(xlWorkBook, CellCnt)
                        'ElseIf CellCnt.CellReference.ToString.Contains("I") = True Then
                        '    datosfila(8) = xmlGetValue(xlWorkBook, CellCnt)
                        'ElseIf CellCnt.CellReference.ToString.Contains("J") = True Then
                        '    datosfila(9) = xmlGetValue(xlWorkBook, CellCnt)
                        'ElseIf CellCnt.CellReference.ToString.Contains("K") = True Then
                        '    datosfila(10) = xmlGetValue(xlWorkBook, CellCnt)
                        '    Exit For
                        'End If

                    Next
                    datosagrilla(datosfila)
                    'Exit For
                    progreso.Text = "cargando lista de items: Fila " & numfila & "     " & datosfila(0) & " " & datosfila(1) & " " & datosfila(2) & " " & datosfila(3)
                    If progreso.ProgressBar1.Value = i + 1 Then
                        Exit For
                    Else
                        progreso.ProgressBar1.Value = progreso.ProgressBar1.Value + 1
                    End If

                Next
                progreso.Close()
            Else
                Dim numfila As Integer = 0
                Dim lineatxt As String
                Dim txtworkbook As System.IO.StreamReader
                Try
                    txtworkbook = New System.IO.StreamReader(ruta, System.Text.Encoding.Default)
                Catch ex As Exception
                    If ex.ToString.Contains("because it is being used by another process") Then
                        MsgBox("debe cerrar el excel antes de montarlo al cargasif, no se puede leer el archivo porque está siendo usado por otro proceso",, "error de usuario")
                    End If
                    Exit Sub
                End Try

                Dim numfilas As Integer = File.ReadAllLines(ruta).Length 'cuenta las filas del archivo plano

                Dim nombrearchivosplit = Split(ruta, "\")
                Dim nombrearchivo = nombrearchivosplit(UBound(nombrearchivosplit))
                Label1.Text = nombrearchivo
                Label2.Text = nombrearchivo
                Label9.Text = nombrearchivo

                Dim progreso As New Frm_BarraCarga
                progreso.ProgressBar1.Minimum = 0

                progreso.ProgressBar1.Maximum = numfilas + 1
                progreso.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
                Me.Hide()
                progreso.Show()
                progreso.TopMost = True
                progreso.Text = "cargando lista de items"

                lineatxt = txtworkbook.ReadLine

                While lineatxt IsNot Nothing
                    numfila = numfila + 1
                    Dim datosfila(), datoslinea() As String

                    ReDim datosfila(10)
                    'datosfila(0) = cant
                    'datosfila(1) = nomenclatura
                    'datosfila(2) = fm ancho1   acc desaux
                    'datosfila(3) = fm alto1    acc dim1
                    'datosfila(4) = fm alto2    acc dim2
                    'datosfila(5) = fm ancho2   acc dim3
                    'datosfila(6) = fm plano    acc dim4
                    'datosfila(7) = fm desaux   acc dim5
                    'datosfila(8) = fm familia  acc null
                    'datosfila(9) = fm areauni  acc null
                    'datosfila(10) = fm aretot  acc null

                    datoslinea = Split(lineatxt, vbTab)
                    For i = 0 To 10
                        Try
                            datosfila(i) = datoslinea(i)
                        Catch ex2 As Exception

                        End Try
                    Next
                    datosagrilla(datosfila)
                    'Exit For

                    lineatxt = txtworkbook.ReadLine

                    progreso.Text = "cargando lista de items: Fila " & numfila & "     " & datosfila(0) & " " & datosfila(1) & " " & datosfila(2) & " " & datosfila(3)

                    If progreso.ProgressBar1.Value = 1000 Then
                    Else
                        progreso.ProgressBar1.Value = progreso.ProgressBar1.Value + 1
                    End If
                End While
                txtworkbook.Close()
                progreso.Close()
            End If
            'Dim xlApp = New Microsoft.Office.Interop.Excel.Application 'Excel.ApplicationClass
            'Dim xlWorkBook As Microsoft.Office.Interop.Excel.Workbook

            'xlWorkBook = xlApp.Workbooks.Open(ruta) '"C:\Users\hugogomez\Desktop\ROBLES_FUNDICION 1.xls"
        End If
        'Label1.Text = xlWorkBook.Name
        'Label2.Text = xlWorkBook.Name
        'Label9.Text = xlWorkBook.Name



        'Dim xlWorkSheet As Microsoft.Office.Interop.Excel.Worksheet
        'For Each esheet As Microsoft.Office.Interop.Excel.Worksheet In xlWorkBook.Worksheets
        'ReDim Preserve vecsheets(UBound())
        'If esheet.Name = "LISTADO_INICIAL" Then
        'xlWorkSheet = esheet
        'Exit For
        'Else
        'xlWorkSheet = esheet 'xlWorkBook.Worksheets("ROBLES_FUNDICION 1")
        'End If
        'Next
        'Dim nombrelibro() As String = Split(xlWorkBook.Name, ".") *
        ' crea una copia del libro que está abriendo
        'xlWorkBook.SaveAs(xlWorkBook.Path & "\" & xlWorkBook.Name & "x", Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook)
        'libroExcel = xlWorkBook.Path & "\" & xlWorkBook.Name
        'Dim range = xlWorkSheet.UsedRange
        'Dim Obj As Microsoft.Office.Interop.Excel.Range
        'Dim Obj2 As Microsoft.Office.Interop.Excel.Range

        'Dim progreso As New Form8
        'progreso.ProgressBar1.Minimum = 0
        'progreso.ProgressBar1.Maximum = range.Rows.Count
        'progreso.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        'Me.Hide()
        'progreso.Show()
        'progreso.TopMost = True
        'progreso.Text = "cargando lista de items"

        'For rowCnt = 1 To range.Rows.Count
        'Obj = CType(range.Cells(rowCnt, 9), Microsoft.Office.Interop.Excel.Range)
        'Obj2 = CType(range.Cells(rowCnt, 1), Microsoft.Office.Interop.Excel.Range)

        'Try
        'If CStr(Obj.Value) = "MUROS" Then
        'cargalinea(Obj, Obj2, Range, rowCnt, DataGridView1)
        'ElseIf CStr(Obj.Value) = "UNION" Then
        'cargalinea(Obj, Obj2, range, rowCnt, DataGridView2)
        'ElseIf CStr(Obj.Value) = "LOSAS" Then
        'cargalinea(Obj, Obj2, range, rowCnt, DataGridView3)
        'ElseIf CStr(Obj.Value) = "CULAT" Then
        'cargalinea(Obj, Obj2, range, rowCnt, DataGridView5)
        'ElseIf CInt(Obj2.Value) > 0 Then
        'Dim codigo, paroimpar, plantaid, nomlargo, cantvalida As String
        'If Label11.Text = "Colombia" Then
        'plantaid = 1
        'ElseIf Label11.Text = "Mexico" Then
        'plantaid = 2
        'ElseIf Label11.Text = "Brasil" Then
        'plantaid = 3
        'End If
        'elijecodigo(codigo, paroimpar, range, rowCnt, plantaid, cantvalida)
        'If codigo = 0 Then
        'nomlargo = ""
        'Else
        'nomlargo = traenomlargo(codigo, plantaid)
        'End If
        '
        'If paroimpar = "ALMACEN" Then
        'cargalineaacc(Obj, Obj2, range, rowCnt, DataGridView4, codigo, nomlargo, cantvalida) 'almacen
        'DataGridView8.Rows.Item(0).Cells(1).Value = DataGridView8.Rows.Item(0).Cells(1).Value + 1
        'DataGridView8.Rows.Item(2).Cells(1).Value = DataGridView8.Rows.Item(2).Cells(1).Value + 1
        'Else
        'cargalineaacc(Obj, Obj2, range, rowCnt, DataGridView7, codigo, nomlargo, cantvalida) 'planta
        'DataGridView8.Rows.Item(1).Cells(1).Value = DataGridView8.Rows.Item(1).Cells(1).Value + 1
        'DataGridView8.Rows.Item(2).Cells(1).Value = DataGridView8.Rows.Item(2).Cells(1).Value + 1
        'End If
        'codigo = ""
        'paroimpar = ""
        'End If
        'Catch ex As Exception

        'End Try

        'progreso.ProgressBar1.Value = rowCnt

        'Next
        If ComboBox7.SelectedItem = "CT" And ComboBox5.Text = "51" And ComboBox10.SelectedItem = 1 Then
            Dim codigo, nomlargo, plantaid, paroimpar, cant, nom, ordennumero As String
            plantaid = 1
            ordennumero = ComboBox6.SelectedItem
            ordennumero = ordennumero.Substring(0, ordennumero.IndexOf("-")).Trim()
            Dim cons = "WITH EC AS (SELECT TOP (1) eect_id FROM fup_enc_entrada_cotizacion ec INNER JOIN Orden_seg os ON os.[fup] = ec.eect_fup_id AND os.[vers] = ec.eect_vercot_id INNER JOIN Orden o ON os.[Id_Seg_Of] = o.[ordenseg_id] WHERE (o.Numero = '" & ordennumero & "') AND (o.Tipo_Sol = '" & ComboBox7.SelectedItem & "') group by ec.eect_id ) SELECT e.eect_fup_id Fup_ID, e.eect_vercot_id Version ,cp.fcp_CodigoAccesorio Id_UnoE ,cp.fcp_Item ,[apa_AdicionalCantidad] * CASE WHEN fcp_CodigoAccesorio2 IS NULL THEN 1 ELSE 0.5 END CantidadAdicional ,[apa_Descripcion] Descripcion ,d.fdom_Descripcion Desc_solicitado FROM fup_enc_entrada_cotizacion e INNER JOIN EC ON E.eect_id = EC.eect_id left outer join [dbo].[fup_alcance_partes] t ON e.[eect_id] = t.[apa_enc_entrada_cot_id] left outer join fup_ConfiguracionPartes cp on cp.fcp_id = t.apa_Itemparte_id left outer join [dbo].[fup_Dominios] d on cp.fcp_ItemDominio = d.fdom_Dominio and d.fdom_CodDominio = t.apa_ItemTextoLista WHERE cp.fcp_IncluirAdicional = 1 AND (isnull(cp.fcp_CodigoAccesorio,'0') <> '0' ) AND d.fdom_EsAccesorio = 1 AND t.apa_AdicionalSiNo = 1 UNION ALL SELECT e.eect_fup_id Fup_ID, e.eect_vercot_id Version ,cp.fcp_CodigoAccesorio2 Id_UnoE ,cp.fcp_Item ,[apa_AdicionalCantidad] * CASE WHEN fcp_CodigoAccesorio IS NULL THEN 1 ELSE 0.5 END CantidadAdicional ,[apa_Descripcion] Descripcion ,d.fdom_Descripcion Desc_solicitado FROM fup_enc_entrada_cotizacion e INNER JOIN EC ON E.eect_id = EC.eect_id left outer join [dbo].[fup_alcance_partes] t ON e.[eect_id] = t.[apa_enc_entrada_cot_id] left outer join fup_ConfiguracionPartes cp on cp.fcp_id = t.apa_Itemparte_id left outer join [dbo].[fup_Dominios] d on cp.fcp_ItemDominio = d.fdom_Dominio and d.fdom_CodDominio = t.apa_ItemTextoLista WHERE cp.fcp_IncluirAdicional = 1 AND (isnull(cp.fcp_CodigoAccesorio2,'0') <> '0' ) AND d.fdom_EsAccesorio = 1 AND t.apa_AdicionalSiNo = 1 UNION ALL SELECT B.eect_fup_id, B.eect_vercot_id, fic_CodigoERP, C.fic_ItemCotizacion, fcc_Cantidad, fcc_ItemCotiza, fcc_Observacion FROM fup_CartaCierre_partes A INNER JOIN fup_enc_entrada_cotizacion B ON B.eect_id = A.fcc_enc_entrada_cot_id INNER JOIN EC ON B.eect_id = EC.eect_id INNER JOIN fup_ItemCotizacion C ON C.fic_id = A.fcc_Item_id AND isnull(C.fic_CodigoERP,'0') <> '0'"
            Using connection As New SqlConnection(Str_Con_Bd)
                Dim command As SqlCommand = New SqlCommand(cons, connection)

                connection.Open()
                Dim reader As SqlDataReader = command.ExecuteReader()
                If reader.HasRows Then
                    Do While reader.Read()
                        codigo = reader.GetValue(2)
                        nomlargo = traenomlargo(codigo, plantaid, paroimpar, nom)
                        cant = CInt(reader.GetValue(4))
                        If paroimpar = "ALMACEN" Then
                            cargalineafup(nom, cant, DataGridView4, codigo, nomlargo, 0) 'almacen
                        Else
                            cargalineafup(nom, cant, DataGridView7, codigo, nomlargo, 0) 'planta
                        End If
                    Loop
                End If
            End Using

        End If

        DataGridView1.AllowUserToAddRows = False
        DataGridView2.AllowUserToAddRows = False
        DataGridView3.AllowUserToAddRows = False
        DataGridView5.AllowUserToAddRows = False
        'xlWorkBook.Close(False)
        'xlApp.Quit()
        If campoPlanoConFormato = False Then
            mensajesfinales(UBound(mensajesfinales)) = "la fila de plano (columna G en el listado) solo debe decir 'P' para indicar que el item requiere plano o vacío para indicar que no lo requiere"
            ReDim Preserve mensajesfinales(UBound(mensajesfinales) + 1)
        End If

        Me.Show()

        mensajefinal()



        Try
            Do While DataGridView1.Rows.Item(DataGridView1.RowCount - 1).Cells.Item(0).Value = Nothing
                DataGridView1.Rows.Remove(DataGridView1.Rows.Item(DataGridView1.RowCount - 1))
            Loop
            Tab_VerInfoCargada.SelectedTab = TabPage_Muros
        Catch ex As Exception

        End Try
        Try
            Do While DataGridView2.Rows.Item(DataGridView2.RowCount - 1).Cells.Item(0).Value = Nothing
                DataGridView2.Rows.Remove(DataGridView2.Rows.Item(DataGridView2.RowCount - 1))
            Loop
            Tab_VerInfoCargada.SelectedTab = TabPage_Union
        Catch ex As Exception

        End Try

        Try
            Do While DataGridView3.Rows.Item(DataGridView3.RowCount - 1).Cells.Item(0).Value = Nothing
                DataGridView3.Rows.Remove(DataGridView3.Rows.Item(DataGridView3.RowCount - 1))
            Loop
            Tab_VerInfoCargada.SelectedTab = TabPage_Losas
        Catch ex As Exception

        End Try
        Try
            Do While DataGridView5.Rows.Item(DataGridView5.RowCount - 1).Cells.Item(0).Value = Nothing
                DataGridView5.Rows.Remove(DataGridView5.Rows.Item(DataGridView5.RowCount - 1))
            Loop
            Tab_VerInfoCargada.SelectedTab = TabPage_Culat
        Catch ex As Exception

        End Try
        Try
            Do While DataGridView4.Rows.Item(DataGridView4.RowCount - 1).Cells.Item(0).Value = Nothing
                DataGridView4.Rows.Remove(DataGridView4.Rows.Item(DataGridView4.RowCount - 1))
            Loop
            Tab_VerInfoCargada.SelectedTab = TabPage_ACC_Ver
        Catch ex As Exception

        End Try
        Try
            Do While DataGridView7.Rows.Item(DataGridView7.RowCount - 1).Cells.Item(0).Value = Nothing
                DataGridView7.Rows.Remove(DataGridView7.Rows.Item(DataGridView7.RowCount - 1))
            Loop
            Tab_VerInfoCargada.SelectedTab = TabPage_ACC_Ver
        Catch ex As Exception

        End Try

        Tab_VerInfoCargada.SelectedTab = TabPage_Muros
        'ordenar()

        'progreso.Close()
        'Me.Show()

    End Sub
    Public Sub mensajefinal()
        If UBound(mensajesfinales) > 5 Then
            If UBound(mensajesfinales) > 0 Then
                Try
                    Do While UBound(mensajesfinales) = ""
                        ReDim mensajesfinales(UBound(mensajesfinales) - 1)
                    Loop
                Catch ex As Exception
                    Do While UBound(mensajesfinales) = Nothing
                        ReDim mensajesfinales(UBound(mensajesfinales) - 1)
                    Loop
                End Try
            End If
            Try
                If mensajesfinales(0) <> "" Then
                    'Dim formamensajes As New Form10
                    'formamensajes.mensajes = mensajesfinales
                    'formamensajes.StartPosition = FormStartPosition.CenterScreen
                    'formamensajes.ShowDialog()
                    Dim FN As Integer
                    FN = FreeFile()
                    Using sw As New StreamWriter("C:\BLOQUES\FORLINE\INVENTOR\Libraries\mensajefinalcargasif.txt", False, System.Text.Encoding.Default) '= File.CreateText(filename)

                        For i = 0 To UBound(mensajesfinales)
                            sw.WriteLine(mensajesfinales(i))
                        Next

                    End Using
                    System.Diagnostics.Process.Start("notepad.exe", "C:\BLOQUES\FORLINE\INVENTOR\Libraries\mensajefinalcargasif.txt")
                    Threading.Thread.Sleep(1000 * 2)
                    My.Computer.FileSystem.DeleteFile("C:\BLOQUES\FORLINE\INVENTOR\Libraries\mensajefinalcargasif.txt")
                End If
            Catch ex As Exception

            End Try
            ReDim Preserve mensajesfinales(4)
            ReDim Preserve mensajesfinales(5)
        End If
    End Sub
    Sub datosagrilla(ByVal datosfila As String())


        Dim intcant As Integer

        If datosfila(8) = "MUROS" Then
            cargalinea(datosfila, DataGridView1)
        ElseIf datosfila(8) = "UNION" Then
            cargalinea(datosfila, DataGridView2)
        ElseIf datosfila(8) = "LOSAS" Then
            cargalinea(datosfila, DataGridView3)
        ElseIf datosfila(8) = "CULAT" Then
            cargalinea(datosfila, DataGridView5)
        ElseIf Integer.TryParse(datosfila(0), intcant) = True And intcant > 0 And Button21.Enabled = True Then
            Dim codigo, paroimpar, plantaid, nomlargo, cantvalida, codigosId As String
            If Label11.Text = "Colombia" Then
                plantaid = 1
            ElseIf Label11.Text = "Mexico" Then
                plantaid = 2
            ElseIf Label11.Text = "Brasil" Then
                plantaid = 3
            End If
            codigosId = "0"

            If datosfila(2) IsNot Nothing Then

                datosfila(2) = datosfila(2).Trim()

            End If

            elijecodigo(codigo, paroimpar, datosfila, plantaid, cantvalida, True, codigosId)
            If codigo = 0 Then
                elijecodigo(codigo, paroimpar, datosfila, plantaid, cantvalida, False, codigosId)
                If codigo <> 0 Then
                    Dim descri, desaux, dim1, dim2, dim3, dim4, dim5, dim6 As String
                    descri = datosfila(1)
                    desaux = datosfila(2)
                    If datosfila(3) = "" Then dim1 = 0 Else dim1 = datosfila(3)
                    If datosfila(4) = "" Then dim2 = 0 Else dim2 = datosfila(4)
                    If datosfila(5) = "" Then dim3 = 0 Else dim3 = datosfila(5)
                    If datosfila(6) = "" Then dim4 = 0 Else dim4 = datosfila(6)
                    If datosfila(7) = "" Then dim5 = 0 Else dim5 = datosfila(7)
                    If datosfila(8) = "" Then dim6 = 0 Else dim6 = datosfila(8)
                    mensajesfinales(UBound(mensajesfinales)) = "el item " & descri & "(" & desaux & ") " & dim1 & "x" & dim2 & "x" & dim3 & "x" & dim4 & "x" & dim5 & "x" & dim6 & " tiene el código " & codigo & " pero está inactivo"
                    ReDim Preserve mensajesfinales(UBound(mensajesfinales) + 1)
                    codigo = 0
                End If
                nomlargo = ""
            Else
                nomlargo = traenomlargo(codigo, plantaid)
            End If

            If paroimpar = "ALMACEN" Then
                cargalineaacc(datosfila, DataGridView4, codigo, nomlargo, cantvalida, codigosId) 'almacen
                DataGridView8.Rows.Item(0).Cells(1).Value = DataGridView8.Rows.Item(0).Cells(1).Value + 1
                DataGridView8.Rows.Item(2).Cells(1).Value = DataGridView8.Rows.Item(2).Cells(1).Value + 1
            Else
                cargalineaacc(datosfila, DataGridView7, codigo, nomlargo, cantvalida, codigosId) 'planta
                DataGridView8.Rows.Item(1).Cells(1).Value = DataGridView8.Rows.Item(1).Cells(1).Value + 1
                DataGridView8.Rows.Item(2).Cells(1).Value = DataGridView8.Rows.Item(2).Cells(1).Value + 1
            End If
            codigo = ""
            paroimpar = ""
        End If
    End Sub
    Private Function xmlGetValue(doc As SpreadsheetDocument, cell As DocumentFormat.OpenXml.Spreadsheet.Cell) As String
        Dim value As String
        Try
            value = cell.CellValue.InnerText
        Catch ex As Exception
            value = ""
        End Try
        If cell.DataType IsNot Nothing AndAlso cell.DataType.Value = CellValues.SharedString Then
            Return doc.WorkbookPart.SharedStringTablePart.SharedStringTable.ChildElements.Item(Integer.Parse(value)).InnerText
        End If
        Return value
    End Function
    Function traenomlargo(ByVal codigo As String, ByVal plantaid As String, Optional ByRef paroimpar As String = "PLANTA", Optional ByRef nom As String = "") As String
        Dim cons = "SELECT item_planta.descripcion, item_planta.cod_erp, Accesorios_Codigos.ParImpar, Accesorios_Codigos.Nomenclatura FROM item_planta WITH (nolock) INNER JOIN Accesorios_Codigos WITH (nolock) ON item_planta.cod_erp = Accesorios_Codigos.Id_UnoE AND item_planta.planta_id = Accesorios_Codigos.planta_id WHERE (item_planta.cod_erp = '" & codigo & "') AND (item_planta.planta_id = '" & plantaid & "') AND (item_planta.activo = 1)"
        Using connection As New SqlConnection(Str_Con_Bd)
            Dim command As SqlCommand = New SqlCommand(cons, connection)

            connection.Open()
            Dim reader As SqlDataReader = command.ExecuteReader()
            If reader.HasRows Then


                Do While reader.Read()
                    traenomlargo = reader.GetValue(0)
                    If IsDBNull(reader.GetValue(2)) Then
                        paroimpar = "PLANTA"
                    Else
                        paroimpar = reader.GetValue(2)
                    End If

                    nom = reader.GetValue(3)
                Loop
            Else
                traenomlargo = ""
            End If

            command = Nothing
            reader.Close()
            reader = Nothing
        End Using

    End Function
    Sub elijecodigo(ByRef codigo As String, ByRef paroimpar As String, datosfila As String(), plantaid As String, ByRef cantvalida As String, activo As Boolean, ByRef codigosId As String)
        Dim cant, descri, desaux, dim1, dim2, dim3, dim4, dim5, dim6, aux, plano As String
        Dim lista As String(,)
        ReDim Preserve lista(18, 0)
        cant = datosfila(0)
        descri = datosfila(1)
        desaux = datosfila(2)
        If datosfila(3) = "" Then dim1 = 0 Else dim1 = datosfila(3)
        If datosfila(4) = "" Then dim2 = 0 Else dim2 = datosfila(4)
        If datosfila(5) = "" Then dim3 = 0 Else dim3 = datosfila(5)
        If datosfila(6) = "" Then dim4 = 0 Else dim4 = datosfila(6)
        If datosfila(7) = "" Then dim5 = 0 Else dim5 = datosfila(7)
        If datosfila(8) = "" Then dim6 = 0 Else dim6 = datosfila(8)
        'dim1 = datosfila(3)
        'dim2 = datosfila(4)
        'dim3 = datosfila(5)
        'dim4 = datosfila(6)
        'dim5 = datosfila(7)
        'dim6 = datosfila(8)
        plano = datosfila(9)

        'agrega los items con tal descripcion en un arreglo "lista" para verificar cuales coinciden con la desaux
        Dim SlrsN1 = "SELECT Accesorios_Codigos.Codigos_Id, Accesorios_Codigos.Id_UnoE, Accesorios_Codigos.Nomenclatura, Accesorios_Codigos.Des_Aux, Accesorios_Codigos.Valor_1_Min, Accesorios_Codigos.Valor_1_Max, Accesorios_Codigos.Valor_2_Min, " +
                  "Accesorios_Codigos.Valor_2_Max, Accesorios_Codigos.Valor_3_Min, Accesorios_Codigos.Valor_3_Max, Accesorios_Codigos.Valor_4_Min, Accesorios_Codigos.Valor_4_Max, Accesorios_Codigos.Valor_5_Min, " +
                  "Accesorios_Codigos.Valor_5_Max, Accesorios_Codigos.Valor_6_Min, Accesorios_Codigos.Valor_6_Max, Accesorios_Codigos.Valor_7_Min, Accesorios_Codigos.Valor_7_Max, Accesorios_Codigos.ParImpar "
        Dim SlrsN2 = "FROM     Accesorios_Codigos " 'INNER JOIN item_planta ON Accesorios_Codigos.Acc_Id_ItemPlanta = item_planta.item_planta_id "
        If ComboBox7.SelectedItem = "CT" Then 'si es ct agrega tabla de precios
            SlrsN2 = SlrsN2 & "INNER JOIN Fup_ListaPreciosAccesoriosERP ON Accesorios_Codigos.Id_UnoE = Fup_ListaPreciosAccesoriosERP.ItemId AND Accesorios_Codigos.planta_id = Fup_ListaPreciosAccesoriosERP.Planta_Id "
        End If
        Dim SlrsN3 = "WHERE (Accesorios_Codigos.Nomenclatura = '" & descri & "') AND (Accesorios_Codigos.planta_id = " & plantaid & ")"
        If activo = True Then
            SlrsN3 = SlrsN3 & " AND (Accesorios_Codigos.Acc_Anulado = 0)" ' AND (item_planta.activo = 1)"
        End If
        If ComboBox7.SelectedItem = "CT" Then 'si es ct filtra por precio mayor a cero
            SlrsN3 = SlrsN3 & " AND (Fup_ListaPreciosAccesoriosERP.CostoPromUni > 0) "
            If CheckBox6.Checked = False Then 'True
                Dim filenametxt = "\\172.21.0.202\ingenieria\Z_VARIOS_SCI\FILTRO.txt"
                If My.Computer.FileSystem.FileExists(filenametxt) Then
                    Dim reader As StreamReader = My.Computer.FileSystem.OpenTextFileReader(filenametxt)
                    Dim linealeida As String
                    Do
                        linealeida = reader.ReadLine
                        If linealeida IsNot Nothing Then
                            SlrsN3 = SlrsN3 & " AND (Accesorios_Codigos.Id_UnoE <> '" & linealeida & "')"
                        End If
                    Loop Until linealeida Is Nothing
                    reader.Close()
                End If
            End If
            SlrsN3 = SlrsN3 & "OR (Accesorios_Codigos.Nomenclatura = '" & descri & "') AND (Accesorios_Codigos.planta_id = " & plantaid & ") AND (Accesorios_Codigos.Acc_Anulado = 0) " + _ 'AND (item_planta.activo = 1) " +
                "AND (Fup_ListaPreciosAccesoriosERP.CostoInventor > 0)"
            If CheckBox6.Checked = False Then 'True
                Dim filenametxt = "\\172.21.0.202\ingenieria\Z_VARIOS_SCI\FILTRO.txt"
                If My.Computer.FileSystem.FileExists(filenametxt) Then
                    Dim reader As StreamReader = My.Computer.FileSystem.OpenTextFileReader(filenametxt)
                    Dim linealeida As String
                    Do
                        linealeida = reader.ReadLine
                        If linealeida IsNot Nothing Then
                            SlrsN3 = SlrsN3 & " AND (Accesorios_Codigos.Id_UnoE <> '" & linealeida & "')"
                        End If
                    Loop Until linealeida Is Nothing
                    reader.Close()
                End If
            End If
        End If
        Dim SlrsN = SlrsN1 & SlrsN2 & SlrsN3
        Using connection As New SqlConnection(Str_Con_Bd)
            Dim command As SqlCommand = New SqlCommand(SlrsN, connection)
            connection.Open()
            Dim reader As SqlDataReader = command.ExecuteReader()
            If reader.HasRows Then
                'Dim i = 0
                Do While reader.Read()
                    Dim cantitems = UBound(lista, 2)

                    lista(0, cantitems) = reader.GetValue(0)
                    lista(1, cantitems) = reader.GetValue(1)
                    lista(2, cantitems) = reader.GetValue(2)
                    lista(3, cantitems) = reader.GetValue(3)
                    lista(4, cantitems) = reader.GetValue(4)
                    lista(5, cantitems) = reader.GetValue(5)
                    lista(6, cantitems) = reader.GetValue(6)
                    lista(7, cantitems) = reader.GetValue(7)
                    lista(8, cantitems) = reader.GetValue(8)
                    lista(9, cantitems) = reader.GetValue(9)
                    lista(10, cantitems) = reader.GetValue(10)
                    lista(11, cantitems) = reader.GetValue(11)
                    lista(12, cantitems) = reader.GetValue(12)
                    lista(13, cantitems) = reader.GetValue(13)
                    lista(14, cantitems) = reader.GetValue(14)
                    lista(15, cantitems) = reader.GetValue(15)
                    lista(16, cantitems) = reader.GetValue(16)
                    lista(17, cantitems) = reader.GetValue(17)
                    If plano = "p" Or plano = "P" Then
                        lista(18, cantitems) = "PLANTA"
                    Else
                        Try
                            lista(18, cantitems) = reader.GetValue(18) 'cuando parimpar es null
                        Catch ex As Exception
                            lista(18, cantitems) = "PLANTA"
                        End Try
                    End If


                    ReDim Preserve lista(18, cantitems + 1)
                    'i = i + 1
                Loop
                ReDim Preserve lista(18, UBound(lista, 2) - 1)
            End If
        End Using

        'reemplaza los guiones por guion bajo y crea un vector separando los diferentes textos de la desaux
        aux = Replace(desaux, "-", "_")
        Dim aux2 As String()
        Dim aux3 As String
        Try
            'Aqui lanza una excepcion, al usuario final no le afecta sin embargo no deberian haber inconvenientes desde aqui.
            'Hay que mirar como se ajusta para que la excepcion no afecte en modo desarrollador
            If aux IsNot Nothing Then
                aux2 = aux.Split("_")
            End If

            If plantaid = 1 Then
                desaux = "0"
            End If

            For j = 0 To UBound(aux2)
                Dim vec As String()
                ReDim vec(UBound(lista, 2))
                For k = 0 To UBound(lista, 2)
                    vec(k) = lista(3, k)
                Next
                'Dim aux3 As String
                aux3 = aux2(j)
                Dim salir = False 'se agrega salida para segundo for
                For Each veco As String In vec
                    If aux3 = veco Then
                        desaux = aux3
                        salir = True
                        Exit For 'aqui solo sale del primer for sino sale del segundo sigue buscando otros desaux
                    End If
                Next
                If salir = True Then
                    Exit For 'aqui sale del segundo for
                End If
                'If IsInArray(aux3, vec) Then
                '    desaux = aux3
                'End If
            Next
        Catch ex As Exception
            desaux = "0"
        End Try
        Dim sqldesaux As String
        If aux3 <> "" Then
            desaux = aux3
        Else
            desaux = "0"
        End If
        sqldesaux = ""
        If desaux = "0" Then
            desaux = ""
            sqldesaux = " = '" & desaux & "'"

            'ElseIf descri = "TB6X4" Or descri = "TB7X7" Or descri = "TB10X5" Or descri = "PGIROH" Or descri = "COMPDT" Or descri = "NGT2" Or descri = "PAS" Then
            '    sqldesaux = "LIKE '%" & desaux & "%'"
        Else
            sqldesaux = " = '" & desaux & "'"
        End If
        'reiniciar la lista
        ReDim lista(18, 0)
        'hace otra lista solo con los items que coinciden con la desaux
        SlrsN1 = "SELECT Accesorios_Codigos.Codigos_Id, Accesorios_Codigos.Id_UnoE, Accesorios_Codigos.Nomenclatura, Accesorios_Codigos.Des_Aux, Accesorios_Codigos.Valor_1_Min, Accesorios_Codigos.Valor_1_Max, Accesorios_Codigos.Valor_2_Min, " +
                  "Accesorios_Codigos.Valor_2_Max, Accesorios_Codigos.Valor_3_Min, Accesorios_Codigos.Valor_3_Max, Accesorios_Codigos.Valor_4_Min, Accesorios_Codigos.Valor_4_Max, Accesorios_Codigos.Valor_5_Min, " +
                  "Accesorios_Codigos.Valor_5_Max, Accesorios_Codigos.Valor_6_Min, Accesorios_Codigos.Valor_6_Max, Accesorios_Codigos.Valor_7_Min, Accesorios_Codigos.Valor_7_Max, Accesorios_Codigos.ParImpar "
        SlrsN2 = "FROM Accesorios_Codigos " 'INNER JOIN item_planta ON Accesorios_Codigos.Acc_Id_ItemPlanta = item_planta.item_planta_id "
        If ComboBox7.SelectedItem = "CT" Then 'si es ct agrega tabla de precios
            SlrsN2 = SlrsN2 & "INNER JOIN Fup_ListaPreciosAccesoriosERP ON Accesorios_Codigos.Id_UnoE = Fup_ListaPreciosAccesoriosERP.ItemId AND Accesorios_Codigos.planta_id = Fup_ListaPreciosAccesoriosERP.Planta_Id "
        End If
        SlrsN3 = "WHERE (Accesorios_Codigos.Nomenclatura = '" & descri & "') AND (Accesorios_Codigos.Des_Aux " & sqldesaux & ") AND (Accesorios_Codigos.planta_id = " & plantaid & ")"
        If activo = True Then
            SlrsN3 = SlrsN3 & " AND (Acc_Anulado = 0)" ' AND (item_planta.activo = 1)"
        End If
        If ComboBox7.SelectedItem = "CT" Then 'si es ct filtra por precio mayor a cero
            SlrsN3 = SlrsN3 & " AND (Fup_ListaPreciosAccesoriosERP.CostoPromUni > 0)"
            If CheckBox6.Checked = False Then 'True
                Dim filenametxt = "\\172.21.0.202\ingenieria\Z_VARIOS_SCI\FILTRO.txt"
                If My.Computer.FileSystem.FileExists(filenametxt) Then
                    Dim reader As StreamReader = My.Computer.FileSystem.OpenTextFileReader(filenametxt)
                    Dim linealeida As String
                    Do
                        linealeida = reader.ReadLine
                        If linealeida IsNot Nothing Then
                            SlrsN3 = SlrsN3 & " AND (Accesorios_Codigos.Id_UnoE <> '" & linealeida & "')"
                        End If
                    Loop Until linealeida Is Nothing
                    reader.Close()
                End If
            End If
        End If

        SlrsN = SlrsN1 & SlrsN2 & SlrsN3
        Using connection As New SqlConnection(Str_Con_Bd)
            Dim command As SqlCommand = New SqlCommand(SlrsN, connection)
            connection.Open()
            Dim reader As SqlDataReader = command.ExecuteReader()
            If reader.HasRows Then
                'Dim i = 0
                Do While reader.Read()
                    Dim cantitems = UBound(lista, 2)

                    lista(0, cantitems) = reader.GetValue(0)
                    lista(1, cantitems) = reader.GetValue(1)
                    lista(2, cantitems) = reader.GetValue(2)
                    lista(3, cantitems) = reader.GetValue(3)
                    lista(4, cantitems) = reader.GetValue(4)
                    lista(5, cantitems) = reader.GetValue(5)
                    lista(6, cantitems) = reader.GetValue(6)
                    lista(7, cantitems) = reader.GetValue(7)
                    lista(8, cantitems) = reader.GetValue(8)
                    lista(9, cantitems) = reader.GetValue(9)
                    lista(10, cantitems) = reader.GetValue(10)
                    lista(11, cantitems) = reader.GetValue(11)
                    lista(12, cantitems) = reader.GetValue(12)
                    lista(13, cantitems) = reader.GetValue(13)
                    lista(14, cantitems) = reader.GetValue(14)
                    lista(15, cantitems) = reader.GetValue(15)
                    lista(16, cantitems) = reader.GetValue(16)
                    lista(17, cantitems) = reader.GetValue(17)
                    If plano = "p" Or plano = "P" Then
                        lista(18, cantitems) = "PLANTA"
                    Else
                        Try
                            lista(18, cantitems) = reader.GetValue(18) 'cuando parimpar es null
                        Catch ex As Exception
                            lista(18, cantitems) = "PLANTA"
                        End Try
                    End If

                    ReDim Preserve lista(18, cantitems + 1)
                    'i = i + 1
                Loop
                ReDim Preserve lista(18, UBound(lista, 2) - 1)
            End If
        End Using
        Dim agregocodigo As String()
        ReDim agregocodigo(UBound(lista, 2))
        Dim indice As Integer = 0
        Dim validacodigo() As String
        ReDim validacodigo(5)
        For j = 0 To UBound(lista, 2)
            For k = 0 To 5
                validacodigo(k) = "False"
            Next
            If CDbl(lista(4, j)) = 0 And CDbl(lista(5, j)) = 0 Then
                validacodigo(0) = "True"
            ElseIf CDbl(dim1) >= CDbl(lista(4, j)) And CDbl(dim1) <= CDbl(lista(5, j)) Then
                validacodigo(0) = "True"
                indice = j
            End If
            If CDbl(lista(6, j)) = 0 And CDbl(lista(7, j)) = 0 Then
                validacodigo(1) = "True"
            ElseIf CDbl(dim2) >= CDbl(lista(6, j)) And CDbl(dim2) <= CDbl(lista(7, j)) Then
                validacodigo(1) = "True"
                indice = j
            End If
            If CDbl(lista(8, j)) = 0 And CDbl(lista(9, j)) = 0 Then
                validacodigo(2) = "True"
            ElseIf CDbl(dim3) >= CDbl(lista(8, j)) And CDbl(dim3) <= CDbl(lista(9, j)) Then
                validacodigo(2) = "True"
                indice = j
            End If
            If CDbl(lista(10, j)) = 0 And CDbl(lista(11, j)) = 0 Then
                validacodigo(3) = "True"
            ElseIf CDbl(dim4) >= CDbl(lista(10, j)) And CDbl(dim4) <= CDbl(lista(11, j)) Then
                validacodigo(3) = "True"
                indice = j
            End If
            If CDbl(lista(12, j)) = 0 And CDbl(lista(13, j)) = 0 Then
                validacodigo(4) = "True"
            ElseIf CDbl(dim5) >= CDbl(lista(12, j)) And CDbl(dim5) <= CDbl(lista(13, j)) Then
                validacodigo(4) = "True"
                indice = j
            End If
            If CDbl(lista(14, j)) = 0 And CDbl(lista(15, j)) = 0 Then
                validacodigo(5) = "True"
            ElseIf CDbl(dim6) >= CDbl(lista(14, j)) And CDbl(dim6) <= CDbl(lista(15, j)) Then
                validacodigo(5) = "True"
                indice = j
            End If
            If IsInArray("False", validacodigo) Then
                agregocodigo(j) = "False"
            Else
                agregocodigo(j) = "True"
                codigo = lista(1, indice)
                paroimpar = lista(18, indice)
                codigosId = lista(0, indice)
                Exit For
            End If
        Next
        If codigo = Nothing Or codigo = "19034" Or codigo = "193" Or codigo = "3315" Then
            If descri.Contains("NG") And ComboBox7.SelectedItem = "CT" Then
                Dim volng, pesong, dim3b As Double
                If dim3 = 0 Or dim3 = Nothing Then
                    dim3b = 10
                Else
                    dim3b = dim3
                End If
                volng = (CDbl(dim1) * CDbl(dim2) * CDbl(dim3b)) / 1000000
                pesong = (628.86 * volng) + 3.3009
                If pesong < 1 Then
                    codigo = "47160"
                ElseIf pesong > 1 And pesong <= 2 Then
                    codigo = "47161"
                ElseIf pesong > 2 And pesong <= 3 Then
                    codigo = "47162"
                ElseIf pesong > 3 And pesong <= 4 Then
                    codigo = "47165"
                ElseIf pesong > 4 And pesong <= 5 Then
                    codigo = "47169"
                ElseIf pesong > 5 And pesong <= 6 Then
                    codigo = "47171"
                ElseIf pesong > 6 And pesong <= 7 Then
                    codigo = "47174"
                ElseIf pesong > 7 And pesong <= 8 Then
                    codigo = "47176"
                ElseIf pesong > 8 And pesong <= 9 Then
                    codigo = "47177"
                ElseIf pesong > 9 And pesong <= 10 Then
                    codigo = "47179"
                ElseIf pesong > 10 And pesong <= 11 Then
                    codigo = "47181"
                ElseIf pesong > 11 And pesong <= 12 Then
                    codigo = "47183"
                ElseIf pesong > 12 And pesong <= 13 Then
                    codigo = "47186"
                ElseIf pesong > 13 And pesong <= 14 Then
                    codigo = "47189"
                ElseIf pesong > 14 And pesong <= 15 Then
                    codigo = "47191"
                ElseIf pesong > 15 And pesong <= 16 Then
                    codigo = "47193"
                ElseIf pesong > 16 And pesong <= 17 Then
                    codigo = "47195"
                ElseIf pesong > 17 And pesong <= 18 Then
                    codigo = "47197"
                ElseIf pesong > 18 And pesong <= 19 Then
                    codigo = "47199"
                ElseIf pesong > 19 And pesong <= 20 Then
                    codigo = "47203"
                ElseIf pesong > 20 And pesong <= 21 Then
                    codigo = "47205"
                ElseIf pesong > 21 And pesong <= 22 Then
                    codigo = "47208"
                ElseIf pesong > 22 And pesong <= 23 Then
                    codigo = "47210"
                ElseIf pesong > 23 And pesong <= 24 Then
                    codigo = "47212"
                ElseIf pesong > 24 And pesong <= 25 Then
                    codigo = "47213"
                ElseIf pesong > 25 And pesong <= 26 Then
                    codigo = "47215"
                ElseIf pesong > 26 And pesong <= 27 Then
                    codigo = "47218"
                ElseIf pesong > 27 And pesong <= 28 Then
                    codigo = "47220"
                ElseIf pesong > 28 And pesong <= 29 Then
                    codigo = "47163"
                ElseIf pesong > 29 And pesong <= 30 Then
                    codigo = "47164"
                ElseIf pesong > 30 And pesong <= 31 Then
                    codigo = "47166"
                ElseIf pesong > 31 And pesong <= 32 Then
                    codigo = "47167"
                ElseIf pesong > 32 And pesong <= 33 Then
                    codigo = "47168"
                ElseIf pesong > 33 And pesong <= 34 Then
                    codigo = "47170"
                ElseIf pesong > 34 And pesong <= 35 Then
                    codigo = "47173"
                ElseIf pesong > 35 And pesong <= 36 Then
                    codigo = "47175"
                ElseIf pesong > 36 And pesong <= 37 Then
                    codigo = "47178"
                ElseIf pesong > 37 And pesong <= 38 Then
                    codigo = "47180"
                ElseIf pesong > 38 And pesong <= 39 Then
                    codigo = "47182"
                ElseIf pesong > 39 And pesong <= 40 Then
                    codigo = "47188"
                ElseIf pesong > 40 And pesong <= 41 Then
                    codigo = "47190"
                ElseIf pesong > 41 And pesong <= 42 Then
                    codigo = "47192"
                ElseIf pesong > 42 And pesong <= 43 Then
                    codigo = "47194"
                ElseIf pesong > 43 And pesong <= 44 Then
                    codigo = "47196"
                ElseIf pesong > 44 And pesong <= 45 Then
                    codigo = "47198"
                ElseIf pesong > 45 And pesong <= 46 Then
                    codigo = "47200"
                ElseIf pesong > 46 And pesong <= 47 Then
                    codigo = "47204"
                ElseIf pesong > 47 And pesong <= 48 Then
                    codigo = "47206"
                ElseIf pesong > 48 And pesong <= 49 Then
                    codigo = "47207"
                ElseIf pesong > 49 And pesong <= 50 Then
                    codigo = "47209"
                ElseIf pesong > 50 And pesong <= 51 Then
                    codigo = "47211"
                ElseIf pesong > 51 And pesong <= 52 Then
                    codigo = "47214"
                ElseIf pesong > 52 And pesong <= 53 Then
                    codigo = "47216"
                ElseIf pesong > 53 And pesong <= 54 Then
                    codigo = "47217"
                ElseIf pesong > 54 And pesong <= 55 Then
                    codigo = "47219"
                ElseIf pesong > 55 And pesong <= 56 Then
                    codigo = "47221"
                ElseIf pesong > 56 And pesong <= 57 Then
                    codigo = "47279"
                ElseIf pesong > 57 And pesong <= 58 Then
                    codigo = "47282"
                ElseIf pesong > 58 And pesong <= 59 Then
                    codigo = "47283"
                ElseIf pesong > 59 And pesong <= 60 Then
                    codigo = "47284"
                ElseIf pesong > 60 And pesong <= 61 Then
                    codigo = "47285"
                ElseIf pesong > 61 And pesong <= 62 Then
                    codigo = "47289"
                ElseIf pesong > 62 And pesong <= 63 Then
                    codigo = "47290"
                ElseIf pesong > 63 And pesong <= 64 Then
                    codigo = "47291"
                ElseIf pesong > 64 Then
                    codigo = "47292"
                End If
                SlrsN = "SELECT Codigos_Id FROM Accesorios_Codigos WHERE (Id_UnoE = " & codigo & ") ORDER BY Id_UnoE DESC"
                Using connection As New SqlConnection(Str_Con_Bd)
                    Dim command As SqlCommand = New SqlCommand(SlrsN, connection)
                    connection.Open()
                    Dim reader As SqlDataReader = command.ExecuteReader()
                    If reader.HasRows Then
                        reader.Read()
                        codigosId = reader.GetValue(0)
                    Else
                        codigosId = "0"
                    End If
                End Using
            End If
        End If
        If codigo = Nothing And ComboBox7.SelectedItem = "CT" Then

            Dim sqlval As String

            'sqlval = " SELECT AC.Codigos_Id, AC.Id_UnoE, AC.Nomenclatura "
            '        " FROM Accesorios_Codigos WITH(NOLOCK) AS AC"

            'Hacer consulta anterior quitando filtro del costo
            'Si encuentra codigo es pq no tiene costo.
            'Si no encuentra codigo no existe
            SlrsN = "SELECT Accesorios_Codigos.Codigos_Id, Accesorios_Codigos.Id_UnoE, Accesorios_Codigos.Nomenclatura, Fup_ListaPreciosAccesoriosERP.CostoPromUni, Accesorios_Codigos.ParImpar " +
                    "FROM Accesorios_Codigos INNER JOIN Fup_ListaPreciosAccesoriosERP ON Accesorios_Codigos.Id_UnoE = Fup_ListaPreciosAccesoriosERP.ItemId AND Accesorios_Codigos.planta_id = Fup_ListaPreciosAccesoriosERP.Planta_Id " +
                    "WHERE (Accesorios_Codigos.planta_id = 1) AND (Accesorios_Codigos.Nomenclatura = '" & descri & "') AND (Fup_ListaPreciosAccesoriosERP.CostoPromUni > 0) AND (Accesorios_Codigos.Acc_Anulado = 0)"

            If CheckBox6.Checked = False Then 'True
                Dim filenametxt = "\\172.21.0.202\ingenieria\Z_VARIOS_SCI\FILTRO.txt"
                If My.Computer.FileSystem.FileExists(filenametxt) Then
                    Dim reader As StreamReader = My.Computer.FileSystem.OpenTextFileReader(filenametxt)
                    Dim linealeida As String
                    Do
                        linealeida = reader.ReadLine
                        If linealeida IsNot Nothing Then
                            SlrsN = SlrsN & " AND (Accesorios_Codigos.Id_UnoE <> '" & linealeida & "')"
                        End If
                    Loop Until linealeida Is Nothing
                    reader.Close()
                End If
            End If
            SlrsN = SlrsN & " ORDER BY Fup_ListaPreciosAccesoriosERP.CostoPromUni DESC"
            Using connection As New SqlConnection(Str_Con_Bd)
                Dim command As SqlCommand = New SqlCommand(SlrsN, connection)
                connection.Open()
                Dim reader As SqlDataReader = command.ExecuteReader()
                If reader.HasRows Then
                    reader.Read()
                    codigo = reader.GetValue(1)
                    codigosId = reader.GetValue(0)
                    Try
                        paroimpar = reader.GetValue(4)
                    Catch ex As Exception
                        paroimpar = "PLANTA"
                    End Try
                End If
            End Using
        End If
        If codigo = Nothing Then
            codigo = 0
            paroimpar = "Planta"
        ElseIf IsInArray("True", agregocodigo) Then

        ElseIf codigo <> Nothing Then

        Else
            codigo = 0
            paroimpar = "Planta"
        End If
        If descri = "ALCP" Or descri = "#ALCP" Or descri = "PBCP" Or descri = "#PBCP" Or descri = "SPALCP" Or descri = "PAS" Or descri = "#MPAS" Or descri = "PORTV" Or descri = "#PORTV" Or descri = "SCCR" Or descri = "SCSR" Or descri = "#SCCR" Or descri = "#SCSR" Or descri = "SPAGR" Or descri = "SUML" Or descri = "#SUML" Or descri = "STB" Or descri = "STBMURO" Or descri = "STBMUROAG" Or descri = "SCM" Or descri = "GRADAND" Or descri = "PELGRAND" Then
            cantvalida = 2
        ElseIf descri = "NGC1" Or descri = "NGT1" Or descri = "NGT2" Or descri = "NGLT1" Or descri = "NGLT2" Or descri = "NGLT3" Or descri = "NGLT4" Or descri = "NGLT5" Or descri = "NGLT6" Or descri = "T1" Or descri = "T2" Or descri = "T3" Or descri = "T4" Or descri = "T5" Then
            cantvalida = cant
        Else
            cantvalida = 0
        End If
    End Sub


    Sub elijeKamban(ByRef codigo As String, ByVal cant As Integer, ByVal descri As String, ByVal desaux As String, ByVal dim1 As Double, ByVal dim2 As Double, ByVal dim3 As Double, ByVal dim4 As Double, plantaid As String)
        'Dim Pass_Bd As String = "forsa2006"
        'Dim Login_Bd As String = "Forsa"
        'Dim Instan_Bd As String = "Forsa"
        'Dim Nom_Bd As String = "172.21.224.130"
        'Dim Str_Con_Bd As String = "Password=" & Pass_Bd & ";Persist Security " +
        '"Info=True;User ID=" & Login_Bd & ";Initial Catalog=" & Instan_Bd & ";Data Source=" & Nom_Bd

        Dim aux As String
        Dim lista As String(,)
        ReDim Preserve lista(18, 0)

        'agrega los items con tal descripcion en un arreglo "lista" para verificar cuales coinciden con la desaux
        Dim SlrsN1 = "SELECT Codigos_Id, Id_UnoE, Nomenclatura, Des_Aux, Valor_1_Min, Valor_1_Max, Valor_2_Min, Valor_2_Max, Valor_3_Min, Valor_3_Max, Valor_4_Min, Valor_4_Max, " +
            "Valor_5_Min , Valor_5_Max, Valor_6_Min, Valor_6_Max, Valor_7_Min, Valor_7_Max, ParImpar "
        Dim SlrsN2 = "FROM Accesorios_Codigos "
        Dim SlrsN3 = "WHERE (Nomenclatura = '" & descri & "') AND (planta_id = " & plantaid & ") AND (Acc_Anulado = 0)"
        Dim SlrsN = SlrsN1 & SlrsN2 & SlrsN3
        Using connection As New SqlConnection(Str_Con_Bd)
            Dim command As SqlCommand = New SqlCommand(SlrsN, connection)
            connection.Open()
            Dim reader As SqlDataReader = command.ExecuteReader()
            If reader.HasRows Then
                'Dim i = 0
                Do While reader.Read()
                    Dim cantitems = UBound(lista, 2)

                    lista(0, cantitems) = reader.GetValue(0)
                    lista(1, cantitems) = reader.GetValue(1)
                    lista(2, cantitems) = reader.GetValue(2)
                    lista(3, cantitems) = reader.GetValue(3)
                    lista(4, cantitems) = reader.GetValue(4)
                    lista(5, cantitems) = reader.GetValue(5)
                    lista(6, cantitems) = reader.GetValue(6)
                    lista(7, cantitems) = reader.GetValue(7)
                    lista(8, cantitems) = reader.GetValue(8)
                    lista(9, cantitems) = reader.GetValue(9)
                    lista(10, cantitems) = reader.GetValue(10)
                    lista(11, cantitems) = reader.GetValue(11)
                    lista(12, cantitems) = reader.GetValue(12)
                    lista(13, cantitems) = reader.GetValue(13)
                    lista(14, cantitems) = reader.GetValue(14)
                    lista(15, cantitems) = reader.GetValue(15)
                    lista(16, cantitems) = reader.GetValue(16)
                    lista(17, cantitems) = reader.GetValue(17)

                    lista(18, cantitems) = "PLANTA"


                    ReDim Preserve lista(18, cantitems + 1)
                    'i = i + 1
                Loop
                ReDim Preserve lista(18, UBound(lista, 2) - 1)
            End If
        End Using

        'reemplaza los guiones por guion bajo y crea un vector separando los diferentes textos de la desaux
        aux = Replace(desaux, "-", "_")
        Dim aux2 As String()
        Try
            If aux IsNot Nothing Then
                aux2 = aux.Split("_")
            End If

            If plantaid = 1 Then
                desaux = "0"
            End If
            For j = 0 To UBound(aux2)
                Dim vec As String()
                ReDim vec(UBound(lista, 2))
                For k = 0 To UBound(lista, 2)
                    vec(k) = lista(3, k)
                Next
                Dim aux3 As String
                aux3 = aux2(j)
                For Each veco As String In vec
                    If aux3 = veco Then
                        desaux = aux3
                        Exit For
                    End If
                Next
                'If IsInArray(aux3, vec) Then
                '    desaux = aux3
                'End If
            Next
        Catch ex As System.Exception
            desaux = "0"
        End Try
        Dim sqldesaux As String
        sqldesaux = ""
        If desaux = "0" Then
            desaux = ""
            sqldesaux = " = '" & desaux & "'"

            'ElseIf descri = "TB6X4" Or descri = "TB7X7" Or descri = "TB10X5" Or descri = "PGIROH" Or descri = "COMPDT" Or descri = "NGT2" Or descri = "PAS" Then
            '    sqldesaux = "LIKE '%" & desaux & "%'"
        Else
            sqldesaux = " = '" & desaux & "'"
        End If
        'reiniciar la lista
        ReDim lista(18, 0)
        'hace otra lista solo con los items que coinciden con la desaux
        SlrsN = "SELECT Codigos_Id, Id_UnoE, Nomenclatura, Des_Aux, Valor_1_Min, Valor_1_Max, Valor_2_Min, Valor_2_Max, Valor_3_Min, Valor_3_Max, Valor_4_Min, Valor_4_Max, " +
                "Valor_5_Min , Valor_5_Max, Valor_6_Min, Valor_6_Max, Valor_7_Min, Valor_7_Max, ParImpar FROM Accesorios_Codigos WHERE (Nomenclatura = '" & descri & "') AND (Des_Aux " & sqldesaux & ") AND (planta_id = " & plantaid & ") AND (Acc_Anulado = 0)" '(Des_Aux LIKE '%" & desaux & "%')
        Using connection As New SqlConnection(Str_Con_Bd)
            Dim command As SqlCommand = New SqlCommand(SlrsN, connection)
            connection.Open()
            Dim reader As SqlDataReader = command.ExecuteReader()
            If reader.HasRows Then
                'Dim i = 0
                Do While reader.Read()
                    Dim cantitems = UBound(lista, 2)

                    lista(0, cantitems) = reader.GetValue(0)
                    lista(1, cantitems) = reader.GetValue(1)
                    lista(2, cantitems) = reader.GetValue(2)
                    lista(3, cantitems) = reader.GetValue(3)
                    lista(4, cantitems) = reader.GetValue(4)
                    lista(5, cantitems) = reader.GetValue(5)
                    lista(6, cantitems) = reader.GetValue(6)
                    lista(7, cantitems) = reader.GetValue(7)
                    lista(8, cantitems) = reader.GetValue(8)
                    lista(9, cantitems) = reader.GetValue(9)
                    lista(10, cantitems) = reader.GetValue(10)
                    lista(11, cantitems) = reader.GetValue(11)
                    lista(12, cantitems) = reader.GetValue(12)
                    lista(13, cantitems) = reader.GetValue(13)
                    lista(14, cantitems) = reader.GetValue(14)
                    lista(15, cantitems) = reader.GetValue(15)
                    lista(16, cantitems) = reader.GetValue(16)
                    lista(17, cantitems) = reader.GetValue(17)

                    lista(18, cantitems) = "PLANTA"


                    ReDim Preserve lista(18, cantitems + 1)
                    'i = i + 1
                Loop
                ReDim Preserve lista(18, UBound(lista, 2) - 1)
            End If
        End Using
        Dim agregocodigo As String()
        ReDim agregocodigo(UBound(lista, 2))
        Dim indice As Integer = 0
        Dim validacodigo() As String
        ReDim validacodigo(5)
        For j = 0 To UBound(lista, 2)
            For k = 0 To 5
                validacodigo(k) = "False"
            Next
            If CDbl(lista(4, j)) = 0 And CDbl(lista(5, j)) = 0 Then
                validacodigo(0) = "True"
            ElseIf CDbl(dim1) >= CDbl(lista(4, j)) And CDbl(dim1) <= CDbl(lista(5, j)) Then
                validacodigo(0) = "True"
                indice = j
            End If
            If CDbl(lista(6, j)) = 0 And CDbl(lista(7, j)) = 0 Then
                validacodigo(1) = "True"
            ElseIf CDbl(dim2) >= CDbl(lista(6, j)) And CDbl(dim2) <= CDbl(lista(7, j)) Then
                validacodigo(1) = "True"
                indice = j
            End If
            If CDbl(lista(8, j)) = 0 And CDbl(lista(9, j)) = 0 Then
                validacodigo(2) = "True"
            ElseIf CDbl(dim3) >= CDbl(lista(8, j)) And CDbl(dim3) <= CDbl(lista(9, j)) Then
                validacodigo(2) = "True"
                indice = j
            End If
            If CDbl(lista(10, j)) = 0 And CDbl(lista(11, j)) = 0 Then
                validacodigo(3) = "True"
            ElseIf CDbl(dim4) >= CDbl(lista(10, j)) And CDbl(dim4) <= CDbl(lista(11, j)) Then
                validacodigo(3) = "True"
                indice = j
            End If
            validacodigo(4) = "True"
            validacodigo(5) = "True"
            If IsInArray("False", validacodigo) Then
                agregocodigo(j) = "False"
            Else
                agregocodigo(j) = "True"
                codigo = lista(1, indice)

                'Exit For
            End If
        Next


        If codigo = Nothing Then
            codigo = 0
            'paroimpar = "Planta"
        ElseIf IsInArray("True", agregocodigo) Then

        Else
            codigo = 0
            'paroimpar = "Planta"
        End If
    End Sub
    Function IsInArray(stringToBeFound As String, arr As Object) As Boolean
        IsInArray = (UBound(Strings.Filter(arr, stringToBeFound,)) > -1)
    End Function
    Sub dataGridView1_RowPostPaint(ByVal sender As Object, ByVal e As DataGridViewRowPostPaintEventArgs) Handles DataGridView1.RowPostPaint
        Dim grilla As DataGridView = CType(sender, DataGridView)
        Dim fila As String = (e.RowIndex + 1).ToString()
        Dim rowFont As New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        'Dim centerFormat As New StringFormat()
        'centerFormat.Alignment = StringAlignment.Far
        'centerFormat.LineAlignment = StringAlignment.Near

        Dim headerBounds As System.Drawing.Rectangle = New System.Drawing.Rectangle(e.RowBounds.Left, e.RowBounds.Top, grilla.RowHeadersWidth, e.RowBounds.Height)
        e.Graphics.DrawString(fila, rowFont, System.Drawing.SystemBrushes.ControlText, headerBounds)
        llenacuadros(DataGridView1)

    End Sub
    Sub dataGridView2_RowPostPaint(ByVal sender As Object, ByVal e As DataGridViewRowPostPaintEventArgs) Handles DataGridView2.RowPostPaint
        Dim grilla As DataGridView = CType(sender, DataGridView)
        Dim fila As String = (e.RowIndex + 1).ToString()
        Dim rowFont As New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        'Dim centerFormat As New StringFormat()
        'centerFormat.Alignment = StringAlignment.Far
        'centerFormat.LineAlignment = StringAlignment.Near

        Dim headerBounds As System.Drawing.Rectangle = New System.Drawing.Rectangle(e.RowBounds.Left, e.RowBounds.Top, grilla.RowHeadersWidth, e.RowBounds.Height)
        e.Graphics.DrawString(fila, rowFont, System.Drawing.SystemBrushes.ControlText, headerBounds)
        llenacuadros(DataGridView2)
    End Sub
    Sub dataGridView3_RowPostPaint(ByVal sender As Object, ByVal e As DataGridViewRowPostPaintEventArgs) Handles DataGridView3.RowPostPaint
        Dim grilla As DataGridView = CType(sender, DataGridView)
        Dim fila As String = (e.RowIndex + 1).ToString()
        Dim rowFont As New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        'Dim centerFormat As New StringFormat()
        'centerFormat.Alignment = StringAlignment.Far
        'centerFormat.LineAlignment = StringAlignment.Near

        Dim headerBounds As System.Drawing.Rectangle = New System.Drawing.Rectangle(e.RowBounds.Left, e.RowBounds.Top, grilla.RowHeadersWidth, e.RowBounds.Height)
        e.Graphics.DrawString(fila, rowFont, System.Drawing.SystemBrushes.ControlText, headerBounds)
        llenacuadros(DataGridView3)
    End Sub
    Sub dataGridView5_RowPostPaint(ByVal sender As Object, ByVal e As DataGridViewRowPostPaintEventArgs) Handles DataGridView5.RowPostPaint
        Dim grilla As DataGridView = CType(sender, DataGridView)
        Dim fila As String = (e.RowIndex + 1).ToString()
        Dim rowFont As New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        'Dim centerFormat As New StringFormat()
        'centerFormat.Alignment = StringAlignment.Far
        'centerFormat.LineAlignment = StringAlignment.Near

        Dim headerBounds As System.Drawing.Rectangle = New System.Drawing.Rectangle(e.RowBounds.Left, e.RowBounds.Top, grilla.RowHeadersWidth, e.RowBounds.Height)
        e.Graphics.DrawString(fila, rowFont, System.Drawing.SystemBrushes.ControlText, headerBounds)
        llenacuadros(DataGridView5)
    End Sub
    Sub dataGridView4_RowPostPaint(ByVal sender As Object, ByVal e As DataGridViewRowPostPaintEventArgs) Handles DataGridView4.RowPostPaint
        Dim grilla As DataGridView = CType(sender, DataGridView)
        Dim fila As String = (e.RowIndex + 1).ToString()
        Dim rowFont As New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        'Dim centerFormat As New StringFormat()
        'centerFormat.Alignment = StringAlignment.Far
        'centerFormat.LineAlignment = StringAlignment.Near

        Dim headerBounds As System.Drawing.Rectangle = New System.Drawing.Rectangle(e.RowBounds.Left, e.RowBounds.Top, grilla.RowHeadersWidth, e.RowBounds.Height)
        e.Graphics.DrawString(fila, rowFont, System.Drawing.SystemBrushes.ControlText, headerBounds)
    End Sub
    Sub dataGridView7_RowPostPaint(ByVal sender As Object, ByVal e As DataGridViewRowPostPaintEventArgs) Handles DataGridView7.RowPostPaint
        Dim grilla As DataGridView = CType(sender, DataGridView)
        Dim fila As String = (e.RowIndex + 1).ToString()
        Dim rowFont As New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        'Dim centerFormat As New StringFormat()
        'centerFormat.Alignment = StringAlignment.Far
        'centerFormat.LineAlignment = StringAlignment.Near

        Dim headerBounds As System.Drawing.Rectangle = New System.Drawing.Rectangle(e.RowBounds.Left, e.RowBounds.Top, grilla.RowHeadersWidth, e.RowBounds.Height)
        e.Graphics.DrawString(fila, rowFont, System.Drawing.SystemBrushes.ControlText, headerBounds)
    End Sub
    Sub dataGridView6_RowPostPaint(ByVal sender As Object, ByVal e As DataGridViewRowPostPaintEventArgs) Handles DataGridView6.RowPostPaint
        If CheckBox2.Checked = False Then

            DataGridView6.Rows.Item(4).Cells(1).Value = CInt(DataGridView6.Rows.Item(0).Cells(1).Value) + CInt(DataGridView6.Rows.Item(1).Cells(1).Value) + CInt(DataGridView6.Rows.Item(2).Cells(1).Value) + CInt(DataGridView6.Rows.Item(3).Cells(1).Value)
            DataGridView6.Rows.Item(4).Cells(2).Value = CInt(DataGridView6.Rows.Item(0).Cells(2).Value) + CInt(DataGridView6.Rows.Item(1).Cells(2).Value) + CInt(DataGridView6.Rows.Item(2).Cells(2).Value) + CInt(DataGridView6.Rows.Item(3).Cells(2).Value)
            DataGridView6.Rows.Item(4).Cells(3).Value = CInt(DataGridView6.Rows.Item(0).Cells(3).Value) + CInt(DataGridView6.Rows.Item(1).Cells(3).Value) + CInt(DataGridView6.Rows.Item(2).Cells(3).Value) + CInt(DataGridView6.Rows.Item(3).Cells(3).Value)
            DataGridView6.Rows.Item(4).Cells(4).Value = CDbl(DataGridView6.Rows.Item(0).Cells(4).Value) + CDbl(DataGridView6.Rows.Item(1).Cells(4).Value) + CDbl(DataGridView6.Rows.Item(2).Cells(4).Value) + CDbl(DataGridView6.Rows.Item(3).Cells(4).Value)
            DataGridView6.Rows.Item(4).Cells(5).Value = CDbl(DataGridView6.Rows.Item(0).Cells(5).Value) + CDbl(DataGridView6.Rows.Item(1).Cells(5).Value) + CDbl(DataGridView6.Rows.Item(2).Cells(5).Value) + CDbl(DataGridView6.Rows.Item(3).Cells(5).Value)
            DataGridView6.Rows.Item(4).Cells(6).Value = CDbl(DataGridView6.Rows.Item(0).Cells(6).Value) + CDbl(DataGridView6.Rows.Item(1).Cells(6).Value) + CDbl(DataGridView6.Rows.Item(2).Cells(6).Value) + CDbl(DataGridView6.Rows.Item(3).Cells(6).Value)
        End If
    End Sub
    Sub datagridview1_CellEndEdit(ByVal sender As Object, ByVal e As DataGridViewCellEventArgs) Handles DataGridView1.CellEndEdit
        Dim fila = DataGridView1.Rows.Item(e.RowIndex)
        recalcular(fila)
    End Sub
    Sub datagridview2_CellEndEdit(ByVal sender As Object, ByVal e As DataGridViewCellEventArgs) Handles DataGridView2.CellEndEdit
        Dim fila = DataGridView2.Rows.Item(e.RowIndex)
        recalcular(fila)
    End Sub
    Sub datagridview3_CellEndEdit(ByVal sender As Object, ByVal e As DataGridViewCellEventArgs) Handles DataGridView3.CellEndEdit
        Dim fila = DataGridView3.Rows.Item(e.RowIndex)
        recalcular(fila)
    End Sub
    Sub datagridview5_CellEndEdit(ByVal sender As Object, ByVal e As DataGridViewCellEventArgs) Handles DataGridView5.CellEndEdit
        Dim fila = DataGridView5.Rows.Item(e.RowIndex)
        recalcular(fila)
    End Sub
    Sub ordenar()




        'ordenalista(DataGridView1, 0, DataGridView1.RowCount - 1, 1)
        DataGridView1.Sort(DataGridView1.Columns.Item(17), Ascending) 'ComponentModel.ListSortDirection.Ascending

        'ordenalista(DataGridView2, 0, DataGridView2.RowCount - 1, 1)
        DataGridView2.Sort(DataGridView2.Columns.Item(17), Ascending) 'ComponentModel.ListSortDirection.Ascending

        'ordenalista(DataGridView3, 0, DataGridView3.RowCount - 1, 1)
        DataGridView3.Sort(DataGridView3.Columns.Item(17), Ascending) 'ComponentModel.ListSortDirection.Ascending

        DataGridView5.Sort(DataGridView5.Columns.Item(17), Ascending) 'ComponentModel.ListSortDirection.Ascending

        'ordenalista(DataGridView4, 0, DataGridView4.RowCount - 1, 1)
        DataGridView4.Sort(DataGridView4.Columns.Item(11), Ascending) 'ComponentModel.ListSortDirection.Ascending



        ordenadesc(DataGridView1, False)

        ordenadesc(DataGridView2, True)

        ordenadesc(DataGridView3, False)

        ordenadesc(DataGridView5, False)

        ordenadesc(DataGridView4, False)


    End Sub

    Sub ordenadesc(grilla As System.Windows.Forms.DataGridView, uml As Boolean)
        Dim indexmenor As Integer
        Dim indexmayor As Integer
        Dim indexmenor2 As Integer
        Dim indexmayor2 As Integer
        Dim ordenaprimero As Integer
        Dim ordenasegundo As Integer

        If uml = True Then
            ordenaprimero = 2
            ordenasegundo = 3
        Else
            ordenaprimero = 3
            ordenasegundo = 2
        End If

        'ordena descendente por grupos de alto1
        indexmenor = 0
        For i = 0 To grilla.RowCount - 1
            Dim ultimoindex
            Try
                ultimoindex = grilla(1, i + 1).Value
            Catch ex As Exception
                ultimoindex = grilla(1, i).Value
            End Try

            If grilla(1, i).Value = ultimoindex And i < grilla.RowCount - 1 Then
                indexmayor = i
            Else
                If indexmenor = i Then
                    indexmenor = i + 1
                Else
                    indexmayor = i
                    ordenalistadesc(grilla, indexmenor, indexmayor, ordenaprimero)
                    indexmenor2 = indexmenor
                    indexmayor2 = indexmenor
                    For j = indexmenor To indexmayor
                        Dim ultimoindex2
                        Try
                            ultimoindex2 = grilla(ordenaprimero, j + 1).Value
                        Catch ex As Exception
                            ultimoindex2 = grilla(ordenaprimero, j).Value
                        End Try
                        If grilla(ordenaprimero, j).Value = ultimoindex2 And j < indexmayor Then
                            indexmayor2 = j + 1
                        Else
                            If indexmenor2 = j Then
                                indexmenor2 = j + 1
                            Else
                                indexmayor2 = j
                                ordenalistadesc(grilla, indexmenor2, indexmayor2, ordenasegundo)
                                indexmenor2 = j + 1
                                'indexmayor = i + 1
                            End If
                        End If
                    Next
                    indexmenor = i + 1
                    'indexmayor = i + 1
                End If
            End If
        Next
        'ordena descendente por grupos de ancho

    End Sub

    Sub ordenalista(grilla As System.Windows.Forms.DataGridView, ByVal indexmenor As Integer, ByVal indexmayor As Integer, columna As Integer)

        Dim indexparticion As String
        Dim guardadortemporal As Object
        Dim i, j As Integer
        If indexmenor >= indexmayor Then Exit Sub

        Dim numero As Integer = System.Math.Truncate((indexmenor + indexmayor) / 2)
        indexparticion = grilla(columna, numero).Value.ToString
        i = indexmenor
        j = indexmayor
        Do
            Do While grilla(columna, i).Value.ToString < indexparticion
                i = i + 1
            Loop
            Do While grilla(columna, j).Value.ToString > indexparticion
                j = j - 1
            Loop
            If i <= j Then
                For k = 0 To grilla.ColumnCount - 1
                    guardadortemporal = grilla(k, j).Value
                    grilla(k, j).Value = grilla(k, i).Value
                    grilla(k, i).Value = guardadortemporal
                Next
                i = i + 1
                j = j - 1
            End If
        Loop Until i > j
        If (j - indexmenor) < (indexmayor - i) Then
            ordenalista(grilla, indexmenor, j, columna)
            ordenalista(grilla, i, indexmayor, columna)
        Else
            ordenalista(grilla, i, indexmayor, columna)
            ordenalista(grilla, indexmenor, j, columna)
        End If

    End Sub
    Sub ordenalistadesc(grilla As System.Windows.Forms.DataGridView, ByVal indexmenor As Integer, ByVal indexmayor As Integer, columna As Integer)

        Dim indexparticion As Double
        Dim guardadortemporal As Object
        Dim i, j As Integer
        If indexmenor >= indexmayor Then Exit Sub
        Dim numero As Integer = System.Math.Truncate((indexmenor + indexmayor) / 2)
        indexparticion = grilla(columna, numero).Value
        i = indexmenor
        j = indexmayor
        Do
            Do While grilla(columna, i).Value > indexparticion
                i = i + 1
            Loop
            Do While grilla(columna, j).Value < indexparticion
                j = j - 1
            Loop
            If i <= j Then
                For k = 0 To grilla.ColumnCount - 1
                    guardadortemporal = grilla(k, j).Value
                    grilla(k, j).Value = grilla(k, i).Value
                    grilla(k, i).Value = guardadortemporal
                Next
                i = i + 1
                j = j - 1
            End If
        Loop Until i > j
        If (j - indexmenor) > (indexmayor - i) Then
            ordenalistadesc(grilla, indexmenor, j, columna)
            ordenalistadesc(grilla, i, indexmayor, columna)
        Else
            ordenalistadesc(grilla, i, indexmayor, columna)
            ordenalistadesc(grilla, indexmenor, j, columna)
        End If

    End Sub
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load

        ReDim mensajesfinales(0)
        mensajesfinales(UBound(mensajesfinales)) = "Debe Corregir estos puntos en el listado para poder cargar"
        ReDim Preserve mensajesfinales(UBound(mensajesfinales) + 1)
        mensajesfinales(UBound(mensajesfinales)) = "***"
        ReDim Preserve mensajesfinales(UBound(mensajesfinales) + 1)
        mensajesfinales(UBound(mensajesfinales)) = "**"
        ReDim Preserve mensajesfinales(UBound(mensajesfinales) + 1)
        mensajesfinales(UBound(mensajesfinales)) = "*"
        ReDim Preserve mensajesfinales(UBound(mensajesfinales) + 1)
        mensajesfinales(UBound(mensajesfinales)) = ""
        ReDim Preserve mensajesfinales(UBound(mensajesfinales) + 1)

        'OJO: Este Entorno se captura directamente de la aplicacion en ejecucion
        'Es decir que como la que se ejecuta es CargaListasMemoriaCalc el app.config sera tomado de alli
        'Mientras que en el proceso normal para el usuario lo tomaria directamente de este mismo proyecto.
        'Recordar siempre antes de ejecutar los proyectos compilarlos Ctrl + B

        Dim Entorno As String
        Dim version As String = versionactual()

        If ConfigurationManager.AppSettings("TC") IsNot Nothing Then

            Entorno = ConfigurationManager.AppSettings("TC")

        Else

            Entorno = [CONST].TC

        End If

        'Remover TabsPage
        'Tab_CargarInfo.TabPages.Remove(TabPrueba)
        Tab_CargarInfo.TabPages.Remove(TabPage_ERP)

        Select Case Entorno

            Case "R"
                Tab_CargarInfo.TabPages.Remove(TabPrueba)
                Me.Text = Me.Text + " - V. " + version + " - Produccion"
            Case "P"
                Me.Text = Me.Text + " - V. " + version + " - Pruebas"
            Case Else
                Me.Text = Me.Text + " - V. " + version + " - Pruebas"

        End Select

        If Label11.Text = "Brasil" Then
            Button18.Visible = True
            ValDesAux.Visible = True
        Else
            Button18.Visible = False
            ValDesAux.Visible = False
        End If

        'Asignar PlantaId
        If Label11.Text = "Colombia" Then
            Me.PlantaId = 1
        ElseIf Label11.Text = "Mexico" Then
            Me.PlantaId = 2
        ElseIf Label11.Text = "Brasil" Then
            Me.PlantaId = 3
        End If

        Button1.Enabled = False
        Button2.Enabled = False
        CheckBox3.Enabled = False

        Dim rsiferp2 = New ADODB.Recordset
        Dim Slrs10 = "Select " +
                "T010_MM_COMPANIAS.F010_ID, " +
                "T010_MM_COMPANIAS.F010_RAZON_SOCIAL " +
            "From " +
                "T010_MM_COMPANIAS " +
            "Where " +
                "(T010_MM_COMPANIAS.F010_ID = 15) Or " +
                "(T010_MM_COMPANIAS.F010_ID = 11) Or " +
                "(T010_MM_COMPANIAS.F010_ID = 12) Or " +
                "(T010_MM_COMPANIAS.F010_ID = 14) Or " +
                "(T010_MM_COMPANIAS.F010_ID = 6) " +
            "Order By " +
                "T010_MM_COMPANIAS.F010_ID"
        llenacombo2()
        Dim cons = "SELECT T_Sol_Tipo FROM Tipos_Sol WITH (nolock) WHERE (T_Sol_Activo = 1) AND (NOT (T_Sol_Tipo = 'OK'))"
        Using connection As New SqlConnection(Str_Con_Bd)
            Dim command As SqlCommand = New SqlCommand(cons, connection)

            connection.Open()
            Dim reader As SqlDataReader = command.ExecuteReader()
            If reader.HasRows Then


                Do While reader.Read()
                    Label12.Text = reader.GetValue(0)
                Loop
            End If

            command = Nothing
            reader.Close()
            reader = Nothing
        End Using
        'Dim Con_Ora As ADODB.Connection = CreaRQI.MyCommands.AbreDb_Orl
        'rsiferp2.Open(Slrs10, Con_Ora, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        Dim tablacias As Object
        'tablacias = rsiferp2.GetRows
        'For i = 0 To UBound(tablacias, 2)
        'ComboBox1.Items.Add(tablacias(0, i) & " - " & tablacias(1, i))
        'Next

        'Tab del lado derecho
        Tab_VerInfoCargada.TabPages.Item(0).Text = "MUROS"
        Tab_VerInfoCargada.TabPages.Item(1).Text = "UNION"
        Tab_VerInfoCargada.TabPages.Item(2).Text = "LOSAS"
        Tab_VerInfoCargada.TabPages.Item(3).Text = "CULAT"
        Tab_VerInfoCargada.TabPages.Item(4).Text = "ACC"
        'Tab_CargarInfo.TabPages.Item(0).Text = "ERP"
        'Tab_CargarInfo.TabPages.Item(1).Text = "SIIF"
        'Tab_CargarInfo.TabPages.Item(2).Text = "ACC"
        Tab_CargarInfo.TabPages.Item(0).Text = "SIIF"
        Tab_CargarInfo.TabPages.Item(1).Text = "ACC"


        For i = 0 To DataGridView1.Columns.Count - 1
            ancho = ancho + DataGridView1.Columns.Item(i).Width
        Next
        ruta = Nothing
        Try
            CreaRQI.MyCommands.RutaDwg()
            ruta = CreaRQI.MyCommands.ruta
        Catch ex As Exception
            'MsgBox("Debe guardar el plano modulado", MsgBoxStyle.Exclamation) 'se deshabilita el llena grillas
            Exit Sub
        End Try
        Dim di As New DirectoryInfo(ruta)
        Dim fiArr As FileInfo() = di.GetFiles()
        Dim frifinal As FileInfo
        Dim j = 0
        For Each fri In fiArr
            If frifinal Is Nothing Then
                Dim frivec As Object
                frivec = Split(fri.Name, ".")
                If frivec(UBound(frivec)) = "xls" Then
                    frifinal = fri
                End If
            End If
            Try
                If frifinal.LastWriteTime < fri.LastWriteTime Then
                    Dim frivec As Object
                    frivec = Split(fri.Name, ".")
                    If frivec(UBound(frivec)) = "xls" Then
                        frifinal = fri
                    End If
                End If
            Catch ex As Exception

            End Try
            j = j + 1
        Next

        Dim sihayruta As Boolean
        sihayruta = False
        Try
            ruta = frifinal.FullName
            sihayruta = True
        Catch ex As Exception
            sihayruta = False
        End Try
        Dim addcolumn1 As New System.Windows.Forms.DataGridViewCheckBoxColumn
        With addcolumn1
            .HeaderText = "Select"
            .Name = "Select"
            .Width = 30
        End With
        Dim addcolumn2 As New System.Windows.Forms.DataGridViewCheckBoxColumn
        With addcolumn1
            .HeaderText = "Select"
            .Name = "Select"
            .Width = 30
        End With
        Dim addcolumn3 As New System.Windows.Forms.DataGridViewCheckBoxColumn
        With addcolumn1
            .HeaderText = "Select"
            .Name = "Select"
            .Width = 30
        End With
        Dim addcolumn4 As New System.Windows.Forms.DataGridViewCheckBoxColumn
        With addcolumn4
            .HeaderText = "Select"
            .Name = "Select"
            .Width = 30
        End With
        DataGridView1.Columns.Insert(DataGridView1.Columns.Count, addcolumn1)
        DataGridView2.Columns.Insert(DataGridView2.Columns.Count, addcolumn2)
        DataGridView3.Columns.Insert(DataGridView3.Columns.Count, addcolumn3)
        DataGridView5.Columns.Insert(DataGridView5.Columns.Count, addcolumn4)
        If sihayruta = True Then
            'LlenaGrillas() se deshabilita por petición del usuario
            'llenacombo2()

        End If


    End Sub
    Public Function versionactual() As String

        Dim cons As String
        Dim version As String = Nothing
        Dim row As DataRow

        cons = "SELECT VersionCargaSif FROM PARAMETROS WITH (NOLOCK) "

        row = Conn.SelReg(cons)

        If row IsNot Nothing Then

            version = row("VersionCargaSif").ToString().Replace(",", ".")

        End If

        Return version

    End Function
    Sub llenacuadros(ByRef grilla As DataGridView)

        Dim filas As System.Windows.Forms.DataGridViewRowCollection
        Dim fila As System.Windows.Forms.DataGridViewRow
        Dim totalplanos, totalesp, cantitem, totalest, area, volumen, peso As String
        totalplanos = 0
        totalesp = 0
        totalest = 0
        area = 0
        volumen = 0
        peso = 0
        Dim j As Integer
        If grilla.Name = "DataGridView1" Then
            j = 0
        ElseIf grilla.Name = "DataGridView2" Then
            j = 1
        ElseIf grilla.Name = "DataGridView3" Then
            j = 2
        ElseIf grilla.Name = "DataGridView5" Then
            j = 3
        End If
        filas = grilla.Rows
        For i = 0 To filas.Count - 1
            fila = filas.Item(i)
            If fila.Cells(0).Value = Nothing Then 'si no hay items fila queda en cero 0
                If DataGridView6.RowCount > 0 Then

                    DataGridView6.Rows.Item(j).Cells(1).Value = 0
                    DataGridView6.Rows.Item(j).Cells(2).Value = 0
                    DataGridView6.Rows.Item(j).Cells(3).Value = 0
                    DataGridView6.Rows.Item(j).Cells(4).Value = 0
                    DataGridView6.Rows.Item(j).Cells(5).Value = 0
                    DataGridView6.Rows.Item(j).Cells(6).Value = 0
                End If
            Else
                cantitem = fila.Cells(0).Value
                area = Math.Round(CDbl(area) + CDbl(fila.Cells(10).Value), 2) 'suma el area total de cada item uno a uno
                volumen = Math.Round(CDbl(volumen) + CDbl(CreaRQI.MyCommands.VolCalc(fila.Cells(1).Value, fila.Cells(9).Value, fila.Cells(8).Value, fila.Cells(3).Value)) * CDbl(fila.Cells(0).Value), 2)
                peso = Math.Round(CDbl(peso) + CDbl(CreaRQI.MyCommands.PesoCalc(fila.Cells(1).Value, fila.Cells(8).Value, fila.Cells(9).Value, fila.Cells(3).Value)) * CDbl(fila.Cells(0).Value), 2)
                If fila.Cells(6).Value = "P" Then
                    totalplanos = CInt(totalplanos) + CInt(1)
                    totalesp = CInt(totalesp) + CInt(cantitem)
                Else
                    totalest = CInt(totalest) + CInt(cantitem)
                End If
            End If
        Next

        'volumen = Math.Round(CDbl(area) * 0.054, 2)
        'peso = Math.Round(CDbl(area) * 25, 2)
        If DataGridView6.RowCount > 0 Then 'si cuadros tiene filas

            DataGridView6.Rows.Item(j).Cells(1).Value = totalplanos
            DataGridView6.Rows.Item(j).Cells(2).Value = totalest
            DataGridView6.Rows.Item(j).Cells(3).Value = totalesp
            DataGridView6.Rows.Item(j).Cells(4).Value = area
            DataGridView6.Rows.Item(j).Cells(5).Value = volumen
            DataGridView6.Rows.Item(j).Cells(6).Value = peso
        End If

        'DataGridView6.Rows.Item(4).Cells(1).Value = CInt(DataGridView6.Rows.Item(0).Cells(1).Value) + CInt(DataGridView6.Rows.Item(1).Cells(1).Value) + CInt(DataGridView6.Rows.Item(2).Cells(1).Value) + CInt(DataGridView6.Rows.Item(3).Cells(1).Value)
        'DataGridView6.Rows.Item(4).Cells(2).Value = CInt(DataGridView6.Rows.Item(0).Cells(2).Value) + CInt(DataGridView6.Rows.Item(1).Cells(2).Value) + CInt(DataGridView6.Rows.Item(2).Cells(2).Value) + CInt(DataGridView6.Rows.Item(3).Cells(2).Value)
        'DataGridView6.Rows.Item(4).Cells(3).Value = CInt(DataGridView6.Rows.Item(0).Cells(3).Value) + CInt(DataGridView6.Rows.Item(1).Cells(3).Value) + CInt(DataGridView6.Rows.Item(2).Cells(3).Value) + CInt(DataGridView6.Rows.Item(3).Cells(3).Value)
        'DataGridView6.Rows.Item(4).Cells(4).Value = CDbl(DataGridView6.Rows.Item(0).Cells(4).Value) + CDbl(DataGridView6.Rows.Item(1).Cells(4).Value) + CDbl(DataGridView6.Rows.Item(2).Cells(4).Value) + CDbl(DataGridView6.Rows.Item(3).Cells(4).Value)
        'DataGridView6.Rows.Item(4).Cells(5).Value = CDbl(DataGridView6.Rows.Item(0).Cells(5).Value) + CDbl(DataGridView6.Rows.Item(1).Cells(5).Value) + CDbl(DataGridView6.Rows.Item(2).Cells(5).Value) + CDbl(DataGridView6.Rows.Item(3).Cells(5).Value)
        'DataGridView6.Rows.Item(4).Cells(6).Value = CDbl(DataGridView6.Rows.Item(0).Cells(6).Value) + CDbl(DataGridView6.Rows.Item(1).Cells(6).Value) + CDbl(DataGridView6.Rows.Item(2).Cells(6).Value) + CDbl(DataGridView6.Rows.Item(3).Cells(6).Value)
    End Sub
    'llena todos los combos del tipo de Orden. Nota: el combo1 o ERP no se usa. El formulario esta deshabilitado.
    Sub llenacombo2()
        Dim cons As String = "SELECT T_Sol_Tipo FROM Tipos_Sol WITH (nolock) WHERE (T_Sol_Activo = 1) AND (NOT (T_Sol_Tipo = 'OK'))"
        Using connection As New SqlConnection(Str_Con_Bd)
            Dim command As SqlCommand = New SqlCommand(cons, connection)

            connection.Open()
            Dim reader As SqlDataReader = command.ExecuteReader()
            If reader.HasRows Then
                ComboBox2.Items.Clear() 'Combo Formaletas
                ComboBox7.Items.Clear() 'Combo Accesorios
                ComboBox1.Items.Clear() 'Combo ERP
                Do While reader.Read()
                    ComboBox2.Items.Add(reader.GetValue(0)) 'Combo Formaletas
                    ComboBox7.Items.Add(reader.GetValue(0)) 'Combo Accesorios
                    ComboBox1.Items.Add(reader.GetValue(0)) 'Combo ERP
                Loop
            End If

            command = Nothing
            reader.Close()
            reader = Nothing
        End Using
    End Sub
    Sub llenacombo3(ByVal selected As String)
        'Dim Cons As String = "SELECT Orden.ano, Orden.Numero, Orden.letra, Orden.Abierta, Orden.Tipo_Of, Orden.Id_Of_P FROM Orden WITH (nolock) GROUP BY " +
        '    "Orden.ano, Orden.Numero, Orden.letra, Orden.Abierta, Orden.Tipo_Of, Orden.Id_Of_P HAVING (((Orden.Tipo_Of) " +
        '    "= '" & selected & "') "
        'If selected = "OG" Or selected = "OM" Or selected = "PV" Or selected = "RC" Or selected = "ID" Then
        '    Cons = Cons & " ) ORDER BY Orden.ano DESC, Orden.Numero DESC, Orden.letra DESC" 'And (Abierta = 1)
        'Else
        '    Cons = Cons & "AND ((Orden.letra)='1')) ORDER BY Orden.ano DESC, Orden.Numero DESC, Orden.letra DESC" 'And (Abierta = 1)
        'End If

        Dim Cons As String = "SELECT    ano," +
                              "Numero,  letra," +
                              "Abierta, Tipo_Of, Id_Of_P " +
                              "FROM ORDEN WITH (NOLOCK)  " +
                              "WHERE Tipo_Of = '" + selected + "' "

        If selected <> "OG" Or selected <> "OM" Or selected <> "PV" Or selected <> "RC" Or selected <> "ID" Then

            Cons = Cons + " AND letra = '1' " +
                          " AND planta_id = '" + Me.PlantaId.ToString() + "'" +
                          " ORDER BY ano DESC, Numero DESC, letra DESC"

        End If

        Using connection As New SqlConnection(Str_Con_Bd)

            Dim command As SqlCommand = New SqlCommand(Cons, connection)

            connection.Open()
            Dim reader As SqlDataReader = command.ExecuteReader()
            If reader.HasRows Then
                ComboBox3.Items.Clear()
                ComboBox6.Items.Clear()

                Do While reader.Read()
                    ComboBox3.Items.Add(reader.GetValue(1) & "-" & reader.GetValue(0))
                    ComboBox6.Items.Add(reader.GetValue(1) & "-" & reader.GetValue(0))
                Loop
            End If

            command = Nothing
            reader.Close()
            reader = Nothing

        End Using

    End Sub
    Sub llenacombo4()
        Dim Cons As String = "SELECT Orden.letra AS Raya, Orden.Id_Ofa FROM Orden WITH (nolock) " +
            "WHERE (Orden.Tipo_Of = '" & ComboBox2.SelectedItem & "') AND (RTRIM(Orden.Numero) + '-' + RTRIM(Orden.ano) LIKE '" & ComboBox3.SelectedItem & "%') AND (NOT (letra LIKE '%M')) AND (Anulada = 0) "
        If ComboBox2.SelectedItem = "CT" Then
            Cons = Cons & "AND (No_Mod = " & ComboBox9.SelectedItem & ") "
        End If
        Cons = Cons & "ORDER BY Id_Ofa" '"ORDER BY RTRIM(Orden.Numero) + '-' + RTRIM(Orden.ano) DESC, Raya"
        '(Orden.Abierta = 0) AND 
        Using connection As New SqlConnection(Str_Con_Bd)
            Dim command As SqlCommand = New SqlCommand(Cons, connection)

            connection.Open()

            Dim reader As SqlDataReader = command.ExecuteReader()

            If reader.HasRows Then
                ComboBox4.Items.Clear()
                Do While reader.Read()
                    If reader.GetValue(0) < 51 Then
                        ComboBox4.Items.Add(reader.GetValue(0))
                    End If

                Loop
                If ComboBox4.Items.Count = 0 Then
                    rayavacia = 1
                    ComboBox4.Items.Add(rayavacia) 'agregar unico item
                Else
                    rayasantesdevacia = ComboBox4.Items.Item(ComboBox4.Items.Count - 1)
                    rayavacia = ComboBox4.Items.Item(ComboBox4.Items.Count - 1) + 1
                    ComboBox4.Items.Add(rayavacia) 'agregar ultimo item
                End If

            Else
                ComboBox4.Items.Clear()
                ComboBox4.Items.Add(1)
                rayavacia = 1
            End If
            command = Nothing
            reader.Close()
            reader = Nothing
        End Using
    End Sub
    Sub llenacombo11()
        Dim UnidadNegocioExiste As Boolean
        Dim cons As String

        cons = "SELECT Unidad_Negocio.Und_Neg_Nombre FROM Orden INNER JOIN Unidad_Negocio ON Orden.Ord_Unidad_Neg = Unidad_Negocio.Und_Neg_Id WHERE  (Orden.Ofa = '" & ComboBox3.Text & "-1')"

        Using connection As New SqlConnection(Str_Con_Bd)
            Dim command As SqlCommand = New SqlCommand(cons, connection)

            connection.Open()

            Dim reader As SqlDataReader = command.ExecuteReader()

            If reader.HasRows Then
                ComboBox11.Items.Clear()
                Do While reader.Read()
                    ComboBox11.Text = reader.GetValue(0)
                    If reader.GetValue(0).ToString = "Ninguna" Then
                        'ComboBox11.Enabled = True
                        UnidadNegocioExiste = False
                    Else
                        'ComboBox11.Enabled = False
                        UnidadNegocioExiste = True
                    End If
                Loop
            End If
            command = Nothing
            reader.Close()
            reader = Nothing
        End Using
        If UnidadNegocioExiste = False Then
            cons = "SELECT Und_Neg_Nombre, Und_Neg_Id FROM Unidad_Negocio WHERE  (Und_Neg_Inactivo = 0)"
            Using connection As New SqlConnection(Str_Con_Bd)
                Dim command As SqlCommand = New SqlCommand(cons, connection)

                connection.Open()

                Dim reader As SqlDataReader = command.ExecuteReader()

                If reader.HasRows Then
                    ComboBox11.Items.Clear()
                    Do While reader.Read()
                        ComboBox11.Items.Add(reader.GetValue(0))
                    Loop
                End If
                command = Nothing
                reader.Close()
                reader = Nothing
            End Using
        End If

    End Sub

    Sub verificacumplimiento()
        Dim ODplanta, ODalmacen, ODalum, IForden, plantaidOD As Integer
        Dim cons As String
        cons = "SELECT Orden_Seg.Id_Ofa, REPLACE(Orden_Seg.Num_Of + '-' + Orden_Seg.Ano_Of, ' ', '') AS NUmero_Orden, Orden_Seg.OP_1EE_Abuelo AS IF_Orden, Orden_Seg.Item_1EE_Abuelo AS ItemIF, Orden_Seg.Item_Ac_Al AS ItemAC, Orden_Seg.Item_Ac_Com AS ItemCOM, Orden_Seg.ND_1EE AS ItemALUM, Orden_Seg.OP_Ac_Al AS OD_PLANTA, Orden_Seg.OP_Ac_Com AS OD_ALMACEN, Orden_Seg.OP_1EE AS OD_ALUM, Orden_Seg.planta_id " +
                "FROM Orden_Seg WITH (NOLOCK) INNER JOIN Orden ON Orden_Seg.Id_Seg_Of = Orden.ordenseg_id AND Orden_Seg.Id_Ofa = Orden.Id_Ofa " +
                "WHERE (Orden.Numero + '-' + Orden.ano = '" & ComboBox6.SelectedItem & "') AND (Orden.Tipo_Of = '" & ComboBox7.SelectedItem & "')"

        Using connection As New SqlConnection(Str_Con_Bd)
            Dim command As SqlCommand = New SqlCommand(cons, connection)

            connection.Open()
            Dim reader As SqlDataReader = command.ExecuteReader()
            If reader.HasRows Then
                reader.Read()
                ODplanta = reader.GetValue(7)
                ODalmacen = reader.GetValue(8)
                ODalum = reader.GetValue(9)
                plantaidOD = reader.GetValue(10)
                IForden = reader.GetValue(2)
            Else

            End If
            command = Nothing
            reader.Close()
            reader = Nothing
        End Using
        If plantaidOD = 3 Then
            ComboBox8.Enabled = True
            TextBox2.Enabled = True
            TextBox2.BackColor = System.Drawing.Color.White
            ComboBox5.Enabled = True
            TextBox3.Enabled = True
            TextBox3.BackColor = System.Drawing.Color.White
            estatuserplibre = True
        Else
            If consultacumplimineto("OD", ODplanta, plantaidOD) = True Then
                ComboBox8.Enabled = False
                TextBox2.Enabled = False
                TextBox2.BackColor = System.Drawing.Color.Red
            Else
                ComboBox8.Enabled = True
                TextBox2.Enabled = True
                TextBox2.BackColor = System.Drawing.Color.White
            End If
            If consultacumplimineto("OD", ODalmacen, plantaidOD) = True Then
                ComboBox5.Enabled = False
                TextBox3.Enabled = False
                TextBox3.BackColor = System.Drawing.Color.Red
            Else
                ComboBox5.Enabled = True
                TextBox3.Enabled = True
                TextBox3.BackColor = System.Drawing.Color.White
            End If
            If consultacumplimineto("IF", IForden, plantaidOD) = True Then
                estatuserplibre = False
            Else
                estatuserplibre = True
            End If
        End If

    End Sub
    Function consultacumplimineto(ByVal tipood As String, ByVal numerood As Integer, ByVal plantaid As Integer) As Boolean
        If numerood = 0 Then
            consultacumplimineto = False
        Else
            Dim cons As String
            cons = "DECLARE @Planta INT = '" & plantaid & "' DECLARE @Tipo VARCHAR(50) = '" & tipood & "' DECLARE @NumOrden VARCHAR(50) = '" & numerood & "' DECLARE @myEstado VARCHAR(50) DECLARE @IdEstado AS INT " +
                    "EXEC @IdEstado = dbo.proErpEstadoOrden @Planta, @Tipo, @NumOrden, @myEstado OUTPUT " +
                    "SELECT @IdEstado as idestado, @myEstado as estado"
            Using connection As New SqlConnection(Str_Con_Bd)
                Dim command As SqlCommand = New SqlCommand(cons, connection)

                connection.Open()
                Dim reader As SqlDataReader = command.ExecuteReader()
                If reader.HasRows Then
                    reader.Read()
                    If reader.GetValue(0) = 3 Then
                        consultacumplimineto = True
                    Else
                        consultacumplimineto = False
                    End If
                Else

                End If
                command = Nothing
                reader.Close()
                reader = Nothing
            End Using
        End If

    End Function
    Sub llenacombo5y8()
        verificacumplimiento()
        If estatuserplibre = False Then
            impidereemplazar = True
            CheckBox3.Checked = False
            Button19.Enabled = False
            CheckBox3.Enabled = False
            Button19.BackColor = System.Drawing.Color.Red
            Button19.Text = "cerrada"
        Else
            impidereemplazar = False
            'CheckBox3.Checked = False
            If ComboBox7.SelectedItem = "CT" And itemssincodigo = True Then
                'no habilita boton solo si hay items en amarillo 
            Else
                Button19.Enabled = True 'si no es CT habilita el boon
                CheckBox3.Enabled = True
            End If

            Button19.BackColor = System.Drawing.Color.Transparent
            Button19.Text = "Cargar Orden"
            ComboBox5.Items.Clear()
            ComboBox8.Items.Clear()
            Dim cons As String
            cons = "SELECT Id_Ofa, Tipo_Of, Numero + '-' + ano AS No_Sol, letra FROM Orden WITH (nolock) " +
                "WHERE (Tipo_Of = '" & ComboBox7.SelectedItem & "') AND (Numero + '-' + ano = '" & ComboBox6.SelectedItem & "') "
            If ComboBox7.SelectedItem = "CT" Then
                cons = cons & "AND (No_Mod = " & ComboBox10.SelectedItem & ") "
            End If
            cons = cons & "ORDER BY letra"
            Using connection As New SqlConnection(Str_Con_Bd)
                Dim command As SqlCommand = New SqlCommand(cons, connection)

                connection.Open()
                Dim reader As SqlDataReader = command.ExecuteReader()
                If reader.HasRows Then


                    Do While reader.Read()
                        If ComboBox7.SelectedItem = "OG" Or ComboBox7.SelectedItem = "OM" Or ComboBox7.SelectedItem = "PV" Or ComboBox7.SelectedItem = "RC" Or ComboBox7.SelectedItem = "ID" Then
                            ComboBox5.Items.Add("0")
                            ComboBox5.SelectedIndex = 0
                            ComboBox8.Items.Add("0")
                            ComboBox8.SelectedIndex = 0
                        Else
                            Try
                                Dim raya As Integer = reader.GetValue(3)
                                If raya > 50 Then
                                    If raya Mod 2 = 0 Then 'par
                                        ComboBox8.Items.Add(raya)
                                    Else 'impar
                                        ComboBox5.Items.Add(raya)
                                    End If
                                End If
                            Catch ex As Exception

                            End Try


                        End If
                    Loop



                End If
                If ComboBox7.SelectedItem = "OG" Or ComboBox7.SelectedItem = "OM" Or ComboBox7.SelectedItem = "PV" Or ComboBox7.SelectedItem = "RC" Or ComboBox7.SelectedItem = "ID" Then

                Else
                    Dim cantimpar = ComboBox5.Items.Count
                    Dim cantpar = ComboBox8.Items.Count
                    If cantimpar = 0 Then
                        rayaimparvacia = 51
                        ComboBox5.Items.Add(51)
                    Else
                        rayaimparvacia = ComboBox5.Items.Item(cantimpar - 1) + 2
                        ComboBox5.Items.Add(rayaimparvacia)
                    End If
                    If cantpar = 0 Then
                        rayaparvacia = 52
                        ComboBox8.Items.Add(52)
                    Else
                        rayaparvacia = ComboBox8.Items.Item(cantpar - 1) + 2
                        ComboBox8.Items.Add(rayaparvacia)
                    End If
                End If
                command = Nothing
                reader.Close()
                reader = Nothing
            End Using

        End If

    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        CreaRQI.MyCommands.verificaitems(True, False, grillas, Tab_VerInfoCargada, , , , , CIA)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        CreaRQI.MyCommands.verificaitems(False, True, grillas, Tab_VerInfoCargada, , , , libroExcel, CIA)
        CreaRQI.MyCommands.GuardaLista(grillas, libroExcel)
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        CreaRQI.MyCommands.separaUML(DataGridView2, libroExcel)
        CreaRQI.MyCommands.GuardaLista(grillas, libroExcel)
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        ReDim grills(3)
        grills(0) = DataGridView1
        grills(1) = DataGridView2
        grills(2) = DataGridView3
        grills(3) = DataGridView5
        tabcontroluno = Tab_VerInfoCargada
        DataGridView1.Enabled = False
        DataGridView2.Enabled = False
        DataGridView3.Enabled = False
        DataGridView5.Enabled = False
        Dim inventario As New Frm_CalculaDisponibles
        With inventario
            .TopMost = True
            .Show()
            .TopMost = False
        End With
        DataGridView1.Enabled = True
        DataGridView2.Enabled = True
        DataGridView3.Enabled = True
        DataGridView5.Enabled = True
    End Sub
    'Public Shared Function datagridviewuno1() As System.Windows.Forms.DataGridView
    'Return datagridviewuno
    'End Function
    'Public Shared Function datagridviewdos2() As System.Windows.Forms.DataGridView
    'Return datagridviewdos
    'End Function
    'Public Shared Function datagridviewtres3() As System.Windows.Forms.DataGridView
    'Return datagridviewtres
    'End Function

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        CreaRQI.MyCommands.separautilizadas(grillas, libroExcel)
    End Sub

    Function grillas() As Object
        Dim grilla2 As Object
        ReDim grilla2(3)
        grilla2(0) = DataGridView1
        grilla2(1) = DataGridView2
        grilla2(2) = DataGridView3
        grilla2(3) = DataGridView5
        Return grilla2
    End Function
    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        grillasvisibles()
        mensa = Nothing
        CreaRQI.MyCommands.verificaitems(True, False, grillas, Tab_VerInfoCargada, False, , , , CIA)
        If mensa = MsgBoxResult.Ok Then
        Else
            Dim forma As System.Windows.Forms.Form = Me
            'Me.TopMost = True
            Me.Enabled = False
            Dim genera As New Frm_GenerarRQI
            With genera
                '.TopMost = True
                .forma2 = forma
                .Show()
            End With
        End If

    End Sub
    'Private Sub Form1_Activated(sender As Object, e As EventArgs) Handles Me.Activated
    'Me.Enabled = True
    'End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        seleclista()
    End Sub
    Sub seleclista()
        Dim rutavec As Object
        rutavec = Split(ruta, "\")
        ruta = Nothing
        For i = 0 To UBound(rutavec) - 1
            ruta = ruta & rutavec(i) & "\"
        Next
        OpenFileDialog1.InitialDirectory = ruta
        ruta = Nothing
        If OpenFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            ruta = OpenFileDialog1.FileName
        End If
        'If FolderBrowserDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
        'ruta = FolderBrowserDialog1.SelectedPath
        'End If
        LlenaGrillas()
    End Sub
    Sub grillasvisibles()
        Try
            For Each grilla In grills
                For Each ro In grilla.rows
                    ro.visible = True
                Next
            Next
        Catch ex As Exception

        End Try

    End Sub
    Sub cargalinea(datosfila As String(), ByRef datagrid As System.Windows.Forms.DataGridView)
        Try
            'Dim ver = CInt(Obj2.Value)
            datagrid.Rows.Add()
            Dim datosredondear() As Integer = {2, 3, 4, 5, 9, 10} 'filas que contienen valores que se podrian redondear
            For Each datoredondear In datosredondear
                If datosfila(datoredondear) <> "" Then 'verificar que no sea cero para que aparezca vacio en la grilla
                    datosfila(datoredondear) = CStr(Math.Round(CDbl(datosfila(datoredondear)), 3))
                End If
            Next

            For j = 0 To 13 'recorreo la celda de excel y agrega los datos al recorrido del datagrid
                'Obj = CType(Range.Cells(rowCnt, j + 1), Microsoft.Office.Interop.Excel.Range)
                If j = 11 Then
                ElseIf j = 6 Then
                    If datosfila(j) <> "p" And datosfila(j) <> "P" And datosfila(j) <> "" Then
                        campoPlanoConFormato = False
                        Button12.Enabled = False
                    End If
                    datagrid.Rows.Item(datagrid.Rows.Count - 2).Cells(j).Value = datosfila(j)
                ElseIf j = 12 Then
                    Dim codigo, cantitem, tipoitem, desaux, an1, al1, al2, an2 As String
                    cantitem = datagrid.Rows.Item(datagrid.Rows.Count - 2).Cells(0).Value
                    tipoitem = datagrid.Rows.Item(datagrid.Rows.Count - 2).Cells(1).Value
                    an1 = datagrid.Rows.Item(datagrid.Rows.Count - 2).Cells(2).Value
                    al1 = datagrid.Rows.Item(datagrid.Rows.Count - 2).Cells(3).Value
                    If datagrid.Rows.Item(datagrid.Rows.Count - 2).Cells(4).Value = "" Then
                        al2 = 0
                    Else
                        al2 = datagrid.Rows.Item(datagrid.Rows.Count - 2).Cells(4).Value
                    End If
                    If datagrid.Rows.Item(datagrid.Rows.Count - 2).Cells(5).Value = "" Then
                        an2 = 0
                    Else
                        an2 = datagrid.Rows.Item(datagrid.Rows.Count - 2).Cells(5).Value
                    End If
                    desaux = datagrid.Rows.Item(datagrid.Rows.Count - 2).Cells(7).Value
                    elijeKamban(codigo, cantitem, tipoitem, desaux, an1, al1, al2, an2, 1)
                    datagrid.Rows.Item(datagrid.Rows.Count - 2).Cells(j).Value = codigo
                ElseIf j = 13 Then
                    Dim idpiezasforsa As String
                    Dim tipoitem = datagrid.Rows.Item(datagrid.Rows.Count - 2).Cells(1).Value
                    Dim cons = "SELECT Id_Piezas FROM Piezas_Forsa WITH (nolock) WHERE ([Desc] = '" & tipoitem & "')"
                    Using connection As New SqlConnection(Str_Con_Bd)
                        Dim command As SqlCommand = New SqlCommand(cons, connection)
                        connection.Open()
                        Dim reader As SqlDataReader = command.ExecuteReader()

                        If reader.HasRows Then
                            reader.Read()
                            idpiezasforsa = reader.GetValue(0)
                        End If
                        command = Nothing
                        reader.Close()
                        reader = Nothing
                    End Using
                    If idpiezasforsa = 0 Or idpiezasforsa = Nothing Then
                        datagrid.Rows.Item(datagrid.Rows.Count - 2).Cells(j).Value = 0
                        datagrid.Rows.Item(datagrid.Rows.Count - 2).DefaultCellStyle.BackColor = System.Drawing.Color.Yellow

                        mensajesfinales(UBound(mensajesfinales)) = "Debe solicitar a Mesa de Ayuda la creacion de ***" & tipoitem & "*** (idPiezasForsa)"
                        ReDim Preserve mensajesfinales(UBound(mensajesfinales) + 1)

                        'MsgBox("Debe solicitar a Mesa de Ayuda la creacion de ***" & tipoitem & "*** (idPiezasForsa)",, "esto NO es un error")
                        Button12.Enabled = False
                        Button12.Visible = False
                    Else
                        datagrid.Rows.Item(datagrid.Rows.Count - 2).Cells(j).Value = idpiezasforsa
                    End If
                ElseIf j = 9 And datosfila(j) = Nothing Then 'si el listado no provee area, la calcula
                    Dim tipoitem = datagrid.Rows.Item(datagrid.Rows.Count - 2).Cells(1).Value

                    Dim an1, al1, al2, an2 As Double
                    Double.TryParse(datagrid.Rows.Item(datagrid.Rows.Count - 2).Cells(2).Value, an1)
                    Double.TryParse(datagrid.Rows.Item(datagrid.Rows.Count - 2).Cells(3).Value, al1)
                    Double.TryParse(datagrid.Rows.Item(datagrid.Rows.Count - 2).Cells(4).Value, al2)
                    Double.TryParse(datagrid.Rows.Item(datagrid.Rows.Count - 2).Cells(5).Value, an2)
                    Dim desaux = datagrid.Rows.Item(datagrid.Rows.Count - 2).Cells(7).Value
                    Dim grupo = datagrid.Rows.Item(datagrid.Rows.Count - 2).Cells(8).Value
                    datagrid.Rows.Item(datagrid.Rows.Count - 2).Cells(j).Value = CreaRQI.MyCommands.AreaCalc(tipoitem, an1, al1, al2, an2, desaux, grupo)
                ElseIf j = 10 Then 'Campiño que los calcule con los m2 unitarios x las cantidades del listado 'And datosfila(j) = Nothing
                    datagrid.Rows.Item(datagrid.Rows.Count - 2).Cells(j).Value = CStr(CInt(datagrid.Rows.Item(datagrid.Rows.Count - 2).Cells(0).Value) * CDbl(datagrid.Rows.Item(datagrid.Rows.Count - 2).Cells(9).Value))
                Else
                    datagrid.Rows.Item(datagrid.Rows.Count - 2).Cells(j).Value = datosfila(j) 'DataGridView1.Rows.Item(i - 1).Cells(j).Value = Obj.Value
                End If

            Next
            datagrid.Rows.Item(datagrid.Rows.Count - 2).Cells(16).Value = datagrid.Rows.Item(datagrid.Rows.Count - 2).Cells(0).Value
            Dim nom As String = datagrid.Rows.Item(datagrid.Rows.Count - 2).Cells(1).Value
            Dim cant = 10 - nom.Length
            For i = 0 To cant - 1
                nom = nom & 0
            Next
            Dim alto = Strings.Format(datagrid.Rows.Item(datagrid.Rows.Count - 2).Cells(3).Value, "0000.00")
            Dim ancho = Strings.Format(datagrid.Rows.Item(datagrid.Rows.Count - 2).Cells(2).Value, "0000.00")
            datagrid.Rows.Item(datagrid.Rows.Count - 2).Cells(17).Value = nom & alto & ancho
            llenacuadros(datagrid)
        Catch ex As Exception

        End Try

    End Sub
    Sub cargalineaacc(datosfila As String(), ByRef datagrid As System.Windows.Forms.DataGridView, codigo As String, nomlargo As String, cantvalida As String, codigosId As String)
        Dim dim1, dim2, dim3, dim4, dim5, dim6 As Double
        Try

            'Dim ver = CInt(Obj2.Value)
            datagrid.Rows.Add()
            datagrid.Rows.Item(datagrid.Rows.Count - 2).Cells(0).Value = datosfila(0) 'cant
            datagrid.Rows.Item(datagrid.Rows.Count - 2).Cells(1).Value = datosfila(1) 'descri
            datagrid.Rows.Item(datagrid.Rows.Count - 2).Cells(9).Value = datosfila(2) 'desaux
            datagrid.Rows.Item(datagrid.Rows.Count - 2).Cells(2).Value = datosfila(3) 'dim1
            datagrid.Rows.Item(datagrid.Rows.Count - 2).Cells(3).Value = datosfila(4) 'dim2
            datagrid.Rows.Item(datagrid.Rows.Count - 2).Cells(4).Value = datosfila(5) 'dim3
            datagrid.Rows.Item(datagrid.Rows.Count - 2).Cells(5).Value = datosfila(6) 'dim4
            datagrid.Rows.Item(datagrid.Rows.Count - 2).Cells(6).Value = datosfila(7) 'dim5
            datagrid.Rows.Item(datagrid.Rows.Count - 2).Cells(7).Value = datosfila(8) 'dim6


            Try
                If datosfila(3) = "" Then dim1 = 0 Else dim1 = datosfila(3)
            Catch ex2 As Exception
                mensajesfinales(UBound(mensajesfinales)) = "la Dim1 del item " & datagrid.Rows.Count - 1 & " - " & datosfila(1) & " no puede ser un texto"
                ReDim Preserve mensajesfinales(UBound(mensajesfinales) + 1)
            End Try
            Try
                If datosfila(4) = "" Then dim2 = 0 Else dim2 = datosfila(4)
            Catch ex2 As Exception
                mensajesfinales(UBound(mensajesfinales)) = "la Dim2 del item " & datagrid.Rows.Count - 1 & " - " & datosfila(1) & " no puede ser un texto"
                ReDim Preserve mensajesfinales(UBound(mensajesfinales) + 1)
            End Try
            Try
                If datosfila(5) = "" Then dim3 = 0 Else dim3 = datosfila(5)
            Catch ex2 As Exception
                mensajesfinales(UBound(mensajesfinales)) = "la Dim3 del item " & datagrid.Rows.Count - 1 & " - " & datosfila(1) & " no puede ser un texto"
                ReDim Preserve mensajesfinales(UBound(mensajesfinales) + 1)
            End Try
            Try
                If datosfila(6) = "" Then dim4 = 0 Else dim4 = datosfila(6)
            Catch ex2 As Exception
                mensajesfinales(UBound(mensajesfinales)) = "la Dim4 del item " & datagrid.Rows.Count - 1 & " - " & datosfila(1) & " no puede ser un texto"
                ReDim Preserve mensajesfinales(UBound(mensajesfinales) + 1)
            End Try
            Try
                If datosfila(7) = "" Then dim5 = 0 Else dim5 = datosfila(7)
            Catch ex2 As Exception
                mensajesfinales(UBound(mensajesfinales)) = "la Dim5 del item " & datagrid.Rows.Count - 1 & " - " & datosfila(1) & " no puede ser un texto"
                ReDim Preserve mensajesfinales(UBound(mensajesfinales) + 1)
            End Try
            Try
                If datosfila(8) = "" Then dim6 = 0 Else dim6 = datosfila(8)
            Catch ex2 As Exception
                mensajesfinales(UBound(mensajesfinales)) = "la Dim6 del item " & datagrid.Rows.Count - 1 & " - " & datosfila(1) & " no puede ser un texto"
                ReDim Preserve mensajesfinales(UBound(mensajesfinales) + 1)
            End Try

            'If TypeOf datagrid.Rows.Item(datagrid.Rows.Count - 2).Cells(2).Value Is Double Or TypeOf datagrid.Rows.Item(datagrid.Rows.Count - 2).Cells(2).Value Is Integer Then
            'ElseIf datagrid.Rows.Item(datagrid.Rows.Count - 2).Cells(2).Value Is Nothing Then
            'Else

            'End If
            'If TypeOf datagrid.Rows.Item(datagrid.Rows.Count - 2).Cells(3).Value Is Double Or TypeOf datagrid.Rows.Item(datagrid.Rows.Count - 2).Cells(3).Value Is Integer Then
            'ElseIf datagrid.Rows.Item(datagrid.Rows.Count - 2).Cells(3).Value Is Nothing Then
            'Else
            '    mensajesfinales(UBound(mensajesfinales)) = "la Dim2 del item " & datagrid.Rows.Count - 1 & " - " & datosfila(1) & " no puede ser un texto"
            '    ReDim Preserve mensajesfinales(UBound(mensajesfinales) + 1)
            'End If
            'If TypeOf datagrid.Rows.Item(datagrid.Rows.Count - 2).Cells(4).Value Is Double Or TypeOf datagrid.Rows.Item(datagrid.Rows.Count - 2).Cells(4).Value Is Integer Then
            'ElseIf datagrid.Rows.Item(datagrid.Rows.Count - 2).Cells(4).Value Is Nothing Then
            'Else
            '    mensajesfinales(UBound(mensajesfinales)) = "la Dim3 del item " & datagrid.Rows.Count - 1 & " - " & datosfila(1) & " no puede ser un texto"
            '    ReDim Preserve mensajesfinales(UBound(mensajesfinales) + 1)
            'End If
            'If TypeOf datagrid.Rows.Item(datagrid.Rows.Count - 2).Cells(5).Value Is Double Or TypeOf datagrid.Rows.Item(datagrid.Rows.Count - 2).Cells(5).Value Is Integer Then
            'ElseIf datagrid.Rows.Item(datagrid.Rows.Count - 2).Cells(5).Value Is Nothing Then
            'Else
            '    mensajesfinales(UBound(mensajesfinales)) = "la Dim4 del item " & datagrid.Rows.Count - 1 & " - " & datosfila(1) & " no puede ser un texto"
            '    ReDim Preserve mensajesfinales(UBound(mensajesfinales) + 1)
            'End If
            'If TypeOf datagrid.Rows.Item(datagrid.Rows.Count - 2).Cells(6).Value Is Double Or TypeOf datagrid.Rows.Item(datagrid.Rows.Count - 2).Cells(6).Value Is Integer Then
            'ElseIf datagrid.Rows.Item(datagrid.Rows.Count - 2).Cells(6).Value Is Nothing Then
            'Else
            '    mensajesfinales(UBound(mensajesfinales)) = "la Dim5 del item " & datagrid.Rows.Count - 1 & " - " & datosfila(1) & " no puede ser un texto"
            '    ReDim Preserve mensajesfinales(UBound(mensajesfinales) + 1)
            'End If
            'If TypeOf datagrid.Rows.Item(datagrid.Rows.Count - 2).Cells(7).Value Is Double Or TypeOf datagrid.Rows.Item(datagrid.Rows.Count - 2).Cells(7).Value Is Integer Then
            'ElseIf datagrid.Rows.Item(datagrid.Rows.Count - 2).Cells(7).Value Is Nothing Then
            'Else
            '    mensajesfinales(UBound(mensajesfinales)) = "la Dim6 del item " & datagrid.Rows.Count - 1 & " - " & datosfila(1) & " no puede ser un texto"
            '    ReDim Preserve mensajesfinales(UBound(mensajesfinales) + 1)
            'End If

            'revisar si se pone P
            'plano
            If datosfila(9) = "p" Then
                datosfila(9) = "P"
            End If
            If datosfila(9) = "P" Then
                TextBox4.Text = CInt(TextBox4.Text) + 1
            End If
            datagrid.Rows.Item(datagrid.Rows.Count - 2).Cells(8).Value = datosfila(9) 'plano
            'verifica codigo
            datagrid.Rows.Item(datagrid.Rows.Count - 2).Cells(10).Value = codigo

            'trae nombre
            datagrid.Rows.Item(datagrid.Rows.Count - 2).Cells(11).Value = nomlargo 'nombre item

            'datagrid.Rows.Item(datagrid.Rows.Count - 2).Cells(16).Value = datagrid.Rows.Item(datagrid.Rows.Count - 2).Cells(0).Value

            datagrid.Rows.Item(datagrid.Rows.Count - 2).Cells(12).Value = cantvalida 'valida

            datagrid.Rows.Item(datagrid.Rows.Count - 2).Cells(15).Value = codigosId 'id de accesorios codigos

            If codigo = 0 Then
                datagrid.Rows.Item(datagrid.Rows.Count - 2).DefaultCellStyle.BackColor = System.Drawing.Color.Yellow
                If ComboBox7.Text = "CT" Then

                    Button19.Enabled = False
                    CheckBox3.Enabled = False
                    itemssincodigo = True

                End If
            End If
            'Dim nom As String = datagrid.Rows.Item(datagrid.Rows.Count - 2).Cells(1).Value
            'Dim cant = 10 - nom.Length
            'For i = 0 To cant - 1
            'nom = nom & 0
            'Next
            'Dim alto = Strings.Format(datagrid.Rows.Item(datagrid.Rows.Count - 2).Cells(3).Value, "0000.00")
            'Dim ancho = Strings.Format(datagrid.Rows.Item(datagrid.Rows.Count - 2).Cells(2).Value, "0000.00")
            'datagrid.Rows.Item(datagrid.Rows.Count - 2).Cells(17).Value = nom & alto & ancho
            'llenacuadros(datagrid)

        Catch ex As Exception

        End Try

    End Sub
    Sub cargalineafup(ByVal nom As String, ByVal cant As String, ByRef datagrid As System.Windows.Forms.DataGridView, codigo As String, nomlargo As String, cantvalida As String)
        Try
            datagrid.Rows.Add()
            datagrid.Rows.Item(datagrid.Rows.Count - 2).DefaultCellStyle.BackColor = System.Drawing.Color.LightBlue
            datagrid.Rows.Item(datagrid.Rows.Count - 2).Cells(0).Value = cant
            datagrid.Rows.Item(datagrid.Rows.Count - 2).Cells(1).Value = nom
            datagrid.Rows.Item(datagrid.Rows.Count - 2).Cells(9).Value = "" 'desaux
            datagrid.Rows.Item(datagrid.Rows.Count - 2).Cells(2).Value = "" 'dim1
            datagrid.Rows.Item(datagrid.Rows.Count - 2).Cells(3).Value = "" 'dim2
            datagrid.Rows.Item(datagrid.Rows.Count - 2).Cells(4).Value = "" 'dim3
            datagrid.Rows.Item(datagrid.Rows.Count - 2).Cells(5).Value = "" 'dim4
            datagrid.Rows.Item(datagrid.Rows.Count - 2).Cells(6).Value = "" 'dim5
            datagrid.Rows.Item(datagrid.Rows.Count - 2).Cells(7).Value = "" 'dim6
            'revisar si se pone P
            datagrid.Rows.Item(datagrid.Rows.Count - 2).Cells(8).Value = "" 'plano
            'verifica codigo
            datagrid.Rows.Item(datagrid.Rows.Count - 2).Cells(10).Value = codigo
            'trae nombre
            datagrid.Rows.Item(datagrid.Rows.Count - 2).Cells(11).Value = nomlargo 'nombre item

            datagrid.Rows.Item(datagrid.Rows.Count - 2).Cells(12).Value = cantvalida 'valida

            datagrid.Rows.Item(datagrid.Rows.Count - 2).Cells(15).Value = 0 'codigo id

            If codigo = 0 Then
                datagrid.Rows.Item(datagrid.Rows.Count - 2).DefaultCellStyle.BackColor = System.Drawing.Color.Yellow
            End If

        Catch ex As Exception

        End Try
    End Sub
    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        CreaRQI.MyCommands.SeparaMarcadas(grillas, libroExcel)
        CreaRQI.MyCommands.GuardaLista(grillas, libroExcel)
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        seleclista()
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Button1.Enabled = True
        Button2.Enabled = True
        Dim CiaVec() As String = Split(ComboBox1.Text, " ")
        CIA = CiaVec(0)
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        Dim a As System.Drawing.Size
        If DataGridView1.Size.Width = 883 Then
            a.Width = 650
            a.Height = 457
            DataGridView1.MaximumSize = a
        Else
            a.Width = 883
            a.Height = 457
            DataGridView1.MaximumSize = a
            DataGridView1.Width = 883
        End If
        If DataGridView2.Size.Width = 883 Then
            a.Width = 650
            a.Height = 457
            DataGridView2.MaximumSize = a
        Else
            a.Width = 883
            a.Height = 457
            DataGridView2.MaximumSize = a
            DataGridView2.Width = 883
        End If
        If DataGridView3.Size.Width = 883 Then
            a.Width = 650
            a.Height = 457
            DataGridView3.MaximumSize = a
        Else
            a.Width = 883
            a.Height = 457
            DataGridView3.MaximumSize = a
            DataGridView3.Width = 883
        End If
        If DataGridView5.Size.Width = 883 Then
            a.Width = 650
            a.Height = 457
            DataGridView5.MaximumSize = a
        Else
            a.Width = 883
            a.Height = 457
            DataGridView5.MaximumSize = a
            DataGridView5.Width = 883
        End If
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        CreaRQI.MyCommands.eliminafilas(grillas)
        llenacuadros(DataGridView1)
        llenacuadros(DataGridView2)
        llenacuadros(DataGridView3)
        llenacuadros(DataGridView5)
        Dim tabactual = Tab_VerInfoCargada.SelectedTab
        Tab_VerInfoCargada.SelectedTab = TabPage_Muros
        Tab_VerInfoCargada.SelectedTab = TabPage_Union
        Tab_VerInfoCargada.SelectedTab = TabPage_Losas
        Tab_VerInfoCargada.SelectedTab = TabPage_Culat
        Tab_VerInfoCargada.SelectedTab = tabactual
    End Sub

    Function verificasaldosp(idofa As String) As Boolean
        Dim cons As String = "SELECT Saldos_P.Saldos_PId FROM Saldos WITH (nolock) INNER JOIN Saldos_P ON Saldos.Identificador = Saldos_P.Saldos_PId WHERE (Saldos.Id_Ofa = " & idofa & ")"
        Using connection As New SqlConnection(Str_Con_Bd)
            Dim command As SqlCommand = New SqlCommand(cons, connection)

            connection.Open()

            Dim reader As SqlDataReader = command.ExecuteReader()
            If reader.HasRows Then
                verificasaldosp = True
            End If
            reader.Close()

            command = Nothing
            reader = Nothing
        End Using
    End Function
    Function traeultimoidofa() As Integer
        Dim cons As String = "SELECT TOP (1) Id_Ofa FROM Orden WITH (nolock) ORDER BY Id_Ofa DESC"
        Using connection As New SqlConnection(Str_Con_Bd)
            Dim command As SqlCommand = New SqlCommand(cons, connection)

            connection.Open()

            Dim reader As SqlDataReader = command.ExecuteReader()

            If reader.HasRows Then
                reader.Read()
                Return reader.GetValue(0)
            End If
            command = Nothing
            reader.Close()
            reader = Nothing
        End Using
    End Function
    Private Sub datosorden(ByRef t_sol As String, ByRef CSG As String, ByRef id_fup As String, ByRef numof As String, ByRef anof As String, ByVal ofa As String, ByRef ordenseg As String, ByRef planta As String)
        Dim cons As String = "SELECT Orden.Tipo_Sol, Orden.Codigo_Sgc, Orden.Yale_Cotiza, Orden.Numero, Orden.ano, Orden.ordenseg_id, Orden_Seg.planta_id " +
                        "FROM Orden WITH (nolock) INNER JOIN Orden_Seg WITH (nolock) ON Orden.ordenseg_id = Orden_Seg.Id_Seg_Of WHERE  (Orden.Ofa = '" & ofa & "')"
        Using connection As New SqlConnection(Str_Con_Bd)
            Dim command As SqlCommand = New SqlCommand(cons, connection)

            connection.Open()

            Dim reader As SqlDataReader = command.ExecuteReader()

            If reader.HasRows Then
                reader.Read()
                t_sol = reader.GetValue(0)
                CSG = reader.GetValue(1)
                id_fup = reader.GetValue(2)
                numof = reader.GetValue(3)
                anof = reader.GetValue(4)
                ordenseg = reader.GetValue(5)
                planta = reader.GetValue(6)
            End If
            command = Nothing
            reader.Close()
            reader = Nothing
        End Using
    End Sub
    Public Sub crearaya(Optional acc As Boolean = False, Optional rayacrear As Integer = 0)
        Dim idofa, ofa, idofap, t_sol, CSG, id_fup, numof, anof, ordenseg, planta As String
        'Dim fecha As Double
        Dim fecha As String = "SYSDATETIME()"
        'fecha = DateAndTime.Now.ToOADate

        If acc = True Then

            ofa = ComboBox6.Text & "-1" 'define la raya 1 para buscar el idofap

        Else

            ofa = ComboBox3.Text & "-1" 'define la raya 1 para buscar el idofap

        End If

        idofap = capturaidofap(ofa)
        If idofap = Nothing Then
            If acc = True Then
                ofa = ComboBox6.Text
            Else
                ofa = ComboBox3.Text
            End If

            idofap = capturaidofap(ofa)
        End If

        datosorden(t_sol, CSG, id_fup, numof, anof, ofa, ordenseg, planta)

        If acc = True Then

            unidadNegocio = 3
            Dim cons As String

            If rayaimparvacia = ComboBox5.SelectedItem Then

                If rayaimparvacia > 0 And ComboBox5.SelectedItem <> Nothing Then

                    'Dim row As DataRow
                    ofa = ComboBox6.Text & "-" & ComboBox5.Text
                    Dim ofaexiste = True 'False 
                    Dim consofa = "SELECT Id_Ofa FROM Orden WITH (nolock) WHERE (Ofa = '" & ofa & "')"

                    If ComboBox7.SelectedItem = "CT" Then
                        consofa = consofa & " AND (No_Mod = " & ComboBox10.SelectedItem & ")"
                    End If

                    Using connection As New SqlConnection(Str_Con_Bd)
                        Dim command As SqlCommand = New SqlCommand(consofa, connection)
                        connection.Open()
                        Dim reader As SqlDataReader = command.ExecuteReader()
                        If reader.HasRows Then
                            'reader.Read()
                            ofaexiste = True
                        Else
                            ofaexiste = False
                        End If
                        command = Nothing
                        reader.Close()
                        reader = Nothing
                    End Using

                    'row = Conn.SelReg(consofa)

                    'If row IsNot Nothing Then

                    '    ofaexiste = True

                    'End If

                    If ofaexiste = False Then

                        idofa = traeultimoidofa() + 1

                        cons = "INSERT INTO Orden " +
                                                   "(Id_Ofa, Id_Of_P, Id_Emp_Ing, Ofa, Tipo_Of, Tipo_Sol, Codigo_Sgc, Yale_Cotiza, Numero, ano, letra, fecha_ingenieria, m2, M2_Reales, Nu_Piezas, Nu_Real_Und, Anulada, Despachada, Asignada, Ref, Abierta, Abierta_Acc, F_Aper, ord_fecha_real, ord_fecha_asigna, ord_fecha_planeada, Bloqueada, Ord_Unidad_Neg, Accesorios, ordenseg_id, planta_id, No_Mod)"
                        cons = cons & " VALUES (" & idofa & ", '" & idofap & "', 0, '" & ofa & "', '" & ComboBox7.SelectedItem & "', '" & t_sol & "', '" & CSG & "', " & id_fup & ", '" & numof & "', " & anof & ", " & rayaimparvacia & ", " & fecha & ", 0, 0, 0, 0, 0, 0, 1, '" & TextBox3.Text & "', 1, 1, " & fecha & ", " & fecha & ", " & fecha & ", " & fecha & ", 0, " & unidadNegocio & ", 1, " & ordenseg & ", " & planta & ", " & ComboBox10.SelectedItem & ")"

                        Try

                            Conn.ejecutarSql(cons)

                        Catch ex As Exception

                            MsgBox("Ocurrio un Error, Intente Mas Tarde")
                            Conn.RegistraExcepcion("SYS", "BdDatos.CargarTabla", "Error de Consulta", cons)

                        End Try

                    End If

                End If

            End If

            If rayaparvacia = ComboBox8.SelectedItem Then

                If rayaparvacia > 0 And ComboBox8.SelectedItem <> Nothing Then

                    Dim row As DataRow

                    ofa = ComboBox6.Text & "-" & ComboBox8.Text
                    Dim ofaexiste = False 'True
                    Dim consofa = "SELECT Id_Ofa FROM Orden WITH (nolock) WHERE (Ofa = '" & ofa & "')"

                    If ComboBox7.SelectedItem = "CT" Then

                        consofa = consofa & " AND (No_Mod = " & ComboBox10.SelectedItem & ")"

                    End If

                    'Valida que la Orden de Fabriacion Existe
                    row = Conn.SelReg(consofa)

                    If row IsNot Nothing Then

                        ofaexiste = True

                    End If

                    If ofaexiste = False Then

                        idofa = traeultimoidofa() + 1

                        cons = "INSERT INTO Orden " +
                                                   "(Id_Ofa, Id_Of_P, Id_Emp_Ing, Ofa, Tipo_Of, Tipo_Sol, Codigo_Sgc, Yale_Cotiza, Numero, ano, letra, fecha_ingenieria, m2, M2_Reales, Nu_Piezas, Nu_Real_Und, Anulada, Despachada, Asignada, Ref, Abierta, Abierta_Acc, F_Aper, ord_fecha_real, ord_fecha_asigna, ord_fecha_planeada, Bloqueada, Ord_Unidad_Neg, Accesorios, ordenseg_id, planta_id, No_Mod)"
                        cons = cons & " VALUES (" & idofa & ", '" & idofap & "', 0, '" & ofa & "', '" & ComboBox7.SelectedItem & "', '" & t_sol & "', '" & CSG & "', " & id_fup & ", '" & numof & "', " & anof & ", " & rayaparvacia & ", " & fecha & ", 0, 0, 0, 0, 0, 0, 1, '" & TextBox2.Text & "', 1, 1, " & fecha & ", " & fecha & ", " & fecha & ", " & fecha & ", 0, " & unidadNegocio & ", 1, " & ordenseg & ", " & planta & ", " & ComboBox10.SelectedItem & ")"

                        Try

                            Conn.ejecutarSql(cons)

                        Catch ex As Exception

                            MsgBox("Ocurrio un Error, Intente Nuevamente")
                            Conn.RegistraExcepcion("SYS", "BdDatos.CargarTabla", "Error de Consulta", cons)

                        End Try

                    End If

                End If

            End If

        Else

            Dim consUN = "SELECT Und_Neg_Id FROM Unidad_Negocio WHERE  (Und_Neg_Nombre = '" & ComboBox11.Text & "')"

            Using connection As New SqlConnection(Str_Con_Bd)
                Dim command As SqlCommand = New SqlCommand(consUN, connection)

                connection.Open()

                Dim reader As SqlDataReader = command.ExecuteReader()

                If reader.HasRows Then
                    Do While reader.Read()
                        unidadNegocio = reader.GetValue(0)
                    Loop
                End If
                command = Nothing
                reader.Close()
                reader = Nothing

            End Using

            If rayacrear <> 0 Then
                ofa = ComboBox3.Text & "-" & rayacrear
                rayavacia = rayacrear
            Else
                ofa = ComboBox3.Text & "-" & ComboBox4.Text
            End If


            Dim ofaexiste = True
            Dim consofa = "SELECT Id_Ofa FROM Orden WITH (nolock) WHERE (Ofa = '" & ofa & "')"
            If ComboBox2.SelectedItem = "CT" Then
                consofa = consofa & " AND (No_Mod = " & ComboBox9.SelectedItem & ")"
            End If
            Using connection As New SqlConnection(Str_Con_Bd)
                Dim command As SqlCommand = New SqlCommand(consofa, connection)
                connection.Open()
                Dim reader As SqlDataReader = command.ExecuteReader()
                If reader.HasRows Then
                    'reader.Read()
                    ofaexiste = True
                Else
                    ofaexiste = False
                End If
                command = Nothing
                reader.Close()
                reader = Nothing
            End Using
            If ofaexiste = False Then

                idofa = traeultimoidofa() + 1
                Dim cons As String

                cons = "INSERT INTO Orden " +
                                               "(Id_Ofa, Id_Of_P, Id_Emp_Ing, Ofa, Tipo_Of, Tipo_Sol, Codigo_Sgc, Yale_Cotiza, Numero, ano, letra, fecha_ingenieria, m2, M2_Reales, Nu_Piezas, Nu_Real_Und, Anulada, Despachada, Asignada, Ref, Abierta, F_Aper, ord_fecha_real, ord_fecha_asigna, ord_fecha_planeada, Bloqueada, Ord_Unidad_Neg, ordenseg_id, planta_id, No_Mod)"
                cons = cons & " VALUES (" & idofa & ", '" & idofap & "', 0, '" & ofa & "', '" & ComboBox2.SelectedItem & "', '" & t_sol & "', '" & CSG & "', " & id_fup & ", '" & numof & "', " & anof & ", " & rayavacia & ", " & fecha & ", 0, 0, 0, 0, 0, 0, 1, '" & TextBox1.Text & "', 1, " & fecha & ", " & fecha & ", " & fecha & ", " & fecha & ", 0, " & unidadNegocio & ", " & ordenseg & ", " & planta & ", " & ComboBox9.SelectedItem & ")"

                Using connection As New SqlConnection(Str_Con_Bd)
                    Dim command As SqlCommand = New SqlCommand(cons, connection)
                    connection.Open()
                    command.ExecuteNonQuery()
                End Using
            End If
        End If

    End Sub
    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        If ComboBox11.Text = "Unidad Negocio" Or ComboBox11.Text = "Ninguna" Then
            MsgBox("Debe elegir unidad de negocio")
        Else
            cargaraya()
        End If

    End Sub

    Public Sub cargaraya()

        Dim idofa, ofa, idofap As String

        ofa = ComboBox3.Text & "-" & ComboBox4.Text

        ''New
        'If ComboBox2.Text = "OG" Then
        '    ofa = ComboBox3.Text
        'End If

        idofa = capturaidofa(ofa, ComboBox2.SelectedItem, ComboBox9.SelectedItem)
        If idofa = 0 Then
            ofa = ComboBox3.Text
            idofa = capturaidofa(ofa, ComboBox2.SelectedItem, ComboBox9.SelectedItem)
        End If
        If semaforo <> "vacio" Then
            Dim tienesaldosp = verificasaldosp(idofa)
            If tienesaldosp = True Then
                MsgBox("no se puede cargar la raya porque tiene etiquetas generadas en producción")
                Exit Sub
            End If
        End If
        If impidereemplazar = True Then
            MsgBox("no se puede cargar la raya porque tiene elementos cargados")
            Exit Sub

        Else
            If semaforo = "vacio" Then

                crearaya()

            Else
                Dim consUN = "SELECT Und_Neg_Id FROM Unidad_Negocio WHERE  (Und_Neg_Nombre = '" & ComboBox11.Text & "')"

                Using connection As New SqlConnection(Str_Con_Bd)
                    Dim command As SqlCommand = New SqlCommand(consUN, connection)

                    connection.Open()

                    Dim reader As SqlDataReader = command.ExecuteReader()

                    If reader.HasRows Then
                        Do While reader.Read()
                            unidadNegocio = reader.GetValue(0)
                        Loop
                    End If
                    command = Nothing
                    reader.Close()
                    reader = Nothing
                End Using
                consUN = "UPDATE Orden SET Ord_Unidad_Neg = " & unidadNegocio & " " +
                                "WHERE (Id_Ofa = " & idofa & ")"
                Using connection As New SqlConnection(Str_Con_Bd)
                    Dim command As SqlCommand = New SqlCommand(consUN, connection)
                    connection.Open()
                    command.ExecuteNonQuery()
                End Using
            End If
            ofa = ComboBox3.Text & "-" & ComboBox4.Text
            idofa = capturaidofa(ofa, ComboBox2.SelectedItem, ComboBox9.SelectedItem)
            If idofa = 0 Then
                ofa = ComboBox3.Text
                idofa = capturaidofa(ofa, ComboBox2.SelectedItem, ComboBox9.SelectedItem)
            End If
            Dim cons As String = "DELETE FROM Saldos WHERE (Id_Ofa = '" & idofa & "')"
            Using connection As New SqlConnection(Str_Con_Bd)
                Dim command As SqlCommand = New SqlCommand(cons, connection)
                command.CommandTimeout = 600
                connection.Open()
                command.ExecuteNonQuery()
            End Using
        End If
        grills = Nothing
        ReDim grills(3)
        grills(0) = DataGridView1
        grills(1) = DataGridView2
        grills(2) = DataGridView3
        grills(3) = DataGridView5

        If ComboBox2.SelectedText = "OG" Or ComboBox2.SelectedText = "OM" Then 'rc y pv no
            idofap = idofa
        Else
            idofap = capturaidofap(ofa)
        End If

        Dim ParaArmado = 0, Pernado = 0, UltimaEntrega = 0, Escalera = 0
        If CheckBox7.Checked = True Then
            ParaArmado = 1
        End If
        If CheckBox8.Checked = True Then
            Pernado = 1
        End If
        If CheckBox9.Checked = True Then
            UltimaEntrega = 1
        End If
        If CheckBox10.Checked = True Then
            Escalera = 1
        End If

        CreaRQI.MyCommands.cargaorden(idofa, grills, ofa, idofap, DataGridView6, TextBox1.Text, TextBox8.Text, loginvalidez, darusuario, "", darpassw, ParaArmado, Pernado, UltimaEntrega, Escalera)

        If mensajes = True Then
            'MsgBox("cargado con exito")
        End If
        llenacombo4()
        'Dim cerrarventana = MsgBox("cargado con éxito ¿desea cerrar?", MsgBoxStyle.YesNo)
        'If cerrarventana = MsgBoxResult.Yes Then
        'Me.Close()
        'End If

    End Sub
    Function capturaidofa(ofa As String, tipoof As String, consecutivo As String) As String
        Dim Cons As String = "SELECT Id_Ofa FROM Orden WITH (nolock) WHERE (Ofa = '" & ofa & "')"
        'Dim Cons As String = "SELECT TOP(1) Id_Ofa FROM Orden WITH (nolock) WHERE (Ofa = '" & ofa & "') AND Anulada = 0 "

        If tipoof = "CT" Then

            Cons = Cons & " AND (No_Mod = " & consecutivo & ")"

        End If
        Using connection As New SqlConnection(Str_Con_Bd)

            Dim command As SqlCommand = New SqlCommand(Cons, connection)

            connection.Open()

            Dim reader As SqlDataReader = command.ExecuteReader()

            If reader.HasRows Then
                reader.Read()
                Return reader.GetValue(0)
            End If
            command = Nothing
            reader.Close()
            reader = Nothing

        End Using
    End Function
    Function capturaidofap(ofa As String) As String
        Dim Cons As String = "SELECT Id_Of_P FROM Orden WITH (nolock) WHERE (Ofa = '" & ofa & "')"
        Using connection As New SqlConnection(Str_Con_Bd)
            Dim command As SqlCommand = New SqlCommand(Cons, connection)

            connection.Open()

            Dim reader As SqlDataReader = command.ExecuteReader()

            If reader.HasRows Then
                reader.Read()
                Return reader.GetValue(0)
            End If
            command = Nothing
            reader.Close()
            reader = Nothing
        End Using
    End Function
    Private Sub verifrayalibre()
        Dim idofa, ofa As String
        ofa = ComboBox3.Text & "-" & ComboBox4.Text
        idofa = capturaidofa(ofa, ComboBox2.SelectedItem, ComboBox9.SelectedItem)
        If idofa = 0 Then
            ofa = ComboBox3.Text
            idofa = capturaidofa(ofa, ComboBox2.SelectedItem, ComboBox9.SelectedItem)
        End If
        Dim cons As String = "SELECT SUM(Cant_Final_Req) AS cant FROM Saldos WITH (nolock) WHERE (Id_Ofa = " & idofa & ")"
        Using connection As New SqlConnection(Str_Con_Bd)
            Dim command As SqlCommand = New SqlCommand(cons, connection)

            connection.Open()
            Dim reader As SqlDataReader
            Try
                reader = command.ExecuteReader()

                If reader.HasRows Then
                    reader.Read()

                    If reader.GetValue(0) > 0 Then
                        rayalibre = False
                    Else
                        rayalibre = True
                    End If
                Else
                    rayalibre = True
                End If
            Catch ex As Exception
                rayalibre = True
                Exit Sub
            End Try
            command = Nothing
            reader.Close()
            reader = Nothing

        End Using
    End Sub
    Private Sub verificaestadoerp()
        Dim idofa, ofa, plantaid As String
        Dim itemabuelo As Integer
        estatuserplibre = True

        plantaid = 1
        ofa = ComboBox3.Text & "-" & ComboBox4.Text
        idofa = capturaidofa(ofa, ComboBox2.SelectedItem, ComboBox9.SelectedItem)
        If idofa = 0 Then
            ofa = ComboBox3.Text
            idofa = capturaidofa(ofa, ComboBox2.SelectedItem, ComboBox9.SelectedItem)
        End If

        Dim cons As String = "SELECT Orden_Seg.OP_1EE_Abuelo, Orden.planta_id, Orden.Op_Sol_Apro_Fecha FROM Orden WITH (nolock) INNER JOIN " +
                         "Orden_Seg WITH (nolock) ON Orden.ordenseg_id = Orden_Seg.Id_Seg_Of " +
                        "WHERE (Orden.Id_Ofa = " & idofa & ")"

        Using Connection As New SqlConnection(Str_Con_Bd)
            Dim command As SqlCommand = New SqlCommand(cons, Connection)

            Connection.Open()
            Dim reader As SqlDataReader
            Try
                reader = command.ExecuteReader()
            Catch ex As Exception
                estatuserplibre = True
                Exit Sub
            End Try

            If reader.HasRows Then
                reader.Read()
                itemabuelo = reader.GetValue(0)
                plantaid = reader.GetValue(1)
                Try
                    If reader.GetValue(2) = Nothing Then 'si no es nulo, entonces tiene fecha y ya no se puede usar
                    Else
                        estatuserplibre = False
                    End If
                Catch ex As Exception

                End Try

            End If
            command = Nothing
            reader.Close()
            reader = Nothing

        End Using
        'Dim rsiferp2 = New ADODB.Recordset
        If estatuserplibre = False Then 'si ya no se puede usar, no revise mas
        ElseIf plantaid = 3 Then 'si es Brasil no busque en el ERP porque Brasil es Gerbo
        Else
            Dim Slrs10 = "SELECT T850_MF_OP_DOCTO.F850_ID_CLASE_OP, T850_MF_OP_DOCTO.F850_IND_ESTADO " +
                    "From T850_MF_OP_DOCTO " +
                    "Where T850_MF_OP_DOCTO.F850_ID_CIA = 6 " +
                    "And	T850_MF_OP_DOCTO.F850_ID_TIPO_DOCTO = 'IF' " +
                    "And T850_MF_OP_DOCTO.F850_ID_INSTALACION = '001' " +
                    "And T850_MF_OP_DOCTO.F850_CONSEC_DOCTO = " & itemabuelo
            'Dim Con_Ora As ADODB.Connection = CreaRQI.MyCommands.AbreDb_Orl
            Dim Str_Con_BdERP = CreaRQI.MyCommands.AbreDB_ERP
            Using connection As New SqlConnection(Str_Con_BdERP)
                Dim command As SqlCommand = New SqlCommand(Slrs10, connection)
                connection.Open()

                Dim reader As SqlDataReader = command.ExecuteReader()
                If reader.HasRows Then
                    reader.Read()
                    Try
                        If reader.GetValue(1) = 3 Then
                            estatuserplibre = False
                        Else
                            estatuserplibre = True
                        End If
                    Catch ex As Exception
                        estatuserplibre = True
                    End Try

                Else
                    estatuserplibre = True
                End If
                command = Nothing
                reader.Close()
                reader = Nothing
            End Using
            'rsiferp2.Open(Slrs10, Con_Ora, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)

        End If

    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        llenacombo3(ComboBox2.SelectedItem)
        If ComboBox2.Text = "CT" Then

            ComboBox9.Text = 1
            Label13.Visible = True
            ComboBox9.Visible = True
            ComboBox3.Enabled = False
            'New
            ComboBox9.Text = ""
            '----------------------------
            ComboBox9.Items.Remove("0")
        Else
            'Añadir al combo "consecutivo" numeracion del 0 al 10
            'Cuando es diferente de CT va de 0 al 10
            If (ComboBox9.Items.Count = 9) Then

                ComboBox9.Items.Insert(0, "0")

            End If
            ComboBox9.Text = 0
            Label13.Visible = False
            ComboBox9.Visible = False
        End If

        'Este tipo de Orden no llevan Rayas. El campo Letra en la tabla Orden de la BD va en NULL.
        'El CargaSIIF permite hacer el cargue sin necesidad de seleccionar la raya.
        'Para evitar errores de usuario lo mejor es bloquear el combo "Raya".
        If ComboBox2.Text = "SR" Or ComboBox2.Text = "ID" Or ComboBox2.Text = "OG" Then

            ComboBox4.Enabled = False
        Else

            ComboBox4.Enabled = True

        End If

    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        llenacombo4()
        llenacombo11()
        If ComboBox2.Text = "CT" Then
            ComboBox3.Enabled = False
        End If

        'Solo llena la pestaña de Aluminio
        cargarinfoCliente(ComboBox3.Text, 1)

    End Sub
    Private Sub referencia()
        If ComboBox4.SelectedItem = rayavacia Then
        Else
            Dim cons As String = "SELECT Ref FROM Orden WITH (nolock) WHERE (Ofa = '" & ComboBox3.Text & "-" & ComboBox4.Text & "')"
            If ComboBox9.SelectedItem = "CT" Then
                cons = cons & " AND (No_Mod = " & ComboBox9.SelectedItem & ")"
            End If
            Using connection As New SqlConnection(Str_Con_Bd)
                Dim command As SqlCommand = New SqlCommand(cons, connection)

                connection.Open()

                Dim reader As SqlDataReader = command.ExecuteReader()

                If reader.HasRows Then
                    reader.Read()
                    Try
                        TextBox1.Text = reader.GetValue(0)
                    Catch ex As Exception
                        TextBox1.Text = "Ref"
                    End Try

                End If
                command = Nothing
                reader.Close()
                reader = Nothing
            End Using
        End If
    End Sub

    Private Sub VerificarOrden()

        Dim idofa, ofa, plantaid As String
        Dim itemabuelo As Integer
        estatuserplibre = True 'Erp
        entregada = False 'Estatus Solicitud
        'New
        Dim objConn As New ConexionBD
        Dim regVal As DataRow

        plantaid = 1
        ofa = ComboBox3.Text & "-" & ComboBox4.Text
        idofa = capturaidofa(ofa, ComboBox2.SelectedItem, ComboBox9.SelectedItem)
        If idofa = 0 Then
            ofa = ComboBox3.Text
            idofa = capturaidofa(ofa, ComboBox2.SelectedItem, ComboBox9.SelectedItem)
        End If

        MessageBox.Show(idofa)
        Dim cons As String = "SELECT Ord_seg.OP_1EE_Abuelo, Ord.planta_id, Ord.Op_Sol_Apro_Fecha, " +
                             "CASE WHEN Ord_fec_impresion Is NULL THEN 0 ELSE 1 END AS Entrega " +
                             "From Orden as Ord WITH (NOLOCK) INNER JOIN " +
                             "Orden_Seg as Ord_seg WITH (NOLOCK) ON Ord.ordenseg_id = Ord_seg.Id_Seg_Of " +
                             "Where (Ord.Id_Ofa = " & idofa & ") "
        'New
        'El try Catch, sobra la verdad sin embargo no voy a pecar... y lo voy a dejar
        Try
            regVal = objConn.SelReg(cons)

            If regVal IsNot Nothing Then

                itemabuelo = regVal("OP_1EE_Abuelo")
                plantaid = regVal("planta_id")

                'VerificaEstadoERP 
                'si no es nulo, entonces tiene fecha y ya no se puede usar
                If regVal("Op_Sol_Apro_Fecha") IsNot Nothing Then
                    estatuserplibre = False
                End If
                'Verifica Estatus Solicitud
                If regVal("Entrega") = 1 Then
                    entregada = True
                End If

            End If

        Catch ex As Exception
            MessageBox.Show("VerificarOrden: \n" + ex.ToString())
        End Try

        'Esto ya estaba y como no conozco la relacion de esas tablas, lo dejo igual.
        '------------------------------------------------------------------------------
        If estatuserplibre = False Then 'si ya no se puede usar, no revise mas
        ElseIf plantaid = 3 Then 'si es Brasil no busque en el ERP porque Brasil es Gerbo
        Else
            Dim Slrs10 = "Select T850_MF_OP_DOCTO.F850_ID_CLASE_OP, T850_MF_OP_DOCTO.F850_IND_ESTADO " +
                    "From T850_MF_OP_DOCTO " +
                    "Where T850_MF_OP_DOCTO.F850_ID_CIA = 6 " +
                    "And	T850_MF_OP_DOCTO.F850_ID_TIPO_DOCTO = 'IF' " +
                    "And T850_MF_OP_DOCTO.F850_ID_INSTALACION = '001' " +
                    "And T850_MF_OP_DOCTO.F850_CONSEC_DOCTO = " & itemabuelo
            'Dim Con_Ora As ADODB.Connection = CreaRQI.MyCommands.AbreDb_Orl
            Dim Str_Con_BdERP = CreaRQI.MyCommands.AbreDB_ERP
            Using connection As New SqlConnection(Str_Con_BdERP)
                Dim command As SqlCommand = New SqlCommand(Slrs10, connection)
                connection.Open()

                Dim reader As SqlDataReader = command.ExecuteReader()
                If reader.HasRows Then
                    reader.Read()
                    Try
                        If reader.GetValue(1) = 3 Then
                            estatuserplibre = False
                        Else
                            estatuserplibre = True
                        End If
                    Catch ex As Exception
                        estatuserplibre = True
                    End Try

                Else
                    estatuserplibre = True
                End If
                command = Nothing
                reader.Close()
                reader = Nothing
            End Using
        End If

    End Sub
    'Tipo producto define si es Aluminio(1) Acero(2)
    Private Sub cargarinfoCliente(ByVal NumOrden As String, ByVal TipoProducto As Integer)

        'Asignamos el cliente de la orden.
        Dim msg As String
        'Asignamos la planta de login del usuario.
        If Me.PlantaId = 0 Then

            Me.PlantaId = 1

        End If

        Orden.PlantaLoginUsu = Me.PlantaId

        'Aluminio
        If TipoProducto = 1 Then
            If Orden.Validar(NumOrden, msg) Then

                Orden.RecuperarInfoCliente()

                lblPaisAlum.Text = "Pais: " + Orden.Pais
                lblClienteAlum.Text = "Cliente: " + Orden.Cliente
                lblObraAlum.Text = "Obra: " + Orden.Obra

            Else
                lblPaisAlum.Text = "Pais:"
                lblClienteAlum.Text = "Cliente:"
                lblObraAlum.Text = "Obra:"
            End If
        End If

        'Acc
        If TipoProducto = 2 Then
            If Orden.Validar(NumOrden, msg) Then

                Orden.RecuperarInfoCliente()

                lblPaisAcc.Text = "Pais: " + Orden.Pais
                lblClienteAcc.Text = "Cliente: " + Orden.Cliente
                lblObraAcc.Text = "Obra: " + Orden.Obra
            Else
                lblPaisAcc.Text = "Pais: " + Orden.Pais
                lblClienteAcc.Text = "Cliente: " + Orden.Cliente
                lblObraAcc.Text = "Obra: " + Orden.Obra
            End If
        End If
    End Sub

    Private Sub ComboBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox4.SelectedIndexChanged

        verificaestadoerp()
        Verif_Estatus_Sol()
        'new - reusme las 2 funciones comentadas
        'VerificarOrden()
        '-----------------
        verifrayalibre()
        referencia()

        Dim reemplazar As MsgBoxResult

        If ComboBox4.SelectedItem = rayavacia Then 'rayavacia es el numero de raya vacia que no existe
            semaforo = "vacio"
            rayalibre = True
            CheckBox1.Checked = False
            impidereemplazar = False
            Button12.Enabled = True
            Button12.BackColor = System.Drawing.Color.Transparent
            Button12.Text = "Cargar Orden"
            TextBox1.Text = "Ref"
        ElseIf estatuserplibre = True And entregada = False And rayalibre = True Then
            semaforo = "abierto"
            impidereemplazar = False
            CheckBox1.Checked = False
            Button12.Enabled = True
            Button12.BackColor = System.Drawing.Color.Transparent
            Button12.Text = "Cargar Orden"
        ElseIf estatuserplibre = True And entregada = False And rayalibre = False Then
            semaforo = "modificable"
            Button12.BackColor = System.Drawing.Color.Yellow
            Button12.Text = "Cargar Orden"
            reemplazar = MsgBox("la raya seleccionada tiene elementos cargados, desea reemplazar?", MsgBoxStyle.OkCancel, "reemplazar")
            If reemplazar = MsgBoxResult.Ok Then
                CheckBox1.Checked = True
                impidereemplazar = False
                Button12.Enabled = True
            Else
                impidereemplazar = True
                CheckBox1.Checked = False
                Button12.Enabled = False
                'ComboBox4.SelectedValue = ""
            End If
        ElseIf estatuserplibre = True And entregada = True Then
            semaforo = "entregado"
            impidereemplazar = True
            CheckBox1.Checked = False
            Button12.Enabled = False
            Button12.BackColor = System.Drawing.Color.Red
            Button12.Text = "ya entregada"
        ElseIf estatuserplibre = False Then
            semaforo = "cerrado"
            impidereemplazar = True
            CheckBox1.Checked = False
            Button12.Enabled = False
            Button12.BackColor = System.Drawing.Color.Red
            Button12.Text = "cerrada"
        End If
    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        ordenar()

    End Sub

    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
        CreaRQI.MyCommands.muevefila(grillas)

    End Sub

    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click
        'Dim url As String = "http://si.forsa.com.co:81/ReportServer/Pages/ReportViewer.aspx?/RepIng/rptEntregaListadoPiezas&rs:ClearSession=true&rs:command=render&Orden=" & ComboBox3.Text & "-" & ComboBox4.Text
        Dim url As String = "http://app.forsa.com.co/siomaestros/ReportViewer.aspx?/RepIng/rptEntregaListadoPiezas&Orden=" & ComboBox3.Text & "-" & ComboBox4.Text
        Process.Start(url)
    End Sub

    Sub recalcular(fila As DataGridViewRow)
        Dim an1, al1, al2, an2, areapieza, areaitem As Double
        Dim tipoitem, desaux, grupo As String
        tipoitem = fila.Cells(1).Value
        desaux = fila.Cells(7).Value
        grupo = fila.Cells(8).Value
        If tipoitem = "MAG" Or tipoitem = "AG" Or tipoitem = "AGR" Then
            fila.Cells(9).Value = 0.01
            fila.Cells(10).Value = 0.01
        Else
            an1 = fila.Cells(2).Value
            al1 = fila.Cells(3).Value
            al2 = fila.Cells(4).Value
            an2 = fila.Cells(5).Value
            areapieza = CreaRQI.MyCommands.AreaCalc(tipoitem, an1, al1, al2, an2, desaux, grupo) 'area item es un una unidad del item 'Math.Round((CDbl(an1) + CDbl(an2)) * (CDbl(al1) + CDbl(al2)) / 10000, 2)
            areaitem = CDbl(areapieza) * CInt(fila.Cells(0).Value)
            fila.Cells(9).Value = areapieza
            fila.Cells(10).Value = areaitem
        End If


    End Sub

    Private Sub Button21_Click(sender As Object, e As EventArgs) Handles Button21.Click
        seleclista()
    End Sub

    Private Sub Button18_Click(sender As Object, e As EventArgs) Handles Button18.Click
        Dim dim1, dim2, dim3, dim4, dim5, dim6, dim7 As String
        Dim dim1d, dim2d, dim3d, dim4d, dim5d, dim6d, dim7d As Double
        Dim plantaid As Integer
        Dim ParImpar As String
        Dim j = 0
        Dim grillasacc(1) As Object
        grillasacc(0) = DataGridView4
        grillasacc(1) = DataGridView7

        For Each grillaacc As DataGridView In grillasacc
            If j = 0 Then
                ParImpar = "Almacen"
                j = j + 1
            Else
                ParImpar = "Planta"
            End If

            For Each fila As DataGridViewRow In grillaacc.Rows

                If fila.Cells(10).Value = 0 Then

                    Dim nom = fila.Cells(1).Value

                    If nom = "" Or nom = Nothing Then
                        Continue For
                    End If

                    Dim desaux = fila.Cells(9).Value

                    dim1 = fila.Cells(2).Value
                    If dim1 = "" Then dim1d = 0 Else dim1d = dim1
                    dim1 = Math.Round(dim1d, 1)
                    dim2 = fila.Cells(3).Value
                    If dim2 = "" Then dim2d = 0 Else dim2d = dim2
                    dim2 = Math.Round(dim2d, 1)
                    dim3 = fila.Cells(4).Value
                    If dim3 = "" Then dim3d = 0 Else dim3d = dim3
                    dim3 = Math.Round(dim3d, 1)
                    dim4 = fila.Cells(5).Value
                    If dim4 = "" Then dim4d = 0 Else dim4d = dim4
                    dim4 = Math.Round(dim4d, 1)
                    dim5 = fila.Cells(6).Value
                    If dim5 = "" Then dim5d = 0 Else dim5d = dim5
                    dim5 = Math.Round(dim5d, 1)
                    dim6 = fila.Cells(7).Value
                    If dim6 = "" Then dim6d = 0 Else dim6d = dim6
                    dim6 = Math.Round(dim6d, 1)
                    dim7 = 0

                    If Label11.Text = "Colombia" Then
                        plantaid = 1
                    ElseIf Label11.Text = "Mexico" Then
                        plantaid = 2
                    ElseIf Label11.Text = "Brasil" Then
                        plantaid = 3
                    End If

                    Dim cons, cons2 As String
                    Dim nomlargo As String

                    cons = "SELECT NomLargoEsp FROM Accesorios_Maestro_Nomenclaturas WITH (nolock) WHERE (Nomenclatura = '" & nom & "')"

                    Using Connection As New SqlConnection(Str_Con_Bd)

                        Dim command As SqlCommand = New SqlCommand(cons, Connection)

                        Connection.Open()
                        Dim reader As SqlDataReader
                        reader = command.ExecuteReader()

                        If reader.HasRows Then
                            reader.Read()
                            nomlargo = reader.GetValue(0)
                        Else
                            nomlargo = InputBox("Indique que es un " & nom, "defina nombre largo para el ACC")
                            If nomlargo = "" Then
                                MsgBox(nom & " nomenclatura no existe en maestro y no se puede generar codigo")
                                Continue For
                            Else
                                nomlargo = nomlargo.ToUpper
                                Dim level As Integer
                                If Integer.TryParse(InputBox("indique nivel del 2 al 5 para " & nom, "defina a que nivel pertenece el ACC"), level) = True Then
                                    If level = 2 Or level = 3 Or level = 4 Or level = 5 Then
                                    Else
                                        MsgBox("Se le indicó que del 2 al 5, se deja nivel 3 por defecto, cualquier cambio comuníquese con costos")
                                        level = 3
                                    End If
                                Else
                                    level = 3
                                End If
                                cons2 = "INSERT INTO Accesorios_Maestro_Nomenclaturas (Nomenclatura,NomLargoEsp,Nivel,Soldadura,ProductoTerminado,Pintura,TiempoPor) " +
                                "VALUES ('" & nom & "','" & nomlargo & "'," & level & ",0.02,0.001,0.001,'Promedio')"
                            End If

                        End If

                        command = Nothing
                        reader.Close()
                        reader = Nothing
                        cons = ""

                    End Using

                    If cons2 <> "" Then

                        Using Connection As New SqlConnection(Str_Con_Bd)
                            Dim command As SqlCommand = New SqlCommand(cons2, Connection)
                            Connection.Open()
                            command.ExecuteNonQuery()
                            cons2 = ""
                        End Using

                    End If

                    cons = "SELECT IdItemGrupoPlanta FROM Accesorios_Cod_GrupoPlanta WHERE (Nomenclatura = '" & nom & "') AND (IdPlanta = 3)"

                    Using Connection As New SqlConnection(Str_Con_Bd)

                        Dim command As SqlCommand = New SqlCommand(cons, Connection)

                        Connection.Open()
                        Dim reader As SqlDataReader
                        reader = command.ExecuteReader()

                        If reader.HasRows Then

                        Else

                            Dim creagrupoplanta = New Frm_GrupoPlanta
                            creagrupoplanta.nom = nom
                            creagrupoplanta.ShowDialog()

                        End If

                        command = Nothing
                        reader.Close()
                        reader = Nothing
                        cons = ""

                    End Using

                    cons = "INSERT INTO Accesorios_Codigos " +
                                                        "(Id_UnoE, Nomenclatura, Des_Aux, Valor_1_Min, Valor_1_Max, Valor_2_Min, Valor_2_Max, Valor_3_Min, Valor_3_Max, Valor_4_Min, Valor_4_Max, Valor_5_Min, Valor_5_Max, Valor_6_Min, Valor_6_Max, Valor_7_Min, Valor_7_Max, planta_id, Acc_Anulado, Acc_Id_ItemPlanta, NombreLargo, ParImpar, Origen)"
                    cons = cons & " VALUES (-1, '" & nom & "', '" & desaux & "', '" & dim1 & "', '" & dim1 & "', '" & dim2 & "', '" & dim2 & "', " & dim3 & ", " & dim3 & ", " & dim4 & ", " & dim4 & ", " & dim5 & ", " & dim5 & ", " & dim6 & ", " & dim6 & ", " & dim7 & ", " & dim7 & ", " & plantaid & ", 0, -1,'" & nomlargo & " " & desaux & " " & dim1 & "x" & dim2 & "x" & dim3 & "x" & dim4 & "x" & dim5 & "x" & dim6 & "','" & ParImpar & "','CARGASIF')"

                    Using connection As New SqlConnection(Str_Con_Bd)
                        Dim command As SqlCommand = New SqlCommand(cons, connection)
                        connection.Open()
                        command.ExecuteNonQuery()
                    End Using

                    dim1 = Nothing
                    dim2 = Nothing
                    dim3 = Nothing
                    dim4 = Nothing
                    dim5 = Nothing
                    dim6 = Nothing
                    dim7 = Nothing

                End If
            Next

        Next

        ''Se crea el Objeto que realizará la solicitud REST
        'Dim webClient As New System.Net.WebClient
        'Try
        '    'Se convierte el dictionary que representa los datos en un String
        '    Dim strJSON As String = JsonConvert.SerializeObject(formatting:=Formatting.None)
        'Catch ex As Exception

        'End Try

        'Web Services
        Dim addres As String = "http://172.21.224.131/ApiTotvs/api/CrearAccesorio/cons=4555&planta=3" '"http://172.21.224.132/siomaestros/wsconsumegerbo.aspx?ItemPlantilla=1"
        Process.Start(addres)
        Dim seconds As Integer = 20
        For i As Integer = 0 To seconds * 100
            System.Threading.Thread.Sleep(10)
            System.Windows.Forms.Application.DoEvents()
        Next
        LlenaGrillas()
    End Sub

    Private Sub Button20_Click(sender As Object, e As EventArgs) Handles Button20.Click
        Dim habilitarcarga As Boolean = True

        Dim grillasacc(1) As Object
        grillasacc(0) = DataGridView4
        grillasacc(1) = DataGridView7
        Dim filas As System.Windows.Forms.DataGridViewRowCollection
        Dim fila As System.Windows.Forms.DataGridViewRow
        Dim j As Integer = 1
        For k = 0 To UBound(grillasacc)
            filas = grillasacc(k).rows

            For i = 0 To filas.Count - j
                If i = filas.Count Then 'verifica que esta en la ultima fila
                    Exit For
                End If
                fila = filas.Item(i)

                If fila.Selected = True Then
                    filas.Remove(fila)
                    j = j + 1 'suma uno para reducir el numero total de filas en el for
                    i = i - 1 'resta una porque al eliminar la fila se reduce la cantidad de filas total
                Else
                    If fila.Cells(0).Value > 0 And fila.Cells(10).Value = 0 Then
                        habilitarcarga = False
                    End If
                End If
            Next
        Next
        If habilitarcarga = True Then
            Button19.Enabled = True
            CheckBox3.Enabled = True
        End If
    End Sub

    Private Sub Button17_Click(sender As Object, e As EventArgs) Handles Button17.Click
        Dim cambiopermanente As DialogResult
        cambiopermanente = MsgBox("desea que el cambio sea permanente?", MsgBoxStyle.YesNo, "permanente o solo una vez")
        Dim grillasacc(1) As Object
        Dim cant, descri, desaux, dim1, dim2, dim3, dim4, dim5, dim6, plano, codigo, nomlargo, cantvalida As String
        'Dim AccesoriosCodigosId
        grillasacc(0) = DataGridView4 'almacen
        grillasacc(1) = DataGridView7 'planta
        Dim filastrae, filaslleva As System.Windows.Forms.DataGridViewRowCollection
        Dim filatrae, filalleva As System.Windows.Forms.DataGridViewRow
        Dim j As Integer = 1

        filastrae = grillasacc(0).rows
        filaslleva = grillasacc(1).rows
        For i = 0 To filastrae.Count - j
            If i = filastrae.Count Then 'verifica que esta en la ultima fila
                Exit For
            End If
            filatrae = filastrae.Item(i)
            If filatrae.Selected = True Then

                cant = filatrae.Cells(0).Value
                descri = filatrae.Cells(1).Value
                desaux = filatrae.Cells(9).Value
                dim1 = filatrae.Cells(2).Value
                dim2 = filatrae.Cells(3).Value
                dim3 = filatrae.Cells(4).Value
                dim4 = filatrae.Cells(5).Value
                dim5 = filatrae.Cells(6).Value
                dim6 = filatrae.Cells(7).Value
                plano = filatrae.Cells(8).Value
                codigo = filatrae.Cells(10).Value
                nomlargo = filatrae.Cells(11).Value
                cantvalida = filatrae.Cells(12).Value
                'AccesoriosCodigosId = filatrae.Cells(15).Value

                Try

                    Dim nuevafila = filaslleva.Add()

                    filalleva = filaslleva.Item(nuevafila)
                    filalleva.Cells(0).Value = cant
                    filalleva.Cells(1).Value = descri
                    filalleva.Cells(9).Value = desaux
                    filalleva.Cells(2).Value = dim1
                    filalleva.Cells(3).Value = dim2
                    filalleva.Cells(4).Value = dim3
                    filalleva.Cells(5).Value = dim4
                    filalleva.Cells(6).Value = dim5
                    filalleva.Cells(7).Value = dim6
                    filalleva.Cells(8).Value = plano
                    filalleva.Cells(10).Value = codigo
                    filalleva.Cells(11).Value = nomlargo
                    filalleva.Cells(12).Value = cantvalida
                    'filalleva.Cells(15).Value = AccesoriosCodigosId
                    If cambiopermanente = DialogResult.Yes Then
                        Dim cons = "update Accesorios_Codigos set ParImpar = 'PLANTA' where (Id_UnoE = " & codigo & ")"
                        Using connection As New SqlConnection(Str_Con_Bd)
                            Dim command As SqlCommand = New SqlCommand(cons, connection)
                            connection.Open()
                            command.ExecuteNonQuery()
                        End Using
                    End If
                Catch ex As Exception

                End Try

                filastrae.Remove(filatrae)
                j = j + 1 'suma uno para reducir el numero total de filas en el for
                i = i - 1 'resta una porque al eliminar la fila se reduce la cantidad de filas total
            End If
        Next
        j = 1
        filastrae = grillasacc(1).rows
        filaslleva = grillasacc(0).rows
        For i = 0 To filastrae.Count - j
            If i = filastrae.Count Then 'verifica que esta en la ultima fila
                Exit For
            End If
            filatrae = filastrae.Item(i)
            If filatrae.Selected = True Then

                cant = filatrae.Cells(0).Value
                descri = filatrae.Cells(1).Value
                desaux = filatrae.Cells(9).Value
                dim1 = filatrae.Cells(2).Value
                dim2 = filatrae.Cells(3).Value
                dim3 = filatrae.Cells(4).Value
                dim4 = filatrae.Cells(5).Value
                dim5 = filatrae.Cells(6).Value
                dim6 = filatrae.Cells(7).Value
                plano = filatrae.Cells(8).Value
                codigo = filatrae.Cells(10).Value
                nomlargo = filatrae.Cells(11).Value
                cantvalida = filatrae.Cells(12).Value
                'AccesoriosCodigosId = filatrae.Cells(15).Value

                Try

                    Dim nuevafila = filaslleva.Add()

                    filalleva = filaslleva.Item(nuevafila)
                    filalleva.Cells(0).Value = cant
                    filalleva.Cells(1).Value = descri
                    filalleva.Cells(9).Value = desaux
                    filalleva.Cells(2).Value = dim1
                    filalleva.Cells(3).Value = dim2
                    filalleva.Cells(4).Value = dim3
                    filalleva.Cells(5).Value = dim4
                    filalleva.Cells(6).Value = dim5
                    filalleva.Cells(7).Value = dim6
                    filalleva.Cells(8).Value = plano
                    filalleva.Cells(10).Value = codigo
                    filalleva.Cells(11).Value = nomlargo
                    filalleva.Cells(12).Value = cantvalida
                    'filalleva.Cells(15).Value = AccesoriosCodigosId

                    If cambiopermanente = DialogResult.Yes Then
                        Dim cons = "update Accesorios_Codigos set ParImpar = 'ALMACEN' where (Id_UnoE = " & codigo & ")"
                        Using connection As New SqlConnection(Str_Con_Bd)
                            Dim command As SqlCommand = New SqlCommand(cons, connection)
                            connection.Open()
                            command.ExecuteNonQuery()
                        End Using
                    End If
                Catch ex As Exception

                End Try

                filastrae.Remove(filatrae)
                j = j + 1 'suma uno para reducir el numero total de filas en el for
                i = i - 1 'resta una porque al eliminar la fila se reduce la cantidad de filas total
            End If
        Next
    End Sub

    Private Sub ComboBox7_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox7.SelectedIndexChanged
        llenacombo3(ComboBox7.SelectedItem)
        If ComboBox7.Text = "CT" Then

            ComboBox10.Text = 1
            Label14.Visible = True
            ComboBox10.Visible = True
            muestrafiltro = False
            CheckBox6.Visible = True
            'CheckBox6.Checked = True
            ComboBox6.Enabled = False
            'new
            ComboBox10.Text = ""
            'CheckBox6.Checked = False

            'Añadir al combo "consecutivo" numeracion del 1 al 10
            'Cuando es CT va del 1 al 10
            ComboBox10.Items.Clear()
            For i As Integer = 1 To 10
                ComboBox10.Items.Insert(i - 1, i) 'Index 0 --> i-1 valor --> i
            Next
            'ComboBox10.Text = 1

        Else
            'Añadir al combo "consecutivo" numeracion del 0 al 10
            'Cuando es diferente de CT va de 0 al 10
            'ComboBox10.Items.Clear()
            For i As Integer = 0 To 10
                ComboBox10.Items.Insert(i, i)
            Next
            'ComboBox10.Text = 0
            ComboBox10.Text = 0
            Label14.Visible = False
            ComboBox10.Visible = False
            CheckBox6.Visible = False
            CheckBox6.Checked = False
        End If
    End Sub

    Private Sub ComboBox6_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox6.SelectedIndexChanged
        ComboBox5.Text = ""
        ComboBox8.Text = ""
        llenacombo5y8()
        If ComboBox7.Text = "CT" Then
            ComboBox6.Enabled = False
        End If

        'Solo llena la pestaña de Acc
        cargarinfoCliente(ComboBox6.Text, 2)

    End Sub

    Private Sub Button19_Click(sender As Object, e As EventArgs) Handles Button19.Click

        Dim cargoOgPvRcOm As Boolean
        cargoOgPvRcOm = False
        Dim usuario As String

        Dim esAcc = True
        If CheckBox5.Checked = True Then
            MsgBox("pregunta raya vacia")
        End If

        If rayaparvacia = ComboBox8.SelectedItem Or rayaimparvacia = ComboBox5.SelectedItem Then
            If ComboBox5.SelectedItem <> Nothing Or ComboBox8.SelectedItem <> Nothing Then

                crearaya(esAcc)
            End If
        End If



        Dim idofa, ofa, idofap As String
        If CheckBox5.Checked = True Then
            MsgBox("va a entrar a los de almacen")
        End If

        If ComboBox5.SelectedItem <> Nothing Then
            ofa = ComboBox6.Text & "-" & ComboBox5.Text
            If CheckBox5.Checked = True Then
                MsgBox("va por la idofa de almacen con ofa " & ofa)
            End If

            idofa = capturaidofa(ofa, ComboBox7.SelectedItem, ComboBox10.SelectedItem)
            If idofa = 0 Then
                ofa = ComboBox6.Text
                idofa = capturaidofa(ofa, ComboBox7.SelectedItem, ComboBox10.SelectedItem)
            End If
            If CheckBox5.Checked = True Then
                MsgBox("va a ver lo del erp de almancen")
            End If

            Dim tienesaldosp = verificasaldosp(idofa)
            'If tienesaldosp = True Then
            'MsgBox("no se puede cargar la raya porque tiene etiquetas generadas en producción")
            'Exit Sub
            'End If

            Dim cons As String


            If impidereemplazar = True Then
                MsgBox("no se puede cargar la raya porque tiene elementos cargados")
                Exit Sub

            ElseIf CheckBox3.Checked Then
                If ComboBox7.SelectedItem = "OG" Or ComboBox7.SelectedItem = "OM" Or ComboBox7.SelectedItem = "PV" Or ComboBox7.SelectedItem = "RC" Or ComboBox7.SelectedItem = "ID" Then
                    cargoOgPvRcOm = True
                Else
                    cargoOgPvRcOm = False
                End If
                cons = "DELETE FROM Of_Accesorios WHERE (Id_Ofa = '" & idofa & "')"
                Using connection As New SqlConnection(Str_Con_Bd)
                    Dim command As SqlCommand = New SqlCommand(cons, connection)
                    command.CommandTimeout = 600
                    connection.Open()
                    command.ExecuteNonQuery()
                End Using
                cons = ""
            End If

            cons = "UPDATE Orden SET Ref = '" & TextBox3.Text & "', Accesorios = 1 " +
            "WHERE (Id_Ofa = " & idofa & ")"
            Using connection As New SqlConnection(Str_Con_Bd)
                Dim command As SqlCommand = New SqlCommand(cons, connection)
                connection.Open()
                command.ExecuteNonQuery()
            End Using

            grills = Nothing
            ReDim grills(0)
            grills(0) = DataGridView4

            idofap = capturaidofap(ofa)
            'cargar accesorios
            CreaRQI.MyCommands.cargaordenacc(idofa, grills, ofa, idofap, TextBox3.Text, usuario, CheckBox4.Checked, CheckBox5.Checked, loginvalidez, darusuario, "", darpassw, ComboBox7.Text)

            rayaimparvacia = rayaimparvacia + 2
        End If

        If CheckBox5.Checked = True Then
            MsgBox("va a entrar a los de planta")
        End If
        If ComboBox8.SelectedItem <> Nothing Then


            ofa = ComboBox6.Text & "-" & ComboBox8.Text
            If CheckBox5.Checked = True Then
                MsgBox("va por la idofa de planta con ofa " & ofa)
            End If
            idofa = capturaidofa(ofa, ComboBox7.SelectedItem, ComboBox10.SelectedItem)
            If idofa = 0 Then
                ofa = ComboBox6.Text
                idofa = capturaidofa(ofa, ComboBox7.SelectedItem, ComboBox10.SelectedItem)
            End If
            If CheckBox5.Checked = True Then
                MsgBox("va a ver lo del erp de planta")
            End If
            Dim tienesaldosp = verificasaldosp(idofa)
            'If tienesaldosp = True Then
            'MsgBox("no se puede cargar la raya porque tiene etiquetas generadas en producción")
            'Exit Sub
            'End If
            Dim cons As String


            If impidereemplazar = True Then
                MsgBox("no se puede cargar la raya porque tiene elementos cargados")
                Exit Sub

            ElseIf CheckBox3.Checked Then
                If cargoOgPvRcOm = False Then
                    cons = "DELETE FROM Of_Accesorios WHERE (Id_Ofa = '" & idofa & "')"
                    Using connection As New SqlConnection(Str_Con_Bd)
                        Dim command As SqlCommand = New SqlCommand(cons, connection)
                        command.CommandTimeout = 600
                        connection.Open()
                        command.ExecuteNonQuery()
                    End Using
                End If
            End If

            cons = "UPDATE Orden SET Ref = '" & TextBox2.Text & "', Accesorios = 1, Ord_Acc_Planos = " & CInt(TextBox4.Text) & " " +
            "WHERE (Id_Ofa = " & idofa & ")"
            Using connection As New SqlConnection(Str_Con_Bd)
                Dim command As SqlCommand = New SqlCommand(cons, connection)
                connection.Open()
                command.ExecuteNonQuery()
            End Using

            grills = Nothing
            ReDim grills(0)
            grills(0) = DataGridView7

            idofap = capturaidofap(ofa)
            'cargar accesorios

            CreaRQI.MyCommands.cargaordenacc(idofa, grills, ofa, idofap, TextBox2.Text, usuario, CheckBox4.Checked, CheckBox5.Checked, loginvalidez, darusuario, "", darpassw, ComboBox7.Text)
            rayaparvacia = rayaparvacia + 2
        End If
        MsgBox("cargado con exito")
        llenacombo5y8()
        Button25.PerformClick()
    End Sub

    Private Sub ComboBox5_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox5.SelectedIndexChanged
        Verif_Estatus_Acc(ComboBox6.Text & "-" & ComboBox5.Text, False) 'aca tambien verifica si esta vacia

        'trae referencia y llevar todo a la par
        Dim reemplazar As MsgBoxResult

        If entregada = False And rayalibre = True Then
            semaforo = "abierto"
            impidereemplazar = False
            CheckBox3.Checked = False
            'Button19.Enabled = True
            Button19.BackColor = System.Drawing.Color.Transparent
            Button19.Text = "Cargar Orden"
        ElseIf entregada = False And rayalibre = False Then
            semaforo = "modificable"
            Button19.BackColor = System.Drawing.Color.Yellow
            Button19.Text = "Cargar Orden"
            reemplazar = MsgBox("la raya seleccionada tiene elementos cargados, desea reemplazar?", MsgBoxStyle.OkCancel, "reemplazar")
            If reemplazar = MsgBoxResult.Ok Then
                CheckBox3.Checked = True
                impidereemplazar = False
                'Button19.Enabled = True
            Else
                impidereemplazar = True
                CheckBox3.Checked = False
                Button19.Enabled = False
                'ComboBox4.SelectedValue = ""
            End If
        ElseIf entregada = True Then
            semaforo = "entregado"
            impidereemplazar = True
            CheckBox3.Checked = False
            Button19.Enabled = False
            Button19.BackColor = System.Drawing.Color.Red
            Button19.Text = "ya entregada"
        End If
        Button21.Enabled = True
    End Sub

    Private Sub ComboBox8_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox8.SelectedIndexChanged
        Verif_Estatus_Acc(ComboBox6.Text & "-" & ComboBox8.Text, True)
        'traereferencia y llevar todo a la par
        Dim reemplazar As MsgBoxResult

        If entregada = False And rayalibre = True Then
            semaforo = "abierto"
            impidereemplazar = False
            CheckBox3.Checked = False
            'Button19.Enabled = True
            Button19.BackColor = System.Drawing.Color.Transparent
            Button19.Text = "Cargar Orden"
        ElseIf entregada = False And rayalibre = False Then
            semaforo = "modificable"
            Button19.BackColor = System.Drawing.Color.Yellow
            Button19.Text = "Cargar Orden"
            reemplazar = MsgBox("la raya seleccionada tiene elementos cargados, desea reemplazar?", MsgBoxStyle.OkCancel, "reemplazar")
            If reemplazar = MsgBoxResult.Ok Then
                CheckBox3.Checked = True
                impidereemplazar = False
                Button19.Enabled = True
                CheckBox3.Enabled = True
            Else
                impidereemplazar = True
                CheckBox3.Checked = False
                Button19.Enabled = False
                'ComboBox4.SelectedValue = ""
            End If
        ElseIf entregada = True Then
            semaforo = "entregado"
            impidereemplazar = True
            CheckBox3.Checked = False
            Button19.Enabled = False
            Button19.BackColor = System.Drawing.Color.Red
            Button19.Text = "ya entregada"
        End If
        Button21.Enabled = True
    End Sub
    Sub Verif_Estatus_Acc(ByVal ofa As String, ByVal par As Boolean)
        Dim cons As String
        'si la orden es par verifica si esta despachada y si tiene OA
        'si la orden es impar verifica la cantidad enviada
        If par = True Then
            cons = "SELECT Of_Accesorios.Despachado, Of_Accesorios.OA_UnoEE FROM Of_Accesorios WITH (nolock) INNER JOIN Vista_Item ON Of_Accesorios.Id_UnoE = Vista_Item.codErp INNER JOIN Orden WITH (nolock) ON Of_Accesorios.Id_Ofa = Orden.Id_Ofa " +
                            "WHERE (Orden.Ofa = '" & ofa & "') AND (Of_Accesorios.Despachado = 0) AND (Of_Accesorios.OA_UnoEE > 1) "
            If ComboBox7.SelectedItem = "CT" Then
                cons = cons & "AND (Orden.No_Mod = " & ComboBox10.SelectedItem & ") "
            End If
            cons = cons & "ORDER BY Of_Accesorios.No_Item"
        Else
            cons = "SELECT Of_Accesorios.Cant_Enviada FROM Of_Accesorios WITH (nolock) INNER JOIN Vista_Item ON Of_Accesorios.Id_UnoE = Vista_Item.codErp INNER JOIN Orden WITH (nolock) ON Of_Accesorios.Id_Ofa = Orden.Id_Ofa " +
                            "WHERE (Orden.Ofa = '" & ofa & "') AND (Of_Accesorios.Cant_Enviada > 0) "
            If ComboBox7.SelectedItem = "CT" Then
                cons = cons & "AND (Orden.No_Mod = " & ComboBox10.SelectedItem & ") "
            End If
            cons = cons & "ORDER BY Of_Accesorios.No_Item"
        End If

        Using connection As New SqlConnection(Str_Con_Bd)
            Dim command As SqlCommand = New SqlCommand(cons, connection)

            connection.Open()
            Dim reader As SqlDataReader
            Try
                reader = command.ExecuteReader()
            Catch ex As Exception
                entregada = False
                rayalibre = True
                Exit Sub
            End Try

            If reader.HasRows Then
                entregada = True
                rayalibre = False
                reader.Read()

            Else
                entregada = False
                rayalibre = True
            End If
            command = Nothing
            reader.Close()
            reader = Nothing

        End Using
        If entregada = False Then
            cons = "SELECT ord_fec_Imp_Acc FROM Orden WITH (nolock) WHERE (Ofa = '" & ofa & "')"
            If ComboBox7.SelectedItem = "CT" Then
                'cons = cons & " AND (No_Mod = " & ComboBox10.SelectedItem & ")"
                cons = "SELECT Id_Emp_Ing FROM Orden WITH (nolock) WHERE (Ofa = '" & ofa & "') AND (No_Mod = " & ComboBox10.SelectedItem & ")"
            End If
            Using connection As New SqlConnection(Str_Con_Bd)
                Dim command As SqlCommand = New SqlCommand(cons, connection)

                connection.Open()
                Dim reader As SqlDataReader
                Try
                    reader = command.ExecuteReader()
                Catch ex As Exception
                    'entregada = False
                    'rayalibre = True
                    Exit Sub
                End Try
                If reader.HasRows Then
                    reader.Read()
                    If IsDBNull(reader.GetValue(0)) Then

                    Else
                        entregada = True
                        rayalibre = False
                    End If
                End If
            End Using
        End If
    End Sub

    Private Sub CheckBox3_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox3.CheckedChanged
        If impidereemplazar = True And CheckBox3.Checked = True Then
            impidereemplazar = False
            Button19.Enabled = True
            Button12.BackColor = System.Drawing.Color.Transparent

        ElseIf impidereemplazar = False And CheckBox3.Checked = False Then
            impidereemplazar = True
            Button19.Enabled = False
            Button12.BackColor = System.Drawing.Color.Red
        End If
    End Sub



    Private Sub Button22_Click(sender As Object, e As EventArgs) Handles Button22.Click
        CreaRQI.MyCommands.eliminaPALERT(grillas)
    End Sub

    Private Sub Button23_Click(sender As Object, e As EventArgs) Handles Button23.Click
        Dim rayacrear As Integer
        Try
            rayacrear = InputBox("Si la raya no existe, digite el numero de raya a crear")
            If verificaraya(rayacrear) = False Then 'falso raya no existe  verdadero raya si existe
                crearaya(, rayacrear)
                MsgBox("seleccione la raya que ha creado para cargar allí")

                llenacombo4()
            Else
                MsgBox("la raya que intentaste crear ya existe")
            End If
        Catch ex As Exception
            MsgBox("no digitó un número entero, intente de nuevo")
        End Try
    End Sub
    Function verificaraya(rayacrear As Integer) As Boolean 'verdadero raya si existe falso raya no existe 
        Dim cons As String
        cons = "SELECT letra AS Raya FROM Orden WITH (nolock) " +
                "WHERE (Tipo_Of = '" & ComboBox2.SelectedItem & "') AND (RTRIM(Numero) + '-' + RTRIM(ano) LIKE '" & ComboBox3.SelectedItem & "') AND (Anulada = 0) AND (letra = '" & rayacrear & "')"
        Using connection As New SqlConnection(Str_Con_Bd)
            Dim command As SqlCommand = New SqlCommand(cons, connection)

            connection.Open()
            Dim reader As SqlDataReader
            reader = command.ExecuteReader()
            If reader.HasRows Then
                reader.Read()
                verificaraya = True
            Else
                verificaraya = False
            End If
            command = Nothing
            reader.Close()
            reader = Nothing
        End Using
    End Function

    Private Sub Button24_Click(sender As Object, e As EventArgs) Handles Button24.Click
        Dim ofa As String
        If ComboBox4.Text <> "" Then
            ofa = ComboBox3.Text & "-" & ComboBox4.Text
            Dim cons As String = "UPDATE Orden SET AutoInventor = 'CARGADA' WHERE (Ofa = '" & ofa & "') AND (Tipo_Of = '" & ComboBox2.Text & "')"
            Using connection As New SqlConnection(Str_Con_Bd)
                Dim command As SqlCommand = New SqlCommand(cons, connection)
                connection.Open()
                command.ExecuteNonQuery()
            End Using
        End If
    End Sub

    Private Sub Button25_Click(sender As Object, e As EventArgs) Handles Button25.Click
        Dim ofa As String
        Dim cons As String
        If ComboBox5.Text <> "" Then
            ofa = ComboBox6.Text & "-" & ComboBox5.Text
            cons = "UPDATE Orden SET AutoInventor = 'CARGADA' WHERE (Ofa = '" & ofa & "') AND (Tipo_Of = '" & ComboBox7.Text & "')"
            Using connection As New SqlConnection(Str_Con_Bd)
                Dim command As SqlCommand = New SqlCommand(cons, connection)
                connection.Open()
                command.ExecuteNonQuery()
            End Using
        End If
        If ComboBox8.Text <> "" Then
            ofa = ComboBox6.Text & "-" & ComboBox8.Text
            cons = "UPDATE Orden SET AutoInventor = 'CARGADA' WHERE (Ofa = '" & ofa & "') AND (Tipo_Of = '" & ComboBox7.Text & "')"
            Using connection As New SqlConnection(Str_Con_Bd)
                Dim command As SqlCommand = New SqlCommand(cons, connection)
                connection.Open()
                command.ExecuteNonQuery()
            End Using
        End If
    End Sub

    Private Sub ComboBox9_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox9.SelectedIndexChanged
        'llenacombo4()
        ComboBox3.Text = ""
        ComboBox4.Items.Clear()
        ComboBox4.Text = ""
        ComboBox3.Enabled = True
    End Sub

    Private Sub ComboBox10_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox10.SelectedIndexChanged
        'llenacombo5y8()
        ComboBox6.Text = ""
        ComboBox5.Items.Clear()
        ComboBox8.Items.Clear()
        ComboBox5.Text = ""
        ComboBox8.Text = ""
        ComboBox6.Enabled = True
    End Sub

    Private Sub CheckBox6_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox6.CheckedChanged

        If muestrafiltro = False Then 'false
            muestrafiltro = True 'true
        Else
            If ComboBox7.Text = "CT" Then
                If CheckBox6.Checked = False Then 'true
                    Dim filenametxt = "\\172.21.0.202\ingenieria\Z_VARIOS_SCI\FILTRO.txt"
                    If My.Computer.FileSystem.FileExists(filenametxt) Then
                        Process.Start(filenametxt)
                    Else
                        Dim FN As Integer
                        FN = FreeFile()
                        Using sw As New StreamWriter(filenametxt, False, System.Text.Encoding.Default)
                        End Using
                        Process.Start(filenametxt)
                    End If
                End If
            End If

        End If

    End Sub

    Private Sub Button26_Click(sender As Object, e As EventArgs) Handles Button26.Click
        'Dim habilitarcarga As Boolean = True

        Dim grillasacc(1) As Object
        grillasacc(0) = DataGridView4
        grillasacc(1) = DataGridView7
        Dim filas As System.Windows.Forms.DataGridViewRowCollection
        Dim fila As System.Windows.Forms.DataGridViewRow
        Dim j As Integer = 1
        For k = 0 To UBound(grillasacc)
            filas = grillasacc(k).rows

            For i = 0 To filas.Count - j
                If i = filas.Count Then 'verifica que esta en la ultima fila
                    Exit For
                End If
                fila = filas.Item(i)

                If fila.Selected = True Then
                    Dim datosfila() As String

                    ReDim datosfila(10)
                    datosfila(0) = fila.Cells(0).Value 'cant
                    datosfila(1) = fila.Cells(1).Value 'nomenclatura
                    datosfila(2) = fila.Cells(9).Value 'desaux
                    datosfila(3) = fila.Cells(2).Value 'dim1
                    datosfila(4) = fila.Cells(3).Value 'dim2
                    datosfila(5) = fila.Cells(4).Value 'dim3
                    datosfila(6) = fila.Cells(5).Value 'dim4
                    datosfila(7) = fila.Cells(6).Value 'dim5
                    datosfila(8) = fila.Cells(7).Value 'dim6
                    'datosfila(9) = fm areauni  acc null
                    'datosfila(10) = fm aretot  acc null

                    'cant = filatrae.Cells(0).Value
                    'descri = filatrae.Cells(1).Value
                    'desaux = filatrae.Cells(9).Value
                    'dim1 = filatrae.Cells(2).Value
                    'dim2 = filatrae.Cells(3).Value
                    'dim3 = filatrae.Cells(4).Value
                    'dim4 = filatrae.Cells(5).Value
                    'dim5 = filatrae.Cells(6).Value
                    'dim6 = filatrae.Cells(7).Value
                    'plano = filatrae.Cells(8).Value
                    'codigo = filatrae.Cells(10).Value
                    'nomlargo = filatrae.Cells(11).Value
                    'cantvalida = filatrae.Cells(12).Value



                    Dim codigo, paroimpar, plantaid, nomlargo, cantvalida, codigosId As String
                    If Label11.Text = "Colombia" Then
                        plantaid = 1
                    ElseIf Label11.Text = "Mexico" Then
                        plantaid = 2
                    ElseIf Label11.Text = "Brasil" Then
                        plantaid = 3
                    End If
                    codigosId = "0"
                    elijecodigo(codigo, paroimpar, datosfila, plantaid, cantvalida, True, codigosId)
                    nomlargo = traenomlargo(codigo, plantaid)
                    fila.Cells(10).Value = codigo
                    fila.Cells(11).Value = nomlargo
                    fila.Cells(15).Value = codigosId
                Else
                    'If fila.Cells(0).Value > 0 And fila.Cells(10).Value = 0 Then
                    'habilitarcarga = False
                    'End If
                End If
            Next
        Next
        'If habilitarcarga = True Then
        'Button19.Enabled = True
        'CheckBox3.Enabled = True
        'End If
    End Sub

    Private Sub Btn_Prueba_Click(sender As Object, e As EventArgs) Handles Btn_Prueba.Click
        'Dim frm As Frm_Prueba = New Frm_Prueba
        Dim frm As Frm_ListFormaletas = New Frm_ListFormaletas
        frm.Show()
    End Sub

    Private Sub btn_listacc_Click(sender As Object, e As EventArgs) Handles btn_listacc.Click
        Dim url As String = "http://app.forsa.com.co/siomaestros/ReportViewer.aspx?/RepIng/rptListaItemAccesorio"
        Process.Start(url)
    End Sub

    Private Sub ValDesAux_Click(sender As Object, e As EventArgs) Handles ValDesAux.Click


        'Grilla de Accesorios - Almacen
        If DataGridView4.Rows.Count > 0 Then
            For Each Fila As DataGridViewRow In DataGridView4.Rows

                Dim desaux As String = Fila.Cells(9).Value

                If InStr(desaux, "TIPO_") Then
                    Fila.Cells(9).Value = desaux.Substring(5)
                End If

            Next
        End If

        'Grilla de Accesorios - Planta
        If DataGridView7.Rows.Count > 0 Then
            For Each Fila As DataGridViewRow In DataGridView7.Rows

                Dim desaux As String = Fila.Cells(9).Value

                If InStr(desaux, "TIPO_") Then
                    Fila.Cells(9).Value = desaux.Substring(5)
                End If

            Next
        End If
    End Sub

    Private Sub btnVerKits_Click(sender As Object, e As EventArgs) Handles btnVerKits.Click
        Dim url As String = "http://app.forsa.com.co/siomaestros/ReportViewer.aspx?/RepIng/rptKitsAccesorios"
        Process.Start(url)
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If impidereemplazar = True And CheckBox1.Checked = True Then
            impidereemplazar = False
            Button12.Enabled = True
        ElseIf impidereemplazar = False And CheckBox1.Checked = False Then
            impidereemplazar = True
            Button12.Enabled = False
        End If
    End Sub
    Sub Verif_Estatus_Sol()

        Dim idofa, ofa As String
        ofa = ComboBox3.Text & "-" & ComboBox4.Text
        idofa = capturaidofa(ofa, ComboBox2.SelectedItem, ComboBox9.SelectedItem)
        If idofa = 0 Then
            ofa = ComboBox3.Text
            idofa = capturaidofa(ofa, ComboBox2.SelectedItem, ComboBox9.SelectedItem)
        End If
        Dim cons As String = "SELECT CASE WHEN Ord_fec_impresion IS NULL THEN 0 ELSE 1 END AS Entrega FROM Orden WITH (nolock) WHERE (Id_Ofa = " & idofa & ")"
        Using connection As New SqlConnection(Str_Con_Bd)
            Dim command As SqlCommand = New SqlCommand(cons, connection)

            connection.Open()
            Dim reader As SqlDataReader
            Try
                reader = command.ExecuteReader()
            Catch ex As Exception
                entregada = False
                Exit Sub
            End Try

            If reader.HasRows Then
                reader.Read()
                Try
                    If reader.GetValue(0) = 1 Then
                        entregada = True
                    End If
                Catch ex As Exception
                    entregada = False
                End Try

            Else
                entregada = False
            End If
            command = Nothing
            reader.Close()
            reader = Nothing

        End Using
    End Sub
    Public Sub controlesexternos(control As String, Optional valor As String = "0")
        Select Case control
            Case "combo2"
                ComboBox2.Text = valor
            Case "combo3"
                ComboBox3.Text = valor
            Case "combo4"
                ComboBox4.Text = valor
        End Select

    End Sub
End Class