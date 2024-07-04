Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports System.IO

Public Class Frm_InventERP
    Public Function AbreDb_Orl()
        Dim Pass_Bd_Unoee = "unoee2"
        Dim User_Bd_Unoee = "unoee2"
        Dim Inst_Bd_Unoee = "ERPFORSA"
        Dim Con_Ora = New ADODB.Connection
        Dim StrOra = "Provider=ORAOLEDB.ORACLE;Password=" & Pass_Bd_Unoee & ";User ID=" & User_Bd_Unoee & _
                ";Data Source=" & Inst_Bd_Unoee & ";Persist Security Info=True"
        Con_Ora.ConnectionString = StrOra
        Con_Ora.CursorLocation = ADODB.CursorLocationEnum.adUseClient
        Con_Ora.Open()
        Return Con_Ora
    End Function
    Private Sub Frm_InventERP_load(sender As Object, e As EventArgs) Handles MyBase.Load
        ComboBox1.Text = "Seleccione Compañía"
        CargaCias()

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        'Dim a As Double = CheckedListBox1.Height
        'Dim b As Double = CheckedListBox1.Location.Y
        'Dim c As Double = PictureBox1.Height
        'Dim d As Double = PictureBox1.Location.Y
        'Dim f As Double = Me.Height
        Dim CiaVec() As String = Split(ComboBox1.Text, " ")
        Dim Cia As String = CiaVec(0)
        CargaBodega(Cia)
        Dim picturelocation As New System.Drawing.Point(PictureBox1.Location.X, CheckedListBox1.Location.Y + CheckedListBox1.Height + 6)
        PictureBox1.Location = picturelocation
    End Sub
    Sub CargaCias()
        Dim rsiferp2 = New ADODB.Recordset
        Dim Slrs10 = "Select " + _
                "T010_MM_COMPANIAS.F010_ID, " + _
                "T010_MM_COMPANIAS.F010_RAZON_SOCIAL " + _
            "From " + _
                "T010_MM_COMPANIAS " + _
            "Where " + _
                "(T010_MM_COMPANIAS.F010_ID = 15) Or " + _
                "(T010_MM_COMPANIAS.F010_ID = 11) Or " + _
                "(T010_MM_COMPANIAS.F010_ID = 12) Or " + _
                "(T010_MM_COMPANIAS.F010_ID = 14) " + _
            "Order By " + _
                "T010_MM_COMPANIAS.F010_ID"
        Dim Con_Ora As ADODB.Connection = AbreDb_Orl()
        rsiferp2.Open(Slrs10, Con_Ora, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        Dim tablaCias(,) As Object
        tablaCias = rsiferp2.GetRows
        'Dim i As Integer = tablaCias.GetUpperBound(1)
        If tablaCias.GetUpperBound(1) >= 0 Then
            ComboBox1.Items.Clear()
            For i = 0 To tablaCias.GetUpperBound(1)

                ComboBox1.Items.Add(tablaCias(0, i) & " - " & tablaCias(1, i))

            Next
        End If
    End Sub
    Sub CargaBodega(ByVal Cia As String)
        Dim rsiferp2 = New ADODB.Recordset
        Dim Slrs10 = "Select " + _
                        "T150_MC_BODEGAS.F150_ID, " + _
                        "T150_MC_BODEGAS.F150_DESCRIPCION " + _
                    "From " + _
                        "T150_MC_BODEGAS " + _
                    "Where " + _
                        "T150_MC_BODEGAS.F150_ID_CIA = " & Cia & " And " + _
                        "T150_MC_BODEGAS.F150_IND_MULTI_UBICACION = 0 And " + _
                        "T150_MC_BODEGAS.F150_IND_ESTADO = 1 " + _
                    "Order By " + _
                        "T150_MC_BODEGAS.F150_ID"
        Dim Con_Ora As ADODB.Connection = AbreDb_Orl()
        rsiferp2.Open(Slrs10, Con_Ora, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        Dim tablaCos(,) As Object
        tablaCos = rsiferp2.GetRows
        If tablaCos.GetUpperBound(1) >= 0 Then
            'ComboBox2.Items.Clear()
            CheckedListBox1.Items.Clear()

            For i = 0 To tablaCos.GetUpperBound(1)
                'ComboBox2.Items.Add(tablaCos(0, i) & " - " & tablaCos(1, i))
                CheckedListBox1.Items.Add(tablaCos(0, i) & " - " & tablaCos(1, i))
            Next

            If CheckedListBox1.Items.Count > 15 Then
                CheckedListBox1.ScrollAlwaysVisible = True
                CheckedListBox1.Height = 259

                MyBase.Size = New System.Drawing.Size(590, CheckedListBox1.Location.Y + CheckedListBox1.Height + PictureBox1.Height + 47)

            Else
                CheckedListBox1.ScrollAlwaysVisible = False
                CheckedListBox1.Height = (CheckedListBox1.Items.Count * 17) + 4

                'Dim tama As Integer
                'tama = CheckedListBox1.Height + 87

                'If tama < 160 Then
                'MyBase.Size = New System.Drawing.Size(590, 160)
                'Else
                'MyBase.Size = New System.Drawing.Size(590, tama)

                'End If
                MyBase.Size = New System.Drawing.Size(590, CheckedListBox1.Location.Y + CheckedListBox1.Height + PictureBox1.Height + 47)
            End If

        End If
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim CiaVec() As String = Split(ComboBox1.Text, " ")
        Dim Cia As String = CiaVec(0)
        Dim tablaItems As Object
        Dim tablaItemsForline(,) As Object
        tablaItems = items(Cia)
        tablaItemsForline = transformaItems(tablaItems)
        Dim filename = CreaRQI.MyCommands.RutaPlanoAcad
        If filename = Nothing Then
            Exit Sub
        End If
        filename = filename & "InventarioERP.txt"
        'filename = "C:\Inventario Arrendadora " & Format(Now(), "dd-MM-yyyy hh-mm") & ".xls"
        'If Not File.Exists(filename) Then
        ' Create a file to write to. 
        Do While tablaItemsForline(0, UBound(tablaItemsForline, 2)) Is Nothing
            ReDim Preserve tablaItemsForline(4, UBound(tablaItemsForline, 2) - 1)
        Loop
        CreaRQI.MyCommands.escribe(filename, tablaItemsForline, False)
        'End If
        'If path <> "" Then
        If filename <> "" Then
            MsgBox("Inventario Generado Satifactoriamente")
        End If
        Me.Close()
    End Sub
    Function items(ByVal Cia As String)
        Dim tablaItemsFinal(,) As Object
        ReDim tablaItemsFinal(1, 0)
        For i = 0 To CheckedListBox1.Items.Count - 1
            If CheckedListBox1.GetItemChecked(i) Then
                Dim bodegavec() As String = Split(CheckedListBox1.Items.Item(i), " ")
                Dim bodega As String = bodegavec(0)

                Dim rsiferp2 = New ADODB.Recordset
                Dim Slrs10 = "Select " + _
                                "rtrim(SALDOS_AL_DIA.NOMBRE_ITEM), " + _
                                "SALDOS_AL_DIA.EXISTENCIA " + _
                            "From " + _
                                "SALDOS_AL_DIA " + _
                            "Where " + _
                                "SALDOS_AL_DIA.CIA = " & Cia & " And " + _
                                "SALDOS_AL_DIA.BODEGA = '" & bodega & "' And " + _
                                "SALDOS_AL_DIA.EXISTENCIA > 0 " + _
                            "Order By " + _
                                "SALDOS_AL_DIA.NOMBRE_ITEM"
                '"Select " + _
                '"SALDOS_ITEMS_ERP.NOMBRE_ITEM, " + _
                '"SALDOS_ITEMS_ERP.""Existencia_1"", " + _
                '"SALDOS_ITEMS_ERP.""Pendiente_1"" " + _
                '"From " + _
                '"SALDOS_ITEMS_ERP Inner Join " + _
                '"T150_MC_BODEGAS " + _
                '"On T150_MC_BODEGAS.F150_ID_CO = SALDOS_ITEMS_ERP.CO And " + _
                '"T150_MC_BODEGAS.F150_ID = SALDOS_ITEMS_ERP.BODEGA And " + _
                '"T150_MC_BODEGAS.F150_ID_CIA = SALDOS_ITEMS_ERP.EMPRESA " + _
                '"Where " + _
                '"SALDOS_ITEMS_ERP.EMPRESA = " & Cia & " And " + _
                '"SALDOS_ITEMS_ERP.TIPO_ITEM = 1 And " + _
                '"SALDOS_ITEMS_ERP.BODEGA = '" & bodega & "' And " + _
                '"SALDOS_ITEMS_ERP.F470_ID_PERIODO = 201606 And " + _
                '"SALDOS_ITEMS_ERP.F122_FACTOR <> 1 " + _
                '"Order By " + _
                '"SALDOS_ITEMS_ERP.NOMBRE_ITEM"
                Dim Con_Ora As ADODB.Connection = AbreDb_Orl()
                rsiferp2.Open(Slrs10, Con_Ora, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                Dim tablaItems(,) As Object
                tablaItems = rsiferp2.GetRows

                Dim sumado As Boolean

                For j = 0 To tablaItems.GetUpperBound(1)
                    sumado = False
                    For k = 0 To tablaItemsFinal.GetUpperBound(1)
                        If tablaItems(0, j) = tablaItemsFinal(1, k) Then
                            'Dim h As Integer = tablaItems(1, j) - tablaItems(2, j)
                            'If tablaItems(1, j) - tablaItems(2, j) <= 0 Then
                            'Exit For
                            'tablaItemsFinal(0, k) = 0
                            'Else
                            tablaItemsFinal(0, k) = tablaItems(1, j) + tablaItemsFinal(0, k) ' - tablaItems(2, j)
                            'End If

                            sumado = True
                        End If
                    Next
                    If sumado = False Then
                        'Dim h As Integer = tablaItems(1, j) - tablaItems(2, j)
                        'If tablaItems(1, j) - tablaItems(2, j) <= 0 Then
                        'Continue For
                        'tablaItemsFinal(0, tablaItemsFinal.GetUpperBound(1)) = 0
                        'Else
                        tablaItemsFinal(0, tablaItemsFinal.GetUpperBound(1)) = tablaItems(1, j) ' - tablaItems(2, j)
                        'End If
                        tablaItemsFinal(1, tablaItemsFinal.GetUpperBound(1)) = tablaItems(0, j)
                        ReDim Preserve tablaItemsFinal(1, tablaItemsFinal.GetUpperBound(1) + 1)
                    End If
                Next

                Do While tablaItemsFinal(0, tablaItemsFinal.GetUpperBound(1)) = Nothing
                    ReDim Preserve tablaItemsFinal(1, tablaItemsFinal.GetUpperBound(1) - 1)
                Loop


            End If
        Next
        Return tablaItemsFinal
    End Function
    Function transformaItems(ByVal tablaItems(,) As Object)
        Dim tablatransformada(,) As Object
        ReDim tablatransformada(4, 0)
        Dim itemVector() As String

        For i = 0 To tablaItems.GetUpperBound(1)
            Dim numero As Double
            Dim J As Integer = tablatransformada.GetUpperBound(1)
            itemVector = Split(tablaItems(1, i), " ")
            If itemVector(0) = "PANEL" Or itemVector(0) = "FILLER" Then
                If itemVector(1) = "DESPUNTE" Then
                    If itemVector(2) = "MONO" Then
                        Try
                            numero = Convert.ToDouble(itemVector(3))
                            tablatransformada(0, J) = tablaItems(0, i)
                            'If itemVector(1) = "DESPUNTE" Then
                            tablatransformada(1, J) = "MFMD"
                            'End If
                            tablatransformada(2, J) = itemVector(3)
                            tablatransformada(3, J) = itemVector(5)

                            ReDim Preserve tablatransformada(4, tablatransformada.GetUpperBound(1) + 1)
                        Catch ex As Exception
                        End Try
                    Else
                        Try
                            numero = Convert.ToDouble(itemVector(2))
                            tablatransformada(0, J) = tablaItems(0, i)
                            'If itemVector(1) = "DESPUNTE" Then
                            tablatransformada(1, J) = "MFMD"
                            'End If
                            tablatransformada(2, J) = itemVector(2)
                            tablatransformada(3, J) = itemVector(4)

                            ReDim Preserve tablatransformada(4, tablatransformada.GetUpperBound(1) + 1)
                        Catch ex As Exception
                        End Try
                    End If
                Else
                    Try
                        numero = Convert.ToDouble(itemVector(1))
                        tablatransformada(0, J) = tablaItems(0, i)
                        If itemVector(0) = "PANEL" Or itemVector(0) = "FILLER" Then
                            tablatransformada(1, J) = "MFM"
                        End If
                        tablatransformada(2, J) = itemVector(1)
                        tablatransformada(3, J) = itemVector(3)

                        ReDim Preserve tablatransformada(4, tablatransformada.GetUpperBound(1) + 1)
                    Catch ex As Exception

                    End Try
                End If

                If IsNumeric(numero) Then
                    'If itemVector(1) = "ALU-L" Or itemVector(1) = "UNIVERSAL" Or itemVector(1) = "DESPUNTE" Or itemVector(1) = "LOSA" Then
                    'ElseIf itemVector(1) = "DT" Or itemVector(1) = "PUNTAL" Or itemVector(1) = "ESPECIAL" Then
                Else

                End If
            ElseIf itemVector(0) = "ESQ" Then
                If itemVector(1) = "INTERNO" Then
                    Try
                        numero = Convert.ToDouble(itemVector(2))
                        tablatransformada(0, J) = tablaItems(0, i)
                        tablatransformada(1, J) = "MEQM"
                        tablatransformada(2, J) = itemVector(2)
                        tablatransformada(3, J) = itemVector(6)
                        If itemVector(2) <> itemVector(4) Then
                            tablatransformada(4, J) = itemVector(4)
                        End If
                        ReDim Preserve tablatransformada(4, tablatransformada.GetUpperBound(1) + 1)
                    Catch ex As Exception

                    End Try
                    'If itemVector(2) = "+" Then
                    'Else

                    'End If
                ElseIf itemVector(1) = "EXTERIOR" Then
                    tablatransformada(0, J) = tablaItems(0, i)
                    tablatransformada(1, J) = "MAG"
                    tablatransformada(2, J) = itemVector(2)
                    tablatransformada(3, J) = itemVector(4)
                    ReDim Preserve tablatransformada(4, tablatransformada.GetUpperBound(1) + 1)
                End If
            ElseIf itemVector(0) = "TAPA" Then
                Try
                    numero = Convert.ToDouble(itemVector(2))
                    tablatransformada(0, J) = tablaItems(0, i)
                    tablatransformada(1, J) = "MTH"
                    tablatransformada(2, J) = itemVector(2)
                    tablatransformada(3, J) = itemVector(4)
                    ReDim Preserve tablatransformada(4, tablatransformada.GetUpperBound(1) + 1)
                Catch ex As Exception
                End Try
            ElseIf itemVector(0) = "CUCHILLA" Then
                If itemVector(1) = "DE" And itemVector(4) = "X" Then
                    Try
                        numero = Convert.ToDouble(itemVector(3))
                        tablatransformada(0, J) = tablaItems(0, i)
                        tablatransformada(1, J) = "CU" & itemVector(2)
                        tablatransformada(2, J) = itemVector(3)
                        tablatransformada(3, J) = itemVector(5)
                        ReDim Preserve tablatransformada(4, tablatransformada.GetUpperBound(1) + 1)
                    Catch ex As Exception
                    End Try
                End If
                End If
        Next
        Return tablatransformada
    End Function
End Class