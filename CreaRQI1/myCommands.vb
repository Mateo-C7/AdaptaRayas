' (C) Copyright 2011 by  
'
Imports System
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.EditorInput
Imports System.Text
Imports System.IO
Imports System.Data.SqlClient
Imports CreaRQI.CapaDatos

' This line is not mandatory, but improves loading performances
'<Assembly: CommandClass(GetType(CreaRQI.MyCommands))>
Namespace CreaRQI

    ' This class is instantiated by AutoCAD for each document when
    ' a command is called by the user the first time in the context
    ' of a given document. In other words, non static data in this class
    ' is implicitly per-document!
    Public Class MyCommands

        Dim Conn As ConexionBD = New ConexionBD 'Instanciamos la clase de conexion
        Public Shared bodegaqueseusa As String
        Public Shared Numero_Reg As String
        Public Shared tablaforlinemuros As System.Windows.Forms.DataGridView
        Public Shared tablaForlineUnion As System.Windows.Forms.DataGridView
        Public Shared tablaForlineLosas As System.Windows.Forms.DataGridView
        Public Shared ruta As String
        Public Shared nombre As String
        Public Shared NomUsuario, Passw As String
        Public Shared sigue As Boolean
        Public Shared validez As Boolean
        Public Shared userid As String
        Public Shared versionapp As Double
        ' The CommandMethod attribute can be applied to any public  member 
        ' function of any public class.
        ' The function should take no arguments and return nothing.
        ' If the method is an instance member then the enclosing class is 
        ' instantiated for each document. If the member is a static member then
        ' the enclosing class is NOT instantiated.
        '
        ' NOTE: CommandMethod has overloads where you can provide helpid and
        ' context menu.

        ' Modal Command with localized name
        ' AutoCAD will search for a resource string with Id "MyCommandLocal" in the 
        ' same namespace as this command class. 
        ' If a resource string is not found, then the string "MyLocalCommand" is used 
        ' as the localized command name.
        ' To view/edit the resx file defining the resource strings for this command, 
        ' * click the 'Show All Files' button in the Solution Explorer;
        ' * expand the tree node for myCommands.vb;
        ' * and double click on myCommands.resx
        '<CommandMethod("MyGroup", "MyCommand", "MyCommandLocal", CommandFlags.Modal)> _
        'Public Sub MyCommand() ' This method can have any name
        ' Put your command code here
        'End Sub

        ' Modal Command with pickfirst selection
        '<CommandMethod("MyGroup", "MyPickFirst", "MyPickFirstLocal", CommandFlags.Modal + CommandFlags.UsePickSet)> _
        'Public Sub MyPickFirst() ' This method can have any name
        'Dim result As PromptSelectionResult = Application.DocumentManager.MdiActiveDocument.Editor.GetSelection()
        'If (result.Status = PromptStatus.OK) Then
        ' There are selected entities
        ' Put your command using pickfirst set code here
        'Else
        ' There are no selected entities
        ' Put your command code here
        'End If
        'End Sub

        ' Application Session Command with localized name
        '<CommandMethod("MyGroup", "MySessionCmd", "MySessionCmdLocal", CommandFlags.Modal + CommandFlags.Session)> _
        'Public Sub MySessionCmd() ' This method can have any name
        ' Put your command code here
        'End Sub
        '<CommandMethod("CreaRQI")>
        'Public Sub CreaRQI() '<CommandMethod("MyGroup", "CreaRQI", "MySessionCMDLocal", CommandFlags.Modal + CommandFlags.Session)> _
        '    Dim formato As New Frm_CargaListados
        '    formato.StartPosition = Windows.Forms.FormStartPosition.CenterScreen
        '    formato.ShowDialog()
        'End Sub
        '<CommandMethod("InventarioERP")>
        'Public Sub InventarioERP() '<CommandMethod("MyGroup", "CreaRQI", "MySessionCMDLocal", CommandFlags.Modal + CommandFlags.Session)> _
        '    Dim formato As New Frm_InventERP
        '    formato.StartPosition = Windows.Forms.FormStartPosition.CenterScreen
        '    formato.ShowDialog()
        'End Sub
        '<CommandMethod("CargaSif")>
        'Public Sub CargaSif()
        '    If esVersionActual() = False Then Exit Sub
        '    Dim formato As New Frm_CargaListados
        '    formato.Text = "Listado de Formaletas V" & versionapp
        '    formato.StartPosition = Windows.Forms.FormStartPosition.CenterScreen
        '    formato.Tab_CargarInfo.SelectedTab = formato.TabPage_SIIF
        '    formato.ShowDialog()
        'End Sub
        '<CommandMethod("CargaAcc")>
        'Public Sub CargaSif2()
        '    If esVersionActual() = False Then Exit Sub
        '    Dim planta As New Frm_SeleccionarPlantaACC
        '    planta.Text = "Listado de Formaletas V" & versionapp
        '    planta.StartPosition = Windows.Forms.FormStartPosition.CenterScreen
        '    planta.ShowDialog()
        '    Dim formato As New Frm_CargaListados
        '    formato.StartPosition = Windows.Forms.FormStartPosition.CenterScreen
        '    formato.Tab_VerInfoCargada.SelectedTab = formato.TabPage_ACC_Ver
        '    formato.Tab_CargarInfo.SelectedTab = formato.TabPage_ACC_Carga
        '    formato.Label11.Text = planta.planta
        '    planta.Close()
        '    formato.ShowDialog()
        'End Sub
        '<CommandMethod("AdaptaRayas")>
        'Public Sub AdaptaRayas() '<CommandMethod("MyGroup", "CreaRQI", "MySessionCMDLocal", CommandFlags.Modal + CommandFlags.Session)> _
        '    If esVersionActual() = False Then Exit Sub
        '    validalog()
        '    If validez = True Then
        '        SolicitaOF()
        '    End If

        'End Sub

        '<CommandMethod("MurosCurvos")>
        'Public Sub MurosCurvos()
        '    If esVersionActual() = False Then Exit Sub
        '    validalog()
        '    If validez = True Then
        '        MurosCurvosForm()
        '    End If
        'End Sub

        '<LispFunction("ADAPO")> _
        'Public Shared Function VBfunction(ByVal rbfArgs As ResultBuffer) As ResultBuffer
        'Dim rbfResult As ResultBuffer
        'Return rbfResult
        'End Function
        Private Sub SolicitaOF()
            Dim forma As New Frm_ListFormaletas
            'forma.GroupBox1.Text = nombre
            forma.StartPosition = Windows.Forms.FormStartPosition.CenterScreen
            forma.ShowDialog()
        End Sub
        Private Sub MurosCurvosForm()
            Dim formamuroscurvos As New Frm_FormaletaRecMuro
            formamuroscurvos.StartPosition = Windows.Forms.FormStartPosition.CenterScreen
            formamuroscurvos.ShowDialog()
        End Sub
        Private Sub validalog()
            If validez = False Then
                Dim Log As New Frm_Login_2
                Log.StartPosition = Windows.Forms.FormStartPosition.CenterScreen
                Log.ShowDialog()
                validez = Log.validez
                nombre = Log.nombre
                If validez = True Then
                    Log.Close()
                End If
            End If
        End Sub

        ' LispFunction is similar to CommandMethod but it creates a lisp 
        ' callable function. Many return types are supported not just string
        ' or integer.
        '<LispFunction("MyLispFunction", "MyLispFunctionLocal")> _
        'Public Function MyLispFunction(ByVal args As ResultBuffer) ' This method can have any name
        ' Put your command code here

        ' Return a value to the AutoCAD Lisp Interpreter
        'Return 1
        'End Function
        Public Shared Function ConsultaCodigoItem(ByVal descriLarga As String, tablaitems(,) As Object) As String

            Dim cod As String()
            For i = 0 To UBound(tablaitems, 2)
                If descriLarga = tablaitems(0, i) Then
                    cod = Split(tablaitems(4, i), " ")
                    ConsultaCodigoItem = cod(0)
                    Exit For
                End If
            Next
        End Function
        Public Shared Function nomenclatura(ByVal descri As String, ancho As String, alto As String) As String
            If descri = "MAG" Then
                nomenclatura = "ESQ EXTERIOR"
            ElseIf descri = "MEQM" Or descri = "MEQL" Or descri = "MFMA" Or descri = "MCPAT" Or descri = "MEQMT" Then
                nomenclatura = "ESQ INTERNO"
            ElseIf descri = "MFM" Or descri = "MCP" Or descri = "MDT" Or descri = "MFL" Or descri = "MFLP" Then
                If Math.Min(CInt(ancho), CInt(alto)) <= 5 Then
                    nomenclatura = "FILLER"
                Else
                    nomenclatura = "PANEL"
                End If
            ElseIf descri = "MTH" Or descri = "MTV" Then
                nomenclatura = "TAPA MURO"
            ElseIf descri = "MFMD" Then
                nomenclatura = "PANEL DESPUNTE MONO"
            Else
                nomenclatura = descri
            End If
        End Function
        Public Shared Function atributos(ByVal descri As String, att1 As String, att2 As String, att3 As String)
            If descri = "ESQ INTERNO" Then
                If CDbl(att3) <> 0 Or att3 = Nothing Then
                    atributos = att1 & " X " & att1 & " X " & att2
                Else
                    atributos = att1 & " X " & att2 & " X " & att3
                End If
            Else
                If IsNumeric(att1) And IsNumeric(att2) Then
                    If descri = "PANEL" Or descri = "FILLER" Or descri = "PANEL DESPUNTE MONO" Then
                        If CDbl(att1) < CDbl(att2) Then
                            atributos = att1 & " X " & att2
                        Else
                            atributos = att2 & " X " & att1
                        End If
                    Else
                        atributos = att1 & " X " & att2
                    End If
                End If
            End If
        End Function
        Shared Function AbreDb_Orl()
            Dim Pass_Bd_Unoee = "unoee2"
            Dim User_Bd_Unoee = "unoee2"
            Dim Inst_Bd_Unoee = "ERPFORSA"
            Dim Con_Ora = New ADODB.Connection
            Dim StrOra = "Provider=ORAOLEDB.ORACLE;Password=" & Pass_Bd_Unoee & ";User ID=" & User_Bd_Unoee &
                    ";Data Source=" & Inst_Bd_Unoee & ";Persist Security Info=True"
            Con_Ora.ConnectionString = StrOra
            Con_Ora.CursorLocation = ADODB.CursorLocationEnum.adUseClient
            Con_Ora.Open()
            Return Con_Ora
        End Function
        Shared Function AbreDB_ERP() As String
            Dim Login_BdERP As String = "Forsaapp"
            Dim Pass_BdERP As String = "Forsa$21$%"
            Dim Instan_BdERP As String = "FORSA"
            Dim Nom_BdERP As String = "siesa-m4-sqlsw-db2.ceqnrhbwqaoo.us-east-1.rds.amazonaws.com,1433" '-->Esta Ip sirve en ambos casos, sin embargo dentro de Forsa es un poco mas lenta la validacion.
            '"172.21.200.3,1433"-->Esta Ip sirve cuando se esta directamente en Forsa 
            AbreDB_ERP = "Password=" & Pass_BdERP & ";Persist Security " +
            "Info=True;User ID=" & Login_BdERP & ";Initial Catalog=" & Instan_BdERP & ";Data Source=" & Nom_BdERP
        End Function
        Shared Function tablaitems(ByVal CIA As String) As Object
            Dim slrs10 As String
            Dim rsiferp2 = New ADODB.Recordset
            slrs10 = "Select " +
                    "rtrim(T120_MC_ITEMS.F120_DESCRIPCION), " +
                    "T120_MC_ITEMS.F120_ID_CIA, " +
                    "T121_MC_ITEMS_EXTENSIONES.F121_IND_ESTADO, " +
                    "T120_MC_ITEMS.F120_IND_TIPO_ITEM, " +
                    "T120_MC_ITEMS.F120_REFERENCIA " +
                "From " +
                    "T120_MC_ITEMS Inner Join " +
                    "T121_MC_ITEMS_EXTENSIONES " +
                    "On T121_MC_ITEMS_EXTENSIONES.F121_ROWID_ITEM = T120_MC_ITEMS.F120_ROWID " +
                "Where " +
                    "T120_MC_ITEMS.F120_ID_CIA = " & CIA & " And T121_MC_ITEMS_EXTENSIONES.F121_IND_ESTADO = 1 And T120_MC_ITEMS.F120_IND_TIPO_ITEM = 1 " +
                "Order By " +
                    "T120_MC_ITEMS.F120_DESCRIPCION"
            Dim Con_Ora As ADODB.Connection = AbreDb_Orl()
            rsiferp2.Open(slrs10, Con_Ora, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            'Dim tablaItems(,) As Object
            tablaitems = rsiferp2.GetRows
        End Function
        Public Shared Sub eliminaPALERT(grillas() As Object) 'convierte P_ALERTA en P en aluminio
            Dim filas As System.Windows.Forms.DataGridViewRowCollection
            Dim fila As System.Windows.Forms.DataGridViewRow
            Dim j As Integer = 1
            For k = 0 To UBound(grillas)
                filas = grillas(k).rows
                For i = 0 To filas.Count - j
                    If i = filas.Count Then 'verifica que esta en la ultima fila
                        Exit For
                    End If
                    fila = filas.Item(i)
                    If fila.Cells(6).Value = "P_ALERTA" Then
                        fila.Cells(6).Value = "P"
                    End If
                Next
            Next
        End Sub
        Public Shared Sub eliminafilas(grillas() As Object)
            Dim filas As System.Windows.Forms.DataGridViewRowCollection
            Dim fila As System.Windows.Forms.DataGridViewRow
            Dim j As Integer = 1
            For k = 0 To UBound(grillas)
                filas = grillas(k).rows

                For i = 0 To filas.Count - j
                    If i = filas.Count Then 'verifica que esta en la ultima fila
                        Exit For
                    End If
                    fila = filas.Item(i)
                    If fila.Selected = True Then
                        filas.Remove(fila)
                        j = j + 1 'suma uno para reducir el numero total de filas en el for
                        i = i - 1 'resta una porque al eliminar la fila se reduce la cantidad de filas total
                    End If
                Next
            Next

        End Sub
        Public Shared Sub muevefila(grillas() As Object)
            Dim numeroitem As Integer
            numeroitem = 0
            Dim preitem = InputBox("hasta cual item deseas mover la selección", "item")

            Do While numeroitem = 0

                If Integer.TryParse(preitem, numeroitem) Then

                Else
                    preitem = InputBox("hasta cual item deseas mover la selección", "item")
                End If

            Loop

            Dim todalagrilla As System.Windows.Forms.DataGridView
            Dim fila As System.Windows.Forms.DataGridViewRow
            Dim filas As System.Windows.Forms.DataGridViewRowCollection
            Dim j As Integer = 1
            For k = 0 To UBound(grillas)
                todalagrilla = grillas(k)
                filas = grillas(k).rows
                For i = 0 To filas.Count - j
                    If i = filas.Count Then 'verifica que esta en la ultima fila
                        Exit For
                    End If
                    fila = filas.Item(i)
                    If fila.Selected = True Then
                        todalagrilla.Rows.Remove(fila)
                        todalagrilla.Rows.Insert(numeroitem - 1, fila)
                        todalagrilla.ClearSelection()
                        todalagrilla.Rows.Item(numeroitem - 1).Selected = True
                        Exit For
                    End If
                Next
            Next
        End Sub
        Public Shared Function traeyalecotiza(ByVal idofa As String) As String

            Dim Str_Con_Bd As String = ConexionBD.getStringConexion()
            traeyalecotiza = 0
            Dim cons As String
            cons = "SELECT Yale_Cotiza FROM Orden WITH (nolock) WHERE (Id_Ofa = " & idofa & ") ORDER BY Id_Ofa DESC"
            Using connection As New SqlConnection(Str_Con_Bd)
                Dim command As SqlCommand = New SqlCommand(cons, connection)

                connection.Open()

                Dim reader As SqlDataReader = command.ExecuteReader()

                If reader.HasRows Then
                    reader.Read()

                    traeyalecotiza = reader.GetValue(0)

                End If
                command = Nothing
                reader.Close()
                reader = Nothing
            End Using
        End Function
        Public Shared Sub cargaordenacc(idofa As String, grillas() As Object, ofa As String, idofap As String, ref As String, ByRef user As String, ByVal checkbox As Boolean, ultra As Boolean, Optional ByRef darvalidez As Boolean = False, Optional ByRef darusuario As String = "", Optional ByRef darnombre As String = "", Optional passw As String = "", Optional tipoorden As String = "")

            Dim limiteitem As Integer
            Dim plantaid As Integer
            Dim Str_Con_Bd As String = ConexionBD.getStringConexion()
            Dim validez2 As Boolean
            Dim codigo, observa, desaux, dime1, dime2, dime3, dime4, dime5, dime6, yalecotiza, nom, valida, claseOP, codigosId As String
            Dim ItemAcc, cantitem, plano, nivel As Integer
            Dim cons As String

            cons = "SELECT TOP (1) No_Item FROM Of_Accesorios WITH (nolock) WHERE (Id_Ofa = " & idofa & ") ORDER BY No_Item DESC"

            Using connection As New SqlConnection(Str_Con_Bd)

                Dim command As SqlCommand = New SqlCommand(cons, connection)

                connection.Open()

                Dim reader As SqlDataReader = command.ExecuteReader()

                If reader.HasRows Then
                    reader.Read()
                    ItemAcc = CDbl(reader.GetValue(0))

                Else
                    ItemAcc = 0
                End If
                command = Nothing
                reader.Close()
                reader = Nothing

            End Using

            cons = "SELECT planta_id FROM Orden WITH (nolock) WHERE (Id_Ofa = " & idofa & ")"

            Using connection As New SqlConnection(Str_Con_Bd)

                Dim command As SqlCommand = New SqlCommand(cons, connection)

                connection.Open()

                Dim reader As SqlDataReader = command.ExecuteReader()

                If reader.HasRows Then
                    reader.Read()
                    plantaid = CInt(reader.GetValue(0))

                Else
                    plantaid = 1
                End If
                command = Nothing
                reader.Close()
                reader = Nothing

            End Using

            Dim filas As System.Windows.Forms.DataGridViewRowCollection
            Dim celdas As System.Windows.Forms.DataGridViewCellCollection
            Dim fila As System.Windows.Forms.DataGridViewRow

            validez2 = darvalidez

            Do While validez2 = False

                Dim Log As New Frm_Login_2
                Log.StartPosition = Windows.Forms.FormStartPosition.CenterScreen

                If passw <> "" And darusuario <> "" Then
                    verifica2(darusuario, passw) 'le da valor a validez
                Else
                    Log.ShowDialog()
                    verifica2(Log.darusuario, Log.darpass) 'le da valor a validez
                    Log.Close()
                End If
                If validez = False Then
                    MsgBox("no se cargó")
                    Exit Sub
                    'Continue Do
                Else
                    'validez ya esta
                    validez2 = validez
                    darvalidez = validez
                    If validez2 = True Then
                        'nombre = Log.nombre ya lo hizo verifica2
                        darnombre = nombre
                        'userid = Log.iduser ya lo hizo verifica2
                        darusuario = userid
                    Else
                        Log.Close()
                        Exit Sub
                    End If
                End If

            Loop

            limiteitem = 1

            Dim progreso As New Frm_BarraCarga
            progreso.ProgressBar1.Minimum = 0
            Dim pro As Integer

            For i = 0 To 4
                Try
                    pro = pro + grillas(i).rows.count
                Catch ex As System.Exception

                End Try
            Next

            yalecotiza = traeyalecotiza(idofa)

            progreso.ProgressBar1.Maximum = pro
            progreso.StartPosition = Windows.Forms.FormStartPosition.CenterScreen
            progreso.Show()
            progreso.TopMost = True
            progreso.Text = "cargando a la orden"

            For Each grilla In grillas
                filas = grilla.rows
                For i = 0 To filas.Count - 1

                    fila = filas.Item(i)
                    codigo = fila.Cells(10).Value
                    If codigo = Nothing Or codigo = 0 Then
                        Continue For
                    End If
                    ItemAcc = ItemAcc + 1
                    cantitem = fila.Cells(0).Value

                    dime1 = fila.Cells(2).Value
                    dime2 = fila.Cells(3).Value
                    dime3 = fila.Cells(4).Value
                    dime4 = fila.Cells(5).Value
                    dime5 = fila.Cells(6).Value
                    dime6 = fila.Cells(7).Value

                    nom = fila.Cells(1).Value

                    If dime1 = Nothing Then
                        dime1 = 0
                    Else
                        If nom = "TWA" Then 'si es twa redondear a 3
                            dime1 = Math.Round(CDbl(dime1), 3).ToString
                        Else
                            dime1 = Math.Round(CDbl(dime1), 2).ToString
                        End If

                    End If
                    If dime2 = Nothing Then
                        dime2 = 0
                    Else
                        dime2 = Math.Round(CDbl(dime2), 2).ToString
                    End If
                    If dime3 = Nothing Then
                        dime3 = 0
                    Else
                        dime3 = Math.Round(CDbl(dime3), 2).ToString
                    End If
                    If dime4 = Nothing Then
                        dime4 = 0
                    Else
                        dime4 = Math.Round(CDbl(dime4), 2).ToString
                    End If
                    If dime5 = Nothing Then
                        dime5 = 0
                    Else
                        dime5 = Math.Round(CDbl(dime5), 2).ToString
                    End If
                    If dime6 = Nothing Then
                        dime6 = 0
                    Else
                        dime6 = Math.Round(CDbl(dime6), 2).ToString
                    End If
                    observa = fila.Cells(9).Value
                    If observa = Nothing Then
                        observa = ""
                    End If
                    Try


                        If fila.Cells(9).Value = "" Then
                            desaux = ""
                        Else
                            desaux = fila.Cells(9).Value & " - "
                        End If
                    Catch ex As System.Exception
                        desaux = ""
                    End Try
                    If dime6 = 0 Then
                        If dime5 = 0 Then
                            If dime4 = 0 Then
                                If dime3 = 0 Then
                                    If dime2 = 0 Then
                                        If dime1 = 0 Then
                                            If fila.Cells(9).Value = "" Then
                                                desaux = ""
                                            Else
                                                desaux = fila.Cells(9).Value
                                            End If
                                        Else
                                            desaux = desaux & dime1
                                        End If
                                    Else
                                        desaux = desaux & dime1 & "x" & dime2
                                    End If
                                Else
                                    desaux = desaux & dime1 & "x" & dime2 & "x" & dime3
                                End If
                            Else
                                desaux = desaux & dime1 & "x" & dime2 & "x" & dime3 & "x" & dime4
                            End If
                        Else
                            desaux = desaux & dime1 & "x" & dime2 & "x" & dime3 & "x" & dime4 & "x" & dime5
                        End If
                    Else
                        desaux = desaux & dime1 & "x" & dime2 & "x" & dime3 & "x" & dime4 & "x" & dime5 & "x" & dime6
                    End If
                    If fila.Cells(8).Value = "P" Or fila.Cells(8).Value = "p" Then
                        plano = 1
                    Else
                        plano = 0
                    End If

                    valida = fila.Cells(12).Value
                    If valida = Nothing Then
                        valida = 0
                    End If

                    codigosId = fila.Cells(15).Value

                    cons = "SELECT Accesorios_Maestro_Nomenclaturas.Nivel, item_planta.tipo_kamban " +
                            "FROM Accesorios_Maestro_Nomenclaturas WITH (nolock) INNER JOIN Accesorios_Codigos WITH (nolock) ON Accesorios_Maestro_Nomenclaturas.Nomenclatura = Accesorios_Codigos.Nomenclatura INNER JOIN item_planta WITH (nolock) ON Accesorios_Codigos.Id_UnoE = item_planta.cod_erp AND Accesorios_Codigos.planta_id = item_planta.planta_id " +
                            "WHERE  (Accesorios_Maestro_Nomenclaturas.Nomenclatura = '" & nom & "') AND (Accesorios_Codigos.Id_UnoE = " & codigo & ") AND (Accesorios_Codigos.planta_id = " & plantaid & ")"
                    Using connection As New SqlConnection(Str_Con_Bd)
                        Dim command As SqlCommand = New SqlCommand(cons, connection)

                        connection.Open()

                        Dim reader As SqlDataReader = command.ExecuteReader()

                        If reader.HasRows Then
                            reader.Read()
                            If observa.Contains("COMPLEMENTARIA") Or observa.Contains("ADICIONAL") Then
                                nivel = 4
                            Else
                                nivel = CInt(reader.GetValue(0))
                            End If
                            If reader.GetValue(1) = "True" Then claseOP = "OK" Else claseOP = "OA"
                        Else
                            nivel = 4
                            claseOP = "OA"
                        End If
                        command = Nothing
                        reader.Close()
                        reader = Nothing
                    End Using


                    Dim id As String
                    '**** ojo of_accesorios no lleva planta id *****
                    cons = "INSERT INTO Of_Accesorios " +
                        "(Id_UnoE, Id_Ofa, No_Item, Cant_Req, Cant_Reci, Cant_Enviada, Despachado, Observa, Anula, fecha_crea, usu_crea, Yale_Cotiza, Peso_ttl, Grupo, Gno, Tiene_Plano, Id_Ofa_Papa, Saldo, Cant_Acc_Valida, Dim1, Dim2, Dim3, Dim4, Dim5, Dim6, Des_Aux, Nomenclatura, Nivel, Clase_Op, AccesoriosCodigosId) " +
                        "VALUES (" & codigo & ", " & idofa & ", " & ItemAcc & ", " & cantitem & ", 0, 0, 0, '" & desaux & "', 0, @fecha, " & userid & ", " & yalecotiza & ", 0, 'ACC', 10, " & plano & ", " & idofap & ", " & cantitem & ", " & valida & ", " & dime1 & ", " & dime2 & ", " & dime3 & ", " & dime4 & ", " & dime5 & ", " & dime6 & ", '" & observa & "', '" & nom & "', " & nivel & ", '" & claseOP & "', " & codigosId & ") " +
                        "SELECT cast(SCOPE_IDENTITY() as int) AS 'ID'"

                    'Pruebas
                    'cons = "INSERT INTO Of_Accesorios " +
                    '    "(Id_UnoE, Id_Ofa, No_Item, Cant_Req, Cant_Reci, Cant_Enviada, Despachado, Observa, Anula, fecha_crea, usu_crea, Yale_Cotiza, Peso_ttl, Grupo, Gno, Tiene_Plano, Id_Ofa_Papa, Saldo, Cant_Acc_Valida, Dim1, Dim2, Dim3, Dim4, Dim5, Dim6, Des_Aux, Nomenclatura, Nivel, Clase_Op) " +
                    '    "VALUES (" & codigo & ", " & idofa & ", " & ItemAcc & ", " & cantitem & ", 0, 0, 0, '" & desaux & "', 0, @fecha, " & userid & ", " & yalecotiza & ", 0, 'ACC', 10, " & plano & ", " & idofap & ", " & cantitem & ", " & valida & ", " & dime1 & ", " & dime2 & ", " & dime3 & ", " & dime4 & ", " & dime5 & ", " & dime6 & ", '" & observa & "', '" & nom & "', " & nivel & ", '" & claseOP & "') " +
                    '    "SELECT cast(SCOPE_IDENTITY() as int) AS 'ID'"

                    If checkbox = True Or ultra = True Then
                        MsgBox(cons)
                    End If


                    Using connection As New SqlConnection(Str_Con_Bd) 'Date.Now.ToString("yyyy-MM-dd HHH:mm:ss")
                        Dim command As SqlCommand = New SqlCommand(cons, connection)
                        command.Parameters.Add("@fecha", SqlDbType.DateTime).Value = DateTime.Now
                        connection.Open()
                        If checkbox = True Then
                            MsgBox("logro cargar")
                        End If

                        id = command.ExecuteScalar
                        'command.ExecuteNonQuery()
                    End Using
                    If plano = 1 Then
                        Dim insertoplano As String
                        cons = "select Tiene_Plano from Of_Accesorios WITH (nolock) where (Id_Orden_Acce = " & id & ")"
                        Using connection As New SqlConnection(Str_Con_Bd)
                            Dim command As SqlCommand = New SqlCommand(cons, connection)
                            connection.Open()
                            Dim reader As SqlDataReader = command.ExecuteReader()
                            If reader.HasRows Then
                                reader.Read()
                                insertoplano = reader.GetValue(0)
                            End If
                            command = Nothing
                            reader.Close()
                            reader = Nothing
                        End Using
                        If insertoplano <> "True" Then
                            cons = "UPDATE Of_Accesorios SET Tiene_Plano = 1 where (Id_Orden_Acce = " & id & ")"
                            Using connection As New SqlConnection(Str_Con_Bd)
                                Dim command As SqlCommand = New SqlCommand(cons, connection)
                                connection.Open()
                                command.ExecuteNonQuery()
                            End Using
                        End If
                    End If

                    limiteitem = limiteitem + 1

                    'actualiza cantidad de planos
                    'cons = "UPDATE Orden SET Id_Emp_Ing = " & userid & ", Planos_Det = " & totalplanos & ", No_Pesp = " & totalesp & ", No_Pest = " & totalest & ", m2 = " & totalm2 & ", vaprx = " & totalvol & ", paprx = " & totalpe & ", Observaciones = '" & obser & "', Ref = '" & ref & "' " +
                    '"WHERE (Id_Ofa = " & idofa & ")"
                    'Using connection As New SqlConnection(Str_Con_Bd)
                    'Dim command As SqlCommand = New SqlCommand(cons, connection)
                    'connection.Open()
                    'Command.ExecuteNonQuery()
                    'End Using

                    progreso.ProgressBar1.Value = limiteitem - 1 'se resta 1 porque limite item empieza en 1, pero la barra inicia en 0

                    infovalidakit(codigo, plantaid, Str_Con_Bd, checkbox, cantitem, idofa, desaux, yalecotiza, idofap, ultra, id, tipoorden, plano, ItemAcc)



                Next
            Next
            cons = "UPDATE Orden SET Id_Emp_Ing = '" & userid & "', Resp_Ingenieria = '" & nombre & "' ,Abierta = 1, Abierta_Acc = 1 " +
            "WHERE (Id_Ofa = " & idofa & ")"
            Using connection As New SqlConnection(Str_Con_Bd)
                Dim command As SqlCommand = New SqlCommand(cons, connection)
                connection.Open()
                command.ExecuteNonQuery()
            End Using
            progreso.Close()
        End Sub

        Public Shared Sub infovalidakit(codigo As String, plantaid As Integer, Str_Con_Bd As String, checkbox As Boolean, cantitem As Integer, idofa As String, desaux As String, yalecotiza As String, idofap As String, ultra As Boolean, id As String, tipoorden As String, plano As Integer, ByRef itemAcc As Integer)
            Dim haykit As Boolean
            haykit = False
            Dim cons, claseOP, nom, CodigosId As String
            Dim nivel As Integer
            cons = "SELECT Acc_Kit_HijoId, Acc_Kit_Can_Hijo, Acc_Kit_CosteaConHijo FROM Accesorio_Kit WITH (nolock) WHERE (Acc_Kit_PapaId = " & codigo & ") AND (planta_id = " & plantaid & ")"
            Using connection As New SqlConnection(Str_Con_Bd)
                Dim command As SqlCommand = New SqlCommand(cons, connection)

                connection.Open()

                Dim reader As SqlDataReader = command.ExecuteReader()

                If reader.HasRows Then
                    If checkbox = True Then
                        MsgBox("lleva kit")
                    End If

                    reader.Read()

                    If tipoorden = "CT" And reader.GetValue(2) = "True" Then 'si sí costea con hijo, no cargar kit hijo 
                        haykit = False
                    Else
                        haykit = True
                    End If
                    codigo = reader.GetValue(0)
                    cantitem = cantitem * CInt(reader.GetValue(1))
                End If
                command = Nothing
                reader.Close()
                reader = Nothing
            End Using

            If haykit = True Then
                cons = "SELECT Accesorios_Maestro_Nomenclaturas.Nivel, item_planta.tipo_kamban, Accesorios_Codigos.Nomenclatura, Accesorios_Codigos.Codigos_Id " +
                        "FROM Accesorios_Maestro_Nomenclaturas WITH (nolock) INNER JOIN Accesorios_Codigos WITH (nolock) ON Accesorios_Maestro_Nomenclaturas.Nomenclatura = Accesorios_Codigos.Nomenclatura INNER JOIN item_planta ON Accesorios_Codigos.planta_id = item_planta.planta_id AND Accesorios_Codigos.Id_UnoE = item_planta.cod_erp " +
                    "WHERE (Accesorios_Codigos.Id_UnoE = " & codigo & ") AND (Accesorios_Codigos.planta_id = " & plantaid & ")"
                Using connection As New SqlConnection(Str_Con_Bd)
                    Dim command As SqlCommand = New SqlCommand(cons, connection)

                    connection.Open()

                    Dim reader As SqlDataReader = command.ExecuteReader()
                    Try
                        If reader.HasRows Then
                            reader.Read()
                            nivel = CInt(reader.GetValue(0))
                            If reader.GetValue(1) = "True" Then claseOP = "OK" Else claseOP = "OA"
                            nom = reader.GetValue(2)
                            CodigosId = reader.GetValue(3)
                        Else
                            nivel = 4
                            claseOP = "OA"
                            nom = ""
                            CodigosId = "0"
                        End If
                    Catch ex As System.Exception
                        nivel = 4
                        claseOP = "OA"
                        nom = ""
                        CodigosId = "0"
                    End Try

                    command = Nothing
                    reader.Close()
                    reader = Nothing
                End Using
                itemAcc = itemAcc + 1

                '**** ojo of_accesorios no lleva planta id *****
                cons = "INSERT INTO Of_Accesorios " +
                "(Id_UnoE, Id_Ofa, No_Item, Cant_Req, Cant_Reci, Cant_Enviada, Despachado, Observa, Anula, fecha_crea, usu_crea, Yale_Cotiza, Peso_ttl, Grupo, Gno, Tiene_Plano, Id_Ofa_Papa, Saldo, Cant_Acc_Valida, Dim1, Dim2, Dim3, Dim4, Dim5, Dim6, Des_Aux, Nivel, Clase_Op, Nomenclatura, CodOrden_Acce_Padre, AccesoriosCodigosId) " +
                "VALUES (" & codigo & ", " & idofa & ", " & itemAcc & ", " & cantitem & ", 0, 0, 0, '" & desaux & "', 0, @fecha, " & userid & ", " & yalecotiza & ", 0, 'ACC', 10, " & plano & ", " & idofap & ", " & cantitem & ", 0, 0, 0, 0, 0, 0, 0, '', " & nivel & ", '" & claseOP & "', '" & nom & "', '" & id & "'," & CodigosId & ") " +
                "SELECT cast(SCOPE_IDENTITY() as int) AS 'ID'"

                If checkbox = True Or ultra = True Then
                    MsgBox(cons)
                End If


                Using connection As New SqlConnection(Str_Con_Bd) 'Date.Now.ToString("yyyy-MM-dd HHH:mm:ss")
                    Dim command As SqlCommand = New SqlCommand(cons, connection)
                    command.Parameters.Add("@fecha", SqlDbType.DateTime).Value = DateTime.Now
                    connection.Open()
                    If checkbox = True Then
                        MsgBox("logro cargar")
                    End If

                    id = command.ExecuteScalar
                    'command.ExecuteNonQuery()
                End Using
                infovalidakit(codigo, plantaid, Str_Con_Bd, checkbox, cantitem, idofa, desaux, yalecotiza, idofap, ultra, id, tipoorden, plano, itemAcc)
            End If
        End Sub
        Public Shared Sub cargaorden(idofa As String, grillas() As Object, ofa As String, idofap As String, cuadros As System.Windows.Forms.DataGridView, ref As String, obser As String, Optional ByRef darvalidez As Boolean = False, Optional ByRef darusuario As String = "", Optional ByRef darnombre As String = "", Optional passw As String = "", Optional ParaArmado As String = "", Optional Pernado As String = "", Optional UltimaEntrega As String = "", Optional Escalera As String = "")

            Dim limiteitem As Integer
            Dim Str_Con_Bd As String = ConexionBD.getStringConexion()
            Dim validez2 As Boolean
            Dim tipoitem, desaux, an1, al1, al2, an2, medinicial, areaitem, grupo As String
            Dim fecha As Double
            Dim filas As System.Windows.Forms.DataGridViewRowCollection
            Dim celdas As System.Windows.Forms.DataGridViewCellCollection
            Dim fila As System.Windows.Forms.DataGridViewRow
            Dim totalm2, totalvol, totalpe As Double
            Dim itemFM, itemFL, itemUM, itemCL, itemTEMP, planoesp, cantitem, gno, idpiezasforsa, totalplanos, totalest, totalesp As Integer
            itemFM = 0
            itemFL = 0
            itemUM = 0
            itemCL = 0
            planoesp = 0
            desaux = ""
            totalplanos = 0
            totalest = 0
            totalesp = 0
            totalm2 = 0
            totalvol = 0
            totalpe = 0
            fecha = DateAndTime.Now.ToOADate

            validez2 = darvalidez
            Do While validez2 = False
                Dim Log As New Frm_Login_2
                Log.StartPosition = Windows.Forms.FormStartPosition.CenterScreen
                If passw <> "" And darusuario <> "" Then
                    verifica2(darusuario, passw) 'le da valor a validez
                Else
                    Log.ShowDialog()
                    verifica2(Log.darusuario, Log.darpass) 'le da valor a validez
                    Log.Close()
                End If
                If validez = False Then
                    MsgBox("no se cargó")
                    Exit Sub
                    'Continue Do
                Else
                    'validez ya esta
                    validez2 = validez
                    darvalidez = validez
                    If validez2 = True Then
                        'nombre = Log.nombre ya lo hizo verifica2
                        darnombre = nombre
                        'userid = Log.iduser ya lo hizo verifica2
                        darusuario = userid
                    Else
                        Log.Close()
                        Exit Sub
                    End If
                End If

            Loop

            limiteitem = 1

            Dim progreso As New Frm_BarraCarga
            progreso.ProgressBar1.Minimum = 0
            Dim pro As Integer
            For i = 0 To 4
                Try
                    pro = pro + grillas(i).rows.count
                Catch ex As System.Exception

                End Try
            Next

            progreso.ProgressBar1.Maximum = pro
            progreso.StartPosition = Windows.Forms.FormStartPosition.CenterScreen
            progreso.Show()
            progreso.TopMost = True
            progreso.Text = "cargando a la orden"

            Dim cons As String

            For Each grilla In grillas
                filas = grilla.rows

                For i = 0 To filas.Count - 1
                    fila = filas.Item(i)

                    If String.IsNullOrEmpty(fila.Cells(8).Value) Then
                        Continue For
                    End If
                    grupo = fila.Cells(8).Value
                    Select Case grupo
                        Case "FM"
                            itemFM = itemFM + 1
                            itemTEMP = itemFM
                            gno = 1
                        Case "MUROS"
                            grupo = "FM"
                            itemFM = itemFM + 1
                            itemTEMP = itemFM
                            gno = 1
                        Case "FL"
                            itemFL = itemFL + 1
                            itemTEMP = itemFL
                            gno = 3
                        Case "LOSAS"
                            grupo = "FL"
                            itemFL = itemFL + 1
                            itemTEMP = itemFL
                            gno = 3
                        Case "UM"
                            itemUM = itemUM + 1
                            itemTEMP = itemUM
                            gno = 2
                        Case "UNION"
                            grupo = "UM"
                            itemUM = itemUM + 1
                            itemTEMP = itemUM
                            gno = 2
                        Case "CL"
                            itemCL = itemCL + 1
                            itemTEMP = itemCL
                            gno = 4
                        Case "CULAT"
                            grupo = "CL"
                            itemCL = itemCL + 1
                            itemTEMP = itemCL
                            gno = 4
                    End Select

                    If String.IsNullOrEmpty(fila.Cells(0).Value) Then
                        Continue For
                    End If
                    cantitem = fila.Cells(0).Value
                    If fila.Cells(6).Value = "P" Or fila.Cells(6).Value = "p" Then
                        planoesp = 1
                        totalplanos = totalplanos + 1
                        totalesp = totalesp + cantitem
                    Else
                        planoesp = 0
                        totalest = totalest + cantitem
                    End If

                    If String.IsNullOrEmpty(fila.Cells(1).Value) Then
                        Continue For
                    End If
                    tipoitem = fila.Cells(1).Value

                    If fila.Cells(7).Value = Nothing Then
                        desaux = ""
                    Else
                        desaux = fila.Cells(7).Value
                    End If

                    For j = 2 To 5
                        Try
                            If fila.Cells(j).Value = Nothing Then
                                fila.Cells(j).Value = 0
                            End If
                        Catch ex As System.Exception

                        End Try
                    Next

                    an1 = fila.Cells(2).Value
                    al1 = fila.Cells(3).Value
                    al2 = fila.Cells(4).Value
                    an2 = fila.Cells(5).Value
                    medinicial = an1 & "x" & al1 & "x" & al2 & "x" & an2
                    areaitem = fila.Cells(9).Value

                    If areaitem = Nothing Then
                        areaitem = AreaCalc(tipoitem, an1, al1, al2, an2, desaux, grupo) 'area item es un una unidad del item (CDbl(an1) + CDbl(an2)) * (CDbl(al1) + CDbl(al2)) / 10000
                        totalm2 = totalm2 + (areaitem * cantitem)
                        totalvol = totalvol + (totalm2 * 0.054)
                        totalpe = totalpe + (totalm2 * 25)
                    End If

                    cons = "SELECT Id_Piezas FROM Piezas_Forsa WITH (nolock) WHERE ([Desc] = '" & tipoitem & "')"

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

                    Dim id, codigo As String
                    codigo = fila.Cells(12).Value
                    'cons = "INSERT INTO Saldos " +
                    '                    "(Id_EmpCrea, Ing_Metal, Fecrea, Metal, Memo, Pnc_P, Id_Ofa, Est_Pie, Id_Piezas_Forsa, Id_Esp_Prod, Tipo_for, Item, Plano_esp, cant, P_T, Cant_Final_Req, Cant_Desp, Saldo, Tipo, Observacion, Anc1, alto1, Alto2, Anc2, MedidaInicial, Area, Especial, Recibidas, Ofa, Grupo, GNo, completados, Despachadas, Alquiler, Imp_Etiquetas, Eti_Calidad, " +
                    '                    "Peso_Unt, Lote, Explosion, Lamina, sal_Kam_asig, fecha_crea, listado, Pieza_k, Sol, Peso_T, Anula, Id_Tp_K, Asig_Stk, Asig_real, Id_Ofa_P, Dilatacion, Id_Planta, Izquierda, Derecha, " +
                    '                    "Adicionales, Sol_K, Pintadas, Cod_1EE, OA_UnoEE, Sci_1EE, Ept_1EE, Cant_Ept, Procesado, Cod_Planta, Sci_a_Pa, Ept_a_Pa, Peso_Madera, Cant_Acc_Valida, Mensaje_memo, Kam_Entrega, Fabricadas, Explo_Inventor, " +
                    '                    "Peso_Compo, Planta_PlaId, Tmpo_Mtal, Tmpo_SolArm, Tmpo_SolFin, Tmpo_ProTer, Tiene_LM, M2_Est_Erp, Pernado, Categoria, Alerta, IzquDere, TmpoAdic_Mtal, TmpoAdic_Sold, TmpoAdic_Pter, TmpoOper_Metal)"
                    'cons = cons & " VALUES (" & userid & ", 0," & fecha & ", 0, 0, 0," & idofa & ", 0, " & idpiezasforsa & ", 0, 0, " & itemTEMP & ", " & planoesp & ", " & cantitem & ", 0, " & cantitem & ", 0, " & cantitem & ", '" & tipoitem & "', '" & desaux & "', " & an1 & ", " & al1 & ", " & al2 & ", " & an2 & ", '" & medinicial & "', " & areaitem & ", 0, 0, '" & ofa & "', '" & grupo & "', " & gno & ", 0, 0, 0, 0, 0, " +
                    '                    "0, 0, 0, 0, 0, " & fecha & ", 0, 0, 0, 0, 0, 0, 0, 0, " & idofap & ", 0, 0, 0, 0, 0, 0, 0, " & codigo & ", 0, 0, 0, 0, 'Fabricado', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, " & areaitem & ", 0, 0, 0, 0, 0, 0, 0, 0) " +
                    '                    "SELECT cast(SCOPE_IDENTITY() as int) AS 'ID'"
                    cons = "INSERT INTO Saldos " +
                                        "(Id_EmpCrea, Ing_Metal, Fecrea, Metal, Memo, Pnc_P, Id_Ofa, Est_Pie, Id_Piezas_Forsa, Id_Esp_Prod, Tipo_for, Item, Plano_esp, cant, P_T, Cant_Final_Req, Cant_Desp, Saldo, Tipo, Observacion, Anc1, alto1, Alto2, Anc2, MedidaInicial, Area, Especial, Recibidas, Ofa, Grupo, GNo, completados, Despachadas, Alquiler, Imp_Etiquetas, Eti_Calidad, " +
                                        "Peso_Unt, Lote, Explosion, Lamina, sal_Kam_asig, fecha_crea, listado, Pieza_k, Sol, Peso_T, Anula, Id_Tp_K, Asig_Stk, Asig_real, Id_Ofa_P, Dilatacion, Id_Planta, Izquierda, Derecha, " +
                                        "Adicionales, Sol_K, Pintadas, Cod_1EE, OA_UnoEE, Sci_1EE, Ept_1EE, Cant_Ept, Procesado, Cod_Planta, Sci_a_Pa, Ept_a_Pa, Peso_Madera, Cant_Acc_Valida, Mensaje_memo, Kam_Entrega, Fabricadas, Explo_Inventor, " +
                                        "Peso_Compo, Planta_PlaId, Tmpo_Mtal, Tmpo_SolArm, Tmpo_SolFin, Tmpo_ProTer, Tiene_LM, M2_Est_Erp, Pernado, Categoria, Alerta, IzquDere, TmpoAdic_Mtal, TmpoAdic_Sold, TmpoAdic_Pter, TmpoOper_Metal)"
                    cons = cons & " VALUES (" & userid & ", 0," & fecha & ", 0, 0, 0," & idofa & ", 0, " & idpiezasforsa & ", 0, 0, " & itemTEMP & ", " & planoesp & ", " & cantitem & ", 0, " & cantitem & ", 0, " & cantitem & ", '" & tipoitem & "', '" & desaux & "', " & an1 & ", " & al1 & ", " & al2 & ", " & an2 & ", '" & medinicial & "', " & areaitem & ", 0, 0, '" & ofa & "', '" & grupo & "', " & gno & ", 0, 0, 0, 0, 0" +
                                        ", 0, 0, 0, 0, 0, " & fecha & ", 0, 0, 0, 0, 0, 0, 0, 0, " & idofap & ", 0, 0, 0, 0, 0, 0, 0, " & codigo & ", 0, 0, 0, 0, 'Fabricado', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, " & areaitem & ", 0, 0, 0, 0, 0, 0, 0, 0) " +
                                        "SELECT cast(SCOPE_IDENTITY() as int) AS 'ID'"


                    Using connection As New SqlConnection(Str_Con_Bd)

                        Dim command As SqlCommand = New SqlCommand(cons, connection)
                        connection.Open()
                        id = command.ExecuteScalar
                        'command.ExecuteNonQuery()
                    End Using

                    If planoesp = 1 Then
                        Dim insertoplano As String
                        cons = "select Plano_esp from Saldos WITH (nolock) where (Identificador = " & id & ")"
                        Using connection As New SqlConnection(Str_Con_Bd)
                            Dim command As SqlCommand = New SqlCommand(cons, connection)
                            connection.Open()
                            Dim reader As SqlDataReader = command.ExecuteReader()
                            If reader.HasRows Then
                                reader.Read()
                                insertoplano = reader.GetValue(0)
                            End If
                            command = Nothing
                            reader.Close()
                            reader = Nothing
                        End Using
                        If insertoplano <> "True" Then
                            cons = "UPDATE Saldos SET Plano_esp = 1 where (Identificador = " & id & ")"
                            Using connection As New SqlConnection(Str_Con_Bd)
                                Dim command As SqlCommand = New SqlCommand(cons, connection)
                                connection.Open()
                                command.ExecuteNonQuery()
                            End Using
                        End If
                    End If

                    limiteitem = limiteitem + 1


                    progreso.ProgressBar1.Value = limiteitem - 1 'se resta 1 porque limite item empieza en 1, pero la barra inicia en 0
                Next
                Dim consfamilia As String
                If grupo = "FM" Or grupo = "MUROS" Then
                    consfamilia = "Muros"
                ElseIf grupo = "FL" Or grupo = "LOSAS" Then
                    consfamilia = "Losa"
                ElseIf grupo = "UM" Or grupo = "UNION" Then
                    consfamilia = "UML"
                ElseIf grupo = "CL" Or grupo = "CULAT" Then
                    consfamilia = "Culata"
                Else
                    consfamilia = ""
                End If
                If consfamilia <> "" Then


                    cons = "UPDATE Orden SET " & consfamilia & " = 1, Abierta = 1 WHERE (Id_Ofa = " & idofa & ")"

                    Using connection As New SqlConnection(Str_Con_Bd)
                        Dim command As SqlCommand = New SqlCommand(cons, connection)
                        connection.Open()
                        command.ExecuteNonQuery()
                    End Using
                End If

            Next

            totalplanos = cuadros.Rows.Item(4).Cells(1).Value
            totalest = cuadros.Rows.Item(4).Cells(2).Value
            totalesp = cuadros.Rows.Item(4).Cells(3).Value
            totalm2 = cuadros.Rows.Item(4).Cells(4).Value
            totalvol = cuadros.Rows.Item(4).Cells(5).Value
            totalpe = cuadros.Rows.Item(4).Cells(6).Value


            cons = "UPDATE Orden SET Id_Emp_Ing = " & userid & ", Resp_Ingenieria = '" & nombre & "'  ,Planos_Det = " & totalplanos & ", No_Pesp = " & totalesp & ", No_Pest = " & totalest & ", Nu_Piezas = " & totalesp + totalest & ", m2 = " & totalm2 & ", M2_Reales = " & totalm2 & ", vaprx = " & totalvol & ", paprx = " & totalpe & ", Observaciones = '" & obser & "', Ref = '" & ref & "', Abierta = " & 1 & ", Ord_Para_Armado = " & ParaArmado & ", Pernado = " & Pernado & ", Ord_ult_entrega = " & UltimaEntrega & ", Escalera = " & Escalera & " " +
                                "WHERE (Id_Ofa = " & idofa & ")"
            Using connection As New SqlConnection(Str_Con_Bd)
                Dim command As SqlCommand = New SqlCommand(cons, connection)
                connection.Open()
                command.ExecuteNonQuery()
            End Using
            progreso.Close()
            MsgBox("cargado con exito")
        End Sub
        Public Shared Sub verifica2(ByVal Usuario As String, pass As String)

            Dim Str_Con_Bd As String = ConexionBD.getStringConexion()

            Dim Cons As String = "SELECT usuario.usu_login, usuario.usu_passwd, empleado.emp_nombre_mayusculas, empleado.emp_apellidos_mayusculas, usuario.usu_emp_usu_num_id " +
                                "FROM  usuario INNER JOIN " +
                                "empleado ON usuario.usu_emp_usu_num_id = empleado.emp_usu_num_id " +
                                "WHERE (NOT (empleado.emp_nombre_mayusculas = 'Sin nombre')) " +
                                "ORDER BY usuario.usu_login"
            Using connection As New SqlConnection(Str_Con_Bd)
                Dim command As SqlCommand = New SqlCommand(Cons, connection)
                connection.Open()
                Dim reader As SqlDataReader = command.ExecuteReader()
                validez = False
                While reader.Read()
                    Dim a = reader("usu_login").ToString
                    If reader("usu_login").ToString = Usuario Then
                        If reader("usu_passwd").ToString = pass Then
                            validez = True
                            nombre = reader("emp_nombre_mayusculas").ToString & " " & reader("emp_apellidos_mayusculas").ToString
                            userid = reader("usu_emp_usu_num_id").ToString
                            reader.Close()
                            reader = Nothing
                            'Me.Hide()
                            Exit While
                        End If
                    End If
                End While
                If validez = True Then

                Else
                    MsgBox("usuario o clave no es correcta")
                    'TextBox2.Focus()
                End If
            End Using
        End Sub
        Public Shared Function PesoCalc(ByVal desc As String, ByVal grupo As String, area As Double, ByVal Alt1 As Double) As String
            'area de una unidad para producir peso de una unidad

            If desc.Length < 4 And desc.Contains("AG") Then
                PesoCalc = Alt1 * 0.00114 ' (10.55kg/482cm)
            ElseIf desc.Substring(0, 1) = "M" Then
                PesoCalc = area * 25
            ElseIf desc.Substring(0, 1) = "P" Then
                PesoCalc = area * 15
            Else
                PesoCalc = area * 21
            End If

        End Function
        Public Shared Function VolCalc(ByVal desc As String, ByVal Area As Double, ByVal grupo As String, ByVal Alt1 As Double) As String
            'area de una unidad para producir volumen de una unidad


            If desc.Length < 4 And desc.Contains("AG") Then 'desc.Substring(0, 2) = "AG" Or desc.Substring(1, 2) = "AG" Then 'puede fallar cuando a descripcion tenga solo dos digitos
                VolCalc = Alt1 * 0.00003
            ElseIf desc.Substring(0, 1) = "M" Then
                VolCalc = Area * 0.064
            ElseIf desc.Substring(0, 1) = "P" Then
                VolCalc = Area * 0.07
            ElseIf desc.Substring(0, 1) = "#" Then
                VolCalc = Area * 0.08
            Else
                VolCalc = Area * 0.054
            End If



        End Function
        Public Shared Function AreaCalc(ByVal desc As String, ByVal ancho1 As Double, ByVal alto1 As Double, ByVal alto20 As Double, ByVal ancho20 As Double, ByVal desmua As String, ByVal familia As String) As String
            Dim lr_Eqm, lr_Eqmt, lr_Angta, lr_Angex, lr_Angei, lr_Angeit, lr_Angin, lr_Ag, lr_Tn As Object
            Dim lr_Uni, lr_Culang, lr_Culesq, lr_Culata, lr_Culaton As Object
            Dim espm0, espuml, espan, posc1 As Double
            Dim alto2, ancho2 As Double

            alto2 = Val(alto20)
            ancho2 = Val(ancho20)

            espan = 5.4 'Espesor de formaleta
            espuml = 10 'Espesor horizontal de unión

            If familia = "UNION" Then
                If InStr(desc, "CU") > 0 Then
                    posc1 = InStr(desc, "CU")
                    espm0 = Val(Mid(desc, posc1 + 2))
                    If espm0 > 0 Then
                        espuml = espm0
                        desmua = ""
                    End If
                End If
                'Si no encontró valor en la descripción principal, lo busca en la Auxiliar.
                While InStr(desmua, "CC") > 0
                    posc1 = InStr(desmua, "CC")
                    espm0 = Val(Mid(desmua, posc1 + 2))
                    If espm0 > 0 Then
                        espuml = espm0
                        desmua = ""
                    Else
                        desmua = Mid(desmua, posc1 + 2)
                    End If
                End While
            End If

            lr_Eqm = {"EQM-AL", "EQME-AL", "EQMF-AL", "EQM", "EQM+2AGH", "EQM-5P", "EQM-DV", "EQMD", "EQMD-5P", "EQMDF", "EQME", "EQMF", "EQMR", "EQMR-5P", "EQMR-DV", "EQMRF", "EQMBL", "EQMBL+DS"}
            lr_Eqmt = {"EQMT-AL", "EQMTE-AL", "EQMTD-AL", "EQMTDF-AL", "EQMTF-AL", "EQMT", "EQMT-CD", "EQMT-5P", "EQMTE", "EQMTD", "EQMTD-5P", "EQMTDF", "EQMTF", "EQMTT"}

            lr_Angta = {"DTA", "FTA", "TALA"} 'Angulares que tienen TA en descripción
            lr_Angex = {"CPAEX", "CPATEX", "EQLDE", "EQLE", "EQLED", "EQCDE", "EQCE"}
            lr_Angei = {"CPA", "CPA-PCC", "CPABL+DS", "CPAD", "CPAE", "CPAED", "FMA", "FMA+2AGH", "FMA-5P", "FMAF"}
            lr_Angeit = {"CPAT", "CPATD", "CPATF", "CPATDF"}
            lr_Angin = {"EQLZ", "EQCZ", "EQLDI", "EQLI", "EQLID", "EQLW", "EQCDI", "EQCI", "EQCW"}
            lr_Ag = {"AG", "AGE", "AGP", "AGR", "AGRD", "AGRD-E", "AGRD-I", "ES"}

            lr_Tn = {"TNV", "TNVD", "TNH", "TNH2.8", "TNH6.9", "TNH11", "TNHP"}

            lr_Uni = {"EQC", "EQC-AL", "EQC-DG", "EQC-PCC", "EQC-2PCC", "EQCD", "EQL", "EQL-AL", "EQL-DG", "EQL-PCC", "EQL-2PCC", "EQL15", "EQLD", "EQLP", "CU", "CU12.5", "CU15", "CU-DG"}

            lr_Culang = {"CLA", "CPABL", "CPABL+DS", "CPAC", "CPAC-5P", "CPACF", "CPATC", "CPATC-5P", "CPATCF", "CPPA"}
            lr_Culesq = {"EQMC", "EQMC-5P", "EQMFC", "EQMTC", "EQMTC-5P", "EQMTCF"}
            lr_Culata = {"CL", "CL-5P", "CLF", "CLC", "CLC-5P", "CLCF"}
            lr_Culaton = {"CLT", "CLT-5P", "CLTF", "CLTC", "CLTC-5P", "CLTCF"}

            Select Case Left(desc, 1) 'Truncar el texto descriptivo para tipo de formaleta
                Case "#"
                    If Left(desc, 2) = "#4" Then
                        desc = Mid(desc, 3)
                        espan = 6.5
                        espuml = 15
                    Else
                        desc = Mid(desc, 2)
                        espan = 5.4
                        espuml = 10
                    End If
                Case "M"
                    desc = Mid(desc, 2)
                    espan = 6.4
                    espuml = 15
                Case "P"
                    desc = Mid(desc, 2)
                    espan = 7
                    espuml = 10
                Case "Q"
                    desc = Mid(desc, 2)
                    espan = 6.5
                    espuml = 15
                Case "X"
                    desc = Mid(desc, 2)
                    espan = 5
                    espuml = 5
            End Select

            AreaCalc = vbEmpty
            If IsInArray(desc, lr_Eqm) Then
                AreaCalc = (ancho1 * 2) * alto1 / 10000 'Para que pasa de cm² a m²
            End If

            If AreaCalc = vbEmpty Then
                If IsInArray(desc, lr_Eqmt) Then
                    AreaCalc = ((ancho1 * 2) * (alto1 + espuml) - (espuml * espuml)) / 10000 'Para que pasa de cm² a m²
                End If
            End If
            If AreaCalc = vbEmpty Then
                If IsInArray(desc, lr_Angta) Then
                    If ancho20 = "" Then ancho2 = ancho1
                    AreaCalc = (ancho1 + ancho2) * alto1 / 10000 'Para que pasa de cm² a m²
                End If
            End If
            If AreaCalc = vbEmpty Then
                If IsInArray(desc, lr_Angex) Then
                    Try
                        If ancho20 = "" Then ancho2 = ancho1
                    Catch ex As System.Exception
                        If ancho20 = 0 Then ancho2 = ancho1
                    End Try
                    If desc = "CPAEX" Then
                        AreaCalc = ((ancho1 - espan) + (ancho2 - espan)) * alto1 / 10000 'Para que pasa de cm² a m²
                    Else
                        AreaCalc = ((ancho1 - espuml) + (ancho2 - espuml)) * alto1  'Para que pasa de cm² a m²
                        AreaCalc = (AreaCalc + (ancho1 - espuml + ancho2) * espuml) / 10000
                    End If
                End If
            End If
            If AreaCalc = vbEmpty Then
                If IsInArray(desc, lr_Angin) Then
                    Try
                        If ancho20 = "" Then ancho2 = ancho1
                    Catch ex As System.Exception
                        If ancho20 = 0 Then ancho2 = ancho1
                    End Try
                    AreaCalc = (((ancho1 + ancho2) * (alto1 + espuml)) - (espuml * espuml)) / 10000 'Para que pasa de cm² a m²
                End If
            End If
            If AreaCalc = vbEmpty Then
                If IsInArray(desc, lr_Angei) Then
                    Try
                        If ancho20 = "" Then ancho2 = ancho1
                    Catch ex As System.Exception
                        If ancho20 = 0 Then ancho2 = ancho1
                    End Try
                    AreaCalc = (ancho1 + ancho2) * alto1 / 10000 'Para que pasa de cm² a m²
                End If
            End If
            If AreaCalc = vbEmpty Then
                If IsInArray(desc, lr_Angeit) Then
                    Try


                        If ancho20 = "" Then ancho2 = ancho1
                    Catch ex As System.Exception
                        If ancho20 = 0 Then ancho2 = ancho1
                    End Try
                    AreaCalc = ((ancho1 + ancho2) * (alto1 + espuml) - (espuml * espuml)) / 10000 'Para que pasa de cm² a m²
                End If
            End If
            If AreaCalc = vbEmpty Then
                If IsInArray(desc, lr_Uni) Then
                    AreaCalc = alto1 * (ancho1 + espuml) / 10000 'Para que pasa de cm² a m²
                End If
            End If
            If AreaCalc = vbEmpty Then
                If IsInArray(desc, lr_Ag) Then
                    AreaCalc = 0
                End If
            End If
            If AreaCalc = vbEmpty Then
                If IsInArray(desc, lr_Culata) Then
                    If alto20 = "" Then alto2 = alto1
                    AreaCalc = (alto1 + alto2) * ancho1 / 20000 'Para que pasa de cm² a m²
                End If
            End If
            If AreaCalc = vbEmpty Then
                If IsInArray(desc, lr_Culaton) Then
                    If alto20 = "" Then alto2 = alto1
                    AreaCalc = ((alto1 + alto2) / 2 + espuml) * ancho1 / 10000 'Para que pasa de cm² a m²
                End If
            End If
            If AreaCalc = vbEmpty Then
                If IsInArray(desc, lr_Tn) Then
                    Try
                        If ancho20 = "" Then ancho2 = ancho1
                    Catch ex As System.Exception
                        If ancho20 = 0 Then ancho2 = ancho1
                    End Try
                    AreaCalc = alto1 * (ancho1 + ancho2 + ancho2) / 10000
                End If
            End If
            If AreaCalc = vbEmpty And IsInArray(desc, lr_Ag) = False Then 'no calcular para lo que está en AG
                AreaCalc = alto1 * ancho1 / 10000 'Para que pasa de cm² a m²
            End If

        End Function
        Public Shared Function IsInArray(desi As String, arr As Object) As Boolean
            IsInArray = (UBound(Filter(arr, desi)) > -1)
            If IsInArray = True Then
                IsInArray = False
                For Each element In arr
                    If element = desi Then
                        IsInArray = True
                    End If
                Next element
            End If
        End Function
        Public Shared Sub separalistados(ByVal grillas() As Object, rutalibro As String)
            Dim filas As System.Windows.Forms.DataGridViewRowCollection
            Dim fila As System.Windows.Forms.DataGridViewRow
            Dim celdas As System.Windows.Forms.DataGridViewCellCollection

            Dim xlApp = New Microsoft.Office.Interop.Excel.Application
            Dim xlWorkBook = xlApp.Workbooks.Open(rutalibro)
            Dim xlWorkSheet = xlWorkBook.Worksheets.Add()
            xlWorkSheet.name = "Sin codigo ERP"
            'xlWorkSheet = xlWorkBook.Worksheets.Item(xlWorkBook.Worksheets.Count)
            Dim numfila As Integer = 1
            For k = 0 To UBound(grillas)
                Select Case k
                    Case 0
                        xlWorkSheet.cells(numfila, 1) = "MUROS" 'sw.WriteLine("MUROS")
                        numfila = numfila + 1
                    Case 1
                        'sw.WriteLine("")
                        'sw.WriteLine("UNION")
                        xlWorkSheet.cells(numfila, 1) = "UNION"
                        numfila = numfila + 1
                    Case 2
                        'sw.WriteLine("")
                        'sw.WriteLine("LOSAS")
                        xlWorkSheet.cells(numfila, 1) = "LOSAS"
                        numfila = numfila + 1
                End Select
                filas = grillas(k).rows
                For i = 0 To filas.Count - 1
                    Try
                        fila = filas.Item(i)
                    Catch ex As System.Exception
                        Exit For
                    End Try
                    celdas = fila.Cells
                    If fila.DefaultCellStyle.BackColor = System.Drawing.Color.Red Then
                        'escribir = ""
                        For j = 0 To 10
                            xlWorkSheet.cells(numfila, j + 1) = celdas.Item(j).Value
                            'escribir = escribir & celdas.Item(j).Value & vbTab
                        Next
                        'escribir = escribir & celdas.Item(10).Value
                        'sw.WriteLine(escribir)
                        filas.Remove(fila)
                        i = i - 1
                        numfila = numfila + 1
                    End If
                Next
                numfila = numfila + 1
            Next
            xlWorkBook.Close(True)
            xlApp.Quit()
            'Dim xlApp = New Microsoft.Office.Interop.Excel.Application
            'Dim xlWorkBook = xlApp.Workbooks.Add()
            'Dim xlWorkSheet = xlWorkBook.Sheets("sheet1")
        End Sub
        Public Shared Sub GuardaLista(ByVal grillas() As Object, rutalibro As String)
            Dim filas As System.Windows.Forms.DataGridViewRowCollection
            Dim fila As System.Windows.Forms.DataGridViewRow
            Dim celdas As System.Windows.Forms.DataGridViewCellCollection
            Dim xlApp = New Microsoft.Office.Interop.Excel.Application
            Dim xlWorkBook = xlApp.Workbooks.Open(rutalibro)
            Dim xlWorkSheet As Microsoft.Office.Interop.Excel.Worksheet
            Dim existe = False
            For Each worksheet In xlWorkBook.Sheets
                If worksheet.name = "Final" Then
                    worksheet.cells.clearcontents()
                    existe = True
                    xlWorkSheet = worksheet
                    Exit For
                End If
            Next
            If existe = False Then
                xlWorkSheet = xlWorkBook.Worksheets.Add()
                xlWorkSheet.Name = "Final"
            End If


            'xlWorkSheet = xlWorkBook.Worksheets.Item(xlWorkBook.Worksheets.Count)
            Dim numfila As Integer = 1
            For k = 0 To UBound(grillas)
                Select Case k
                    Case 0
                        xlWorkSheet.Cells(numfila, 1) = "MUROS" 'sw.WriteLine("MUROS")
                        numfila = numfila + 1
                    Case 1
                        'sw.WriteLine("")
                        'sw.WriteLine("UNION")
                        xlWorkSheet.Cells(numfila, 1) = "UNION"
                        numfila = numfila + 1
                    Case 2
                        'sw.WriteLine("")
                        'sw.WriteLine("LOSAS")
                        xlWorkSheet.Cells(numfila, 1) = "LOSAS"
                        numfila = numfila + 1
                End Select
                filas = grillas(k).rows
                For i = 0 To filas.Count - 1
                    Try
                        fila = filas.Item(i)
                    Catch ex As System.Exception
                        Exit For
                    End Try
                    celdas = fila.Cells
                    'escribir = ""
                    For j = 0 To 10
                        xlWorkSheet.Cells(numfila, j + 1) = celdas.Item(j).Value
                        'escribir = escribir & celdas.Item(j).Value & vbTab
                    Next
                    'escribir = escribir & celdas.Item(10).Value
                    'sw.WriteLine(escribir)
                    numfila = numfila + 1
                Next

            Next
            xlWorkBook.Close(True)
            xlApp.Quit()
        End Sub
        Public Shared Sub SeparaMarcadas(ByVal grillas() As Object, rutalibro As String)
            Dim filas As System.Windows.Forms.DataGridViewRowCollection
            Dim fila As System.Windows.Forms.DataGridViewRow
            Dim celdas As System.Windows.Forms.DataGridViewCellCollection
            Dim xlApp = New Microsoft.Office.Interop.Excel.Application
            Dim xlWorkBook = xlApp.Workbooks.Open(rutalibro)
            Dim xlWorkSheet = xlWorkBook.Worksheets.Add()
            xlWorkSheet.name = "Separadas"
            'xlWorkSheet = xlWorkBook.Worksheets.Item(xlWorkBook.Worksheets.Count)
            Dim numfila As Integer = 1
            For k = 0 To UBound(grillas)
                Select Case k
                    Case 0
                        xlWorkSheet.cells(numfila, 1) = "MUROS" 'sw.WriteLine("MUROS")
                        numfila = numfila + 1
                    Case 1
                        'sw.WriteLine("")
                        'sw.WriteLine("UNION")
                        xlWorkSheet.cells(numfila, 1) = "UNION"
                        numfila = numfila + 1
                    Case 2
                        'sw.WriteLine("")
                        'sw.WriteLine("LOSAS")
                        xlWorkSheet.cells(numfila, 1) = "LOSAS"
                        numfila = numfila + 1
                End Select
                filas = grillas(k).rows
                For i = 0 To filas.Count - 1
                    Try
                        fila = filas.Item(i)
                    Catch ex As System.Exception
                        Exit For
                    End Try
                    celdas = fila.Cells
                    If celdas.Item(17).Value = True Then
                        'escribir = ""
                        For j = 0 To 10
                            xlWorkSheet.cells(numfila, j + 1) = celdas.Item(j).Value
                            'escribir = escribir & celdas.Item(j).Value & vbTab
                        Next
                        'escribir = escribir & celdas.Item(10).Value
                        'sw.WriteLine(escribir)
                        filas.Remove(fila)
                        i = i - 1
                        numfila = numfila + 1
                    End If
                Next
                numfila = numfila + 1
            Next
            xlWorkBook.Close(True)
            xlApp.Quit()
        End Sub
        Public Shared Sub verificaitems(ByVal consulta As Boolean, separa As Boolean, grillas() As Object, tabcontrol1 As Windows.Forms.TabControl, Optional ByVal verificainventario As Boolean = False, Optional ByVal tablaitemsverifica As Object = Nothing, Optional ByVal bodegausada As String = "", Optional ByVal libroexcel As String = Nothing, Optional CIA As String = "")
            Dim validadoM As Boolean = True
            Dim validadoU As Boolean = True
            Dim validadoL As Boolean = True
            Dim tablaitems(,) As Object = MyCommands.tablaitems(CIA)
            Dim filas As System.Windows.Forms.DataGridViewRowCollection
            Dim fila As System.Windows.Forms.DataGridViewRow
            For j = 0 To UBound(grillas)
                filas = grillas(j).Rows
                For i = 0 To filas.Count - 1
                    fila = filas.Item(i)
                    Dim temp1 As String = nomenclatura(fila.Cells(1).Value, fila.Cells(2).Value, fila.Cells(3).Value)
                    Dim temp2 As String = atributos(temp1, fila.Cells(2).Value, fila.Cells(3).Value, fila.Cells(4).Value)
                    fila.Cells(11).Value = temp1 & " " & temp2
                    fila.Cells(12).Value = ConsultaCodigoItem(fila.Cells(11).Value, tablaitems)
                    If fila.Cells(12).Value = Nothing Then
                        fila.DefaultCellStyle.BackColor = System.Drawing.Color.Red
                        Select Case j
                            Case 0
                                validadoM = False
                            Case 1
                                validadoU = False
                            Case 2
                                validadoL = False
                        End Select
                    Else
                        fila.DefaultCellStyle.BackColor = System.Drawing.Color.White
                    End If
                Next
            Next
            If consulta = True Or verificainventario = True Then
                If validadoM = False And validadoU = False And validadoL = False Then
                    tabcontrol1.SelectTab(0)
                    Frm_CargaListados.mensa = MsgBox("Algunos Ítems de Muro, Union y Losa no existen en el ERP", MsgBoxStyle.Exclamation)
                    Exit Sub
                ElseIf validadoU = False And validadoL = False Then
                    tabcontrol1.SelectTab(1)
                    Frm_CargaListados.mensa = MsgBox("Algunos Ítems de Union y Losa no existen en el ERP", MsgBoxStyle.Exclamation)
                    Exit Sub
                ElseIf validadoM = False And validadoL = False Then
                    tabcontrol1.SelectTab(0)
                    Frm_CargaListados.mensa = MsgBox("Algunos Ítems de Muro y Losa no existen en el ERP", MsgBoxStyle.Exclamation)
                    Exit Sub
                ElseIf validadoM = False And validadoU = False Then
                    tabcontrol1.SelectTab(0)
                    Frm_CargaListados.mensa = MsgBox("Algunos Ítems de Muro y Union no existen en el ERP", MsgBoxStyle.Exclamation)
                    Exit Sub
                ElseIf validadoM = False Then
                    tabcontrol1.SelectTab(0)
                    Frm_CargaListados.mensa = MsgBox("Algunos Ítems de Muro no existen en el ERP", MsgBoxStyle.Exclamation)
                    Exit Sub
                ElseIf validadoU = False Then
                    tabcontrol1.SelectTab(1)
                    Frm_CargaListados.mensa = MsgBox("Algunos Ítems de Union no existen en el ERP", MsgBoxStyle.Exclamation)
                    Exit Sub
                ElseIf validadoL = False Then
                    tabcontrol1.SelectTab(2)
                    Frm_CargaListados.mensa = MsgBox("Algunos Ítems de Losa no existen en el ERP", MsgBoxStyle.Exclamation)
                    Exit Sub
                Else
                    If verificainventario = False Then
                        MsgBox("todos los Ítems existen en el ERP", MsgBoxStyle.Information)
                    End If

                End If

            End If

            If separa = True Then
                separalistados(grillas, libroexcel)
            End If
            If verificainventario = True Then
                bodegaqueseusa = bodegausada
                verinventario(grillas, tablaitemsverifica)
            End If
            If consulta = True And verificainventario = False Then
                tablaforlinemuros = grillas(0)
                tablaForlineUnion = grillas(1)
                tablaForlineLosas = grillas(2)
            End If
        End Sub

        Public Shared Sub verinventario(ByVal grillas() As Object, tablaitems As Object)
            Dim BodegaI, quedan, faltan As Integer
            Dim encontrado As Boolean
            Dim filas As System.Windows.Forms.DataGridViewRowCollection
            Dim celdas As System.Windows.Forms.DataGridViewCellCollection
            For Each grilla In grillas
                filas = grilla.rows
                For i = 0 To filas.Count - 1
                    BodegaI = 0
                    celdas = filas.Item(i).Cells
                    celdas.Item(13).Value = 0
                    celdas.Item(14).Value = 0
                    celdas.Item(15).Value = 0
                    quedan = 0
                    encontrado = False
                    faltan = celdas.Item(16).Value
                    For j = 0 To UBound(tablaitems, 2)
                        If tablaitems(2, j) = celdas.Item(12).Value Then
                            quedan = tablaitems(0, j) - celdas.Item(16).Value
                            faltan = celdas.Item(16).Value - tablaitems(0, j)
                            BodegaI = tablaitems(0, j)
                            If quedan > 0 Then
                                tablaitems(0, j) = quedan
                            Else
                                tablaitems(0, j) = 0
                            End If
                            encontrado = True
                            Exit For
                        End If
                    Next
                    If encontrado = False Then
                        celdas.Item(13).Value = 0
                        celdas.Item(14).Value = 0
                        celdas.Item(15).Value = 0
                    Else
                        celdas.Item(13).Value = BodegaI
                        If quedan < 0 Then
                            celdas.Item(14).Value = BodegaI
                            celdas.Item(15).Value = 0
                            celdas.Item(16).Value = faltan
                        Else
                            celdas.Item(14).Value = celdas.Item(16).Value
                            celdas.Item(15).Value = quedan
                            celdas.Item(16).Value = 0
                        End If
                    End If

                Next
            Next
        End Sub
        Public Shared Sub separaUML(ByVal datagridview2 As Windows.Forms.DataGridView, rutalibro As String)
            Dim filas As System.Windows.Forms.DataGridViewRowCollection
            Dim fila As System.Windows.Forms.DataGridViewRow
            Dim celdas As System.Windows.Forms.DataGridViewCellCollection
            Dim xlApp = New Microsoft.Office.Interop.Excel.Application
            Dim xlWorkBook = xlApp.Workbooks.Open(rutalibro)
            Dim xlWorkSheet = xlWorkBook.Worksheets.Add()
            xlWorkSheet.name = "UML Separada"
            filas = datagridview2.Rows
            Dim numfila As Integer = 1
            For i = 0 To filas.Count - 1
                Try
                    fila = filas.Item(i)
                Catch ex As System.Exception
                    Exit For
                End Try
                celdas = fila.Cells
                For j = 0 To 10
                    xlWorkSheet.cells(numfila, j + 1) = celdas.Item(j).Value
                    'escribir = escribir & celdas.Item(j).Value & vbTab
                Next
                'escribir = escribir & celdas.Item(10).Value
                'sw.WriteLine(escribir)
                filas.Remove(fila)
                i = i - 1
                numfila = numfila + 1
            Next
            xlWorkBook.Close(True)
            xlApp.Quit()
        End Sub
        Public Shared Sub separautilizadas(ByVal grillas() As Object, rutalibro As String)
            Dim filas As System.Windows.Forms.DataGridViewRowCollection
            Dim fila As System.Windows.Forms.DataGridViewRow
            Dim celdas As System.Windows.Forms.DataGridViewCellCollection
            Dim xlApp = New Microsoft.Office.Interop.Excel.Application
            Dim xlWorkBook = xlApp.Workbooks.Open(rutalibro)
            Dim xlWorkSheet = xlWorkBook.Worksheets.Add()
            xlWorkSheet.name = "Usadas Bodega " & bodegaqueseusa
            'filename = filename & "usadas bodega " & bodegaqueseusa & ".xls"
            Dim numfila As Integer = 1
            For k = 0 To UBound(grillas)
                Select Case k
                    Case 0
                        xlWorkSheet.cells(numfila, 1) = "MUROS" 'sw.WriteLine("MUROS")
                        numfila = numfila + 1
                    Case 1
                        'sw.WriteLine("")
                        'sw.WriteLine("UNION")
                        xlWorkSheet.cells(numfila, 1) = "UNION"
                        numfila = numfila + 1
                    Case 2
                        'sw.WriteLine("")
                        'sw.WriteLine("LOSAS")
                        xlWorkSheet.cells(numfila, 1) = "LOSAS"
                        numfila = numfila + 1
                End Select
                filas = grillas(k).rows
                For i = 0 To filas.Count - 1
                    fila = filas.Item(i)
                    celdas = fila.Cells
                    If celdas.Item(14).Value > 0 And fila.Visible = True Then
                        'escribir = ""
                        'escribir = escribir & celdas.Item(14).Value & vbTab
                        xlWorkSheet.cells(numfila, 1) = celdas.Item(14).Value
                        For j = 1 To 10
                            xlWorkSheet.cells(numfila, j + 1) = celdas.Item(j).Value
                            'escribir = escribir & celdas.Item(j).Value & vbTab
                        Next
                        'escribir = escribir & celdas.Item(10).Value
                        'sw.WriteLine(escribir)
                        numfila = numfila + 1
                    End If
                    If celdas.Item(16).Value = 0 Then
                        fila.Visible = False
                    End If
                Next
            Next
            xlWorkBook.Close(True)
            xlApp.Quit()
        End Sub

        Public Shared Function SeparaCaracteres(ByVal texto As String) As Object
            Dim ascii As Encoding = Encoding.ASCII
            Dim unicode As Encoding = Encoding.Unicode
            Dim unicodeBytes As Byte() = unicode.GetBytes(texto)
            Dim asciiBytes As Byte() = Encoding.Convert(unicode, ascii, unicodeBytes)
            Dim asciiChars(ascii.GetCharCount(asciiBytes, 0, asciiBytes.Length) - 1) As Char
            ascii.GetChars(asciiBytes, 0, asciiBytes.Length, asciiChars, 0)
            Dim separados() As String
            ReDim separados(asciiChars.Length - 1)
            'SeparaCaracteres = asciiChars
            For i = 0 To UBound(asciiChars)
                separados(i) = Convert.ToString(asciiChars(i))
            Next
            SeparaCaracteres = separados
        End Function

        Public Shared Sub CreaPlano()
            Dim encabezado, inicial, final As String
            Dim Tipo_Cot() As String
            Dim vacio As String
            Dim Tipo_Reg, Subtipo_Reg, Version_Reg, Consec_Auto_Reg, Id_Tipo_Docto, Consec_Docto As String ', Cia, Id_Co, Fecha, Id_Solicitante, Num_Dias_Entrega, Fecha_Entrega, notas, Bodega_Salida, Bodega_Entrada
            Dim Id_Tercero, Id_Clase_Docto, Ind_Estado, Ind_Impresion As String
            Dim Concepto As String
            Dim Referencia_item, Barras_Detalle1_Detalle2, Motivo, Campo As String
            Dim cCosto_Movto, Id_Proyecto, Desc_Variable, Id_Un_Movto As String

            vacio = ""

            'GoTo otro

            Tipo_Reg = "0440"
            Subtipo_Reg = "00"
            Version_Reg = "02"
            'Cia = Selcia(True) '"002" '
            Consec_Auto_Reg = "1"
            'Id_Co = SelCO(Cia) '"001" '
            Id_Tipo_Docto = "RQI"
            Consec_Docto = "00000000"

            'Id_Tercero = SelTercero(Cia, Id_Co) '"123456789      " '
            'Id_Solicitante = SelTercero(Cia, Id_Co, Id_Tercero) '"CLV  " '
            'Num_Dias_Entrega = diasEntrega '"123"'
            'Fecha_Entrega = diasEntrega(Num_Dias_Entrega, Fecha)


            Id_Clase_Docto = "075"
            Ind_Estado = "0"
            Ind_Impresion = "0"
            'Notas = selNotas '"a b c d e f g h i j k l m n o p q r s t u v w x y z 1 2 3 4 5 6 7 8 70a b c d e f g h i j k l m n o p q r s t u v w x y z 1 2 3 4 5 6 7 8140a b c d e f g h i j k l m n o p q r s t u v w x y z 1 2 3 4 5 6 7 8210a b c d e f g h i j k l m n o p q r s t u 255" '
            Concepto = "607"
            'Bodega_Salida = SelBodega(Cia, Id_Co, True) '"001  " '
            'Bodega_Entrada = SelBodega(Cia, Id_Co, False) '"001-1" '
            'Referencia = selRef '"ESTANDAR            " '
            'Ubicacion = selUbica(Cia, Bodega_Entrada) '"ARQ001    " '
            inicial = "00000001" & Frm_GenerarRQI.CIA 'consecutivo '
            encabezado = Tipo_Reg & Subtipo_Reg & Version_Reg & Frm_GenerarRQI.CIA & Consec_Auto_Reg & Frm_GenerarRQI.Id_Co & Id_Tipo_Docto & Consec_Docto
            encabezado = encabezado & Frm_GenerarRQI.fecha & "               " & Frm_GenerarRQI.Id_Solicitante & Frm_GenerarRQI.Fecha_Entrega & Frm_GenerarRQI.Num_Dias_Entrega & Id_Clase_Docto & Ind_Estado & Ind_Impresion & Frm_GenerarRQI.notas
            encabezado = encabezado & Concepto & Frm_GenerarRQI.Bodega_Salida & Frm_GenerarRQI.Bodega_Entrada
            final = "99990001" & Frm_GenerarRQI.CIA 'consecutivo

            'Referencia = "               MUROS" '"            ESTANDAR"
            Tipo_Cot = Split(Frm_GenerarRQI.Referencia, " ")

            Dim tablafinal(,) As Object
            Dim tablafinalMuros(,) As Object
            Dim tablafinalLosas(,) As Object
            ReDim tablafinalMuros(1, 0)
            ReDim tablafinalLosas(1, 0)
            ReDim tablafinal(1, 0)
            'Dim sumar As Boolean
            'Dim sumado As Boolean
            'Dim item As Integer
            If Tipo_Cot(0) = "ESTANDAR" Then
                tablafinalMuros = tablafinalfuncion(tablaforlinemuros)
                tablafinalLosas = tablafinalfuncion(tablaForlineUnion)
                tablafinalLosas = tablafinalfuncion(tablaForlineLosas, tablafinalLosas)
            ElseIf Tipo_Cot(0) = "MUROS" Or Tipo_Cot(0) = "LOSAS" Then
                tablafinal = tablafinalfuncion(tablaforlinemuros)
                tablafinal = tablafinalfuncion(tablaForlineUnion, tablafinal)
                tablafinal = tablafinalfuncion(tablaForlineLosas, tablafinal)
            End If
            Tipo_Reg = "0441"
            Referencia_item = "                                                  "
            Barras_Detalle1_Detalle2 = "                                                            "
            Motivo = "01"
            Campo = "00"
            cCosto_Movto = "               "
            Id_Proyecto = "               "
            Desc_Variable = " "

            For i = 1 To 1999
                Desc_Variable = Desc_Variable & " "
            Next
            'otro:
            Id_Un_Movto = "001                 "
            'pq = Len(Desc_Variable
            Try
                RutaDwg()
            Catch ex As System.Exception
                MsgBox("Debe guardar el plano modulado", MsgBoxStyle.Exclamation)
                Exit Sub
            End Try



            'filename = filename & "InventarioERP.xls"

            usupassw()
            If sigue = True Then
                If Tipo_Cot(0) = "ESTANDAR" Then
                    Numero_Reg = "0000001"
                    escribe(nombre & " MUROS", tablafinalMuros, ruta, inicial, encabezado, final, Tipo_Reg, Subtipo_Reg, Version_Reg, Id_Tipo_Docto, Consec_Docto, Referencia_item, Barras_Detalle1_Detalle2, Concepto, Motivo, Campo, cCosto_Movto, Id_Proyecto, Desc_Variable, Id_Un_Movto)
                    Numero_Reg = "0000001"
                    escribe(nombre & " LOSAS", tablafinalLosas, ruta, inicial, encabezado, final, Tipo_Reg, Subtipo_Reg, Version_Reg, Id_Tipo_Docto, Consec_Docto, Referencia_item, Barras_Detalle1_Detalle2, Concepto, Motivo, Campo, cCosto_Movto, Id_Proyecto, Desc_Variable, Id_Un_Movto)
                ElseIf Tipo_Cot(0) = "MUROS" Or Tipo_Cot(0) = "LOSAS" Then
                    Numero_Reg = "0000001"
                    escribe(nombre & " " & Tipo_Cot(0), tablafinal, ruta, inicial, encabezado, final, Tipo_Reg, Subtipo_Reg, Version_Reg, Id_Tipo_Docto, Consec_Docto, Referencia_item, Barras_Detalle1_Detalle2, Concepto, Motivo, Campo, cCosto_Movto, Id_Proyecto, Desc_Variable, Id_Un_Movto)
                End If
            End If
            'NomUsuario = InputBox("Ingrese nombre de Usuario de ERP", "Usuario") '"Siif_Unoee"
            'Passw = InputBox("Ingrese Contraseña", "Contraseña") '"siif"



        End Sub
        Public Shared Sub usupassw()
            NomUsuario = ""
            Passw = ""
            Dim usupass As New Frm_Login
            usupass.StartPosition = Windows.Forms.FormStartPosition.CenterScreen
            With usupass
                .ShowDialog()
            End With
        End Sub
        Public Shared Function NumReg(ByVal Numero_Reg As String) As String

            Dim nreg As Integer
            nreg = CInt(Numero_Reg) + 1
            Dim NumVector1(6) As String
            For i = 0 To 6
                NumVector1(i) = "0"
            Next
            Dim NumVector2() As String
            NumVector2 = SeparaCaracteres(nreg) 'Split(StrConv(nreg, vbUnicode), Chr$(0))
            Dim j = NumVector2.Length - 1
            Dim k = 6
            NumReg = "0"
            Do While j >= 0
                NumVector1(k) = NumVector2(j)
                j = j - 1
                k = k - 1
            Loop
            For i = 1 To UBound(NumVector1)
                NumReg = CStr(NumReg) & NumVector1(i)
            Next
        End Function
        Public Shared Function und_Med(ByVal item As String, ByVal CIA As String) As String
            Dim slrs10 As String
            Dim rsiferp2 = New ADODB.Recordset
            slrs10 = "Select " +
                            "T120_MC_ITEMS.F120_ID_UNIDAD_INVENTARIO " +
                        "From " +
                            "T120_MC_ITEMS Inner Join " +
                            "T121_MC_ITEMS_EXTENSIONES " +
                            "On T121_MC_ITEMS_EXTENSIONES.F121_ROWID_ITEM = T120_MC_ITEMS.F120_ROWID " +
                        "Where " +
                            "T120_MC_ITEMS.F120_ID_CIA = " & CIA & " And " +
                            "T121_MC_ITEMS_EXTENSIONES.F121_IND_ESTADO = 1 And " +
                            "T120_MC_ITEMS.F120_IND_TIPO_ITEM = 1 And " +
                            "T120_MC_ITEMS.F120_REFERENCIA = '" & item & "' " +
                        "Order by " +
                            "T120_MC_ITEMS.F120_DESCRIPCION"

            Dim Con_Ora As ADODB.Connection = AbreDb_Orl()
            rsiferp2.Open(slrs10, Con_Ora, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            'Dim TablaUnd() As Variant
            'TablaUnd = rsiferp2.GetRows

            'und_Med = rsiferp2.Fields.item(0).Value
            und_Med = Left(rsiferp2.Fields.Item(0).Value, 3) & " "
            'und_Med = " " & und_Med
        End Function
        Public Shared Function cantBase(ByVal cant As String) As String
            Dim cantVector1(19) As String
            For i = 0 To 19
                If i = 15 Then
                    cantVector1(i) = "."
                Else
                    cantVector1(i) = "0"
                End If
            Next
            Dim cantVector2() As String
            cantVector2 = SeparaCaracteres(cant) 'Split(StrConv(cant, vbUnicode), Chr$(0))
            Dim j = Len(cant) - 1
            Dim k = 14
            cantBase = ""
            Do While j >= 0
                cantVector1(k) = cantVector2(j)
                j = j - 1
                k = k - 1
            Loop
            For i = 0 To UBound(cantVector1)
                cantBase = cantBase & cantVector1(i)
            Next
        End Function
        Public Shared Function cant_M2(ByVal item As String, ByVal CIA As String, ByVal Id_Unidad_Medida As String, ByVal Cant_Base As Integer) As String
            Dim cantidad As Double
            Dim slrs10 As String
            Dim rsiferp2 = New ADODB.Recordset
            slrs10 = "Select " +
                            "T122_MC_ITEMS_UNIDADES.F122_FACTOR " +
                        "From " +
                            "T120_MC_ITEMS Inner Join " +
                            "T121_MC_ITEMS_EXTENSIONES " +
                            "On T121_MC_ITEMS_EXTENSIONES.F121_ROWID_ITEM = T120_MC_ITEMS.F120_ROWID " +
                        "Inner Join " +
                            "T122_MC_ITEMS_UNIDADES " +
                            "On T122_MC_ITEMS_UNIDADES.F122_ROWID_ITEM = T120_MC_ITEMS.F120_ROWID " +
                        "Where " +
                            "T120_MC_ITEMS.F120_ID_CIA = " & CIA & " And " +
                            "T120_MC_ITEMS.F120_REFERENCIA = '" & item & "'"
            Dim Con_Ora As ADODB.Connection = AbreDb_Orl()
            rsiferp2.Open(slrs10, Con_Ora, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            cant_M2 = rsiferp2.Fields.Item(0).Value
            cantidad = Math.Round(Cant_Base / CDbl(cant_M2), 4)

            Dim m2Vector1(19) As String
            For i = 0 To 19
                If i = 15 Then
                    m2Vector1(i) = "."
                Else
                    m2Vector1(i) = "0"
                End If
            Next
            Dim cantVector2() As String

            Dim m2Vector2 = Split(cantidad, ".")
            Dim decimales() As String
            Dim unidades = SeparaCaracteres(m2Vector2(0)) 'Split(StrConv(m2Vector2(0), vbUnicode), Chr$(0))
            On Error Resume Next
            decimales = SeparaCaracteres(m2Vector2(1)) 'Split(StrConv(m2Vector2(1), vbUnicode), Chr$(0))
            Dim modulo = CDbl(cant_M2) - Fix(CDbl(cant_M2))
            If modulo = 0 Then
                ReDim decimales(0)
                decimales(0) = "0"
            End If
            Dim j = UBound(unidades) - 1
            Dim k = 14
            cant_M2 = ""
            Do While j >= 0
                m2Vector1(k) = unidades(j)
                j = j - 1
                k = k - 1
            Loop
            If UBound(decimales) > 3 Then
                ReDim Preserve decimales(3)
            ElseIf UBound(decimales) <= 3 Then
                ReDim Preserve decimales(3)
                For l = 0 To UBound(decimales)
                    If decimales(l) = "" Then
                        decimales(l) = "0"
                    End If
                Next
            End If




            'decimales(2) = "0"
            'decimales(3) = "0"




            j = UBound(decimales)
            k = 19
            Do While j >= 0
                m2Vector1(k) = decimales(j)
                j = j - 1
                k = k - 1
            Loop




            For i = 0 To UBound(m2Vector1)
                cant_M2 = cant_M2 & m2Vector1(i)
            Next
        End Function

        Public Shared Sub escribe(ByVal tipocot As String, tabla As Object, ruta As String, inicial As String, encabezado As String, final As String, Tipo_Reg As String, Subtipo_Reg As String, Version_Reg As String, Id_Tipo_Docto As String, Consec_Docto As String, Referencia_item As String, Barras_Detalle1_Detalle2 As String, Concepto As String, Motivo As String, Campo As String, cCosto_Movto As String, Id_Proyecto As String, Desc_Variable As String, Id_Un_Movto As String)

            Dim pvstrDatos, nro_Registro, Id_Item, Id_Unidad_Medida, Cant_Base, cant_2 As String
            Dim FN As Integer
            Dim filename As String
            FN = FreeFile()
            filename = ruta & " " & tipocot & ".txt"


            'Open filename For Output As #FN


            Using sw As StreamWriter = File.CreateText(filename)
                sw.WriteLine("<?xml version='1.0' encoding='utf-8'?>" & Environment.NewLine & "<Importar>" & Environment.NewLine & "<NombreConexion>Siif_Unoee</NombreConexion>" & Environment.NewLine & "<IdCia>" & CInt(Frm_GenerarRQI.CIA) & "</IdCia>" & Environment.NewLine & "<Usuario>" & NomUsuario & "</Usuario>" & Environment.NewLine & "<Clave>" & Passw & "</Clave>" & Environment.NewLine & "<Datos>" & Environment.NewLine)
                pvstrDatos = "<?xml version='1.0' encoding='utf-8'?>" & Environment.NewLine & "<Importar>" & Environment.NewLine & "<NombreConexion>Siif_Unoee</NombreConexion>" & Environment.NewLine & "<IdCia>" & CInt(Frm_GenerarRQI.CIA) & "</IdCia>" & Environment.NewLine & "<Usuario>" & NomUsuario & "</Usuario>" & Environment.NewLine & "<Clave>" & Passw & "</Clave>" & Environment.NewLine & "<Datos>" & Environment.NewLine
                'Print #FN, Numero_Reg & inicial
                'sw.WriteLine(Numero_Reg & inicial)
                sw.WriteLine("<Linea>" & Numero_Reg & inicial & "</Linea>" & Environment.NewLine)
                pvstrDatos = pvstrDatos & "<Linea>" & Numero_Reg & inicial & "</Linea>" & Environment.NewLine
                'Print #FN, vacio
                Numero_Reg = NumReg(Numero_Reg)
                Frm_GenerarRQI.Referencia = "ESTANDAR MUROS      "
                'Print #FN, Numero_Reg & encabezado & Referencia & Ubicacion
                'sw.WriteLine(Numero_Reg & encabezado & Form3.Referencia & Form3.Ubicacion)
                sw.WriteLine("<Linea>" & Numero_Reg & encabezado & Frm_GenerarRQI.Referencia & Frm_GenerarRQI.Ubicacion & "</Linea>" & Environment.NewLine)
                pvstrDatos = pvstrDatos & "<Linea>" & Numero_Reg & encabezado & Frm_GenerarRQI.Referencia & Frm_GenerarRQI.Ubicacion & "</Linea>" & Environment.NewLine
                'Print #FN, vacio
                For i = 0 To UBound(tabla, 2)
                    Numero_Reg = NumReg(Numero_Reg)
                    nro_Registro = "000" & Numero_Reg
                    Id_Item = tabla(1, i)
                    Id_Unidad_Medida = und_Med(Id_Item, Frm_GenerarRQI.CIA)
                    Cant_Base = cantBase(tabla(0, i))
                    cant_2 = cant_M2(Id_Item, Frm_GenerarRQI.CIA, Id_Unidad_Medida, Cant_Base)
                    Dim plano As String
                    plano = Tipo_Reg & Subtipo_Reg & Version_Reg & Frm_GenerarRQI.CIA & Frm_GenerarRQI.Id_Co & Id_Tipo_Docto & Consec_Docto & nro_Registro & Id_Item
                    plano = plano & Referencia_item & Barras_Detalle1_Detalle2 & Frm_GenerarRQI.Bodega_Salida & Concepto & Motivo & Id_Unidad_Medida
                    plano = plano & Cant_Base & cant_2 & Frm_GenerarRQI.Fecha_Entrega & Frm_GenerarRQI.Num_Dias_Entrega & Frm_GenerarRQI.Id_Co & Campo & cCosto_Movto & Id_Proyecto
                    plano = plano & Frm_GenerarRQI.notas & Desc_Variable & Id_Un_Movto
                    'Print #FN, Numero_Reg & plano
                    'sw.WriteLine(Numero_Reg & plano)
                    sw.WriteLine("<Linea>" & Numero_Reg & plano & "</Linea>" & Environment.NewLine)
                    pvstrDatos = pvstrDatos & "<Linea>" & Numero_Reg & plano & "</Linea>" & Environment.NewLine
                    'Print #FN, vacio
                Next
                Numero_Reg = NumReg(Numero_Reg)
                'Print #FN, Numero_Reg & final
                'sw.WriteLine(Numero_Reg & final)
                sw.WriteLine("<Linea>" & Numero_Reg & final & "</Linea>" & Environment.NewLine & "</Datos>" & Environment.NewLine & "</Importar>")
                pvstrDatos = pvstrDatos & "<Linea>" & Numero_Reg & final & "</Linea>" & Environment.NewLine & "</Datos>" & Environment.NewLine & "</Importar>"
                'Close #FN
            End Using
            'Dim x As Object
            'Dim x = numeroRQI()
            Dim x As Short = CShort(0)
            Dim CallWebService As New WebReference1.WSUNOEE
            'CallWebService.ImportarXML(pvstrDatos, x)

            'Dim resultado = CallWebService.ImportarXML(pvstrDatos, x)
            'If x = 0 Then
            'recuperar número de RQI
            'MsgBox(x(0, 0).ToString)
            'Dim otroresultado = resultado.GetXml
            'MsgBox(x.ToString)
            Frm_GenerarRQI.cargada = False
            If x = 0 Then
                MsgBox("RQI cargada satisfactoriamente con el número " & numeroRQI())
                Frm_GenerarRQI.cargada = True
            ElseIf x = 2 Then
                MsgBox("el nombre de usuario o contraseña no debe tener carácteres especiales")
            ElseIf x = 3 Then
                MsgBox("Nombre de Usuario y/o Contraseña no válido")
            Else
                MsgBox("la RQI no se pudo cargar debido al código de error " & x & Environment.NewLine & "consulte al departamento de informática")
            End If
            'End If
            'MsgBox(x.ToString)
            'Dim sGetValue As String = CallWebService.e

        End Sub
        Public Shared Function numeroRQI() As String
            Dim slrs10 As String
            Dim rsiferp2 = New ADODB.Recordset
            slrs10 = "Select " +
                        "T440_DOCTO_REQ_INT.F440_ID_CIA, " +
                        "T440_DOCTO_REQ_INT.F440_ID_CO, " +
                        "T440_DOCTO_REQ_INT.F440_ID_TIPO_DOCTO, " +
                        "T440_DOCTO_REQ_INT.F440_CONSEC_DOCTO, " +
                        "T200_MM_TERCEROS.F200_RAZON_SOCIAL, " +
                        "T440_DOCTO_REQ_INT.F440_IND_ESTADO, " +
                        "T440_DOCTO_REQ_INT.F440_TS, " +
                        "T155_MC_UBICACION_AUXILIARES.F155_ID, " +
                        "T440_DOCTO_REQ_INT.F440_REFERENCIA, " +
                        "T440_DOCTO_REQ_INT.F440_NOTAS, " +
                        "T440_DOCTO_REQ_INT.F440_USUARIO_CREACION, " +
                        "T211_MM_FUNCIONARIOS.F211_ID, " +
                        "T200_MM_TERCEROS.F200_ROWID " +
                    "From " +
                        "T440_DOCTO_REQ_INT Inner Join " +
                        "T155_MC_UBICACION_AUXILIARES " +
                        "On T440_DOCTO_REQ_INT.F440_ID_UBICACION_ENT = " +
                        "T155_MC_UBICACION_AUXILIARES.F155_ID Inner Join " +
                        "T200_MM_TERCEROS " +
                        "On T155_MC_UBICACION_AUXILIARES.F155_ROWID_TERCERO = " +
                        "T200_MM_TERCEROS.F200_ROWID Inner Join " +
                        "T211_MM_FUNCIONARIOS " +
                        "On T211_MM_FUNCIONARIOS.F211_ROWID_TERCERO = " +
                        "T440_DOCTO_REQ_INT.F440_ROWID_TERCERO_SOL " +
                    "Where " +
                        "T440_DOCTO_REQ_INT.F440_ID_CIA = " & Frm_GenerarRQI.CIA & " And " +
                        "T440_DOCTO_REQ_INT.F440_ID_CO = '" & Frm_GenerarRQI.Id_Co & "' And " +
                        "T440_DOCTO_REQ_INT.F440_ID_TIPO_DOCTO = 'RQI' " +
                    "Order By " +
                        "T440_DOCTO_REQ_INT.F440_CONSEC_DOCTO Desc"
            Dim Con_Ora As ADODB.Connection = AbreDb_Orl()
            rsiferp2.Open(slrs10, Con_Ora, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            Dim tablaRQI = rsiferp2.GetRows
            For i = 0 To 10
                Dim b = tablaRQI(7, i).ToString
                Dim c = Split(Frm_GenerarRQI.Id_Solicitante.ToString, " ")
                Dim d = tablaRQI(12, i).ToString
                If b = Frm_GenerarRQI.Ubicacion Then
                    If tablaRQI(11, i) = c(0) Then
                        If d = Frm_GenerarRQI.Cliente Then
                            numeroRQI = tablaRQI(3, i)
                            Exit For
                        End If
                    End If
                End If
            Next
        End Function
        Public Shared Sub RutaDwg()
            'Dim doc As Document = Application.DocumentManager.MdiActiveDocument
            'Dim hs As HostApplicationServices = HostApplicationServices.Current
            'Dim path As String
            'path = hs.FindFile(doc.Name, doc.Database, FindFileHint.Default)
            'Dim PathVec() As String = Split(path, "\")
            'ruta = ""
            'For i = 0 To UBound(PathVec) - 1
            '    ruta = ruta & PathVec(i) & "\"
            'Next
            'Dim nomvec = Split(PathVec(UBound(PathVec)), ".")
            'nombre = nomvec(0)
        End Sub
        Public Shared Function tablafinalfuncion(ByVal tablaforline As System.Windows.Forms.DataGridView, Optional ByVal tablainicial As Object = Nothing) As Object

            Dim tabladevuelve(,) As Object
            If tablainicial Is Nothing Then
                ReDim tabladevuelve(1, 0)
            Else
                tabladevuelve = tablainicial
                ReDim Preserve tabladevuelve(1, UBound(tabladevuelve, 2) + 1)
            End If
            If tablaforline.RowCount = 0 Then
                Exit Function
            End If
            ReDim tablafinalfuncion(1, 0)

            For i = 0 To tablaforline.RowCount - 1
                Dim sumado = False
                For j = 0 To UBound(tabladevuelve, 2)

                    Dim sumar = False
                    If tablaforline.Rows.Item(i).Cells.Item(12).Value = tabladevuelve(1, j) Then
                        sumar = True
                        'item = j
                    End If
                    If sumar = True Then
                        tabladevuelve(0, j) = CStr(CInt(tabladevuelve(0, j)) + CInt(tablaforline.Rows.Item(i).Cells.Item(0).Value))
                        sumado = True
                    End If
                Next
                If sumado = False Then

                    tabladevuelve(0, UBound(tabladevuelve, 2)) = tablaforline.Rows.Item(i).Cells.Item(0).Value.ToString
                    tabladevuelve(1, UBound(tabladevuelve, 2)) = tablaforline.Rows.Item(i).Cells.Item(12).Value
                    ReDim Preserve tabladevuelve(1, UBound(tabladevuelve, 2) + 1)
                End If
            Next
            Do While tabladevuelve(0, UBound(tabladevuelve, 2)) = ""
                ReDim Preserve tabladevuelve(1, UBound(tabladevuelve, 2) - 1)
            Loop
            tablafinalfuncion = tabladevuelve
        End Function
        Public Shared Sub buscagarantiaacc(ByRef lista As String(,), ByVal numofa As String, ByVal Str_Con_Bd As String)
            Dim cons As String = "SELECT DISTINCT " +
                         "Of_Accesorios.Cant_Req, CASE WHEN isnull(Of_Accesorios.Nomenclatura, Accesorios_Codigos.Nomenclatura) IS NULL THEN item_planta.descripcion ELSE isnull(Of_Accesorios.Nomenclatura, Accesorios_Codigos.Nomenclatura) END AS Nomenclatura, " +
                         "ISNULL(Of_Accesorios.Des_Aux, Of_Accesorios.Observa) AS Des_Aux, Of_Accesorios.Dim1, Of_Accesorios.Dim2, Of_Accesorios.Dim3, Of_Accesorios.Dim4, Of_Accesorios.Dim5, Of_Accesorios.Dim6, RTRIM(Orden.Tipo_Of) + ' ' + RTRIM(Orden.Ofa) AS Ofa, " +
                         "Of_Accesorios.No_Item, Orden.Id_Of_P, Orden.Id_Ofa " +
                         "FROM Of_Accesorios WITH (nolock) INNER JOIN " +
                         "Orden WITH (nolock) ON Of_Accesorios.Id_Ofa = Orden.Id_Ofa INNER JOIN " +
                         "item_planta WITH (nolock) ON Orden.planta_id = item_planta.planta_id AND Of_Accesorios.Id_UnoE = item_planta.cod_erp INNER JOIN " +
                         "Orden_Seg WITH (nolock) ON Orden.Id_Of_P = Orden_Seg.Id_Ofa LEFT OUTER JOIN " +
                         "Accesorios_Codigos WITH (nolock) ON Of_Accesorios.Id_UnoE = Accesorios_Codigos.Id_UnoE AND Orden.planta_id = Accesorios_Codigos.planta_id " +
                        "WHERE (RTRIM(Orden_Seg.Num_Of) + '-' + RTRIM(Orden_Seg.Ano_Of) = '" & numofa & "') 
                        AND (Orden.Tipo_Of <> 'FP') AND (Orden.Tipo_Of <> 'OF') 
                        AND (Orden.Tipo_Of <> 'PR') AND (Orden.Tipo_Of <> 'F4') 
                        AND (Orden.Tipo_Of <> 'CT') AND (Orden.Tipo_Of <> 'IO') " +
                        "ORDER BY Ofa, Of_Accesorios.No_Item"
            Using connection As New SqlConnection(Str_Con_Bd)
                Dim command As SqlCommand = New SqlCommand(cons, connection)
                connection.Open()
                Dim reader As SqlDataReader = command.ExecuteReader()
                If reader.HasRows Then
                    Do While reader.Read()

                        'Dim idofaP = reader.GetValue(11)

                        Dim nomenclatura, desaux, dim1, dim2, dim3, dim4, dim5, dim6 As String
                        nomenclatura = Frm_ListFormaletas.RemoveXtraSpaces(reader.GetValue(1))
                        If reader.GetValue(3) = "0" Then
                            dim1 = ""
                        Else
                            dim1 = reader.GetValue(3)
                        End If
                        If reader.GetValue(4) = "0" Then
                            dim2 = ""
                        Else
                            dim2 = reader.GetValue(4)
                        End If
                        If reader.GetValue(5) = "0" Then
                            dim3 = ""
                        Else
                            dim3 = reader.GetValue(5)
                        End If
                        If reader.GetValue(6) = "0" Then
                            dim4 = ""
                        Else
                            dim4 = reader.GetValue(6)
                        End If
                        If reader.GetValue(7) = "0" Then
                            dim5 = ""
                        Else
                            dim5 = reader.GetValue(7)
                        End If
                        If reader.GetValue(8) = "0" Then
                            dim6 = ""
                        Else
                            dim6 = reader.GetValue(8)
                        End If
                        If IsDBNull(reader.GetValue(2)) Then
                            desaux = ""
                        ElseIf reader.GetValue(2) = " " Then
                            desaux = ""
                        Else
                            desaux = reader.GetValue(2)
                        End If

                        lista(0, UBound(lista, 2)) = reader.GetValue(0)     'cantidad
                        lista(1, UBound(lista, 2)) = nomenclatura           'nomenclatura
                        lista(2, UBound(lista, 2)) = desaux                 'aux
                        lista(3, UBound(lista, 2)) = dim1                   'alto1
                        lista(4, UBound(lista, 2)) = dim2                   'alto2
                        lista(5, UBound(lista, 2)) = dim3                   'ancho2
                        lista(6, UBound(lista, 2)) = dim4                   'Plano Especial
                        lista(7, UBound(lista, 2)) = dim5                   'observacion
                        lista(8, UBound(lista, 2)) = dim6                   'Familia
                        lista(9, UBound(lista, 2)) = ""                     'Area
                        lista(10, UBound(lista, 2)) = reader.GetValue(9)    'OF
                        lista(11, UBound(lista, 2)) = reader.GetValue(10)   'Item
                        lista(12, UBound(lista, 2)) = ""                    'Memo
                        lista(13, UBound(lista, 2)) = ""                    'Tipo de cambio
                        lista(14, UBound(lista, 2)) = ""                    'Observacion Memo
                        lista(15, UBound(lista, 2)) = ""                    'Ref Orden
                        lista(16, UBound(lista, 2)) = If(dim1 = "", 0, dim1) & "x" & If(dim2 = "", 0, dim2) & "x" & If(dim3 = "", 0, dim3) & "x" & If(dim4 = "", 0, dim4) & "x" & If(dim5 = "", 0, dim5) & "x" & If(dim6 = "", 0, dim6)
                        ReDim Preserve lista(18, UBound(lista, 2) + 1)

                    Loop
                End If
                command = Nothing
                reader.Close()
                reader = Nothing

            End Using
        End Sub
        Public Shared Sub buscagarantiaalum(ByRef lista As String(,), ByVal idofaP As String, ByVal Str_Con_Bd As String, ByVal tipoOF As String)
            Dim cons As String = "SELECT Saldos.Cant_Final_Req, Saldos.Tipo, Saldos.Anc1, Saldos.alto1, Saldos.Alto2, Saldos.Anc2, Saldos.Plano_esp, " +
            "IIF(saldos.Observacion = '.', '', Saldos.Observacion) AS Observacion, Saldos.Grupo, Saldos.Area, Orden.Ofa, Saldos.Item, Saldos.Ult_Memo, Memo_Det.Memo_DetOperacion, Memo_Det.Memo_Obs, Orden.Id_Of_P " +
            "FROM Saldos WITH (nolock) INNER JOIN " +
            "Orden WITH (nolock) ON Saldos.Id_Ofa = Orden.Id_Ofa LEFT OUTER JOIN " +
            "Memo_Det WITH (nolock) ON Saldos.Identificador = Memo_Det.Memo_DetIdSaldosOri " +
            "WHERE (Orden.Id_Of_P = " & idofaP & ") AND (Orden.Tipo_Of = '" & tipoOF & "') " +
            "ORDER BY Saldos.Grupo, Saldos.Item"

            Using connection As New SqlConnection(Str_Con_Bd)
                Dim command As SqlCommand = New SqlCommand(cons, connection)
                connection.Open()
                Dim reader As SqlDataReader = command.ExecuteReader()
                If reader.HasRows Then
                    Do While reader.Read()
                        Dim desaux, dim3, dim4, memo, TipoCambio, Obser, Plano As String
                        If reader.GetValue(4) = "0" Then
                            dim3 = ""
                        Else
                            dim3 = reader.GetValue(4)
                        End If
                        If reader.GetValue(5) = "0" Then
                            dim4 = ""
                        Else
                            dim4 = reader.GetValue(5)
                        End If
                        If reader.Item(6) = True Then
                            Plano = "P"
                        Else
                            Plano = ""
                        End If
                        If reader.GetValue(7) = " " Then
                            desaux = ""
                        Else
                            desaux = reader.GetValue(7)
                        End If
                        If IsDBNull(reader.GetValue(12)) Then
                            memo = ""
                        Else
                            memo = "'" & reader.GetValue(12)
                        End If
                        If IsDBNull(reader.GetValue(13)) Then
                            TipoCambio = ""
                        Else
                            TipoCambio = reader.GetValue(13)
                        End If
                        If IsDBNull(reader.GetValue(14)) Then
                            Obser = ""
                        Else
                            Obser = reader.GetValue(14)
                        End If
                        Dim concatenado = reader.GetValue(1) & "-" & reader.GetValue(2) & "-" & reader.GetValue(3) & "-" & dim3 & "-" & dim4 & "-" & desaux & "-" & reader.GetValue(8) & "-" & reader.GetValue(10) & "-" & reader.GetValue(11) & "-" & memo & "-" & TipoCambio & "-" & Obser
                        Dim sumar As Boolean = False
                        Dim k As Integer
                        For j = 0 To UBound(lista, 2)
                            Dim prueba = lista(1, j) & "-" & lista(2, j) & "-" & lista(3, j) & "-" & lista(4, j) & "-" & lista(5, j) & "-" & lista(7, j) & "-" & lista(8, j) & "-" & lista(10, j) & "-" & lista(11, j) & "-" & lista(12, j) & "-" & lista(13, j) & "-" & lista(14, j)
                            If concatenado = prueba Then
                                sumar = True
                                k = j
                                Exit For
                            End If
                        Next
                        If sumar = True Then
                            lista(0, k) = CDbl(lista(0, k)) + CDbl(reader.GetValue(0))
                        Else
                            lista(0, UBound(lista, 2)) = reader.GetValue(0)         'cantidad
                            lista(1, UBound(lista, 2)) = reader.GetValue(1)         'nomenclatura
                            lista(2, UBound(lista, 2)) = reader.GetValue(2)         'ancho1
                            lista(3, UBound(lista, 2)) = reader.GetValue(3)         'alto1
                            lista(4, UBound(lista, 2)) = dim3                       'alto2
                            lista(5, UBound(lista, 2)) = dim4                       'ancho2
                            lista(6, UBound(lista, 2)) = Plano                      'Plano Especial
                            lista(7, UBound(lista, 2)) = desaux                     'observacion
                            lista(8, UBound(lista, 2)) = reader.GetValue(8)         'Familia
                            lista(9, UBound(lista, 2)) = reader.GetValue(9)         'Area
                            lista(10, UBound(lista, 2)) = tipoOF & " " & reader.GetValue(10) 'OG o OM o RC
                            lista(11, UBound(lista, 2)) = reader.GetValue(11)         'Item
                            lista(12, UBound(lista, 2)) = memo                      'Memo
                            lista(13, UBound(lista, 2)) = TipoCambio                'Tipo de cambio
                            lista(14, UBound(lista, 2)) = Obser                     'Observacion Memo
                            ReDim Preserve lista(18, UBound(lista, 2) + 1)
                        End If
                    Loop

                End If

                command = Nothing
                reader.Close()
                reader = Nothing

            End Using
            cons = "SELECT Ofa, Observaciones FROM Orden WITH (nolock) WHERE (Id_Of_P = " & idofaP & ") AND (Tipo_Of = '" & tipoOF & "')"

            Using connection As New SqlConnection(Str_Con_Bd)
                Dim command As SqlCommand = New SqlCommand(cons, connection)
                connection.Open()
                Dim reader As SqlDataReader = command.ExecuteReader()
                If reader.HasRows Then
                    Do While reader.Read()
                        lista(0, UBound(lista, 2)) = 0                          'cantidad
                        lista(1, UBound(lista, 2)) = reader.GetValue(1)         'nomenclatura
                        lista(2, UBound(lista, 2)) = ""                         'ancho1
                        lista(3, UBound(lista, 2)) = ""                         'alto1
                        lista(4, UBound(lista, 2)) = ""                         'alto2
                        lista(5, UBound(lista, 2)) = ""                         'ancho2
                        lista(6, UBound(lista, 2)) = ""                         'Plano Especial
                        lista(7, UBound(lista, 2)) = ""                         'observacion
                        lista(8, UBound(lista, 2)) = ""                        'Familia
                        lista(9, UBound(lista, 2)) = ""                        'Area
                        lista(10, UBound(lista, 2)) = tipoOF & " " & reader.GetValue(0) 'OG o OM o RC
                        lista(11, UBound(lista, 2)) = ""                         'Item
                        lista(12, UBound(lista, 2)) = ""                        'Memo
                        lista(13, UBound(lista, 2)) = ""                        'Tipo de cambio
                        lista(14, UBound(lista, 2)) = ""                        'Observacion Memo
                        ReDim Preserve lista(18, UBound(lista, 2) + 1)
                    Loop
                End If

                command = Nothing
                reader.Close()
                reader = Nothing

            End Using
        End Sub
        Public Shared Sub escribe(ByVal filename As String, tablaitems As Object, convertir As Boolean)
            Dim FN As Integer
            FN = FreeFile()
            Using sw As New StreamWriter(filename, False, System.Text.Encoding.Default) '= File.CreateText(filename)
                'verifica si la primera fila es el encabezado
                Dim j As Integer
                If tablaitems(0, 0) = "CANT" Then
                    sw.WriteLine(tablaitems(0, 0) & vbTab & tablaitems(1, 0) & vbTab & tablaitems(2, 0) & vbTab & tablaitems(3, 0) & vbTab & tablaitems(4, 0) & vbTab & tablaitems(5, 0) & vbTab & tablaitems(6, 0) & vbTab & tablaitems(7, 0) & vbTab & tablaitems(8, 0) & vbTab & tablaitems(9, 0) & vbTab & "AREA_TOTAL" & vbTab & vbTab & tablaitems(10, 0) & vbTab & tablaitems(11, 0) & vbTab & tablaitems(12, 0) & vbTab & tablaitems(13, 0) & vbTab & tablaitems(14, 0) & vbTab & tablaitems(15, 0) & vbTab & tablaitems(16, 0) & vbTab & tablaitems(17, 0) & vbTab & tablaitems(18, 0))

                    j = 1
                Else
                    j = 0
                End If

                'inicia una posición despues del encabezado
                For i = j To tablaitems.GetUpperBound(1)
                    Dim familia, observ, areaT As String

                    Try
                        Select Case tablaitems(8, i)
                            Case "FM"
                                familia = "MUROS"
                            Case "FL"
                                familia = "LOSAS"
                            Case "UM"
                                familia = "UNION"
                            Case "CL"
                                familia = "CULAT"
                            Case ""
                                familia = ""
                            Case Else
                                familia = tablaitems(8, i) 'dimension 6
                        End Select
                    Catch ex As System.Exception
                        familia = "MUROS"
                    End Try
                    Try

                        observ = tablaitems(7, i)


                    Catch ex As System.Exception
                        observ = ""
                    End Try

                    Try
                        areaT = CStr(CInt(tablaitems(0, i)) * CDbl(tablaitems(9, i)))
                    Catch ex As System.Exception
                        areaT = ""
                    End Try

                    Try
                        sw.WriteLine(CInt(tablaitems(0, i)) & vbTab & tablaitems(1, i) & vbTab & tablaitems(2, i) & vbTab & tablaitems(3, i) & vbTab & tablaitems(4, i) & vbTab & tablaitems(5, i) & vbTab & tablaitems(6, i) & vbTab & observ & vbTab & familia & vbTab & tablaitems(9, i) & vbTab & areaT & vbTab & vbTab & tablaitems(10, i) & vbTab & tablaitems(11, i) & vbTab & tablaitems(12, i) & vbTab & tablaitems(13, i) & vbTab & tablaitems(14, i) & vbTab & tablaitems(15, i) & vbTab & tablaitems(16, i) & vbTab & tablaitems(17, i) & vbTab & tablaitems(18, i))
                    Catch ex As System.Exception
                        sw.WriteLine(CInt(tablaitems(0, i)) & vbTab & tablaitems(1, i) & vbTab & tablaitems(2, i) & vbTab & tablaitems(3, i) & vbTab & tablaitems(4, i))
                    End Try

                    'sw.WriteLine(CInt(tablaitems(0, i)) & vbTab & tablaitems(1, i) & vbTab & tablaitems(2, i) & vbTab & tablaitems(3, i) & vbTab & tablaitems(4, i) & vbTab & "" & vbTab & vbTab & observ & vbTab & familia & vbTab & vbTab & vbTab & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "")
                    'sw.WriteLine(CInt(tablaitems(0, i)) & vbTab & tablaitems(1, i) & vbTab & tablaitems(2, i) & vbTab & tablaitems(3, i) & vbTab & tablaitems(4, i) & vbTab & tablaitems(5, i) & vbTab & vbTab & observ & vbTab & familia & vbTab & vbTab & vbTab & vbTab & tablaitems(8, i) & vbTab & tablaitems(9, i) & vbTab & tablaitems(10, i) & vbTab & tablaitems(11, i) & vbTab & tablaitems(12, i))

                Next

            End Using

            'Dim reader As StreamReader = New StreamReader(filename)
            'Dim xlApp = New Microsoft.Office.Interop.Excel.Application
            'Dim xlWorkBook = xlApp.Workbooks.Add 'xlApp.Workbooks.Open(filename)
            'Dim sheet = xlWorkBook.ActiveSheet
            'Dim xlWorkSheet As Microsoft.Office.Interop.Excel.Worksheet
            'Dim line As String

            'Dim lineIndex As Long = 1

            'Dim progreso As New Form8
            'progreso.ProgressBar1.Minimum = 0
            'progreso.ProgressBar1.Maximum = tablaitems.GetUpperBound(1) + 1 'porque el line index esta quedando mas grande
            'progreso.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen

            'progreso.Show()
            'progreso.TopMost = True
            'progreso.Text = "descargando lista"

            'Do While reader.Peek() >= 0
            '    line = reader.ReadLine()
            '    WriteToExcel(line, sheet, lineIndex)
            '    Try
            '        progreso.ProgressBar1.Value = lineIndex 'porque el line index esta quedando mas grande
            '    Catch ex As system.Exception

            '    End Try
            '    lineIndex += 1
            'Loop

            'progreso.Close()

            'Dim filenamesplit = Split(filename, ".")
            ''reader.Close()
            ''My.Computer.FileSystem.DeleteFile(filename)
            'filename = filenamesplit(0)
            'For i = 1 To UBound(filenamesplit) - 1
            '    filename = filename & "." & filenamesplit(i)
            '    'i = i + 1
            'Next



            'If convertir = True Then
            '    'revisa hasta 100 veces si el archivo existe
            '    For i = 0 To 100
            '        If My.Computer.FileSystem.FileExists(filename & ".xls") Then
            '            filename = filename & i
            '        Else
            '            Continue For
            '        End If
            '    Next
            '    filename = filename & ".xls"
            'Else
            '    'revisa hasta 100 veces si el archivo existe
            '    For i = 0 To 100
            '        If My.Computer.FileSystem.FileExists(filename & ".xls") Then
            '            filename = filename & i
            '        Else
            '            Exit For
            '        End If
            '    Next
            '    filename = filename & ".xls"
            'End If



            'Try
            '    xlWorkBook.SaveAs(filename, Microsoft.Office.Interop.Excel.XlFileFormat.xlTextWindows) 'xlWorkBook.SaveCopyAs("C:\Users\hugogomez\Desktop\Nueva carpeta\ensayo.xls")
            '    xlWorkBook.Close()
            '    MsgBox("se ha guardado el listado con el nombre " & filename)

            'Catch ex As System.Exception
            '    MsgBox("por favor cierre el archivo del listado antes de generar otro nuevo")

            'End Try
            'Try
            '    xlWorkBook.Close()
            'Catch ex2 As system.Exception

            'End Try
        End Sub
        Public Shared Sub WriteToExcel(line As String, targetWorksheet As Microsoft.Office.Interop.Excel.Worksheet, lineIndex As Long)
            Dim column As Integer = 1
            For Each part As String In line.Split(vbTab)
                targetWorksheet.Cells(lineIndex, column).Value = part
                If column = 13 Then
                    Dim a As Object = targetWorksheet.Cells(lineIndex, column).numberformat
                    Dim b As Object = targetWorksheet.Cells(lineIndex, column).value
                    If a = "m/d/yyyy" Then
                        targetWorksheet.Cells(lineIndex, column).numberformat = "@"
                        targetWorksheet.Cells(lineIndex, column).Value = "' " & part
                        a = targetWorksheet.Cells(lineIndex, column).numberformat
                        b = targetWorksheet.Cells(lineIndex, column).value
                    End If
                End If
                column += 1
            Next
        End Sub
        Public Shared Function RutaPlanoAcad() As String
            'Dim doc As Document = Application.DocumentManager.MdiActiveDocument
            'Dim hs As HostApplicationServices = HostApplicationServices.Current
            'Dim path As String
            'Try
            '    path = hs.FindFile(doc.Name, doc.Database, FindFileHint.Default)
            'Catch ex As System.Exception
            '    MsgBox("Debe guardar el archivo de autocad para continuar")
            '    Exit Function
            'End Try
            'Dim PathVec() As String = Split(path, "\")
            'Dim filename As String
            'filename = ""
            'For i = 0 To UBound(PathVec) - 1
            '    filename = filename & PathVec(i) & "\"
            'Next
            'Return filename
        End Function

        Public Shared Function esVersionActual() As Boolean
            Dim versionserver As Double
            Dim cons As String
            cons = "SELECT VersionCargaSif FROM Parametros WITH (nolock)"

            Dim Str_Con_Bd As String = ConexionBD.getStringConexion()

            Using connection As New SqlConnection(Str_Con_Bd)
                Dim command As SqlCommand = New SqlCommand(cons, connection)
                connection.Open()
                Dim reader As SqlDataReader = command.ExecuteReader
                reader.Read()

                versionserver = reader.Item(0)

            End Using
            versionapp = 5.46
            If versionserver <= versionapp Then
                esVersionActual = True
            Else
                esVersionActual = False
                MsgBox("actualice Forline para usar la ultima version del software")

            End If

        End Function
    End Class
End Namespace