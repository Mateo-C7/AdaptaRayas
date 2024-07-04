Imports System.Data.SqlTypes
Imports System.Windows
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Windows
Imports CreaRQI.CapaDatos

Public Class Orden

    Public Property IdPlanta As Integer ' Identificador de la planta
    Public Property IdOfa As Integer ' Identificador de la raya
    Public Property IdOfaP As Integer ' Identificador de la orden principal
    Public Property TipoOf As String ' Tipo de Orden
    Public Property NumOrden As String ' Numero de Orden
    Public Property Raya As String ' Numero de Orden
    Public Property FUP As Integer ' Numero unico de Proyecto
    Public Property TipoIF As String
    Public Property NumIF As Integer
    Public Property ItemIF As Integer
    Public Property ItemALUM As Integer ' Estructura de aluminio
    Public Property ItemAC As Integer ' Estructura de aluminio
    Public Property ItemCOM As Integer ' Estructura de aluminio
    Public Property ItemOP As Integer ' Estructura de aluminio
    Public Property OD_ALUM As Integer
    Public Property OD_AC As Integer
    Public Property OD_COM As Integer
    Public Property VerModelo As String ' Version Modelo de Costos
    Public Property TipoOP As String ' Tipo de Orden
    Public Property NumOP As Integer ' Numero de Orden
    Public Property Pais As String
    Public Property Ciudad As String
    Public Property Cliente As String
    Public Property Obra As String
    Public Property M2_IF As Single ' M2 Cotizados de la IF
    Public Property M2 As Single ' M2 de formaletas en la Raya
    Public Property Kg As Single ' Kilos de Acero
    Public Property Letra As Integer ' Numero de Raya
    Public Property Crea_Ac As Boolean ' Crear orden de Acc Produccion
    Public Property Crea_Com As Boolean ' Crear orden de Acc Almacen
    Public Property RayaMemo As Boolean ' Indica si es una Raya de memorando
    Public Property Requisicion As Integer ' Numero de Requisicion de Acc
    Public Property CantReq As Integer ' Cantidad de piezas, aplica para kanban
    Public Property DescPieza As String ' Descripcion de la Pieza de Kanban
    Public Property PlantaLoginUsu As Integer 'Planta con la que el usuario se loguea

    Dim Conn As ConexionBD = New ConexionBD

    'De la orden trae: Pais - Cliente - Obra
    Public Sub New()
        'Me.Reset()
    End Sub

    ' Validacion de una Orden Padre Por el Numero de Orden
    Public Function Validar(ByVal NumOrden As String, ByRef stMsg As String) As Boolean
        Dim retorno As Boolean = False
        Dim numSep As Integer
        stMsg = ""

        ' Reinicializar Objeto
        Reset()
        numSep = NumOrden.Split("-").Length - 1

        ' Si no tiene Separador de Numero de Orden
        If numSep = 0 OrElse numSep > 2 Then
            stMsg = "Nùmero de Orden no es Valido"
            Return False
        End If

        ' Si tiene un Separador, se valida la Orden Padre
        If numSep = 1 Then
            retorno = Validar(0, NumOrden, stMsg)
        End If

        Return retorno
    End Function

    ' Validacion de una Orden Padre Por IdOfa Padre
    Public Function Validar(ByVal IdOfaP As Long, ByRef stMsg As String) As Boolean
        Dim retorno As Boolean = False
        stMsg = ""

        ' Si no tiene Separador de Numero de Orden
        If IdOfaP <= 0 Then
            stMsg = "Nùmero de Orden no es Valido"
            Return False
        End If

        retorno = Validar(IdOfaP, "", stMsg)
        Return retorno
    End Function

    ' Validacion de una Orden Padre
    Public Function Validar(ByVal IdOfaP As Long, ByVal NumOrden As String, ByRef stMsg As String) As Boolean
        Dim regVal As DataRow = Nothing
        Dim sqlSel As String = ""
        Dim retorno As Boolean = False
        stMsg = ""
        Try
            sqlSel = " SELECT S.Id_Ofa IdOfaP, S.Tipo_Of TipoOf, REPLACE(S.Num_Of + '-' + S.Ano_Of, ' ', '')  NumOrden, " &
                 " S.FUP,	S.Tip_Ord_Abu TipoIF, S.OP_1EE_Abuelo NumIF, S.Item_1EE_Abuelo ItemIF, S.Item_Ac_Al ItemAC, " &
                 " S.Item_Ac_Com ItemCOM, S.ND_1EE ItemALUM, S.OP_Ac_Al OD_AC, S.OP_Ac_Com OD_COM, S.OP_1EE OD_ALUM, " &
                 " S.Mod_Ver VerModelo, S.Planta_Id, S.M2_Cotizados M2_IF " &
                 " FROM ORDEN_SEG S WITH(NOLOCK) " &
                 " WHERE S.Anulado = 0 " '& Me.PlantaLoginUsu &  '" WHERE S.Planta_Id = " & objUsu.Planta_id.ToString() &

            ' Si tiene un Separador, se valida la Orden Padre
            If IdOfaP > 0 Then
                sqlSel += " AND S.IdOfa  = " & IdOfaP.ToString()
            Else
                ' Se valida una Raya
                sqlSel += "  AND REPLACE(S.Num_Of+'-' + S.Ano_Of,' ','') = '" & NumOrden & "'"
            End If

            regVal = Conn.SelReg(sqlSel)

            If regVal IsNot Nothing Then
                retorno = True
                Me.IdPlanta = Integer.Parse(regVal("Planta_Id").ToString())
                Me.IdOfaP = Integer.Parse(regVal("IdOfaP").ToString())
                Me.TipoOf = regVal("TipoOf").ToString().Trim()
                Me.NumOrden = regVal("NumOrden").ToString().Trim()
                Me.FUP = Integer.Parse(regVal("FUP").ToString())
                Me.TipoIF = regVal("TipoIF").ToString().Trim()
                Me.NumIF = Integer.Parse(regVal("NumIF").ToString())
                Me.ItemIF = Integer.Parse(regVal("ItemIF").ToString())
                Me.ItemAC = Integer.Parse(regVal("ItemAC").ToString())
                Me.ItemCOM = Integer.Parse(regVal("ItemCOM").ToString())
                Me.ItemALUM = Integer.Parse(regVal("ItemALUM").ToString())
                Me.OD_AC = Integer.Parse(regVal("OD_AC").ToString())
                Me.OD_COM = Integer.Parse(regVal("OD_COM").ToString())
                Me.OD_ALUM = Integer.Parse(regVal("OD_ALUM").ToString())
                Me.VerModelo = regVal("VerModelo").ToString().Trim()
                Me.M2_IF = Single.Parse(regVal("M2_IF").ToString())
            Else
                ' Orden no es Valida
                stMsg = "Nùmero de Orden no es Valido"
            End If
            Return retorno
        Catch e As Exception
            ' Manejo de excepcion
            'myGn.ProErrorException("Orden.Validar", e, sqlSel)
            Conn.RegistraExcepcion("SYS", "Orden.Validar", e.ToString(), sqlSel)

            Return False
        End Try
    End Function

    ' Recuperar Informacion del Cliente
    Public Sub RecuperarInfoCliente()
        Dim regVal As DataRow = Nothing
        Dim sqlSel As String = ""
        Try
            ' Si esta inicializada la Orden
            If Me.IdOfaP > 0 AndAlso Me.TipoOf.Length > 0 Then
                ' Si es una orden de Garantia o Mejora
                If Me.TipoOf = "OG" OrElse Me.TipoOf = "OM" Then
                    sqlSel = " SELECT OD.Id_Ofa IdOfaP, OD.Ofa NumOrden, OD.Yale_Cotiza FUP, P.pai_id IdPais, UPPER(P.pai_nombre) NomPais," &
                         " CL.cli_id IdCliente, CL.cli_nombre NomCliente, CI.ciu_id IdCiudad, UPPER(CI.ciu_nombre) NomCiudad, " &
                         " OB.obr_id IdObra, UPPER(OB.obr_nombre) NomObra " &
                         " FROM QRS Q WITH(NOLOCK)  INNER JOIN CLIENTE CL WITH(NOLOCK)  ON Q.qr_cliente_id = CL.cli_id " &
                         " INNER JOIN CIUDAD CI WITH(NOLOCK) ON CL.cli_ciu_id = CI.ciu_id " &
                         " INNER JOIN PAIS P WITH(NOLOCK) ON CL.cli_pai_id = P.pai_id " &
                         " INNER JOIN OBRA OB  WITH(NOLOCK) ON Q.qr_obra_id = OB.obr_id " &
                         " INNER JOIN ORDEN OD   WITH(NOLOCK) ON CASE WHEN Q.qr_ofa_gar_id > 0 THEN Q.qr_ofa_gar_id ELSE  Q.qr_ofa_ome_id END = OD.Id_Ofa " &
                         " WHERE  OD.Id_Ofa  = " & Me.IdOfaP.ToString() &  '" WHERE OD.Planta_id  = " & objUsu.Planta_id.ToString() & -- & Me.PlantaLoginUsu &
                         " ORDER BY qr_id  DESC "
                Else
                    sqlSel = " SELECT OS.Id_Ofa IdOfaP, REPLACE( OS.Num_Of + '-' + OS.Ano_Of, ' ', '')  NumOrden, FU.fup_id FUP, " &
                         " P.pai_id IdPais, UPPER(P.pai_nombre) NomPais, CL.cli_id IdCliente, CL.cli_nombre NomCliente, " &
                         " CI.ciu_id IdCiudad, UPPER(CI.ciu_nombre) NomCiudad, OB.obr_id IdObra, UPPER(OB.obr_nombre) NomObra " &
                         " FROM   FORMATO_UNICO FU WITH(NOLOCK) INNER JOIN  CLIENTE CL WITH(NOLOCK)    ON FU.fup_cli_id = CL.cli_id " &
                         " INNER JOIN  PAIS P  WITH(NOLOCK) ON CL.cli_pai_id = P.pai_id " &
                         " INNER JOIN  OBRA OB ON FU.fup_obr_id = OB.obr_id " &
                         " INNER JOIN ORDEN OD WITH(NOLOCK) ON FU.fup_id = OD.Yale_Cotiza " &
                         " INNER JOIN ORDEN_SEG OS   WITH(NOLOCK) ON OS.Id_Ofa = OD.Id_Ofa  " &
                         " INNER JOIN  PAIS P2  WITH(NOLOCK) ON OB.obr_pai_id = P2.pai_id " &
                         " INNER JOIN CIUDAD CI WITH(NOLOCK)   ON CL.cli_ciu_id = CI.ciu_id  AND CL.cli_pai_id = P.pai_id " &
                         " WHERE OS.Id_Ofa  = " & Me.IdOfaP.ToString() & '" WHERE OS.Planta_id  = " & objUsu.Planta_id.ToString() & -- & Me.PlantaLoginUsu &
                         " ORDER BY OS.Id_Ofa "
                End If
                regVal = Conn.SelReg(sqlSel)
                If regVal IsNot Nothing Then
                    Me.Pais = regVal("NomPais").ToString()
                    Me.Ciudad = regVal("NomCiudad").ToString()
                    Me.Cliente = regVal("NomCliente").ToString()
                    Me.Obra = regVal("NomObra").ToString()
                    '--------------------------------
                End If
            End If
        Catch e As Exception
            'Manejo de excepcion
            'myGn.ProErrorException("Orden.RecuperarInfoCliente", e, sqlSel)
            Conn.RegistraExcepcion("SYS", "Orden.RecuperarInfoCliente", e.ToString(), sqlSel)
        End Try
    End Sub

    'Realiza una consulta que trae toda la informacion cargada(Aluminio) a la orden, para presentar al cliente.
    'Nota: No separa por rayas
    Public Function ConsultarListadoClienteAlum(ByVal NumOrden As String) As DataTable

        Dim dt As DataTable = Nothing

        Dim sql As String = "SELECT SUM(ISNULL(MD.Memo_DetCantPFin, S.cant)) AS Cant, " &
                            "S.Tipo  AS Descripcion, S.Observacion AS 'Observacion',  " &
                            "S.Anc1  AS 'Ancho 1', " &
                            "S.alto1 AS 'Alto 1',  " &
                            "S.Alto2 As 'Alto 2',  " &
                            "S.Anc2  AS 'Ancho 2', " &
                            "CASE WHEN S.Plano_esp = '1'THEN 'P' ELSE '' END AS Plano, " &
                            " '', '', " &
                            "I.Nombre + ' '+ ISNULL(PF.Observacion,'') AS 'Refencia Alum' " &
                            "FROM Explo_Saldos As ES WITH (NOLOCK) " &
                            "INNER JOIN Formaleta_Lib_Inv AS FLI WITH (NOLOCK) ON ES.Explo_Lib_Id = FLI.Formaleta_Lib_Inv_Id " &
                            "RIGHT OUTER JOIN Saldos AS S With (NOLOCK) " &
                            "INNER JOIN Orden AS O WITH (NOLOCK) ON S.Id_Ofa = O.Id_Ofa ON ES.Saldos_Id = S.Identificador " &
                            "LEFT OUTER JOIN Memos AS M WITH (NOLOCK) INNER JOIN Memo_Det AS MD WITH (NOLOCK) ON M.Id_Memo = MD.Id_MemoId ON S.Ult_Memo = M.Memo_No " &
                            "AND S.Identificador = MD.Memo_DetIdSaldosOri " &
                            "INNER JOIN Piezas_Forsa AS PF WITH(NOLOCK) ON PF.Id_Piezas = S.Id_Piezas_Forsa " &
                            "INNER JOIN Items AS I WITH(NOLOCK) ON I.Id_Item = PF.Id_Item " &
                            "WHERE (O.Ofa LIKE '%" & NumOrden & "%') " &
                            "AND   (S.Anula = 0) " &
                            "AND   (PF.Anulado = 0) " &
                            "GROUP BY Descripcion,S.Tipo, S.Observacion, S.Anc1 , S.alto1 , S.Alto2 , S.Anc2, S.Plano_esp,S.Grupo, I.Nombre, PF.Observacion " &
                            "ORDER BY S.Descripcion "
        Try

            dt = Conn.CargarTabla(sql)

        Catch ex As Exception

            Conn.RegistraExcepcion("SYS", "Orden.ConsultarListadoClienteAlum", ex.ToString(), sql)

        End Try

        Return dt

    End Function

    'Realiza una consulta que trae toda la informacion cargada(Accesorios) a la orden, para presentar al cliente.
    'Nota: No separa por rayas
    Public Function ConsultarListadoClienteACC(ByVal NumOrden As String) As DataTable

        Dim dt As DataTable = Nothing
        Dim sql As String = "SELECT SUM(ISNULL(MACD.Memo_Acc_Cant_Fin, OFACC.Cant_Req)) AS Cant, " &
                            "CASE " &
                            "WHEN OFACC.Nomenclatura !='' THEN OFACC.Nomenclatura " &
                            "WHEN AC.Nomenclatura !='' THEN AC.Nomenclatura " &
                            "ELSE ITP.descripcion END AS Descripcion, " &
                            "CASE " &
                            "WHEN OFACC.Des_Aux != '' THEN OFACC.Des_Aux " &
                            "ELSE AC.Des_Aux END AS 'Des_Aux'," &
                            "CASE WHEN CAST(OFACC.Dim1 AS varchar) = '0' THEN '' ELSE CAST(OFACC.Dim1 AS varchar) END AS Dim1, " &
                            "CASE WHEN CAST(OFACC.Dim2 AS varchar) = '0' THEN '' ELSE CAST(OFACC.Dim2 AS varchar) END AS Dim2, " &
                            "CASE WHEN CAST(OFACC.Dim3 AS varchar) = '0' THEN '' ELSE CAST(OFACC.Dim3 AS varchar) END AS Dim3, " &
                            "CASE WHEN CAST(OFACC.Dim4 AS varchar) = '0' THEN '' ELSE CAST(OFACC.Dim4 AS varchar) END AS Dim4, " &
                            "CASE WHEN CAST(OFACC.Dim5 AS varchar) = '0' THEN '' ELSE CAST(OFACC.Dim5 AS varchar) END AS Dim5, " &
                            "CASE WHEN CAST(OFACC.Dim6 AS varchar) = '0' THEN '' ELSE CAST(OFACC.Dim6 AS varchar) END AS Dim6, " &
                            "CASE WHEN CAST(AC.Valor_7_Min AS varchar) = '0' THEN '' ELSE CAST(AC.Valor_7_Min AS varchar) END AS Dim7, " &
                            "ITP.descripcion AS 'Referencia ACC' " &
                            "FROM Of_Accesorios AS OFACC WITH (nolock)  " &
                            "INNER JOIN Orden AS ORD WITH (nolock) ON OFACC.Id_Ofa = ORD.Id_Ofa  " &
                            "INNER JOIN item_planta AS ITP WITH (nolock) ON ORD.planta_id  = ITP.planta_id AND OFACC.Id_UnoE = ITP.cod_erp " &
                            "LEFT OUTER JOIN Memo_Acc_Det AS MACD WITH (nolock)  ON OFACC.Id_Orden_Acce   = MACD.Memo_Acc_OfAccId " &
                            "LEFT JOIN Accesorios_Codigos AS AC WITH(nolock)   ON AC.Codigos_Id = OFACC.AccesoriosCodigosId AND ORD.planta_id = AC.planta_id  " &
                            "WHERE (ORD.Ofa LIKE '%" & NumOrden & "%') " &
                            "AND (OFACC.Anula = 0) " &
                            "AND ITP.activo = 1 " &
                            "GROUP BY OFACC.Nomenclatura, AC.Nomenclatura,ITP.descripcion, OFACC.Des_Aux, AC.Des_Aux ,OFACC.Dim1, OFACC.Dim2, OFACC.Dim3, OFACC.Dim4, OFACC.Dim5, OFACC.Dim6, AC.Valor_7_Min " &
                            "ORDER BY Descripcion "
        Try

            dt = Conn.CargarTabla(sql)

        Catch ex As Exception

            Conn.RegistraExcepcion("SYS", "Orden.ConsultarListadoClienteACC", ex.ToString(), sql)

        End Try

        Return dt

    End Function

    Public Sub Reset()
        Me.IdPlanta = 0
        Me.IdOfa = 0
        Me.IdOfaP = 0
        Me.TipoOf = ""
        Me.NumOrden = ""
        Me.Raya = ""
        Me.FUP = 0
        Me.TipoIF = ""
        Me.ItemIF = 0
        Me.NumIF = 0
        Me.ItemALUM = 0
        Me.ItemAC = 0
        Me.ItemCOM = 0
        Me.ItemOP = 0
        Me.OD_ALUM = 0
        Me.OD_AC = 0
        Me.OD_COM = 0
        Me.TipoOP = ""
        Me.NumOP = 0
        Me.Pais = ""
        Me.Ciudad = ""
        Me.Cliente = ""
        Me.Obra = ""
        Me.M2 = 0
        Me.Crea_Ac = False
        Me.Crea_Com = False
        Me.Letra = 0
        Me.RayaMemo = False
        Me.Requisicion = 0
        Me.M2_IF = 0
        Me.Kg = 0
        Me.CantReq = 0
        Me.DescPieza = ""
        Me.VerModelo = ""
    End Sub


End Class
