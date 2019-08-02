Imports System.Net.Sockets
Imports System.Threading
Imports System.Data.Odbc
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Text
Imports Microsoft.VisualBasic
Imports System.Net
Imports System


'*****************************************
'IMPR#caceres.ltg
'Y a continuación:
'VQK#NOMBRE=TOLEDO19#CODBARRAS=8432971052856

'10.106.5.70 :IMP SUPERIOR  :UNICODE 1 : Contraseña VNC: calamardo Contraseña admin: 12345
'10.106.5.71 :IMP LATERAL 1 :UNICODE 0 : LIMITAG V5 LITE 8.3.0.1 L FS
'10.106.5.119:IMP LATERAL 2 :UNICODE 1 : LIMITAG COMPACTA V5 1.0.33

'TODO ES COREECTO CON UNICODE EN LAS 3 IMPRESORAS Y ENVIANDO ESTOS COMANDOS SE CARGAN LOS DATOS
'IMPR#caceres.ltg
'VQK#NOMBRE = LANZAROTE 25#CODBARRAS=8432971048835

'IMPRESORA LATERAL  : ALIAS1LAT:   ESTA VARIABLE RECIBE EL DATO DEL ALIAS DE MOVEX TIPO UNO 
'                   : ALIAS2EAN13: ESTA VARIABLE RECIBE EL DATO DEL ALIAS DE MOVEX TIPO DOS 
'                   : ETIQUETALAT:    BR01LAT
'IMPRESORA SUPERIOR : ALIAS2EANB:  ESTA VARIABLE RECIBE EL DATO DEL ALIAS DE MOVEX TIPO DOS 
'                   : DESCRIPCION: ESTA VARIABLE RECIBE EL DATO DEL OIS005
'                   : ALIAS1SUP:   ESTA VARIABLE RECIBE EL DATO DEL ALIAS DE MOVEX TIPO UNO 
'                   : ALIAS6:      ESTA VARIABLE RECIBE EL DATO DEL ALIAS DE MOVEX TIPO SEIS
'                   : ETIQUETASUP:    BR01SUP
'PLC CAJAS OMRON CS1: 10.78.38.151 PUERTO 9600
'                   : VARIABLE ANCHO_PANT_CAJAS: D00320[1][1]
'                   : VARIABLE LARGO_PANT_CAJAS: D00356[1][1]
'                   : VARIABLE ALTO_PANT_CAJAS:  D00344[1][1]
'*****************************************


Public Class FormLimitronic
    Dim WithEvents WinSockCliente As New Cliente
    Dim WithEvents WinSockCliente2 As New Cliente2
    Dim WithEvents WinSockCliente3 As New Cliente3
    Dim WithEvents WinSockCliente4 As New Cliente4

    Private Sub btnEnviarConfiguracion_Click(sender As Object, e As EventArgs) Handles btnEnviar.Click
        txtDialogo.Text = ""

        Try
            'RPM EnviarDatosImpresora1()
        Catch
            txtDialogo.Text = "Fallo al enviar datos a Impresora 1"
        End Try
        FormLimitronic.ActiveForm.Refresh()
        Try
            'RPM EnviarDatosImpresora2()
        Catch
            txtDialogo.Text = "Fallo al enviar datos a Impresora 2"
        End Try
        FormLimitronic.ActiveForm.Refresh()
        Try
            'RPM EnviarDatosImpresora3()

        Catch
            txtDialogo.Text = "Fallo al enviar datos a impresora 3"
        End Try

        Try
            EnviarDatosPLC()
        Catch
            txtDialogo.Text = "Fallo al enviar datos a PLC"
        End Try
        FormLimitronic.ActiveForm.Refresh()

    End Sub

    Private Sub EnviarDatosPLC()
        Dim ComandoPLC As String
        Dim Largo As Integer
        Dim Ancho As Integer
        Dim Grueso As Integer

        Largo = txtLargo.Text + txtSuplementoLargo.Text
        Ancho = txtAncho.Text + txtSuplementoAncho.Text
        Grueso = txtGrueso.Text + txtSuplementoGrueso.Text

        ' RPM ComandoPLC = "800002"   'comando de escritura ICF RSV GCT
        ComandoPLC = "800007"   ' RPM comando de escritura ICF RSV GCT
        ComandoPLC += "00"      'DNA red de destino, como es local = 0 OJO NO ES LOCAL
        ComandoPLC += Hex("151")      'DA1 destination node. Ultimo octeto de la IP de destino.
        ComandoPLC += "00"      'DA2 número de unidad de destino 00 = PLC
        ComandoPLC += "00"      'SNA -> Source network address: (hex)00, local network
        ComandoPLC += Hex("88")      'SA1 -> Source node address: (El último octeto de la IP del PC: 88 SERVIDOR APPNET)
        ComandoPLC += "00"      'SA2 Source unit address: (hex)00, PC only has one ethernet
        ComandoPLC += "00"      'SID -> ID del servicio. Puede ser de 00 a ff
        ComandoPLC += "0102"      'MRC SRC Write memory command. Main Request Code: 01, memory area write
        ComandoPLC += "82"      'DM Memory area code (1 byte): 82(DM)
        ' RPM ComandoPLC += Hex("000364")  'Dirección del área en Hex (En este caso D00356)164
        ComandoPLC += Hex("0356")  ' RPM Dirección del área en Hex (En este caso D00356)
        ComandoPLC += "00" ' RPM 
        ComandoPLC += Hex("0001")    'Cantidad de datos a escribir (2 byte). Solo escribe uno    
        ComandoPLC += Hex("1001")    'Dato. En este caso, DM0 quedaría como 0xFF01: 012c = 300
        'ComandoPLC += Strings.Left("0000", 4 - Hex(Largo).Length) + Hex(Largo) sustituir por el anterior

        ComandoPLC = txtComando.Text

        'rpm  Dim ComandoPLCByte As Byte() = System.Text.Encoding.ASCII.GetBytes(ComandoPLC)

        Dim ComandoPLCByte As Byte() = System.Text.Encoding.Unicode.GetBytes(ComandoPLC)

        Dim Puerto As Int32 = 9600

        Dim client As System.Net.Sockets.TcpClient
        Dim stream As System.Net.Sockets.NetworkStream
        client = New System.Net.Sockets.TcpClient("10.78.38.151", Puerto)
        txtDialogo.Text += "Enviando datos a PLC: " & ComandoPLC
        stream = client.GetStream()
        Try
            stream.Write(ComandoPLCByte, 0, ComandoPLCByte.Length)

        Catch
        End Try
        System.Threading.Thread.Sleep(4000)

        '-----------------------------------------------------
        Try
            Dim buffer(2048) As Byte

            Dim a As Integer = 1
            While Not stream.DataAvailable And a < 10  ' Máximo 10 iteraciones... 10 segundos
                Threading.Thread.Sleep(100)  ' milésimas de segundo.
                a += 1
            End While
            If stream.CanRead And stream.DataAvailable Then
                Dim respuesta As Integer = stream.Read(buffer, 0, 2048)

                Dim cadena As String
                cadena = System.Text.Encoding.ASCII.GetString(buffer, 0, respuesta)
                Dim longitud As String = Mid(cadena, 15, 4)
                Dim codError As String = Mid(cadena, 19, 4)
                codError = cadena
                If longitud = "0004" And codError = "0000" Then
                Else
                    txtDialogo.Text = "Intento: " + ComandoPLC + a.ToString + ". ERROR " + codError + " Respuesta: " + cadena
                End If
            Else
                txtDialogo.Text = "Intento: " + a.ToString + ". No hay respuesta del PLC."
            End If
        Catch ex As Exception
            txtDialogo.Text += "Error respuesta del PLC. " + ex.Message
        End Try
        '-----------------------------------------------------

        Try
            stream.Close()
            client.Close()
        Catch
            txtDialogo.Text += "Error al cerrar la comunicacion" + vbNewLine
        End Try


    End Sub

    Private Sub EnviarDatosImpresora1()
        Dim SeleccionarEtiquetaLateral As String
        Dim SeleccionarEtiquetaSuperior As String
        Dim CadenaImpresoraLateral As String
        Dim CadenaImpresoraSuperior As String


        'Primero carga la etiqueta a utilizar:

        SeleccionarEtiquetaSuperior = Chr(2) + "IMPR#" + txtEtiquetaSup.Text + Chr(3)


        If txtEtiquetaSup.Text.Substring(0, 2) = "LM" Then
            CadenaImpresoraSuperior = Chr(2) + "VQK#ALIAS1SUP=" + txtAlias1Sup.Text + "#ALIAS2EAN13=" + txtAlias2EANB.Text + "#ALIAS1LAT=" + txtAlias1Lat.Text + "#ALIAS6=" + txtalias6.Text + "#ALIAS6=" + txtalias6.Text + Chr(3) ' Repito la última variable porque de otra forma, la impresora no la coje
        Else
            CadenaImpresoraSuperior = Chr(2) + "VQK#ALIAS1SUP=" + txtAlias1Sup.Text + "#DESCRIPCION=" + txtDescripcion.Text + "#ALIAS6=" + txtalias6.Text + "#ALIAS2EAN13=" + txtAlias2EAN13.Text + "#ALIAS2EAN13=" + txtAlias2EAN13.Text + Chr(3) ' Repito la última variable porque de otra forma, la impresora no la coje
        End If

        txtDialogo.Text = txtDialogo.Text + "Enviando datos a " + txtImpresora1.Text + " "


        With WinSockCliente
            .IPDelHost = txtImpresora1.Text
            .PuertoDelHost = txtPuerto.Text
            .Conectar()
            .EnviarDatos(SeleccionarEtiquetaSuperior)
            System.Threading.Thread.Sleep(4000) ' se deja un tiempo de espera para que la impresora cargue la etiqueta
            .EnviarDatos(CadenaImpresoraSuperior)
            txtDialogo.Text = txtDialogo.Text + "OK" + txtImpresora1.Text + vbNewLine
            .Desconectar() ' RPM 16/05/2019
        End With

    End Sub

    Private Sub EnviarDatosImpresora2()
        Dim SeleccionarEtiquetaLateral As String
        Dim CadenaImpresoraLateral As String

        SeleccionarEtiquetaLateral = Chr(2) + "IMPR#" + txtEtiquetaLat.Text + Chr(3)
        'CadenaImpresoraLateral = Chr(2) + "VQK#ALIAS1LAT=" + txtAlias1Lat.Text + "#ALIAS2EAN13=" + txtAlias2EAN13.Text + "#ALIAS2EAN13=" + txtAlias2EAN13.Text + Chr(3) ' Repito la última variable porque de otra forma, la impresora no la coje
        If txtEtiquetaSup.Text.Substring(0, 2) = "LM" Then
            CadenaImpresoraLateral = Chr(2) + "VQK#ALIAS1SUP=" + txtAlias1Sup.Text + "#ALIAS2EAN13=" + txtAlias2EANB.Text + "#ALIAS1LAT=" + txtAlias1Lat.Text + "#ALIAS6=" + txtalias6.Text + "#ALIAS6=" + txtalias6.Text + Chr(3) ' Repito la última variable porque de otra forma, la impresora no la coje
        Else
            CadenaImpresoraLateral = Chr(2) + "VQK#ALIAS1LAT=" + txtAlias1Lat.Text + "#DESCRIPCION=" + txtDescripcion.Text + "#ALIAS2EAN13=" + txtAlias2EAN13.Text + "#ALIAS2EAN13=" + txtAlias2EAN13.Text + Chr(3) ' Repito la última variable porque de otra forma, la impresora no la coje
        End If

        txtDialogo.Text = txtDialogo.Text + "Enviando datos a " + txtImpresora2.Text + " "


        With WinSockCliente2

            .IPDelHost2 = txtImpresora2.Text
            .PuertoDelHost2 = txtPuerto.Text
            .Conectar2()
            .EnviarDatos2(SeleccionarEtiquetaLateral)
            System.Threading.Thread.Sleep(4000) ' se deja un tiempo de espera para que la impresora cargue la etiqueta
            .EnviarDatos2(CadenaImpresoraLateral)
            txtDialogo.Text += "OK" + txtImpresora2.Text + vbNewLine
            .Desconectar2() ' RPM 16/05/2019

        End With


    End Sub

    Private Sub EnviarDatosImpresora3()
        Dim SeleccionarEtiquetaLateral As String
        Dim CadenaImpresoraLateral As String

        SeleccionarEtiquetaLateral = Chr(2) + "IMPR#" + txtEtiquetaLat.Text + "#" + Chr(3) 'OJO, La impresora 10.106.5.119 exije terminar el comando con una #.
        '  CadenaImpresoraLateral = Chr(2) + "VQK#ALIAS1LAT=" + txtAlias1Lat.Text + "#ALIAS2EAN13=" + txtAlias2EAN13.Text + "#ALIAS2EAN13=" + txtAlias2EAN13.Text + Chr(3) ' Repito la última variable porque de otra forma, la impresora no la coje

        If txtEtiquetaSup.Text.Substring(0, 2) = "LM" Then
            CadenaImpresoraLateral = Chr(2) + "VQK#ALIAS1SUP=" + txtAlias1Sup.Text + "#ALIAS2EAN13=" + txtAlias2EANB.Text + "#ALIAS1LAT=" + txtAlias1Lat.Text + "#ALIAS6=" + txtalias6.Text + "#ALIAS6=" + txtalias6.Text + Chr(3) ' Repito la última variable porque de otra forma, la impresora no la coje
        Else
            CadenaImpresoraLateral = Chr(2) + "VQK#ALIAS1LAT=" + txtAlias1Lat.Text + "#DESCRIPCION=" + txtDescripcion.Text + "#ALIAS2EAN13=" + txtAlias2EAN13.Text + "#ALIAS2EAN13=" + txtAlias2EAN13.Text + Chr(3) ' Repito la última variable porque de otra forma, la impresora no la coje
        End If

        txtDialogo.Text = txtDialogo.Text + "Enviando datos a " + txtImpresora3.Text + " "


        With WinSockCliente3
            .IPDelHost3 = txtImpresora3.Text
            .PuertoDelHost3 = txtPuerto.Text
            .Conectar3()
            .EnviarDatos3(SeleccionarEtiquetaLateral)
            System.Threading.Thread.Sleep(4000) ' se deja un tiempo de espera para que la impresora cargue la etiqueta
            .EnviarDatos3(CadenaImpresoraLateral)
            txtDialogo.Text += "OK" + txtImpresora3.Text + vbNewLine
            .Desconectar3() ' RPM 16/05/2019
        End With


    End Sub

    Private Sub WinSockCliente_DatosRecibidos(ByVal datos As String) Handles WinSockCliente.DatosRecibidos
        'txtDialogo.Text = txtDialogo.Text + "Respuesta de la impresora: " & datos + Chr(10) + Chr(13)
    End Sub

    Private Sub WinSockCliente_ConexionTerminada() Handles WinSockCliente.ConexionTerminada
        ' txtDialogo.Text = txtDialogo.Text + "Conexion finalizada " + Chr(10) + Chr(13) da error
    End Sub


    Private Sub txtOF_LostFocus(sender As Object, e As EventArgs) Handles txtOF.LostFocus
        '*****RECUPERAR DATOS DEL ARTICULO A PARTIR DE LA O.F.
        Dim myConnection As OdbcConnection = New OdbcConnection()
        myConnection.ConnectionString = "DSN=IBM_MVXPROD;UID=CONSULTAS;Pwd=CONSULTAS"
        myConnection.Open()

        Dim Articulo As String
        Articulo = ""

        Dim Cliente As String
        Cliente = ""

        Dim SqlCommand As String
        Dim myCommand As New OdbcCommand()

        If txtOF.Text <> "" Then '

            myCommand.Connection = myConnection
            Dim myAdapter As New OdbcDataAdapter

            SqlCommand = "SELECT VHPRNO FROM MVXJDTA.MWOHED WHERE VHCONO = 1 AND VHFACI = 'RP0' AND VHMFNO = '" + txtOF.Text + "'"

            myCommand.CommandText = SqlCommand 'start query
            myAdapter.SelectCommand = myCommand
            Dim moddata As OdbcDataReader
            moddata = myCommand.ExecuteReader()
            While moddata.Read = True
                Try
                    Articulo = (moddata("VHPRNO"))
                Catch ex As Exception
                    txtDialogo.Text = txtDialogo.Text + "error al recuperar los datos del articulo desde la base de datos " + vbNewLine
                End Try
            End While

            moddata.Close()

            SqlCommand = " SELECT VMMTNO AS CodArticulo, MMFUDS AS DescripcionCaja,"
            SqlCommand = SqlCommand + " IFNULL(TPPIE.QJOPTN, '') AS CajaTipo, "
            SqlCommand = SqlCommand + " Int(IFNULL(SUBSTRING(LARGO.QJOPTN, 6, 4), 0)) As CajaLargo,"
            SqlCommand = SqlCommand + " INT(IFNULL(SUBSTRING(ANCHO.QJOPTN, 6, 4), 0)) AS CajaAncho,"
            SqlCommand = SqlCommand + " INT(IFNULL(SUBSTRING(GRUESO.QJOPTN, 6, 4), 0)) AS CajaGrueso,"
            SqlCommand = SqlCommand + " IFNULL(D1.MPPOPN, '') AS DesEtiLateral,"
            SqlCommand = SqlCommand + " IFNULL(D2.MPPOPN, '') AS DesEtiSuperior,"
            SqlCommand = SqlCommand + " CASE WHEN SUBSTRING(SCJLA.QKMEVA, 15, 1)='&' THEN -1 ELSE 1 END * CAST(IFNULL(SUBSTRING(SCJLA.QKMEVA, 1, 15), 0) AS DOUBLE)/1000000 AS SupLargo,"
            SqlCommand = SqlCommand + " CASE WHEN SUBSTRING(SCJAN.QKMEVA, 15, 1)='&' THEN -1 ELSE 1 END * CAST(IFNULL(SUBSTRING(SCJLA.QKMEVA, 1, 15), 0) AS DOUBLE)/1000000 AS SupAncho,"
            SqlCommand = SqlCommand + " CASE WHEN SUBSTRING(SCJGR.QKMEVA, 15, 1)='&' THEN -1 ELSE 1 END * CAST(IFNULL(SUBSTRING(SCJLA.QKMEVA, 1, 15), 0) AS DOUBLE)/1000000 AS SupGrueso"
            SqlCommand = SqlCommand + " FROM MVXJDTA.MWOHED"
            SqlCommand = SqlCommand + " LEFT JOIN MVXJDTA.MWOMAT ON VMCONO=1 AND VMFACI='RP0' AND VMMFNO=VHMFNO AND VMPRNO=VHPRNO"
            SqlCommand = SqlCommand + " LEFT JOIN MVXJDTA.MITMAS ON MMCONO=1 AND MMITNO=VMMTNO"
            SqlCommand = SqlCommand + " LEFT JOIN MVXJDTA.MPDVAN ON PVCONO=1 AND PVFACI='RP0' AND PVSTRT='001' AND PVPRNO=MMHDPR AND PVVANO=VMMTNO"
            SqlCommand = SqlCommand + " LEFT JOIN MVXJDTA.MPDCDF TPPIE ON TPPIE.QJCONO=1 AND TPPIE.QJECVS=0 AND TPPIE.QJCFIN=PVCFIN AND TPPIE.QJFTID='TPPIE'"
            SqlCommand = SqlCommand + " LEFT JOIN MVXJDTA.MPDCDF LARGO ON LARGO.QJCONO=1 AND LARGO.QJECVS=0 AND LARGO.QJCFIN=PVCFIN AND LARGO.QJFTID='LARGO'"
            SqlCommand = SqlCommand + " LEFT JOIN MVXJDTA.MPDCDF ANCHO ON ANCHO.QJCONO=1 AND ANCHO.QJECVS=0 AND ANCHO.QJCFIN=PVCFIN AND ANCHO.QJFTID='ANCHO'"
            SqlCommand = SqlCommand + " LEFT JOIN MVXJDTA.MPDCDF GRUESO ON GRUESO.QJCONO=1 AND GRUESO.QJECVS=0 AND GRUESO.QJCFIN=PVCFIN AND GRUESO.QJFTID='GRUES'"
            SqlCommand = SqlCommand + " LEFT JOIN MVXJDTA.MITPOP D1 ON D1.MPCONO=1 AND D1.MPITNO=VHPRNO AND D1.MPALWT=1 AND D1.MPALWQ='' AND D1.MPSEQN='10' AND D1.MPREMK='LATERAL'"
            SqlCommand = SqlCommand + " LEFT JOIN MVXJDTA.MITPOP D2 ON D2.MPCONO=1 AND D2.MPITNO=VHPRNO AND D2.MPALWT=1 AND D2.MPALWQ='' AND D2.MPSEQN='20' AND D2.MPREMK='SUPERIOR'"
            SqlCommand = SqlCommand + " LEFT JOIN MVXJDTA.MPDCDM SCJLA ON SCJLA.QKCONO=1 AND SCJLA.QKECVS=0 AND SCJLA.QKCFIN=VHCFIN AND SCJLA.QKDMID='SCJLA' AND SCJLA.QKPRNO=CASE WHEN VHHDPR='' THEN VHPRNO ELSE VHHDPR END"
            SqlCommand = SqlCommand + " LEFT JOIN MVXJDTA.MPDCDM SCJAN ON SCJAN.QKCONO=1 AND SCJAN.QKECVS=0 AND SCJAN.QKCFIN=VHCFIN AND SCJAN.QKDMID='SCJAN' AND SCJAN.QKPRNO=CASE WHEN VHHDPR='' THEN VHPRNO ELSE VHHDPR END"
            SqlCommand = SqlCommand + " LEFT JOIN MVXJDTA.MPDCDM SCJGR ON SCJGR.QKCONO=1 AND SCJGR.QKECVS=0 AND SCJGR.QKCFIN=VHCFIN AND SCJGR.QKDMID='SCJGR' AND SCJGR.QKPRNO=CASE WHEN VHHDPR='' THEN VHPRNO ELSE VHHDPR END"
            SqlCommand = SqlCommand + " WHERE VHCONO=1 AND VHFACI='RP0' AND MMITTY='EMB' AND VHMFNO='" + txtOF.Text + "'"

            myCommand.CommandText = SqlCommand 'start query
            myAdapter.SelectCommand = myCommand
            moddata = myCommand.ExecuteReader()
            While moddata.Read = True
                Try
                    txtDescripcionCaja.Text = RTrim(moddata("DescripcionCaja"))
                    txtTipoCaja.Text = RTrim(moddata("CajaTipo"))
                    txtAlias1Sup.Text = RTrim(moddata("DesEtiSuperior"))
                    txtAlias1Lat.Text = RTrim(moddata("DesEtiLateral"))
                    txtLargo.Text = (moddata("CajaLargo"))
                    txtAncho.Text = (moddata("CajaAncho"))
                    txtGrueso.Text = (moddata("CajaGrueso"))
                    txtSuplementoLargo.Text = (moddata("SupLargo"))
                    txtSuplementoAncho.Text = (moddata("SupAncho"))
                    txtSuplementoGrueso.Text = (moddata("SupGrueso"))
                Catch ex As Exception
                    txtDialogo.Text = txtDialogo.Text + "error al recuperar los datos de la O.F. desde la base de datos " + vbNewLine
                End Try
            End While

            moddata.Close()

            SqlCommand = "SELECT MPPOPN FROM MVXJDTA.MITPOP WHERE MPCONO = 1  AND MPALWT = '02' AND MPITNO = '" + Articulo + "'"

            myCommand.CommandText = SqlCommand 'start query
            myAdapter.SelectCommand = myCommand
            moddata = myCommand.ExecuteReader()
            While moddata.Read = True
                Try
                    txtAlias2EAN13.Text = RTrim(moddata("MPPOPN"))
                    txtAlias2EANB.Text = RTrim(moddata("MPPOPN"))
                Catch ex As Exception
                    txtDialogo.Text = txtDialogo.Text + "error al recuperar los datos de la variable ALIAS2EANB desde la base de datos " + vbNewLine
                End Try
            End While

            moddata.Close()

            SqlCommand = "SELECT MPPOPN,MPE0PA FROM MVXJDTA.MITPOP WHERE MPCONO = 1  AND MPALWT = '06' AND MPITNO = '" + Articulo + "'"

            myCommand.CommandText = SqlCommand 'start query
            myAdapter.SelectCommand = myCommand
            moddata = myCommand.ExecuteReader()
            While moddata.Read = True
                Try
                    txtalias6.Text = RTrim(moddata("MPPOPN"))
                    Cliente = RTrim(moddata("MPE0PA"))
                Catch ex As Exception
                    txtDialogo.Text = txtDialogo.Text + "error al recuperar los datos de la variable ASLIAS6 desde la base de datos " + vbNewLine
                End Try
            End While

            moddata.Close()

            SqlCommand = "SELECT QBMRC1 FROM MVXJDTA.MPMXVA WHERE QBCONO = 1 AND QBMXID = 'ETQGR' AND QBMVC1 = '" + Cliente + "'"

            myCommand.CommandText = SqlCommand 'start query
            myAdapter.SelectCommand = myCommand
            moddata = myCommand.ExecuteReader()
            While moddata.Read = True
                Try
                    txtEtiquetaLat.Text = RTrim(moddata("QBMRC1")) + "LAT.ltg"
                    txtEtiquetaSup.Text = RTrim(moddata("QBMRC1")) + "SUP.ltg"
                Catch ex As Exception
                    txtDialogo.Text = txtDialogo.Text + "error al recuperar los datos de la variable ASLIAS6 desde la base de datos " + vbNewLine
                End Try
            End While

            moddata.Close()

            SqlCommand = "SELECT MAX(ORTEDS) AS ORTEDS FROM MVXJDTA.OCUSIT WHERE ORCONO = 1 AND ORITNO = '" + Articulo + "'"

            myCommand.CommandText = SqlCommand 'start query
            myAdapter.SelectCommand = myCommand
            moddata = myCommand.ExecuteReader()
            While moddata.Read = True
                Try
                    txtDescripcion.Text = RTrim(moddata("ORTEDS"))

                Catch ex As Exception
                    txtDialogo.Text = txtDialogo.Text + "error al recuperar los datos de la variable DESCRIPCION desde la base de datos" + vbNewLine
                End Try
            End While

            moddata.Close()

            myConnection.Close()
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        WinSockCliente.Desconectar()
        WinSockCliente2.Desconectar2()
        WinSockCliente3.Desconectar3()

        System.Windows.Forms.Application.Exit()
        'RemoveHandler MyBase.FormClosing, AddressOf FormLimitronic_FormClosing
        'RemoveHandler MyBase.FormClosed, AddressOf FormLimitronic_FormClosed
        'Me.Close()
    End Sub
End Class

'*************************************************************************
' SOCKET PARA LA IMPRESORA 1
'*************************************************************************

Public Class Cliente

#Region "VARIABLES"
    Private Stm As Stream 'Utilizado para enviar datos al Servidor y recibir datos del mismo
    Private m_IPDelHost As String 'Direccion del objeto de la clase Servidor
    Private m_PuertoDelHost As String 'Puerto donde escucha el objeto de la clase Servidor
    Public Mensaje As String 'Mensajes del socket para motrar en el textbox de mensajes de comunicaciones
#End Region

#Region "EVENTOS"
    Public Event ConexionTerminada()
        Public Event DatosRecibidos(ByVal datos As String)
#End Region

#Region "PROPIEDADES"
        Public Property IPDelHost() As String
            Get
                IPDelHost = m_IPDelHost
            End Get

            Set(ByVal Value As String)
                m_IPDelHost = Value
            End Set
        End Property

        Public Property PuertoDelHost() As String
            Get
                PuertoDelHost = m_PuertoDelHost
            End Get

            Set(ByVal Value As String)
                m_PuertoDelHost = Value
            End Set
        End Property

#End Region

#Region "METODOS"

    Public Sub Conectar()
        Dim tcpClnt As TcpClient
        Dim tcpThd As Thread 'Se encarga de escuchar mensajes enviados por el Servidor
        tcpClnt = New TcpClient()

        tcpClnt.Connect(IPDelHost, PuertoDelHost)

        Stm = tcpClnt.GetStream()
        'Creo e inicio un thread para que escuche los mensajes enviados por el Servidor
        tcpThd = New Thread(AddressOf LeerSocket)
        tcpThd.Start()
    End Sub

    Public Sub EnviarDatos(ByVal Datos As String)
        Dim BufferDeEscritura() As Byte
        BufferDeEscritura = Encoding.Unicode.GetBytes(Datos)
        If Not (Stm Is Nothing) Then
            'Envio los datos al Servidor
            Stm.Write(BufferDeEscritura, 0, BufferDeEscritura.Length)

        End If
    End Sub

    Public Sub Desconectar()
        Stm.Close()
    End Sub

#End Region

#Region "FUNCIONES PRIVADAS"

    Private Sub LeerSocket()
        Dim BufferDeLectura() As Byte
        While True
            Try
                BufferDeLectura = New Byte(100) {}
                'Me quedo esperando a que llegue algun mensaje
                Stm.Read(BufferDeLectura, 0, BufferDeLectura.Length)
                'Genero el evento DatosRecibidos, ya que se han recibido datos desde el Servidor
                RaiseEvent DatosRecibidos(Encoding.Unicode.GetString(BufferDeLectura))

                Mensaje = Encoding.Unicode.GetString(BufferDeLectura)

            Catch e As Exception
                Exit While
            End Try
        End While
        'Finalizo la conexion, por lo tanto genero el evento correspondiente
        RaiseEvent ConexionTerminada()
    End Sub

#End Region

End Class
'*************************************************************************
' SOCKET PARA LA IMPRESORA 2
'*************************************************************************

Public Class Cliente2

#Region "VARIABLES"
    Private Stm2 As Stream 'Utilizado para enviar datos al Servidor y recibir datos del mismo
    Private m_IPDelHost2 As String 'Direccion del objeto de la clase Servidor
    Private m_PuertoDelHost2 As String 'Puerto donde escucha el objeto de la clase Servidor
    Public Mensaje2 As String 'Mensajes del socket para motrar en el textbox de mensajes de comunicaciones
#End Region

#Region "EVENTOS"
    Public Event ConexionTerminada2()
    Public Event DatosRecibidos2(ByVal datos2 As String)
#End Region

#Region "PROPIEDADES"
    Public Property IPDelHost2() As String
        Get
            IPDelHost2 = m_IPDelHost2
        End Get

        Set(ByVal Value2 As String)
            m_IPDelHost2 = Value2
        End Set
    End Property

    Public Property PuertoDelHost2() As String
        Get
            PuertoDelHost2 = m_PuertoDelHost2
        End Get

        Set(ByVal Value2 As String)
            m_PuertoDelHost2 = Value2
        End Set
    End Property

#End Region

#Region "METODOS"

    Public Sub Conectar2()
        Dim tcpClnt2 As TcpClient
        Dim tcpThd2 As Thread 'Se encarga de escuchar mensajes enviados por el Servidor
        tcpClnt2 = New TcpClient()

        tcpClnt2.Connect(IPDelHost2, PuertoDelHost2)

        Stm2 = tcpClnt2.GetStream()
        'Creo e inicio un thread para que escuche los mensajes enviados por el Servidor
        tcpThd2 = New Thread(AddressOf LeerSocket2)
        tcpThd2.Start()
    End Sub

    Public Sub EnviarDatos2(ByVal Datos2 As String)
        Dim BufferDeEscritura2() As Byte
        BufferDeEscritura2 = Encoding.Unicode.GetBytes(Datos2)
        If Not (Stm2 Is Nothing) Then
            'Envio los datos al Servidor
            Stm2.Write(BufferDeEscritura2, 0, BufferDeEscritura2.Length)

        End If
    End Sub
    Public Sub Desconectar2()
        Stm2.Close()
    End Sub


#End Region

#Region "FUNCIONES PRIVADAS"

    Private Sub LeerSocket2()
        Dim BufferDeLectura2() As Byte
        While True
            Try
                BufferDeLectura2 = New Byte(100) {}
                'Me quedo esperando a que llegue algun mensaje
                Stm2.Read(BufferDeLectura2, 0, BufferDeLectura2.Length)
                'Genero el evento DatosRecibidos, ya que se han recibido datos desde el Servidor
                RaiseEvent DatosRecibidos2(Encoding.Unicode.GetString(BufferDeLectura2))

                Mensaje2 = Encoding.Unicode.GetString(BufferDeLectura2)

            Catch e2 As Exception
                Exit While
            End Try
        End While
        'Finalizo la conexion, por lo tanto genero el evento correspondiente
        RaiseEvent ConexionTerminada2()
    End Sub

#End Region

End Class


'*************************************************************************
' SOCKET PARA LA IMPRESORA 3
'*************************************************************************

Public Class Cliente3

#Region "VARIABLES"
    Private Stm3 As Stream 'Utilizado para enviar datos al Servidor y recibir datos del mismo
    Private m_IPDelHost3 As String 'Direccion del objeto de la clase Servidor
    Private m_PuertoDelHost3 As String 'Puerto donde escucha el objeto de la clase Servidor
    Public Mensaje3 As String 'Mensajes del socket para motrar en el textbox de mensajes de comunicaciones
#End Region

#Region "EVENTOS"
    Public Event ConexionTerminada3()
    Public Event DatosRecibidos3(ByVal datos3 As String)
#End Region

#Region "PROPIEDADES"
    Public Property IPDelHost3() As String
        Get
            IPDelHost3 = m_IPDelHost3
        End Get

        Set(ByVal Value3 As String)
            m_IPDelHost3 = Value3
        End Set
    End Property

    Public Property PuertoDelHost3() As String
        Get
            PuertoDelHost3 = m_PuertoDelHost3
        End Get

        Set(ByVal Value3 As String)
            m_PuertoDelHost3 = Value3
        End Set
    End Property

#End Region

#Region "METODOS"

    Public Sub Conectar3()
        Dim tcpClnt3 As TcpClient
        Dim tcpThd3 As Thread 'Se encarga de escuchar mensajes enviados por el Servidor
        tcpClnt3 = New TcpClient()

        tcpClnt3.Connect(IPDelHost3, PuertoDelHost3)

        Stm3 = tcpClnt3.GetStream()
        'Creo e inicio un thread para que escuche los mensajes enviados por el Servidor
        tcpThd3 = New Thread(AddressOf LeerSocket3)
        tcpThd3.Start()
    End Sub

    Public Sub EnviarDatos3(ByVal Datos3 As String)
        Dim BufferDeEscritura3() As Byte
        BufferDeEscritura3 = Encoding.Unicode.GetBytes(Datos3)
        If Not (Stm3 Is Nothing) Then
            'Envio los datos al Servidor
            Stm3.Write(BufferDeEscritura3, 0, BufferDeEscritura3.Length)

        End If
    End Sub
    Public Sub Desconectar3()
        Stm3.Close()
    End Sub


#End Region

#Region "FUNCIONES PRIVADAS"

    Private Sub LeerSocket3()
        Dim BufferDeLectura3() As Byte
        While True
            Try
                BufferDeLectura3 = New Byte(100) {}
                'Me quedo esperando a que llegue algun mensaje
                Stm3.Read(BufferDeLectura3, 0, BufferDeLectura3.Length)
                'Genero el evento DatosRecibidos, ya que se han recibido datos desde el Servidor
                RaiseEvent DatosRecibidos3(Encoding.Unicode.GetString(BufferDeLectura3))

                Mensaje3 = Encoding.Unicode.GetString(BufferDeLectura3)

            Catch e3 As Exception
                Exit While
            End Try
        End While
        'Finalizo la conexion, por lo tanto genero el evento correspondiente
        RaiseEvent ConexionTerminada3()
    End Sub

#End Region

End Class


'*************************************************************************
' SOCKET PARA PLC
'*************************************************************************

Public Class Cliente4

#Region "VARIABLES"
    Private Stm4 As Stream 'Utilizado para enviar datos al Servidor y recibir datos del mismo
    Private m_IPDelHost4 As String 'Direccion del objeto de la clase Servidor
    Private m_PuertoDelHost4 As String 'Puerto donde escucha el objeto de la clase Servidor
    Public Mensaje4 As String 'Mensajes del socket para motrar en el textbox de mensajes de comunicaciones
#End Region

#Region "EVENTOS"
    Public Event ConexionTerminada4()
    Public Event DatosRecibidos4(ByVal datos4 As String)
#End Region

#Region "PROPIEDADES"
    Public Property IPDelHost4() As String
        Get
            IPDelHost4 = m_IPDelHost4
        End Get

        Set(ByVal Value4 As String)
            m_IPDelHost4 = Value4
        End Set
    End Property

    Public Property PuertoDelHost4() As String
        Get
            PuertoDelHost4 = m_PuertoDelHost4
        End Get

        Set(ByVal Value4 As String)
            m_PuertoDelHost4 = Value4
        End Set
    End Property

#End Region

#Region "METODOS"

    Public Sub Conectar4()
        Dim tcpClnt4 As TcpClient
        'Dim tcpThd4 As Thread 'Se encarga de escuchar mensajes enviados por el Servidor
        tcpClnt4 = New TcpClient()

        tcpClnt4.Connect(IPDelHost4, PuertoDelHost4)

        Stm4 = tcpClnt4.GetStream()
        'Creo e inicio un thread para que escuche los mensajes enviados por el Servidor
        'tcpThd4 = New Thread(AddressOf LeerSocket4)
        'tcpThd4.Start()
    End Sub

    Public Sub EnviarDatos4(ByVal Datos4 As Byte())
        Dim BufferDeEscritura4() As Byte
        BufferDeEscritura4 = Datos4
        If Not (Stm4 Is Nothing) Then
            'Envio los datos al Servidor
            Stm4.Write(BufferDeEscritura4, 0, BufferDeEscritura4.Length)

        End If
    End Sub
    Public Sub Desconectar4()
        Stm4.Close()
    End Sub


#End Region

#Region "FUNCIONES PRIVADAS"

    Private Sub LeerSocket4()
        Dim BufferDeLectura4() As Byte
        While True
            Try
                BufferDeLectura4 = New Byte(100) {}
                'Me quedo esperando a que llegue algun mensaje
                Stm4.Read(BufferDeLectura4, 0, BufferDeLectura4.Length)
                'Genero el evento DatosRecibidos, ya que se han recibido datos desde el Servidor
                RaiseEvent DatosRecibidos4(Encoding.Unicode.GetString(BufferDeLectura4))

                Mensaje4 = Encoding.Unicode.GetString(BufferDeLectura4)

            Catch e3 As Exception
                Exit While
            End Try
        End While
        'Finalizo la conexion, por lo tanto genero el evento correspondiente
        RaiseEvent ConexionTerminada4()
    End Sub

#End Region

End Class
