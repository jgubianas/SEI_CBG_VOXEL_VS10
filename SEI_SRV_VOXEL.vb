Option Explicit On
'
Imports System.Data.SqlClient
Imports System.Text
Imports System.IO
Imports SAPbobsCOM.BoDataServerTypes
Imports SAPbobsCOM.BoSuppLangs

Public Class SEI_SRV_VOXEL

    Private dInicio As Date
    Private dFinal As Date
    Private lResultado As Long

    Private Sub btnEjecutar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEjecutar.Click

        Me.lblmsg.Text = ""
        '
        dInicio = Now

        If Not Me.ConectarSQLNative(go_conn) Then
            Exit Sub
        End If

        ' '' '' ''If Not ConectarSBO() Then
        ' '' '' ''    Me.DesconectarSQLNative()
        ' '' '' ''    Exit Sub
        ' '' '' ''End If
        '

        If Me.chkAlbarans.Checked Then
            txtProceso.Text = "Generando Albaranes"
            AlbaransElectronics()
        End If

        If Me.chkFacturacion.Checked Then
            txtProceso.Text = "Facturación"
            FacturacionElectronica()
            txtProceso.Text = "Abonos"
            AbonamentsElectronics()
        End If


        If Me.chkConfirmacions.Checked Then
            txtProceso.Text = "Leiendo ficheros de aceptación de facturas "
            AcetpacionFacturasElectronicas()
        End If


        '
        dFinal = Now
        lResultado = DateDiff(DateInterval.Minute, dInicio, dFinal)
        Me.lblmsg.Text = "Facturación finalizada. Minutos: " & lResultado.ToString
        '
        '-------------------------------------------------------------------------------
        ' DESECONEXIONES DE LAS BASES DE DATOS
        '-------------------------------------------------------------------------------
        Me.DesconectarSQLNative()
        ' '' ''Me.DesconectarSB0()
        Me.Dispose()

    End Sub

    Public Function ConectarSQLNative(ByRef go_connaux As SqlConnection) As Boolean
        '
        Dim ls As String = ""
        '
        ConectarSQLNative = False
        Try
            '
            'go_conn.ConnectionString = "Server=JJM\SBO_2005;Database=Sap;User ID=sa;Password=seidor.65;Connect Timeout=120"
            ls = "Server=" & IniGet(Application.StartupPath & "\S_SEI_CBG_VOXEL.ini", "Parametros", "S") & ";" & _
                 "Database=" & IniGet(Application.StartupPath & "\S_SEI_CBG_VOXEL.ini", "Parametros", "D") & ";" & _
                 "User id=" & IniGet(Application.StartupPath & "\S_SEI_CBG_VOXEL.ini", "Parametros", "U") & ";" & _
                 "Password=" & IniGet(Application.StartupPath & "\S_SEI_CBG_VOXEL.ini", "Parametros", "P") & ";"
            '
            go_connaux = New SqlConnection
            go_connaux.ConnectionString = ls
            go_connaux.Open()

            Me.txtUsuario.Text = IniGet(Application.StartupPath & "\S_SEI_CBG_VOXEL.ini", "Parametros", "U")
            Me.txtempresa.Text = IniGet(Application.StartupPath & "\S_SEI_CBG_VOXEL.ini", "Parametros", "D")

            ConectarSQLNative = True

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        '
    End Function
    '
    Private Function ConectarSBO() As Boolean

        Dim lRetCode As Long
        Dim lErrCode As Long
        Dim sErrMsg As String
        '
        ConectarSBO = False
        lRetCode = 0
        sErrMsg = ""
        '
        oCompany = New SAPbobsCOM.Company
        oCompany.Server = IniGet(Application.StartupPath & "\S_SEI_CBG_VAOXEL.ini", "Parametros", "S")    ' Server
        oCompany.CompanyDB = IniGet(Application.StartupPath & "\S_SEI_CBG_VOXEL.ini", "Parametros", "B") ' Base de Dades
        oCompany.UserName = "manager" 'User SBO
        oCompany.Password = "1234"    'Password SBO
        oCompany.DbUserName = IniGet(Application.StartupPath & "\S_SEI_CBG_VOXEL.ini", "Parametros", "U")       'User BD
        oCompany.DbPassword = IniGet(Application.StartupPath & "\S_SEI_CBG_VOXEL.ini", "Parametros", "P")       'Password BD
        oCompany.UseTrusted = IniGet(Application.StartupPath & "\S_SEI_CBG_VOXEL.ini", "Parametros", "Trusted") 'Trusted
        oCompany.language = ln_Spanish
        oCompany.DbServerType = dst_MSSQL2005
        '
        '// Connecting to a company DB
        lRetCode = oCompany.Connect

        If lRetCode <> 0 Then
            Me.txtempresa.Text = oCompany.CompanyName
            Me.txtUsuario.Text = oCompany.UserName
            Application.DoEvents()
        Else
            '
            Me.txtempresa.Text = oCompany.CompanyName
            Me.txtUsuario.Text = oCompany.UserName
            '
            oCompany.GetLastError(lErrCode, sErrMsg)
            'grabar log de errores
            '
            Me.lblmsg.Text = "Error: " & lErrCode.ToString & " " & sErrMsg
            Me.lblmsg.BackColor = Color.Red
            Application.DoEvents()
        End If

        '// Use Windows authentication for database server.
        '// True for NT server authentication,
        '// False for database server authentication.
        'oCompany.UseTrusted = True

    End Function
    '
    Private Sub DesconectarSQLNative()
        go_conn.Close()
    End Sub

    Private Sub DesconectarSB0()
        oCompany.Disconnect()
    End Sub
    '
    Private Sub FacturacionElectronica()
        Dim oFacturas As SEI_Facturas
        oFacturas = New SEI_Facturas(Me)
        oFacturas.GENERAR_FACTURES_TLY()
    End Sub


    Private Sub AlbaransElectronics()
        Dim oAlbarans As SEI_Albarans
        oAlbarans = New SEI_Albarans(Me)
        oAlbarans.GENERAR_ALBARANS_TLY()
    End Sub

    Private Sub AbonamentsElectronics()
        Dim oAbonament As SEI_Abonaments
        oAbonament = New SEI_Abonaments(Me)
        oAbonament.GENERAR_ABONAMENTS_TLY()
    End Sub

    '
    Private Sub AcetpacionFacturasElectronicas()
        Dim oAcceptacio As SEI_AcceptacioF
        oAcceptacio = New SEI_AcceptacioF(Me)
        oAcceptacio.LLEGIR_ACCEPTACIONSF_TLY()
    End Sub
    '
    Private Sub SEI_SRV_VOXEL_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        chkFacturacion.Checked = True
        chkAlbarans.Checked = True
        chkConfirmacions.Checked = True
    End Sub
    '
    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkConfirmacions.CheckedChanged
    End Sub
    '
    Private Sub SEI_SRV_VOXEL_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown
        If IniGet(Application.StartupPath & "\S_SEI_CBG_VOXEL.ini", "Parametros", "A") = "S" Then
            btnEjecutar_Click(sender, e)
        End If
    End Sub
    '
    Private Sub btnSalir_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSalir.Click
        Me.Dispose()
    End Sub

    Private Sub chkFacturacion_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkFacturacion.CheckedChanged

    End Sub
End Class
