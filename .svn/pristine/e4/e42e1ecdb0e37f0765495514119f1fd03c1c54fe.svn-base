'
Imports System.Data.SqlClient
Imports System.Data.OleDb
'
Module SEI_Globals

    Public go_connOLEDB As OleDbConnection
    Public go_conn As SqlConnection
    Public oCompany As SAPbobsCOM.Company

#Region "Funciones Fichero.INI"

    '
    ' Leer una clave de un fichero INI
    Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
    Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Integer, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
    '
    Public Function IniGet(ByVal sFileName As String, ByVal sSection As String, ByVal sKeyName As String, Optional ByVal sDefault As String = "") As String
        '--------------------------------------------------------------------------
        ' Devuelve el valor de una clave de un fichero INI
        ' Los parámetros son:
        '   sFileName   El fichero INI
        '   sSection    La sección de la que se quiere leer
        '   sKeyName    Clave
        '   sDefault    Valor opcional que devolverá si no se encuentra la clave
        '--------------------------------------------------------------------------
        ' sSection ->   "Parametros"
        ' sKeyName ->   "U" , "I" , "P"
        '
        ' [Parametros]
        ' U = sa
        ' I = IG
        ' P =seidor.65

        Dim ret As Integer
        Dim sRetVal As String
        '
        sRetVal = New String(Chr(0), 255)
        '
        ret = GetPrivateProfileString(sSection, sKeyName, sDefault, sRetVal, Len(sRetVal), sFileName)
        If ret = 0 Then
            Return sDefault
        Else
            Return Left(sRetVal, ret)
        End If
    End Function

#End Region


#Region "Funciones Tipos de Datos"

    Function Formato_Decimales_IG(ByVal Valor As Object) As String

        Valor = Valor.ToString.Replace(".", "")
        Valor = Valor.ToString.Replace(",", ".")
        Return Valor.ToString

    End Function

    Function NullToText(ByVal Valor As Object) As String

        If IsDBNull(Valor) Or Valor.ToString.Trim = "" Then
            Return " "
        Else
            Return Valor.ToString
        End If

    End Function

    Function NullToInt(ByVal Valor As Object) As Integer

        If IsDBNull(Valor) Or Valor.ToString.Trim = "" Then
            Return 0
        Else
            Return Convert.ToInt32(Valor.ToString)   ' Pasar a integer
        End If

    End Function

    Function NullToDoble(ByRef Valor As Object) As Double

        If IsDBNull(Valor) Or Trim(Valor.ToString) = "" Then
            Return 0
        Else
            Return Convert.ToDouble(Valor.ToString)  ' Pasar a double
        End If

    End Function

    Function NullToData(ByVal Valor As Object) As String

        If IsDBNull(Valor) Or Valor.GetType.ToString = "" Then
            Return "NULL"
        Else
            Return String.Format("{0:d}", Valor)
        End If

    End Function

    Function NullToHora(ByVal Valor As Object) As String

        If IsDBNull(Valor) Or Valor.GetType.ToString = "" Then
            Return "NULL"
        Else
            Return String.Format("{0:t}", Valor)
        End If

    End Function

    Function NullToLong(ByVal Valor As Object) As Long

        If IsDBNull(Valor) Or Trim(Valor.ToString) = "" Then
            Return 0
        Else
            Return CType(Valor, Long)   ' Pasar a Long
        End If

    End Function

    Function NullToSiNo(ByVal Valor As Object) As String

        If IsDBNull(Valor) Or Valor.GetType.ToString = "" Then
            Return "N"
        Else
            Return "Y"
        End If

    End Function

    Function IntToBooleanS_N(ByVal Valor As Object) As String

        If IsDBNull(Valor) Then
            Return "N"
        ElseIf Valor = 0 Then
            Return "N"
        Else
            Return "S"
        End If

    End Function

    Function IntToBoolean(ByVal Valor As Object) As String

        If IsDBNull(Valor) Then
            Return "N"
        ElseIf Valor = 0 Then
            Return "N"
        Else
            Return "Y"
        End If

    End Function

    Function BooleanToInt(ByVal Valor As Object) As Integer

        If Valor = "True" Then
            Return "1"
        Else
            Return "0"
        End If

    End Function
    '
    Public Function NowDateToString() As String
        NowDateToString = Now.Date.ToString("yyyyMMdd")
    End Function
    '
    ' Poner un valor entre comillas
    Public Function sC(ByVal sValor As String) As String
        sC = "'" & sValor.Replace("'", "''") & "'"
    End Function

#End Region

    '
    Public Sub LiberarObjCOM(ByRef oObjCOM As Object, Optional ByVal bCollect As Boolean = False)
        '
        'Liberar y destruir Objecto com 
        ' En los UDO'S es necesario utilizar GC.Collect  para eliminarlos de la memoria
        If Not IsNothing(oObjCOM) Then
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oObjCOM)
            oObjCOM = Nothing
            If bCollect Then
                GC.Collect()
            End If
        End If

    End Sub
    '
    Public Function RecuperarErrorSap() As String

        Dim sError As String
        Dim lErrCode As Long
        Dim sErrMsg As String
        '
        lErrCode = 0
        sErrMsg = ""
        oCompany.GetLastError(lErrCode, sErrMsg)
        sError = "Error: " & lErrCode.ToString & " " & sErrMsg
        '
        Return sError
        '
    End Function

End Module
