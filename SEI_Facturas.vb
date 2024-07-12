'
Option Explicit On
'
Imports System.Data.SqlClient
Imports System.Text
Imports System.IO
Imports SAPbobsCOM.BoObjectTypes

Imports System.Collections.Generic
Imports System.Collections.ObjectModel

Imports System.Reflection
Imports System.Xml
Imports System.Windows.Forms

 
Imports System.Collections
Imports System.Net


Imports System
Imports Microsoft.VisualBasic

Imports ADODB

Public Class SEI_Facturas
    Private Form As SEI_SRV_VOXEL

#Region "Contructor"
    '
    Public Sub New(ByRef o_Form As SEI_SRV_VOXEL)
        Form = o_Form
        'if creating controls via code, use initialize
        Initialize()
    End Sub

    Private Sub Initialize()

    End Sub
    '
    'Public Sub GENERAR_FACTURES_TLY()
    '    Dim ls As String
    '    Dim oSqlcomand As SqlCommand
    '    Dim oDataReader As SqlClient.SqlDataReader = Nothing
    '    Dim oDataReader2 As SqlClient.SqlDataReader = Nothing
    '    Dim oRcsFactLin As SqlClient.SqlDataReader = Nothing
    '    Dim oXml As XmlDocument = Nothing
    '    Dim oItem As System.Xml.XmlNodeList = Nothing
    '    Dim sPath, sPathD As String
    '    Dim sFichero, sFicheroD As String
    '    Dim numSerie As String = "0"
    '    Dim iFila As Integer
    '    Dim go_conn3 As SqlConnection = Nothing
    '    Dim HashFEnviats As Hashtable = New Hashtable
    '    Dim HashResumIVAS As Hashtable = New Hashtable
    '    Dim Conn1, conn2 As New ADODB.Connection
    '    Dim Cmd1 As New ADODB.Command
    '    Dim oRecordset, orecordset2 As ADODB.Recordset '''SAPbobsCOM.Recordset
    '    '
    '    Dim esHotelsCatalonia As Boolean = False
    '    Dim esAngrup As Boolean = False

    '    '' dades de CBG
    '    Dim xCIF As String = ""
    '    Dim xAdreca As String = ""
    '    Dim xCompany As String = ""
    '    Dim xAddress As String = ""
    '    Dim xCity As String = ""
    '    Dim xProvince As String = ""
    '    Dim xCountry As String = ""
    '    Dim xZipCode As String = ""
    '    '' 
    '    ls = ""
    '    ls = ls & " select top 1 OADM.TaxIdNum,OADM.CompnyName,OADM.CompnyAddr, adm1.City, adm1.State, adm1.Country, adm1.ZipCode, OCST.Name as nomprov from OADM "
    '    ls = ls & " left join ADM1 on oadm.currPeriod = adm1.currPeriod  "
    '    ls = ls & "   left join OCST   on OCST.code = adm1.State "
    '    oSqlcomand = New SqlCommand(ls, go_conn)
    '    oDataReader = oSqlcomand.ExecuteReader()
    '    While oDataReader.Read()
    '        xCIF = oDataReader("TaxIdNum").ToString.Substring(2, 9)
    '        xCompany = "COMERCIAL CBG S.A."   '''''oDataReader("CompnyName").ToString
    '        xAddress = oDataReader("CompnyAddr").ToString
    '        xCity = oDataReader("City").ToString
    '        xProvince = oDataReader("nomprov").ToString
    '        If oDataReader("Country").ToString = "ES" Then
    '            xCountry = "ESP"
    '        Else
    '            xCountry = oDataReader("Country").ToString
    '        End If
    '        xZipCode = oDataReader("ZipCode").ToString
    '    End While
    '    ''''''' fi agafar dades cbg 
    '    go_conn.Close()
    '    SEI_SRV_VOXEL.ConectarSQLNative(go_conn)
    '    '
    '    '''' aquí configuro connexió addob.
    '    Conn1 = New ADODB.Connection
    '    Obre_Connexio_ADO(Conn1)
    '    '
    '    Conn2 = New ADODB.Connection
    '    Obre_Connexio_ADO(Conn2)
    '    '
    '    'Consutla Capçalera 
    '    ls = ""
    '    ls = ls & " SELECT"
    '    ls = ls & " T0.CardCode,  T0.DocNum,     T0.DocEntry,  T1.U_SEIEdiC,"
    '    ls = ls & " T0.DocDate,   T0.DocDueDate, T0.U_SEI_EDIR, "
    '    ls = ls & " T0.U_SEI_EDIF,T0.U_SEI_EDIE,  T0.U_SEI_EDIL,"
    '    ls = ls & " T0.DocDate,   T0.CardName,   T1.Address,    T1.City      ,T1.ZipCode,"
    '    ls = ls & " T1.LicTradNum,T0.Doccur,     T0.GroupNum,  T2.U_SEI_EDIC,"
    '    ls = ls & " (T0.DocTotal- T0.VatSumSy + T0.DiscSumSy) as BASEIMP,"
    '    ls = ls & " T0.VatSumSy as TOTIMP,"
    '    ls = ls & " T0.DocTotal as TOTAL,"
    '    ls = ls & " T0.Discprcnt as PORCEN1,"   ' Porcentaje Cabecera
    '    ls = ls & " T0.DiscSumSy as IMPDES1,"   ' Importe Porcentaje Cabecera
    '    ls = ls & " T0.Comments,"                ' Observaciones
    '    ls = ls & " T1.MailAddres, T1.MailCity, T1.MailZipCod, isnull(T0.NumAtCard,'') as NumAtCard   "              ' PO Quien emite EDI es el "DESTINATARIO" de la factura
    '    ls = ls & ", T1.CardFName as NomExtranger, t0.shiptocode  "
    '    ls = ls & ", isnull(T1.U_SEIgrcli,'') as U_SEIgrcli, isnull(U_SEINUMPE,'') as U_SEINUMPE "
    '    ls = ls & " FROM OINV T0"
    '    ls = ls & " INNER JOIN OCRD T1"
    '    ls = ls & " ON T0.CardCode= T1.CardCode "
    '    ls = ls & " LEFT OUTER JOIN OCTG T2"
    '    ls = ls & " ON T0.GroupNum= T2.GroupNum "
    '    ls = ls & " WHERE T1.QryGroup41 = 'Y' "         ' Cliente con Flag Facturas VOXEL
    '    ls = ls & " AND ISNULL(T0.U_SEIFiVox,'')=''"    ' Factura no exportada a Voxel   
    '    '
    '    Try
    '        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '        Dim cn As ADODB.Connection = New ADODB.Connection()
    '        Dim rs As ADODB.Recordset = New ADODB.Recordset()
    '        Dim cnStr As String
    '        Dim query As String
    '        cnStr = "Provider=SQLNCLI10;" & _
    '               " Server=" & IniGet(Application.StartupPath & "\S_SEI_CBG_VOXEL.ini", "Parametros", "S") & ";" & _
    '               " Database=" & IniGet(Application.StartupPath & "\S_SEI_CBG_VOXEL.ini", "Parametros", "D") & ";" & _
    '               " UID=" & IniGet(Application.StartupPath & "\S_SEI_CBG_VOXEL.ini", "Parametros", "U") & ";" & _
    '               " PWD=" & IniGet(Application.StartupPath & "\S_SEI_CBG_VOXEL.ini", "Parametros", "P") & ";"
    '        query = ls
    '        oRecordset = New ADODB.Recordset()
    '        oRecordset.Open(query, cnStr, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic, -1)
    '        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' 
    '        Do While Not oRecordset.EOF
    '            '  
    '            HashResumIVAS.Clear()
    '            iFila = 0
    '            '
    '            oXml = ObtenerXML("FacturaVoxel.xml")
    '            oItem = oXml.SelectNodes("//GeneralData")
    '            '
    '            oItem.Item(0).Attributes("Ref").InnerText = oRecordset.Fields.Item("DocNum").Value.ToString '''' oDataReader("DocNum").ToString
    '            oItem.Item(0).Attributes("Type").InnerText = "FacturaComercial"
    '            oItem.Item(0).Attributes("Date").InnerText = Convert.ToDateTime(oRecordset.Fields.Item("Docdate").Value).ToString("yyyy-MM-dd")   '''' Convert.ToDateTime(oRecordset.Fields.Item("Docdate").Value.ToString).ToShortDateString ''' oDataReader("Docdate")
    '            oItem = oXml.SelectNodes("//Supplier")
    '            oItem.Item(0).Attributes("CIF").InnerText = xCIF
    '            oItem.Item(0).Attributes("Company").InnerText = xCompany
    '            oItem.Item(0).Attributes("Address").InnerText = xAddress
    '            oItem.Item(0).Attributes("City").InnerText = xCity
    '            oItem.Item(0).Attributes("PC").InnerText = xZipCode
    '            oItem.Item(0).Attributes("Province").InnerText = xProvince
    '            oItem.Item(0).Attributes("Country").InnerText = xCountry
    '            '
    '            oItem = oXml.SelectNodes("//Client")
    '            oItem.Item(0).Attributes("SupplierClientID").InnerText = oRecordset.Fields.Item("CardCode").Value.ToString   '''oDataReader("CardCode").ToString
    '            If oRecordset.Fields.Item("LicTradNum").Value.ToString.Length > 9 Then
    '                oItem.Item(0).Attributes("CIF").InnerText = oRecordset.Fields.Item("LicTradNum").Value.ToString.Substring(2, 9)   '''' oDataReader("LicTradNum").ToString
    '            Else
    '                oItem.Item(0).Attributes("CIF").InnerText = oRecordset.Fields.Item("LicTradNum").Value.ToString
    '            End If
    '            '
    '            oItem.Item(0).Attributes("Company").InnerText = oRecordset.Fields.Item("CardName").Value.ToString
    '            oItem.Item(0).Attributes("Address").InnerText = oRecordset.Fields.Item("Address").Value.ToString
    '            oItem.Item(0).Attributes("City").InnerText = oRecordset.Fields.Item("City").Value.ToString
    '            oItem.Item(0).Attributes("PC").InnerText = oRecordset.Fields.Item("ZipCode").Value.ToString
    '            '
    '            ''' aquí poso les dades de facturació 


    '            ''' miro si el client es de catalonia hotels per fer coses despres
    '            '''
    '            If oRecordset.Fields.Item("U_SEIgrcli").Value.ToString = "008" Or oRecordset.Fields.Item("U_SEIgrcli").Value.ToString = "052" Then
    '                esHotelsCatalonia = True
    '            Else
    '                esHotelsCatalonia = False
    '            End If
    '            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '            If oRecordset.Fields.Item("U_SEIgrcli").Value.ToString = "002" Then
    '                esAngrup = True
    '            Else
    '                esAngrup = False
    '            End If

    '            ls = ""
    '            ls = ls & " SELECT state,CRD1.Country,OCST.Name from CRD1  "
    '            ls = ls & "   left join OCST   on OCST.code = state "
    '            ls = ls & "where cardcode =  '" & oRecordset.Fields.Item("CardCode").Value.ToString & "' and adrestype = 'B'"  '' oDataReader("CardCode").ToString
    '            SEI_SRV_VOXEL.ConectarSQLNative(go_conn3)
    '            oSqlcomand = New SqlCommand(ls, go_conn3)
    '            oDataReader2 = oSqlcomand.ExecuteReader()
    '            While oDataReader2.Read()
    '                oItem.Item(0).Attributes("Province").InnerText = oDataReader2("Name").ToString
    '                If oDataReader2("Country").ToString = "ES" Then
    '                    oItem.Item(0).Attributes("Country").InnerText = "ESP"
    '                Else
    '                    oItem.Item(0).Attributes("Country").InnerText = oDataReader2("Country").ToString
    '                End If
    '            End While
    '            go_conn3.Close()
    '            '
    '            ''' aquí poso les dades de enviament 
    '            '''' nou nou nou per posar me de una direcció d'enviament.
    '            ''' mirar shiptocode de la factura a quin CRD1 pertan i aquí agafar el  SEIVOXSECUND 
    '            '
    '            ls = ""
    '            ls = ls & " SELECT state,CRD1.Country,OCST.Name, isnull(U_SEIVOXSECUND,'') as secundario from CRD1  "
    '            ls = ls & "   left join OCST   on OCST.code = state "
    '            ls = ls & " where cardcode =  '" & oRecordset.Fields.Item("CardCode").Value.ToString & "' and adrestype = 'S'"  ''' oDataReader("CardCode").ToString
    '            ls = ls & " and crd1.Address =  '" & Replace(oRecordset.Fields.Item("shiptocode").Value.ToString, "'", "''") & "'"  '''' nou nou 
    '            SEI_SRV_VOXEL.ConectarSQLNative(go_conn3)
    '            oSqlcomand = New SqlCommand(ls, go_conn3)
    '            oDataReader2 = oSqlcomand.ExecuteReader()
    '            Dim xSecundari As String = ""
    '            While oDataReader2.Read()
    '                oItem = oXml.SelectNodes("//Customers/Customer")
    '                oItem.Item(0).Attributes("SupplierClientID").InnerText = oRecordset.Fields.Item("CardCode").Value.ToString   '''' oDataReader("CardCode").ToString
    '                xSecundari = oDataReader2("secundario").ToString ''''' oDataReader2.Fields.Item("secundario").Value
    '                oItem.Item(0).Attributes("SupplierCustomerID").InnerText = oDataReader2("secundario").ToString '''''' oDataReader2.Fields.Item("secundario").Value.ToString   ''' nou nou nou

    '                oItem.Item(0).Attributes("Customer").InnerText = oRecordset.Fields.Item("NomExtranger").Value.ToString   '''''oRecordset.Fields.Item("CardName").Value.ToString    ''' oDataReader("CardName").ToString
    '                oItem.Item(0).Attributes("Address").InnerText = oRecordset.Fields.Item("MailAddres").Value.ToString  '''  oDataReader("MailAddres").ToString
    '                oItem.Item(0).Attributes("PC").InnerText = oRecordset.Fields.Item("MailZipCod").Value.ToString   ''' oDataReader("MailZipCod").ToString
    '                oItem.Item(0).Attributes("City").InnerText = oRecordset.Fields.Item("MailCity").Value.ToString   '''  oDataReader("MailCity").ToString
    '                oItem.Item(0).Attributes("Province").InnerText = oDataReader2("Name").ToString

    '                If oDataReader2("Country").ToString = "ES" Then
    '                    oItem.Item(0).Attributes("Country").InnerText = "ESP"
    '                Else
    '                    oItem.Item(0).Attributes("Country").InnerText = oDataReader2("Country").ToString
    '                End If
    '            End While
    '            go_conn3.Close()
    '            '''' això ho poso nou per tornar a posar el supplierid a nivell de client 
    '            oItem = oXml.SelectNodes("//Client")
    '            oItem.Item(0).Attributes("SupplierClientID").InnerText = xSecundari
    '            '''''''''''''''''''''''''''''''''''''''''''''''''
    '            ''' aquí miro les referències '''''''''''''''''''
    '            '''''''''''''''''''''''''''''''''''''''''''''''''
    '            ls = ""
    '            'ls = ls & "  select top 1 baseref  FROM  OINV T0 "
    '            'ls = ls & "  INNER JOIN  INV1 T1  ON T0.DocEntry=T1.DocEntry  "
    '            'ls = ls & "   where(T0.DocEntry = " & oRecordset.Fields.Item("DocEntry").Value.ToString & ")" '''oDataReader("DocEntry").ToString
    '            'ls = ls & "  group by baseref,NumAtCard "

    '            ls = ""
    '            If esAngrup Then
    '                ls = ls & "  select  baseref  FROM  OINV T0 "
    '                ls = ls & "  INNER JOIN  INV1 T1  ON T0.DocEntry=T1.DocEntry  "
    '                ls = ls & "   where(T0.DocEntry = " & oRecordset.Fields.Item("DocEntry").Value.ToString & ")" '''oDataReader("DocEntry").ToString
    '                ls = ls & "  group by baseref,NumAtCard "
    '            Else
    '                ls = ls & "  select top 1 baseref  FROM  OINV T0 "
    '                ls = ls & "  INNER JOIN  INV1 T1  ON T0.DocEntry=T1.DocEntry  "
    '                ls = ls & "   where(T0.DocEntry = " & oRecordset.Fields.Item("DocEntry").Value.ToString & ")" '''oDataReader("DocEntry").ToString
    '                ls = ls & "  group by baseref,NumAtCard "
    '            End If


    '            ''' fi aquí de mirar les referències
    '            SEI_SRV_VOXEL.ConectarSQLNative(go_conn3)
    '            oSqlcomand = New SqlCommand(ls, go_conn3)
    '            oDataReader2 = oSqlcomand.ExecuteReader()
    '            Dim xLinia As Integer = 0
    '            Dim oDocumentLines As Xml.XmlNode
    '            Dim oFirstRow As Xml.XmlNode
    '            Dim oNewRow As Xml.XmlNode

    '            If esAngrup Then

    '                Dim referenciesAlbarans As Hashtable = New Hashtable
    '                ''''' posar les referències al detall en 
    '                If oDataReader2.Read Then
    '                    Dim xxx() As String = oRecordset.Fields.Item("NumAtCard").Value.ToString.Split(";")
    '                    For iter As Integer = 0 To xxx.Length - 1
    '                        If xLinia > 0 Then
    '                            oDocumentLines = oXml.SelectSingleNode("//References")
    '                            oFirstRow = oDocumentLines.FirstChild
    '                            oNewRow = oFirstRow.CloneNode(True)
    '                            oDocumentLines.AppendChild(oNewRow)
    '                        End If
    '                        oItem = oXml.SelectNodes("//References/Reference")
    '                        oItem.Item(xLinia).Attributes("PORef").InnerText = Replace(xxx(xLinia), ".", "")
    '                        xLinia = xLinia + 1
    '                    Next
    '                    If Not IsNothing(oDataReader2("baseref")) And oDataReader2("baseref").ToString <> "" Then
    '                        If Not referenciesAlbarans.ContainsKey(oDataReader2("baseref").ToString) Then
    '                            If xLinia > 0 Then
    '                                oDocumentLines = oXml.SelectSingleNode("//References")
    '                                oFirstRow = oDocumentLines.FirstChild
    '                                oNewRow = oFirstRow.CloneNode(True)
    '                                oDocumentLines.AppendChild(oNewRow)
    '                            End If
    '                            oItem = oXml.SelectNodes("//References/Reference")
    '                            oItem.Item(xLinia).Attributes("DNRef").InnerText = oDataReader2("baseref").ToString
    '                            oItem.Item(xLinia).Attributes("PORef").InnerText = ""
    '                            referenciesAlbarans.Add(oDataReader2("baseref").ToString, oDataReader2("baseref").ToString)
    '                            xLinia = xLinia + 1
    '                        End If
    '                    End If
    '                    While oDataReader2.Read()
    '                        If Not IsNothing(oDataReader2("baseref")) And oDataReader2("baseref").ToString <> "" Then
    '                            If Not referenciesAlbarans.ContainsKey(oDataReader2("baseref").ToString) Then
    '                                If xLinia > 0 Then
    '                                    oDocumentLines = oXml.SelectSingleNode("//References")
    '                                    oFirstRow = oDocumentLines.FirstChild
    '                                    oNewRow = oFirstRow.CloneNode(True)
    '                                    oDocumentLines.AppendChild(oNewRow)
    '                                End If
    '                                oItem = oXml.SelectNodes("//References/Reference")
    '                                oItem.Item(xLinia).Attributes("DNRef").InnerText = oDataReader2("baseref").ToString
    '                                oItem.Item(xLinia).Attributes("PORef").InnerText = ""
    '                                referenciesAlbarans.Add(oDataReader2("baseref").ToString, oDataReader2("baseref").ToString)
    '                                xLinia = xLinia + 1
    '                            End If
    '                        End If
    '                    End While
    '                End If
    '                ''''' fi de posar 



    '            Else

    '                While oDataReader2.Read()
    '                    If xLinia > 0 Then
    '                        oDocumentLines = oXml.SelectSingleNode("//References")
    '                        oFirstRow = oDocumentLines.FirstChild
    '                        oNewRow = oFirstRow.CloneNode(True)
    '                        oDocumentLines.AppendChild(oNewRow)
    '                    End If
    '                    oItem = oXml.SelectNodes("//References/Reference")
    '                    If IsNothing(oDataReader2("baseref")) Or oDataReader2("baseref").ToString = "" Then
    '                        oItem.Item(xLinia).Attributes("DNRef").InnerText = oRecordset.Fields.Item("DocNum").Value.ToString '''' oDataReader("DocNum").ToString
    '                    Else
    '                        oItem.Item(xLinia).Attributes("DNRef").InnerText = oDataReader2("baseref").ToString
    '                    End If
    '                    ''
    '                    If Not esHotelsCatalonia Then
    '                        oItem.Item(xLinia).Attributes("PORef").InnerText = oRecordset.Fields.Item("NumAtCard").Value.ToString
    '                    Else
    '                        oItem.Item(xLinia).Attributes("PORef").InnerText = oRecordset.Fields.Item("U_SEINUMPE").Value.ToString
    '                    End If

    '                    xLinia = xLinia + 1
    '                    ''
    '                End While


    '            End If

    '            go_conn3.Close()
    '            Me.Form.lblmsg.Text = "Factura: " & oRecordset.Fields.Item("CardCode").Value.ToString ''' oDataReader("CardCode").ToString
    '            '''''''''''''''-------------------------------------------------------------------------------------------
    '            '############ LINFAC.TXT Detalle de la Factura (Sumatorio de Lineas necesario para la cabecera) ##########
    '            '''''''''''''''-------------------------------------------------------------------------------------------
    '            oRcsFactLin = ObtenerRcsFactLin(oRecordset.Fields.Item("DocEntry").Value.ToString) ''' oDataReader("DocEntry").ToString
    '            '
    '            While oRcsFactLin.Read()
    '                XML_Linea(oXml, oRcsFactLin, iFila, esHotelsCatalonia)
    '                iFila = iFila + 1
    '            End While
    '            ' 
    '            '!!!!! atenció al obtenir ivas ''''Descripció i poder
    '            Dim xLiniaIVA As Integer = 0
    '            Dim oDocumentLinesIVA As Xml.XmlNode
    '            Dim oFirstRowIVA As Xml.XmlNode
    '            Dim oNewRowIVA As Xml.XmlNode
    '            Dim rsIVES As ADODB.Recordset = New ADODB.Recordset()
    '            rsIVES = ObtenerRcsTiposIvas(oRecordset.Fields.Item("DocEntry").Value.ToString)
    '            Do While Not rsIVES.EOF
    '                If xLiniaIVA > 0 Then
    '                    oDocumentLinesIVA = oXml.SelectSingleNode("//TaxSummary")
    '                    oFirstRowIVA = oDocumentLinesIVA.FirstChild
    '                    oNewRowIVA = oFirstRowIVA.CloneNode(True)
    '                    oDocumentLinesIVA.AppendChild(oNewRowIVA)
    '                End If
    '                oItem = oXml.SelectNodes("//TaxSummary/Tax")
    '                If Not esHotelsCatalonia Then
    '                    oItem.Item(xLiniaIVA).Attributes("Type").InnerText = donaImpost(rsIVES.Fields.Item("VatGroup").Value.ToString.Replace(",", "."))
    '                    oItem.Item(xLiniaIVA).Attributes("Rate").InnerText = rsIVES.Fields.Item("rate").Value.ToString.Replace(",", ".")
    '                    oItem.Item(xLiniaIVA).Attributes("Amount").InnerText = rsIVES.Fields.Item("IMPIMP").Value.ToString.Replace(",", ".")
    '                    oItem.Item(xLiniaIVA).Attributes("Base").InnerText = rsIVES.Fields.Item("BASEIMP").Value.ToString.Replace(",", ".")
    '                Else
    '                    oItem.Item(xLiniaIVA).Attributes("Type").InnerText = donaImpost(rsIVES.Fields.Item("VatGroup").Value.ToString.Replace(",", "."))
    '                    oItem.Item(xLiniaIVA).Attributes("Rate").InnerText = String.Format("{0:0.000}", rsIVES.Fields.Item("rate").Value).ToString.Replace(",", ".")
    '                    oItem.Item(xLiniaIVA).Attributes("Amount").InnerText = String.Format("{0:0.000}", rsIVES.Fields.Item("IMPIMP").Value).ToString.Replace(",", ".")
    '                    oItem.Item(xLiniaIVA).Attributes("Base").InnerText = String.Format("{0:0.000}", rsIVES.Fields.Item("BASEIMP").Value).ToString.Replace(",", ".")
    '                End If

    '                xLiniaIVA = xLiniaIVA + 1
    '                ' aquí s'ha tret el principi
    '                rsIVES.MoveNext()
    '            Loop
    '            ''' fi atenció obtenir ives !!!! 
    '            '
    '            oItem = oXml.SelectNodes("//TotalSummary")
    '            If Not esHotelsCatalonia Then
    '                oItem.Item(0).Attributes("SubTotal").InnerText = oRecordset.Fields.Item("BASEIMP").Value.ToString.Replace(",", ".") '''oDataReader("BASEIMP").ToString
    '                oItem.Item(0).Attributes("Tax").InnerText = oRecordset.Fields.Item("TOTIMP").Value.ToString.Replace(",", ".")  ''' oDataReader("TOTIMP").ToString
    '                oItem.Item(0).Attributes("Total").InnerText = oRecordset.Fields.Item("TOTAL").Value.ToString.Replace(",", ".")  '''  oDataReader("TOTAL").ToString
    '            Else
    '                oItem.Item(0).Attributes("SubTotal").InnerText = String.Format("{0:0.000}", oRecordset.Fields.Item("BASEIMP").Value).ToString.Replace(",", ".")
    '                oItem.Item(0).Attributes("Tax").InnerText = String.Format("{0:0.000}", oRecordset.Fields.Item("TOTIMP").Value).ToString.Replace(",", ".")
    '                oItem.Item(0).Attributes("Total").InnerText = String.Format("{0:0.000}", oRecordset.Fields.Item("TOTAL").Value).ToString.Replace(",", ".")
    '            End If

    '            '
    '            sPath = Application.StartupPath() & "\"
    '            sPath = IniGet(Application.StartupPath & "\S_SEI_CBG_VOXEL.ini", "Parametros", "C") '''' "c:\PROVES VOXEL\"
    '            sPathD = IniGet(Application.StartupPath & "\S_SEI_CBG_VOXEL.ini", "Parametros", "R")
    '            sFichero = sPath & "Factura_" & oRecordset.Fields.Item("DocNum").Value.ToString & "_" & "000" & ".xml"   '''oDataReader("DocNum").ToString
    '            sFicheroD = sPathD & "\" & "Factura_" & oRecordset.Fields.Item("DocNum").Value.ToString & "_" & "000" & ".xml"   '''oDataReader("DocNum").ToString
    '            ''''''' 
    '            oXml.Save(sFichero)
    '            oXml.Save(sFicheroD)
    '            '
    '            ls = ""
    '            ls = ls & " update  OINV  set  U_SEIFivox = '" & sFichero & "'where docentry  = " & oRecordset.Fields.Item("DocEntry").Value.ToString
    '            '' ''oDataReader = oSqlcomand.ExecuteReader()
    '            orecordset2 = Nothing
    '            orecordset2 = conn2.Execute(ls)
    '            '
    '            HashFEnviats(oRecordset.Fields.Item("DocEntry").Value.ToString) = sFichero '''oDataReader("DocEntry").ToString
    '            '
    '            oDataReader2 = Nothing
    '            oSqlcomand = Nothing
    '            oRecordset.MoveNext()
    '            ' 
    '        Loop
    '        '
    '        Dim oEnumerador As IDictionaryEnumerator
    '        oEnumerador = HashFEnviats.GetEnumerator
    '        While oEnumerador.MoveNext
    '            ls = " update  OINV  set  U_SEIFivox = '" & oEnumerador.Value & "'where docentry  = " & oEnumerador.Key
    '            SEI_SRV_VOXEL.ConectarSQLNative(go_conn3)
    '            oSqlcomand = Nothing
    '            oSqlcomand = New SqlCommand(ls, go_conn3)
    '            oSqlcomand.ExecuteNonQuery()
    '            go_conn3.Close()
    '        End While
    '        HashFEnviats.Clear()
    '        '
    '    Catch ex As Exception
    '        Dim oEnumerador As IDictionaryEnumerator
    '        oEnumerador = HashFEnviats.GetEnumerator
    '        While oEnumerador.MoveNext
    '            ls = " update  OINV  set  U_SEIFivox = '" & oEnumerador.Value & "'where docentry  = " & oEnumerador.Key
    '            SEI_SRV_VOXEL.ConectarSQLNative(go_conn3)
    '            oSqlcomand = Nothing
    '            oSqlcomand = New SqlCommand(ls, go_conn3)
    '            oSqlcomand.ExecuteNonQuery()
    '            go_conn3.Close()
    '        End While
    '        HashFEnviats.Clear()
    '        Me.Form.lblmsg.Text = ex.Message
    '    Finally
    '        If Not IsNothing(oDataReader) Then
    '            oDataReader.Close()
    '        End If
    '    End Try
    'End Sub

    Public Sub GENERAR_FACTURES_TLY()
        Dim ls As String
        Dim oSqlcomand As SqlCommand
        Dim oDataReader As SqlClient.SqlDataReader = Nothing
        Dim oDataReader2 As SqlClient.SqlDataReader = Nothing
        Dim oRcsFactLin As SqlClient.SqlDataReader = Nothing
        Dim oXml As XmlDocument = Nothing
        Dim oItem As System.Xml.XmlNodeList = Nothing
        Dim sPath, sPathD As String
        Dim sFichero, sFicheroD As String
        Dim numSerie As String = "0"
        Dim iFila As Integer
        Dim go_conn3 As SqlConnection = Nothing
        Dim HashFEnviats As Hashtable = New Hashtable
        Dim HashResumIVAS As Hashtable = New Hashtable
        Dim Conn1, conn2 As New ADODB.Connection
        Dim Cmd1 As New ADODB.Command
        Dim oRecordset, orecordset2 As ADODB.Recordset '''SAPbobsCOM.Recordset
        '
        Dim esHotelsCatalonia As Boolean = False
        Dim esAngrup As Boolean = False

        '' dades de CBG
        Dim xCIF As String = ""
        Dim xAdreca As String = ""
        Dim xCompany As String = ""
        Dim xAddress As String = ""
        Dim xCity As String = ""
        Dim xProvince As String = ""
        Dim xCountry As String = ""
        Dim xZipCode As String = ""
        '' 
        ls = ""
        ls = ls & " select top 1 OADM.TaxIdNum,OADM.CompnyName,OADM.CompnyAddr, adm1.City, adm1.State, adm1.Country, adm1.ZipCode, OCST.Name as nomprov from OADM "
        ls = ls & " left join ADM1 on oadm.currPeriod = adm1.currPeriod  "
        ls = ls & "   left join OCST   on OCST.code = adm1.State "
        oSqlcomand = New SqlCommand(ls, go_conn)
        oDataReader = oSqlcomand.ExecuteReader()
        While oDataReader.Read()
            xCIF = oDataReader("TaxIdNum").ToString.Substring(2, 9)
            xCompany = "COMERCIAL CBG S.A."   '''''oDataReader("CompnyName").ToString
            xAddress = oDataReader("CompnyAddr").ToString
            xCity = oDataReader("City").ToString
            xProvince = oDataReader("nomprov").ToString
            If oDataReader("Country").ToString = "ES" Then
                xCountry = "ESP"
            Else
                xCountry = oDataReader("Country").ToString
            End If
            xZipCode = oDataReader("ZipCode").ToString
        End While
        ''''''' fi agafar dades cbg 
        go_conn.Close()
        SEI_SRV_VOXEL.ConectarSQLNative(go_conn)
        '
        ''''' aquí configuro connexió addob.
        Conn1 = New ADODB.Connection
        Obre_Connexio_ADO(Conn1)
        '
        Conn2 = New ADODB.Connection
        Obre_Connexio_ADO(Conn2)
        '
        'Consutla Capçalera 
        ls = ""
        ls = ls & " SELECT"
        ls = ls & " T0.CardCode,  T0.DocNum,     T0.DocEntry,  T1.U_SEIEdiC,"
        ls = ls & " T0.DocDate,   T0.DocDueDate, T0.U_SEI_EDIR, "
        ls = ls & " T0.U_SEI_EDIF,T0.U_SEI_EDIE,  T0.U_SEI_EDIL,"
        ls = ls & " T0.DocDate,   T0.CardName,   T1.Address,    T1.City      ,T1.ZipCode,"
        ls = ls & " T1.LicTradNum,T0.Doccur,     T0.GroupNum,  T2.U_SEI_EDIC,"
        ls = ls & " (T0.DocTotal- T0.VatSumSy + T0.DiscSumSy) as BASEIMP,"
        ls = ls & " T0.VatSumSy as TOTIMP,"
        ls = ls & " T0.DocTotal as TOTAL,"
        ls = ls & " T0.Discprcnt as PORCEN1,"   ' Porcentaje Cabecera
        ls = ls & " T0.DiscSumSy as IMPDES1,"   ' Importe Porcentaje Cabecera
        ls = ls & " T0.Comments,"                ' Observaciones
        ls = ls & " T1.MailAddres, T1.MailCity, T1.MailZipCod, isnull(T0.NumAtCard,'') as NumAtCard   "              ' PO Quien emite EDI es el "DESTINATARIO" de la factura
        ls = ls & ", T1.CardFName as NomExtranger, t0.shiptocode  "
        ls = ls & ", isnull(T1.U_SEIgrcli,'') as U_SEIgrcli, isnull(U_SEINUMPE,'') as U_SEINUMPE "
        ls = ls & ", isnull(t0.PeyMethod, '') as PeyMethod  "
        ls = ls & " FROM OINV T0"
        ls = ls & " INNER JOIN OCRD T1"
        ls = ls & " ON T0.CardCode= T1.CardCode "
        ls = ls & " LEFT OUTER JOIN OCTG T2"
        ls = ls & " ON T0.GroupNum= T2.GroupNum "
        ls = ls & " WHERE T1.QryGroup41 = 'Y' "         ' Cliente con Flag Facturas VOXEL
        ls = ls & " AND ISNULL(T0.U_SEIFiVox,'')=''"    ' Factura no exportada a Voxel   

        ''''''''''''''''TREUREEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEe
        'ls = ls & " and  (T0.docentry = '1033574' or T0.docentry = '1033574') "
        '
        Try
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Dim cn As ADODB.Connection = New ADODB.Connection()
            Dim rs As ADODB.Recordset = New ADODB.Recordset()
            Dim cnStr As String
            Dim query As String
            cnStr = "Provider=SQLNCLI11;" &
                   " Server=" & IniGet(Application.StartupPath & "\S_SEI_CBG_VOXEL.ini", "Parametros", "S") & ";" &
                   " Database=" & IniGet(Application.StartupPath & "\S_SEI_CBG_VOXEL.ini", "Parametros", "D") & ";" &
                   " UID=" & IniGet(Application.StartupPath & "\S_SEI_CBG_VOXEL.ini", "Parametros", "U") & ";" &
                   " PWD=" & IniGet(Application.StartupPath & "\S_SEI_CBG_VOXEL.ini", "Parametros", "P") & ";"
            query = ls
            oRecordset = New ADODB.Recordset()
            oRecordset.Open(query, cnStr, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic, -1)
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' 
            Do While Not oRecordset.EOF
                '  
                HashResumIVAS.Clear()
                iFila = 0
                '
                oXml = ObtenerXML("FacturaVoxel.xml")
                oItem = oXml.SelectNodes("//GeneralData")
                '
                oItem.Item(0).Attributes("Ref").InnerText = oRecordset.Fields.Item("DocNum").Value.ToString '''' oDataReader("DocNum").ToString
                oItem.Item(0).Attributes("Type").InnerText = "FacturaComercial"
                oItem.Item(0).Attributes("Date").InnerText = Convert.ToDateTime(oRecordset.Fields.Item("Docdate").Value).ToString("yyyy-MM-dd")   '''' Convert.ToDateTime(oRecordset.Fields.Item("Docdate").Value.ToString).ToShortDateString ''' oDataReader("Docdate")

                If oRecordset.Fields.Item("DocCur").Value.ToString <> "EUR" And oRecordset.Fields.Item("DocCur").Value.ToString <> "" Then
                    oItem.Item(0).Attributes("Currency").InnerText = oRecordset.Fields.Item("DocCur").Value.ToString
                End If

                oItem = oXml.SelectNodes("//Supplier")
                oItem.Item(0).Attributes("CIF").InnerText = xCIF
                oItem.Item(0).Attributes("Company").InnerText = xCompany
                oItem.Item(0).Attributes("Address").InnerText = xAddress
                oItem.Item(0).Attributes("City").InnerText = xCity
                oItem.Item(0).Attributes("PC").InnerText = xZipCode
                oItem.Item(0).Attributes("Province").InnerText = xProvince
                oItem.Item(0).Attributes("Country").InnerText = xCountry
                '
                oItem = oXml.SelectNodes("//Client")
                oItem.Item(0).Attributes("SupplierClientID").InnerText = oRecordset.Fields.Item("CardCode").Value.ToString   '''oDataReader("CardCode").ToString
                If oRecordset.Fields.Item("LicTradNum").Value.ToString.Length > 9 Then
                    oItem.Item(0).Attributes("CIF").InnerText = oRecordset.Fields.Item("LicTradNum").Value.ToString.Substring(2, 9)   '''' oDataReader("LicTradNum").ToString
                Else
                    oItem.Item(0).Attributes("CIF").InnerText = oRecordset.Fields.Item("LicTradNum").Value.ToString
                End If
                '
                oItem.Item(0).Attributes("Company").InnerText = oRecordset.Fields.Item("CardName").Value.ToString
                oItem.Item(0).Attributes("Address").InnerText = oRecordset.Fields.Item("Address").Value.ToString
                oItem.Item(0).Attributes("City").InnerText = oRecordset.Fields.Item("City").Value.ToString
                oItem.Item(0).Attributes("PC").InnerText = oRecordset.Fields.Item("ZipCode").Value.ToString
                '
                ''' aquí poso les dades de facturació 


                ''' miro si el client es de catalonia hotels per fer coses despres
                '''
                If oRecordset.Fields.Item("U_SEIgrcli").Value.ToString = "008" Or oRecordset.Fields.Item("U_SEIgrcli").Value.ToString = "052" Then
                    esHotelsCatalonia = True
                Else
                    esHotelsCatalonia = False
                End If
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                If oRecordset.Fields.Item("U_SEIgrcli").Value.ToString = "002" Then
                    esAngrup = True
                Else
                    esAngrup = False
                End If

                ls = ""
                ls = ls & " SELECT state,CRD1.Country,OCST.Name from CRD1  "
                ls = ls & "   left join OCST   on OCST.code = state "
                ls = ls & "where cardcode =  '" & oRecordset.Fields.Item("CardCode").Value.ToString & "' and adrestype = 'B'"  '' oDataReader("CardCode").ToString
                SEI_SRV_VOXEL.ConectarSQLNative(go_conn3)
                oSqlcomand = New SqlCommand(ls, go_conn3)
                oDataReader2 = oSqlcomand.ExecuteReader()
                While oDataReader2.Read()
                    oItem.Item(0).Attributes("Province").InnerText = oDataReader2("Name").ToString
                    If oDataReader2("Country").ToString = "ES" Then
                        oItem.Item(0).Attributes("Country").InnerText = "ESP"
                    Else
                        oItem.Item(0).Attributes("Country").InnerText = oDataReader2("Country").ToString
                    End If
                End While
                go_conn3.Close()
                '
                ''' aquí poso les dades de enviament 
                '''' nou nou nou per posar me de una direcció d'enviament.
                ''' mirar shiptocode de la factura a quin CRD1 pertan i aquí agafar el  SEIVOXSECUND 
                '
                ls = ""
                ls = ls & " SELECT state,CRD1.Country,OCST.Name, isnull(U_SEIVOXSECUND,'') as secundario from CRD1  "
                ls = ls & "   left join OCST   on OCST.code = state "
                ls = ls & " where cardcode =  '" & oRecordset.Fields.Item("CardCode").Value.ToString & "' and adrestype = 'S'"  ''' oDataReader("CardCode").ToString
                ls = ls & " and crd1.Address =  '" & Replace(oRecordset.Fields.Item("shiptocode").Value.ToString, "'", "''") & "'"  '''' nou nou 
                SEI_SRV_VOXEL.ConectarSQLNative(go_conn3)
                oSqlcomand = New SqlCommand(ls, go_conn3)
                oDataReader2 = oSqlcomand.ExecuteReader()
                Dim xSecundari As String = ""
                While oDataReader2.Read()
                    oItem = oXml.SelectNodes("//Customers/Customer")
                    oItem.Item(0).Attributes("SupplierClientID").InnerText = oRecordset.Fields.Item("CardCode").Value.ToString   '''' oDataReader("CardCode").ToString
                    xSecundari = oDataReader2("secundario").ToString ''''' oDataReader2.Fields.Item("secundario").Value
                    oItem.Item(0).Attributes("SupplierCustomerID").InnerText = oDataReader2("secundario").ToString '''''' oDataReader2.Fields.Item("secundario").Value.ToString   ''' nou nou nou

                    oItem.Item(0).Attributes("Customer").InnerText = oRecordset.Fields.Item("NomExtranger").Value.ToString   '''''oRecordset.Fields.Item("CardName").Value.ToString    ''' oDataReader("CardName").ToString
                    oItem.Item(0).Attributes("Address").InnerText = oRecordset.Fields.Item("MailAddres").Value.ToString  '''  oDataReader("MailAddres").ToString
                    oItem.Item(0).Attributes("PC").InnerText = oRecordset.Fields.Item("MailZipCod").Value.ToString   ''' oDataReader("MailZipCod").ToString
                    oItem.Item(0).Attributes("City").InnerText = oRecordset.Fields.Item("MailCity").Value.ToString   '''  oDataReader("MailCity").ToString
                    oItem.Item(0).Attributes("Province").InnerText = oDataReader2("Name").ToString

                    If oDataReader2("Country").ToString = "ES" Then
                        oItem.Item(0).Attributes("Country").InnerText = "ESP"
                    Else
                        oItem.Item(0).Attributes("Country").InnerText = oDataReader2("Country").ToString
                    End If
                End While
                go_conn3.Close()
                '''' això ho poso nou per tornar a posar el supplierid a nivell de client 
                oItem = oXml.SelectNodes("//Client")
                oItem.Item(0).Attributes("SupplierClientID").InnerText = xSecundari
                '''''''''''''''''''''''''''''''''''''''''''''''''
                ''' aquí miro les referències '''''''''''''''''''
                '''''''''''''''''''''''''''''''''''''''''''''''''
                ls = ""
                'ls = ls & "  select top 1 baseref  FROM  OINV T0 "
                'ls = ls & "  INNER JOIN  INV1 T1  ON T0.DocEntry=T1.DocEntry  "
                'ls = ls & "   where(T0.DocEntry = " & oRecordset.Fields.Item("DocEntry").Value.ToString & ")" '''oDataReader("DocEntry").ToString
                'ls = ls & "  group by baseref,NumAtCard "

                ls = ""
                If esAngrup Then
                    ls = ls & "  select  baseref, t1.baseentry , t2.U_SEINUMPE, t2.DocDate   FROM  OINV T0 "
                    ls = ls & "  INNER JOIN  INV1 T1  ON T0.DocEntry=T1.DocEntry  "
                    ls = ls & "  left join ODLN t2 on T1.BaseEntry = t2.docentry "
                    ls = ls & "   where(T0.DocEntry = " & oRecordset.Fields.Item("DocEntry").Value.ToString & ")" '''oDataReader("DocEntry").ToString
                    ls = ls & "  group by baseref, t1.baseentry, t2.U_SEINUMPE, t2.DocDate  "
                Else
                    '' ls = ls & "  select top 1 baseref  FROM  OINV T0 "
                    ''' tret el top 1 perquè retorni tots els albrans.
                    ls = ls & "  select   baseref , t2.DocDate     FROM  OINV T0 "
                    ls = ls & "  INNER JOIN  INV1 T1  ON T0.DocEntry=T1.DocEntry  "
                    ls = ls & "  left join ODLN t2 on T1.BaseEntry = t2.docentry "
                    ls = ls & "   where(T0.DocEntry = " & oRecordset.Fields.Item("DocEntry").Value.ToString & ")" '''oDataReader("DocEntry").ToString
                    ls = ls & "  group by baseref,t0.NumAtCard, t2.DocDate   "
                End If


                ''' fi aquí de mirar les referències
                SEI_SRV_VOXEL.ConectarSQLNative(go_conn3)
                oSqlcomand = New SqlCommand(ls, go_conn3)
                oDataReader2 = oSqlcomand.ExecuteReader()
                Dim xLinia As Integer = 0
                Dim oDocumentLines As Xml.XmlNode
                Dim oFirstRow As Xml.XmlNode
                Dim oNewRow As Xml.XmlNode

                If esAngrup Then

                    Dim referenciesAlbarans As Hashtable = New Hashtable
                    Dim referenciesComandes As Hashtable = New Hashtable
                    ''''' posar les referències al detall en 
                    '' If oDataReader2.Read Then
                    'Dim xxx() As String = oRecordset.Fields.Item("NumAtCard").Value.ToString.Split(";")
                    'For iter As Integer = 0 To xxx.Length - 1
                    '    If xLinia > 0 Then
                    '        oDocumentLines = oXml.SelectSingleNode("//References")
                    '        oFirstRow = oDocumentLines.FirstChild
                    '        oNewRow = oFirstRow.CloneNode(True)
                    '        oDocumentLines.AppendChild(oNewRow)
                    '    End If
                    '    oItem = oXml.SelectNodes("//References/Reference")
                    '    oItem.Item(xLinia).Attributes("PORef").InnerText = Replace(xxx(xLinia), ".", "")
                    '    xLinia = xLinia + 1
                    'Next
                    'If Not IsNothing(oDataReader2("baseref")) And oDataReader2("baseref").ToString <> "" Then
                    '    If Not referenciesAlbarans.ContainsKey(oDataReader2("baseref").ToString) Then
                    '        If xLinia > 0 Then
                    '            oDocumentLines = oXml.SelectSingleNode("//References")
                    '            oFirstRow = oDocumentLines.FirstChild
                    '            oNewRow = oFirstRow.CloneNode(True)
                    '            oDocumentLines.AppendChild(oNewRow)
                    '        End If
                    '        oItem = oXml.SelectNodes("//References/Reference")
                    '        oItem.Item(xLinia).Attributes("DNRef").InnerText = oDataReader2("baseref").ToString
                    '        oItem.Item(xLinia).Attributes("PORef").InnerText = ""
                    '        referenciesAlbarans.Add(oDataReader2("baseref").ToString, oDataReader2("baseref").ToString)
                    '        xLinia = xLinia + 1
                    '    End If
                    'End If

                    While oDataReader2.Read()


                        If Not IsNothing(oDataReader2("U_SEINUMPE")) And oDataReader2("U_SEINUMPE").ToString <> "" Then
                            If Not referenciesComandes.ContainsKey(oDataReader2("U_SEINUMPE").ToString) Then
                                If xLinia > 0 Then
                                    oDocumentLines = oXml.SelectSingleNode("//References")
                                    oFirstRow = oDocumentLines.FirstChild
                                    oNewRow = oFirstRow.CloneNode(True)
                                    oDocumentLines.AppendChild(oNewRow)
                                End If
                                oItem = oXml.SelectNodes("//References/Reference")
                                oItem.Item(xLinia).Attributes("DNRef").InnerText = ""
                                oItem.Item(xLinia).Attributes("PORef").InnerText = oDataReader2("U_SEINUMPE").ToString  ''''""
                                referenciesComandes.Add(oDataReader2("U_SEINUMPE").ToString, oDataReader2("U_SEINUMPE").ToString)
                                xLinia = xLinia + 1
                            End If
                        End If
                        'If Not IsNothing(oDataReader2("baseref")) And oDataReader2("baseref").ToString <> "" Then
                        '    If Not referenciesAlbarans.ContainsKey(oDataReader2("baseref").ToString) Then
                        '        If xLinia > 0 Then
                        '            oDocumentLines = oXml.SelectSingleNode("//References")
                        '            oFirstRow = oDocumentLines.FirstChild
                        '            oNewRow = oFirstRow.CloneNode(True)
                        '            oDocumentLines.AppendChild(oNewRow)
                        '        End If
                        '        oItem = oXml.SelectNodes("//References/Reference")
                        '        oItem.Item(xLinia).Attributes("DNRef").InnerText = oDataReader2("baseref").ToString
                        '        oItem.Item(xLinia).Attributes("PORef").InnerText = ""
                        '        referenciesAlbarans.Add(oDataReader2("baseref").ToString, oDataReader2("baseref").ToString)
                        '        xLinia = xLinia + 1
                        '    End If
                        'End If

                    End While
                    oDataReader2.Close()
                    oDataReader2 = Nothing
                    oDataReader2 = oSqlcomand.ExecuteReader()

                    While oDataReader2.Read()


                        'If Not IsNothing(oDataReader2("U_SEINUMPE")) And oDataReader2("U_SEINUMPE").ToString <> "" Then
                        '    If Not referenciesComandes.ContainsKey(oDataReader2("U_SEINUMPE").ToString) Then
                        '        If xLinia > 0 Then
                        '            oDocumentLines = oXml.SelectSingleNode("//References")
                        '            oFirstRow = oDocumentLines.FirstChild
                        '            oNewRow = oFirstRow.CloneNode(True)
                        '            oDocumentLines.AppendChild(oNewRow)
                        '        End If
                        '        oItem = oXml.SelectNodes("//References/Reference")
                        '        oItem.Item(xLinia).Attributes("DNRef").InnerText = ""
                        '        oItem.Item(xLinia).Attributes("PORef").InnerText = oDataReader2("U_SEINUMPE").ToString  ''''""
                        '        referenciesComandes.Add(oDataReader2("U_SEINUMPE").ToString, oDataReader2("U_SEINUMPE").ToString)
                        '        xLinia = xLinia + 1
                        '    End If
                        'End If
                        If Not IsNothing(oDataReader2("baseref")) And oDataReader2("baseref").ToString <> "" Then
                            If Not referenciesAlbarans.ContainsKey(oDataReader2("baseref").ToString) Then
                                If xLinia > 0 Then
                                    oDocumentLines = oXml.SelectSingleNode("//References")
                                    oFirstRow = oDocumentLines.FirstChild
                                    oNewRow = oFirstRow.CloneNode(True)
                                    oDocumentLines.AppendChild(oNewRow)
                                End If
                                oItem = oXml.SelectNodes("//References/Reference")
                                oItem.Item(xLinia).Attributes("DNRef").InnerText = oDataReader2("baseref").ToString
                                oItem.Item(xLinia).Attributes("DNRefDate").InnerText = Convert.ToDateTime(oDataReader2("DocDate").ToString).ToString("yyyy-MM-dd")   ''' jgt nou 19/10/2022
                                oItem.Item(xLinia).Attributes("PORef").InnerText = ""
                                referenciesAlbarans.Add(oDataReader2("baseref").ToString, oDataReader2("baseref").ToString)
                                xLinia = xLinia + 1
                            End If
                        End If

                    End While




                    '' End If
                    ''''' fi de posar 



                Else

                    While oDataReader2.Read()
                        If xLinia > 0 Then
                            oDocumentLines = oXml.SelectSingleNode("//References")
                            oFirstRow = oDocumentLines.FirstChild
                            oNewRow = oFirstRow.CloneNode(True)
                            oDocumentLines.AppendChild(oNewRow)
                        End If
                        oItem = oXml.SelectNodes("//References/Reference")
                        If IsNothing(oDataReader2("baseref")) Or oDataReader2("baseref").ToString = "" Then
                            oItem.Item(xLinia).Attributes("DNRef").InnerText = oRecordset.Fields.Item("DocNum").Value.ToString '''' oDataReader("DocNum").ToString
                            oItem.Item(xLinia).Attributes("DNRefDate").InnerText = Convert.ToDateTime(oRecordset.Fields.Item("Docdate").Value).ToString("yyyy-MM-dd")   ''' jgt nou 19/10/2022
                        Else
                            oItem.Item(xLinia).Attributes("DNRef").InnerText = oDataReader2("baseref").ToString
                            oItem.Item(xLinia).Attributes("DNRefDate").InnerText = Convert.ToDateTime(oDataReader2("DocDate").ToString).ToString("yyyy-MM-dd")   ''' jgt nou 19/10/2022
                        End If
                        ''
                        If Not esHotelsCatalonia Then
                            oItem.Item(xLinia).Attributes("PORef").InnerText = oRecordset.Fields.Item("NumAtCard").Value.ToString
                        Else
                            oItem.Item(xLinia).Attributes("PORef").InnerText = oRecordset.Fields.Item("U_SEINUMPE").Value.ToString
                        End If

                        xLinia = xLinia + 1
                        ''
                    End While
                End If

                go_conn3.Close()
                Me.Form.lblmsg.Text = "Factura: " & oRecordset.Fields.Item("CardCode").Value.ToString ''' oDataReader("CardCode").ToString
                '''''''''''''''-------------------------------------------------------------------------------------------
                '############ LINFAC.TXT Detalle de la Factura (Sumatorio de Lineas necesario para la cabecera) ##########
                '''''''''''''''-------------------------------------------------------------------------------------------
                oRcsFactLin = ObtenerRcsFactLin(oRecordset.Fields.Item("DocEntry").Value.ToString) ''' oDataReader("DocEntry").ToString
                '
                While oRcsFactLin.Read()
                    XML_Linea(oXml, oRcsFactLin, iFila, esHotelsCatalonia)
                    iFila = iFila + 1
                End While
                ' 
                '!!!!! atenció al obtenir ivas ''''Descripció i poder
                Dim xLiniaIVA As Integer = 0
                Dim oDocumentLinesIVA As Xml.XmlNode
                Dim oFirstRowIVA As Xml.XmlNode
                Dim oNewRowIVA As Xml.XmlNode
                Dim rsIVES As ADODB.Recordset = New ADODB.Recordset()
                rsIVES = ObtenerRcsTiposIvas(oRecordset.Fields.Item("DocEntry").Value.ToString)
                Do While Not rsIVES.EOF
                    If xLiniaIVA > 0 Then
                        oDocumentLinesIVA = oXml.SelectSingleNode("//TaxSummary")
                        oFirstRowIVA = oDocumentLinesIVA.FirstChild
                        oNewRowIVA = oFirstRowIVA.CloneNode(True)
                        oDocumentLinesIVA.AppendChild(oNewRowIVA)
                    End If
                    oItem = oXml.SelectNodes("//TaxSummary/Tax")
                    If Not esHotelsCatalonia Then
                        oItem.Item(xLiniaIVA).Attributes("Type").InnerText = donaImpost(rsIVES.Fields.Item("VatGroup").Value.ToString.Replace(",", "."))
                        oItem.Item(xLiniaIVA).Attributes("Rate").InnerText = rsIVES.Fields.Item("rate").Value.ToString.Replace(",", ".")
                        oItem.Item(xLiniaIVA).Attributes("Amount").InnerText = rsIVES.Fields.Item("IMPIMP").Value.ToString.Replace(",", ".")
                        oItem.Item(xLiniaIVA).Attributes("Base").InnerText = rsIVES.Fields.Item("BASEIMP").Value.ToString.Replace(",", ".")
                    Else
                        oItem.Item(xLiniaIVA).Attributes("Type").InnerText = donaImpost(rsIVES.Fields.Item("VatGroup").Value.ToString.Replace(",", "."))
                        oItem.Item(xLiniaIVA).Attributes("Rate").InnerText = String.Format("{0:0.000}", rsIVES.Fields.Item("rate").Value).ToString.Replace(",", ".")
                        oItem.Item(xLiniaIVA).Attributes("Amount").InnerText = String.Format("{0:0.000}", rsIVES.Fields.Item("IMPIMP").Value).ToString.Replace(",", ".")
                        oItem.Item(xLiniaIVA).Attributes("Base").InnerText = String.Format("{0:0.000}", rsIVES.Fields.Item("BASEIMP").Value).ToString.Replace(",", ".")
                    End If

                    xLiniaIVA = xLiniaIVA + 1
                    ' aquí s'ha tret el principi
                    rsIVES.MoveNext()
                Loop
                ''' fi atenció obtenir ives !!!! 
                '

                ''''' poso el tag duedate
                oItem = oXml.SelectNodes("//DueDates/DueDate")
                oItem.Item(0).Attributes("Date").InnerText = Convert.ToDateTime(oRecordset.Fields.Item("DocDueDate").Value).ToString("yyyy-MM-dd")   ''' '''oDataReader("BASEIMP").ToString
                oItem.Item(0).Attributes("Amount").InnerText = String.Format("{0:0.000}", oRecordset.Fields.Item("TOTAL").Value).ToString.Replace(",", ".")  ''' oDataReader("TOTIMP").ToString
                Dim descaux As String = ""
                oItem.Item(0).Attributes("PaymentID").InnerText = donaMetodoPago(Trim(oRecordset.Fields.Item("CardCode").Value), descaux)  '''  oDataReader("TOTAL").ToString
                oItem.Item(0).Attributes("Description").InnerText = descaux '''  oDataReader("TOTAL").ToString
                ''''' fi de posar el tag duedate

                oItem = oXml.SelectNodes("//TotalSummary")
                If Not esHotelsCatalonia Then
                    oItem.Item(0).Attributes("SubTotal").InnerText = oRecordset.Fields.Item("BASEIMP").Value.ToString.Replace(",", ".") '''oDataReader("BASEIMP").ToString
                    oItem.Item(0).Attributes("Tax").InnerText = oRecordset.Fields.Item("TOTIMP").Value.ToString.Replace(",", ".")  ''' oDataReader("TOTIMP").ToString
                    oItem.Item(0).Attributes("Total").InnerText = oRecordset.Fields.Item("TOTAL").Value.ToString.Replace(",", ".")  '''  oDataReader("TOTAL").ToString
                Else
                    oItem.Item(0).Attributes("SubTotal").InnerText = String.Format("{0:0.000}", oRecordset.Fields.Item("BASEIMP").Value).ToString.Replace(",", ".")
                    oItem.Item(0).Attributes("Tax").InnerText = String.Format("{0:0.000}", oRecordset.Fields.Item("TOTIMP").Value).ToString.Replace(",", ".")
                    oItem.Item(0).Attributes("Total").InnerText = String.Format("{0:0.000}", oRecordset.Fields.Item("TOTAL").Value).ToString.Replace(",", ".")
                End If
                '
                sPath = Application.StartupPath() & "\"
                sPath = IniGet(Application.StartupPath & "\S_SEI_CBG_VOXEL.ini", "Parametros", "C") '''' "c:\PROVES VOXEL\"
                sPathD = IniGet(Application.StartupPath & "\S_SEI_CBG_VOXEL.ini", "Parametros", "R")
                sFichero = sPath & "Factura_" & oRecordset.Fields.Item("DocNum").Value.ToString & "_" & "000" & ".xml"   '''oDataReader("DocNum").ToString
                sFicheroD = sPathD & "\" & "Factura_" & oRecordset.Fields.Item("DocNum").Value.ToString & "_" & "000" & ".xml"   '''oDataReader("DocNum").ToString
                '
                oXml.Save(sFichero)
                oXml.Save(sFicheroD)
                '
                ls = ""
                ls = ls & " update  OINV  set  U_SEIFivox = '" & sFichero & "'where docentry  = " & oRecordset.Fields.Item("DocEntry").Value.ToString
                '' ''oDataReader = oSqlcomand.ExecuteReader()
                orecordset2 = Nothing
                orecordset2 = conn2.Execute(ls)
                '
                HashFEnviats(oRecordset.Fields.Item("DocEntry").Value.ToString) = sFichero '''oDataReader("DocEntry").ToString
                '
                oDataReader2 = Nothing
                oSqlcomand = Nothing
                oRecordset.MoveNext()
                ' 
            Loop
            '
            Dim oEnumerador As IDictionaryEnumerator
            oEnumerador = HashFEnviats.GetEnumerator
            While oEnumerador.MoveNext
                ls = " update  OINV  set  U_SEIFivox = '" & oEnumerador.Value & "'where docentry  = " & oEnumerador.Key
                SEI_SRV_VOXEL.ConectarSQLNative(go_conn3)
                oSqlcomand = Nothing
                oSqlcomand = New SqlCommand(ls, go_conn3)
                oSqlcomand.ExecuteNonQuery()
                go_conn3.Close()
            End While
            HashFEnviats.Clear()
            '
        Catch ex As Exception
            Dim oEnumerador As IDictionaryEnumerator
            oEnumerador = HashFEnviats.GetEnumerator
            While oEnumerador.MoveNext
                ls = " update  OINV  set  U_SEIFivox = '" & oEnumerador.Value & "'where docentry  = " & oEnumerador.Key
                SEI_SRV_VOXEL.ConectarSQLNative(go_conn3)
                oSqlcomand = Nothing
                oSqlcomand = New SqlCommand(ls, go_conn3)
                oSqlcomand.ExecuteNonQuery()
                go_conn3.Close()
            End While
            HashFEnviats.Clear()
            Me.Form.lblmsg.Text = ex.Message
        Finally
            If Not IsNothing(oDataReader) Then
                oDataReader.Close()
            End If
        End Try
    End Sub
    '
    Private Function donaMetodoPago(ByVal sappagid As String, ByRef descripcio As String) As String
        Dim ls As String
        Dim oRcs As SqlClient.SqlDataReader = Nothing
        Dim oSqlcomand As SqlCommand
        Dim coonLocal As SqlConnection
        '

        ''' aquí ha d'anar el mapeig  amb el metodos pago de Voxel , la taual seva
        ls = ""
        'ls = ls & " Select U_CBGMetodo, Name from [@CBGPAYMETHVOXEL]  "
        'ls = ls & " where code = '" & sappagid & "'"

        ls = ls & " Select   U_CBGMetodo, Name  "
        ls = ls & " from [@CBGPAYMETHVOXEL] "
        ls = ls & " Left Join crd2 on [@CBGPAYMETHVOXEL].Code = CRD2.PymCode  "
        ls = ls & " where CRD2.CardCode = '" & sappagid & "' "

        '
        Try
            '
            SEI_SRV_VOXEL.ConectarSQLNative(coonLocal)
            oSqlcomand = New SqlCommand(ls, coonLocal)
            oRcs = oSqlcomand.ExecuteReader()
            While oRcs.Read()
                donaMetodoPago = oRcs("U_CBGMetodo").ToString
                descripcio = oRcs("Name").ToString
                Exit Function
            End While



            donaMetodoPago = ""

            '
            oRcs = Nothing
        Catch ex As Exception
            Throw New Exception(ex.Message)
        Finally
            oRcs = Nothing
            oSqlcomand = Nothing
        End Try
        '
    End Function

    '
    Private Function donaImpost(ByVal VatGroup As String) As String

        If VatGroup.Contains("IGIC") Then
            Return "IGIC"
        End If
        If VatGroup.Contains("IVA") Then
            Return "IVA"
        End If
        Select Case VatGroup

            Case Is = "R0", "R0TR", "R1", "R2", "R3", "RA", "SI0", "SI1", "SI2", "SI3", "I0", "I1", "I2", "I3", "IBI0", "IBI1", "IBI2", "IBI3", "ND0", "ND1", "ND2", "ND3"
                Return "IVA"
            Case Is = "RIGIC0", "RIGIC1", "RIGIC13", "RIGIC2", "RIGIC5", "SIGIC0", "SIGIC13", "SIGIC2", "SIGIC5", "RIGIC3", "RIGIC6", "RIGIC7"
                Return "IGIC"
            Case Else
                Return "IVA"
        End Select
    End Function
    Private Function ObtenerRcsTiposIvas(ByVal lDocEntry As Long) As ADODB.Recordset
        '
        Dim ls As String
        Dim oRcs As SAPbobsCOM.Recordset
        Dim sCab As String = "OINV"
        Dim sDet As String = "INV1"
        Dim oRecordset As ADODB.Recordset
        Dim Conn1, conn2 As New ADODB.Connection
        '
        '''' aquí configuro connexió addob.
        Conn1 = New ADODB.Connection
        Obre_Connexio_ADO(Conn1)
        '
        ls = ""
        ls = ls & " SELECT  T1.VatGroup,T3.Rate, "
        ls = ls & " SUM(ISNULL(T1.VatSum,0))    as IMPIMP,"
        ls = ls & " SUM(ISNULL(T1.LineTotal,0)) AS BASEIMP"
        ls = ls & " FROM  " & sCab & " T0 INNER JOIN  " & sDet & " T1"
        ls = ls & " ON T0.DocEntry=T1.DocEntry"
        ' Tabla grupo de descuentos
        ls = ls & " LEFT OUTER JOIN OVTG T3 "
        ls = ls & " ON T1.VatGroup = T3.Code "
        ls = ls & " Where T0.DocEntry = " & Trim(lDocEntry)
        ls = ls & " GROUP BY T1.VatGroup,T3.Rate"
        ls = ls & " ORDER BY T1.VatGroup"
        '' ''oRecordset = Nothing
        '' ''oRecordset = Conn1.Execute(ls)
        '
        Dim cnStr As String
        Dim query As String
        cnStr = "Provider=SQLNCLI11;" &
                       "Server=" & IniGet(Application.StartupPath & "\S_SEI_CBG_VOXEL.ini", "Parametros", "S") & ";" &
                       "Database=" & IniGet(Application.StartupPath & "\S_SEI_CBG_VOXEL.ini", "Parametros", "D") & ";" &
                       "UID=" & IniGet(Application.StartupPath & "\S_SEI_CBG_VOXEL.ini", "Parametros", "U") & ";" &
                       "PWD=" & IniGet(Application.StartupPath & "\S_SEI_CBG_VOXEL.ini", "Parametros", "P") & ";"
        query = ls
        oRecordset = New ADODB.Recordset()
        oRecordset.Open(query, cnStr, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic, -1)
        ObtenerRcsTiposIvas = oRecordset
        '
        oRecordset = Nothing
        '
    End Function

    Private Sub UPDATE_SBO_CLIENTE(ByRef oDataReader As SqlClient.SqlDataReader)
        Dim oCliente As SAPbobsCOM.BusinessPartners = Nothing
        Try
            oCliente = oCompany.GetBusinessObject(oBusinessPartners)
            If oCliente.GetByKey(oDataReader("Code").ToString) Then
                oCliente.EmailAddress = oDataReader("Email").ToString
                oCliente.Fax = oDataReader("Fax").ToString
                oCliente.Phone1 = oDataReader("Phone").ToString
                oCliente.Phone2 = oDataReader("Phone2").ToString
                oCliente.Cellular = oDataReader("Movil").ToString
                '
                If oCliente.Update <> 0 Then
                    Throw New Exception(RecuperarErrorSap())
                End If
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        Finally
            SEI_Globals.LiberarObjCOM(oCliente)
        End Try
        ''''  
    End Sub
    '

    Private Function ObtenerRcsFactLin(ByVal lDocEntry As String) As SqlClient.SqlDataReader
        Dim ls As String
        Dim oRcs As SqlClient.SqlDataReader = Nothing
        Dim oSqlcomand As SqlCommand
        Dim coonLocal As SqlConnection
        '
        ls = ""
        ls = ls & " SELECT  T1.DocEntry, T1.LineNum,T2.CodeBars,T1.ItemCode ,"
        ls = ls & " T1.Dscription,T1.Quantity,T1.PriceBefDi,T1.Price,T1.VatGroup,T3.Rate,T1.VatSum,T1.LineTotal,"
        ls = ls & "  T1.DiscPrcnt  , T2.SalUnitMsr  "
        ls = ls & " FROM  OINV T0 INNER JOIN  INV1 T1"
        ls = ls & " ON T0.DocEntry=T1.DocEntry"
        ls = ls & " INNER JOIN OITM T2"
        ls = ls & " ON T1.ItemCode=T2.ItemCode"
        ls = ls & " LEFT OUTER JOIN OVTG T3 "
        ls = ls & " ON T1.VatGroup = T3.Code "
        ls = ls & " Where T0.DocEntry = " & lDocEntry
        ls = ls & " ORDER BY T1.DocEntry,T1.LineNum"
        '
        Try
            '
            SEI_SRV_VOXEL.ConectarSQLNative(coonLocal)
            oSqlcomand = New SqlCommand(ls, coonLocal)
            oRcs = oSqlcomand.ExecuteReader()
            ObtenerRcsFactLin = oRcs
            '
            oRcs = Nothing
        Catch ex As Exception
            Throw New Exception(ex.Message)
        Finally
            oRcs = Nothing
            oSqlcomand = Nothing
        End Try
        '
    End Function



    Public Shared Function GetEmbeddedResource(ByVal p_objTypeForNameSpace As Type, ByVal p_strScriptFileName As String) As String
        Dim s As StringBuilder = New StringBuilder
        Dim ass As [Assembly] = [Assembly].GetAssembly(p_objTypeForNameSpace)
        Dim sr As StreamReader
        '
        sr = New StreamReader(ass.GetManifestResourceStream(p_objTypeForNameSpace, p_strScriptFileName))
        s.Append(sr.ReadToEnd())
        '
        Return s.ToString()
        '
    End Function


    Private Function ObtenerXML(ByVal sFileName As String) As XmlDocument
        Dim oXMLDocument As XmlDocument = New XmlDocument
        oXMLDocument.LoadXml(GetEmbeddedResource(Me.GetType, sFileName))
        'SetFormPosition(oXMLDocument)
        Return oXMLDocument
    End Function


    Private Sub XML_Linea(ByRef oXML As Xml.XmlDocument,
                          ByRef oRcs As SqlClient.SqlDataReader,
                          ByRef iFila As Integer, ByVal esHotelsCatalonia As Boolean)
        '
        Dim oItem As Xml.XmlNodeList
        Dim oDocumentLines As Xml.XmlNode
        Dim oFirstRow As Xml.XmlNode
        Dim oNewRow As Xml.XmlNode
        '
        If iFila > 0 Then
            ' hauré de fer les referències abans 
            'Lineas Documento (Pedido de  Ventas)
            oDocumentLines = oXML.SelectSingleNode("//ProductList")
            'get the first row 
            oFirstRow = oDocumentLines.FirstChild
            'copy the first row the th new one -> for getting the same structure
            oNewRow = oFirstRow.CloneNode(True)
            'add the new row to the DocumentLines object
            oDocumentLines.AppendChild(oNewRow)
        End If
        '
        'Items
        oItem = oXML.SelectNodes("//ProductList/Product")
        oItem.Item(iFila).Attributes("SupplierSKU").InnerText = oRcs("ItemCode").ToString   ' Código de artículo interno del proveedor
        '''' oItem.Item(iFila).Attributes("CustomerSKU").InnerText = oRcs("ItemCode").ToString  ' Código de artículo interno del cliente
        oItem.Item(iFila).Attributes("Item").InnerText = String.Format("{0:0}", oRcs("Dscription").ToString).Replace(",", ".")
        If Not esHotelsCatalonia Then
            oItem.Item(iFila).Attributes("Qty").InnerText = String.Format("{0:0.000000}", oRcs("Quantity").ToString).Replace(",", ".")
        Else
            oItem.Item(iFila).Attributes("Qty").InnerText = String.Format("{0:0.000}", oRcs("Quantity")).ToString.Replace(",", ".")
        End If

        oItem.Item(iFila).Attributes("MU").InnerText = oRcs("SalUnitMsr").ToString
        oItem.Item(iFila).Attributes("Total").InnerText = String.Format("{0:0.00}", oRcs("PriceBefDi") * oRcs("Quantity")).Replace(",", ".")
        If Not esHotelsCatalonia Then
            oItem.Item(iFila).Attributes("UP").InnerText = oRcs("PriceBefDi").ToString.Replace(",", ".")
        Else
            oItem.Item(iFila).Attributes("UP").InnerText = String.Format("{0:0.000}", oRcs("PriceBefDi")).ToString.Replace(",", ".")
        End If

        '''   If Convert.ToDouble(oRcs("DiscPrcnt").ToString) <> 0 Then
        oItem = oXML.SelectNodes("//ProductList/Product/Discounts/Discount")
        oItem.Item(iFila).Attributes("Amount").InnerText = String.Format("{0:0.00}", (oRcs("Quantity") * (oRcs("PriceBefDi") * oRcs("DiscPrcnt") / 100))).Replace(",", ".")

        ' '' ''If oRcs("DiscPrcnt") > 0 Then
        ' '' ''    MsgBox("Factura amb descompte " & oRcs("DocNum"))
        ' '' ''End If
        ''' End If
        ''' If Convert.ToDouble(oRcs("VatSum").ToString) <> 0 Then
        oItem = oXML.SelectNodes("//ProductList/Product/Taxes/Tax")
        If Not esHotelsCatalonia Then
            oItem.Item(iFila).Attributes("Type").InnerText = donaImpost(oRcs("VatGroup")).ToString

            oItem.Item(iFila).Attributes("Amount").InnerText = oRcs("VatSum").ToString.Replace(",", ".")
            oItem.Item(iFila).Attributes("Rate").InnerText = oRcs("rate").ToString.Replace(",", ".")
        Else
            oItem.Item(iFila).Attributes("Type").InnerText = donaImpost(oRcs("VatGroup")).ToString
            oItem.Item(iFila).Attributes("Amount").InnerText = String.Format("{0:0.000}", oRcs("VatSum")).ToString.Replace(",", ".")
            oItem.Item(iFila).Attributes("Rate").InnerText = String.Format("{0:0.000}", oRcs("rate")).ToString.Replace(",", ".")
        End If
        ''' End If
    End Sub


    Private Sub Obre_Connexio_ADO(ByRef Conn1 As ADODB.Connection)
        '''   Dim Conn1 As New ADODB.Connection
        Dim Cmd1 As New ADODB.Command
        Dim Errs1 As ADODB.Errors
        Dim Rs1 As New ADODB.Recordset
        Dim i As Integer
        Dim AccessConnect As String
        ' Error Handling Variables
        Dim strTmp As String
        On Error GoTo AdoError  ' Full Error Handling which traverses
        ' Connection object
        Dim sSQLstring As String

        sSQLstring = "Provider=SQLNCLI11;" &
                            "Server=" & IniGet(Application.StartupPath & "\S_SEI_CBG_VOXEL.ini", "Parametros", "S") & ";" &
                            "Database=" & IniGet(Application.StartupPath & "\S_SEI_CBG_VOXEL.ini", "Parametros", "D") & ";" &
                            "UID=" & IniGet(Application.StartupPath & "\S_SEI_CBG_VOXEL.ini", "Parametros", "U") & ";" &
                            "PWD=" & IniGet(Application.StartupPath & "\S_SEI_CBG_VOXEL.ini", "Parametros", "P") & ";"

        Conn1.Open(sSQLstring)

Done:
        Rs1 = Nothing
        Cmd1 = Nothing
        Exit Sub

AdoError:
        i = 1
        On Error Resume Next

        ' Enumerate Errors collection and display properties of
        ' each Error object (if Errors Collection is filled out)
        Errs1 = Conn1.Errors
        Dim errLoop As ADODB.Error

        'tret pa no se de quina clase es tracat errLoop
        For Each errLoop In Errs1
            With errLoop
                strTmp = strTmp & vbCrLf & "ADO Error # " & i & ":"
                strTmp = strTmp & vbCrLf & "   ADO Error   # " & .Number
                strTmp = strTmp & vbCrLf & "   Description   " & .Description
                strTmp = strTmp & vbCrLf & "   Source        " & .Source
                i = i + 1
            End With
        Next

AdoErrorLite:
        ' Get VB Error Object's information
        strTmp = strTmp & vbCrLf & "VB Error # " & Str(Err.Number)
        strTmp = strTmp & vbCrLf & "   Generated by " & Err.Source
        strTmp = strTmp & vbCrLf & "   Description  " & Err.Description
        MsgBox(strTmp)
        ' Clean up gracefully without risking infinite loop in error handler
        On Error GoTo 0
        GoTo Done

    End Sub
#End Region

End Class
