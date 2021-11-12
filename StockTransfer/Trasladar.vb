Imports System.Windows.Forms

Public Class Trasladar


    Private cSBOApplication As SAPbouiCOM.Application '//OBJETO DE APLICACION
    Private cSBOCompany As SAPbobsCOM.Company     '//OBJETO DE CONEXION
    Dim conexionSQL As Sap.Data.Hana.HanaConnection

    '//----- METODO DE CREACION DE LA CLASE
    Public Sub New()
        MyBase.New()
        cSBOApplication = oCatchingEvents.SBOApplication
        cSBOCompany = oCatchingEvents.SBOCompany
        conectar()
    End Sub

    Public Function conectar() As Boolean
        Dim stCadenaConexion As String
        Try

            conectar = False

            ''---- Cargamos datos de archivo de configuracion

            '---- objeto compañia
            conexionSQL = New Sap.Data.Hana.HanaConnection

            '---- armamos cadena de conexion
            stCadenaConexion = "DRIVER={B1CRHPROXY32};UID=" & My.Settings.UserSQL & ";PWD=" & My.Settings.PassSQL & ";SERVERNODE=" & My.Settings.Server

            '---- realizamos conexion
            conexionSQL = New Sap.Data.Hana.HanaConnection(stCadenaConexion)

            conexionSQL.Open()

        Catch ex As Exception
            cSBOApplication.MessageBox("Error al conectar con HANA . " & ex.Message)
        End Try
    End Function

    Public Function AddTransfer(ByVal csDirectory As String, ByVal DocNum As String, ByVal FormUID As String)

        Dim DocEntry, ObjType, LineNum, ItemCode, VisOrder, FromWhsCod, WhsCode, BatchNumber, DocNumST, Lote As String
        Dim Quantity As Double
        Dim stQueryH1, stQueryH2, stQueryH3 As String
        Dim oRecSetH1, oRecSetH2, oRecSetH3 As SAPbobsCOM.Recordset
        Dim oStockTransfer As SAPbobsCOM.StockTransfer
        Dim CantidadR, CantidadL As Double
        Dim llError As Long
        Dim lsError As String
        Dim AOWTR As Integer
        Dim oED As FrmtekEDocument
        Dim tabla As DataTable
        Dim comm As New Sap.Data.Hana.HanaCommand
        Dim DA As New Sap.Data.Hana.HanaDataAdapter
        Dim ds As New DataSet

        oRecSetH1 = cSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecSetH2 = cSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecSetH3 = cSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        oStockTransfer = cSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oStockTransfer)

        Try

            stQueryH1 = "Select T1.""DocEntry"",T0.""ObjType"",T1.""LineNum"",T1.""ItemCode"",T1.""VisOrder"",T1.""FromWhsCod"",T1.""WhsCode"",T1.""Quantity"",T2.""ManBtchNum"" from """ & cSBOCompany.CompanyDB & """.OWTQ T0 Inner Join """ & cSBOCompany.CompanyDB & """.WTQ1 T1 on T1.""DocEntry""=T0.""DocEntry"" Inner Join """ & cSBOCompany.CompanyDB & """.OITM T2 on T2.""ItemCode""=T1.""ItemCode"" where T0.""DocNum""=" & DocNum
            oRecSetH1.DoQuery(stQueryH1)
            comm.CommandText = stQueryH1
            comm.Connection = conexionSQL
            DA.SelectCommand = comm
            DA.Fill(ds)


            If ds.Tables(0).Rows.Count > 0 Then

                tabla = ds.Tables(0)

                oStockTransfer.DocDate = DateTime.Now
                oStockTransfer.FromWarehouse = oRecSetH1.Fields.Item("FromWhsCod").Value
                oStockTransfer.ToWarehouse = oRecSetH1.Fields.Item("WhsCode").Value
                oStockTransfer.Comments = "Basado en la solicitud " & DocNum
                oStockTransfer.ElectronicProtocols.GenerationType = 1
                oStockTransfer.ElectronicProtocols.Add()

                oRecSetH1.MoveFirst()

                For Each i As DataRow In tabla.Rows
                    'Each i As DataRow In tabla.Rows '11 seg
                    'i = 0 To oRecSetH1.RecordCount - 1 '9, 10, 11

                    DocEntry = oRecSetH1.Fields.Item("DocEntry").Value
                    ObjType = oRecSetH1.Fields.Item("ObjType").Value
                    LineNum = oRecSetH1.Fields.Item("LineNum").Value
                    VisOrder = oRecSetH1.Fields.Item("VisOrder").Value
                    ItemCode = oRecSetH1.Fields.Item("ItemCode").Value
                    FromWhsCod = oRecSetH1.Fields.Item("FromWhsCod").Value
                    WhsCode = oRecSetH1.Fields.Item("WhsCode").Value
                    Quantity = oRecSetH1.Fields.Item("Quantity").Value
                    Lote = oRecSetH1.Fields.Item("ManBtchNum").Value

                    'oStockTransfer.Lines.BaseEntry = DocEntry
                    'oStockTransfer.Lines.BaseType = 5
                    'oStockTransfer.Lines.BaseLine = LineNum
                    oStockTransfer.Lines.ItemCode = ItemCode
                    oStockTransfer.Lines.FromWarehouseCode = FromWhsCod
                    oStockTransfer.Lines.WarehouseCode = WhsCode
                    oStockTransfer.Lines.Quantity = Quantity

                    If Lote = "Y" Then

                        stQueryH2 = "Select T0.*,T1.""CreateDate"" from
                                                    (Select ""BatchNum"",""ItemCode"",""WhsCode"",
                                                    sum(case when ""Direction""=0 then ""Quantity"" else -1*""Quantity"" end) as ""CantidadLote"" 
                                                    from IBT1 where ""ItemCode""='" & ItemCode & "' AND ""WhsCode""='" & FromWhsCod & "'
                                                    Group by  ""BatchNum"",""ItemCode"",""WhsCode"") T0
                                                    Inner Join OBTN T1 on T1.""DistNumber""=T0.""BatchNum"" and T1.""ItemCode""=T0.""ItemCode""
                                                    where T0.""CantidadLote"">0
                                                    order by T1.""CreateDate"""
                        oRecSetH2.DoQuery(stQueryH2)

                        If oRecSetH2.RecordCount > 0 Then

                            oRecSetH2.MoveFirst()
                            CantidadR = Quantity

                            For l = 0 To oRecSetH2.RecordCount - 1

                                CantidadL = oRecSetH2.Fields.Item("CantidadLote").Value

                                If CantidadR > CantidadL Then

                                    CantidadR = CantidadR - CantidadL

                                    BatchNumber = oRecSetH2.Fields.Item("BatchNum").Value

                                    oStockTransfer.Lines.BatchNumbers.BatchNumber = BatchNumber
                                    oStockTransfer.Lines.BatchNumbers.Quantity = CantidadL
                                    'oStockTransfer.Lines.BatchNumbers.BaseLineNumber = LineNum

                                    oStockTransfer.Lines.BatchNumbers.Add()

                                    l = 0

                                Else

                                    BatchNumber = oRecSetH2.Fields.Item("BatchNum").Value

                                    oStockTransfer.Lines.BatchNumbers.BatchNumber = BatchNumber
                                    oStockTransfer.Lines.BatchNumbers.Quantity = CantidadR
                                    'oStockTransfer.Lines.BatchNumbers.BaseLineNumber = LineNum

                                    oStockTransfer.Lines.BatchNumbers.Add()

                                    l = oRecSetH2.RecordCount - 1

                                End If

                                oRecSetH2.MoveNext()

                            Next

                        End If

                    End If

                    oStockTransfer.Lines.Add()
                    oRecSetH1.MoveNext()

                Next

                If oStockTransfer.Add() <> 0 Then

                    cSBOCompany.GetLastError(llError, lsError)
                    Err.Raise(-1, 1, lsError)
                    conexionSQL.Close()

                Else

                    AOWTR = cSBOCompany.GetNewObjectKey().ToString()
                    stQueryH3 = "Select ""DocNum"" from OWTR where ""DocEntry""=" & AOWTR
                    oRecSetH3.DoQuery(stQueryH3)

                    If oRecSetH3.RecordCount = 1 Then

                        DocNumST = oRecSetH3.Fields.Item("DocNum").Value

                        oStockTransfer = cSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryTransferRequest)

                        oStockTransfer.GetByKey(DocEntry)
                        oStockTransfer.Comments = "Cerrado por el traslado " & DocNumST
                        oStockTransfer.Update()

                        oStockTransfer.Close()

                        oED = New FrmtekEDocument
                        oED.openForm(csDirectory, AOWTR)

                        conexionSQL.Close()

                    End If

                End If

            End If

        Catch ex As Exception

            cSBOApplication.MessageBox("Error al crear el traslado. " & ex.Message)
            conexionSQL.Close()

        End Try

    End Function


End Class
