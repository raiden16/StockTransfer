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

        Dim DocEntry, ObjType, LineNum, ItemCode, VisOrder, FromWhsCod, WhsCode, BatchNumber, DocNumST, Lote, Boxes, Delivery, Package As String
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

            stQueryH1 = "Select T1.""DocEntry"",T0.""ObjType"",T1.""LineNum"",T1.""ItemCode"",T1.""VisOrder"",T1.""FromWhsCod"",T1.""WhsCode"",T1.""Quantity"",T1.""U_CajasReq"",T1.""U_DeliveryType"",T1.""U_NumPaq"",T2.""ManBtchNum"" from """ & cSBOCompany.CompanyDB & """.OWTQ T0 Inner Join """ & cSBOCompany.CompanyDB & """.WTQ1 T1 on T1.""DocEntry""=T0.""DocEntry"" Inner Join """ & cSBOCompany.CompanyDB & """.OITM T2 on T2.""ItemCode""=T1.""ItemCode"" where T0.""DocNum""=" & DocNum
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

                CreateTemporalTable()

                InsertTemporalTable(DocNum)

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
                    Boxes = oRecSetH1.Fields.Item("U_CajasReq").Value
                    Delivery = oRecSetH1.Fields.Item("U_DeliveryType").Value
                    Package = oRecSetH1.Fields.Item("U_NumPaq").Value

                    'oStockTransfer.Lines.BaseEntry = DocEntry
                    'oStockTransfer.Lines.BaseType = 5
                    'oStockTransfer.Lines.BaseLine = LineNum
                    oStockTransfer.Lines.ItemCode = ItemCode
                    oStockTransfer.Lines.FromWarehouseCode = FromWhsCod
                    oStockTransfer.Lines.WarehouseCode = WhsCode
                    oStockTransfer.Lines.Quantity = Quantity
                    oStockTransfer.Lines.UserFields.Fields.Item("U_CajasReq").Value = Boxes
                    oStockTransfer.Lines.UserFields.Fields.Item("U_DeliveryType").Value = Delivery
                    oStockTransfer.Lines.UserFields.Fields.Item("U_NumPaq").Value = Package

                    If Lote = "Y" Then

                        'stQueryH2 = "Select T0.*,T1.""CreateDate"" from
                        '                            (Select ""BatchNum"",""ItemCode"",""WhsCode"",
                        '                            sum(case when ""Direction""=0 then ""Quantity"" else -1*""Quantity"" end) as ""CantidadLote"" 
                        '                            from IBT1 where ""ItemCode""='" & ItemCode & "' AND ""WhsCode""='" & FromWhsCod & "'
                        '                            Group by  ""BatchNum"",""ItemCode"",""WhsCode"") T0
                        '                            Inner Join OBTN T1 on T1.""DistNumber""=T0.""BatchNum"" and T1.""ItemCode""=T0.""ItemCode""
                        '                            where T0.""CantidadLote"">0
                        '                            order by T1.""CreateDate"""
                        stQueryH2 = "Select * from """ & cSBOCompany.CompanyDB & """.ListaLotes where ""ITEMCODE""='" & ItemCode & "' and ""CANTIDADLOTE"">0 order by ""CREATEDATE"" Desc"
                        oRecSetH2.DoQuery(stQueryH2)

                        If oRecSetH2.RecordCount > 0 Then

                            oRecSetH2.MoveFirst()
                            CantidadR = Format(Quantity, "0.000")

                            For l = 0 To oRecSetH2.RecordCount - 1

                                CantidadL = Format(oRecSetH2.Fields.Item("CANTIDADLOTE").Value, "0.000")

                                If CantidadR > CantidadL Then

                                    CantidadR = Format(CantidadR - CantidadL, "0.000")

                                    BatchNumber = oRecSetH2.Fields.Item("BatchNum").Value

                                    oStockTransfer.Lines.BatchNumbers.BatchNumber = BatchNumber
                                    oStockTransfer.Lines.BatchNumbers.Quantity = CantidadL
                                    'oStockTransfer.Lines.BatchNumbers.BaseLineNumber = LineNum

                                    oStockTransfer.Lines.BatchNumbers.Add()

                                    UpdateTemporalTable(BatchNumber, CantidadL - CantidadL)

                                    l = 0

                                Else

                                    BatchNumber = oRecSetH2.Fields.Item("BatchNum").Value

                                    oStockTransfer.Lines.BatchNumbers.BatchNumber = BatchNumber
                                    oStockTransfer.Lines.BatchNumbers.Quantity = CantidadR
                                    'oStockTransfer.Lines.BatchNumbers.BaseLineNumber = LineNum

                                    oStockTransfer.Lines.BatchNumbers.Add()

                                    UpdateTemporalTable(BatchNumber, CantidadL - CantidadR)

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

                    DropTemporalTable()

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
            DropTemporalTable()

        End Try

    End Function


    Public Function CreateTemporalTable()

        Dim stQueryH1 As String
        Dim comm As New Sap.Data.Hana.HanaCommand
        Dim DA As New Sap.Data.Hana.HanaDataAdapter
        Dim ds As New DataSet

        Try

            stQueryH1 = "CREATE COLUMN TABLE """ & cSBOCompany.CompanyDB & """.ListaLotes (BatchNum NVARCHAR(50), ItemCode NVARCHAR(50), WhsCode NVARCHAR(5), CantidadLote Double, CreateDate NVARCHAR(20));"
            comm.CommandText = stQueryH1
            comm.Connection = conexionSQL
            DA.SelectCommand = comm
            DA.Fill(ds)

        Catch ex As Exception

            cSBOApplication.MessageBox("Error al CreateTemporalTable. " & ex.Message)
            conexionSQL.Close()

        End Try

    End Function


    Public Function InsertTemporalTable(ByVal DocNum As String)

        Dim stQueryH1, stQueryH2 As String
        Dim comm, comm2 As New Sap.Data.Hana.HanaCommand
        Dim DA, DA2 As New Sap.Data.Hana.HanaDataAdapter
        Dim ds, ds2 As New DataSet
        Dim tabla As DataTable
        Dim BatchNum, ItemCode, WhsCode As String
        Dim CantidadLote As Double
        Dim CreateDate As String

        Try

            stQueryH1 = "Call """ & cSBOCompany.CompanyDB & """.Lotes(" & DocNum & ")"
            comm.CommandText = stQueryH1
            comm.Connection = conexionSQL
            DA.SelectCommand = comm
            DA.Fill(ds)

            If ds.Tables(0).Rows.Count > 0 Then

                tabla = ds.Tables(0)

                For i = 0 To ds.Tables(0).Rows.Count - 1

                    BatchNum = ds.Tables(0).Rows(i).Item("BatchNum")
                    ItemCode = ds.Tables(0).Rows(i).Item("ItemCode")
                    WhsCode = ds.Tables(0).Rows(i).Item("WhsCode")
                    CantidadLote = ds.Tables(0).Rows(i).Item("CantidadLote")
                    CreateDate = ds.Tables(0).Rows(i).Item("CreateDate")

                    stQueryH2 = "Insert Into """ & cSBOCompany.CompanyDB & """.ListaLotes values ('" & BatchNum & "','" & ItemCode & "','" & WhsCode & "'," & CantidadLote & ",'" & CreateDate & "')"
                    comm2.CommandText = stQueryH2
                    comm2.Connection = conexionSQL
                    DA2.SelectCommand = comm2
                    DA2.Fill(ds2)

                Next

            End If

        Catch ex As Exception

            cSBOApplication.MessageBox("Error al InsertTemporalTable. " & ex.Message)
            conexionSQL.Close()

        End Try

    End Function


    Public Function UpdateTemporalTable(ByVal Lote As String, ByVal Cantidad As Double)

        Dim stQueryH1 As String
        Dim comm As New Sap.Data.Hana.HanaCommand
        Dim DA As New Sap.Data.Hana.HanaDataAdapter
        Dim ds As New DataSet

        Try

            stQueryH1 = "Update """ & cSBOCompany.CompanyDB & """.ListaLotes set ""CANTIDADLOTE""=" & Cantidad & " where ""BATCHNUM""='" & Lote & "'"
            comm.CommandText = stQueryH1
            comm.Connection = conexionSQL
            DA.SelectCommand = comm
            DA.Fill(ds)

        Catch ex As Exception

            cSBOApplication.MessageBox("Error al UpdateTemporalTable. " & ex.Message)
            conexionSQL.Close()

        End Try

    End Function


    Public Function DropTemporalTable()

        Dim stQueryH1 As String
        Dim comm As New Sap.Data.Hana.HanaCommand
        Dim DA As New Sap.Data.Hana.HanaDataAdapter
        Dim ds As New DataSet

        Try

            stQueryH1 = "Drop table """ & cSBOCompany.CompanyDB & """.ListaLotes"
            comm.CommandText = stQueryH1
            comm.Connection = conexionSQL
            DA.SelectCommand = comm
            DA.Fill(ds)

        Catch ex As Exception

            cSBOApplication.MessageBox("Error al DropTemporalTable. " & ex.Message)
            conexionSQL.Close()

        End Try

    End Function


End Class
