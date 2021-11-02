Public Class CatchingEvents

    Friend WithEvents SBOApplication As SAPbouiCOM.Application '//OBJETO DE APLICACION
    Friend SBOCompany As SAPbobsCOM.Company '//OBJETO COMPAÑIA
    Friend csDirectory As String '//DIRECTORIO DONDE SE ENCUENTRAN LOS .SRF
    Dim DocNum As String


    Public Sub New()

        MyBase.New()
        SetAplication()
        SetConnectionContext()
        ConnectSBOCompany()

        setFilters()

    End Sub


    '//----- ESTABLECE LA COMUNICACION CON SBO
    Private Sub SetAplication()
        Dim SboGuiApi As SAPbouiCOM.SboGuiApi
        Dim sConnectionString As String
        Try
            SboGuiApi = New SAPbouiCOM.SboGuiApi
            sConnectionString = Environment.GetCommandLineArgs.GetValue(1)
            SboGuiApi.Connect(sConnectionString)
            SBOApplication = SboGuiApi.GetApplication()
        Catch ex As Exception
            SBOApplication.MessageBox("Falló la conexión con la aplicación SBO " & ex.Message)
            System.Windows.Forms.Application.Exit()
            End '//termina aplicación
        End Try
    End Sub


    '//----- ESTABLECE EL CONTEXTO DE LA APLICACION
    Private Sub SetConnectionContext()
        Try
            SBOCompany = SBOApplication.Company.GetDICompany
        Catch ex As Exception
            SBOApplication.MessageBox("Falló la conexión con el DI")
            SBOApplication = Nothing
            System.Windows.Forms.Application.Exit()
            End '//termina aplicación
            'Finally
        End Try
    End Sub


    '//----- CONEXION CON LA BASE DE DATOS
    Private Sub ConnectSBOCompany()
        Dim loRecSet As SAPbobsCOM.Recordset
        Try
            '//ESTABLECE LA CONEXION A LA COMPAÑIA
            csDirectory = My.Application.Info.DirectoryPath
            If (csDirectory = "") Then
                System.Windows.Forms.Application.Exit()
                End
            End If
        Catch ex As Exception
            SBOApplication.MessageBox("Falló la conexión con la BD. " & ex.Message)
            SBOApplication = Nothing
            System.Windows.Forms.Application.Exit()
            End '//termina aplicación
        Finally
            loRecSet = Nothing
        End Try
    End Sub


    '//----- ESTABLECE FILTROS DE EVENTOS DE LA APLICACION
    Private Sub setFilters()
        Dim lofilter As SAPbouiCOM.EventFilter
        Dim lofilters As SAPbouiCOM.EventFilters

        Try
            lofilters = New SAPbouiCOM.EventFilters
            lofilter = lofilters.Add(SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED)
            lofilter.AddEx(1250000940) '// FORMA Solicitud de traslado

            SBOApplication.SetFilter(lofilters)

        Catch ex As Exception
            SBOApplication.MessageBox("SetFilter: " & ex.Message)
        End Try

    End Sub


    ''//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    ''// METODOS PARA EVENTOS DE LA APLICACION
    ''//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

    Private Sub SBO_Application_AppEvent(ByVal EventType As SAPbouiCOM.BoAppEventTypes) Handles SBOApplication.AppEvent
        Select Case EventType
            Case SAPbouiCOM.BoAppEventTypes.aet_ShutDown, SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition, SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged
                System.Windows.Forms.Application.Exit()
                End
        End Select

    End Sub


    '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '// METODOS PARA MANEJO DE EVENTOS ITEM
    '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Private Sub SBOApplication_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles SBOApplication.ItemEvent

        If pVal.Before_Action = True And pVal.FormTypeEx <> "" Then

            Select Case pVal.FormTypeEx

                Case 1250000940                           '////// FORMA Solicitud de traslado
                    frmOWTRControllerBefore(FormUID, pVal)

            End Select

        Else
            If pVal.Before_Action = False And pVal.FormTypeEx <> "" Then
                Select Case pVal.FormTypeEx

                    Case 1250000940                           '////// FORMA Solicitud de traslado
                        frmOWTRControllerAfter(FormUID, pVal)

                End Select
            End If
        End If

    End Sub


    Private Sub frmOWTRControllerBefore(ByVal FormUID As String, ByVal pVal As SAPbouiCOM.ItemEvent)

        Dim coForm As SAPbouiCOM.Form
        Dim stTabla As String
        Dim oDatatable As SAPbouiCOM.DBDataSource

        Try

            Select Case pVal.EventType

                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                    Select Case pVal.ItemUID

                        Case 1

                            stTabla = "OWTQ"
                            coForm = SBOApplication.Forms.Item(FormUID)

                            oDatatable = coForm.DataSources.DBDataSources.Item(stTabla)
                            DocNum = oDatatable.GetValue("DocNum", 0)

                    End Select

            End Select

        Catch ex As Exception
            SBOApplication.MessageBox("Error en el evento sobre Forma Facturacion Clientes. " & ex.Message)
        Finally
            'oPO = Nothing
        End Try

    End Sub


    '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '// CONTROLADOR DE EVENTOS FORMA Solicitud de traslado
    '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Private Sub frmOWTRControllerAfter(ByVal FormUID As String, ByVal pVal As SAPbouiCOM.ItemEvent)
        Dim DocEntry, ObjType, LineNum, ItemCode, FromWhsCod, WhsCode, BatchNumber, DocNumST As String
        Dim Quantity As Double
        Dim stQueryH1, stQueryH2, stQueryH3 As String
        Dim oRecSetH1, oRecSetH2, oRecSetH3 As SAPbobsCOM.Recordset
        Dim oStockTransfer As SAPbobsCOM.StockTransfer
        Dim CantidadR, CantidadL As Double
        Dim llError As Long
        Dim lsError As String
        Dim AOWTR As Integer
        Dim oED As FrmtekEDocument

        oRecSetH1 = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecSetH2 = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecSetH3 = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        oStockTransfer = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oStockTransfer)

        Try

            Select Case pVal.EventType

                '//////Evento Presionar Item
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                    Select Case pVal.ItemUID

                        Case 1

                            stQueryH1 = "Select T1.""DocEntry"",T0.""ObjType"",T1.""LineNum"",T1.""ItemCode"",T1.""FromWhsCod"",T1.""WhsCode"",T1.""Quantity"" from OWTQ T0 Inner Join WTQ1 T1 on T1.""DocEntry""=T0.""DocEntry"" where T0.""DocNum""=" & DocNum
                            oRecSetH1.DoQuery(stQueryH1)

                            If oRecSetH1.RecordCount > 0 Then

                                oStockTransfer.DocDate = DateTime.Now
                                oStockTransfer.ElectronicProtocols.GenerationType = 1
                                oStockTransfer.ElectronicProtocols.Add()

                                oRecSetH1.MoveFirst()

                                For i = 0 To oRecSetH1.RecordCount - 1

                                    DocEntry = oRecSetH1.Fields.Item("DocEntry").Value
                                    ObjType = oRecSetH1.Fields.Item("ObjType").Value
                                    LineNum = oRecSetH1.Fields.Item("LineNum").Value
                                    ItemCode = oRecSetH1.Fields.Item("ItemCode").Value
                                    FromWhsCod = oRecSetH1.Fields.Item("FromWhsCod").Value
                                    WhsCode = oRecSetH1.Fields.Item("WhsCode").Value
                                    Quantity = oRecSetH1.Fields.Item("Quantity").Value

                                    oStockTransfer.Lines.BaseEntry = DocEntry
                                    oStockTransfer.Lines.BaseType = 5
                                    oStockTransfer.Lines.BaseLine = LineNum
                                    oStockTransfer.Lines.ItemCode = ItemCode
                                    oStockTransfer.Lines.FromWarehouseCode = FromWhsCod
                                    oStockTransfer.Lines.WarehouseCode = WhsCode
                                    oStockTransfer.Lines.Quantity = Quantity

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
                                                oStockTransfer.Lines.BatchNumbers.BaseLineNumber = LineNum

                                                oStockTransfer.Lines.BatchNumbers.Add()

                                                l = 0

                                            Else

                                                BatchNumber = oRecSetH2.Fields.Item("BatchNum").Value

                                                oStockTransfer.Lines.BatchNumbers.BatchNumber = BatchNumber
                                                oStockTransfer.Lines.BatchNumbers.Quantity = CantidadR
                                                oStockTransfer.Lines.BatchNumbers.BaseLineNumber = LineNum

                                                oStockTransfer.Lines.BatchNumbers.Add()

                                                l = oRecSetH2.RecordCount - 1

                                            End If

                                            oRecSetH2.MoveNext()

                                        Next

                                    End If

                                    oStockTransfer.Lines.Add()
                                    oRecSetH1.MoveNext()

                                Next

                                If oStockTransfer.Add() <> 0 Then

                                    SBOCompany.GetLastError(llError, lsError)
                                    Err.Raise(-1, 1, lsError)

                                Else

                                    AOWTR = SBOCompany.GetNewObjectKey().ToString()
                                    stQueryH3 = "Select ""DocNum"" from OWTR where ""DocEntry""=" & AOWTR
                                    oRecSetH3.DoQuery(stQueryH3)

                                    If oRecSetH3.RecordCount = 1 Then

                                        DocNumST = oRecSetH3.Fields.Item("DocNum").Value
                                        oED = New FrmtekEDocument
                                        oED.openForm(csDirectory, AOWTR)

                                    End If

                                End If

                            End If

                    End Select

            End Select

        Catch ex As Exception
            SBOApplication.MessageBox("Error en el evento sobre Forma Solicitud de traslado. " & ex.Message)
        Finally
            'oPO = Nothing
        End Try
    End Sub


End Class
