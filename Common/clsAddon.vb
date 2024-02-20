﻿Imports SAPbouiCOM.Framework
Imports System.IO

Namespace Decimal_Rounding
    Public Class clsAddon
        Public WithEvents objapplication As SAPbouiCOM.Application
        Public objcompany As SAPbobsCOM.Company
        Public objmenuevent As clsMenuEvent
        Public objrightclickevent As clsRightClickEvent
        Public objglobalmethods As clsGlobalMethods
        'Public objProdOrder As ClsProductionOrder
        'Public objBatchSetup As ClsBatchSetup
        Public objManFactOrder As ClsManufacturingOrder
        Public objPickOrder As ClsPickOrder
        Public objQualityControlT As ClsQualityControlTest
        Dim objform As SAPbouiCOM.Form
        Dim strsql As String = "", GetValue As String = "", PickOrderEntry As String = ""
        Dim Qty As Double
        Dim objrs As SAPbobsCOM.Recordset
        Dim print_close As Boolean = False
        Public HANA As Boolean = True
        'Public HANA As Boolean = False
        Dim tempform As SAPbouiCOM.Form
        Dim CurRow As Integer = -1

        Public HWKEY() As String = New String() {"L1653539483", "O1710639360"}

        Public Sub Intialize(ByVal args() As String)
            Try
                Dim oapplication As Application
                If (args.Length < 1) Then oapplication = New Application Else oapplication = New Application(args(0))
                objapplication = Application.SBO_Application
                If isValidLicense() Then
                    objapplication.StatusBar.SetText("Establishing Company Connection Please Wait...", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    objcompany = Application.SBO_Application.Company.GetDICompany()

                    Create_DatabaseFields() 'UDF & UDO Creation Part    
                    'Menu() 'Menu Creation Part
                    Create_Objects() 'Object Creation Part
                    'If HANA Then
                    '    Localization = objglobalmethods.getSingleValue("select ""LawsSet"" from CINF")
                    '    CostCenter = objaddon.objglobalmethods.getSingleValue("select ""MDStyle"" from OADM")
                    '    MainCurr = objaddon.objglobalmethods.getSingleValue("select ""MainCurncy"" from OADM")
                    'Else
                    '    Localization = objglobalmethods.getSingleValue("select LawsSet from CINF")
                    '    CostCenter = objaddon.objglobalmethods.getSingleValue("select MDStyle from OADM")
                    '    MainCurr = objaddon.objglobalmethods.getSingleValue("select MainCurncy from OADM")
                    'End If
                    objapplication.StatusBar.SetText("Addon Connected Successfully..!!!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                    oapplication.Run()

                Else
                    objapplication.StatusBar.SetText("Addon Disconnected due to license mismatch..!!!", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Exit Sub
                End If
                'System.Windows.Forms.Application.Run()
            Catch ex As Exception
                objaddon.objapplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
            End Try
        End Sub

        Function isValidLicense() As Boolean
            Try
                Try
                    If objapplication.Forms.ActiveForm.TypeCount > 0 Then
                        For i As Integer = 0 To objapplication.Forms.ActiveForm.TypeCount - 1
                            objapplication.Forms.ActiveForm.Close()
                        Next
                    End If
                Catch ex As Exception
                End Try

                'If Not HANA Then
                '    objapplication.Menus.Item("1030").Activate()
                'End If
                objapplication.Menus.Item("257").Activate()
                Dim CrrHWKEY As String = objapplication.Forms.ActiveForm.Items.Item("79").Specific.Value.ToString.Trim
                objapplication.Forms.ActiveForm.Close()

                For i As Integer = 0 To HWKEY.Length - 1
                    If HWKEY(i).Trim = CrrHWKEY.Trim Then
                        Return True
                    End If
                Next
                MsgBox("Installing Add-On failed due to License mismatch", MsgBoxStyle.OkOnly, "License Management")
                Return False
            Catch ex As Exception
                objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                'MsgBox(ex.ToString)
            End Try
            Return True
        End Function

        Private Sub Create_Objects()
            objmenuevent = New clsMenuEvent
            objrightclickevent = New clsRightClickEvent
            objglobalmethods = New clsGlobalMethods
            'objProdOrder = New ClsProductionOrder
            'objBatchSetup = New ClsBatchSetup
            objManFactOrder = New ClsManufacturingOrder
            objPickOrder = New ClsPickOrder
            objQualityControlT = New ClsQualityControlTest
        End Sub

        Private Sub Create_DatabaseFields()
            'If objapplication.Company.UserName.ToString.ToUpper <> "MANAGER" Then

            'If objapplication.MessageBox("Do you want to execute the field Creations?", 2, "Yes", "No") <> 1 Then Exit Sub
            objapplication.StatusBar.SetText("Creating Database Fields.Please Wait...", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

            Dim objtable As New clsTable
            objtable.FieldCreation()
            'End If

        End Sub

#Region "Menu Creation Details"

        Private Sub Menu()
            Dim Menucount As Integer = 1
            If objapplication.Menus.Item("43520").SubMenus.Exists("CHKPRT") And objapplication.Menus.Item("43520").SubMenus.Exists("TNKMSTR") Then Return
            'CreateMenu("", Menucount, "Multi-Branch A/P Service Invoice", SAPbouiCOM.BoMenuType.mt_STRING, "MBAPSI", "2304") : Menucount += 1
            Menucount = objapplication.Menus.Item("43520").SubMenus.Count
            CreateMenu("", Menucount, "Check Print", SAPbouiCOM.BoMenuType.mt_POPUP, "CHKPRT", "43520")
            CreateMenu("", Menucount, "Multi-Check Printing", SAPbouiCOM.BoMenuType.mt_STRING, "MULCHK", "CHKPRT") : Menucount += 1 ' "43537"

            CreateMenu("", Menucount, "Tank Master", SAPbouiCOM.BoMenuType.mt_STRING, "TNKMSTR", "4352") : Menucount += 1


        End Sub

        Private Sub CreateMenu(ByVal ImagePath As String, ByVal Position As Int32, ByVal DisplayName As String, ByVal MenuType As SAPbouiCOM.BoMenuType, ByVal UniqueID As String, ByVal ParentMenuID As String)
            Try
                Dim oMenuPackage As SAPbouiCOM.MenuCreationParams
                Dim parentmenu As SAPbouiCOM.MenuItem
                parentmenu = objapplication.Menus.Item(ParentMenuID)
                If parentmenu.SubMenus.Exists(UniqueID.ToString) Then Exit Sub
                oMenuPackage = objapplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
                oMenuPackage.Image = ImagePath
                oMenuPackage.Position = Position
                oMenuPackage.Type = MenuType
                oMenuPackage.UniqueID = UniqueID
                oMenuPackage.String = DisplayName
                parentmenu.SubMenus.AddEx(oMenuPackage)
            Catch ex As Exception
                objapplication.StatusBar.SetText("Menu Already Exists", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_None)
            End Try
            'Return ParentMenu.SubMenus.Item(UniqueID)
        End Sub

#End Region

#Region "ItemEvent_Link Button"

        Private Sub objapplication_ItemEvent(FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles objapplication.ItemEvent
            Try
                'Dim oform As SAPbouiCOM.Form
                Dim objGrid As SAPbouiCOM.Grid
                Dim odbdsDetails, odbdsHeader As SAPbouiCOM.DBDataSource
                'If objapplication.Forms.Count > 0 Then oform = objapplication.Forms.ActiveForm
                objform = objaddon.objapplication.Forms.Item(FormUID)
                If pVal.BeforeAction Then
                    objform = objaddon.objapplication.Forms.Item(FormUID)
                    Select Case pVal.EventType
                        Case SAPbouiCOM.BoEventTypes.et_CLICK
                            If objform.TypeEx.ToUpper() = "CT_PF_QCTEST" And pVal.ItemUID = "TstRstTab" And CurRow <> -1 Then
                                objform.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                                objform.Items.Item("SeriesTbx").Enabled = True
                                objform.Items.Item("SeriesTbx").Specific.string = Trim(tempform.Items.Item("7").Specific.Columns.Item("DocNum").Cells.Item(CurRow).Specific.String)
                                objform.Items.Item("CrdDatTbx").Specific.string = Trim(tempform.Items.Item("7").Specific.Columns.Item("U_Created").Cells.Item(CurRow).Specific.String)
                                objform.Items.Item("1").Click()
                                CurRow = -1 : tempform = Nothing
                                'objform.Select()
                                'objform.ActiveItem = "TstPrNTbx"
                                'objapplication.SendKeys("^+R")

                            ElseIf objform.TypeEx.ToUpper() = "CT_PF_MANUFACORD" And pVal.ItemUID = "tab_2" And CurRow <> -1 Then
                                'objform = objaddon.objapplication.Forms.ActiveForm
                                objform.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                                objform.Items.Item("5").Enabled = True
                                'strsql = tempform.Items.Item("7").Specific.Columns.Item("U_DocNum").Cells.Item(CurRow).Specific.String
                                'strsql = tempform.Items.Item("7").Specific.Columns.Item("U_RequiredDate").Cells.Item(CurRow).Specific.String
                                objform.Items.Item("5").Specific.string = Trim(tempform.Items.Item("7").Specific.Columns.Item("U_DocNum").Cells.Item(CurRow).Specific.String)
                                objform.Items.Item("21").Specific.string = Trim(tempform.Items.Item("7").Specific.Columns.Item("U_RequiredDate").Cells.Item(CurRow).Specific.String)
                                objform.Items.Item("1").Click()
                                CurRow = -1 : tempform = Nothing
                            End If
                            If bModal And (objaddon.objapplication.Forms.ActiveForm.TypeEx = "65211") Then
                                BubbleEvent = False
                                objapplication.Forms.Item("SELFRM").Select()
                            ElseIf bModal And (objaddon.objapplication.Forms.ActiveForm.TypeEx = "CT_PF_ManufacOrd") Then
                                BubbleEvent = False
                                objapplication.Forms.Item("DATAVIEW").Select()
                            End If
                            If objaddon.objapplication.Forms.ActiveForm.TypeEx = "CT_PF_MORSEMRESH" Then ' Semi Finished Product Scheduling
                                If pVal.ItemUID <> "Ok" Then Exit Sub
                                objform = objaddon.objapplication.Forms.GetForm("CT_PF_MORSEMRESH", objaddon.objapplication.Forms.ActiveForm.TypeCount)
                                objGrid = objform.Items.Item("9").Specific
                                For i As Integer = 0 To objGrid.Rows.Count - 1
                                    If objGrid.DataTable.GetValue("U_ResCode", i) = "" Then Continue For
                                    Dim Qty As Double = CDbl(objGrid.DataTable.GetValue("Quantity", i))
                                    strsql = objglobalmethods.getSingleValue("Select ""U_CycleCap"" ""Cycle Capacity"" from ""@CT_PF_ORSC"" where ""U_RscCode""='" & objGrid.DataTable.GetValue("U_ResCode", i) & "' ")
                                    If CDbl(strsql) <> Qty Then objaddon.objapplication.StatusBar.SetText("Planned Qty is not matching with the Resource Qty...On line: " & i + 1, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error) : BubbleEvent = False : Exit Sub
                                Next
                            End If

                            If pVal.FormTypeEx = "CT_PF_DCF" And pVal.ItemUID = "1" Then
                                If PickOrderEntry = "" Then Exit Select
                                strsql = objglobalmethods.getSingleValue("Select 1 as ""Status"" from ""@CT_PF_POR1"" Where ""DocEntry""=" & PickOrderEntry & " and ""U_ItemCode"" not in (Select ""U_ItemCode"" from ""@CT_PF_MOR3"" Where ""DocEntry""=" & GetValue & ")")
                                If strsql = "1" Then
                                    objaddon.objapplication.StatusBar.SetText("Pick Order Items is not matching with the Manufacturing Order line items...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error) : BubbleEvent = False : Exit Sub
                                End If
                            End If


                            If objform.TypeEx.ToUpper() = "41" And pVal.ItemUID = "1" Then
                                Dim mtx As SAPbouiCOM.Matrix = objform.Items.Item("3").Specific

                                For index = 1 To mtx.RowCount
                                    If String.IsNullOrEmpty(mtx.Columns.Item("10").Cells.Item(index).Specific.String) Then
                                        objaddon.objapplication.StatusBar.SetText("Kindly Fill Expiry Date", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                        BubbleEvent = False
                                        Return
                                    End If
                                    If String.IsNullOrEmpty(mtx.Columns.Item("11").Cells.Item(index).Specific.String) Then
                                        objaddon.objapplication.StatusBar.SetText("Kindly Fill Manufacture Date", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                        BubbleEvent = False
                                        Return
                                    End If
                                Next
                            End If

                        Case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE
                            Dim EventEnum As SAPbouiCOM.BoEventTypes
                            EventEnum = pVal.EventType
                            If FormUID = "SELFRM" And (EventEnum = SAPbouiCOM.BoEventTypes.et_FORM_CLOSE) And bModal Then
                                bModal = False
                            ElseIf FormUID = "DATAVIEW" And (EventEnum = SAPbouiCOM.BoEventTypes.et_FORM_CLOSE) And bModal Then
                                bModal = False
                            End If
                        Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                            Try
                                If objaddon.objapplication.Forms.ActiveForm.TypeEx = "65214" Then
                                    objform = objaddon.objapplication.Forms.GetForm("65214", objaddon.objapplication.Forms.ActiveForm.TypeCount)
                                    odbdsDetails = objform.DataSources.DBDataSources.Item("IGN1")
                                    ProdOrderEntry = odbdsDetails.GetValue("BaseEntry", 0)
                                End If
                            Catch ex As Exception
                            End Try
                        Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED
                            If objform.Title.ToUpper() = "LAB COUNTS REPORT" Then ' Query Report
                                BubbleEvent = False
                                tempform = objform 'objapplication.Forms.ActiveForm
                                CurRow = pVal.Row
                                objapplication.ActivateMenuItem("CT_PF_OQCT")
                            ElseIf objform.Title.ToUpper() = "SALES ORDER VS AR INVOICE REPORT13122022" Then ' Query Report
                                BubbleEvent = False
                                tempform = objform ' objapplication.Forms.ActiveForm
                                CurRow = pVal.Row
                                objapplication.ActivateMenuItem("CT_PF_6")
                            End If

                    End Select
                Else
                    Select Case pVal.EventType

                        Case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE
                            If pVal.FormTypeEx = "CT_PF_PickOrder" Then
                                PickOrderEntry = objform.DataSources.DBDataSources.Item("@CT_PF_POR1").GetValue("DocEntry", 0)
                                GetValue = objform.DataSources.DBDataSources.Item("@CT_PF_POR1").GetValue("U_BaseEntry", 0)
                            End If
                        Case SAPbouiCOM.BoEventTypes.et_FORM_VISIBLE
                            If objaddon.objapplication.Forms.ActiveForm.TypeEx = "CT_PF_PickReceipt" Then
                                objform = objaddon.objapplication.Forms.GetForm("CT_PF_PickReceipt", objaddon.objapplication.Forms.ActiveForm.TypeCount)
                                odbdsDetails = objform.DataSources.DBDataSources.Item("@CT_PF_PRE1")
                                GetValue = odbdsDetails.GetValue("U_BaseEntry", 0)
                                Qty = odbdsDetails.GetValue("U_PickedQty", 0)
                                'ElseIf objaddon.objapplication.Forms.ActiveForm.TypeEx = "CT_PF_MORSEMRESH" Then
                                '    objform = objaddon.objapplication.Forms.GetForm("CT_PF_MORSEMRESH", objaddon.objapplication.Forms.ActiveForm.TypeCount)
                                '    objGrid = objform.Items.Item("9").Specific
                                '    Try
                                '        objform.Freeze(True)
                                '        'For i As Integer = objGrid.Rows.Count - 1 To 0 Step -1
                                '        '    Dim ItemCode As String = objGrid.DataTable.GetValue("ItemCode", i)
                                '        '    strsql = objglobalmethods.getSingleValue("Select ""U_ASFGMOC"" ""Auto SFG MO Creation"" from OITB where ""ItmsGrpCod""=(Select ""ItmsGrpCod"" from OITM where ""ItemCode""='" & ItemCode & "')")
                                '        '    If strsql = "N" Then objGrid.DataTable.Rows.Remove(i) ': objform.DataSources.DataTables.Item("Data").SetValue("Create", i, strsql)
                                '        'Next
                                '        'For i As Integer = 0 To objGrid.Rows.Count - 1
                                '        '    objform.DataSources.DataTables.Item("Data").SetValue("Level", i, i + 1)
                                '        'Next
                                '        For i As Integer = 0 To objGrid.Rows.Count - 1
                                '            Dim ItemCode As String = objGrid.DataTable.GetValue("ItemCode", i)
                                '            strsql = objglobalmethods.getSingleValue("Select ""U_ASFGMOC"" ""Auto SFG MO Creation"" from OITB where ""ItmsGrpCod""=(Select ""ItmsGrpCod"" from OITM where ""ItemCode""='" & ItemCode & "')")
                                '            If strsql = "N" Then objform.DataSources.DataTables.Item("Data").SetValue("Create", i, strsql)
                                '        Next
                                '        objform.Freeze(False)
                                '    Catch ex As Exception
                                '        objform.Freeze(False)
                                '    End Try

                            End If
                        Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                            Try
                                'If pVal.FormTypeEx = "CT_PF_PReceiptBatchS" And pVal.ColUID = "Batch" Then 'objaddon.objapplication.Forms.ActiveForm.TypeEx = "CT_PF_PReceiptBatchS" 
                                If pVal.FormTypeEx = "CT_PF_PReceiptBatchS" And pVal.ColUID = "DocNum" Then
                                    objform = objaddon.objapplication.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)
                                    strsql = "Select T1.""U_Quantity"" ""Qty"",T1.""U_BatchNo"" ""BatchNo"",TO_Varchar(T1.""U_MfgDate"",'yyyyMMdd') ""MfgDate"","
                                    strsql += vbCrLf + "TO_Varchar(T1.""U_ExpDate"",'yyyyMMdd') ""ExpDate"",TO_Varchar(CURRENT_DATE,'yyyyMMdd') ""Admission Date"",T1.""U_Warehouse"" ""Warehouse"" "
                                    strsql += vbCrLf + "from ""@CT_PF_OMOR"" T1 where T1.""DocEntry""=" & GetValue & " and T1.""U_BatchNo"" <>''" 'odbdsDetails.GetValue("U_BaseEntry", 0)
                                    objrs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    objrs.DoQuery(strsql)
                                    If objrs.RecordCount = 0 Then objaddon.objapplication.StatusBar.SetText("No Batches Found...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error) : Exit Sub
                                    If objrs.Fields.Item("ExpDate").Value = "" Then objaddon.objapplication.StatusBar.SetText("Expiration Date is not found...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error) : Exit Sub
                                    If objrs.Fields.Item("MfgDate").Value = "" Then objaddon.objapplication.StatusBar.SetText("Mfg Date is not found...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error) : Exit Sub
                                    objform.Freeze(True)
                                    objform.DataSources.DataTables.Item("dtBatches").SetValue("Batch", 0, objrs.Fields.Item("BatchNo").Value)
                                    objform.DataSources.DataTables.Item("dtBatches").SetValue("BatchQty", 0, Qty) 'objrs.Fields.Item("Qty").Value
                                    strsql = objglobalmethods.getSingleValue("Select 1 as ""Status"" from OWHS Where ""WhsCode""='" & objrs.Fields.Item("Warehouse").Value & "' and ""BinActivat""='Y'")
                                    If strsql = "1" Then
                                        objform.DataSources.DataTables.Item("dtBatches").SetValue("BinQty", 0, objaddon.objglobalmethods.CtoD(Qty).ToString())
                                    End If
                                    objform.DataSources.DataTables.Item("dtBatches").SetValue("ExpDate", 0, objrs.Fields.Item("ExpDate").Value)
                                    objform.DataSources.DataTables.Item("dtBatches").SetValue("ManDate", 0, objrs.Fields.Item("MfgDate").Value)
                                    objform.DataSources.DataTables.Item("dtBatches").SetValue("RecDate", 0, objrs.Fields.Item("Admission Date").Value)
                                    If objform.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                    objaddon.objapplication.StatusBar.SetText("Batches Loaded Successfully...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                    objform.Freeze(False)
                                    If strsql = "1" Then
                                        Dim col As SAPbouiCOM.EditTextColumn
                                        Dim tempform As SAPbouiCOM.Form
                                        Dim objmatrix As SAPbouiCOM.Matrix
                                        objGrid = objform.Items.Item("grdBatches").Specific
                                        col = objGrid.Columns.Item("BinQty")
                                        col.PressLink(0)
                                        tempform = objaddon.objapplication.Forms.ActiveForm
                                        'tempform.Visible = False
                                        objmatrix = tempform.Items.Item("mtxBins").Specific
                                        objmatrix.Columns.Item("Alloc").Cells.Item(1).Specific.String = Qty
                                        If tempform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then tempform.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                        tempform.Close()
                                    End If
                                End If

                            Catch ex As Exception
                                objform.Freeze(False)
                                objaddon.objapplication.StatusBar.SetText("Batch Load :" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            End Try


                    End Select
                End If
                Select Case pVal.FormTypeEx
                    'Case "65211"
                    '    objProdOrder.ItemEvent(FormUID, pVal, BubbleEvent)
                    'Case "41"
                    '    objBatchSetup.ItemEvent(FormUID, pVal, BubbleEvent)
                    Case "CT_PF_ManufacOrd"
                        objManFactOrder.ItemEvent(FormUID, pVal, BubbleEvent)
                    Case "CT_PF_PickOrder"
                        objPickOrder.ItemEvent(FormUID, pVal, BubbleEvent)
                    Case "CT_PF_QCTest"
                        objQualityControlT.ItemEvent(FormUID, pVal, BubbleEvent)
                End Select
            Catch ex As Exception
                'objform.Freeze(False)
            End Try
        End Sub

#End Region

#Region "Form Data event"
        Private Sub objApplication_FormDataEvent(ByRef pVal As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean) Handles objapplication.FormDataEvent
            Try
                'If BusinessObjectInfo.BeforeAction = True Then

                'Select Case pVal.EventType
                '    Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD

                '    Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD

                'End Select
                Select Case pVal.FormTypeEx
                    Case "CT_PF_ManufacOrd"
                    Case "CT_PF_AdtBtch"
                        objManFactOrder.FormDataEvent(pVal, BubbleEvent)
                End Select
                'End If

            Catch ex As Exception
                'objAddOn.objApplication.StatusBar.SetText(ex.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            End Try
        End Sub
#End Region

#Region "Menu Event"

        Public Sub SBO_Application_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles objapplication.MenuEvent
            Try
                Select Case pVal.MenuUID
                    Case "1281", "1282", "1283", "1284", "1285", "1286", "1287", "1300", "1288", "1289", "1290", "1291", "1304", "1292", "1293", "CPYD", "AT_Show_Batches", "6005", "6913", "AT_Remove_Row"
                        objmenuevent.MenuEvent_For_StandardMenu(pVal, BubbleEvent)
                    Case "TNKMSTR", "MULCHK"
                        MenuEvent_For_FormOpening(pVal, BubbleEvent)
                        'Case "1293"
                        '    BubbleEvent = False
                    Case "519"
                        MenuEvent_For_Preview(pVal, BubbleEvent)
                End Select
            Catch ex As Exception
                'objaddon.objapplication.SetStatusBarMessage("Error in SBO_Application MenuEvent" + ex.ToString, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
            End Try
        End Sub

        Public Sub MenuEvent_For_Preview(ByRef pval As SAPbouiCOM.MenuEvent, ByRef bubbleevent As Boolean)
            Dim oform = objaddon.objapplication.Forms.ActiveForm()
            'If pval.BeforeAction Then
            '    If oform.TypeEx = "TRANOLVA" Then MenuEvent_For_PrintPreview(oform, "8f481d5cf08e494f9a83e1e46ab2299e", "txtentry") : bubbleevent = False
            '    If oform.TypeEx = "TRANOLAP" Then MenuEvent_For_PrintPreview(oform, "f15ee526ac514070a9d546cda7f94daf", "txtentry") : bubbleevent = False
            '    If oform.TypeEx = "OLSE" Then MenuEvent_For_PrintPreview(oform, "e47ed373e0cc48efb47c9773fba64fc3", "txtentry") : bubbleevent = False
            'End If
        End Sub

        Private Sub MenuEvent_For_PrintPreview(ByVal oform As SAPbouiCOM.Form, ByVal Menuid As String, ByVal Docentry_field As String)
            'Try
            '    Dim Docentry_Est As String = oform.Items.Item(Docentry_field).Specific.String
            '    If Docentry_Est = "" Then Exit Sub
            '    print_close = False
            '    objaddon.objapplication.Menus.Item(Menuid).Activate()
            '    oform = objaddon.objapplication.Forms.ActiveForm()
            '    oform.Items.Item("1000003").Specific.string = Docentry_Est
            '    oform.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            '    print_close = True
            'Catch ex As Exception
            'End Try
        End Sub

        Public Function FormExist(ByVal FormID As String) As Boolean
            Try
                FormExist = False
                For Each uid As SAPbouiCOM.Form In objaddon.objapplication.Forms
                    If uid.TypeEx = FormID Then
                        FormExist = True
                        Exit For
                    End If
                Next
                If FormExist Then
                    objaddon.objapplication.Forms.Item(FormID).Visible = True
                    objaddon.objapplication.Forms.Item(FormID).Select()
                End If
            Catch ex As Exception
                objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try

        End Function

        Public Sub MenuEvent_For_FormOpening(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
            Try
                If pVal.BeforeAction = False Then
                    'Select Case pVal.MenuUID
                    '    Case "TNKMSTR"
                    '        Dim activeform As New FrmTankMaster
                    '        activeform.Show()
                    '    Case "MULCHK"
                    '        If Not FormExist("MULCHK") Then
                    '            Dim activeform As New FrmCheckPrinting
                    '            activeform.Show()
                    '        End If

                    'End Select

                End If
            Catch ex As Exception
                'objaddon.objapplication.SetStatusBarMessage("Error in Form Opening MenuEvent" + ex.ToString, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
            End Try
        End Sub

#End Region

#Region "LayoutKeyEvent"

        Public Sub SBO_Application_LayoutKeyEvent(ByRef eventInfo As SAPbouiCOM.LayoutKeyInfo, ByRef BubbleEvent As Boolean) Handles objapplication.LayoutKeyEvent
            'Dim oForm_Layout As SAPbouiCOM.Form = Nothing
            'If SAPbouiCOM.Framework.Application.SBO_Application.Forms.ActiveForm.BusinessObject.Type = "NJT_CES" Then
            '    oForm_Layout = SAPbouiCOM.Framework.Application.SBO_Application.Forms.Item(eventInfo.FormUID)
            'End If
        End Sub

#End Region

#Region "Application Event"

        Public Sub SBO_Application_AppEvent(EventType As SAPbouiCOM.BoAppEventTypes) Handles objapplication.AppEvent
            If EventType = SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged Or EventType = SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition Or EventType = SAPbouiCOM.BoAppEventTypes.aet_ShutDown Then
                Try

                    If objcompany.Connected Then objcompany.Disconnect()
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(objcompany)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(objapplication)
                    objcompany = Nothing
                    'objapplication = Nothing
                    GC.Collect()
                    System.Windows.Forms.Application.Exit()
                    End
                Catch ex As Exception
                End Try
            End If
            'Select Case EventType
            '    Case SAPbouiCOM.BoAppEventTypes.aet_ShutDown Or SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged Or SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition
            '        Try
            '            'If objcompany.Connected Then objcompany.Disconnect()
            '            ''System.Runtime.InteropServices.Marshal.ReleaseComObject(objcompany)
            '            ''System.Runtime.InteropServices.Marshal.ReleaseComObject(objapplication)
            '            'objcompany = Nothing
            '            ''objapplication = Nothing
            '            'GC.Collect()
            '            'System.Windows.Forms.Application.Exit()
            '            'End
            '        Catch ex As Exception
            '        End Try
            '        'Case SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged
            '        '    End
            '        '    'Case SAPbouiCOM.BoAppEventTypes.aet_FontChanged
            '        '    '    End
            '        '    'Case SAPbouiCOM.BoAppEventTypes.aet_LanguageChanged
            '        '    '    End
            '        'Case SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition
            '        '    End
            'End Select
        End Sub

#End Region

#Region "Right Click Event"

        Private Sub objapplication_RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean) Handles objapplication.RightClickEvent
            Try
                Select Case objaddon.objapplication.Forms.ActiveForm.TypeEx
                    Case "MBAPSI", "PAYINIT", "PAYM", "170", "426", "FOITR", "CT_PF_ManufacOrd"
                        objrightclickevent.RightClickEvent(eventInfo, BubbleEvent)

                End Select
            Catch ex As Exception

            End Try
        End Sub

#End Region

        Public Sub objapplication_UDOEvent(ByRef udoEvent As SAPbouiCOM.UDOEvent, ByRef BubbleEvent As Boolean) Handles objapplication.UDOEvent
            Try

            Catch ex As Exception

            End Try
        End Sub

    End Class

End Namespace
