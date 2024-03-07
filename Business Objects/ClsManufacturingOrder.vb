
Imports System.Drawing
Imports System.Windows.Forms

Namespace Decimal_Rounding
    Public Class ClsManufacturingOrder
        Public Const Formtype = "CT_PF_ManufacOrd"
        Dim objform, objUDFForm As SAPbouiCOM.Form
        Dim objmatrix As SAPbouiCOM.Matrix
        Dim strSQL As String
        Dim objRs As SAPbobsCOM.Recordset
        Dim odbdsHeader, odbdsDetails As SAPbouiCOM.DBDataSource
        Dim CalcFlag As Boolean = False

        Public Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
            Try
                objform = objaddon.objapplication.Forms.Item(FormUID)
                objRs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                odbdsDetails = objform.DataSources.DBDataSources.Item("@CT_PF_MOR3")
                odbdsHeader = objform.DataSources.DBDataSources.Item("@CT_PF_OMOR")
                If pVal.BeforeAction Then
                    Select Case pVal.EventType
                        Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                            If pVal.ItemUID = "Items" And pVal.ColUID = "U_OQty" Then
                                BubbleEvent = False
                            End If
                        Case SAPbouiCOM.BoEventTypes.et_CLICK
                            If pVal.ItemUID = "btnDCP" And objform.Mode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                                If objform.Mode <> SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                    strSQL = objaddon.objglobalmethods.getSingleValue("Select 1 as ""Status"" from IGE1 Where ""U_PickType""='CT_PF_PickOrder' and ""U_PickEntry"" in (Select distinct ""DocEntry"" from ""@CT_PF_POR1"" Where ""U_BaseEntry""='" & odbdsHeader.GetValue("DocEntry", 0) & "')")
                                    If strSQL = "1" Then objaddon.objapplication.StatusBar.SetText("Decimal cannot be adjust once goods has been issued...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error) : BubbleEvent = False : Exit Sub
                                End If
                                'If objform.Items.Item("9").Specific.Selected.Value = "RL" Then objaddon.objapplication.StatusBar.SetText("Decimal cannot be adjust in released mode...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error) : BubbleEvent = False : Exit Sub
                                ''If objform.Items.Item("etDCP").Specific.String <> "" Then
                                ''    If objaddon.objapplication.MessageBox("Do you want to adjust the Planned Quantity?", 2, "Yes", "No") <> 1 Then BubbleEvent = False : Exit Sub
                                ''End If
                            End If
                            If pVal.ItemUID = "1" And (objform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                                For i As Integer = 0 To odbdsDetails.Size - 1
                                    If odbdsDetails.GetValue("U_ItemCode", i) = "" Then Continue For
                                    strSQL = objaddon.objglobalmethods.getSingleValue("Select T1.""U_MND_STDQTY"" from ""@CT_PF_OBOM"" T0 join ""@CT_PF_BOM1"" T1 On T0.""Code""=T1.""Code"" Where T0.""U_ItemCode""='" & odbdsHeader.GetValue("U_ItemCode", 0) & "' and T1.""U_ItemCode""='" & odbdsDetails.GetValue("U_ItemCode", i) & "'")
                                    If strSQL.ToUpper = "Y" Then odbdsDetails.SetValue("U_Result", i, odbdsDetails.GetValue("U_Quantity", i))
                                Next
                                objmatrix = objform.Items.Item("Items").Specific
                                objform.Freeze(True)
                                objmatrix.LoadFromDataSource()
                                objform.Freeze(False)
                                strSQL = objaddon.objglobalmethods.getSingleValue("Select ""U_AutoBatch"" ""Auto Batch"" from OITB where ""ItmsGrpCod""=(Select ""ItmsGrpCod"" from OITM where ""ItemCode""='" & odbdsHeader.GetValue("U_ItemCode", 0) & "')")
                                If strSQL = "Y" Then
                                    Dim Batch As String
                                    Dim objedit As SAPbouiCOM.EditText
                                    Dim objGrid As SAPbouiCOM.Grid
                                    objUDFForm = objaddon.objapplication.Forms.Item(objform.UDFFormUID)
                                    If objUDFForm.Items.Item("U_BatchNo").Specific.String <> "" Then Exit Sub
                                    objedit = objUDFForm.Items.Item("U_MfgDate").Specific
                                    If objedit.Value = "" Then
                                        objaddon.objapplication.StatusBar.SetText("Manufacturing Date is Missing...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error) : Exit Sub
                                    End If
                                    Dim DocDate As Date = Date.ParseExact(objedit.Value, "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                                    Batch = DocDate.ToString("yyMMdd")
                                    strSQL = objaddon.objglobalmethods.getSingleValue("Select ""U_StrainCode"" from OITM where ""ItemCode""='" & odbdsHeader.GetValue("U_ItemCode", 0) & "'")
                                    If strSQL = "" Then
                                        objaddon.objapplication.StatusBar.SetText("Strain Code is Missing...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error) : Exit Sub
                                    End If
                                    Batch += strSQL
                                    objGrid = objform.Items.Item("OprRscGrd").Specific
                                    If objGrid.DataTable.Rows.Count > 0 Then strSQL = objGrid.DataTable.Columns.Item("U_RscCode").Cells.Item(0).Value Else strSQL = ""
                                    If strSQL = "" Then
                                        objaddon.objapplication.StatusBar.SetText("Resource Code is Missing...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error) : Exit Sub
                                    End If
                                    Batch = Batch + "-" + strSQL
                                    objUDFForm.Items.Item("U_BatchNo").Specific.String = Batch
                                    objform.Items.Item("11").Click()
                                    objaddon.objapplication.StatusBar.SetText("Auto-Batch assigned successfully...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                End If
                            End If

                    End Select
                Else
                    Select Case pVal.EventType
                        Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                            If pVal.ItemUID = "1" And pVal.ActionSuccess = True And objform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                objform.Items.Item("etDCP").Specific.String = CInt(1)
                                objform.Items.Item("11").Click()
                            End If
                        Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                            If pVal.ActionSuccess Then
                                CreateButton(FormUID)
                            End If
                        Case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE
                            CalcFlag = False
                        Case SAPbouiCOM.BoEventTypes.et_VALIDATE
                            If pVal.ItemUID = "11" And objform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE And pVal.ItemChanged = True Then
                                If CalcFlag = True Then CalcFlag = False
                            ElseIf pVal.ItemUID = "Quantity" And (objform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) And pVal.ItemChanged = True Then
                                If CalcFlag = True Then CalcFlag = False
                            End If
                        Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                            If (pVal.ItemUID = "11" Or pVal.ItemUID = "Quantity") And (objform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                                If CalcFlag = True Then Exit Sub
                                'If Val(objform.Items.Item("etDCP").Specific.String) <> 0 Then Exit Sub
                                objmatrix = objform.Items.Item("Items").Specific

                                Dim Qty As Double
                                For i As Integer = 0 To odbdsDetails.Size - 1
                                    If odbdsDetails.GetValue("U_ItemCode", i) = "" Then Continue For
                                    odbdsDetails.SetValue("U_OQty", i, odbdsDetails.GetValue("U_Quantity", i) * odbdsHeader.GetValue("U_Quantity", 0))
                                    Qty = CDbl(odbdsDetails.GetValue("U_Result", i))
                                    CalcFlag = True
                                Next
                                objform.Freeze(True)
                                objmatrix.LoadFromDataSource()
                                objform.Freeze(False)
                            End If
                            CreateButton(FormUID)
                        Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                            If pVal.ItemUID = "txtTANK" And pVal.CharPressed = 9 And pVal.Modifiers = SAPbouiCOM.BoModifiersEnum.mt_None And objform.Mode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then

                            End If
                        Case SAPbouiCOM.BoEventTypes.et_CLICK
                            If pVal.ItemUID = "btnDCP" And (objform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                                'If objform.Items.Item("9").Specific.Selected.Value = "RL" Then Exit Sub
                                If objform.Items.Item("etDCP").Specific.String <> "" Then
                                    Rounding_PlannedQty(FormUID, CInt(objform.Items.Item("etDCP").Specific.String))
                                End If
                                'ElseIf pVal.ItemUID = "1" And (objform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then

                            End If
                        Case SAPbouiCOM.BoEventTypes.et_FORM_VISIBLE
                            If clsModule.objaddon.objapplication.Menus.Item("6913").Checked = False Then
                                clsModule.objaddon.objapplication.SendKeys("^+U")
                            End If
                    End Select
                End If
            Catch ex As Exception
                objform.Freeze(False)
            End Try
        End Sub

        Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
            Try
                objform = objaddon.objapplication.Forms.Item(BusinessObjectInfo.FormUID)
                'objRs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                If BusinessObjectInfo.BeforeAction = True Then
                    Select Case BusinessObjectInfo.EventType
                        Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                        Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD

                    End Select
                Else
                    Select Case BusinessObjectInfo.EventType
                        Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                        Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                            Try
                                CreateButton(BusinessObjectInfo.FormUID)
                                strSQL = objaddon.objglobalmethods.getSingleValue("Select 1 as ""Status"" from IGE1 Where ""U_PickType""='CT_PF_PickOrder' and ""U_PickEntry"" in (Select distinct ""DocEntry"" from ""@CT_PF_POR1"" Where ""U_BaseEntry""='" & odbdsHeader.GetValue("DocEntry", 0) & "')")
                                If strSQL = "1" Then objform.Items.Item("btnDCP").Enabled = False
                            Catch ex As Exception

                            End Try

                            If BusinessObjectInfo.FormTypeEx = "CT_PF_AdtBtch" Then
                                Dim obj As SAPbouiCOM.Form = objaddon.objapplication.Forms.Item(BusinessObjectInfo.FormUID)
                                Dim Btdoc As String = obj.DataSources.DBDataSources.Item("@CT_PF_OABT").GetValue("DocEntry", 0)
                                Dim Batchno As String = obj.DataSources.DBDataSources.Item("@CT_PF_OABT").GetValue("U_DistNumber", 0)
                                Dim U_BADETAIL As String = obj.DataSources.DBDataSources.Item("@CT_PF_OABT").GetValue("U_BADETAIL", 0)

                                If (U_BADETAIL Is Nothing Or U_BADETAIL.Trim = "") Then
                                    Dim lstrquery As String = "SELECT MAx(""U_BachDetls"") as ""U_BachDetls""  FROM ""@CT_PF_OMOR"" cpo WHERE ""U_BatchNo""  ='" + Batchno + "'"
                                    Dim updateval As String = clsModule.objaddon.objglobalmethods.getSingleValue(lstrquery)
                                    If (Not updateval.Trim = "") Then
                                        obj.DataSources.DBDataSources.Item("@CT_PF_OABT").SetValue("U_BADETAIL", 0, updateval)
                                        obj.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                        obj.Items.Item("1").Click()
                                    End If
                                End If

                            End If
                            'If objform.Items.Item("9").Specific.Selected.Value = "RL" Then objform.Items.Item("btnDCP").Enabled = False
                    End Select
                End If


            Catch ex As Exception
                objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try

        End Sub

        Public Sub CreateButton(ByVal FormUID As String)
            Try
                Dim objItem As SAPbouiCOM.Item
                'Dim objLabel As SAPbouiCOM.StaticText
                Dim objButton As SAPbouiCOM.Button
                Dim objedit As SAPbouiCOM.EditText
                'Dim objlink As SAPbouiCOM.LinkedButton
                objform = objaddon.objapplication.Forms.Item(FormUID)
                If objform.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then Exit Sub
                Try
                    If objform.Items.Item("btnDCP").UniqueID = "btnDCP" Then Exit Sub
                Catch ex As Exception
                End Try

                objItem = objform.Items.Add("btnDCP", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
                objItem.Left = objform.Items.Item("17").Left + objform.Items.Item("17").Width + 10
                'objItem.Left = objForm.Items.Item("10002056").Left + objForm.Items.Item("10002056").Width + 60
                Dim Fieldsize As Size = TextRenderer.MeasureText("Adjust Decimal", New Font("Arial", 12.0F))
                objItem.Width = Fieldsize.Width '120
                objItem.Top = objform.Items.Item("17").Top
                objItem.Height = 19 ' objform.Items.Item("6").Height
                objItem.LinkTo = "17"
                objButton = objItem.Specific
                objButton.Caption = "Adjust Decimal"

                'Dim objedit As SAPbouiCOM.EditText
                objItem = objform.Items.Add("etDCP", SAPbouiCOM.BoFormItemTypes.it_EDIT)
                objItem.Left = objform.Items.Item("btnDCP").Left + objform.Items.Item("btnDCP").Width + 5
                objItem.Width = 50
                objItem.Top = objform.Items.Item("17").Top
                objItem.Height = objform.Items.Item("17").Height ' 14 'objform.Items.Item("btnDCP").Height
                objItem.LinkTo = "btnDCP"
                objedit = objItem.Specific
                'objedit.Item.Enabled = False
                objedit.DataBind.SetBound(True, "@CT_PF_OMOR", "U_AdjDec")
                Try
                    objedit.Value = CInt(1)
                    objform.Items.Item("11").Click()
                Catch ex As Exception

                End Try

            Catch ex As Exception
            End Try

        End Sub

        Private Function Rounding_PlannedQty(ByVal FormUID As String, ByVal Rounddec As Integer)
            Try
                objform = objaddon.objapplication.Forms.Item(FormUID)
                objmatrix = objform.Items.Item("Items").Specific
                odbdsDetails = objform.DataSources.DBDataSources.Item("@CT_PF_MOR3")
                odbdsHeader = objform.DataSources.DBDataSources.Item("@CT_PF_OMOR")
                Dim PlanQty As Double
                objform.Freeze(True)
                For i As Integer = 0 To odbdsDetails.Size - 1
                    If odbdsDetails.GetValue("U_ItemCode", i) = "" Then Continue For
                    odbdsDetails.SetValue("U_OQty", i, odbdsDetails.GetValue("U_Quantity", i) * odbdsHeader.GetValue("U_Quantity", 0))
                    PlanQty = CDbl(odbdsDetails.GetValue("U_OQty", i))
                    If Math.Round(PlanQty, Rounddec) = 0 Then
                        odbdsDetails.SetValue("U_Result", i, 0.1)
                    Else
                        odbdsDetails.SetValue("U_Result", i, Math.Round(CDbl(odbdsDetails.GetValue("U_OQty", i)), Rounddec)) 'If Math.Round(PlanQty, Rounddec) <> 0.1 Then 
                    End If
                Next
                objmatrix.LoadFromDataSource()
                objaddon.objapplication.StatusBar.SetText("Planned Qty adjusted successfully!!!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                objmatrix.Columns.Item("col_1").Cells.Item(1).Click()
                objform.Freeze(False)
            Catch ex As Exception
                objform.Freeze(False)
            End Try
        End Function

    End Class

End Namespace
