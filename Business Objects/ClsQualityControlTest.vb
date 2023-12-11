
Namespace Decimal_Rounding

    Public Class ClsQualityControlTest
        Public Const Formtype = "CT_PF_QCTest"
        Dim objform As SAPbouiCOM.Form
        Dim objmatrix As SAPbouiCOM.Matrix
        Dim strSQL As String
        Dim objRs As SAPbobsCOM.Recordset
        Dim odbdsDetails As SAPbouiCOM.DBDataSource

        'S.No	Screen	Functionality
        '1   Manufacturing Order	Rounding the decimal points in MO Screen until the goods issue Is posted
        '2   Pick Order	Picked Qty loading from MO Line Qty after clicked the Pick Qty Button
        '3   Quality Control Test	Total Billion FMS replaced
        '4   Semi Finished Product Scheduling	"1. ""Planned Qty is not matching with the Resource Qty"" Validation added.
        '2. Create check box value assigning based on item groups udf config"
        '5   Pick Receipt Batch Screen	Batch details Loading

        Public Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
            Try
                objform = objaddon.objapplication.Forms.Item(FormUID)
                'objRs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                If pVal.BeforeAction Then
                    Select Case pVal.EventType
                        Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                        Case SAPbouiCOM.BoEventTypes.et_CLICK
                        Case SAPbouiCOM.BoEventTypes.et_VALIDATE


                    End Select
                Else
                    Select Case pVal.EventType
                        Case SAPbouiCOM.BoEventTypes.et_VALIDATE
                            If pVal.InnerEvent = True Or pVal.ItemChanged = False Then Exit Sub
                            If (pVal.ColUID = "U_101" Or pVal.ColUID = "U_102" Or pVal.ColUID = "U_103" Or pVal.ColUID = "U_91" Or pVal.ColUID = "U_92" Or pVal.ColUID = "U_93" Or pVal.ColUID = "U_81" Or pVal.ColUID = "U_82" Or pVal.ColUID = "U_83") Then
                                objmatrix = objform.Items.Item("QctTsRMtx").Specific
                                odbdsDetails = objform.DataSources.DBDataSources.Item("@CT_PF_QCT1")
                                'odbdsDetails.SetValue("U_TotalInBillion", pVal.Row - 1, Get_Total(FormUID, pVal.Row - 1))
                                'objmatrix.Columns.Item("U_TotalInB").Cells.Item(pVal.Row).Specific.String = Get_Total(FormUID, pVal.Row)
                                objmatrix.SetCellWithoutValidation(pVal.Row, "U_TotalInB", Get_Total(FormUID, pVal.Row))
                            End If

                        Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                        Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                        Case SAPbouiCOM.BoEventTypes.et_CLICK
                        Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                    End Select
                End If
            Catch ex As Exception
                objform.Freeze(False)
            End Try
        End Sub

        Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
            Try
                objform = objaddon.objapplication.Forms.Item(BusinessObjectInfo.FormUID)
                objRs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                If BusinessObjectInfo.BeforeAction = True Then
                    Select Case BusinessObjectInfo.EventType
                        Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD

                        Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD

                    End Select
                Else
                    Select Case BusinessObjectInfo.EventType
                        Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                        Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                    End Select
                End If

            Catch ex As Exception
                objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try

        End Sub

        Private Function Get_Total(ByVal FormUID As String, ByVal Row As Integer)
            Try
                Dim GetValue As String
                Dim A, B, C, V As Double

                objform = objaddon.objapplication.Forms.Item(FormUID)
                objmatrix = objform.Items.Item("QctTsRMtx").Specific
                odbdsDetails = objform.DataSources.DBDataSources.Item("@CT_PF_QCT1")
                'GetValue = odbdsDetails.GetValue("U_101", Row)
                'A = IIf(odbdsDetails.GetValue("U_101", Row) >= 25, odbdsDetails.GetValue("U_101", Row), 0) + IIf(odbdsDetails.GetValue("U_102", Row) >= 25, odbdsDetails.GetValue("U_102", Row), 0) + IIf(odbdsDetails.GetValue("U_103", Row) >= 25, odbdsDetails.GetValue("U_103", Row), 0)
                'B = IIf(odbdsDetails.GetValue("U_91", Row) >= 25, odbdsDetails.GetValue("U_91", Row), 0) + IIf(odbdsDetails.GetValue("U_92", Row) >= 25, odbdsDetails.GetValue("U_92", Row), 0) + IIf(odbdsDetails.GetValue("U_93", Row) >= 25, odbdsDetails.GetValue("U_93", Row), 0)
                'C = IIf(odbdsDetails.GetValue("U_81", Row) >= 25, odbdsDetails.GetValue("U_81", Row), 0) + IIf(odbdsDetails.GetValue("U_82", Row) >= 25, odbdsDetails.GetValue("U_82", Row), 0) + IIf(odbdsDetails.GetValue("U_83", Row) >= 25, odbdsDetails.GetValue("U_83", Row), 0)

                'V = IIf(odbdsDetails.GetValue("U_101", Row) >= 25, 1, 0) + IIf(odbdsDetails.GetValue("U_102", Row) >= 25, 1, 0) + IIf(odbdsDetails.GetValue("U_103", Row) >= 25, 1, 0)
                'V += IIf(odbdsDetails.GetValue("U_91", Row) >= 25, 1, 0) + IIf(odbdsDetails.GetValue("U_92", Row) >= 25, 1, 0) + IIf(odbdsDetails.GetValue("U_93", Row) >= 25, 1, 0)
                'V += IIf(odbdsDetails.GetValue("U_81", Row) >= 25, 1, 0) + IIf(odbdsDetails.GetValue("U_82", Row) >= 25, 1, 0) + IIf(odbdsDetails.GetValue("U_83", Row) >= 25, 1, 0)

                A = IIf(Val(objmatrix.Columns.Item("U_101").Cells.Item(Row).Specific.String) >= 25 And Val(objmatrix.Columns.Item("U_101").Cells.Item(Row).Specific.String) <= 350, Val(objmatrix.Columns.Item("U_101").Cells.Item(Row).Specific.String), 0)
                A += IIf(Val(objmatrix.Columns.Item("U_102").Cells.Item(Row).Specific.String) >= 25 And Val(objmatrix.Columns.Item("U_102").Cells.Item(Row).Specific.String) <= 350, Val(objmatrix.Columns.Item("U_102").Cells.Item(Row).Specific.String), 0)
                A += IIf(Val(objmatrix.Columns.Item("U_103").Cells.Item(Row).Specific.String) >= 25 And Val(objmatrix.Columns.Item("U_103").Cells.Item(Row).Specific.String) <= 350, Val(objmatrix.Columns.Item("U_103").Cells.Item(Row).Specific.String), 0)

                B = IIf(Val(objmatrix.Columns.Item("U_91").Cells.Item(Row).Specific.String) >= 25 And Val(objmatrix.Columns.Item("U_91").Cells.Item(Row).Specific.String) <= 350, Val(objmatrix.Columns.Item("U_91").Cells.Item(Row).Specific.String), 0)
                B += IIf(Val(objmatrix.Columns.Item("U_92").Cells.Item(Row).Specific.String) >= 25 And Val(objmatrix.Columns.Item("U_92").Cells.Item(Row).Specific.String) <= 350, Val(objmatrix.Columns.Item("U_92").Cells.Item(Row).Specific.String), 0)
                B += IIf(Val(objmatrix.Columns.Item("U_93").Cells.Item(Row).Specific.String) >= 25 And Val(objmatrix.Columns.Item("U_93").Cells.Item(Row).Specific.String) <= 350, Val(objmatrix.Columns.Item("U_93").Cells.Item(Row).Specific.String), 0)

                C = IIf(Val(objmatrix.Columns.Item("U_81").Cells.Item(Row).Specific.String) >= 25 And Val(objmatrix.Columns.Item("U_81").Cells.Item(Row).Specific.String) <= 350, Val(objmatrix.Columns.Item("U_81").Cells.Item(Row).Specific.String), 0)
                C += IIf(Val(objmatrix.Columns.Item("U_82").Cells.Item(Row).Specific.String) >= 25 And Val(objmatrix.Columns.Item("U_82").Cells.Item(Row).Specific.String) <= 350, Val(objmatrix.Columns.Item("U_82").Cells.Item(Row).Specific.String), 0)
                C += IIf(Val(objmatrix.Columns.Item("U_83").Cells.Item(Row).Specific.String) >= 25 And Val(objmatrix.Columns.Item("U_83").Cells.Item(Row).Specific.String) <= 350, Val(objmatrix.Columns.Item("U_83").Cells.Item(Row).Specific.String), 0)

                V = IIf(Val(objmatrix.Columns.Item("U_101").Cells.Item(Row).Specific.String) >= 25 And Val(objmatrix.Columns.Item("U_101").Cells.Item(Row).Specific.String) <= 350, 1, 0) + IIf(Val(objmatrix.Columns.Item("U_102").Cells.Item(Row).Specific.String) >= 25 And Val(objmatrix.Columns.Item("U_102").Cells.Item(Row).Specific.String) <= 350, 1, 0) + IIf(Val(objmatrix.Columns.Item("U_103").Cells.Item(Row).Specific.String) >= 25 And Val(objmatrix.Columns.Item("U_103").Cells.Item(Row).Specific.String) <= 350, 1, 0)
                V += IIf(Val(objmatrix.Columns.Item("U_91").Cells.Item(Row).Specific.String) >= 25 And Val(objmatrix.Columns.Item("U_91").Cells.Item(Row).Specific.String) <= 350, 1, 0) + IIf(Val(objmatrix.Columns.Item("U_92").Cells.Item(Row).Specific.String) >= 25 And Val(objmatrix.Columns.Item("U_92").Cells.Item(Row).Specific.String) <= 350, 1, 0) + IIf(Val(objmatrix.Columns.Item("U_93").Cells.Item(Row).Specific.String) >= 25 And Val(objmatrix.Columns.Item("U_93").Cells.Item(Row).Specific.String) <= 350, 1, 0)
                V += IIf(Val(objmatrix.Columns.Item("U_81").Cells.Item(Row).Specific.String) >= 25 And Val(objmatrix.Columns.Item("U_81").Cells.Item(Row).Specific.String) <= 350, 1, 0) + IIf(Val(objmatrix.Columns.Item("U_82").Cells.Item(Row).Specific.String) >= 25 And Val(objmatrix.Columns.Item("U_82").Cells.Item(Row).Specific.String) <= 350, 1, 0) + IIf(Val(objmatrix.Columns.Item("U_83").Cells.Item(Row).Specific.String) >= 25 And Val(objmatrix.Columns.Item("U_83").Cells.Item(Row).Specific.String) <= 350, 1, 0)

                GetValue = CDbl(((A * 10) + B + (C / 10)) / V)
                Return GetValue
            Catch ex As Exception

            End Try
        End Function

    End Class
End Namespace