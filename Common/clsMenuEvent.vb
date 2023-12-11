Imports System.Windows.Forms
Imports SAPbobsCOM
Imports SAPbouiCOM.Framework

Namespace Decimal_Rounding

    Public Class clsMenuEvent
        Dim objform As SAPbouiCOM.Form
        Dim objglobalmethods As New clsGlobalMethods
        Public ItemCode, StrSQL As String

        Public Sub MenuEvent_For_StandardMenu(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
            Try
                objform = objaddon.objapplication.Forms.ActiveForm

                Select Case objaddon.objapplication.Forms.ActiveForm.TypeEx

                    Case "CT_PF_PickOrder"
                        If pVal.BeforeAction = False And pVal.MenuUID = "1281" Then
                            objform.Items.Item("btnqty").Enabled = False
                        End If

                    Case "CT_PF_ManufacOrd"
                        If pVal.BeforeAction = False And pVal.MenuUID = "1281" Then
                            objform.Items.Item("btnDCP").Enabled = False
                        End If

                        If pVal.MenuUID = "AT_Show_Batches" Then
                            If Not objaddon.FormExist("DataView") Then
                                Dim TranViewForm As New FrmDataView
                                TranViewForm.Show()
                                TranViewForm.objform.Title = "Available Batch/Serial"
                                TranViewForm.objform.Left = objform.Left + 100
                                TranViewForm.objform.Top = objform.Top + 100
                                TranViewForm.GetData(objform.DataSources.DBDataSources.Item("@CT_PF_OMOR").GetValue("DocEntry", 0), IIf(objaddon.objmenuevent.ItemCode = "", "", objaddon.objmenuevent.ItemCode), "4")
                            End If
                        ElseIf pVal.MenuUID = "AT_Remove_Row" Then 'CT_PF_ManufacOrd|Items|del "1293"
                            'StrSQL = "Select distinct 1 as ""Status"" from ""@CT_PF_POR1"" Where ""U_BaseEntry""='" & objform.DataSources.DBDataSources.Item("@CT_PF_OMOR").GetValue("DocEntry", 0) & "' and ""U_ItemCode"" in ('" & objaddon.objmenuevent.ItemCode & "')" 'Pick Order
                            StrSQL = "Select distinct 1 as ""Status"" from IGE1 Where ""U_DocEntry""='" & objform.DataSources.DBDataSources.Item("@CT_PF_OMOR").GetValue("DocEntry", 0) & "' and ""ItemCode"" in ('" & objaddon.objmenuevent.ItemCode & "')" 'Goods Issue
                            StrSQL = objaddon.objglobalmethods.getSingleValue(StrSQL)
                            If StrSQL = "" Then
                                If pVal.BeforeAction = True Then Exit Sub
                                Try
                                    Dim objMatrix As SAPbouiCOM.Matrix
                                    Dim DBSource As SAPbouiCOM.DBDataSource
                                    objMatrix = objform.Items.Item("Items").Specific
                                    objMatrix.FlushToDataSource()
                                    objMatrix.DeleteRow(curMORow)
                                    DBSource = objform.DataSources.DBDataSources.Item("@CT_PF_MOR3")
                                    'DeleteRow(objMatrix, "@CT_PF_MOR3")
                                    DBSource.RemoveRecord(curMORow - 1) 'DBSource.Size - 1
                                    objMatrix.LoadFromDataSource()
                                    If objform.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                Catch ex As Exception
                                    objaddon.objapplication.StatusBar.SetText("Delete Row  Method Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_None)
                                Finally
                                End Try
                            Else
                                If pVal.BeforeAction = False Then Exit Sub
                                objaddon.objapplication.StatusBar.SetText("You cannot delete the row of the Item """ & objaddon.objmenuevent.ItemCode & """ issued in Pick Order", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                BubbleEvent = False
                            End If
                        End If
                        'Default_Sample_MenuEvent(pVal, BubbleEvent)
                        'Case "65211"
                        '    Production_Order_MenuEvent(pVal, BubbleEvent)
                End Select
            Catch ex As Exception

            End Try
        End Sub

        Private Sub Default_Sample_MenuEvent(ByVal pval As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
            Try
                objform = objaddon.objapplication.Forms.ActiveForm
                Dim oUDFForm As SAPbouiCOM.Form
                If pval.BeforeAction = True Then
                    Select Case pval.MenuUID
                        Case "6005"

                        Case "6913"

                        Case "1284" 'Cancel

                    End Select
                Else
                    oUDFForm = objaddon.objapplication.Forms.Item(objform.UDFFormUID)
                    Select Case pval.MenuUID
                        Case "1284" 'Cancel

                        Case "1281" 'Find

                        Case "1287" 'Duplicate



                        Case "1282"


                        Case Else

                    End Select
                End If
            Catch ex As Exception
                'objaddon.objapplication.SetStatusBarMessage("Error in Standart Menu Event" + ex.ToString, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
            End Try
        End Sub

#Region "Production Order"

        Private Sub Production_Order_MenuEvent(ByRef pval As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
            Dim Matrix0 As SAPbouiCOM.Matrix
            Try
                objform = objaddon.objapplication.Forms.ActiveForm
                'Matrix0 = objform.Items.Item("mtxcont").Specific
                If pval.BeforeAction = True Then
                    Select Case pval.MenuUID
                        Case "1283", "1284" 'Remove & Cancel
                            'objaddon.objapplication.SetStatusBarMessage("Remove or Cancel is not allowed ", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            'BubbleEvent = False
                        Case "1293"  'Delete Row
                    End Select
                Else
                    'Dim DBSource As SAPbouiCOM.DBDataSource
                    'DBSource = objform.DataSources.DBDataSources.Item("@MIPL_OAPI")
                    Select Case pval.MenuUID
                        Case "1281" 'Find Mode
                            objform.Items.Item("btnDCP").Enabled = False

                        Case "1282" ' Add Mode
                            objform.Items.Item("btnDCP").Enabled = True


                        Case "1288", "1289", "1290", "1291"

                        Case "1293"
                            'DeleteRow(Matrix0, "@MIPL_API1")
                        Case "1292"
                            'objaddon.objglobalmethods.Matrix_Addrow(Matrix0, "vcode", "#")
                        Case "1304" 'Refresh
                    End Select
                End If
            Catch ex As Exception
                objform.Freeze(False)
                ' objaddon.objapplication.SetStatusBarMessage("Error in Menu Event" + ex.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, True)
            End Try
        End Sub

#End Region

        Sub DeleteRow(ByVal objMatrix As SAPbouiCOM.Matrix, ByVal TableName As String)
            Try
                Dim DBSource As SAPbouiCOM.DBDataSource
                'objMatrix = objform.Items.Item("20").Specific
                objMatrix.FlushToDataSource()
                DBSource = objform.DataSources.DBDataSources.Item(TableName) '"@MIREJDET1"
                For i As Integer = 1 To objMatrix.VisualRowCount
                    objMatrix.GetLineData(i)
                    DBSource.Offset = i - 1
                    DBSource.SetValue("LineId", DBSource.Offset, i)
                    objMatrix.SetLineData(i)
                    objMatrix.FlushToDataSource()
                Next
                DBSource.RemoveRecord(DBSource.Size - 1)
                objMatrix.LoadFromDataSource()

            Catch ex As Exception
                objaddon.objapplication.StatusBar.SetText("DeleteRow  Method Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_None)
            Finally
            End Try
        End Sub

    End Class
End Namespace