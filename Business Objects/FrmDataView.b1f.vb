Option Strict Off
Option Explicit On

Imports SAPbouiCOM.Framework

Namespace Decimal_Rounding
    <FormAttribute("DATAVIEW", "Business Objects/FrmDataView.b1f")>
    Friend Class FrmDataView
        Inherits UserFormBase
        Public WithEvents objform As SAPbouiCOM.Form
        Public Sub New()
        End Sub

        Public Overrides Sub OnInitializeComponent()
            Me.Button0 = CType(Me.GetItem("101").Specific, SAPbouiCOM.Button)
            Me.Button1 = CType(Me.GetItem("2").Specific, SAPbouiCOM.Button)
            Me.Grid0 = CType(Me.GetItem("griddata").Specific, SAPbouiCOM.Grid)
            Me.Button2 = CType(Me.GetItem("bexpand").Specific, SAPbouiCOM.Button)
            Me.Button3 = CType(Me.GetItem("bcollap").Specific, SAPbouiCOM.Button)
            Me.OnCustomInitialize()

        End Sub

        Public Overrides Sub OnInitializeFormEvents()
            AddHandler CloseAfter, AddressOf Me.Form_CloseAfter
            AddHandler ResizeAfter, AddressOf Me.Form_ResizeAfter

        End Sub

#Region "Fields"

        Private WithEvents Button0 As SAPbouiCOM.Button
        Private WithEvents Button1 As SAPbouiCOM.Button
        Private WithEvents Grid0 As SAPbouiCOM.Grid
        Private WithEvents Button2 As SAPbouiCOM.Button
        Private WithEvents Button3 As SAPbouiCOM.Button

#End Region

        Private Sub OnCustomInitialize()
            Try
                objform = objaddon.objapplication.Forms.GetForm("DATAVIEW", 1)
                objform = objaddon.objapplication.Forms.ActiveForm
                bModal = True
                'objform.Title = "Available Batches/Serial"
            Catch ex As Exception

            End Try
        End Sub



        Private Sub Button0_ClickAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Button0.ClickAfter
            Try
                objform.Close()
            Catch ex As Exception

            End Try

        End Sub

        Public Sub GetData(ByVal DocEntry As String, ByVal ItemCode As String, ByVal LinkedID As String)
            Dim objrs As SAPbobsCOM.Recordset
            Try
                Dim str_sql As String = ""
                If objaddon.HANA Then
                    str_sql = "Select T0.""U_ItemCode"" ""Item No."",T0.""U_Description"" ""Item Description"",T0.""U_WhsCode"" ""Whse Code"",T0.""U_Result"" ""Total Needed Quantity"",A.""BatchSerial"" ""Batch"",ifnull(A.""Qty"",0) ""Available Quantity"""
                    str_sql += vbCrLf + ",B.""U_Status"" ""Status"",B.""U_MnfSerial"" ""Strain Code ID"",B.""U_MND_ASPT"" ""Assay Pot Perc"" from ""@CT_PF_MOR3"" T0 left Join (Select T0.""ItemCode"",T0.""WhsCode"",T0.""BatchNum"" ""BatchSerial"",T0.""Quantity"" ""Qty"""
                    str_sql += vbCrLf + "from OBTN T2 left join OIBT T0 on T2.""DistNumber""=T0.""BatchNum"" and T0.""ItemCode""=T2.""ItemCode"" where T0.""Direction""=0 and T0.""Quantity"">0"
                    str_sql += vbCrLf + "Union All"
                    str_sql += vbCrLf + "SELECT T4.""ItemCode"",T4.""WhsCode"",T4.""IntrSerial"" ""BatchSerial"", T4.""Quantity"""
                    str_sql += vbCrLf + "from SRI1 I1 join OSRI T4 on T4.""ItemCode""=I1.""ItemCode"" and I1.""SysSerial""=T4.""SysSerial"" and I1.""WhsCode"" = T4.""WhsCode"" where T4.""Quantity"">0 and I1.""Direction""=0 "
                    str_sql += vbCrLf + ") A On T0.""U_ItemCode""= A.""ItemCode"" and T0.""U_WhsCode""=A.""WhsCode"" left join ""@CT_PF_OABT"" B On A.""BatchSerial""=B.""U_DistNumber"" and A.""ItemCode""=B.""U_ItemCode"" "
                    str_sql += vbCrLf + "Where T0.""DocEntry""=" & DocEntry & ""
                    If ItemCode <> "" Then str_sql += vbCrLf + "And T0.""U_ItemCode""= '" & ItemCode & "'"
                Else
                    str_sql = "Select T0.U_ItemCode Item No.,T0.U_Description Item Description,T0.U_WhsCode Whse Code,T0.U_Result Total Needed Quantity,A.BatchSerial Batch,ifnull(A.Qty,0) Available Quantity"
                    str_sql += vbCrLf + "from [@CT_PF_MOR3] T0 left Join (Select T0.ItemCode,T0.WhsCode,T0.BatchNum BatchSerial,T0.Quantity Qty"
                    str_sql += vbCrLf + "from OBTN T2 left join OIBT T0 on T2.DistNumber=T0.BatchNum and T0.ItemCode=T2.ItemCode where T0.Direction=0 and T0.Quantity>0"
                    str_sql += vbCrLf + "Union All"
                    str_sql += vbCrLf + "SELECT T4.ItemCode,T4.WhsCode,T4.IntrSerial BatchSerial, T4.Quantity"
                    str_sql += vbCrLf + "from SRI1 I1 join OSRI T4 on T4.ItemCode=I1.ItemCode and I1.SysSerial=T4.SysSerial and I1.WhsCode = T4.WhsCode where T4.Quantity>0 and I1.Direction=0 "
                    str_sql += vbCrLf + ") A On T0.U_ItemCode= A.ItemCode and T0.U_WhsCode=A.WhsCode Where T0.DocEntry= " & DocEntry & ""
                    If ItemCode <> "" Then str_sql += vbCrLf + "and T0.U_ItemCode= '" & ItemCode & "'"
                End If
                objrs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                objrs.DoQuery(str_sql)
                If objrs.RecordCount = 0 Then objaddon.objapplication.StatusBar.SetText("No Entries Found...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning) : objform.Close() : objrs = Nothing : Exit Sub
                Dim objDT As SAPbouiCOM.DataTable
                If objform.DataSources.DataTables.Count = 0 Then
                    objform.DataSources.DataTables.Add("DT_0")
                End If

                objDT = objform.DataSources.DataTables.Item("DT_0")
                objDT.ExecuteQuery(str_sql)
                objform.DataSources.DataTables.Item("DT_0").ExecuteQuery(str_sql)

                Grid0.DataTable = objform.DataSources.DataTables.Item("DT_0")
                Grid0.CollapseLevel = 1
                objform.Freeze(True)
                For i As Integer = 0 To Grid0.Columns.Count - 1
                    'Grid0.Columns.Item(i).TitleObject.Sortable = True
                    Grid0.Columns.Item(i).Editable = False
                Next
                If ItemCode = "" Then Grid0.Rows.CollapseAll()
                Grid0.AutoResizeColumns()
                objform.Freeze(False)
                Grid0.Columns.Item(0).Type = SAPbouiCOM.BoGridColumnType.gct_EditText
                Dim col As SAPbouiCOM.EditTextColumn
                col = Grid0.Columns.Item(0)
                col.LinkedObjectType = LinkedID
                objform.Visible = True
                objform.Update()

            Catch ex As Exception
                objform.Freeze(False)
            End Try
        End Sub

        Private Sub Button2_ClickAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Button2.ClickAfter
            Try
                'If Grid0.Rows.IsLeaf(1) Then
                '    Grid0.DataTable.Rows.R
                '    objform.Update()
                'End If
                Grid0.Rows.ExpandAll()
                objaddon.objapplication.Menus.Item("1300").Activate()
            Catch ex As Exception

            End Try

        End Sub

        Private Sub Button3_ClickAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Button3.ClickAfter
            Try
                Grid0.Rows.CollapseAll()
                objaddon.objapplication.Menus.Item("1300").Activate()
            Catch ex As Exception

            End Try

        End Sub

        Private Sub Grid0_ClickAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Grid0.ClickAfter
            Try
                If pVal.Row <> -1 Then Grid0.Rows.SelectedRows.Add(pVal.Row)
                'Dim clrRow As Integer = -1, ParentRow As Integer = -1
                'For i = 0 To Grid0.DataTable.Rows.Count - 1
                '    Dim sss1 As Integer = Grid0.Rows.GetParent(i)
                '    Dim ibool1 As Boolean = Grid0.Rows.IsLeaf(i)
                '    ibool1 = ibool1

                '    If Grid0.Rows.GetParent(i) <> -1 And Grid0.Rows.IsLeaf(i) = True Then
                '        If clrRow = -1 Then clrRow = i : ParentRow = Grid0.Rows.GetParent(i)
                '    End If
                '    If i <> clrRow And clrRow <> -1 Then

                '        Grid0.DataTable.Columns.Item(1).Cells.Item(i).Value = ""
                '        Grid0.DataTable.Columns.Item(2).Cells.Item(i).Value = ""
                '        Grid0.DataTable.Columns.Item(3).Cells.Item(i).Value = 0
                '    End If
                '    'Grid0.DataTable.Columns.Item(1).Cells.Item(i).Value = ""
                '    If clrRow <> -1 And Grid0.Rows.GetParent(i) = -1 Then
                '        clrRow = -1
                '    End If
                'Next
                'For i = 0 To Grid0.Rows.Count - 1
                '    Dim sss As Integer = Grid0.Rows.GetParent(i)
                '    Dim ibool As Boolean = Grid0.Rows.IsLeaf(i)
                '    If Grid0.Rows.GetParent(i) <> -1 And Grid0.Rows.IsLeaf(i) = True Then
                '        If clrRow = -1 Then clrRow = i : ParentRow = Grid0.Rows.GetParent(i)
                '    End If
                '    If i <> clrRow And clrRow <> -1 Then
                '        Grid0.DataTable.Columns.Item(1).Cells.Item(i).Value = ""
                '        Grid0.DataTable.Columns.Item(2).Cells.Item(i).Value = ""
                '        Grid0.DataTable.Columns.Item(3).Cells.Item(i).Value = 0
                '    End If
                '    'Grid0.DataTable.Columns.Item(1).Cells.Item(i).Value = ""
                '    If clrRow <> -1 And Grid0.Rows.GetParent(i) = -1 Then
                '        clrRow = -1
                '    End If

                'Next

            Catch ex As Exception
            End Try

        End Sub

        Private Sub Form_CloseAfter(pVal As SAPbouiCOM.SBOItemEventArg)
            Try
                objaddon.objmenuevent.ItemCode = ""
            Catch ex As Exception
            End Try

        End Sub

        Private Sub Form_ResizeAfter(pVal As SAPbouiCOM.SBOItemEventArg)
            Try
                Grid0.AutoResizeColumns()
            Catch ex As Exception
            End Try

        End Sub
    End Class
End Namespace
