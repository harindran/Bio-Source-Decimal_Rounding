Namespace Decimal_Rounding

    Public Class clsRightClickEvent
        Dim objform As SAPbouiCOM.Form
        Dim objglobalmethods As New clsGlobalMethods
        Dim ocombo As SAPbouiCOM.ComboBox
        Dim objmatrix As SAPbouiCOM.Matrix
        Dim strsql As String
        Dim objrs As SAPbobsCOM.Recordset

        Public Sub RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)
            Try
                If objaddon.objapplication.Menus.Item("1280").SubMenus.Exists("AT_Show_Batches") Then objaddon.objapplication.Menus.RemoveEx("AT_Show_Batches") : objaddon.objapplication.Menus.RemoveEx("AT_Remove_Row")
                Select Case objaddon.objapplication.Forms.ActiveForm.TypeEx
                    Case "170", "426"
                        objform = objaddon.objapplication.Forms.ActiveForm
                    Case "CT_PF_ManufacOrd"
                        objform = objaddon.objapplication.Forms.ActiveForm
                        ''objaddon.objapplication.Menus.RemoveEx("CT_PF_ManufacOrd|Items|del")
                        'objaddon.objapplication.Menus.Item("1280").SubMenus.RemoveEx("CT_PF_ManufacOrd|Items|del")
                        ''RightClickMenu_Delete("1280", "CT_PF_ManufacOrd|Items|del")
                        'If eventInfo.BeforeAction Then eventInfo.RemoveFromContent("CT_PF_ManufacOrd|Items|del")
                        'objform.EnableMenu("1293", True)
                        If eventInfo.Row <> -1 Then
                            objaddon.objmenuevent.ItemCode = objform.Items.Item("Items").Specific.Columns.Item("col_1").Cells.Item(eventInfo.Row).Specific.String
                            curMORow = eventInfo.Row
                        End If
                        If eventInfo.BeforeAction = True And (eventInfo.ItemUID = "" Or eventInfo.ItemUID = "Items") Then
                            RightClickMenu_Add("1280", "AT_Show_Batches", "Show Available Batch/Serial", 0)
                            If eventInfo.ItemUID = "Items" Then RightClickMenu_Add("1280", "AT_Remove_Row", "Remove Row", 0)
                        Else
                            RightClickMenu_Delete("1280", "AT_Show_Batches")
                            RightClickMenu_Delete("1280", "AT_Remove_Row")
                        End If

                End Select
            Catch ex As Exception
            End Try
        End Sub

        Private Sub RightClickMenu_Add(ByVal MainMenu As String, ByVal NewMenuID As String, ByVal NewMenuName As String, ByVal position As Integer)
            Dim omenus As SAPbouiCOM.Menus
            Dim omenuitem As SAPbouiCOM.MenuItem
            Dim oCreationPackage As SAPbouiCOM.MenuCreationParams
            oCreationPackage = objaddon.objapplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
            omenuitem = objaddon.objapplication.Menus.Item(MainMenu) 'Data'

            If Not omenuitem.SubMenus.Exists(NewMenuID) Then
                oCreationPackage.UniqueID = NewMenuID
                oCreationPackage.String = NewMenuName
                oCreationPackage.Position = position
                oCreationPackage.Enabled = True
                omenus = omenuitem.SubMenus
                omenus.AddEx(oCreationPackage)
            End If
        End Sub

        Private Sub RightClickMenu_Delete(ByVal MainMenu As String, ByVal NewMenuID As String)
            Dim omenuitem As SAPbouiCOM.MenuItem
            omenuitem = objaddon.objapplication.Menus.Item(MainMenu) 'Data'
            If omenuitem.SubMenus.Exists(NewMenuID) Then
                objaddon.objapplication.Menus.RemoveEx(NewMenuID)
            End If
        End Sub



    End Class

End Namespace
