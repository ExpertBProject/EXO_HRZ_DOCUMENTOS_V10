Imports System.Xml
Imports SAPbouiCOM

Public Class SAP_STOCK
    Inherits EXO_UIAPI.EXO_DLLBase
    Dim _iLineNumRightClick As Integer = -1
    Public Sub New(ByRef oObjGlobal As EXO_UIAPI.EXO_UIAPI, ByRef actualizar As Boolean, usaLicencia As Boolean, idAddOn As Integer)
        MyBase.New(oObjGlobal, actualizar, usaLicencia, idAddOn)
        If actualizar Then
            cargaDatos()

        End If
    End Sub

#Region "Inicialización"
    Private Sub cargaDatos()
        Dim sXML As String = ""
        Dim res As String = ""

        If objGlobal.refDi.comunes.esAdministrador Then

            sXML = objGlobal.funciones.leerEmbebido(Me.GetType(), "UDFs_OPOR.xml")
            objGlobal.refDi.comunes.LoadBDFromXML(sXML)
            objGlobal.SBOApp.StatusBar.SetText("Validando: UDFs_OPOR", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            res = objGlobal.SBOApp.GetLastBatchResults
        End If
    End Sub

    Public Overrides Function filtros() As SAPbouiCOM.EventFilters
        Dim fXML As String = objGlobal.funciones.leerEmbebido(Me.GetType(), "XML_FILTROS.xml")
        Dim filtro As SAPbouiCOM.EventFilters = New SAPbouiCOM.EventFilters()
        filtro.LoadFromXML(fXML)
        Return filtro
    End Function

    Public Overrides Function menus() As System.Xml.XmlDocument
        Return Nothing
    End Function

#End Region

#Region "Eventos"
    Public Overrides Function SBOApp_FormDataEvent(ByVal infoEvento As BusinessObjectInfo) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim sDocEntry As String = "0"
        Dim oXml As New Xml.XmlDocument
        Dim bolModificar As Boolean = True
        Try
            If infoEvento.BeforeAction = True Then
                Select Case infoEvento.FormTypeEx
                    Case "142", "143", "139", "140"
                        Select Case infoEvento.EventType

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE


                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE

                        End Select

                End Select

            Else

                Select Case infoEvento.FormTypeEx
                    Case "142", "143", "139", "140"
                        Select Case infoEvento.EventType

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD


                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE
                                If infoEvento.ActionSuccess Then
                                    oForm = objGlobal.SBOApp.Forms.Item(infoEvento.FormUID)
                                    If bolModificar = True Then
                                        oForm = objGlobal.SBOApp.Forms.Item(infoEvento.FormUID)
                                        'creación de pedido de venta en otra emprea
                                        oXml.LoadXml(infoEvento.ObjectKey)
                                        sDocEntry = oXml.SelectSingleNode("DocumentParams/DocEntry").InnerText
                                        GrabarStock(oForm, oForm.BusinessObject.Type, -1, CInt(sDocEntry))
                                        bolModificar = False
                                    End If
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                                If infoEvento.ActionSuccess Then
                                    oForm = objGlobal.SBOApp.Forms.Item(infoEvento.FormUID)
                                    'creación de pedido de venta en otra emprea
                                    oXml.LoadXml(infoEvento.ObjectKey)
                                    sDocEntry = oXml.SelectSingleNode("DocumentParams/DocEntry").InnerText

                                    GrabarStock(oForm, oForm.BusinessObject.Type, -1, CInt(sDocEntry))

                                End If

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD

                        End Select

                End Select

            End If

            Return MyBase.SBOApp_FormDataEvent(infoEvento)

        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)

            Return False
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)

            Return False
        Finally
            EXO_CleanCOM.CLiberaCOM.Form(oForm)
        End Try
    End Function


    Public Overrides Function SBOApp_ItemEvent(ByVal infoEvento As ItemEvent) As Boolean
        Try
            If infoEvento.InnerEvent = False Then
                If infoEvento.BeforeAction = False Then
                    Select Case infoEvento.FormTypeEx
                        Case "142"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED


                                Case SAPbouiCOM.BoEventTypes.et_CLICK

                                Case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK

                                Case SAPbouiCOM.BoEventTypes.et_VALIDATE

                                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                                Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE



                            End Select
                    End Select
                ElseIf infoEvento.BeforeAction = True Then
                    Select Case infoEvento.FormTypeEx
                        Case "142"

                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                                Case SAPbouiCOM.BoEventTypes.et_CLICK

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_VALIDATE

                                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                                Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE

                            End Select

                    End Select
                End If

            Else
                If infoEvento.BeforeAction = False Then
                    Select Case infoEvento.FormTypeEx
                        Case "142"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_FORM_VISIBLE

                                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD


                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED


                                Case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS

                                Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS

                            End Select

                    End Select
                Else
                    Select Case infoEvento.FormTypeEx
                        Case "142"

                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                                Case SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED

                                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE

                            End Select

                    End Select
                End If
            End If

            Return MyBase.SBOApp_ItemEvent(infoEvento)

        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)

            Return False
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)

            Return False
        End Try
    End Function



#End Region
#Region "Auxiliares"

    Public Sub GrabarStock(ByRef oForm As SAPbouiCOM.Form, ByRef sObjType As String, ByRef iNumLin As Integer, ByRef iDocEntry As Integer)

        Dim oRs As SAPbobsCOM.Recordset = Nothing
        Dim oRsCrear As SAPbobsCOM.Recordset = Nothing
        Dim oRsStock As SAPbobsCOM.Recordset = Nothing
        Dim oXml As System.Xml.XmlDocument = New System.Xml.XmlDocument
        Dim oNodes As System.Xml.XmlNodeList = Nothing
        Dim oNode As System.Xml.XmlNode = Nothing
        Dim oNodes2 As System.Xml.XmlNodeList = Nothing
        Dim oNode2 As System.Xml.XmlNode = Nothing
        Dim NumLin As Integer = 0
        Dim DocNum As String
        Dim sCodart As String = ""
        Dim sCodAlm As String = ""
        Dim sTablaC As String = ""
        Dim sTablaL As String = ""

        Try
            Dim sSql As String = ""

            Dim dblStockAlm As Double = 0
            Dim dblStockTot As Double = 0
            Select Case sObjType
                Case "17"
                    sTablaC = "ORDR"
                    sTablaL = "RDR1"
                Case "15"
                    sTablaC = "ODLN"
                    sTablaL = "DLN1"
                Case "22"
                    sTablaC = "OPOR"
                    sTablaL = "POR1"
                Case "20"
                    sTablaC = "OPDN"
                    sTablaL = "PDN1"
            End Select

            oRs = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
            oRsCrear = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
            oRsStock = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)

            DocNum = (CType(oForm.Items.Item("8").Specific, SAPbouiCOM.EditText).Value)
            'sObjType = oForm.TypeEx
            If _iLineNumRightClick = -1 Then
                'todas las lineas
                sSql = "SELECT T0.""DocNum"", T0.""ObjType"", T1.""LineNum"", T1.""ItemCode"", T1.""WhsCode"" FROM """ & sTablaC & """  T0  INNER JOIN  """ & sTablaL & """ T1 ON T0.""DocEntry"" = T1.""DocEntry"" WHERE t0.""DocEntry"" =" & iDocEntry & ""
                oRs.DoQuery(sSql)
                oXml.LoadXml(oRs.GetAsXML())
                oNodes = oXml.SelectNodes("//row")
                If oRs.RecordCount > 0 Then
                    For i As Integer = 0 To oNodes.Count - 1
                        oNode = oNodes.Item(i)

                        NumLin = CInt(oNode.SelectSingleNode("LineNum").InnerText.ToString)
                        sCodart = oNode.SelectSingleNode("ItemCode").InnerText.ToString
                        sCodAlm = oNode.SelectSingleNode("WhsCode").InnerText.ToString
                        sSql = "SELECT COALESCE(SUM(T0.""OnHand""),0) ""StockTotal"" FROM ""OITW"" T0 WHERE T0.""Locked"" ='N' and ""ItemCode"" ='" & sCodart & "'"
                        dblStockTot = CDbl(objGlobal.refDi.SQL.sqlStringB1(sSql))

                        'delete
                        sSql = "DELETE FROM ""STOCKALM"" WHERE   ""OBJTYPE"" ='" & sObjType & "' AND ""DOCNUM"" =" & DocNum & " AND ""LINENUM""=" & NumLin & ""
                        oRsCrear.DoQuery(sSql)

                        'cargar stock
                        sSql = "SELECT T0.""ItemCode"", T0.""WhsCode"", T0.""OnHand"" FROM OITW T0 WHERE T0.""Locked"" ='N' and ""ItemCode"" ='" & sCodart & "'"
                        oRsStock.DoQuery(sSql)
                        oXml.LoadXml(oRsStock.GetAsXML())
                        oNodes2 = oXml.SelectNodes("//row")

                        If oRsStock.RecordCount > 0 Then
                            For j As Integer = 0 To oNodes.Count - 1
                                oNode2 = oNodes2.Item(j)
                                'insert
                                sSql = "INSERT INTO ""STOCKALM"" (""OBJTYPE"",""DOCNUM"",""LINENUM"",""ITEMCODE"",""WHSCODE"",""STOCKALM"",""STOCKTOT"") 
                    VALUES ('" & sObjType & "'," & DocNum & "," & NumLin & ",'" & sCodart & "', '" & oNode2.SelectSingleNode("WhsCode").InnerText.ToString & "'," & CDbl(oNode2.SelectSingleNode("OnHand").InnerText.ToString.Replace(",", ".")) & ", " & CDbl(dblStockTot) & ")"
                                oRsCrear.DoQuery(sSql)
                            Next

                        End If
                    Next
                End If
            Else

                NumLin = CInt(CType(CType(oForm.Items.Item("38").Specific, SAPbouiCOM.Matrix).Columns.Item("110").Cells.Item(_iLineNumRightClick).Specific, SAPbouiCOM.EditText).Value)
                sCodart = CType(CType(oForm.Items.Item("38").Specific, SAPbouiCOM.Matrix).Columns.Item("1").Cells.Item(_iLineNumRightClick).Specific, SAPbouiCOM.EditText).Value
                sCodAlm = CType(CType(oForm.Items.Item("38").Specific, SAPbouiCOM.Matrix).Columns.Item("24").Cells.Item(_iLineNumRightClick).Specific, SAPbouiCOM.EditText).Value

                sSql = "SELECT COALESCE(SUM(T0.""OnHand""),0) ""StockTotal"" FROM ""OITW"" T0 WHERE T0.""Locked"" ='N' and ""ItemCode"" ='" & sCodart & "'"
                dblStockTot = CDbl(objGlobal.refDi.SQL.sqlStringB1(sSql))

                'delete
                sSql = "DELETE FROM ""STOCKALM"" WHERE   ""OBJTYPE"" ='" & sObjType & "' AND ""DOCNUM"" =" & DocNum & " AND ""LINENUM""=" & NumLin & ""
                oRsCrear.DoQuery(sSql)

                'cargar stock
                sSql = "SELECT T0.""ItemCode"", T0.""WhsCode"", T0.""OnHand"" FROM OITW T0 WHERE T0.""Locked"" ='N' and ""ItemCode"" ='" & sCodart & "'"
                oRsStock.DoQuery(sSql)
                oXml.LoadXml(oRsStock.GetAsXML())
                oNodes = oXml.SelectNodes("//row")

                If oRsStock.RecordCount > 0 Then
                    For i As Integer = 0 To oNodes.Count - 1
                        oNode = oNodes.Item(i)
                        'insert
                        sSql = "INSERT INTO ""STOCKALM"" (""OBJTYPE"",""DOCNUM"",""LINENUM"",""ITEMCODE"",""WHSCODE"",""STOCKALM"",""STOCKTOT"") 
                    VALUES ('" & sObjType & "'," & DocNum & "," & NumLin & ",'" & sCodart & "', '" & oNode.SelectSingleNode("WhsCode").InnerText.ToString & "'," & CDbl(oNode.SelectSingleNode("OnHand").InnerText.ToString.Replace(",", ".")) & ", " & CDbl(dblStockTot) & ")"
                        oRsCrear.DoQuery(sSql)
                    Next

                End If
            End If




        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally

            'EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRsCrear, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRsStock, Object))
        End Try
    End Sub

    Public Sub AbrirStock(ByRef oForm As SAPbouiCOM.Form, ByRef sObjType As String)

        OpenForm()

    End Sub
    Public Function OpenForm() As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim sSQL As String = ""
        Dim oFP As SAPbouiCOM.FormCreationParams = Nothing
        Dim EXO_Xml As New EXO_UIAPI.EXO_XML(objGlobal)
        Dim NumLin As Integer = 0
        Dim sObjType As String
        Dim oForm2 As SAPbouiCOM.Form = Nothing

        OpenForm = False

        Try

            oFP = CType(objGlobal.SBOApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams), SAPbouiCOM.FormCreationParams)
            oFP.XmlData = objGlobal.leerEmbebido(Me.GetType(), "EXOCONSTOCK.srf")
            oFP.Modality = BoFormModality.fm_Modal
            Try

                oForm = objGlobal.SBOApp.Forms.AddEx(oFP)
                oForm2 = objGlobal.SBOApp.Forms.ActiveForm
                sObjType = oForm2.BusinessObject.Type
                NumLin = CInt(CType(CType(oForm2.Items.Item("38").Specific, SAPbouiCOM.Matrix).Columns.Item("110").Cells.Item(_iLineNumRightClick).Specific, SAPbouiCOM.EditText).Value)
                oForm.DataSources.UserDataSources.Item("UDDOCNUM").ValueEx = CType(oForm2.Items.Item("8").Specific, SAPbouiCOM.EditText).Value
                oForm.DataSources.UserDataSources.Item("UDLINENUM").ValueEx = (CType(CType(oForm2.Items.Item("38").Specific, SAPbouiCOM.Matrix).Columns.Item("110").Cells.Item(_iLineNumRightClick).Specific, SAPbouiCOM.EditText).Value)
                oForm.DataSources.UserDataSources.Item("UDTYPE").Value = sObjType

            Catch ex As Exception
                If ex.Message.StartsWith("Form - already exists") = True Then
                    objGlobal.SBOApp.StatusBar.SetText("El formulario ya está abierto.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Exit Function
                ElseIf ex.Message.StartsWith("Se produjo un error interno") = True Then 'Falta de autorización
                    Exit Function
                End If
            End Try



            OpenForm = True
        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            oForm.Visible = True
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
        End Try
    End Function

    Public Overrides Function SBOApp_RightClickEvent(ByVal infoEvento As ContextMenuInfo) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim oCreationPackage As SAPbouiCOM.MenuCreationParams = Nothing
        Dim oMenuItem As SAPbouiCOM.MenuItem = Nothing
        Dim oMenus As SAPbouiCOM.Menus = Nothing

        Try
            oForm = objGlobal.SBOApp.Forms.Item(infoEvento.FormUID)

            If infoEvento.BeforeAction = False Then
                Select Case oForm.TypeEx
                    Case "142", "143", "139", "140"
                        If objGlobal.SBOApp.Menus.Exists("EXO_MNUSTOCK") Then
                            objGlobal.SBOApp.Menus.RemoveEx("EXO_MNUSTOCK")
                        End If

                End Select

            Else

                Select Case oForm.TypeEx
                    Case "142", "143", "139", "140"
                        If infoEvento.ItemUID = "38" Then
                            If infoEvento.Row >= 0 Then
                                'oForm.DataSources.UserDataSources.Item("LineRight").ValueEx = infoEvento.Row.ToString
                                _iLineNumRightClick = infoEvento.Row
                                If Not objGlobal.SBOApp.Menus.Exists("EXO_MNUSTOCK") Then
                                    oCreationPackage = CType(objGlobal.SBOApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams), SAPbouiCOM.MenuCreationParams)

                                    oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
                                    oCreationPackage.UniqueID = "EXO_MNUSTOCK"
                                    oCreationPackage.String = "Consultar Stock”
                                    oCreationPackage.Enabled = True

                                    oMenuItem = objGlobal.SBOApp.Menus.Item("1280") 'Data'
                                    oMenus = oMenuItem.SubMenus
                                    oMenus.AddEx(oCreationPackage)
                                End If

                            End If
                        End If

                End Select

            End If

            Return MyBase.SBOApp_RightClickEvent(infoEvento)

        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
            Return False
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
            Return False
        Finally
            EXO_CleanCOM.CLiberaCOM.MenuItem(oMenuItem)
            EXO_CleanCOM.CLiberaCOM.Menus(oMenus)
            EXO_CleanCOM.CLiberaCOM.MenuCreation(oCreationPackage)
            EXO_CleanCOM.CLiberaCOM.Form(oForm)
        End Try
    End Function

    Public Overrides Function SBOApp_MenuEvent(ByVal infoEvento As MenuEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim sMensaje As String = ""


        Try
            oForm = objGlobal.SBOApp.Forms.ActiveForm
            If infoEvento.BeforeAction = True Then

            Else
                Select Case infoEvento.MenuUID
                    Case "EXO_MNUSTOCK"
                        If _iLineNumRightClick > 0 Then
                            'si esta en modo consulta, no se graba.
                            If oForm.Mode = BoFormMode.fm_ADD_MODE Or oForm.Mode = BoFormMode.fm_EDIT_MODE Then
                                GrabarStock(oForm, oForm.BusinessObject.Type, _iLineNumRightClick, 0)
                            End If
                            AbrirStock(oForm, oForm.BusinessObject.Type)

                        Else
                            sMensaje = "Tiene que seleccionar una línea."
                            objGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            objGlobal.SBOApp.MessageBox(sMensaje)
                        End If


                End Select
            End If

            Return True

        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
            Return False
        Finally

        End Try
    End Function

#End Region

End Class
