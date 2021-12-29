Imports System.Xml
Imports SAPbobsCOM
Imports SAPbouiCOM

Public Class EXO_CONFTRASDOC
    Inherits EXO_UIAPI.EXO_DLLBase
    Public Sub New(ByRef oObjGlobal As EXO_UIAPI.EXO_UIAPI, ByRef actualizar As Boolean, usaLicencia As Boolean, idAddOn As Integer)
        MyBase.New(oObjGlobal, actualizar, usaLicencia, idAddOn)

        If actualizar Then
            cargaDatos()
            cargaAutorizaciones()
        End If
    End Sub

#Region "Inicialización"

    Public Overrides Function filtros() As SAPbouiCOM.EventFilters
        Dim fXML As String = objGlobal.funciones.leerEmbebido(Me.GetType(), "XML_FILTROS.xml")
        Dim filtro As SAPbouiCOM.EventFilters = New SAPbouiCOM.EventFilters()
        filtro.LoadFromXML(fXML)
        Return filtro
    End Function

    Public Overrides Function menus() As System.Xml.XmlDocument
        Dim menuXML As Xml.XmlDocument = New Xml.XmlDocument
        Dim sPath As String = ""

        sPath = objGlobal.refDi.OGEN.pathGeneral & "\02.Menus" & "\XML_MENU.xml"
        menuXML.Load(sPath)
        Return menuXML

    End Function

    Private Sub cargaDatos()
        Dim sXML As String = ""
        Dim res As String = ""

        If objGlobal.refDi.comunes.esAdministrador Then
            'UDO Tarificador 401
            sXML = objGlobal.funciones.leerEmbebido(Me.GetType(), "UDO_EXO_CONFTRASDOC.xml")
            objGlobal.refDi.comunes.LoadBDFromXML(sXML)
            objGlobal.SBOApp.StatusBar.SetText("Validando: UDO UDO_EXO_CONFTRASDOC", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            res = objGlobal.SBOApp.GetLastBatchResults
        End If
    End Sub

    Private Sub cargaAutorizaciones()
        Dim autorizacionXML As String = objGlobal.funciones.leerEmbebido(Me.GetType(), "EXO_AUCONFTRASDOC.xml")
        objGlobal.refDi.comunes.LoadBDFromXML(autorizacionXML)
        Dim res As String = objGlobal.SBOApp.GetLastBatchResults
    End Sub

#End Region
#Region "Eventos"


    Public Overrides Function SBOApp_ItemEvent(infoEvento As ItemEvent) As Boolean
        Try
            If infoEvento.InnerEvent = False Then
                If infoEvento.BeforeAction = False Then
                    Select Case infoEvento.FormTypeEx
                        Case "UDO_FT_EXO_CONFTRASDOC"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_CLICK

                                Case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK

                                Case SAPbouiCOM.BoEventTypes.et_VALIDATE
                                    If EventHandler_Validate_After(infoEvento) = False Then
                                        Return False
                                    End If
                                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                                Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE

                            End Select
                    End Select
                ElseIf infoEvento.BeforeAction = True Then
                    Select Case infoEvento.FormTypeEx
                        Case "UDO_FT_EXO_CONFTRASDOC"

                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                                Case SAPbouiCOM.BoEventTypes.et_CLICK

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_VALIDATE

                                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                                Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED

                            End Select
                    End Select
                End If

            Else
                If infoEvento.BeforeAction = False Then
                    Select Case infoEvento.FormTypeEx
                        Case "UDO_FT_EXO_CONFTRASDOC"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_FORM_VISIBLE

                                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD

                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                                    If EventHandler_Choose_FromList_After(infoEvento) = False Then
                                        Return False
                                    End If
                            End Select
                    End Select
                Else
                    Select Case infoEvento.FormTypeEx
                        Case "UDO_FT_EXO_CONFTRASDOC"

                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                                    If EventHandler_Choose_FromList_Before(infoEvento) = False Then
                                        Return False
                                    End If
                                Case SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED

                                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                    'If EventHandler_Form_Load(infoEvento) = False Then
                                    '    GC.Collect()
                                    '    Return False
                                    'End If
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
    Public Overrides Function SBOApp_MenuEvent(infoEvento As MenuEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing

        Try
            If infoEvento.BeforeAction = True Then

            Else

                Select Case infoEvento.MenuUID
                    Case "mConfTras"
                        If CargarForm() = False Then
                            Return False
                        End If
                End Select
            End If

            Return MyBase.SBOApp_MenuEvent(infoEvento)

        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
            Return False
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
            Return False
        Finally

            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
        End Try
    End Function

    Private Function EventHandler_Choose_FromList_Before(ByRef pVal As ItemEvent) As Boolean
        Dim sSQL As String = ""

        Dim oXml As System.Xml.XmlDocument = New System.Xml.XmlDocument
        Dim oNodes As System.Xml.XmlNodeList = Nothing
        Dim oNode As System.Xml.XmlNode = Nothing
        Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent = Nothing
        Dim oConds As SAPbouiCOM.Conditions = Nothing
        Dim oCond As SAPbouiCOM.Condition = Nothing
        Dim oForm As SAPbouiCOM.Form = Nothing
    

        EventHandler_Choose_FromList_Before = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)


            If pVal.ItemUID = "21_U_E" Then '
                'Recuperamos el formulario de origen


                oCFLEvento = CType(pVal, SAPbouiCOM.IChooseFromListEvent)

                oConds = New SAPbouiCOM.Conditions

                oCond = oConds.Add
                oCond.Alias = "CardType"
                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCond.CondVal = "S"
                oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND

                oCond = oConds.Add
                oCond.Alias = "validFor"
                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCond.CondVal = "Y"


                oForm.ChooseFromLists.Item(oCFLEvento.ChooseFromListUID).SetConditions(oConds)
            End If

            If pVal.ItemUID = "23_U_E" Then '
                'Recuperamos el formulario de origen

                oCFLEvento = CType(pVal, SAPbouiCOM.IChooseFromListEvent)

                oConds = New SAPbouiCOM.Conditions

                oCond = oConds.Add
                oCond.Alias = "CardType"
                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCond.CondVal = "C"
                oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND

                oCond = oConds.Add
                oCond.Alias = "validFor"
                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCond.CondVal = "Y"

                oForm.ChooseFromLists.Item(oCFLEvento.ChooseFromListUID).SetConditions(oConds)
            End If

            EventHandler_Choose_FromList_Before = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally

            EXO_CleanCOM.CLiberaCOM.FormConditions(oConds)
            EXO_CleanCOM.CLiberaCOM.FormCondition(oCond)
            EXO_CleanCOM.CLiberaCOM.Form(oForm)
        End Try
    End Function

    Private Function EventHandler_Choose_FromList_After(ByRef pVal As ItemEvent) As Boolean
        Dim sSQL As String = ""

        Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent = Nothing
        Dim oDataTable As SAPbouiCOM.DataTable = Nothing
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim sCFL_ID As String = ""
        Dim oFormOrigen As SAPbouiCOM.Form = Nothing

        EventHandler_Choose_FromList_After = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)



            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                oForm = Nothing
                Return True
            End If

            oCFLEvento = CType(pVal, SAPbouiCOM.IChooseFromListEvent)

            Select Case oCFLEvento.ChooseFromListUID
                Case "CFLPRO" 'PROVEEDOR
                    sCFL_ID = oCFLEvento.ChooseFromListUID
                    oDataTable = oCFLEvento.SelectedObjects
                    If oDataTable IsNot Nothing Then
                        If pVal.ItemUID = "21_U_E" Then 'proveedor
                            'Recuperamos el formulario de origen
                            Try
                                CType(oForm.Items.Item("22_U_E").Specific, SAPbouiCOM.EditText).Value = oDataTable.GetValue("CardName", 0).ToString
                            Catch ex As Exception

                            End Try

                            Try
                                CType(oForm.Items.Item("21_U_E").Specific, SAPbouiCOM.EditText).Value = oDataTable.GetValue("CardCode", 0).ToString
                            Catch ex As Exception

                            End Try

                        End If
                    End If
                Case "CFLCLI" 'CLIENTE
                    sCFL_ID = oCFLEvento.ChooseFromListUID
                    oDataTable = oCFLEvento.SelectedObjects
                    If oDataTable IsNot Nothing Then
                        If pVal.ItemUID = "23_U_E" Then 'cliente
                            'Recuperamos el formulario de origen
                            Try
                                CType(oForm.Items.Item("24_U_E").Specific, SAPbouiCOM.EditText).Value = oDataTable.GetValue("CardName", 0).ToString
                            Catch ex As Exception

                            End Try

                            Try
                                CType(oForm.Items.Item("23_U_E").Specific, SAPbouiCOM.EditText).Value = oDataTable.GetValue("CardCode", 0).ToString
                            Catch ex As Exception

                            End Try

                        End If
                    End If
            End Select

            EventHandler_Choose_FromList_After = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally

            EXO_CleanCOM.CLiberaCOM.FormDatatable(oDataTable)
            EXO_CleanCOM.CLiberaCOM.Form(oFormOrigen)
            EXO_CleanCOM.CLiberaCOM.Form(oForm)
        End Try
    End Function

    Private Function EventHandler_Validate_After(ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim sSql As String = ""
        Dim oRs As SAPbobsCOM.Recordset = Nothing

        EventHandler_Validate_After = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)

            If pVal.ItemUID = "13_U_E" Then '
                If pVal.ItemChanged = True Then
                    CType(oForm.Items.Item("0_U_E").Specific, SAPbouiCOM.EditText).Value = CType(oForm.Items.Item("13_U_E").Specific, SAPbouiCOM.EditText).Value
                End If
            End If

            If pVal.ItemUID = "17_U_E" Then '
                If pVal.ItemChanged = True Then
                    'mirar si existe el codigo en la empresa de "nombre de la base de datos
                    If CType(oForm.Items.Item("17_U_E").Specific, SAPbouiCOM.EditText).Value <> "" Then
                        oRs = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
                        sSql = "SELECT T0.""CardName"" FROM """ & CType(oForm.Items.Item("13_U_E").Specific, SAPbouiCOM.EditText).Value & """.""OCRD"" T0 WHERE T0.""CardType"" ='S' AND T0.""CardCode"" ='" & CType(oForm.Items.Item("17_U_E").Specific, SAPbouiCOM.EditText).Value & "' "
                        oRs.DoQuery(sSql)
                        If oRs.RecordCount > 0 Then
                            CType(oForm.Items.Item("18_U_E").Specific, SAPbouiCOM.EditText).Value = oRs.Fields.Item("CardName").Value.ToString()
                        Else
                            objGlobal.SBOApp.StatusBar.SetText("(EXO) - El código introducido no es válido en la empresa " & CType(oForm.Items.Item("13_U_E").Specific, SAPbouiCOM.EditText).Value & " .", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        End If
                    Else
                        CType(oForm.Items.Item("18_U_E").Specific, SAPbouiCOM.EditText).Value = ""
                    End If
                End If
            End If

            If pVal.ItemUID = "19_U_E" Then 'CLIENTE
                If pVal.ItemChanged = True Then
                    'mirar si existe el codigo en la empresa de "nombre de la base de datos
                    If CType(oForm.Items.Item("19_U_E").Specific, SAPbouiCOM.EditText).Value <> "" Then

                        oRs = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
                        sSql = "SELECT T0.""CardName"" FROM """ & CType(oForm.Items.Item("13_U_E").Specific, SAPbouiCOM.EditText).Value & """.""OCRD"" T0 WHERE T0.""CardType"" ='C' AND T0.""CardCode"" ='" & CType(oForm.Items.Item("19_U_E").Specific, SAPbouiCOM.EditText).Value & "' "
                        oRs.DoQuery(sSql)
                        If oRs.RecordCount > 0 Then
                            CType(oForm.Items.Item("20_U_E").Specific, SAPbouiCOM.EditText).Value = oRs.Fields.Item("CardName").Value.ToString()
                        Else
                            objGlobal.SBOApp.StatusBar.SetText("(EXO) - El código introducido no es válido en la empresa " & CType(oForm.Items.Item("13_U_E").Specific, SAPbouiCOM.EditText).Value & " .", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        End If
                    Else
                        CType(oForm.Items.Item("20_U_E").Specific, SAPbouiCOM.EditText).Value = ""
                    End If
                End If
            End If
            EventHandler_Validate_After = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
            EXO_CleanCOM.CLiberaCOM.Form(oForm)
        End Try
    End Function
#End Region
    Public Function CargarForm() As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim sSQL As String = ""
        Dim oFP As SAPbouiCOM.FormCreationParams = Nothing
        Dim EXO_Xml As New EXO_UIAPI.EXO_XML(objGlobal)
        Dim bt As SAPbouiCOM.Item = Nothing
        CargarForm = False

        Try

            oFP = CType(objGlobal.SBOApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams), SAPbouiCOM.FormCreationParams)
            oFP.XmlData = objGlobal.leerEmbebido(Me.GetType(), "EXO_CONFTRASDOC.srf")

            Try
                oForm = objGlobal.SBOApp.Forms.AddEx(oFP)
            Catch ex As Exception
                If ex.Message.StartsWith("Form - already exists") = True Then
                    objGlobal.SBOApp.StatusBar.SetText("El formulario ya está abierto.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Exit Function
                ElseIf ex.Message.StartsWith("Se produjo un error interno") = True Then 'Falta de autorización
                    Exit Function
                End If
            End Try



            CargarForm = True
        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            oForm.Visible = True
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
        End Try
    End Function

    Public Function ComprobarCliProv(ByRef oForm As Form) As Boolean
        Dim sSQL As String = ""
        Dim oRs As SAPbobsCOM.Recordset = Nothing

        ComprobarCliProv = True
        Try
            oRs = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
            sSQL = "SELECT T0.""CardCode"",T0.""CardName"" FROM """ & CType(oForm.Items.Item("13_U_E").Specific, SAPbouiCOM.EditText).Value & """.""OCRD"" T0 WHERE T0.""CardType"" ='C' AND T0.""CardCode"" ='" & CType(oForm.Items.Item("19_U_E").Specific, SAPbouiCOM.EditText).Value & "' "
            oRs.DoQuery(sSQL)
            If oRs.RecordCount > 0 Then
                CType(oForm.Items.Item("20_U_E").Specific, SAPbouiCOM.EditText).Value = oRs.Fields.Item("CardName").Value.ToString()
            Else
                objGlobal.SBOApp.StatusBar.SetText("(EXO) - El código introducido no es válido en la empresa " & CType(oForm.Items.Item("13_U_E").Specific, SAPbouiCOM.EditText).Value & " .", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                ComprobarCliProv = False
            End If




        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            oForm.Visible = True
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
        End Try

    End Function

    Public Overrides Function SBOApp_FormDataEvent(ByVal infoEvento As BusinessObjectInfo) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim sDocEntry As String = "0"
        Dim oXml As New Xml.XmlDocument
        Dim bolModificar As Boolean = True
        Try
            If infoEvento.BeforeAction = True Then
                Select Case infoEvento.FormTypeEx
                    Case "UDO_FT_EXO_CONFTRASDOC"
                        Select Case infoEvento.EventType

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE
                                'antes de actualizar comprobar si el pedido es con destino
                                oForm = objGlobal.SBOApp.Forms.Item(infoEvento.FormUID)
                                If ComprobarCliProv(oForm) = False Then
                                    Return False
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                                oForm = objGlobal.SBOApp.Forms.Item(infoEvento.FormUID)
                                If ComprobarCliProv(oForm) = False Then
                                    Return False
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE

                        End Select

                End Select

            Else

                Select Case infoEvento.FormTypeEx
                    Case "UDO_FT_EXO_CONFTRASDOC"
                        Select Case infoEvento.EventType

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                                If infoEvento.ActionSuccess Then

                                End If


                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE
                                If infoEvento.ActionSuccess Then

                                End If
                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                                If infoEvento.ActionSuccess Then

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

End Class
