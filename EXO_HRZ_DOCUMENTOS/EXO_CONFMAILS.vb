Imports System.Xml
Imports SAPbobsCOM
Imports SAPbouiCOM

Public Class EXO_CONFMAILS
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
            sXML = objGlobal.funciones.leerEmbebido(Me.GetType(), "UDO_EXO_CONFIG.xml")
            objGlobal.refDi.comunes.LoadBDFromXML(sXML)
            objGlobal.SBOApp.StatusBar.SetText("Validando: UDO UDO_EXO_CONFIG", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            res = objGlobal.SBOApp.GetLastBatchResults
        End If
    End Sub

    Private Sub cargaAutorizaciones()
        Dim autorizacionXML As String = objGlobal.funciones.leerEmbebido(Me.GetType(), "EXO_AUEXOCONFIG.xml")
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
                        Case "UDO_FT_EXO_CONFIG"
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
                        Case "UDO_FT_EXO_CONFIG"

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
                        Case "UDO_FT_EXO_CONFIG"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_FORM_VISIBLE

                                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD

                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                            End Select
                    End Select
                Else
                    Select Case infoEvento.FormTypeEx
                        Case "UDO_FT_EXO_CONFIG"

                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

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
                    Case "mConfMail"
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


#End Region
    Public Function CargarForm() As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim oFP As SAPbouiCOM.FormCreationParams = Nothing
        Dim oRs As SAPbobsCOM.Recordset = Nothing
        Dim oDI_CONFAC As EXO_DIAPI.EXO_UDOEntity = Nothing
        Dim sCodeEnvio As String = ""

        CargarForm = False

        Try

            oFP = CType(objGlobal.SBOApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams), SAPbouiCOM.FormCreationParams)
            oFP.XmlData = objGlobal.leerEmbebido(Me.GetType(), "EXO_CONFIG.srf")

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

            'comprobar si tiene registro, y si no tiene crearlo
            oRs = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)

            oRs.DoQuery("SELECT ""Code"" FROM ""@EXO_CONFIG""  ")

            If oRs.RecordCount = 0 Then
                'CREO EL REGISTRO
                '.conexionSAP.refCompañia
                oDI_CONFAC = New EXO_DIAPI.EXO_UDOEntity(objGlobal.refDi.comunes, "EXO_CONFIG")
                oDI_CONFAC.GetNew()
                oDI_CONFAC.SetValue("Code") = "CONFENVIO"
                oDI_CONFAC.SetValue("Name") = "CONFENVIO"
                If oDI_CONFAC.UDO_Add() = False Then
                    Throw New Exception("Error al crear registro CONFENVIO: " & oDI_CONFAC.GetLastError)
                End If

                sCodeEnvio = oDI_CONFAC.GetNewKey
            Else
                sCodeEnvio = oRs.Fields.Item("Code").Value.ToString()

            End If
            If sCodeEnvio <> "" Then
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                oForm.Items.Item("0_U_E").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 4, BoModeVisualBehavior.mvb_True)
                CType(oForm.Items.Item("0_U_E").Specific, SAPbouiCOM.EditText).Value = sCodeEnvio
                oForm.Items.Item("1").Click(BoCellClickType.ct_Regular)

            End If

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

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD

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
