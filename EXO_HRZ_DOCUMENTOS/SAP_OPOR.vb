Imports System.Xml
Imports SAPbouiCOM

Public Class SAP_OPOR
    Inherits EXO_UIAPI.EXO_DLLBase

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

    'Public Overrides Function SBOApp_FormDataEvent(ByVal infoEvento As BusinessObjectInfo) As Boolean
    '    Dim oForm As SAPbouiCOM.Form = Nothing
    '    Dim sDocEntry As String = "0"
    '    Dim oXml As New Xml.XmlDocument
    '    Dim bolModificar As Boolean = True
    '    Try
    '        If infoEvento.BeforeAction = True Then
    '            Select Case infoEvento.FormTypeEx
    '                Case "142"
    '                    Select Case infoEvento.EventType

    '                        Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD

    '                        Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE
    '                            'antes de actualizar comprobar si el pedido es con destino
    '                            oForm = objGlobal.SBOApp.Forms.Item(infoEvento.FormUID)
    '                            If oForm.DataSources.DBDataSources.Item("OPOR").GetValue("U_EXO_EMPRESAD", 0).ToString <> "-1" Then
    '                                If oForm.DataSources.DBDataSources.Item("OPOR").GetValue("U_EXO_EMPRESAD", 0).ToString <> "" Then
    '                                    If ComprobarDocDestino(oForm) = False Then
    '                                        bolModificar = False
    '                                        Return False
    '                                    End If
    '                                End If

    '                            End If
    '                        Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD

    '                        Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE

    '                    End Select

    '            End Select

    '        Else

    '            Select Case infoEvento.FormTypeEx
    '                Case "142"
    '                    Select Case infoEvento.EventType

    '                        Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
    '                            If infoEvento.ActionSuccess Then
    '                                oForm = objGlobal.SBOApp.Forms.Item(infoEvento.FormUID)
    '                                CargaComboEmpresa(oForm)
    '                            End If


    '                        Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE
    '                            If infoEvento.ActionSuccess Then
    '                                oForm = objGlobal.SBOApp.Forms.Item(infoEvento.FormUID)
    '                                If bolModificar = True Then
    '                                    oForm = objGlobal.SBOApp.Forms.Item(infoEvento.FormUID)
    '                                    'creación de pedido de venta en otra emprea
    '                                    oXml.LoadXml(infoEvento.ObjectKey)
    '                                    sDocEntry = oXml.SelectSingleNode("DocumentParams/DocEntry").InnerText
    '                                    CrearPedidoVenta(oForm, sDocEntry)
    '                                    bolModificar = False
    '                                End If
    '                            End If
    '                        Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
    '                            If infoEvento.ActionSuccess Then
    '                                oForm = objGlobal.SBOApp.Forms.Item(infoEvento.FormUID)
    '                                'creación de pedido de venta en otra emprea
    '                                oXml.LoadXml(infoEvento.ObjectKey)
    '                                sDocEntry = oXml.SelectSingleNode("DocumentParams/DocEntry").InnerText
    '                                CrearPedidoVenta(oForm, sDocEntry)

    '                            End If

    '                        Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD

    '                    End Select

    '            End Select

    '        End If

    '        Return MyBase.SBOApp_FormDataEvent(infoEvento)

    '    Catch exCOM As System.Runtime.InteropServices.COMException
    '        objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)

    '        Return False
    '    Catch ex As Exception
    '        objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)

    '        Return False
    '    Finally
    '        EXO_CleanCOM.CLiberaCOM.Form(oForm)
    '    End Try
    'End Function

    Public Overrides Function SBOApp_ItemEvent(ByVal infoEvento As ItemEvent) As Boolean
        Try
            If infoEvento.InnerEvent = False Then
                If infoEvento.BeforeAction = False Then
                    Select Case infoEvento.FormTypeEx
                        Case "142"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                    If EventHandler_ItemPressed_After(infoEvento) = False Then
                                        Return False
                                    End If

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
                                    If EventHandler_ItemPressed_After(infoEvento) = False Then
                                        Return False
                                    End If

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
                                    If EventHandler_Form_Load(infoEvento) = False Then
                                        Return False
                                    End If
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

    Private Function EventHandler_Form_Load(ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim oItem As SAPbouiCOM.Item

        Dim Path As String = ""
        Dim XmlDoc As New System.Xml.XmlDocument


        EventHandler_Form_Load = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)
            'oForm.Visible = True
            If pVal.ActionSuccess = False Then
                objGlobal.SBOApp.StatusBar.SetText("Presentando información...Espere por favor", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                ' combo empresas
                oItem = oForm.Items.Add("cmbEmpresa", BoFormItemTypes.it_COMBO_BOX)
                oItem.Top = oForm.Items.Item("70").Top + 15
                oItem.Left = oForm.Items.Item("14").Left
                oItem.Height = oForm.Items.Item("14").Height
                oItem.Width = oForm.Items.Item("14").Width
                oItem.FromPane = 0
                oItem.ToPane = 0


                CType(oItem.Specific, SAPbouiCOM.ComboBox).DataBind.SetBound(True, "OPOR", "U_EXO_EMPRESAD")
                CType(oItem.Specific, SAPbouiCOM.ComboBox).ExpandType = BoExpandType.et_DescriptionOnly
                CType(oItem.Specific, SAPbouiCOM.ComboBox).Item.DisplayDesc = True

                oItem = oForm.Items.Add("lblEmpresa", BoFormItemTypes.it_STATIC)
                oItem.Top = oForm.Items.Item("70").Top + 15
                oItem.Left = oForm.Items.Item("15").Left
                oItem.Height = oForm.Items.Item("15").Height
                oItem.Width = oForm.Items.Item("15").Width
                oItem.LinkTo = "cmbEmpresa"
                oItem.FromPane = 0
                oItem.ToPane = 0
                CType(oItem.Specific, SAPbouiCOM.StaticText).Caption = "Empresa"

                CargaComboEmpresa(oForm)

            End If

            EventHandler_Form_Load = True


        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            'oForm.Freeze(False)
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
        End Try
    End Function

    Private Function EventHandler_ItemPressed_After(ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing

        EventHandler_ItemPressed_After = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)

            'If pVal.ItemUID = "1" Then
            '    If pVal.ActionSuccess = True Then
            '        oForm.DataSources.UserDataSources.Item("IsDraft").ValueEx = "N"
            '    End If
            'End If

            EventHandler_ItemPressed_After = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
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

            If pVal.ItemUID = "4" Then '
                'If pVal.ItemChanged = True Then
                CargaComboEmpresa(oForm)
                'oForm.Items.Item("cmbEmpresa").Click()
                'End If
            End If

            If pVal.ItemUID = "54" Then '
                'If pVal.ItemChanged = True Then
                CargaComboEmpresa(oForm)
                'CType(oForm.Items.Item("cmbEmpresa").Specific, SAPbouiCOM.ComboBox).Active = True
                'End If
            End If

            'If pVal.ItemUID = "17_U_E" Then '
            '    If pVal.ItemChanged = True Then
            '        'mirar si existe el codigo en la empresa de "nombre de la base de datos
            '        If CType(oForm.Items.Item("17_U_E").Specific, SAPbouiCOM.EditText).Value <> "" Then
            '            oRs = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
            '            sSql = "SELECT T0.""CardName"" FROM """ & CType(oForm.Items.Item("13_U_E").Specific, SAPbouiCOM.EditText).Value & """.""OCRD"" T0 WHERE T0.""CardType"" ='S' AND T0.""CardCode"" ='" & CType(oForm.Items.Item("17_U_E").Specific, SAPbouiCOM.EditText).Value & "' "
            '            oRs.DoQuery(sSql)
            '            If oRs.RecordCount > 0 Then
            '                CType(oForm.Items.Item("18_U_E").Specific, SAPbouiCOM.EditText).Value = oRs.Fields.Item("CardName").Value.ToString
            '            Else
            '                objGlobal.SBOApp.StatusBar.SetText("(EXO) - El código introducido no es válido en la empresa " & CType(oForm.Items.Item("13_U_E").Specific, SAPbouiCOM.EditText).Value & " .", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            '            End If
            '        Else
            '            CType(oForm.Items.Item("18_U_E").Specific, SAPbouiCOM.EditText).Value = ""
            '        End If
            '    End If
            'End If

            'If pVal.ItemUID = "19_U_E" Then 'CLIENTE
            '    If pVal.ItemChanged = True Then
            '        'mirar si existe el codigo en la empresa de "nombre de la base de datos
            '        If CType(oForm.Items.Item("19_U_E").Specific, SAPbouiCOM.EditText).Value <> "" Then

            '            oRs = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
            '            sSql = "SELECT T0.""CardName"" FROM """ & CType(oForm.Items.Item("13_U_E").Specific, SAPbouiCOM.EditText).Value & """.""OCRD"" T0 WHERE T0.""CardType"" ='C' AND T0.""CardCode"" ='" & CType(oForm.Items.Item("19_U_E").Specific, SAPbouiCOM.EditText).Value & "' "
            '            oRs.DoQuery(sSql)
            '            If oRs.RecordCount > 0 Then
            '                CType(oForm.Items.Item("20_U_E").Specific, SAPbouiCOM.EditText).Value = oRs.Fields.Item("CardName").Value.ToString
            '            Else
            '                objGlobal.SBOApp.StatusBar.SetText("(EXO) - El código introducido no es válido en la empresa " & CType(oForm.Items.Item("13_U_E").Specific, SAPbouiCOM.EditText).Value & " .", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            '            End If
            '        Else
            '            CType(oForm.Items.Item("20_U_E").Specific, SAPbouiCOM.EditText).Value = ""
            '        End If
            '    End If
            'End If
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

#Region "Métodos auxiliares"
    Private Function CargaComboEmpresa(ByRef oForm As SAPbouiCOM.Form) As Boolean

        CargaComboEmpresa = False
        Dim sSQL As String = ""
        Dim oRs As SAPbobsCOM.Recordset = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Try
            oForm.Freeze(True)
            sSQL = "SELECT ""Code"", ""U_EXO_NEMP"" from ""@EXO_CONFTRASDOC"" " _
                & " WHERE U_EXO_CPDO='" & CType(oForm.Items.Item("4").Specific, SAPbouiCOM.EditText).Value & "'" _
                & " UNION ALL SELECT '-1' , '' FROM ""DUMMY"" "

            oRs.DoQuery(sSQL)
            If oRs.RecordCount > 0 Then
                objGlobal.funcionesUI.cargaCombo(CType(oForm.Items.Item("cmbEmpresa").Specific, SAPbouiCOM.ComboBox).ValidValues, sSQL)
                'si es uno nuevo, lo asgiamos
                If oForm.Mode = BoFormMode.fm_ADD_MODE Then
                    CType(oForm.Items.Item("cmbEmpresa").Specific, SAPbouiCOM.ComboBox).Select(oRs.Fields.Item(0).Value.ToString, BoSearchKey.psk_ByValue)
                End If


            Else
                objGlobal.SBOApp.StatusBar.SetText("(EXO) - Por favor, antes de continuar, revise la Configuración de Empresas Destino.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End If
            CargaComboEmpresa = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            oForm.Freeze(False)
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
        End Try
    End Function
    Public Shared Sub Connect_Company_Destino(ByRef oCompanyDes As SAPbobsCOM.Company, ByRef oCompOrigen As SAPbobsCOM.Company, ByRef sUser As String, ByRef sPass As String, ByVal SBD As String)

        Try
            'Conectar DI SAP


            oCompanyDes = New SAPbobsCOM.Company
            oCompanyDes.language = SAPbobsCOM.BoSuppLangs.ln_Spanish
            oCompanyDes.Server = oCompOrigen.Server
            oCompanyDes.LicenseServer = oCompOrigen.LicenseServer
            oCompanyDes.UserName = sUser
            oCompanyDes.Password = sPass
            oCompanyDes.UseTrusted = False
            oCompanyDes.DbPassword = "Password2018" 'oCompOrigen.DbPassword
            oCompanyDes.DbUserName = oCompOrigen.DbUserName '"B1SLDUSER"
            oCompanyDes.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB
            ' oCompany.SLDServer = oCompOrigen.SLDServer
            oCompanyDes.CompanyDB = SBD
            'oLog.escribeMensaje("database:" & oCompany.CompanyDB, EXO_Log.EXO_Log.Tipo.advertencia)
            If oCompanyDes.Connect <> 0 Then
                Throw New System.Exception("Error en la conexión a la compañia:" & oCompanyDes.GetLastErrorDescription.Trim)

            End If


        Catch exCOM As System.Runtime.InteropServices.COMException

            Throw New System.Exception("Error en la conexión a la compañia:" & oCompanyDes.GetLastErrorDescription.Trim & " Error: " & exCOM.Message.ToString)
        Catch ex As Exception
            Throw New System.Exception("Error en la conexión a la compañia:" & oCompanyDes.GetLastErrorDescription.Trim & " Error: " & ex.Message.ToString)

        Finally

        End Try
    End Sub
    Public Shared Sub Disconnect_Company(ByRef oCompanyDes As SAPbobsCOM.Company)
        Try
            If Not oCompanyDes Is Nothing Then
                If oCompanyDes.Connected = True Then
                    oCompanyDes.Disconnect()
                End If
            End If

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            If oCompanyDes IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oCompanyDes)
            oCompanyDes = Nothing
        End Try
    End Sub
    'crear pedido en la otra base de datos
    Private Function CrearPedidoVenta(ByRef oForm As Form, ByRef sDocEntryPedC As String) As Boolean
        Dim oORDR As SAPbobsCOM.Documents = Nothing
        Dim sSQL As String = ""
        Dim oRsOPOR As SAPbobsCOM.Recordset = Nothing
        Dim oRsORDR As SAPbobsCOM.Recordset = Nothing

        Dim oCompanyDes As SAPbobsCOM.Company = Nothing
        Dim sCodCliDestino As String = ""
        Dim oXml As System.Xml.XmlDocument = New System.Xml.XmlDocument
        Dim oNodes As System.Xml.XmlNodeList = Nothing
        Dim oNode As System.Xml.XmlNode = Nothing
        Dim sError As String = ""
        Dim sDocEntry As String = ""
        Dim strExiste As String = ""
        Dim sDocNum As String = ""
        Dim sSubject As String = ""
        Dim sComen As String = ""

        Try
            CrearPedidoVenta = True
            oRsOPOR = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)

            'cabecera
            sSQL = "SELECT DISTINCT t0.""U_EXO_EMPRESAD"", t0.""DocEntry"", t0.""DocNum"", t0.""TaxDate"", t0.""CardCode"", t0.""CardName"",t0.""NumAtCard"", t1.""DiscPrcnt"" ""DtoCab""," _
               & " t1.""ItemCode"",t1.""Dscription"",t1.""Quantity"",t1.""Price"",t1.""DiscPrcnt"",t1.""WhsCode"",t1.""DocEntry"",t1.""LineNum"",T1.""UomCode"", t1.""InvQty"",t1.""UomCode"", t1.""ShipDate""," _
                & " t2.""U_EXO_USER"", t2.""U_EXO_CLAVE"", t2.""U_EXO_NBD"",t2.""U_EXO_CCOD"",t0.""U_EXO_DOCENTRYD"",t0.""U_EXO_TIPODOCD"" " _
                & " FROM ""OPOR"" t0" _
                & " INNER JOIN ""POR1"" t1 On t0.""DocEntry"" = t1.""DocEntry""" _
                & " INNER JOIN ""@EXO_CONFTRASDOC"" t2 On t0.""U_EXO_EMPRESAD"" = t2.""Code""" _
                & " where t0.""DocEntry"" = " & CInt(sDocEntryPedC) & ""
            oRsOPOR.DoQuery(sSQL)
            oXml.LoadXml(oRsOPOR.GetAsXML())
            oNodes = oXml.SelectNodes("//row")

            If oRsOPOR.RecordCount > 0 Then
                objGlobal.SBOApp.StatusBar.SetText("Conectando con empresa destino", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success)
                Connect_Company_Destino(oCompanyDes, objGlobal.compañia, oRsOPOR.Fields.Item("U_EXO_USER").Value.ToString, oRsOPOR.Fields.Item("U_EXO_CLAVE").Value.ToString, oRsOPOR.Fields.Item("U_EXO_NBD").Value.ToString)
                ' conexiones con la base de datos destino para hacer el pedido de venta
                oORDR = CType(oCompanyDes.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders), SAPbobsCOM.Documents)
                oRsORDR = CType(oCompanyDes.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
                sCodCliDestino = oRsOPOR.Fields.Item("U_EXO_CCOD").Value.ToString
                'comprobar si existe el documento
                sSQL = "SELECT ""DocEntry"" FROM """ & oRsOPOR.Fields.Item("U_EXO_EMPRESAD").Value.ToString & """.""ORDR"" WHERE ""U_EXO_DOCENTRYD""='" & CInt(sDocEntryPedC) & "' "
                oRsORDR.DoQuery(sSQL)
                If oRsORDR.RecordCount > 0 Then
                    strExiste = oRsORDR.Fields.Item("DocEntry").Value.ToString
                End If
                If strExiste <> "" Then
                    objGlobal.SBOApp.StatusBar.SetText("Modificando pedido en empresa destino", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success)
                    'comprobar que el pedido no este impreso, y no haya niguna línea servida
                    sSQL = "SELECT    t1.""DocEntry"",t1.""DocNum"",t0.""ItemCode"",t1.""Printed""  " _
                    & " FROM """ & oForm.DataSources.DBDataSources.Item("OPOR").GetValue("U_EXO_EMPRESAD", 0).ToString & """.""RDR1"" t0 " _
                    & " INNER JOIN """ & oForm.DataSources.DBDataSources.Item("OPOR").GetValue("U_EXO_EMPRESAD", 0).ToString & """.""ORDR"" t1 ON t0.""DocEntry"" = t1.""DocEntry"" " _
                    & "  WHERE t1.""U_EXO_DOCENTRYD""=" & oForm.DataSources.DBDataSources.Item("OPOR").GetValue("DocEntry", 0).ToString & " And ((t0.""Quantity"" <> t0.""OpenCreQty"")  " _
                    & " Or  t1.""Printed""<> 'N')"
                    oRsORDR.DoQuery(sSQL)
                    If oRsORDR.RecordCount = 0 Then

                        oORDR.GetByKey(CInt(strExiste))

                        oORDR.DiscountPercent = CDbl(oRsOPOR.Fields.Item("DtoCab").Value.ToString.Replace(".", ","))
                        'comenntarios
                        oORDR.Comments = "Empresa origen:  " & oRsOPOR.Fields.Item("U_EXO_EMPRESAD").Value.ToString & vbCrLf & " Tipo doc orig: " & "Pedido de compra " & oRsOPOR.Fields.Item("DocNum").Value.ToString
                        'docentry
                        oORDR.UserFields.Fields.Item("U_EXO_DOCENTRYD").Value = sDocEntryPedC
                        'tipo
                        oORDR.UserFields.Fields.Item("U_EXO_TIPODOCD").Value = 22
                        'bd
                        oORDR.UserFields.Fields.Item("U_EXO_EMPRESAD").Value = objGlobal.compañia.CompanyDB

                        'borrar las lineas
                        'recorrer lineas y borrar y las hago de nuevo
                        For i = 0 To oORDR.Lines.Count - 1
                            oORDR.Lines.SetCurrentLine(0)
                            oORDR.Lines.Delete()
                        Next

                        'crear lineas lineas
                        For i As Integer = 0 To oNodes.Count - 1
                            oNode = oNodes.Item(i)
                            If i <> 0 Then
                                oORDR.Lines.Add()
                            End If
                            oORDR.Lines.ItemCode = oNode.SelectSingleNode("ItemCode").InnerText
                            oORDR.Lines.ItemDescription = oNode.SelectSingleNode("Dscription").InnerText
                            oORDR.Lines.Quantity = CDbl(oNode.SelectSingleNode("Quantity").InnerText.ToString.Replace(".", ","))
                            'precio
                            oORDR.Lines.UnitPrice = CDbl(oNode.SelectSingleNode("Price").InnerText.ToString.Replace(".", ","))
                            'descuento
                            oORDR.Lines.DiscountPercent = CDbl(oNode.SelectSingleNode("DiscPrcnt").InnerText.ToString.Replace(".", ","))
                            'If CDbl(oNode.SelectSingleNode("DiscPrcnt").InnerText.ToString) = 0.0 Then
                            '    oORDR.Lines.DiscountPercent = CDbl(0.0)
                            'Else


                            'End If

                        Next
                    End If

                Else
                    objGlobal.SBOApp.StatusBar.SetText("Creando pedido en empresa destino", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success)
                    oORDR.CardCode = sCodCliDestino
                    oORDR.TaxDate = CDate(oRsOPOR.Fields.Item("TaxDate").Value.ToString)
                    oORDR.DocDueDate = CDate(oRsOPOR.Fields.Item("TaxDate").Value.ToString)
                    'comenntarios
                    oORDR.Comments = "Empresa origen:  " & oRsOPOR.Fields.Item("U_EXO_EMPRESAD").Value.ToString & vbCrLf & " Tipo doc orig: " & "Pedido de compra " & oRsOPOR.Fields.Item("DocNum").Value.ToString
                    'docentry
                    oORDR.UserFields.Fields.Item("U_EXO_DOCENTRYD").Value = sDocEntryPedC
                    'tipo
                    oORDR.UserFields.Fields.Item("U_EXO_TIPODOCD").Value = 22
                    'bd
                    oORDR.UserFields.Fields.Item("U_EXO_EMPRESAD").Value = objGlobal.compañia.CompanyDB

                    oORDR.DiscountPercent = CDbl(oRsOPOR.Fields.Item("DtoCab").Value.ToString.Replace(".", ","))
                    'lineas
                    For i As Integer = 0 To oNodes.Count - 1
                        oNode = oNodes.Item(i)
                        If i <> 0 Then
                            oORDR.Lines.Add()
                        End If
                        oORDR.Lines.ItemCode = oNode.SelectSingleNode("ItemCode").InnerText
                        oORDR.Lines.ItemDescription = oNode.SelectSingleNode("Dscription").InnerText
                        oORDR.Lines.Quantity = CDbl(oNode.SelectSingleNode("Quantity").InnerText.ToString.Replace(".", ","))
                        'precio
                        oORDR.Lines.UnitPrice = CDbl(oNode.SelectSingleNode("Price").InnerText.ToString.Replace(".", ","))
                        'descuento
                        oORDR.Lines.DiscountPercent = CDbl(oNode.SelectSingleNode("DiscPrcnt").InnerText.ToString.Replace(".", ","))
                    Next

                End If
                If strExiste = "" Then
                    If oORDR.Add() <> 0 Then
                        CrearPedidoVenta = False
                        sError = oCompanyDes.GetLastErrorCode.ToString & " / " & oCompanyDes.GetLastErrorDescription.Replace("'", "")
                        'error al crear el pedido de venta en la empresa destino
                        sComen = sError
                        ''Enviamos alerta a los usuarios que estén marcados en la ficha del usuario con el campo Alertas 
                        EnviarAlerta(oCompanyDes, oForm, sDocEntry, oForm.DataSources.DBDataSources.Item("OPOR").GetValue("U_EXO_EMPRESAD", 0).ToString, sDocNum, "17", sSubject, sComen)
                    Else
                        oCompanyDes.GetNewObjectCode(sDocEntry)
                        'udpate
                        sSQL = "UPDATE ""OPOR"" SET ""U_EXO_DOCENTRYD"" =" & sDocEntry & ", ""U_EXO_TIPODOCD""='17', ""U_EXO_EMPRESAD"" ='" & oRsOPOR.Fields.Item("U_EXO_EMPRESAD").Value.ToString & "' WHERE ""DocEntry""=" & sDocEntryPedC & " "
                        oRsOPOR.DoQuery(sSQL)

                        'Enviamos alerta a los usuarios que estén marcados en la ficha del usuario con el campo Alertas 
                        sComen = ""
                        EnviarAlerta(oCompanyDes, oForm, sDocEntry, oForm.DataSources.DBDataSources.Item("OPOR").GetValue("U_EXO_EMPRESAD", 0).ToString, sDocNum, "17", sSubject, sComen)
                    End If
                Else
                    If oORDR.Update() <> 0 Then
                        CrearPedidoVenta = False
                        sError = oCompanyDes.GetLastErrorCode.ToString & " / " & oCompanyDes.GetLastErrorDescription.Replace("'", "")
                        'error al crear el pedido de venta en la empresa destino

                        ''Enviamos alerta a los usuarios que estén marcados en la ficha del usuario con el campo Alertas 
                        'sSubject = "Pedido Compra Intercompany " & sDocNumQUI & " ha tenido un error"
                        'sTipo = "Pedido de compra Intercompany"
                        sComen = sError
                        objGlobal.SBOApp.StatusBar.SetText("Enviando alerta a empresa destino", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success)
                        EnviarAlerta(oCompanyDes, oForm, sDocEntry, oForm.DataSources.DBDataSources.Item("OPOR").GetValue("U_EXO_EMPRESAD", 0).ToString, sDocNum, "17", sSubject, sComen)
                    Else
                        oCompanyDes.GetNewObjectCode(sDocEntry)
                        'udpate
                        sSQL = "UPDATE ""OPOR"" SET ""U_EXO_DOCENTRYD"" =" & sDocEntry & ", ""U_EXO_TIPODOCD""='17', ""U_EXO_EMPRESAD"" ='" & oRsOPOR.Fields.Item("U_EXO_EMPRESAD").Value.ToString & "' WHERE ""DocEntry""=" & sDocEntryPedC & " "
                        oRsOPOR.DoQuery(sSQL)

                        'Enviamos alerta a los usuarios que estén marcados en la ficha del usuario con el campo Alertas 
                        'sDocNum = Conexiones.GetValueDB(oDBSAP, "ORDR WITH (NOLOCK)", "DocNum", "DocEntry = " & sDocEntry & "")
                        'sSubject = "Pedido Venta Intercompany " & sDocNum & " se ha registrado correctamente como pedido de cliente en SAP"

                        sComen = ""
                        objGlobal.SBOApp.StatusBar.SetText("Enviando alerta a empresa destino", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success)
                        EnviarAlerta(oCompanyDes, oForm, sDocEntry, oForm.DataSources.DBDataSources.Item("OPOR").GetValue("U_EXO_EMPRESAD", 0).ToString, sDocNum, "17", sSubject, sComen)
                    End If
                End If
                Disconnect_Company(oCompanyDes)

            End If
        Catch exCOM As System.Runtime.InteropServices.COMException
            CrearPedidoVenta = False
            Throw exCOM
        Catch ex As Exception
            CrearPedidoVenta = False
            Throw ex
        Finally
            Disconnect_Company(oCompanyDes)
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRsOPOR, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRsORDR, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oCompanyDes, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oORDR, Object))
        End Try

    End Function
    Function ComprobarDocDestino(ByRef oForm As Form) As Boolean
        Dim sSQL As String = ""
        Dim oRsORDR As SAPbobsCOM.Recordset = Nothing

        Try
            ComprobarDocDestino = True
            oRsORDR = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
            sSQL = "SELECT    ""t1"".""DocEntry"",""t1"".""DocNum"",""t0"".""ItemCode"",""t1"".""Printed""  " _
                    & " FROM """ & oForm.DataSources.DBDataSources.Item("OPOR").GetValue("U_EXO_EMPRESAD", 0).ToString & """.""RDR1"" ""t0"" " _
                    & " INNER JOIN """ & oForm.DataSources.DBDataSources.Item("OPOR").GetValue("U_EXO_EMPRESAD", 0).ToString & """.""ORDR"" ""t1"" ON ""t0"".""DocEntry"" = ""t1"".""DocEntry"" " _
                    & "  WHERE ""t1"".""U_EXO_DOCENTRYD""=" & oForm.DataSources.DBDataSources.Item("OPOR").GetValue("DocEntry", 0).ToString & " And ((""t0"".""Quantity"" <>""t0"".""OpenCreQty"")  " _
                    & " Or  ""t1"".""Printed""<> 'N')"
            oRsORDR.DoQuery(sSQL)
            If oRsORDR.RecordCount > 0 Then
                ComprobarDocDestino = False
                'no se puede modifcar, sino dar aviso y no dejar
                objGlobal.SBOApp.MessageBox("El documento destino tiene líneas servidas o está impreso, no se puede modificar")
            End If


        Catch ex As Exception
            ComprobarDocDestino = False
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRsORDR, Object))
        End Try

    End Function

    Private Function EnviarAlerta(ByRef oComDes As SAPbobsCOM.Company, ByRef oForm As Form, ByVal sDEntry As String, ByVal sEmpresa As String, ByVal sNumDoc As String, ByVal sObject As String, ByVal sSubject As String, ByVal sText As String) As Boolean
        Dim oCmpSrv As SAPbobsCOM.CompanyService = Nothing
        Dim oMessageService As SAPbobsCOM.MessagesService = Nothing
        Dim oMessage As SAPbobsCOM.Message = Nothing
        Dim pMessageDataColumns As SAPbobsCOM.MessageDataColumns = Nothing
        Dim pMessageDataColumn As SAPbobsCOM.MessageDataColumn = Nothing
        Dim oLines As SAPbobsCOM.MessageDataLines = Nothing
        Dim oLine As SAPbobsCOM.MessageDataLine = Nothing
        Dim oRecipientCollection As SAPbobsCOM.RecipientCollection = Nothing

        Dim oRs As SAPbobsCOM.Recordset = Nothing
        Dim sSQL As String = ""
        Dim oXml As System.Xml.XmlDocument = New System.Xml.XmlDocument
        Dim oNodes As System.Xml.XmlNodeList = Nothing
        Dim oNode As System.Xml.XmlNode = Nothing

        Try
            EnviarAlerta = True
            sSubject = "Traspaso Pedido Venta "
            oRs = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
            sSQL = "Select t1.""USER_CODE"" " &
                       "FROM """ & sEmpresa & """.""OUSR"" t1 " &
                       "WHERE ""U_EXO_AVISO""='Y'"
            oRs.DoQuery(sSQL)
            oXml.LoadXml(oRs.GetAsXML())
            oNodes = oXml.SelectNodes("//row")

            If oRs.RecordCount > 0 Then
                For i As Integer = 0 To oNodes.Count - 1
                    oNode = oNodes.Item(i)
                    oCmpSrv = oComDes.GetCompanyService

                    oMessageService = CType(oCmpSrv.GetBusinessService(SAPbobsCOM.ServiceTypes.MessagesService), SAPbobsCOM.MessagesService)
                    oMessage = CType(oMessageService.GetDataInterface(SAPbobsCOM.MessagesServiceDataInterfaces.msdiMessage), SAPbobsCOM.Message)
                    oMessage.Subject = sSubject
                    oMessage.Text = sText
                    oRecipientCollection = oMessage.RecipientCollection

                    oRecipientCollection.Add()
                    oRecipientCollection.Item(0).SendInternal = SAPbobsCOM.BoYesNoEnum.tYES
                    oRecipientCollection.Item(0).UserCode = oNode.SelectSingleNode("USER_CODE").InnerText

                    pMessageDataColumns = oMessage.MessageDataColumns

                    If sDEntry <> "" Then
                        pMessageDataColumn = pMessageDataColumns.Add()
                        pMessageDataColumn.ColumnName = "Número interno"
                        pMessageDataColumn.Link = SAPbobsCOM.BoYesNoEnum.tYES
                        oLines = pMessageDataColumn.MessageDataLines
                        oLine = oLines.Add()
                        oLine.Value = sDEntry
                        oLine.Object = "17" 'pedido de venta
                        oLine.ObjectKey = sDEntry
                    End If

                    oMessageService.SendMessage(oMessage)
                Next
            End If

        Catch exCOM As System.Runtime.InteropServices.COMException
            EnviarAlerta = False
            Throw exCOM
        Catch ex As Exception
            EnviarAlerta = False
            Throw ex
        Finally
            If oCmpSrv IsNot Nothing Then
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oCmpSrv)
                oCmpSrv = Nothing
            End If
            If pMessageDataColumns IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(pMessageDataColumns)
            If pMessageDataColumn IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(pMessageDataColumn)
            If oLines IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oLines)
            If oLine IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oLine)
            If oRecipientCollection IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oRecipientCollection)
            'If oCmpSrv IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oCmpSrv)
            If oMessageService IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oMessageService)
            If oMessage IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oMessage)
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))

        End Try
    End Function
#End Region

End Class
