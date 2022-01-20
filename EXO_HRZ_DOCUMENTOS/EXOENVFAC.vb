Imports System.Xml
Imports CrystalDecisions.CrystalReports.Engine
Imports SAPbobsCOM
Imports SAPbouiCOM
Imports System.Drawing
Imports System.Drawing.Printing
Imports System.IO
Imports EXO_HRZ_DOCUMENTOS.Extensions
Imports CrystalDecisions.Shared
Imports System.Text.RegularExpressions

Public Class EXOENVFAC
    Inherits EXO_UIAPI.EXO_DLLBase
    Public Sub New(ByRef oObjGlobal As EXO_UIAPI.EXO_UIAPI, ByRef actualizar As Boolean, usaLicencia As Boolean, idAddOn As Integer)
        MyBase.New(oObjGlobal, actualizar, usaLicencia, idAddOn)

        If actualizar Then
            cargaDatos()
            cargaAutorizaciones()
            ParametrizacionGeneral()
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

            sXML = objGlobal.funciones.leerEmbebido(Me.GetType(), "UT_EXO_INFORMEFACV.xml")
            objGlobal.refDi.comunes.LoadBDFromXML(sXML)
            objGlobal.SBOApp.StatusBar.SetText("Validando:  UT_EXO_INFORMEFACV", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)


            sXML = objGlobal.funciones.leerEmbebido(Me.GetType(), "UDFs_OUSR.xml")
            objGlobal.refDi.comunes.LoadBDFromXML(sXML)
            objGlobal.SBOApp.StatusBar.SetText("Validando:  UDFs_OUSR", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            res = objGlobal.SBOApp.GetLastBatchResults
        End If
    End Sub
    Private Sub cargaAutorizaciones()
        Dim autorizacionXML As String = objGlobal.funciones.leerEmbebido(Me.GetType(), "EXO_AUEXOENVFAC.xml")
        objGlobal.refDi.comunes.LoadBDFromXML(autorizacionXML)
        Dim res As String = objGlobal.SBOApp.GetLastBatchResults
    End Sub

    Private Sub ParametrizacionGeneral()
        'If Not objGlobal.refDi.OGEN.existeVariable("EXO_RUTA_IMPFACTURAS") Then
        '    objGlobal.refDi.OGEN.fijarValorVariable("EXO_RUTA_IMPFACTURAS", "")
        'End If
        'If Not objGlobal.refDi.OGEN.existeVariable("EXO_LAYOUT_FACTURA") Then
        '    objGlobal.refDi.OGEN.fijarValorVariable("EXO_LAYOUT_FACTURA", "")
        'End If
        If Not objGlobal.refDi.OGEN.existeVariable("EXO_FORMATOFAC") Then
            objGlobal.refDi.OGEN.fijarValorVariable("EXO_FORMATOFAC", "")
        End If

        If Not objGlobal.refDi.OGEN.existeVariable("EXO_CONFEMAIL") Then
            objGlobal.refDi.OGEN.fijarValorVariable("EXO_CONFEMAIL", "")
        End If


    End Sub

#End Region
#Region "Eventos"


    Public Overrides Function SBOApp_ItemEvent(infoEvento As ItemEvent) As Boolean
        Try
            If infoEvento.InnerEvent = False Then
                If infoEvento.BeforeAction = False Then
                    Select Case infoEvento.FormTypeEx
                        Case "EXOENVFAC"
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
                        Case "EXOENVFAC"

                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                                Case SAPbouiCOM.BoEventTypes.et_CLICK

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_VALIDATE

                                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                                Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK
                                    If EventHandler_DoubleClick_Before(infoEvento) = False Then
                                        GC.Collect()
                                        Return False
                                    End If
                            End Select
                    End Select
                End If

            Else
                If infoEvento.BeforeAction = False Then
                    Select Case infoEvento.FormTypeEx
                        Case "EXOENVFAC"
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
                        Case "EXOENVFAC"

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
                    Case "EXO-MnInfFacV"
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


            If pVal.ItemUID = "txtCodCli" Then '
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
            If pVal.ItemUID = "txtNom" Then '
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

                Case "CFLCli" 'CLIENTE
                    sCFL_ID = oCFLEvento.ChooseFromListUID
                    oDataTable = oCFLEvento.SelectedObjects
                    If oDataTable IsNot Nothing Then
                        If pVal.ItemUID = "txtCodCli" Then 'cliente
                            'Recuperamos el formulario de origen
                            Try
                                CType(oForm.Items.Item("txtCodCli").Specific, SAPbouiCOM.EditText).Value = oDataTable.GetValue("CardCode", 0).ToString
                            Catch ex As Exception

                            End Try
                            Try
                                CType(oForm.Items.Item("txtNom").Specific, SAPbouiCOM.EditText).Value = oDataTable.GetValue("CardName", 0).ToString
                            Catch ex As Exception

                            End Try



                        End If
                    End If

                Case "CFLCliN" 'CLIENTE
                    sCFL_ID = oCFLEvento.ChooseFromListUID
                    oDataTable = oCFLEvento.SelectedObjects
                    If oDataTable IsNot Nothing Then
                        If pVal.ItemUID = "txtNom" Then 'cliente
                            'Recuperamos el formulario de origen
                            Try
                                CType(oForm.Items.Item("txtNom").Specific, SAPbouiCOM.EditText).Value = oDataTable.GetValue("CardName", 0).ToString
                            Catch ex As Exception

                            End Try

                            Try
                                CType(oForm.Items.Item("txtCodCli").Specific, SAPbouiCOM.EditText).Value = oDataTable.GetValue("CardCode", 0).ToString
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

        EventHandler_Validate_After = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)



            If pVal.ItemUID = "txtCodCli" Then '
                If pVal.ItemChanged = True Then

                    If CType(oForm.Items.Item("txtCodCli").Specific, SAPbouiCOM.EditText).Value = "" Then

                        CType(oForm.Items.Item("txtNom").Specific, SAPbouiCOM.EditText).Value = ""
                    End If
                End If
            End If

            If pVal.ItemUID = "txtNom" Then 'CLIENTE
                If pVal.ItemChanged = True Then

                    If CType(oForm.Items.Item("txtNom").Specific, SAPbouiCOM.EditText).Value = "" Then

                        CType(oForm.Items.Item("txtCodCli").Specific, SAPbouiCOM.EditText).Value = ""
                    End If
                End If
            End If
            EventHandler_Validate_After = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally

            EXO_CleanCOM.CLiberaCOM.Form(oForm)
        End Try
    End Function
    Private Function EventHandler_ItemPressed_After(ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing

        EventHandler_ItemPressed_After = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)

            If pVal.ItemUID = "btCon" Then
                If pVal.ActionSuccess = True Then
                    CargarGrid(oForm)
                End If
            End If

            If pVal.ItemUID = "btProcesar" Then
                If pVal.ActionSuccess = True Then
                    ProcesarFac(oForm)
                End If
            End If

            EventHandler_ItemPressed_After = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.Form(oForm)
        End Try
    End Function
    Private Function EventHandler_DoubleClick_Before(ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim sValor As String = ""
        EventHandler_DoubleClick_Before = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)

            If pVal.ItemUID = "gridDoc" AndAlso pVal.ColUID = "Procesar" Then
                If CType(oForm.Items.Item("gridDoc").Specific, SAPbouiCOM.Grid).DataTable.GetValue("Procesar", 0).ToString = "Y" Then
                    sValor = "N"
                    objGlobal.SBOApp.StatusBar.SetText("....Desmarcando selección .... ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                Else
                    sValor = "Y"
                    objGlobal.SBOApp.StatusBar.SetText("....Marcando selección .... ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                End If

                ''RECORRO Y CAMBIO
                oForm.Freeze(True)
                For i As Integer = 0 To oForm.DataSources.DataTables.Item("dtDoc").Rows.Count - 1
                    CType(oForm.Items.Item("gridDoc").Specific, SAPbouiCOM.Grid).DataTable.SetValue("Procesar", i, sValor)
                Next
                oForm.Freeze(False)
            End If

            EventHandler_DoubleClick_Before = True
            objGlobal.SBOApp.StatusBar.SetText("Finalizado ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            oForm.Freeze(False)
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
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
            oFP.XmlData = objGlobal.leerEmbebido(Me.GetType(), "EXOENVFAC.srf")

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
            'CargarComboRPT(oForm)
            CType(oForm.Items.Item("chkImp").Specific, SAPbouiCOM.CheckBox).Checked = True
            CType(oForm.Items.Item("chkCance").Specific, SAPbouiCOM.CheckBox).Checked = True

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
    'Private Sub CargarComboRPTSAP(ByRef oForm As SAPbouiCOM.Form)
    '    Dim sSQL As String = ""
    '    Dim oRs As SAPbobsCOM.Recordset = Nothing
    '    Dim sValor As String = ""

    '    Try
    '        oRs = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)

    '        'combo facturas
    '        sSQL = "SELECT ""DocCode"",""DocName"" FROM ""RDOC"" WHERE ""TypeCode""='INV2'"
    '        oRs.DoQuery(sSQL)
    '        If oRs.RecordCount > 0 Then
    '            While Not oRs.EoF
    '                CType(oForm.Items.Item("cmbFac").Specific, SAPbouiCOM.ComboBox).ValidValues.Add(oRs.Fields.Item("DocCode").Value.ToString, oRs.Fields.Item("DocName").Value.ToString)
    '                oRs.MoveNext()
    '            End While
    '        End If
    '        CType(oForm.Items.Item("cmbFac").Specific, SAPbouiCOM.ComboBox).ExpandType = BoExpandType.et_DescriptionOnly

    '    Catch ex As Exception

    '        EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
    '    End Try
    'End Sub
    'Private Sub CargarComboRPT(ByRef oForm As SAPbouiCOM.Form)
    '    Dim sRuta = ""
    '    Dim intPos As Integer
    '    Dim strError As String = ""


    '    Try

    '        sRuta = objGlobal.funcionesUI.refDi.OGEN.valorVariable("EXO_RUTA_IMPFACTURAS")
    '        Dim fileNames = My.Computer.FileSystem.GetFiles(sRuta, FileIO.SearchOption.SearchTopLevelOnly, "*.rpt")
    '        If sRuta = "" Then
    '            objGlobal.SBOApp.MessageBox("Introduzca ruta informe de facturas en la parametrizacion general de Expert One")
    '        Else
    '            'CType(oForm.Items.Item("cmbFac").Specific, SAPbouiCOM.ComboBox).ValidValues.Add("Factura Ventas Huryza", "Factura Ventas Huryza99999")
    '            For Each fileName As String In fileNames
    '                If (Path.GetFileName(fileName.ToString)).Contains("133") Then
    '                    If fileName.Contains("^") Then
    '                        intPos = InStr(fileName, "^", CompareMethod.Text)
    '                        CType(oForm.Items.Item("cmbFac").Specific, SAPbouiCOM.ComboBox).ValidValues.Add(fileName.ToString, Mid(fileName, intPos + 2, fileName.Length - 3))


    '                    End If
    '                End If
    '            Next
    '        End If

    '        CType(oForm.Items.Item("cmbFac").Specific, SAPbouiCOM.ComboBox).ExpandType = BoExpandType.et_DescriptionOnly
    '        'combo facturas
    '        'sSQL = "SELECT ""DocCode"",""DocName"" FROM ""RDOC"" WHERE ""TypeCode""='INV2'"
    '        'oRs.DoQuery(sSQL)
    '        'If oRs.RecordCount > 0 Then
    '        '    While Not oRs.EoF
    '        '        CType(oForm.Items.Item("cmbFac").Specific, SAPbouiCOM.ComboBox).ValidValues.Add(oRs.Fields.Item("DocCode").Value.ToString, oRs.Fields.Item("DocName").Value.ToString)
    '        '        oRs.MoveNext()
    '        '    End While
    '        'End If
    '        'CType(oForm.Items.Item("cmbFac").Specific, SAPbouiCOM.ComboBox).ExpandType = BoExpandType.et_DescriptionOnly

    '    Catch ex As Exception


    '    End Try
    'End Sub
    Private Sub CargarGrid(ByVal oForm As SAPbouiCOM.Form)
        Dim sSql As String = ""
        Dim oColumnTxt As SAPbouiCOM.EditTextColumn = Nothing
        Dim oColumnChk As SAPbouiCOM.CheckBoxColumn = Nothing
        Dim dFechaD As Date = Now.Date : Dim sFechaD As String = ""
        Dim dFechaH As Date = Now.Date : Dim sFechaH As String = ""
        Try
            'dFechaD = CDate(oForm.DataSources.UserDataSources.Item("FecD").ValueEx)
            'dFechaH = CDate(oForm.DataSources.UserDataSources.Item("FecH").ValueEx)
            'sFechaD = dFechaD.Year.ToString("0000") & dFechaD.Month.ToString("00") & dFechaD.Day.ToString("00")
            'sFechaH = dFechaH.Year.ToString("0000") & dFechaH.Month.ToString("00") & dFechaH.Day.ToString("00")
            '& " WHERE T0.""TaxDate"" >= '" & sFechaD & "' and T0.""TaxDate""<= '" & sFechaH & "' " _
            sSql = "SELECT * FROM (" _
                & " SELECT 'N' ""Procesar"", T0.""DocEntry"",T0.""DocNum"", T0.""TaxDate"", T0.""CardCode"", T0.""CardName"",T0.""DocTotal"", " _
                & " CASE WHEN T1.""U_stec_imp"" = 'S' THEN 'Y' ELSE 'N' END ""Imprimir""," _
                & " CASE WHEN T1.""U_stec_mai"" = 'S' THEN 'Y' ELSE 'N' END ""Email"" ," _
                & " T1.""E_Mail"" ""EnvEmail"" " _
                & " FROM ""OINV"" T0" _
                & " INNER JOIN ""OCRD"" T1 ON T0.""CardCode"" = T1.""CardCode"" " _
                & " WHERE TO_CHAR(COALESCE(T0.""TaxDate"", ''), 'YYYYMMDD') >= '" & oForm.DataSources.UserDataSources.Item("FecD").ValueEx & "' " _
            & " AND TO_CHAR(COALESCE(T0.""TaxDate"", ''), 'YYYYMMDD') <= '" & oForm.DataSources.UserDataSources.Item("FecH").ValueEx & "' "
            If oForm.DataSources.UserDataSources.Item("CodCli").Value <> "" Then
                sSql = sSql & " AND T0.""CardCode"" ='" & oForm.DataSources.UserDataSources.Item("CodCli").Value & "'"
            End If

            If CType(oForm.Items.Item("chkCance").Specific, SAPbouiCOM.CheckBox).Checked = True Then
                sSql = sSql & " and COALESCE(T0.""CANCELED"",'N')='N' "

            End If
            sSql = sSql & " And T1.""U_stec_imp"" = 'S' "

            If CType(oForm.Items.Item("chkImp").Specific, SAPbouiCOM.CheckBox).Checked = True Then
                sSql = sSql & " and COALESCE(T0.""Printed"",'N')='N' "
            End If

            sSql = sSql & " UNION ALL SELECT 'N' ""Procesar"", T0.""DocEntry"",T0.""DocNum"", T0.""TaxDate"", T0.""CardCode"", T0.""CardName"",T0.""DocTotal"", " _
                & " CASE WHEN T1.""U_stec_imp"" = 'S' THEN 'Y' ELSE 'N' END ""Imprimir""," _
                & " CASE WHEN T1.""U_stec_mai"" = 'S' THEN 'Y' ELSE 'N' END ""Email"" ," _
                & " T1.""E_Mail"" ""EnvEmail"" " _
                & " FROM ""OINV"" T0" _
                & " INNER JOIN ""OCRD"" T1 ON T0.""CardCode"" = T1.""CardCode"" " _
                & " WHERE TO_CHAR(COALESCE(T0.""TaxDate"", ''), 'YYYYMMDD') >= '" & oForm.DataSources.UserDataSources.Item("FecD").ValueEx & "' " _
            & " AND TO_CHAR(COALESCE(T0.""TaxDate"", ''), 'YYYYMMDD') <= '" & oForm.DataSources.UserDataSources.Item("FecH").ValueEx & "' "
            If oForm.DataSources.UserDataSources.Item("CodCli").Value <> "" Then
                sSql = sSql & " AND T0.""CardCode"" ='" & oForm.DataSources.UserDataSources.Item("CodCli").Value & "'"
            End If
            If CType(oForm.Items.Item("chkCance").Specific, SAPbouiCOM.CheckBox).Checked = True Then
                sSql = sSql & " and COALESCE(T0.""CANCELED"",'N')='N' "

            End If
            sSql = sSql & " And T1.""U_stec_mai"" = 'S' ) "

            sSql = sSql & " ORDER BY  ""CardCode"" ,""TaxDate"""
            oForm.DataSources.DataTables.Item("dtDoc").ExecuteQuery(sSql)
            If CType(oForm.Items.Item("gridDoc").Specific, SAPbouiCOM.Grid).Columns.Count > 0 Then
                CType(oForm.Items.Item("gridDoc").Specific, SAPbouiCOM.Grid).Columns.Item(0).Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
                CType(oForm.Items.Item("gridDoc").Specific, SAPbouiCOM.Grid).Columns.Item(0).AffectsFormMode = False

                CType(oForm.Items.Item("gridDoc").Specific, SAPbouiCOM.Grid).Columns.Item(7).Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
                CType(oForm.Items.Item("gridDoc").Specific, SAPbouiCOM.Grid).Columns.Item(7).AffectsFormMode = False

                CType(oForm.Items.Item("gridDoc").Specific, SAPbouiCOM.Grid).Columns.Item(8).Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
                CType(oForm.Items.Item("gridDoc").Specific, SAPbouiCOM.Grid).Columns.Item(8).AffectsFormMode = False

                'CType(oForm.Items.Item("gridDoc").Specific, SAPbouiCOM.Grid).Columns.Item(9).Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
                'CType(oForm.Items.Item("gridDoc").Specific, SAPbouiCOM.Grid).Columns.Item(9).AffectsFormMode = False

                oColumnChk = CType(CType(oForm.Items.Item("gridDoc").Specific, SAPbouiCOM.Grid).Columns.Item(0), SAPbouiCOM.CheckBoxColumn)
                oColumnChk.Editable = True
                oColumnChk.TitleObject.Caption = "Procesar"
                oColumnChk.TitleObject.Sortable = True

                oColumnTxt = CType(CType(oForm.Items.Item("gridDoc").Specific, SAPbouiCOM.Grid).Columns.Item(1), SAPbouiCOM.EditTextColumn)
                oColumnTxt.Editable = False
                oColumnTxt.TitleObject.Caption = "Número Interno"
                oColumnTxt.TitleObject.Sortable = True
                oColumnTxt.LinkedObjectType = "13"

                oColumnTxt = CType(CType(oForm.Items.Item("gridDoc").Specific, SAPbouiCOM.Grid).Columns.Item(2), SAPbouiCOM.EditTextColumn)
                oColumnTxt.Editable = False
                oColumnTxt.TitleObject.Caption = "Número Documento"
                oColumnTxt.TitleObject.Sortable = True

                oColumnTxt = CType(CType(oForm.Items.Item("gridDoc").Specific, SAPbouiCOM.Grid).Columns.Item(3), SAPbouiCOM.EditTextColumn)
                oColumnTxt.Editable = False
                oColumnTxt.TitleObject.Caption = "Fecha Documento"
                oColumnTxt.TitleObject.Sortable = True

                oColumnTxt = CType(CType(oForm.Items.Item("gridDoc").Specific, SAPbouiCOM.Grid).Columns.Item(4), SAPbouiCOM.EditTextColumn)
                oColumnTxt.Editable = False
                oColumnTxt.TitleObject.Caption = "Código Cliente"
                oColumnTxt.TitleObject.Sortable = True
                oColumnTxt.LinkedObjectType = "2"

                oColumnTxt = CType(CType(oForm.Items.Item("gridDoc").Specific, SAPbouiCOM.Grid).Columns.Item(5), SAPbouiCOM.EditTextColumn)
                oColumnTxt.Editable = False
                oColumnTxt.TitleObject.Caption = "Nombre Cliente"
                oColumnTxt.TitleObject.Sortable = True
                oColumnTxt.LinkedObjectType = "2"

                oColumnTxt = CType(CType(oForm.Items.Item("gridDoc").Specific, SAPbouiCOM.Grid).Columns.Item(6), SAPbouiCOM.EditTextColumn)
                oColumnTxt.Editable = False
                oColumnTxt.TitleObject.Caption = "Total Documento "
                oColumnTxt.TitleObject.Sortable = True

                oColumnChk = CType(CType(oForm.Items.Item("gridDoc").Specific, SAPbouiCOM.Grid).Columns.Item(7), SAPbouiCOM.CheckBoxColumn)
                oColumnChk.Editable = False
                oColumnChk.TitleObject.Caption = "Imprimir"
                oColumnChk.TitleObject.Sortable = True

                oColumnChk = CType(CType(oForm.Items.Item("gridDoc").Specific, SAPbouiCOM.Grid).Columns.Item(8), SAPbouiCOM.CheckBoxColumn)
                oColumnChk.Editable = False
                oColumnChk.TitleObject.Caption = "Envio Email"
                oColumnChk.TitleObject.Sortable = True

                oColumnTxt = CType(CType(oForm.Items.Item("gridDoc").Specific, SAPbouiCOM.Grid).Columns.Item(9), SAPbouiCOM.EditTextColumn)
                oColumnTxt.Editable = False
                oColumnTxt.TitleObject.Caption = "Email Facturas "
                oColumnTxt.TitleObject.Sortable = True
            End If
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        End Try
    End Sub

    Private Sub ProcesarFac(ByVal oForm As SAPbouiCOM.Form)

        Dim sSql As String = ""
        Dim oDt As System.Data.DataTable = Nothing
        Dim sOutFileName As String = ""
        Dim oReport As New CrystalDecisions.CrystalReports.Engine.ReportDocument()
        Dim sServer As String = ""
        Dim sBBDD As String = ""
        Dim sUser As String = ""
        Dim sPwd As String = ""
        Dim sDriver As String = ""
        Dim intImpr As Integer = 0
        Dim intEmail As Integer = 0
        Dim iRespuesta As Integer = 0
        Dim sConnection As String = ""
        Dim oLogonProps As NameValuePairs2 = Nothing

        Try
            For i As Integer = 0 To oForm.DataSources.DataTables.Item("dtDoc").Rows.Count - 1
                objGlobal.SBOApp.StatusBar.SetText("...Comprobando número de facturas seleccionadas.... " & i + 1 & " de " & oForm.DataSources.DataTables.Item("dtDoc").Rows.Count & "", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                If oForm.DataSources.DataTables.Item("dtDoc").GetValue("Imprimir", i).ToString = "Y" And oForm.DataSources.DataTables.Item("dtDoc").GetValue("Procesar", i).ToString = "Y" Then
                    intImpr = intImpr + 1
                End If

                If oForm.DataSources.DataTables.Item("dtDoc").GetValue("Email", i).ToString = "Y" And oForm.DataSources.DataTables.Item("dtDoc").GetValue("Procesar", i).ToString = "Y" Then
                    intEmail = intEmail + 1
                End If
            Next



            If intImpr > 0 Or intEmail > 0 Then

                iRespuesta = objGlobal.SBOApp.MessageBox("Se van a imprimir " & intImpr & " Facturas y se van a enviar por email " & intEmail & " Facturas " & vbCrLf & "¿Desea Continuar?", 2, "Ok", "Cancel")
                If iRespuesta = 2 Then
                    Exit Sub
                End If
                If oForm.DataSources.DataTables.Item("dtDoc").Rows.Count > 0 Then
                    objGlobal.SBOApp.StatusBar.SetText("...Estableciendo conexión con el impreso de facturas....", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

                    'abro la conexion con el crystal para que el proceso tarde menos
                    sOutFileName = objGlobal.funcionesUI.refDi.OGEN.valorVariable("EXO_FORMATOFAC")
                    oReport.Load(sOutFileName)
                    'objGlobal.SBOApp.StatusBar.SetText("...Fichero : " & sOutFileName, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                    'Establecemos las conexiones a la BBDD
                    'oReport.DataSourceConnections.Clear()
                    sServer = "192.168.0.97:30013" ' objGlobal.compañia.Server
                    'sServer = objGlobal.refDi.SQL.dameCadenaConexion.ToString
                    sBBDD = objGlobal.compañia.CompanyDB
                    sUser = objGlobal.refDi.SQL.usuarioSQL
                    sPwd = objGlobal.refDi.SQL.claveSQL


                    sDriver = "B1CRHPROXY"
                    sConnection = "DRIVER={" & sDriver & "};UID=" & sUser & ";PWD=" & sPwd & ";SERVERNODE=" & sServer & ";DATABASENAME=NDB;DATABASE=" & sBBDD & ";"


                    'Dim conrepor As CrystalDecisions.Shared.DataSourceConnections = oReport.DataSourceConnections
                    'conrepor(0).SetConnection(sServer, sBBDD, sUser, sPwd)


                    'objGlobal.SBOApp.StatusBar.SetText("...Connection " & sConnection & "....", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

                    oLogonProps = oReport.DataSourceConnections(0).LogonProperties
                    oLogonProps.Set("Provider", sDriver)
                    oLogonProps.Set("Connection String", sConnection)
                    oLogonProps.Set("Provider", sDriver)

                    'objGlobal.SBOApp.StatusBar.SetText("...Después de ologonpropos....", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

                    oReport.DataSourceConnections(0).SetLogonProperties(oLogonProps)
                    oReport.DataSourceConnections(0).SetConnection(sServer, sBBDD, False)

                    'objGlobal.SBOApp.StatusBar.SetText("...Después de set logon properties....", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                    For Each oSubReport As ReportDocument In oReport.Subreports
                        For Each oConnection As IConnectionInfo In oSubReport.DataSourceConnections
                            oConnection.SetConnection(sServer, sBBDD, False)
                            oConnection.SetLogon(sUser, sPwd)
                        Next
                    Next

                    'objGlobal.SBOApp.StatusBar.SetText("...Despues de preparar la conexión con el impreso....", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

                    'If Right(objGlobal.pathDLL, 6).ToUpper = "DLL_64" Then
                    '    sDriver = "{HDBODBC}"
                    'Else
                    '    sDriver = "{HDBODBC32}"
                    'End If
                    'oReport.ApplyNewServer(sDriver, sServer, sUser, sPwd, sBBDD)

                End If

                For i As Integer = 0 To oForm.DataSources.DataTables.Item("dtDoc").Rows.Count - 1
                    objGlobal.SBOApp.StatusBar.SetText("....Comprobando facturas marcadas para tratar.... " & i + 1 & " de " & oForm.DataSources.DataTables.Item("dtDoc").Rows.Count & "", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                    If oForm.DataSources.DataTables.Item("dtDoc").GetValue("Procesar", i).ToString = "Y" Then
                        'impresion directa
                        If oForm.DataSources.DataTables.Item("dtDoc").GetValue("Imprimir", i).ToString = "Y" Then
                            'sOutFileName = IO.Path.GetTempPath() & "Doc.rpt"
                            'GetCrystalReportFile(objGlobal.compañia, CType(oForm.Items.Item("cmbFac").Specific, SAPbouiCOM.ComboBox).Value, sOutFileName)
                            'sOutFileName = CType(oForm.Items.Item("cmbFac").Specific, SAPbouiCOM.ComboBox).Value
                            'objGlobal.SBOApp.StatusBar.SetText("....Impriendo factura .... ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                            sOutFileName = objGlobal.funcionesUI.refDi.OGEN.valorVariable("EXO_FORMATOFAC")
                            'objGlobal.SBOApp.StatusBar.SetText("....Formato factura .... " & sOutFileName & "", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                            Imprimir(objGlobal.compañia, sOutFileName, CInt(oForm.DataSources.DataTables.Item("dtDoc").GetValue("DocEntry", i).ToString), oReport)
                        End If

                        'pdf y envio de email
                        If oForm.DataSources.DataTables.Item("dtDoc").GetValue("Email", i).ToString = "Y" Then
                            sOutFileName = objGlobal.funcionesUI.refDi.OGEN.valorVariable("EXO_FORMATOFAC")

                            'objGlobal.SBOApp.StatusBar.SetText("Antes de generar pdf", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                            'envio email
                            'sOutFileName = IO.Path.GetTempPath() & "Doc.rpt"
                            'GetCrystalReportFile(objGlobal.compañia, CType(oForm.Items.Item("cmbFac").Specific, SAPbouiCOM.ComboBox).Value, sOutFileName)
                            ' objGlobal.SBOApp.StatusBar.SetText("....Generando factura .... ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                            Try
                                GenerarPDF(CInt(oForm.DataSources.DataTables.Item("dtDoc").GetValue("DocEntry", i).ToString), i, sOutFileName, oForm.DataSources.DataTables.Item("dtDoc").GetValue("EnvEmail", i).ToString, oForm.DataSources.DataTables.Item("dtDoc").GetValue("DocNum", i).ToString, oReport)
                            Catch ex As Exception

                            End Try
                        End If


                    End If
                Next
            End If
        Catch ex As Exception
            objGlobal.SBOApp.StatusBar.SetText("Error " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            oReport.Close()
            oReport.Dispose()
            GC.Collect()
        Finally
            oReport.Close()
            oReport.Dispose()
            GC.Collect()
        End Try
    End Sub
    Public Sub Imprimir(ByVal oCompany As SAPbobsCOM.Company, ByVal sRptFileName As String, ByVal iDocEntry As Integer, ByRef oReport As ReportDocument)
        Dim sDesImp As String = ""
        Dim NombreFichero As String = ""
        Dim pd As New PrintDocument()
        'Dim oCRReport As New CrystalDecisions.CrystalReports.Engine.ReportDocument()
        Dim sServer As String = ""
        Dim sBBDD As String = ""
        Dim sUser As String = ""
        Dim sPwd As String = ""
        Dim sDriver As String = ""
        Dim oRs As SAPbobsCOM.Recordset = Nothing
        Dim sSql As String = ""

        Try
            sDesImp = "Imprimir Factura"

            'oCRReport.Load(sRptFileName)
            'oReport.SetParameterValue("DOCKEY", iDocEntry)
            'oReport.SetParameterValue("OBJECTID", "13")

            oReport.SetParameterValue("DocKey@", iDocEntry)

            'sServer = objGlobal.compañia.Server
            'sBBDD = objGlobal.compañia.CompanyDB
            'sUser = objGlobal.compañia.DbUserName
            'sPwd = objGlobal.refDi.SQL.claveSQL

            'If Right(objGlobal.pathDLL, 6).ToUpper = "DLL_64" Then
            '    sDriver = "{B1CRHPROXY}"
            'Else
            '    sDriver = "{B1CRHPROXY32}"
            'End If



            'oCRReport.ApplyNewServer(sDriver, sServer, sUser, sPwd, sBBDD)
            oReport.PrintToPrinter(1, False, 0, 0)

            'ACTUALIZAR PRINTED LO HAGO POR QUERY PORQUE POR OBJETO TARDA MUCHO
            oRs = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
            sSql = "UPDATE ""OINV"" SET ""Printed"" ='Y',  ""U_stec_fimp"" ='" & DateTime.Now.ToString("yyyy-MM-dd") & "' ,""U_stec_himp"" ='" & DateTime.Now.Hour.ToString("00") + DateTime.Now.Minute.ToString("00") & "'  WHERE ""DocEntry"" =" & iDocEntry & ""
            oRs.DoQuery(sSql)

        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.SBOApp.StatusBar.SetText("Error: " & exCOM.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Throw exCOM
        Catch ex As Exception

            objGlobal.SBOApp.StatusBar.SetText("Error: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Throw ex
        Finally
            'oCRReport.Close()
            'oCRReport.Dispose()
            'GC.Collect()
        End Try

    End Sub
    Public Sub GetCrystalReportFile(ByVal oCompany As SAPbobsCOM.Company, ByVal sFormatoImp As String, ByVal sOutFileName As String)
        Dim oBlobParams As SAPbobsCOM.BlobParams = Nothing
        Dim oKeySegment As SAPbobsCOM.BlobTableKeySegment = Nothing
        Dim oBlob As SAPbobsCOM.Blob = Nothing
        Dim sContent As String = ""
        Dim obuff() As Byte = Nothing

        Try
            oBlobParams = CType(oCompany.GetCompanyService().GetDataInterface(SAPbobsCOM.CompanyServiceDataInterfaces.csdiBlobParams), SAPbobsCOM.BlobParams)

            oBlobParams.Table = "RDOC"
            oBlobParams.Field = "Template"

            oKeySegment = oBlobParams.BlobTableKeySegments.Add()
            oKeySegment.Name = "DocCode"
            oKeySegment.Value = objGlobal.funcionesUI.refDi.OGEN.valorVariable("EXO_LAYOUT_FACTURA")


            oBlob = oCompany.GetCompanyService().GetBlob(oBlobParams)
            sContent = oBlob.Content

            obuff = Convert.FromBase64String(sContent)

            Using oFile As New System.IO.FileStream(sOutFileName, System.IO.FileMode.Create)
                oFile.Write(obuff, 0, obuff.Length)

                oFile.Close()
            End Using

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oBlobParams, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oKeySegment, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oBlob, Object))
        End Try
    End Sub

    Public Sub GenerarPDF(ByVal iDocEntry As Integer, ByVal i As Integer, ByVal sRptFileName As String, ByVal sEmail As String, ByVal sNumFac As String, ByRef oReport As ReportDocument)

        'Dim oCRReport As New CrystalDecisions.CrystalReports.Engine.ReportDocument()
        Dim oFileDestino As CrystalDecisions.Shared.DiskFileDestinationOptions = Nothing
        Dim sServer As String = ""
        Dim sBBDD As String = ""
        Dim sUser As String = ""
        Dim sPwd As String = ""
        Dim sFileName As String = ""
        Dim sDocEntrys As String = ""
        Dim sDriver As String = ""
        Dim sError As String = ""
        Dim sSql As String = ""
        Dim oUserTable As SAPbobsCOM.UserTable

        Try


            'oCRReport.Load(sRptFileName)

            'Establecemos los parámetros para el report.
            oReport.SetParameterValue("DocKey@", iDocEntry)
            'oReport.SetParameterValue(2, "13") '"OBJECTID", "13")

            'objGlobal.SBOApp.StatusBar.SetText("...despues de oreport set ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            ''Establecemos las conexiones a la BBDD

            'sServer = objGlobal.compañia.Server
            'sBBDD = objGlobal.compañia.CompanyDB
            'sUser = objGlobal.compañia.DbUserName
            'sPwd = objGlobal.refDi.SQL.claveSQL


            'If Right(objGlobal.pathDLL, 6).ToUpper = "DLL_64" Then
            '    sDriver = "{B1CRHPROXY}"
            'Else
            '    sDriver = "{B1CRHPROXY32}"
            'End If
            'oCRReport.ApplyNewServer(sDriver, sServer, sUser, sPwd, sBBDD)

            'objGlobal.SBOApp.StatusBar.SetText("...ruta temporal:" & IO.Path.GetTempPath() & " ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            sFileName = IO.Path.GetTempPath() & "Factura_" & sNumFac & ".pdf"
            'objGlobal.SBOApp.StatusBar.SetText("...filename:" & sFileName & " ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

            oReport.ExportOptions.ExportFormatType = CrystalDecisions.Shared.ExportFormatType.PortableDocFormat

            oFileDestino = New CrystalDecisions.Shared.DiskFileDestinationOptions
            oFileDestino.DiskFileName = sFileName

            'Le pasamos al reporte el parámetro destino del reporte (ruta)
            oReport.ExportOptions.DestinationOptions = oFileDestino

            'Le indicamos que el reporte no es para mostrarse en pantalla, sino, que es para guardar en disco
            oReport.ExportOptions.ExportDestinationType = CrystalDecisions.Shared.ExportDestinationType.DiskFile

            ''Indicamos el formato de la página del reporte
            'oCRReport.PrintOptions.PaperOrientation = CrystalDecisions.Shared.PaperOrientation.Landscape

            'Finalmente exportamos el reporte a PDF

            oReport.Export()

            'Cerramos
            'oReport.Close()
            'oReport.Dispose()

            'enviar por email
            If sEmail = "" Then
                objGlobal.SBOApp.MessageBox("El cliente no tiene email's de envio de facturas")
                Exit Sub
            End If
            'objGlobal.SBOApp.StatusBar.SetText("...Enviando email...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)


            If EnvioEmail(sEmail, sFileName, iDocEntry, sError) = True Then

                oUserTable = objGlobal.compañia.UserTables.Item("EXO_INFORMEFACV")
                oUserTable.UserFields.Fields.Item("U_EXO_DOCN").Value = sNumFac
                oUserTable.UserFields.Fields.Item("U_EXO_MAIL").Value = sEmail
                oUserTable.UserFields.Fields.Item("U_EXO_ENVIADO").Value = "Y"
                oUserTable.UserFields.Fields.Item("U_EXO_FECHA").Value = CDate(Now.ToString("yyyy-MM-dd"))
                oUserTable.UserFields.Fields.Item("U_EXO_ERROR").Value = sError

                If oUserTable.Add() = 0 Then

                End If


            Else
                oUserTable = objGlobal.compañia.UserTables.Item("EXO_INFORMEFACV")
                oUserTable.UserFields.Fields.Item("U_EXO_DOCN").Value = sNumFac
                oUserTable.UserFields.Fields.Item("U_EXO_MAIL").Value = sEmail
                oUserTable.UserFields.Fields.Item("U_EXO_ENVIADO").Value = "N"
                oUserTable.UserFields.Fields.Item("U_EXO_FECHA").Value = CDate(Now.ToString("yyyy-MM-dd"))
                oUserTable.UserFields.Fields.Item("U_EXO_ERROR").Value = sError

                If oUserTable.Add() = 0 Then

                End If
            End If


            'borrar los ficheros temporales este codigo se descomenta cuando se firme
            My.Computer.FileSystem.DeleteFile(sFileName)

            'actualizar el campo generado PDF y ruta del PDF creado


        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
            Throw exCOM
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
            Throw ex
        Finally
            'oCRReport.Close()
            'oCRReport.Dispose()
            'GC.Collect()
        End Try
    End Sub
    Public Function IsValidEmail(ByVal email As String) As Boolean
        If email = String.Empty Then Return False
        ' Compruebo si el formato de la dirección es correcto.
        Dim re As Regex = New Regex("^([0-9a-zA-Z]([-\.\w]*[0-9a-zA-Z])*@([0-9a-zA-Z][-\w]*[0-9a-zA-Z]\.)+[a-zA-Z]{2,9})$")
        Dim m As Match = re.Match(email)
        Return (m.Captures.Count <> 0)
    End Function

    Public Function EnvioEmail(ByVal sEmailDestino As String, ByVal sRutaEnvio As String, ByVal iDocEntry As Integer, ByRef sError As String) As Boolean
        Dim oCorreo As System.Net.Mail.MailMessage = Nothing
        Dim oSmtp As System.Net.Mail.SmtpClient = Nothing
        Dim oAttach As System.Net.Mail.Attachment = Nothing

        Dim sSecure As String = ""
        Dim sSMTP As String = ""
        Dim sUserSMTP As String = ""
        Dim sPwdSMTP As String = ""
        Dim sPortSMTP As String = ""

        Dim sFrom As String = ""
        Dim sSendTo As String() = Nothing
        Dim sEnvioMail As String = ""


        Dim sTitulo As String = "Envio Factura"
        Dim sMensaje As String = "Adjuntamos formato de factura  "
        Dim oRs As SAPbobsCOM.Recordset = Nothing
        Dim sSql As String = ""

        Try
            sEnvioMail = objGlobal.refDi.OGEN.valorVariable("EXO_CONFEMAIL")
            Dim Datos() As String
            Dim Valor() As String
            If sEnvioMail = "" Then
                objGlobal.SBOApp.MessageBox("Introduzca la configuración del envio de email en la parametrizacion general de Expert One")
                Return False
            End If

            'TITULO Y CUERPO
            'ACTUALIZAR PRINTED LO HAGO POR QUERY PORQUE POR OBJETO TARDA MUCHO
            oRs = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
            sSql = "SELECT T0.""U_EXO_ASUNTO"", T0.""U_EXO_CUERPO"" FROM ""@EXO_CONFIG""  T0"
            oRs.DoQuery(sSql)
            If oRs.RecordCount > 0 Then
                sTitulo = oRs.Fields.Item("U_EXO_ASUNTO").Value.ToString
                sMensaje = oRs.Fields.Item("U_EXO_CUERPO").Value.ToString
            End If


            Datos = sEnvioMail.ToString.Split(CType(";", Char()))

            For i = 0 To 5
                Valor = Datos(i).ToString.Split(CType("=", Char()))
                Select Case i
                    Case 0
                        sSecure = Valor(1).ToString
                    Case 1
                        sSMTP = Valor(1).ToString
                    Case 2
                        sUserSMTP = Valor(1).ToString
                    Case 3
                        sPwdSMTP = Valor(1).ToString
                    Case 4
                        sPortSMTP = Valor(1).ToString
                    Case 5
                        sFrom = Valor(1).ToString
                End Select
            Next
            If IsValidEmail(sFrom.Trim) = False Then
                Dim sMensajeVal As String = "Nº Factura Interno: " & iDocEntry.ToString & ChrW(13) & ChrW(10) &
                "Dirección de correo electrónico no válida """ & sFrom.Trim & """, el correo debe tener el formato: nombre@dominio.com, " & ChrW(13) & ChrW(10) &
                    "por favor seleccione un correo valido."
                objGlobal.SBOApp.MessageBox(sMensajeVal)
                objGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensajeVal, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                Return False
            End If
            If sSecure = "N" Then
                oSmtp = New System.Net.Mail.SmtpClient
                oSmtp.Host = sSMTP
                oSmtp.Credentials = New System.Net.NetworkCredential(sUserSMTP, sPwdSMTP)
                oSmtp.EnableSsl = False
            Else
                oSmtp = New System.Net.Mail.SmtpClient
                oSmtp.Host = sSMTP

                If sPortSMTP <> "" Then
                    oSmtp.Port = CInt(sPortSMTP)
                End If

                oSmtp.Credentials = New System.Net.NetworkCredential(sUserSMTP, sPwdSMTP)
                oSmtp.EnableSsl = True
            End If

            'Instanciamos el objeto correo y asociamos el usuario de origen
            oCorreo = New System.Net.Mail.MailMessage()
            oCorreo.From = New System.Net.Mail.MailAddress(sFrom.Trim)
            'sEmailDestino = "shernandez@expertone.es"
            'Recuperamos la información del correo de envio
            If sEmailDestino <> "" Then
                sSendTo = Split(sEmailDestino, ";")
                For i As Integer = 0 To sSendTo.Length - 1
                    If sSendTo(i).Trim <> "" Then
                        If IsValidEmail(sSendTo(i).Trim) = False Then
                            Dim sMensajeVal As String = "Nº Factura Interno: " & iDocEntry.ToString & ChrW(13) & ChrW(10) &
                "Dirección de correo electrónico no válida """ & sSendTo(i).Trim & """, el correo debe tener el formato: nombre@dominio.com, " & ChrW(13) & ChrW(10) &
                    "por favor seleccione un correo valido."
                            objGlobal.SBOApp.MessageBox(sMensajeVal)
                            objGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensajeVal, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                            Return False
                        Else
                            oCorreo.To.Add(sSendTo(i).Trim)
                        End If
                    End If
                Next i
            End If

            ' TODO descomentar
            If sError <> "" Then
                oCorreo.To.Add(sFrom)
            Else
                'Ponemos en copia a Timac
                'oCorreo.To.Add("adrian.castano@timacagro.es")
                'oCorreo.CC.Add(sFrom)
            End If

            oCorreo.Subject = sTitulo
            oCorreo.Body = sMensaje
            oCorreo.IsBodyHtml = False
            oCorreo.Priority = System.Net.Mail.MailPriority.Normal

            'adjuntos
            oAttach = New System.Net.Mail.Attachment(sRutaEnvio)
            oCorreo.Attachments.Add(oAttach)
            oSmtp.Send(oCorreo)

            'ACTUALIZAR PRINTED LO HAGO POR QUERY PORQUE POR OBJETO TARDA MUCHO
            oRs = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
            sSql = "UPDATE ""OINV"" SET   ""U_stec_fmai"" ='" & DateTime.Now.ToString("yyyy-MM-dd") & "' ,""U_stec_himp"" ='" & DateTime.Now.Hour.ToString("00") + DateTime.Now.Minute.ToString("00") & "'  WHERE ""DocEntry"" =" & iDocEntry & ""
            oRs.DoQuery(sSql)
            Return True

        Catch exSMTP As System.Net.Mail.SmtpException
            EnvioEmail = False
            Throw exSMTP

        Catch exCOM As Runtime.InteropServices.COMException
            EnvioEmail = False
            Throw exCOM

        Catch ex As Exception
            EnvioEmail = False
            Throw ex

        Finally
            If oCorreo IsNot Nothing Then oCorreo.Dispose()
            oCorreo = Nothing
            If oAttach IsNot Nothing Then oAttach.Dispose()
            oAttach = Nothing
            If oSmtp IsNot Nothing Then oSmtp.Dispose()
            oSmtp = Nothing
        End Try
    End Function

End Class
