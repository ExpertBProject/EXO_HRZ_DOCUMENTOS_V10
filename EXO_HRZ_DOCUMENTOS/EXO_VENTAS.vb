Imports SAPbouiCOM
Public Class EXO_VENTAS
    Inherits EXO_UIAPI.EXO_DLLBase
    Public Sub New(ByRef oObjGlobal As EXO_UIAPI.EXO_UIAPI, ByRef actualizar As Boolean, usaLicencia As Boolean, idAddOn As Integer)
        MyBase.New(oObjGlobal, actualizar, usaLicencia, idAddOn)
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

    Public Overrides Function SBOApp_FormDataEvent(ByVal infoEvento As BusinessObjectInfo) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim oXml As New Xml.XmlDocument
        Try
            If infoEvento.BeforeAction = True Then
                Select Case infoEvento.FormTypeEx
                    Case "149", "139", "140", "234234567", "180", "65300", "133", "60090", "179", "60091"
                        Select Case infoEvento.EventType

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE
                                oForm = objGlobal.SBOApp.Forms.Item(infoEvento.FormUID)
                                If ComprobarAlmacenes(oForm) = False Then
                                    Return False
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                                'antes de actualizar comprobar si el pedido es con destino
                                oForm = objGlobal.SBOApp.Forms.Item(infoEvento.FormUID)
                                If ComprobarAlmacenes(oForm) = False Then
                                    Return False
                                Else
                                    Return True
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE

                        End Select
                End Select
            Else
                Select Case infoEvento.FormTypeEx
                    Case "149", "139", "140", "234234567", "180", "65300", "133", "60090", "179", "60091"
                        Select Case infoEvento.EventType

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD

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

    Private Function ComprobarAlmacenes(ByRef oForm As SAPbouiCOM.Form) As Boolean
#Region "variables"
        Dim dtContador As System.Data.DataTable = New System.Data.DataTable
#End Region
        ComprobarAlmacenes = False
        Try
            dtContador.Columns.Add("Almacen", GetType(String))
            dtContador.Clear()

            MatrixToNet(oForm, dtContador)
            If dtContador.Rows.Count > 1 Then
                If objGlobal.SBOApp.MessageBox("Existen diferentes almacenes. ¿Está seguro de continuar?", 1, "Sí", "No") = 1 Then
                    ComprobarAlmacenes = True
                Else
                    ComprobarAlmacenes = False
                End If
            Else
                ComprobarAlmacenes = True
            End If

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally

        End Try
    End Function
    Private Function MatrixToNet(ByRef oForm As SAPbouiCOM.Form, ByRef dtContador As System.Data.DataTable) As Boolean
        Dim sXML As String = ""
        Dim oMatrixXML As New Xml.XmlDocument
        Dim oXmlListRow As Xml.XmlNodeList = Nothing
        Dim oXmlListColumn As Xml.XmlNodeList = Nothing
        Dim oXmlNodeField As Xml.XmlNode = Nothing
        Dim sWhsCode As String = ""
        Dim sMatrixUID As String = ""

        MatrixToNet = False

        Try
            sXML = CType(oForm.Items.Item("38").Specific, SAPbouiCOM.Matrix).SerializeAsXML(SAPbouiCOM.BoMatrixXmlSelect.mxs_All)
            oMatrixXML.LoadXml(sXML)

            sMatrixUID = oMatrixXML.SelectSingleNode("//Matrix/UniqueID").InnerText
            oXmlListRow = oMatrixXML.SelectNodes("//Matrix/Rows/Row")

            For Each oXmlNodeRow As Xml.XmlNode In oXmlListRow
                oXmlListColumn = oXmlNodeRow.SelectNodes("Columns/Column")

                'Inicializamos los datos del registro
                sWhsCode = ""

                For Each oXmlNodeColumn As Xml.XmlNode In oXmlListColumn
                    oXmlNodeField = oXmlNodeColumn.SelectSingleNode("ID")

                    If oXmlNodeField.InnerXml = "24" Then 'Almacén
                        oXmlNodeField = oXmlNodeColumn.SelectSingleNode("Value")

                        sWhsCode = oXmlNodeField.InnerText
                        If sWhsCode <> "" Then
                            Dim dataRows As DataRow() = dtContador.Select("Almacen='" & sWhsCode & "'")
                            If dataRows.Count = 0 Then
                                dtContador.Rows.Add(sWhsCode)
                            End If
                        End If
                    End If

                    If dtContador.Rows.Count > 1 Then
                        Exit For
                    End If
                Next
            Next

            MatrixToNet = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        End Try
    End Function
End Class
