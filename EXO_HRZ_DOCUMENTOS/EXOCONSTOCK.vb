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

Public Class EXOCONSTOCK
    Inherits EXO_UIAPI.EXO_DLLBase
    Public intNumLin As Integer
    Public sObjType As String
    Public SDocEntry As String
    Public Sub New(ByRef oObjGlobal As EXO_UIAPI.EXO_UIAPI, ByRef actualizar As Boolean, usaLicencia As Boolean, idAddOn As Integer)
        MyBase.New(oObjGlobal, actualizar, usaLicencia, idAddOn)


    End Sub

#Region "Inicialización"

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


    Public Overrides Function SBOApp_ItemEvent(infoEvento As ItemEvent) As Boolean
        Try
            If infoEvento.InnerEvent = False Then
                If infoEvento.BeforeAction = False Then
                    Select Case infoEvento.FormTypeEx
                        Case "EXOCONSTOCK"
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
                        Case "EXOCONSTOCK"

                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                                Case SAPbouiCOM.BoEventTypes.et_CLICK

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_VALIDATE

                                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN


                                Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK

                            End Select
                    End Select
                End If

            Else
                If infoEvento.BeforeAction = False Then
                    Select Case infoEvento.FormTypeEx
                        Case "EXOCONSTOCK"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_FORM_VISIBLE
                                    Dim oForm As SAPbouiCOM.Form = Nothing
                                    oForm = objGlobal.SBOApp.Forms.Item(infoEvento.FormUID)
                                    If oForm.Visible = True Then
                                        CargarGrid(oForm)
                                    End If

                                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD

                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                            End Select
                    End Select
                Else
                    Select Case infoEvento.FormTypeEx
                        Case "EXOCONSTOCK"

                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                                Case SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED

                                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD

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


            EventHandler_ItemPressed_After = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.Form(oForm)
        End Try
    End Function


#End Region

    Private Sub CargarGrid(ByVal oForm As SAPbouiCOM.Form)
        Dim sSql As String = ""
        Dim oColumnTxt As SAPbouiCOM.EditTextColumn = Nothing
        Dim oColumnChk As SAPbouiCOM.CheckBoxColumn = Nothing
        Dim dFechaD As Date = Now.Date : Dim sFechaD As String = ""
        Dim dFechaH As Date = Now.Date : Dim sFechaH As String = ""
        Try

            sSql = "SELECT T0.""ITEMCODE"" ""Número Artículo"", T1.""ItemName"" ""Descripción"", T0.""WHSCODE"" ""Código Almacen"" ,T2.""WhsName"" ""Nombre"", t0.""STOCKALM"" ""Stock"",T0.""STOCKTOT"" ""Stock Total""
            FROM ""STOCKALM"" T0
            LEFT OUTER JOIN ""OITM"" T1 ON T0.""ITEMCODE"" = T1.""ItemCode""
            INNER JOIN ""OWHS"" T2 ON T0.""WHSCODE"" = T2.""WhsCode"""
            sSql = sSql & " WHERE  T0.""OBJTYPE"" ='" & oForm.DataSources.UserDataSources.Item("UDTYPE").Value & "'"
            sSql = sSql & " AND  T0.""DOCNUM"" =" & oForm.DataSources.UserDataSources.Item("UDDOCNUM").Value & ""
            sSql = sSql & " AND  T0.""LINENUM"" =" & oForm.DataSources.UserDataSources.Item("UDLINENUM").Value & ""

            'where
            'oForm.DataSources.UserDataSources.Item("CodCli").Value
            oForm.DataSources.DataTables.Item("dtStock").ExecuteQuery(sSql)

        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        End Try
    End Sub

End Class
