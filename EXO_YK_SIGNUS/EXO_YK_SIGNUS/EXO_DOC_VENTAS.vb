Imports System.Xml
Imports SAPbouiCOM
Public Class EXO_DOC_VENTAS
    Inherits EXO_Generales.EXO_DLLBase
    Public Sub New(ByRef general As EXO_Generales.EXO_General, actualizar As Boolean)
        MyBase.New(general, actualizar)
    End Sub
    Public Overrides Function menus() As XmlDocument
        Return Nothing
    End Function

    Public Overrides Function filtros() As EventFilters
        Dim filtrosXML As Xml.XmlDocument = New Xml.XmlDocument
        filtrosXML.LoadXml(objGlobal.Functions.leerEmbebido(Me.GetType(), "XML_FILTRO_SIGNUS.xml"))
        Dim filtro As SAPbouiCOM.EventFilters = New SAPbouiCOM.EventFilters()
        filtro.LoadFromXML(filtrosXML.OuterXml)

        Return filtro
    End Function
    Public Overrides Function SBOApp_FormDataEvent(ByRef infoEvento As EXO_Generales.EXO_BusinessObjectInfo) As Boolean
        Dim res As Boolean = True
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim oXml As New Xml.XmlDocument
        Dim sItemCode As String = ""
        oForm = SboApp.Forms.Item(infoEvento.FormUID)
        Try
            oForm.Freeze(True)
            If infoEvento.BeforeAction = True Then
                Select Case infoEvento.FormTypeEx
                    Case "140"
                        Select Case infoEvento.EventType

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD

                        End Select

                End Select

            Else
                Select Case infoEvento.FormTypeEx
                    Case "140", "180", "179"
                        Select Case infoEvento.EventType

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE
                                If infoEvento.ActionSuccess = True Then
                                    'Dim sDocEntry As String = ""
                                    'Select Case infoEvento.FormTypeEx
                                    '    Case "140" : sDocEntry = oForm.DataSources.DBDataSources.Item("ODLN").GetValue("DocEntry", 0).ToUpper
                                    '    Case "180" : sDocEntry = oForm.DataSources.DBDataSources.Item("ORDN").GetValue("DocEntry", 0).ToUpper
                                    '    Case "179" : sDocEntry = oForm.DataSources.DBDataSources.Item("ORIN").GetValue("DocEntry", 0).ToUpper
                                    'End Select

                                    'If IsNumeric(sDocEntry) Then
                                    '    If AsignaCodigos(oForm, sDocEntry) = False Then
                                    '        GC.Collect()
                                    '        Return False
                                    '    End If
                                    'End If
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                                If infoEvento.ActionSuccess = True Then
                                    Dim sDocEntry As String = ""
                                    Select Case infoEvento.FormTypeEx
                                        Case "140" : sDocEntry = oForm.DataSources.DBDataSources.Item("ODLN").GetValue("DocEntry", 0).ToUpper
                                        Case "180" : sDocEntry = oForm.DataSources.DBDataSources.Item("ORDN").GetValue("DocEntry", 0).ToUpper
                                        Case "179" : sDocEntry = oForm.DataSources.DBDataSources.Item("ORIN").GetValue("DocEntry", 0).ToUpper
                                    End Select

                                    If IsNumeric(sDocEntry) Then
                                        If AsignaCodigos(oForm, sDocEntry) = False Then
                                            GC.Collect()
                                            Return False
                                        End If
                                    End If
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD

                        End Select
                End Select
            End If
            Return res

        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.conexionSAP.Mostrar_Error(exCOM, EXO_Generales.EXO_SAP.EXO_TipoMensaje.Excepcion)
            Return False
        Catch ex As Exception
            objGlobal.conexionSAP.Mostrar_Error(ex, EXO_Generales.EXO_SAP.EXO_TipoMensaje.Excepcion)
            Return False
        Finally
            oForm.Freeze(False)
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
        End Try
    End Function
    Private Function AsignaCodigos(ByRef oForm As SAPbouiCOM.Form, ByVal sDocEntry As String) As Boolean
        AsignaCodigos = False

        Try
            If oForm.Visible = True Then
                If (oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE) Then
                    oForm.Freeze(True)
                    Select Case oForm.Type
                        Case "140"
                            If EXO_FUNCIONES.MatrixToNet(oForm, Me.Company, objGlobal, sDocEntry) = True Then
                                AsignaCodigos = True
                            End If
                        Case "180"
                            If EXO_FUNCIONES.AsignarDev(oForm, Me.Company, objGlobal, sDocEntry) = True Then
                                AsignaCodigos = True
                            End If
                        Case "179"
                            If EXO_FUNCIONES.AsignarAbo(oForm, Me.Company, objGlobal, sDocEntry) = True Then
                                AsignaCodigos = True
                            End If
                    End Select
                End If
            End If

        Catch ex As Exception
            Throw ex
        Finally
            oForm.Freeze(False)

        End Try
    End Function
End Class
