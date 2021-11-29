Imports System.Xml
Imports SAPbouiCOM

Public Class EXO_SIGNUSP
    Inherits EXO_Generales.EXO_DLLBase
    Public Sub New(ByRef general As EXO_Generales.EXO_General, actualizar As Boolean)
        MyBase.New(general, actualizar)

        If actualizar Then
            cargaCampos()
        End If
    End Sub
    Public Sub cargaCampos()
        If objGlobal.conexionSAP.esAdministrador() Then
            objGlobal.conexionSAP.escribeMensaje("El usuario es administrador")
            'Definicion descuentos financieros
            Dim contenidoXML As String


            Try
                contenidoXML = objGlobal.Functions.leerEmbebido(Me.GetType(), "EXO_UDO_SIGNUP.xml")
                objGlobal.conexionSAP.refCompañia.LoadBDFromXML(contenidoXML)
                objGlobal.conexionSAP.SBOApp.StatusBar.SetText("(EXO) - Validado EXO_UDO_SIGNUP", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                'Introducir los datos
                CargarDatos()

            Catch exCOM As System.Runtime.InteropServices.COMException
                objGlobal.conexionSAP.Mostrar_Error(exCOM, EXO_Generales.EXO_SAP.EXO_TipoMensaje.Excepcion)
            Catch ex As Exception
                objGlobal.conexionSAP.Mostrar_Error(ex, EXO_Generales.EXO_SAP.EXO_TipoMensaje.Excepcion)
            Finally

            End Try
        Else
            objGlobal.conexionSAP.escribeMensaje("(EXO) - El usuario NO es administrador")
        End If
    End Sub
    Private Function CargarDatos() As Boolean
        CargarDatos = False
        Dim sSQL As String = ""
        Dim oRs As SAPbobsCOM.Recordset = CType(Me.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim sPeriodo As String = ""
        Dim oDI_COM As EXO_DIAPI.EXO_UDOEntity = Nothing 'Instancia del UDO para Insertar datos
        Try
            oDI_COM = New EXO_DIAPI.EXO_UDOEntity(objGlobal.conexionSAP.refCompañia, "EXO_SIGNUSP") 'UDO de Campos de SAP
#Region "CAMPOSSAP"
            sSQL = "SELECT * FROM ""@EXO_SIGNUSP"" WHERE ""Code""='SIGNUS' "
            oRs.DoQuery(sSQL)
            If oRs.RecordCount = 0 Then

                objGlobal.conexionSAP.SBOApp.StatusBar.SetText("(EXO) - Rellenando Parametros de SIGNUS...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                oDI_COM.GetNew()
                oDI_COM.SetValue("Code") = "SIGNUS"
                oDI_COM.SetValue("CodEntry") = "99"
                oDI_COM.SetValue("Name") = "SIGNUS Parámetros"
                oDI_COM.SetValue("U_EXO_URL") = "https://indusapli.signus.es/api/rest/codigosSg"
                oDI_COM.SetValue("U_EXO_COD") = "P2870303"
                oDI_COM.SetValue("U_EXO_CLAVE") = "pru7009"
                If oDI_COM.UDO_Add = False Then
                    Throw New Exception("(EXO) - Error al añadir datos de Parámetros " & oDI_COM.GetLastError)
                End If
            End If
#End Region
            objGlobal.conexionSAP.SBOApp.StatusBar.SetText("(EXO) - Tabla parámetros cargados...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            CargarDatos = True

        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oDI_COM, Object))
        End Try
    End Function
    Public Overrides Function filtros() As EventFilters
        Dim filtrosXML As Xml.XmlDocument = New Xml.XmlDocument
        filtrosXML.LoadXml(objGlobal.Functions.leerEmbebido(Me.GetType(), "XML_FILTRO_SIGNUS.xml"))
        Dim filtro As SAPbouiCOM.EventFilters = New SAPbouiCOM.EventFilters()
        filtro.LoadFromXML(filtrosXML.OuterXml)

        Return filtro
    End Function

    Public Overrides Function menus() As XmlDocument
        Return Nothing
    End Function
    Public Overrides Function SBOApp_MenuEvent(ByRef infoEvento As EXO_Generales.EXO_MenuEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing

        Try
            If infoEvento.BeforeAction = True Then
            Else
                Select Case infoEvento.MenuUID
                    Case "EXO-MnSIGP"
                        'Cargamos UDO Campos SAP.
                        objGlobal.conexionSAP.cargaFormUdoBD("EXO_SIGNUSP")
                End Select
            End If

            Return MyBase.SBOApp_MenuEvent(infoEvento)

        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.conexionSAP.Mostrar_Error(exCOM, EXO_Generales.EXO_SAP.EXO_TipoMensaje.Excepcion)
            Return False
        Catch ex As Exception
            objGlobal.conexionSAP.Mostrar_Error(ex, EXO_Generales.EXO_SAP.EXO_TipoMensaje.Excepcion)
            Return False
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
        End Try
    End Function
    Public Overrides Function SBOApp_ItemEvent(ByRef infoEvento As EXO_Generales.EXO_infoItemEvent) As Boolean
        Dim res As Boolean = True
        Dim oForm As SAPbouiCOM.Form = SboApp.Forms.Item(infoEvento.FormUID)

        Dim Path As String = ""
        Dim XmlDoc As New System.Xml.XmlDocument
        Dim EXO_Functions As New EXO_BasicDLL.EXO_Generic_Forms_Functions(Me.objGlobal.conexionSAP)

        Try
            If infoEvento.FormTypeEx = "UDO_FT_EXO_SIGNUSP" Then
                If infoEvento.InnerEvent = True Then
                    Select Case infoEvento.EventType
                        Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD

                        Case SAPbouiCOM.BoEventTypes.et_FORM_VISIBLE
                            If infoEvento.BeforeAction = False Then
                                If EventHandler_Form_Visible(infoEvento) = False Then
                                    GC.Collect()
                                    Return False
                                End If
                            End If
                    End Select
                Else
                    Select Case infoEvento.EventType
                        Case BoEventTypes.et_COMBO_SELECT

                        Case SAPbouiCOM.BoEventTypes.et_VALIDATE

                        Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                        Case SAPbouiCOM.BoEventTypes.et_CLICK

                    End Select
                End If
            End If


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

        Return res
    End Function
    Private Function EventHandler_Form_Visible(ByRef pVal As EXO_Generales.EXO_infoItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim oRs As SAPbobsCOM.Recordset = CType(Me.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim sSQL As String = ""
        EventHandler_Form_Visible = False

        Try
            oForm = SboApp.Forms.Item(pVal.FormUID)

            If oForm.Visible = True Then
                sSQL = "SELECT * FROM ""@EXO_SIGNUSP"" "
                oRs.DoQuery(sSQL)
                If oRs.RecordCount > 0 Then
                    oForm.Mode = BoFormMode.fm_OK_MODE
                    If oForm.Mode <> SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        objGlobal.conexionSAP.SBOApp.ActivateMenuItem("1290") ' Ir al primer registro
                    End If
                Else
                    oForm.Mode = BoFormMode.fm_ADD_MODE
                End If
            End If

            EventHandler_Form_Visible = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.conexionSAP.Mostrar_Error(exCOM, EXO_Generales.EXO_SAP.EXO_TipoMensaje.Excepcion)
        Catch ex As Exception
            objGlobal.conexionSAP.Mostrar_Error(ex, EXO_Generales.EXO_SAP.EXO_TipoMensaje.Excepcion)
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
        End Try
    End Function
End Class
