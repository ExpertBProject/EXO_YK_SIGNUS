Imports System.Xml
Imports RestSharp
Imports SAPbouiCOM
Imports System.Web.Script.Serialization
Public Class EXO_SIGNUSS
    Inherits EXO_Generales.EXO_DLLBase
    Public Sub New(ByRef general As EXO_Generales.EXO_General, actualizar As Boolean)
        MyBase.New(general, actualizar)
        cargamenu()

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
                contenidoXML = objGlobal.Functions.leerEmbebido(Me.GetType(), "EXO_UDO_SIGNUS.xml")
                objGlobal.conexionSAP.refCompañia.LoadBDFromXML(contenidoXML)
                objGlobal.conexionSAP.SBOApp.StatusBar.SetText("(EXO) - Validado EXO_UDO_SIGNUS", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)


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
    Private Sub cargamenu()
        Dim Path As String = ""
        Dim menuXML As String = objGlobal.Functions.leerEmbebido(Me.GetType(), "XML_MENU.xml")
        SboApp.LoadBatchActions(menuXML)
        Dim res As String = SboApp.GetLastBatchResults

        If SboApp.Menus.Exists("EXO-MnSIG") = True Then
            Path = objGlobal.conexionSAP.path & "\02.Menus"
            If Path <> "" Then
                If IO.File.Exists(Path & "\MnSIG.png") = True Then
                    SboApp.Menus.Item("EXO-MnSIG").Image = Path & "\MnSIG.png"
                End If
            End If
        End If
    End Sub
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
                    Case "EXO-MnSIGS"
                        'Cargamos pantalla de gestión.
                        If CargarFormCDOC() = False Then
                            Exit Function
                        End If
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
    Public Function CargarFormCDOC() As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim sSQL As String = ""
        Dim oRs As SAPbobsCOM.Recordset = CType(Me.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim Path As String = ""
        Dim oFP As SAPbouiCOM.FormCreationParams = Nothing
        Dim EXO_Xml As New EXO_Generales.EXO_XML(objGlobal.conexionSAP.refCompañia, objGlobal.conexionSAP.refSBOApp)

        CargarFormCDOC = False

        Try
            Path = objGlobal.conexionSAP.pathPantallas
            If Path = "" Then
                Return False
            End If

            oFP = CType(SboApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams), SAPbouiCOM.FormCreationParams)
            oFP.XmlData = objGlobal.conexionSAP.leerEmbebido(Me.GetType(), "EXO_SIGNUSS.srf")

            Try
                oForm = SboApp.Forms.AddEx(oFP)
            Catch ex As Exception
                If ex.Message.StartsWith("Form - already exists") = True Then
                    SboApp.StatusBar.SetText("El formulario ya está abierto.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Exit Function
                ElseIf ex.Message.StartsWith("Se produjo un error interno") = True Then 'Falta de autorización
                    Exit Function
                End If
            End Try
            CargaComboFormato(oForm)
            CType(oForm.Items.Item("txtCant").Specific, SAPbouiCOM.EditText).String = "0"
            If CType(oForm.Items.Item("cbTipo").Specific, SAPbouiCOM.ComboBox).ValidValues.Count > 1 Then
                CType(oForm.Items.Item("cbTipo").Specific, SAPbouiCOM.ComboBox).Select(0, BoSearchKey.psk_Index)
            Else
                CType(oForm.Items.Item("cbTipo").Specific, SAPbouiCOM.ComboBox).Select("--", BoSearchKey.psk_ByValue)
            End If
            CargarFormCDOC = True
        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.conexionSAP.Mostrar_Error(exCOM, EXO_Generales.EXO_SAP.EXO_TipoMensaje.Excepcion)
        Catch ex As Exception
            objGlobal.conexionSAP.Mostrar_Error(ex, EXO_Generales.EXO_SAP.EXO_TipoMensaje.Excepcion)
        Finally
            oForm.Visible = True
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
        End Try
    End Function
    Private Function CargaComboFormato(ByRef oForm As SAPbouiCOM.Form) As Boolean

        CargaComboFormato = False
        Dim sSQL As String = ""
        Dim oRs As SAPbobsCOM.Recordset = CType(Me.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Try
            oForm.Freeze(True)
            If objGlobal.conexionSAP.compañia.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                sSQL = "SELECT ""U_EXO_CAT"" ""Categoría"",(""Name"" + ' / ' + ""U_SEI_Desc"") ""Tipo""""Tipo""  "
                sSQL &= " FROM ""@SEITXECOVATORSP"" "
                sSQL &= " WHERE ifnull( ""U_EXO_CAT"",'')<>'' "
            Else
                sSQL = "SELECT ""U_EXO_CAT"" ""Categoría"", (""Name"" + ' / ' + ""U_SEI_Desc"") ""Tipo""  "
                sSQL &= " FROM ""@SEITXECOVATORSP"" "
                sSQL &= " WHERE isnull( ""U_EXO_CAT"",'')<>'' "
            End If

            oRs.DoQuery(sSQL)
            If oRs.RecordCount > 0 Then
                objGlobal.conexionSAP.refSBOApp.cargaCombo(CType(oForm.Items.Item("cbTipo").Specific, SAPbouiCOM.ComboBox).ValidValues, sSQL)
            Else
                objGlobal.conexionSAP.SBOApp.StatusBar.SetText("(EXO) - Por favor, antes de continuar, revise la parametrización.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End If
            CargaComboFormato = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            oForm.Freeze(False)
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
        End Try
    End Function
    Public Overrides Function SBOApp_ItemEvent(ByRef infoEvento As EXO_Generales.EXO_infoItemEvent) As Boolean
        Try
            If infoEvento.InnerEvent = False Then
                If infoEvento.BeforeAction = False Then
                    Select Case infoEvento.FormTypeEx
                        Case "EXO_SIGNUSS"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                    If EventHandler_ItemPressed_After(infoEvento) = False Then
                                        GC.Collect()
                                        Return False
                                    End If
                                Case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK

                                Case SAPbouiCOM.BoEventTypes.et_VALIDATE

                                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                                Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE

                            End Select
                    End Select
                ElseIf infoEvento.BeforeAction = True Then
                    Select Case infoEvento.FormTypeEx
                        Case "EXO_SIGNUSS"

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
                        Case "EXO_SIGNUSS"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD

                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                            End Select

                    End Select
                Else
                    Select Case infoEvento.FormTypeEx
                        Case "EXO_SIGNUSS"

                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                                Case SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED

                            End Select
                    End Select
                End If
            End If

            Return MyBase.SBOApp_ItemEvent(infoEvento)
        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.conexionSAP.Mostrar_Error(exCOM, EXO_Generales.EXO_SAP.EXO_TipoMensaje.Excepcion)
            Return False
        Catch ex As Exception
            objGlobal.conexionSAP.Mostrar_Error(ex, EXO_Generales.EXO_SAP.EXO_TipoMensaje.Excepcion)
            Return False
        End Try
    End Function
    Private Function EventHandler_ItemPressed_After(ByRef pVal As EXO_Generales.EXO_infoItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing

        EventHandler_ItemPressed_After = False

        Try
            oForm = SboApp.Forms.Item(pVal.FormUID)
            Select Case pVal.ItemUID
                Case "btnSol"
                    If Sol_Codigos(oForm) = False Then
                        GC.Collect()
                        Return False
                    End If
            End Select

            EventHandler_ItemPressed_After = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.conexionSAP.Mostrar_Error(exCOM, EXO_Generales.EXO_SAP.EXO_TipoMensaje.Excepcion)
        Catch ex As Exception
            objGlobal.conexionSAP.Mostrar_Error(ex, EXO_Generales.EXO_SAP.EXO_TipoMensaje.Excepcion)
        Finally
            oForm.Freeze(False)
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
        End Try
    End Function
    Private Function Sol_Codigos(ByRef oForm As SAPbouiCOM.Form) As Boolean
        Sol_Codigos = False

        Dim sSQL As String = ""
        Dim oRs As SAPbobsCOM.Recordset = Nothing
        Dim iCantidad As Integer = 0 : Dim sTipo As String = ""
        Dim sRespuesta As String = ""

        Try
            iCantidad = CInt(oForm.DataSources.UserDataSources.Item("UD_Cant").Value)
            sTipo = oForm.DataSources.UserDataSources.Item("UD_Tipo").Value.ToString
            If iCantidad > 0 Then
                objGlobal.conexionSAP.SBOApp.StatusBar.SetText("(EXO) - Solicitando códigos...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                CType(oForm.Items.Item("btnSol").Specific, SAPbouiCOM.Button).Item.Enabled = False
                oForm.Freeze(True)
                oRs = CType(Me.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
#Region "Solicitar códigos"
                'Buscamos datos de conexión
                sSQL = "SELECT * FROM ""@EXO_SIGNUSP"" WHERE Code='SIGNUS'"
                oRs.DoQuery(sSQL)
                If oRs.RecordCount > 0 Then
                    Dim sURL As String = oRs.Fields.Item("U_EXO_URL").Value.ToString
                    Dim sUSU As String = oRs.Fields.Item("U_EXO_COD").Value.ToString
                    Dim sPASS As String = oRs.Fields.Item("U_EXO_CLAVE").Value.ToString
                    oRs = Nothing
                    Dim client As RestClient = New RestClient(sURL)
                    client.Timeout = -1
                    Dim request As RestRequest = New RestRequest(Method.POST)
                    'Usuario:Password en base64
                    Dim string64 As String = EXO_FUNCIONES.EncodeStrToBase64(sUSU & ":" & sPASS)
                    request.AddHeader("Authorization", "Basic " & string64)
                    request.AddHeader("Content-Type", "application/x-www-form-urlencoded")
                    request.AddParameter("categoria", sTipo)
                    request.AddParameter("cantidad", iCantidad.ToString)
                    Dim response As IRestResponse = client.Execute(request)
                    If response.StatusCode = 200 Then
                        objGlobal.conexionSAP.SBOApp.StatusBar.SetText("(EXO) - Recepción OK.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                        sRespuesta = response.Content
                        Dim serializer As New JavaScriptSerializer()
                        Dim oCodigos As CodigosSIGNUM = New CodigosSIGNUM
                        oCodigos = serializer.Deserialize(Of CodigosSIGNUM)(sRespuesta)
                        'Guardamos códigos
                        If EXO_FUNCIONES.Guarda_codigos(oCodigos, Me.Company, objGlobal) = True Then
                            objGlobal.conexionSAP.SBOApp.StatusBar.SetText("(EXO) - Recuperación de códigos terminada.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                            objGlobal.conexionSAP.SBOApp.MessageBox("Recuperación de códigos terminada.")
                        End If
                    Else
                        objGlobal.conexionSAP.SBOApp.StatusBar.SetText("(EXO) - Recepción Errónea. Motivo: " & response.StatusDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        objGlobal.conexionSAP.SBOApp.MessageBox("Recepción Errónea. Se interrumpe el proceso. Motivo: " & response.StatusDescription)
                    End If
                Else
                    objGlobal.conexionSAP.SBOApp.StatusBar.SetText("(EXO) - No existen datos de conexión. Revise los parámetros.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    objGlobal.conexionSAP.SBOApp.MessageBox("No existen datos de conexión. Revise los parámetros.")
                End If
#End Region
            Else
                objGlobal.conexionSAP.SBOApp.StatusBar.SetText("(EXO) - La cantidad tiene que ser superior a 0", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                objGlobal.conexionSAP.SBOApp.MessageBox("Se ha cancelado la la solicitud. La cantidad tiene que ser superior a 0.")
            End If

            Sol_Codigos = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.conexionSAP.Mostrar_Error(exCOM, EXO_Generales.EXO_SAP.EXO_TipoMensaje.Excepcion)
        Catch ex As Exception
            objGlobal.conexionSAP.Mostrar_Error(ex, EXO_Generales.EXO_SAP.EXO_TipoMensaje.Excepcion)
        Finally
            oForm.Freeze(False)
            CType(oForm.Items.Item("btnSol").Specific, SAPbouiCOM.Button).Item.Enabled = True
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
        End Try
    End Function
End Class
