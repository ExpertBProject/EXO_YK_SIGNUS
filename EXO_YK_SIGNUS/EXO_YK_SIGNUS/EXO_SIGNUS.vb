Imports System.Xml
Imports SAPbouiCOM
Public Class EXO_SIGNUS
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
                contenidoXML = objGlobal.Functions.leerEmbebido(Me.GetType(), "UDF_SEITXECOVATORSP.xml")
                objGlobal.conexionSAP.refCompañia.LoadBDFromXML(contenidoXML)
                objGlobal.conexionSAP.SBOApp.StatusBar.SetText("(EXO) - Validado UDF_SEITXECOVATORSP", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                CargarDatos()
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
    Private Function CargarDatos() As Boolean
        CargarDatos = False
        Dim sSQL As String = ""
        Dim oRs As SAPbobsCOM.Recordset = CType(Me.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim sCodigo As String = "" : Dim sDesc As String = "" : Dim sCat As String = ""
        Dim oDI_COM As EXO_DIAPI.EXO_UDOEntity = Nothing 'Instancia del UDO para Insertar datos
        Try

            For i = 1 To 13
                Select Case i
                    Case 1 : sCodigo = "E5S" : sDesc = "S - Manutención, Macizo, Quad, Kart, Jardinería, otros (excepto agrícola, obra pública e industrial) > 100 kg" : sCat = "SE521"
                    Case 2 : sCodigo = "F1S" : sDesc = "S - Agrícola < 50 kg" : sCat = "SF121"
                    Case 3 : sCodigo = "F2S" : sDesc = "S - Agrícola ≥ 50 y < 100 kg" : sCat = "SF221"
                    Case 4 : sCodigo = "F3S" : sDesc = "S - Agrícola ≥ 100 y < 200 kg" : sCat = "SF321"
                    Case 5 : sCodigo = "F4S" : sDesc = "S - Agrícola ≥ 200 kg" : sCat = "SF421"
                    Case 6 : sCodigo = "G1S" : sDesc = "S - Obra Pública e Industrial < 50 kg" : sCat = "SG121"
                    Case 7 : sCodigo = "G2S" : sDesc = "S - Obra Pública e Industrial ≥ 50 y < 100 kg" : sCat = "SG221"
                    Case 8 : sCodigo = "G3S" : sDesc = "S - Obra Pública e Industrial ≥ 100 y < 500 kg" : sCat = "SG321"
                    Case 9 : sCodigo = "G4S" : sDesc = "S - Obra Pública e Industrial ≥ 500 y < 1000 kg" : sCat = "SG421"
                    Case 10 : sCodigo = "G5S" : sDesc = "S - Obra Pública e Industrial ≥ 1000 kg" : sCat = "SG521"
                    Case 11 : sCodigo = "G5AS" : sDesc = "S - Obra Pública e Industrial ≥ 1000 y < 2000 kg" : sCat = "SG5A1"
                    Case 12 : sCodigo = "G5BS" : sDesc = "S - Obra Pública e Industrial ≥ 2000 y < 3500 kg" : sCat = "SG5B1"
                    Case 13 : sCodigo = "G5BS" : sDesc = "S - Obra Pública e Industrial ≥ 3500 kg" : sCat = "SG5C1"
                End Select
                sSQL = "SELECT * FROM ""@SEITXECOVATORSP"" WHERE ""Code""='" & sCodigo & "' "
                oRs.DoQuery(sSQL)
                oDI_COM = New EXO_DIAPI.EXO_UDOEntity(objGlobal.conexionSAP.refCompañia, "SEITXECOVATORSP") 'UDO de Campos de SAP
                objGlobal.conexionSAP.SBOApp.StatusBar.SetText("(EXO) - Rellenando Tabla SEITXECOVATORSP...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                If oRs.RecordCount = 0 Then
                    sSQL = "INSERT INTO ""@SEITXECOVATORSP"" (""Code"",""Name"",""U_SEI_Desc"",""U_EXO_CAT"") "
                    sSQL &= " VALUES('" & sCodigo & "','" & sCodigo & "','" & sDesc & "','" & sCat & "')"
                    objGlobal.conexionSAP.SBOApp.StatusBar.SetText("(EXO) - Insertando código:" & sCodigo, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                Else
                    objGlobal.conexionSAP.SBOApp.StatusBar.SetText("(EXO) - Actualizando Tabla SEITXECOVATORSP... Código: " & sCodigo, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    sSQL = "UPDATE ""@SEITXECOVATORSP"" "
                    sSQL &= " SET ""U_SEI_Desc"" ='" & sDesc & "', ""U_EXO_CAT""='" & sCat & "' "
                    sSQL &= " WHERE ""Code""='" & sCodigo & "' "
                End If
                oRs.DoQuery(sSQL)
            Next

            objGlobal.conexionSAP.SBOApp.StatusBar.SetText("(EXO) - Tabla SEITXECOVATORSP cargada...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            CargarDatos = True

        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oDI_COM, Object))
        End Try
    End Function
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
                    Case "EXO-MnSIGC"
                        'Cargamos UDO Campos SAP.
                        objGlobal.conexionSAP.cargaFormUdoBD("EXO_SIGNUS")
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
            If infoEvento.FormTypeEx = "UDO_FT_EXO_SIGNUS" Then
                If infoEvento.InnerEvent = True Then
                    Select Case infoEvento.EventType
                        Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD

                        Case SAPbouiCOM.BoEventTypes.et_FORM_VISIBLE
                            If infoEvento.BeforeAction = False Then
                                'If EventHandler_Form_Visible(infoEvento) = False Then
                                '    GC.Collect()
                                '    Return False
                                'End If
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
End Class
