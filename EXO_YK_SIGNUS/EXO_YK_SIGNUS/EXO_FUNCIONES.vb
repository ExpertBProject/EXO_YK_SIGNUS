Public Class EXO_FUNCIONES
    Public Shared Function EncodeStrToBase64(valor As String) As String
        Dim myByte As Byte() = System.Text.Encoding.UTF8.GetBytes(valor)
        Dim myBase64 As String = Convert.ToBase64String(myByte)
        Return myBase64
    End Function
    Public Shared Function Guarda_codigos(ByRef oCodigos As CodigosSIGNUM, ByRef oCompany As SAPbobsCOM.Company, ByRef oobjGlobal As EXO_Generales.EXO_General) As Boolean
        Guarda_codigos = False
        Dim sSQL As String = ""
        Dim oRs As SAPbobsCOM.Recordset = Nothing
        Dim oDI_COM As EXO_DIAPI.EXO_UDOEntity = Nothing 'Instancia del UDO para Insertar datos
        Dim sAnno As String = Now.Year.ToString("0000")
        Dim bActualiza As Boolean = False
        Try
            oDI_COM = New EXO_DIAPI.EXO_UDOEntity(oobjGlobal.conexionSAP.refCompañia, "EXO_SIGNUS") 'UDO de Campos de SAP
            oRs = CType(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
            'Buscamos si existe por año
            sSQL = "SELECT * FROM ""@EXO_SIGNUS"" WHERE Code='" & sAnno & "' "
            oRs.DoQuery(sSQL)
            If oRs.RecordCount > 0 Then
                'Apuntamos al UDO del año
                oDI_COM.GetByKey(sAnno)
                bActualiza = True
            Else
                'Creamos uno nuevo con el año 
                oDI_COM.GetNew()
                oDI_COM.SetValue("Code") = sAnno
            End If

            For Each Linea As data In oCodigos.data
                oDI_COM.GetNewChild("EXO_SIGNUSL")
                oDI_COM.SetValueChild("U_EXO_TIPO") = Linea.catAbrev
                oDI_COM.SetValueChild("U_EXO_COD") = Linea.codigo
                oDI_COM.SetValueChild("U_EXO_LOTE") = Linea.lote
            Next
            If bActualiza = True Then
                If oDI_COM.UDO_Update = False Then
                    Throw New Exception("(EXO) - " & oDI_COM.GetLastError)
                End If
            Else
                If oDI_COM.UDO_Add = False Then
                    Throw New Exception("(EXO) - " & oDI_COM.GetLastError)
                End If
            End If
            Guarda_codigos = True
        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oDI_COM, Object))
        End Try
    End Function
#Region "Métodos auxiliares"
    Public Shared Function MatrixToNet(ByRef oForm As SAPbouiCOM.Form, ByRef oCompany As SAPbobsCOM.Company, ByRef oobjglobal As EXO_Generales.EXO_General, ByVal sDocEntry As String) As Boolean
        Dim sXML As String = ""
        Dim oMatrixXML As New Xml.XmlDocument
        Dim oXmlListRow As Xml.XmlNodeList = Nothing
        Dim oXmlListColumn As Xml.XmlNodeList = Nothing
        Dim oXmlNodeField As Xml.XmlNode = Nothing
        Dim sItemCode As String = ""
        Dim sCantidad As String = "0"
        Dim sLinea As String = ""
        Dim oArrCampos() As Boolean = {False, False, False}
        Dim sMatrixUID As String = ""
        Dim sSQL As String = ""
        Dim oRs As SAPbobsCOM.Recordset = Nothing : Dim oRsArt As SAPbobsCOM.Recordset = Nothing
        MatrixToNet = False

        Try
            oRs = CType(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
            oRsArt = CType(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
            sXML = CType(oForm.Items.Item("38").Specific, SAPbouiCOM.Matrix).SerializeAsXML(SAPbouiCOM.BoMatrixXmlSelect.mxs_All)
            oMatrixXML.LoadXml(sXML)

            sMatrixUID = oMatrixXML.SelectSingleNode("//Matrix/UniqueID").InnerText
            oXmlListRow = oMatrixXML.SelectNodes("//Matrix/Rows/Row")

            For Each oXmlNodeRow As Xml.XmlNode In oXmlListRow
                oXmlListColumn = oXmlNodeRow.SelectNodes("Columns/Column")

                'Inicializamos para de dejar a False
                For i As Integer = 0 To 2
                    oArrCampos(i) = False
                Next

                'Inicializamos los datos del registro
                sItemCode = ""
                sCantidad = "0"
                sLinea = "0"
                For Each oXmlNodeColumn As Xml.XmlNode In oXmlListColumn
                    oXmlNodeField = oXmlNodeColumn.SelectSingleNode("ID")

                    If oXmlNodeField.InnerXml = "110" Then 'linea
                        oXmlNodeField = oXmlNodeColumn.SelectSingleNode("Value")
                        sLinea = oXmlNodeField.InnerText.ToString
                        oArrCampos(0) = True
                    ElseIf oXmlNodeField.InnerXml = "11" Then 'Cantidad
                        oXmlNodeField = oXmlNodeColumn.SelectSingleNode("Value")

                        If oXmlNodeField.InnerText = "" Then sCantidad = "0" Else sCantidad = oXmlNodeField.InnerText

                        oArrCampos(1) = True
                    ElseIf oXmlNodeField.InnerXml = "1" Then 'Código Articulo
                        oXmlNodeField = oXmlNodeColumn.SelectSingleNode("Value")

                        sItemCode = oXmlNodeField.InnerText

                        oArrCampos(2) = True
                    End If

                    If oArrCampos(0) = True AndAlso oArrCampos(1) = True AndAlso oArrCampos(2) = True Then
                        oArrCampos(0) = False : oArrCampos(1) = False : oArrCampos(2) = False
                        'tenemos que ver si es de los tipos de signus o no hacemos nada
                        If oobjglobal.conexionSAP.compañia.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                            sSQL = "Select I.""U_TXECOVATORSP"" FROM OITM I "
                            sSQL &= " INNER JOIN ""@SEITXECOVATORSP"" S On I.""U_TXECOVATORSP""=S.""Code"" "
                            sSQL &= " WHERE I.""Itemcode""='" & sItemCode & "' and ifnull(S.""U_EXO_CAT"",'')<>''"
                        Else
                            sSQL = "SELECT I.""U_TXECOVATORSP"" FROM OITM I "
                            sSQL &= " INNER JOIN ""@SEITXECOVATORSP"" S ON I.""U_TXECOVATORSP""=S.""Code"" "
                            sSQL &= " WHERE I.""Itemcode""='" & sItemCode & "' and isnull(S.""U_EXO_CAT"",'')<>''"
                        End If
                        oRsArt.DoQuery(sSQL)
                        If oRsArt.RecordCount > 0 Then
                            If oobjglobal.conexionSAP.compañia.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                                sSQL = "Select COUNT(""U_EXO_COD"") ""CantD"" "
                                sSQL &= " From ""OITM"" I INNER JOIN ""@EXO_SIGNUSL"" SL On I.""U_TXECOVATORSP"" =SL.""U_EXO_TIPO"" INNER JOIN ""@EXO_SIGNUS"" S On S.""Code""=SL.""Code"" "
                                sSQL &= " WHERE I.""Itemcode""='" & sItemCode & "' and ifnull(SL.U_EXO_DocEE,'')='' and S.Code='" & Now.Year.ToString("0000") & "' "
                            Else
                                sSQL = "SELECT COUNT(""U_EXO_COD"") ""CantD"" "
                                sSQL &= " From ""OITM"" I INNER JOIN ""@EXO_SIGNUSL"" SL ON I.""U_TXECOVATORSP"" =SL.""U_EXO_TIPO"" INNER JOIN ""@EXO_SIGNUS"" S ON S.""Code""=SL.""Code"" "
                                sSQL &= " WHERE I.""Itemcode""='" & sItemCode & "' and isnull(SL.U_EXO_DocEE,'')='' and S.Code='" & Now.Year.ToString("0000") & "' "
                            End If

                            oRs.DoQuery(sSQL)
                            If oRs.RecordCount > 0 Then
                                Dim iCantidadDisponible As Integer = CInt(oRs.Fields.Item("CantD").Value.ToString)
                                Dim icantidadLinea As Integer = CInt(sCantidad)
                                If icantidadLinea <= iCantidadDisponible Then
                                    'asignamos linea y entrega en signus
                                    For i = 1 To icantidadLinea
                                        sSQL = "UPDATE L "
                                        sSQL &= " SET ""U_EXO_DocEE""='" & sDocEntry & "', ""U_EXO_DocE""='" & oForm.DataSources.DBDataSources.Item("ODLN").GetValue("DocNum", 0).ToUpper & "',"
                                        sSQL &= " ""U_EXO_linE""='" & sLinea & "' "
                                        sSQL &= " FROM ""@EXO_SIGNUS"" C INNER JOIN ""@EXO_SIGNUSL"" L ON C.""Code""=L.""Code"" "
                                        sSQL &= " WHERE C.""Code""='" & Now.Year.ToString("0000") & "' and isnull(""U_EXO_DocEE"",'')='' and isnull(""U_EXO_linE"",'')='' "
                                        sSQL &= " and L.""LineId""=( Select MIN(""LineId"") ""Linea"" "
                                        sSQL &= " From ""OITM"" I INNER JOIN ""@EXO_SIGNUSL"" SL On I.""U_TXECOVATORSP"" =SL.""U_EXO_TIPO"" INNER JOIN ""@EXO_SIGNUS"" S On S.""Code""=SL.""Code"" "
                                        sSQL &= " WHERE I.""Itemcode""='" & sItemCode & "' and isnull(SL.U_EXO_DocEE,'')='' and S.Code='" & Now.Year.ToString("0000") & "') "
                                        oRs.DoQuery(sSQL)
                                    Next
                                Else
                                    oobjglobal.conexionSAP.SBOApp.StatusBar.SetText("(EXO) - SIGNUS: No tiene suficientes códigos asignados en el año actual. Solicite más códigos para " & sItemCode & ". Actualice la entrega", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                    oobjglobal.conexionSAP.SBOApp.MessageBox("SIGNUS: No tiene suficientes códigos asignados en el año actual. Solicite más códigos para " & sItemCode & ". Actualice la entrega")
                                End If
                            End If
                        Else
                            'NO ES CÓDIGO SIGNUS, NO HACEMOS NADA
                        End If

                        '#############################################################
                        'borramos la línea para actualizar completa
                        'sSQL = "UPDATE L "
                        'sSQL &= " SET ""U_EXO_DocEE""='', ""U_EXO_DocE""='',""U_EXO_linE""='' "
                        'sSQL &= " FROM ""@EXO_SIGNUS"" C INNER JOIN ""@EXO_SIGNUSL"" L ON C.""Code""=L.""Code"" "
                        'sSQL &= " WHERE C.""Code""='" & Now.Year.ToString("0000") & "' and ""U_EXO_DocEE""='" & sDocEntry & "' and ""U_EXO_linE""='" & sLinea & "' "
                        'oRs.DoQuery(sSQL)
                    End If
                Next
            Next

            MatrixToNet = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRsArt, Object))
        End Try
    End Function
    Public Shared Function AsignarDev(ByRef oForm As SAPbouiCOM.Form, ByRef oCompany As SAPbobsCOM.Company, ByRef oobjglobal As EXO_Generales.EXO_General, ByVal sDocEntry As String) As Boolean

        Dim sItemCode As String = ""
        Dim sCantidad As String = "0"
        Dim sLinea As String = ""

        Dim sSQL As String = ""
        Dim oRs As SAPbobsCOM.Recordset = Nothing
        Dim oRsLin As SAPbobsCOM.Recordset = Nothing
        AsignarDev = False

        Try
            oRs = CType(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
            oRsLin = CType(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
            sSQL = " SELECT * FROM ""RDN1"" WHERE DocEntry=" & sDocEntry & " Order by ""LineNum"" "
            oRsLin.DoQuery(sSQL)
            If oRsLin.RecordCount > 0 Then
                For L = 0 To oRsLin.RecordCount - 1
                    If oRsLin.Fields.Item("BaseEntry").Value.ToString <> "" And oRsLin.Fields.Item("Baseline").Value.ToString <> "" Then
                        'borramos la línea para actualizar completa
                        sSQL = "UPDATE L "
                        sSQL &= " Set ""U_EXO_DocEntry""='', ""U_EXO_DocNum""='',""U_EXO_linD""='', ""U_EXO_objType""='-' "
                        sSQL &= " FROM ""@EXO_SIGNUS"" C INNER JOIN ""@EXO_SIGNUSL"" L ON C.""Code""=L.""Code"" "
                        sSQL &= " WHERE C.""Code""='" & Now.Year.ToString("0000") & "' and ""U_EXO_DocE""='" & oRsLin.Fields.Item("BaseEntry").Value.ToString & "' "
                        sSQL &= " And ""U_EXO_linE""='" & oRsLin.Fields.Item("BaseEntry").Value.ToString & "' "
                        sSQL &= " and L.""U_EXO_DocEntry""='" & sDocEntry & "' and ""U_EXO_linD""='" & oRsLin.Fields.Item("Baseline").Value.ToString & "' "
                        sSQL &= " and ""U_EXO_objType""='16'"
                        oRs.DoQuery(sSQL)

                        Dim icantidadLinea As Integer = CInt(oRsLin.Fields.Item("Quantity").Value.ToString)
                        For i = 1 To icantidadLinea
                            sSQL = "UPDATE L "
                            sSQL &= " Set ""U_EXO_DocEntry""='" & sDocEntry & "', ""U_EXO_DocNum""='" & oForm.DataSources.DBDataSources.Item("ORDN").GetValue("DocNum", 0).ToUpper & "',"
                            sSQL &= " ""U_EXO_linD""='" & oRsLin.Fields.Item("LineNum").Value.ToString & "', ""U_EXO_ObjType""='16' "
                            sSQL &= " FROM ""@EXO_SIGNUS"" C INNER JOIN ""@EXO_SIGNUSL"" L ON C.""Code""=L.""Code"" "
                            sSQL &= " WHERE C.""Code""='" & Now.Year.ToString("0000") & "' and ""U_EXO_DocEE""='" & oRsLin.Fields.Item("BaseEntry").Value.ToString & "' "
                            sSQL &= " And ""U_EXO_linE""='" & oRsLin.Fields.Item("Baseline").Value.ToString & "' "
                            sSQL &= " and L.""LineId""=( Select MIN(""LineId"") ""Linea"" "
                            sSQL &= " From ""OITM"" I INNER JOIN ""@EXO_SIGNUSL"" SL On I.""U_TXECOVATORSP"" =SL.""U_EXO_TIPO"" INNER JOIN ""@EXO_SIGNUS"" S On S.""Code""=SL.""Code"" "
                            sSQL &= " WHERE I.""Itemcode""='" & oRsLin.Fields.Item("ItemCode").Value.ToString & "' "
                            sSQL &= " And ""U_EXO_DocEE""='" & oRsLin.Fields.Item("BaseEntry").Value.ToString & "' "
                            sSQL &= " And ""U_EXO_linE""='" & oRsLin.Fields.Item("Baseline").Value.ToString & "' "
                            sSQL &= " and ""U_EXO_objType""='-' "
                            sSQL &= " and S.Code='" & Now.Year.ToString("0000") & "') "
                            oRs.DoQuery(sSQL)
                        Next
                    End If
                    oRsLin.MoveNext()
                Next
            End If
            AsignarDev = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRsLin, Object))
        End Try
    End Function
    Public Shared Function AsignarAbo(ByRef oForm As SAPbouiCOM.Form, ByRef oCompany As SAPbobsCOM.Company, ByRef oobjglobal As EXO_Generales.EXO_General, ByVal sDocEntry As String) As Boolean

        Dim sItemCode As String = ""
        Dim sCantidad As String = "0"
        Dim sLinea As String = ""

        Dim sSQL As String = ""
        Dim oRs As SAPbobsCOM.Recordset = Nothing
        Dim oRsLin As SAPbobsCOM.Recordset = Nothing
        AsignarAbo = False

        Try
            oRs = CType(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
            oRsLin = CType(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
            sSQL = " SELECT R.""DocEntry"" ""DEAbo"", R.""LineNum"" ""LNAbo"",R.""BaseEntry"", R.""Baseline"", D.""DocEntry"" ""DEEnt"", D.""LineNum"" ""LNEnt"", "
            sSQL &= " R.""Quantity"", R.""ItemCode"" "
            sSQL &= " FROM ""RIN1"" R INNER JOIN ""INV1"" I ON I.""DocEntry""=R.""BaseEntry"" And I.""LineNum""=R.""BaseLine"" "
            sSQL &= " INNER JOIN ""DLN1"" D On D.""DocEntry""=I.""BaseEntry"" And D.""LineNum""=I.""BaseLine"" "
            sSQL &= " WHERE R.""DocEntry""=" & sDocEntry & " Order by R.""LineNum"" "
            oRsLin.DoQuery(sSQL)
            If oRsLin.RecordCount > 0 Then
                For L = 0 To oRsLin.RecordCount - 1
                    If oRsLin.Fields.Item("BaseEntry").Value.ToString <> "" And oRsLin.Fields.Item("Baseline").Value.ToString <> "" Then
                        'borramos la línea para actualizar completa
                        sSQL = "UPDATE L "
                        sSQL &= " Set ""U_EXO_DocEntry""='', ""U_EXO_DocNum""='',""U_EXO_linD""='', ""U_EXO_objType""='-' "
                        sSQL &= " FROM ""@EXO_SIGNUS"" C INNER JOIN ""@EXO_SIGNUSL"" L ON C.""Code""=L.""Code"" "
                        sSQL &= " WHERE C.""Code""='" & Now.Year.ToString("0000") & "' and ""U_EXO_DocE""='" & oRsLin.Fields.Item("DEEnt").Value.ToString & "' "
                        sSQL &= " And ""U_EXO_linE""='" & oRsLin.Fields.Item("LNEnt").Value.ToString & "' "
                        sSQL &= " and L.""U_EXO_DocEntry""='" & sDocEntry & "' and ""U_EXO_linD""='" & oRsLin.Fields.Item("LNAbo").Value.ToString & "' "
                        sSQL &= " and ""U_EXO_ObjType""='14' "
                        oRs.DoQuery(sSQL)

                        Dim icantidadLinea As Integer = CInt(oRsLin.Fields.Item("Quantity").Value.ToString)
                        For i = 1 To icantidadLinea
                            sSQL = "UPDATE L "
                            sSQL &= " SET ""U_EXO_DocEntry""='" & sDocEntry & "', ""U_EXO_DocNum""='" & oForm.DataSources.DBDataSources.Item("ORIN").GetValue("DocNum", 0).ToUpper & "',"
                            sSQL &= " ""U_EXO_linD""='" & oRsLin.Fields.Item("LNAbo").Value.ToString & "', ""U_EXO_ObjType""='14' "
                            sSQL &= " FROM ""@EXO_SIGNUS"" C INNER JOIN ""@EXO_SIGNUSL"" L ON C.""Code""=L.""Code"" "
                            sSQL &= " WHERE C.""Code""='" & Now.Year.ToString("0000") & "' and ""U_EXO_DocEE""='" & oRsLin.Fields.Item("DEEnt").Value.ToString & "' "
                            sSQL &= " And ""U_EXO_linE""='" & oRsLin.Fields.Item("LNEnt").Value.ToString & "' "
                            sSQL &= " and L.""LineId""=( Select MIN(""LineId"") ""Linea"" "
                            sSQL &= " From ""OITM"" I INNER JOIN ""@EXO_SIGNUSL"" SL On I.""U_TXECOVATORSP"" =SL.""U_EXO_TIPO"" INNER JOIN ""@EXO_SIGNUS"" S On S.""Code""=SL.""Code"" "
                            sSQL &= " WHERE I.""Itemcode""='" & oRsLin.Fields.Item("ItemCode").Value.ToString & "' "
                            sSQL &= " And ""U_EXO_DocEE""='" & oRsLin.Fields.Item("DEEnt").Value.ToString & "' "
                            sSQL &= " And ""U_EXO_linE""='" & oRsLin.Fields.Item("LNEnt").Value.ToString & "' "
                            sSQL &= " and ""U_EXO_objType""='-' "
                            sSQL &= " and S.Code='" & Now.Year.ToString("0000") & "') "
                            oRs.DoQuery(sSQL)
                        Next
                    End If
                    oRsLin.MoveNext()
                Next
            End If
            AsignarAbo = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRsLin, Object))
        End Try
    End Function
#End Region
End Class
