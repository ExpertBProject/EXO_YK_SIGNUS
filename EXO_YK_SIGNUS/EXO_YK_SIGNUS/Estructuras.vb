<DataContract()>
<Serializable>
Friend Class DataContractAttribute
    Inherits Attribute
End Class

Public Class CodigosSIGNUM
    <DataMember()>
    Public msgCodigo As String
    <DataMember()>
    Public msgDescripcion As String
    <DataMember()>
    Public msgCampo As String
    <DataMember()>
    Public data As List(Of data)
End Class
Public Class data
    <DataMember()>
    Public codigo As String
    <DataMember()>
    Public estado As String
    <DataMember()>
    Public catComCod As String
    <DataMember()>
    Public catAbrev As String
    <DataMember()>
    Public lote As String
    <DataMember()>
    Public auditSecuencia As String
End Class

Friend Class DataMemberAttribute
    Inherits Attribute
End Class
