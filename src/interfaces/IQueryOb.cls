VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IQueryOb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Public Sub QueryObInit(db As DAO.database, qName As String)
End Sub
Public Sub SetParam(pName As String, pData As Variant)
    
End Sub

Public Function GetQueryCode() As String
End Function

Public Function ExecQuery() As IErrBuf

End Function
Public Function GetQueryRecordSet(rec As DAO.Recordset, isDynaset As Boolean) As IErrBuf

End Function



