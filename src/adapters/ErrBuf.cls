VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ErrBuf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database

Implements IErrBuf

Private m_description As String
Private m_number As Integer
Private m_source As String

Property Get IErrBuf_Description() As String
    IErrBuf_Description = m_description
End Property

Property Let IErrBuf_Description(arg_description As String)
    m_description = arg_description
End Property

Property Get IErrBuf_Number() As Integer
    IErrBuf_Number = m_number
End Property

Property Let IErrBuf_Number(arg_number As Integer)
    m_number = arg_number
End Property

Property Get IErrBuf_Source() As String
    IErrBuf_Source = m_source
End Property

Property Let IErrBuf_Source(arg_source As String)
    m_source = arg_source
End Property
