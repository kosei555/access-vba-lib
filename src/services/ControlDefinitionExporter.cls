VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ControlDefinitionExporter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database

Implements IObjectDefinitionExporter

Private Const OPENED_FORM_DETECTED_MESSAGE As String = "フォームの情報を取得するため、全てのフォームを閉じてください。フォームを閉じたくない場合はCancelを押して処理を中止してください。"
Private Const FORM_NAME_COLUMN_NUM As Integer = 0
Private Const CONTROL_NAME_COLUMN_NUM As Integer = 1
Private Const CONTROL_TYPE_COLUMN_NUM As Integer = 2
Private Const CONTROL_PROPERTY_NAME_COLUMN_NUM As Integer = 3
Private Const CONTROL_PROPERTY_VALUE_COLUMN_NUM As Integer = 4
Private Const CAPTION As String = "Caption"
Private Const CONTROLSOURCE As String = "ControlSource"
Private Const ROWSOURCE As String = "RowSource"
Private Const ROWSOURCETYPE As String = "RowSourceType"


Private m_controlTypeToStr As Object
Private m_project As CurrentProject
Private m_propertiesList As Object
Private m_propertiesConf(34) As Collection


Public Sub Class_Initialize()
    Set m_project = CurrentProject
    SetControlTypeToStr
    SetPropertiesConf
End Sub

Public Sub IObjectDefinitionExporter_ExportToCsv(fullPath As String)
    'On Error GoTo Catch
    Dim csvW As ICsvWriter
    Dim csvD(4) As Variant
    Dim accessOb As AccessObject
    Dim formName As String
    Dim frm As Form
    Dim ctl As Control
    Dim ctlName As String
    Dim ctlType As Integer
    Dim prop As Variant
    
    Set csvW = New CsvWriter
        
    csvW.CreateFile (fullPath)
    csvD(FORM_NAME_COLUMN_NUM) = "フォーム名"
    csvD(CONTROL_NAME_COLUMN_NUM) = "コントロール名"
    csvD(CONTROL_TYPE_COLUMN_NUM) = "コントロールタイプ"
    csvD(CONTROL_PROPERTY_NAME_COLUMN_NUM) = "プロパティ名"
    csvD(CONTROL_PROPERTY_VALUE_COLUMN_NUM) = "プロパティ値"
    csvW.AddLine csvD()
    
    
    If (CheckOpenedForm <> True) Then
        GoTo Catch
    End If
    For Each accessOb In m_project.AllForms
        formName = accessOb.Name
        DoCmd.OpenForm formName
        Set frm = Forms(formName)
        csvD(FORM_NAME_COLUMN_NUM) = formName
        For Each ctl In frm.Controls
            ctlName = ctl.Name
            ctlType = ctl.ControlType
            csvD(CONTROL_NAME_COLUMN_NUM) = ctlName
            csvD(CONTROL_TYPE_COLUMN_NUM) = m_controlTypeToStr(ctlType)
            For Each prop In m_propertiesConf(ctlType - 100)
                csvD(CONTROL_PROPERTY_NAME_COLUMN_NUM) = prop
                csvD(CONTROL_PROPERTY_VALUE_COLUMN_NUM) = CallByName(ctl, prop, VbGet)
                csvW.AddLine csvD()
                csvD(FORM_NAME_COLUMN_NUM) = ""
                csvD(CONTROL_NAME_COLUMN_NUM) = ""
                csvD(CONTROL_TYPE_COLUMN_NUM) = ""
            Next
            If (m_propertiesConf(ctlType - 100).Count = 0) Then
                csvD(CONTROL_PROPERTY_NAME_COLUMN_NUM) = ""
                csvD(CONTROL_PROPERTY_VALUE_COLUMN_NUM) = ""
                csvW.AddLine csvD()
                csvD(FORM_NAME_COLUMN_NUM) = ""
            End If
        Next
        
        DoCmd.Close acForm, formName
    Next
    csvW.Complete
    GoTo Finally
Catch:
Finally:
    
End Sub
Private Function CheckOpenedForm() As Boolean
    'False→処理実行をキャンセル
    'True→すべてのフォームを閉じた
    Dim frm As Form
    CheckOpenedForm = True
Retry:
    For Each frm In Forms
        If (ShowOpenedFormDetectedWarning) Then
            GoTo Retry
        Else
            CheckOpenedForm = False
            GoTo Cancel
        End If
        
    Next frm
Cancel:
    
End Function

Private Function ShowOpenedFormDetectedWarning() As Boolean
    Dim boxResult As VbMsgBoxResult
    boxResult = MsgBox(OPENED_FORM_DETECTED_MESSAGE, vbRetryCancel)
    If (boxResult = vbRetry) Then
        ShowOpenedFormDetectedWarning = True
    Else
        ShowOpenedFormDetectedWarning = False
    End If
    
End Function


Private Sub SetControlTypeToStr()
    Set m_controlTypeToStr = CreateObject("Scripting.Dictionary")
    With m_controlTypeToStr
        .Add 100, "acLabel"
        .Add 101, "acRectangle"
        .Add 102, "acLine"
        .Add 103, "acImage"
        .Add 104, "acCommandButton"
        .Add 105, "acOptionButton"
        .Add 106, "acCheckBox"
        .Add 107, "acOptionGroup"
        .Add 108, "acBoundObjectFrame"
        .Add 109, "acTextBox"
        .Add 110, "acListBox"
        .Add 111, "acComboBox"
        .Add 112, "acSubForm"
        .Add 114, "acObjectFrame"
        .Add 118, "acPageBreak"
        .Add 119, "acCustomControl"
        .Add 122, "acToggleButton"
        .Add 123, "acTabCtl"
        .Add 124, "acPage"
        .Add 126, "acAttachment"
        .Add 127, "acEmptyCell"
        .Add 128, "acWebBrowser"
        .Add 129, "acNavigationControl"
        .Add 130, "acNavigationButton"
        .Add 133, "acChart"
        .Add 134, "acEdgeBrowser"
    End With
End Sub
Private Sub SetPropertiesConf()
    Dim i As Integer
    For i = 0 To 34
        Set m_propertiesConf(i) = New Collection
    Next i
    m_propertiesConf(0).Add CAPTION
    m_propertiesConf(4).Add CAPTION
    m_propertiesConf(6).Add CONTROLSOURCE
    m_propertiesConf(9).Add CONTROLSOURCE
    m_propertiesConf(10).Add CONTROLSOURCE
    m_propertiesConf(10).Add ROWSOURCE
    m_propertiesConf(10).Add ROWSOURCETYPE
    m_propertiesConf(11).Add CONTROLSOURCE
    m_propertiesConf(11).Add ROWSOURCE
    m_propertiesConf(11).Add ROWSOURCETYPE
End Sub
