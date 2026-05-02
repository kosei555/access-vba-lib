VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CsvWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'-------------------------------------------------------------------------------
'クラス名   :   CsvWriter
'役割       :   Csvファイル出力
'依存関係   :   ICsvWriter(インターフェイス)
'使い方     :   CreateFileでフルパスを渡し、一時ファイル作成
'                               ↓
'               AddLineで一行ずつcsvを書き込む
'                               ↓
'               Completeで一時ファイルをcsvファイルに変換し、csv完成
'-------------------------------------------------------------------------------
Option Compare Database

Implements ICsvWriter

Const DEFAULT_SEPARATOR As String = "|"
Const ERR_FOLDER_NOT_EXIST_SOURCE As String = "CsvWriter.CreateFile"
Const ERR_FOLER_NOT_EXIST_DESCRIPTION As String = "ファイル作成先のフォルダーが存在しません。"
Const ERR_FILE_EXIST_SOURCE As String = "CsvWriter.CreateFile"
Const ERR_FILE_EXIST_DESCRIPTION As String = "既に存在しているファイルを作成しようとしました。"
Const ERR_SEPARATOR_INCLUDED_SOURCE As String = "CsvWriter.AddLine"
Const ERR_SEPARATOR_INCLUDED_DESCRIPTION As String = "区切り文字が含まれた文字列をcsvファイルに書き込もうとしました。"
Const ANSI As Integer = 0
Const TMP_EXTENSION As String = ".tmp"

Private m_fso As Object
Private m_ts As Object
Private m_csvPath As String
Private m_tmpPath As String
Private m_separator As String

Public Sub Class_Initialize()
    m_separator = DEFAULT_SEPARATOR
    Set m_ts = Nothing
    Set m_fso = CreateObject("Scripting.FileSystemObject")
End Sub

Public Sub ICsvWriter_SetSeparator(separator As String)
    m_separator = separator
End Sub

Public Sub ICsvWriter_CreateFile(path As String)
    Dim fileName As String
    Dim fileNameLen As Long
    Dim pathLen As Long
    m_csvPath = ICsvWriter_ExpandEnvironment(path)
    m_tmpPath = ICsvWriter_ExpandEnvironment(path) & TMP_EXTENSION
    fileName = m_fso.GetFileName(path)
    fileNameLen = Len(fileName)
    pathLen = Len(path)
    If (Not m_fso.FolderExists(Mid(path, 1, pathLen - fileNameLen))) Then
        ICsvWriter_Clean
        Err.Raise 513, ERR_FOLDER_NOT_EXIST_SOURCE, ERR_FOLER_NOT_EXIST_DESCRIPTION
    ElseIf (m_fso.FileExists(m_csvPath) Or m_fso.FileExists(m_tmpPath)) Then
        Debug.Print m_csvPath & " " & m_tmpPath
        ICsvWriter_Clean
        Err.Raise 513, ERR_FILE_EXIST_SOURCE, ERR_FILE_EXIST_DESCRIPTION
    End If
    Set m_ts = m_fso.OpenTextFile(m_tmpPath, ForWriting, True, ANSI)
End Sub

Public Sub ICsvWriter_AddLine(data() As String)
    Dim line As String
    Dim str As Variant
    line = ""
    For Each str In data
        If (InStr(1, str, m_separator) > 0) Then
            ICsvWriter_Clean
            Err.Raise 513, ERR_SEPARATOR_INCLUDED_SOURCE, ERR_SEPARATOR_INCLUDED_DESCRIPTION
        End If
        line = line & str & m_separator
    Next
    line = Mid(line, 1, Len(line) - 1)
    m_ts.WriteLine (line)
End Sub

Public Sub ICsvWriter_Complete()
    m_ts.Close
    Name m_tmpPath As m_csvPath
End Sub

Private Function ICsvWriter_ExpandEnvironment(str As String) As String
    ICsvWriter_ExpandEnvironment = CreateObject("WScript.Shell").ExpandEnvironmentStrings(str)
End Function

Private Sub ICsvWriter_Clean()
    If (Not m_ts Is Nothing) Then
        m_ts.Close
    End If
    If (m_fso.FileExists(m_csvPath)) Then
        m_fso.DeleteFile (m_csvPath)
    End If
End Sub


