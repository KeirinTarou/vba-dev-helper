Attribute VB_Name = "DevProjIO"
Option Explicit

Private m_fso As Object

Private Sub AA_DevHelpers(): End Sub
' 開発者向けヘルパ


Private Sub AA_HelperFunctions(): End Sub


Private Sub AA_Experiments(): End Sub

Private Sub Exp_EnumerateModuleNames()
    ' プロジェクトにぶら下がるモジュール名とTypeを列挙する
    Dim comp As Object
    For Each comp In ThisWorkbook.VBProject.VBComponents
        Debug.Print _
            comp.Name & ": " & comp.Type & _
            "(" & ComponentTypeName(comp.Type) & ")"
    Next
End Sub

Private Sub Exp_EnumerateCodeModules()
    ' コードモジュールの属性を表示する
    Dim comp As Object
    For Each comp In ThisWorkbook.VBProject.VBComponents
        Debug.Print comp.Name & ": " & comp.CodeModule.CountOfLines & "line(s)."
    Next
End Sub

Private Sub Exp_ExportModule()
    ' モジュールをエクスポートする
    Const REPO_REL_DIR As String = "\repo\"
    Const STD_MOD_REL_DIR As String = "std_mod\"
    Set m_fso = CreateObject("Scripting.FileSystemObject")
    Dim repoDir As String
    repoDir = ThisWorkbook.Path & REPO_REL_DIR
    ' `repo`フォルダがなかったら作る
    If Not m_fso.FolderExists(repoDir) Then
        Call m_fso.CreateFolder(repoDir)
    End If
    Dim repo As Object
    Dim comp As Object, stdDir As String, f As String
    stdDir = repoDir & STD_MOD_REL_DIR
    If Not m_fso.FolderExists(stdDir) Then
        Call m_fso.CreateFolder(stdDir)
    End If
    For Each comp In ThisWorkbook.VBProject.VBComponents
        If comp.Type = vbext_ct_StdModule Then
            f = comp.Name & ".bas"
            Call comp.Export(stdDir & f)
            Exit For
        End If
    Next
End Sub
