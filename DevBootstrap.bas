Attribute VB_Name = "DevBootstrap"
Option Explicit

Private m_fso As Object

Private Sub PushDevModules()
    ' dev_helperプロジェクト関係のモジュールを所定フォルダに全push
    '   - 対象の接頭辞はべた書きしているので、変更時はメンテが必要
    Const DEV_HELPER_REPO_NAME As String = "\.dev_helper\"
    Dim repoDir As String
    repoDir = ThisWorkbook.Path & DEV_HELPER_REPO_NAME
    Set m_fso = CreateObject("Scripting.FileSystemObject")
    ' `.dev_helper`フォルダがなかったら作る
    If Not m_fso.FolderExists(repoDir) Then
        Call m_fso.CreateFolder(repoDir)
    End If
    Dim comp As Object, cn As String, f As String, p As String
    ' 接頭辞`Dev`のもののみリポジトリに送り込む
    For Each comp In ThisWorkbook.VBProject.VBComponents
        cn = comp.Name
        If StrComp(Left$(cn, 3), "Dev", vbBinaryCompare) <> 0 Then GoTo Continue
        ' エクスポート用ファイルパス取得
        f = cn & ComponentExtension(comp.Type)
        p = repoDir & f
        ' 既存の同名ファイルがあればポア
        '   - DeleteFile(filespec, [force])
        If m_fso.FileExists(p) Then Call m_fso.DeleteFile(p, True)
        Call comp.Export(p)
        Debug.Print "Pushed: " & f
Continue:
    Next
End Sub

Private Sub PullDevModules()
    ' dev_helperプロジェクト関係のモジュールを一括pull
    '   - このモジュールだけは手動pull
        
End Sub
