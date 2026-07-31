Attribute VB_Name = "DevProjIO"
Option Explicit

Private Const SELF_MOD_NAME As String = "DevProjIO"

Private Const REPO_NAME As String = "repo"
Private Const STD_MOD As String = "std_mod\"
Private Const CLS_MOD As String = "cls_mod\"
Private Const FRM_MOD As String = "frm_mod\"
Private Const DOC_MOD As String = "doc_mod\"

Private m_fso As Object
' =============================================================================
'   公開API
' =============================================================================
Public Sub ExportComponentsToRepository()
    Set m_fso = CreateObject("Scripting.FileSystemObject")
    Dim repoDir As String
    repoDir = ThisWorkbook.Path & "\" & REPO_NAME & "\"
    ' リポジトリフォルダがなければ作る
    Dim stdDir As String, clsDir As String, frmDir As String, docDir As String
    With m_fso
        If Not .FolderExists(repoDir) Then Call .CreateFolder(repoDir)
        stdDir = repoDir & STD_MOD
        If Not .FolderExists(stdDir) Then Call .CreateFolder(stdDir)
        clsDir = repoDir & CLS_MOD
        If Not .FolderExists(clsDir) Then Call .CreateFolder(clsDir)
        frmDir = repoDir & FRM_MOD
        If Not .FolderExists(frmDir) Then Call .CreateFolder(frmDir)
        docDir = repoDir & DOC_MOD
        If Not .FolderExists(docDir) Then Call .CreateFolder(docDir)
    End With
    Call ExportComponents(repoDir)
End Sub

Private Sub AA_HelperFunctions(): End Sub
' =============================================================================
'   開発者向けヘルパ
' =============================================================================
Private Sub ExportComponents( _
            ByVal a_RepositoryPath As String)
    ' プロジェクトの全モジュールをリポジトリにエクスポートする
    '   - `Dev`系は除く
    '   - リポジトリのパスにフォルダがないときは例外スロー
    '   - 想定外の利用法なので、例外を吐いて利用者に知らせる
    Const ERR_SOURCE As String = SELF_MOD_NAME & ".ExportComponents"
    If m_fso Is Nothing Then _
        Set m_fso = CreateObject("Scripting.FileSystemObject")
    ' リポジトリがない -> 例外スロー
    If Not m_fso.FolderExists(a_RepositoryPath) Then _
        Call RaiseError(ERR_NOT_FOUND, ERR_SOURCE, "リポジトリ用フォルダが存在しない。")
    ' ExportComponent()を呼び出す
    Dim comp As Object, outDir As String
    For Each comp In ThisWorkbook.VBProject.VBComponents
        ' `Dev...`モジュールは除外
        If IsDevHelper(comp.Name) Then GoTo Continue
        ' 出力先フォルダパスを取得
        outDir = OutputDirPath(comp.Type)
        Call ExportComponent(comp, outDir)
Continue:
    Next
End Sub

Private Sub ExportComponent( _
            a_Component As Object, _
            a_OutputDir As String)
    ' コンポーネントと保存先フォルダパスを受け取ってエクスポートする
    '   - `a_OutputDir`が存在しない場合は例外スロー
    '   - 100%開発者のせいなので例外キャッチは必要なし
    Const ERR_SOURCE As String = SELF_MOD_NAME & ".ExportComponent"
    If m_fso Is Nothing Then _
        Set m_fso = CreateObject("Scripting.FileSystemObject")
    ' push先フォルダがない -> 例外スロー
    If Not m_fso.FolderExists(a_OutputDir) Then _
        Call RaiseError(ERR_NOT_FOUND, ERR_SOURCE, "保存先フォルダが存在しない。")

    ' エクスポートファイル名を取得 -> 保存先パス作成
    Dim f As String, p As String
    f = a_Component.Name & ComponentExtension(a_Component.Type)
    Dim dlmt As String
    dlmt = "\"
    If Right(a_OutputDir, 1) = "\" Then dlmt = ""
    p = a_OutputDir & dlmt & f
    ' エクスポート
    '   - 例外で失敗したときはイミディエイトに表示
    On Error Resume Next
    ' 既存ファイルがあればポア
    If m_fso.FileExists(p) Then Call m_fso.DeleteFile(p, True)
    ' 既存ファイルのポアに失敗 -> イミディエイトで通知 -> 握りつぶしてExit
    If Err.Number <> 0 Then
        Debug.Print "FAILED: existing `" & f & "` not deleted."
        Debug.Print "Number: " & Err.Number
        Debug.Print "Source: " & Err.Source
        Debug.Print "Desc  : " & Err.Description
        Call Err.Clear
        Exit Sub
    End If
    Call a_Component.Export(p)
    ' エクスポートに失敗していたらイミディエイトに表示して握りつぶす
    If Err.Number <> 0 Then
        Debug.Print "FAILED: `" & f & "` not exported."
        Debug.Print "Number: " & Err.Number
        Debug.Print "Source: " & Err.Source
        Debug.Print "Desc  : " & Err.Description
        Call Err.Clear
    Else
        Debug.Print "Exported: " & f
    End If
    On Error GoTo 0
End Sub

Private Function OutputDirPath( _
            ByVal a_ComponentType As vbext_ComponentType) As String
    Const ERR_SOURCE As String = SELF_MOD_NAME & ".OutputDirPath()"
    Dim ret As String
    
    Dim sf As String    ' sf: SubFolderの略
    Select Case a_ComponentType
        Case vbext_ct_StdModule: sf = STD_MOD
        Case vbext_ct_ClassModule: sf = CLS_MOD
        Case vbext_ct_MSForm: sf = FRM_MOD
        Case vbext_ct_Document: sf = DOC_MOD
        Case vbext_ct_ActiveXDesigner
            Call RaiseError( _
                ERR_NOT_SUPPORTED, ERR_SOURCE, "ActiveX Designerは非対応である。")
        Case Else
            Call RaiseError( _
                ERR_INVALID_ARGUMENT, ERR_SOURCE, "意味不明な引数が渡された。")
    End Select
    ret = ThisWorkbook.Path & "\" & REPO_NAME & "\" & sf
    
    OutputDirPath = ret
End Function

Private Sub AA_Experiments(): End Sub

Private Sub Exp_ExportModules()
    ' プロジェクトの全モジュールをエクスポートする
    Call ExportComponents(ThisWorkbook.Path & "\" & REPO_NAME)
End Sub

Private Sub Exp_ExportModule()
    ' モジュールをエクスポートする
    Set m_fso = CreateObject("Scripting.FileSystemObject")
    Dim comp As Object, stdDir As String, f As String
    stdDir = OutputDirPath(vbext_ct_StdModule)
    If Not m_fso.FolderExists(stdDir) Then
        Call m_fso.CreateFolder(stdDir)
    End If
    For Each comp In ThisWorkbook.VBProject.VBComponents
        If comp.Type = vbext_ct_StdModule Then
            Call ExportComponent(comp, stdDir)
            Exit For
        End If
    Next
End Sub

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

