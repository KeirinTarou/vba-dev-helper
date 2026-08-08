Attribute VB_Name = "DevBootstrap"
Option Explicit

' =============================================================================
'   dev_helper関係モジュールpush・pull用モジュール
' =============================================================================
' 仕様を変更するときはここをメンテ
'   - dev_helper開発用リポジトリフォルダ名
Private Const DEV_HELPER_REPO_NAME As String = "\.dev_helper\"
'   - 自動pull除外用モジュール名（このモジュールの名前）
Private Const SELF_MOD_NAME As String = "DevBootstrap"
'   - 自動push対象ファイルの接頭辞
Public Const DEV_HELPER_PREFIX As String = "Dev"

Private m_fso As Object

Private Sub PushDevModules()
    ' SharePoint上のプロジェクトから実行しようとしたらブロック
    If Left(ThisWorkbook.Path, 5) = "https" Then
        Call MsgBox( _
            Prompt:="SharePoint上のプロジェクトで実行してはいけない。（クソが。）", _
            Buttons:=vbCritical, _
            Title:="(　ﾟдﾟ)､ﾍﾟｯ < クソが。")
        Exit Sub
    End If
    
    ' dev_helperプロジェクト関係のモジュールを所定フォルダに全push
    Set m_fso = CreateObject("Scripting.FileSystemObject")
    
    Dim repodir As String
    repodir = ThisWorkbook.Path & DEV_HELPER_REPO_NAME
    ' `.dev_helper`フォルダがなかったら作る
    If Not m_fso.FolderExists(repodir) Then
        Call m_fso.CreateFolder(repodir)
    End If
    Dim comp As Object, cn As String, f As String, p As String
    ' 接頭辞`Dev`のもののみリポジトリに送り込む
    For Each comp In ThisWorkbook.VBProject.VBComponents
        cn = comp.Name
        If Not IsDevHelper(cn) Then GoTo Continue
        ' エクスポート用ファイルパス取得
        f = cn & ComponentExtension(comp.Type)
        p = repodir & f
        ' 既存の同名ファイルがあるとき
        '   - DeleteFile(filespec, [force])
        If m_fso.FileExists(p) Then
            ' プロジェクト側とリポジトリ側のコードを取得
            Dim projCode As String, repoCode As String
            projCode = ExtractProjectCode(comp)
            repoCode = ExtractCodeBody(p)
            ' 内容が同じだったらスキップ
            If Not HasChanged(projCode, repoCode) Then
                Debug.Print "Skipped: " & f
                GoTo Continue
            ' 内容が異なっていたら、既存ファイルをポア
            Else
                Call m_fso.DeleteFile(p, True)
            End If
        End If
        ' 既存の同名ファイルがない場合
        ' 既存の同名ファイルがあり、内容が異なる場合
        '   -> エクスポート
        Call comp.Export(p)
        Debug.Print "Pushed: " & f
Continue:
    Next
End Sub

Private Sub PullDevModules()
    ' SharePoint上のプロジェクトから実行しようとしたらブロック
    If Left(ThisWorkbook.Path, 5) = "https" Then
        Call MsgBox( _
            Prompt:="SharePoint上のプロジェクトで実行してはいけない。（クソが。）", _
            Buttons:=vbCritical, _
            Title:="(　ﾟдﾟ)､ﾍﾟｯ < クソが。")
        Exit Sub
    End If
    
    ' dev_helperプロジェクト関係のモジュールを一括pull
    '   - このモジュールだけは手動pull
    Set m_fso = CreateObject("Scripting.FileSystemObject")
    
    Dim repodir As String
    repodir = ThisWorkbook.Path & DEV_HELPER_REPO_NAME
    Dim f As Object, ext As String, bs As String
    For Each f In m_fso.GetFolder(repodir).Files
        ext = LCase$(m_fso.GetExtensionName(f.Path))
        bs = m_fso.GetBaseName(f.Path)
        ' `.bas`、`.cls`以外はスルー
        If Not (ext = "bas" Or ext = "cls") Then GoTo Continue
        ' dev_helper以外はスルー
        If Not IsDevHelper(bs) Then GoTo Continue
        ' このモジュールもスルー
        If bs = SELF_MOD_NAME Then GoTo Continue
        ' ここまで来たらインポート候補
        Dim repoCode As String, comp As Object, projCode As String
        ' リポジトリのコードを取得
        '   - 全体を取得 -> 正味のコード部分を取得
        repoCode = ExtractCodeBody(f.Path)
        ' モジュールのコードを取得
        Set comp = FindComponent(bs)
        ' 対応するモジュールがプロジェクト側にない -> pullしてContinue
        If comp Is Nothing Then
            Call ImportComponent(f.Path)
            Debug.Print "Pulled: " & f.Name
            GoTo Continue
        End If
        projCode = ExtractProjectCode(comp)
        ' モジュールのコードとリポジトリのコードを比較
        If HasChanged(projCode, repoCode) Then
            ' 不一致だったらコードを闘魂注入
            Call ImportComponent(f.Path)
            Debug.Print "Pulled: " & f.Name
        Else
            ' 一致していたらスキップ
            Debug.Print "Skipped: " & f.Name
        End If
Continue:
    Next
End Sub

Private Sub AA_HelperFunctions(): End Sub
' =============================================================================
'   Helper Functions
' =============================================================================
Public Function IsDevHelper( _
            ByVal a_ComponentName As String) As Boolean
    Const ERR_SOURCE As String = SELF_MOD_NAME & ".IsTarget()"
    IsDevHelper = False
    Dim cn As String
    cn = a_ComponentName
    ' 3文字未満は問答無用でFalse
    If Len(cn) < 3 Then Exit Function
    IsDevHelper = _
        (StrComp(Left$(cn, 3), DEV_HELPER_PREFIX, vbBinaryCompare) = 0)
End Function

Public Function ComponentFolderName( _
            ByVal a_ComponentType As vbext_ComponentType) As String
    ' モジュールの種別に応じたリポジトリのサブフォルダ名を返す
    Const ERR_SOURCE As String = "DevProj.GetFolderName()"
    Dim ret As String
    Select Case a_ComponentType
        Case vbext_ct_StdModule: ret = "std_mod"
        Case vbext_ct_ClassModule: ret = "cls_mod"
        Case vbext_ct_MSForm: ret = "frm_mod"
        Case vbext_ct_Document: ret = "doc_mod"
        Case vbext_ct_ActiveXDesigner
            Call RaiseError( _
                ERR_NOT_SUPPORTED, ERR_SOURCE, "ActiveX Designer はサポート対象外である。")
        Case Else
            Call RaiseError( _
                ERR_INVALID_ARGUMENT, ERR_SOURCE, "不正な引数が渡された。")
    End Select
    ComponentFolderName = ret
End Function

Public Function ComponentTypeName( _
            ByVal a_ComponentType As vbext_ComponentType) As String
    ' コンポーネントの種別を返す
    Const ERR_SOURCE As String = "DevProj.ComponentTypeName()"
    Dim ret As String
    Select Case a_ComponentType
        Case vbext_ct_StdModule: ret = "Standard Module"
        Case vbext_ct_ClassModule: ret = "Class Module"
        Case vbext_ct_MSForm: ret = "UserForm"
        Case vbext_ct_Document: ret = "Document Module"
        Case vbext_ct_ActiveXDesigner: ret = "ActiveX Designer"
        Case Else
            Call RaiseError( _
                ERR_INVALID_ARGUMENT, ERR_SOURCE, "不正な引数が渡された。")
    End Select
    ComponentTypeName = ret
End Function

Public Sub RaiseError( _
            a_Number As Long, _
            a_Source As String, _
            a_Description As String)
    ' カスタム例外をスローする
    Call Err.Raise( _
        Number:=a_Number, _
        Source:=a_Source, _
        Description:=a_Description & "(ﾟдﾟ)､ < クソが。")
End Sub


