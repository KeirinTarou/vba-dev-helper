Attribute VB_Name = "DevProjIO"
Option Explicit

Private Const SELF_MOD_NAME As String = "DevProjIO"

Private Const REPO_NAME As String = "repo"
Private Const STD_MOD As String = "std_mod\"
Private Const CLS_MOD As String = "cls_mod\"
Private Const FRM_MOD As String = "frm_mod\"
Private Const DOC_MOD As String = "doc_mod\"

' SharePoint上のファイルかどうか
Private m_OnSharePoint As Boolean

' エクスポート/インポート成否カウント用
Private m_SucceededCount As Long
Private m_FailedCount As Long
Private m_SkippedCount As Long

' 警告表示用
Private m_Warnings As String

Private m_fso As Object
' =============================================================================
'   公開API
' =============================================================================
Public Sub ImportComponentsFromRepository()
    Set m_fso = CreateObject("Scripting.FileSystemObject")
    
    ' SharePoint上のファイルのときはフラグを立てる
    If Left(ThisWorkbook.Path, 5) = "https" Then
        m_OnSharePoint = True
    Else
        m_OnSharePoint = False
    End If
    
    ' リポジトリのパスを取得
    Dim repodir As String
    ' SharePoint上のファイルのときは、ユーザの「ドキュメント」フォルダの
    ' `vba_repositories`フォルダ直下にあるブック名のフォルダがリポジトリ
    ' のルートフォルダ、という運用
    If m_OnSharePoint Then
        Dim usrDir As String, projName As String
        usrDir = VBA.Interaction.Environ("USERPROFILE") & "\Documents"
        projName = m_fso.GetBaseName(ThisWorkbook.Name)
        ' ユーザの「ドキュメント」フォルダ直下の`vba_repositories`フォルダ
        Dim parentDir As String
        parentDir = usrDir & "\" & "vba_repositories"
        repodir = parentDir & "\" & projName & "\"
    ' ローカルファイルのときは、隣の既定名のフォルダがリポジトリのルート
    Else
        repodir = ThisWorkbook.Path & "\" & REPO_NAME & "\"
    End If
    
    ' カウンタを初期化
    m_SucceededCount = 0
    m_FailedCount = 0
    m_SkippedCount = 0
    ' 警告を初期化
    m_Warnings = ""
    
    ' モジュールのコードをリポジトリからpull
    Call ImportComponents(repodir)
    
    ' サマリを表示
    Debug.Print String(40, "-")
    Debug.Print "Summary: "
    Debug.Print vbTab & "Pull succeeded: " & CStr(m_SucceededCount) & " module(s)."
    Debug.Print vbTab & "Pull failed   : " & CStr(m_FailedCount) & " module(s)."
    Debug.Print vbTab & "Pull skipped  : " & CStr(m_SkippedCount) & " module(s)."
    If m_Warnings = "" Then Exit Sub
    Debug.Print "【Warning!!】"
    Debug.Print m_Warnings
End Sub

Public Sub ExportComponentsToRepository()
    Set m_fso = CreateObject("Scripting.FileSystemObject")
    
    ' SharePoint上のファイルのときはフラグを立てる
    If Left(ThisWorkbook.Path, 5) = "https" Then
        m_OnSharePoint = True
    Else
        m_OnSharePoint = False
    End If
    
    Dim repodir As String
    ' SharePoint上のファイルのときは、ユーザの「ドキュメント」フォルダに
    ' リポジトリフォルダを作る
    If m_OnSharePoint Then
        Dim usrDir As String, projName As String
        usrDir = VBA.Interaction.Environ("USERPROFILE") & "\Documents"
        projName = m_fso.GetBaseName(ThisWorkbook.Name)
        ' ユーザの「ドキュメント」フォルダに`vba_repositories`フォルダを作る
        Dim parentDir As String
        parentDir = usrDir & "\" & "vba_repositories"
        If Not m_fso.FolderExists(parentDir) Then _
            Call m_fso.CreateFolder(parentDir)
        repodir = parentDir & "\" & projName & "\"
        If Not m_fso.FolderExists(repodir) Then _
            Call m_fso.CreateFolder(repodir)
    Else
        repodir = ThisWorkbook.Path & "\" & REPO_NAME & "\"
    End If
    
    ' リポジトリフォルダがなければ作る
    Dim stdDir As String, clsDir As String, frmDir As String, docDir As String
    With m_fso
        If Not .FolderExists(repodir) Then Call .CreateFolder(repodir)
        stdDir = repodir & STD_MOD
        If Not .FolderExists(stdDir) Then Call .CreateFolder(stdDir)
        clsDir = repodir & CLS_MOD
        If Not .FolderExists(clsDir) Then Call .CreateFolder(clsDir)
        frmDir = repodir & FRM_MOD
        If Not .FolderExists(frmDir) Then Call .CreateFolder(frmDir)
        docDir = repodir & DOC_MOD
        If Not .FolderExists(docDir) Then Call .CreateFolder(docDir)
    End With
    
    ' カウンタを初期化
    m_SucceededCount = 0
    m_FailedCount = 0
    m_SkippedCount = 0
    
    Call ExportComponents(repodir)
    Debug.Print String(40, "-")
    Debug.Print "Summary: "
    Debug.Print vbTab & "Push succeeded: " & CStr(m_SucceededCount) & " module(s)."
    Debug.Print vbTab & "Push failed   : " & CStr(m_FailedCount) & " module(s)."
    Debug.Print vbTab & "Push skipped  : " & CStr(m_SkippedCount) & " module(s)."
End Sub

Private Sub AA_HelperFunctions(): End Sub
' =============================================================================
'   開発者向けヘルパ
' =============================================================================
Private Sub ImportComponents( _
            ByVal a_RepositoryPath As String)
    ' プロジェクト内のモジュールのコードを
    ' リポジトリ内のモジュールのコードで置き換える
    '   - 内容が一致していたらスキップする
        
    '   - リポジトリのパスにフォルダがないときは例外スロー
    '   - 想定外の利用法なので、例外を吐いて利用者に知らせる
    Const ERR_SOURCE As String = SELF_MOD_NAME & ".ImportComponents"
    If m_fso Is Nothing Then _
        Set m_fso = CreateObject("Scripting.FileSystemObject")
    ' リポジトリがない -> 例外スロー
    If Not m_fso.FolderExists(a_RepositoryPath) Then _
        Call DevUtils.RaiseError( _
            ERR_NOT_FOUND, ERR_SOURCE, "リポジトリ用フォルダが存在しない。")
    
    ' ImportComponent()を呼び出す
    Dim repo As Object, f As Object, fd As Object
    ' リポジトリのScripting.Folderオブジェクトを取得
    Set repo = m_fso.GetFolder(a_RepositoryPath)
    For Each fd In repo.SubFolders
        ' サブフォルダのサブフォルダは考慮外で良い（再帰不要）
        For Each f In fd.Files
            Dim ext As String
            ext = LCase$(m_fso.GetExtensionName(f.Name))
            If ext = "bas" Or ext = "cls" Or ext = "frm" Then
                Call ImportComponent(f.Path)
            End If
        Next
    Next
End Sub

Public Sub ImportComponent( _
            ByVal a_ComponentPath As String)
    Const ERR_SOURCE As String = SELF_MOD_NAME & ".ImportComponent()"
    ' モジュール名を取得
    Dim modName As String
    modName = DevUtils.ExtractModuleName(a_ComponentPath)
    Dim comp As Object
    Set comp = DevUtils.FindComponent(modName)
    
    ' 例外発生時はメッセージを表示して握りつぶす
    On Error GoTo HandleError
    ' プロジェクト内に同名モジュールがない -> そのままpull
    If comp Is Nothing Then
        Call ThisWorkbook.VBProject.VBComponents.Import(a_ComponentPath)
    ' プロジェクト内に同名モジュールあり -> コードのみpull
    Else
        Dim repoCode As String, projCode As String, _
            hasCstAttr As Boolean
        
        ' プロジェクト内のモジュールからコード部分を取得
        projCode = DevUtils.ExtractProjectCode(comp)
        ' リポジトリから正味のコード部分を取得
        '   - カスタムAttributeを取り除いて比較する
        hasCstAttr = DevUtils.HasCustomAttribute(a_ComponentPath)
        ' カスタムAttributeがあるときは取り除いて取得
        repoCode = DevUtils.ExtractCodeBody(a_ComponentPath, hasCstAttr)
        ' カスタムAttributeの変更が同期されない旨、警告を表示する
        If hasCstAttr Then
            If m_Warnings <> "" Then m_Warnings = m_Warnings & vbCrLf
            m_Warnings = _
                m_Warnings & _
                vbTab & "`" & modName & "`" & _
                "のカスタムAttributeに変更があっても同期されない。"
        End If
        ' 両者の内容が一致していたらpullしない
        If Not DevUtils.HasChanged(projCode, repoCode) Then
            ' スキップカウンタをインクリメント
            Debug.Print "Skipped: " & modName
            m_SkippedCount = m_SkippedCount + 1
            Exit Sub
        End If
        ' カスタムAttributeの有無を判定
        '   - カスタムAttributeあり -> モジュールをまるごと差し替える
        '   - カスタムAttributeなし -> コードのみ差し替える
        If hasCstAttr Then
            ' モジュールのまるごと差し替え
            ' ただし、ドキュメントモジュールはポアできないので、警告を表示してスキップ
            If comp.Type = vbext_ct_Document Then
                Debug.Print "Skipped: " & modName
                m_SkippedCount = m_SkippedCount + 1
                If m_Warnings <> "" Then m_Warnings = m_Warnings & vbCrLf
                m_Warnings = _
                    m_Warnings & _
                    vbTab & "`" & modName & "`" & _
                    "にはカスタムAttributeがあるため自動同期不可。" & vbCrLf & _
                    vbTab & vbTab & "手動でインポートしやがれクズが。(ﾟдﾟ)､ﾍﾟｯ"
                Exit Sub
            End If
            
            ' 既存モジュールをポア
            Call ThisWorkbook.VBProject.VBComponents.Remove(comp)
            ' リポジトリのモジュールをインポート
            Call ThisWorkbook.VBProject.VBComponents.Import(a_ComponentPath)
        Else
            ' コードのみ差し替え
            ' 既存のコードを削除
            '   - モジュールが空の場合をガード
            If comp.CodeModule.CountOfLines > 0 Then
                Call DevUtils.DeleteCodeLines(comp)
            End If
            ' プロジェクトのモジュールに闘魂注入
            Call comp.CodeModule.AddFromString(repoCode)
        End If
    End If
    ' 成功カウンタをインクリメント
    Debug.Print "Pulled: " & modName
    m_SucceededCount = m_SucceededCount + 1
    Exit Sub
HandleError:
    Dim errNum As Long, errSrc As String, errDesc As String
    errNum = Err.Number
    errSrc = Err.Source & "." & ERR_SOURCE
    errDesc = Err.Description & vbCrLf & _
        "(モジュール`" & modName & "`のpullに失敗した。)"
    Debug.Print "Pull `" & modName & "` module failed."
    Debug.Print "Number: " & errNum
    Debug.Print "Source: " & errSrc
    Debug.Print "Desc  : " & errDesc
    ' 失敗カウンタをインクリメント
    m_FailedCount = m_FailedCount + 1
    ' 例外は握りつぶす
    Call Err.Clear
End Sub

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
        Call DevUtils.RaiseError( _
            ERR_NOT_FOUND, ERR_SOURCE, "リポジトリ用フォルダが存在しない。")
    ' ExportComponent()を呼び出す
    Dim comp As Object, outDir As String
    For Each comp In ThisWorkbook.VBProject.VBComponents
        ' `Dev...`モジュールは除外
        If DevUtils.IsDevHelper(comp.Name) Then GoTo Continue
        ' 出力先フォルダパスを取得
        outDir = OutputDirPath(comp.Type, a_RepositoryPath)
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
        Call DevUtils.RaiseError( _
            ERR_NOT_FOUND, ERR_SOURCE, "保存先フォルダが存在しない。")

    ' エクスポートファイル名を取得 -> 保存先パス作成
    Dim f As String, p As String
    f = a_Component.Name & DevUtils.ComponentExtension(a_Component.Type)
    Dim dlmt As String
    dlmt = "\"
    If Right(a_OutputDir, 1) = "\" Then dlmt = ""
    p = a_OutputDir & dlmt & f
    ' エクスポート
    '   - 例外で失敗したときはイミディエイトに表示
    ' 既存ファイルあり
    If m_fso.FileExists(p) Then
        ' コードを比較
        Dim projCode As String, repoCode As String
        projCode = DevUtils.ExtractProjectCode(a_Component)
        repoCode = DevUtils.ExtractCodeBody(p, DevUtils.HasCustomAttribute(p))
        ' 変更がない場合はエクスポートしなくて良い
        If Not DevUtils.HasChanged(projCode, repoCode) Then
            ' スキップ
            Debug.Print "Skipped: " & a_Component.Name
            m_SkippedCount = m_SkippedCount + 1
            Exit Sub
        End If
        On Error Resume Next
        ' 既存ファイルをポア
        Call m_fso.DeleteFile(p, True)
        ' 既存ファイルのポアに失敗 -> イミディエイトで通知 -> 握りつぶしてExit
        If Err.Number <> 0 Then
            Debug.Print "FAILED: existing `" & f & "` not deleted."
            Debug.Print "Number: " & Err.Number
            Debug.Print "Source: " & Err.Source
            Debug.Print "Desc  : " & Err.Description
            Call Err.Clear
            ' 失敗カウンタをインクリメント
            m_FailedCount = m_FailedCount + 1
            Exit Sub
        End If
        On Error GoTo 0
    End If
    On Error Resume Next
    ' モジュールをエクスポート
    Call a_Component.Export(p)
    ' エクスポートに失敗していたらイミディエイトに表示して握りつぶす
    If Err.Number <> 0 Then
        Debug.Print "FAILED: `" & f & "` not exported."
        Debug.Print "Number: " & Err.Number
        Debug.Print "Source: " & Err.Source
        Debug.Print "Desc  : " & Err.Description
        Call Err.Clear
        m_FailedCount = m_FailedCount + 1
    Else
        Debug.Print "Exported: " & f
        m_SucceededCount = m_SucceededCount + 1
    End If
    On Error GoTo 0
End Sub

Private Function OutputDirPath( _
            ByVal a_ComponentType As vbext_ComponentType, _
            ByVal a_RepositoryPath As String) As String
    Const ERR_SOURCE As String = SELF_MOD_NAME & ".OutputDirPath()"
    Dim ret As String
    
    Dim sf As String    ' sf: SubFolderの略
    Select Case a_ComponentType
        Case vbext_ct_StdModule: sf = STD_MOD
        Case vbext_ct_ClassModule: sf = CLS_MOD
        Case vbext_ct_MSForm: sf = FRM_MOD
        Case vbext_ct_Document: sf = DOC_MOD
        Case vbext_ct_ActiveXDesigner
            Call DevUtils.RaiseError( _
                ERR_NOT_SUPPORTED, ERR_SOURCE, "ActiveX Designerは非対応である。")
        Case Else
            Call DevUtils.RaiseError( _
                ERR_INVALID_ARGUMENT, ERR_SOURCE, "意味不明な引数が渡された。")
    End Select
    ' `a_RepositoryPath`のケツには必ず`\`が付いている
    ret = a_RepositoryPath & sf
    
    OutputDirPath = ret
End Function

Private Sub AA_Experiments(): End Sub
' =============================================================================
'   Experimental procedures
' =============================================================================
Private Sub Exp_ImportComponent_SkipIfNotChanged()
    ' モジュールインポート時、同内容ならスキップする
    Dim modPath As String
    modPath = ThisWorkbook.Path & "\.dev_helper\ForPullTest.bas"
    Call ImportComponent(modPath)
End Sub

Private Sub Exp_ImportComponent_PullNotExistingModule()
    Dim modPath As String
    modPath = ThisWorkbook.Path & "\.dev_helper\ForPullTest.bas"
    Call ImportComponent(modPath)
End Sub

Private Sub Exp_ImportComponent()
    ' モジュールをインポートする
    Dim modPath As String
    modPath = ThisWorkbook.Path & "\.dev_helper\ForPullTest.bas"
    Call ImportComponent(modPath)
End Sub

Private Sub Exp_ExportModules()
    ' プロジェクトの全モジュールをエクスポートする
    Call ExportComponents(ThisWorkbook.Path & "\" & REPO_NAME)
End Sub

Private Sub Exp_ExportModule()
    ' モジュールをエクスポートする
    Set m_fso = CreateObject("Scripting.FileSystemObject")
    Dim comp As Object, stdDir As String, f As String
    stdDir = _
        OutputDirPath(vbext_ct_StdModule, ThisWorkbook.Path & "\" & REPO_NAME)
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

