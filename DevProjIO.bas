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

' エクスポート成否カウント用
Private m_SucceededCount As Long
Private m_FailedCount As Long

Private m_fso As Object
' =============================================================================
'   公開API
' =============================================================================
Public Sub ExportComponentsToRepository()
    Set m_fso = CreateObject("Scripting.FileSystemObject")
    
    ' SharePoint上のファイルのときはフラグを立てる
    If Left(ThisWorkbook.Path, 5) = "https" Then
        m_OnSharePoint = True
    Else
        m_OnSharePoint = False
    End If
    
    Dim repoDir As String
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
        repoDir = parentDir & "\" & projName & "\"
        If Not m_fso.FolderExists(repoDir) Then _
            Call m_fso.CreateFolder(repoDir)
    Else
        repoDir = ThisWorkbook.Path & "\" & REPO_NAME & "\"
    End If
    
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
    ' カウンタを初期化
    m_SucceededCount = 0
    m_FailedCount = 0
    Call ExportComponents(repoDir)
    Debug.Print String(40, "-")
    Debug.Print "Summary: "
    Debug.Print vbTab & "Push succeeded: " & CStr(m_SucceededCount) & " module(s)."
    Debug.Print vbTab & "Push failed   : " & CStr(m_FailedCount) & " module(s)."
End Sub

Private Sub AA_HelperFunctions(): End Sub
' =============================================================================
'   開発者向けヘルパ
' =============================================================================
Private Sub ImportComponent( _
            ByVal a_ComponentPath As String)
    ' モジュール名を取得
    Dim modName As String
    modName = ExtractModuleName(a_ComponentPath)
    Dim comp As Object
    Set comp = FindComponent(modName)
    ' プロジェクト内に同名モジュールがない -> そのままpull
    If comp Is Nothing Then
        Call ThisWorkbook.VBProject.VBComponents.Import(a_ComponentPath)
    ' プロジェクト内に同名モジュールあり -> コードのみpull
    Else
        
        
    End If
End Sub

' プロジェクト側のモジュールからコードを削除する
Private Sub DeleteCodeLines( _
            ByVal a_Component As Object)
    Call a_Component.CodeModule.DeleteLines( _
        1, a_Component.CodeModule.CountOfLines)
End Sub

' モジュールから、コードの正味の部分のみ取り出す
Private Function ExtractCodeBody( _
            ByVal a_ComponentPath As String) As String
    Dim ret As String
    
    Dim conts As String, startLn As Long
    conts = LoadComponentCode(a_ComponentPath)
    startLn = FindCodeStartLine(a_ComponentPath)
    Dim codeArr As Variant
    codeArr = Split(conts, vbCrLf)
    
    Dim i As Long
    For i = startLn - 1 To UBound(codeArr)
        If ret <> "" Then ret = ret & vbCrLf
        ret = ret & codeArr(i)
    Next
    
    ExtractCodeBody = ret
End Function

' モジュールから取り出したコードの正味の開始位置（行）を取得する
'   - `Attribute`で始まる行の最後の行（とそれに続く空行）までが
'     ヘッダ情報部分
Private Function FindCodeStartLine( _
            ByVal a_ComponentPath As String) As Long
    Const ERR_SOURCE As String = SELF_MOD_NAME & ".FindCodeStartLine()"
    Dim ret As Long
    ret = 0
    
    Dim conts As String
    conts = LoadComponentCode(a_ComponentPath)
    Dim i As Long, codeArr As Variant
    codeArr = Split(conts, vbCrLf)
    Dim foundAttr As Boolean
    foundAttr = False
    For i = LBound(codeArr) To UBound(codeArr)
        If Left$(codeArr(i), Len("Attribute ")) = "Attribute " Then
            foundAttr = True
            GoTo Continue
        End If
        If Not foundAttr Then GoTo Continue
        ' 空行はスキップ
        If Len(Trim$(codeArr(i))) = 0 Then GoTo Continue
        ' ここに来た時点で`Attribute`で始まる行を1つは通過済み
        ' かつ`Attribute `で始まらない
        ' かつ空行でもない
        '   -> この行から正味のコード開始
        ' Split()で作った配列は`0`始まりなので`1`を足して補正
        ret = i + 1
        ' ここで行番号を返却
        GoTo Finally
Continue:
    Next
    ' `Attribute`で始まる行がなかった -> 異常 -> 例外スロー
    If Not foundAttr Then _
        Call RaiseError( _
            ERR_INVALID_FILE, _
            ERR_SOURCE, _
            "`Attribute`がない。VBAのモジュールファイルではないようだ。")
    ' 返り値が`1`以下になる
    If ret <= 1 Then _
        Call RaiseError( _
            ERR_INVALID_FILE, _
            ERR_SOURCE, _
            "コードの開始位置が1行目以下であるはずがない。")
Finally:
    FindCodeStartLine = ret
End Function

' 指定したモジュールファイルから、純粋なコード部分のみ取り出す
Private Function LoadComponentCode( _
            ByVal a_ComponentPath As String) As String
    Const ERR_SOURCE As String = SELF_MOD_NAME & ".LoadComponentCode()"
    Dim ret As String
    
    On Error GoTo HandleError:
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim ts As Object
    Set ts = fso.OpenTextFile(a_ComponentPath, ForReading)
    ret = ts.ReadAll()
    GoTo Finally
    
    Dim errNum As Long, errSrc As String, errDesc As String
    errNum = 0
HandleError:
    errNum = Err.Number
    errSrc = Err.Source
    errDesc = Err.Description
    Debug.Print "Load """ & a_ComponentPath & """ failed: "
    Debug.Print vbTab & "Number: " & errNum
    Debug.Print vbTab & "Source: " & errSrc
    Debug.Print vbTab & "Desc  : " & errDesc
Finally:
    If Not ts Is Nothing Then
        Call ts.Close
        Set ts = Nothing
    End If
    ' 例外発生時は呼び出し元に再スロー
    If errNum <> 0 Then Call Err.Raise(errNum, errSrc, errDesc)
    LoadComponentCode = ret
End Function
            

' 指定した名前のコンポーネントを取得する（なければNothing）
Private Function FindComponent( _
            ByVal a_ComponentName As String) As Object
    Dim ret As Object
    Set ret = Nothing
    
    Dim comp As Object
    For Each comp In ThisWorkbook.VBProject.VBComponents
        If comp.Name = a_ComponentName Then
            Set ret = comp
            GoTo Finally
        End If
    Next
    
Finally:
    Set FindComponent = ret
End Function

' モジュールファイルからモジュール名を取り出す
Private Function ExtractModuleName( _
            ByVal a_ComponentPath As String) As String
    Const ERR_SOURCE As String = SELF_MOD_NAME & ".ExtractModuleName"
    Dim ret As String
    
    Dim f As Integer
    f = FreeFile()
    
    Dim ln As String
    Dim p1 As Long, p2 As Long
    
    Open a_ComponentPath For Input As #f
    
    Do Until EOF(f)
        Line Input #f, ln
        
        If Left$(ln, Len("Attribute VB_Name")) = "Attribute VB_Name" Then
            ' 先頭から`"`を探す
            p1 = InStr(ln, """")
            ' 末尾から`"`を探す
            p2 = InStrRev(ln, """")
            
            ' `"`がない or `"`が閉じられていない -> 異常
            '   -> 例外スロー
            If p1 = 0 Or p1 = p2 Then
                Close #f
                Call RaiseError( _
                    ERR_INVALID_ARGUMENT, _
                    ERR_SOURCE, _
                    "VB_Name属性が壊れている。")
            End If
            ' `""`で囲まれた中身（＝モジュール名）をスライス
            ret = Mid$(ln, p1 + 1, p2 - p1 - 1)
            
            ' ファイルを閉じてモジュール名を返す
            Close #f
            GoTo Finally
        End If
    Loop
    
    ' ここへ到達 -> モジュール名が取得できなかった -> 異常事態
    '   -> 例外スロー
    Close #f
    Call RaiseError( _
        ERR_NOT_FOUND, _
        ERR_SOURCE, _
        "モジュール名が見つからなかった。")
    
Finally:
    ExtractModuleName = ret
End Function

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
        ' 失敗カウンタをインクリメント
        m_FailedCount = m_FailedCount + 1
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
            Call RaiseError( _
                ERR_NOT_SUPPORTED, ERR_SOURCE, "ActiveX Designerは非対応である。")
        Case Else
            Call RaiseError( _
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
Private Sub Exp_DeleteCodeLines()
    Dim comp As Object
    Set comp = FindComponent("ForPullTest")
    'Set comp = FindComponent("SOEntry")
    Call DeleteCodeLines(comp)
End Sub

Private Sub Exp_ExtractCodeBody()
    Dim modPath As String
    modPath = ThisWorkbook.Path & "\.dev_helper\ForPullTest.bas"
    Debug.Print ExtractCodeBody(modPath)
End Sub

Private Sub Exp_FindCodeStartLine()
    Dim modPath As String
    modPath = ThisWorkbook.Path & "\.dev_helper\ForPullTest.bas"
    Debug.Print FindCodeStartLine(modPath)
End Sub

Private Sub Exp_LoadComponentCode()
    Dim modPath As String
    modPath = ThisWorkbook.Path & "\.dev_helper\ForPullTest.bas"
    Debug.Print LoadComponentCode(modPath)
End Sub

Private Sub Exp_ImportComponent_PullNotExistingModule()
    Dim modPath As String
    modPath = ThisWorkbook.Path & "\.dev_helper\ForPullTest.bas"
    Call ImportComponent(modPath)
End Sub

Private Sub Exp_FindComponent()
    Dim comp1 As Object, comp2 As Object
    ' 存在するモジュール
    Set comp1 = FindComponent("DevBootstrap")
    Debug.Print comp1.Name
    Debug.Assert comp1.Name = "DevBootstrap"
    ' 存在しないモジュール
    Set comp2 = FindComponent("pachinko123")
    Debug.Assert comp2 Is Nothing
End Sub

Private Sub Exp_ExtractModuleName()
    ' モジュール名を抽出する
    Dim modPath As String
    modPath = ThisWorkbook.Path & "\.dev_helper\DevBootstrap.bas"
    Dim modName As String
    modName = ExtractModuleName(modPath)
    Debug.Print modName
    Debug.Assert modName = "DevBootstrap"
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

