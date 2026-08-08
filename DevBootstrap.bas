Attribute VB_Name = "DevBootstrap"
Option Explicit

' =============================================================================
'   dev_helper関係モジュールpush・pull用モジュール
' =============================================================================
' VBExtention
Public Enum vbext_CodePaneView
    vbext_cv_ProcedureView = 0
    vbext_cv_FullModuleView = 1
End Enum

Public Enum vbext_ComponentType
    vbext_ct_StdModule = 1
    vbext_ct_ClassModule = 2
    vbext_ct_MSForm = 3
    vbext_ct_ActiveXDesigner = 11
    vbext_ct_Document = 100
End Enum

Public Enum vbext_ProcKind
    vbext_pk_Proc = 0
    vbext_pk_Let = 1
    vbext_pk_Set = 2
    vbext_pk_Get = 3
End Enum

Public Enum vbext_ProjectProtection
    vbext_pp_none = 0
    vbext_pp_locked = 1
End Enum

Public Enum vbext_ProjectType
    vbext_pt_HostProject = 100
    vbext_pt_StandAlone = 101
End Enum

Public Enum vbext_RefKind
    vbext_rk_TypeLib = 0
    vbext_rk_Project = 1
End Enum

Public Enum vbext_VBAMode
    vbext_vm_Run = 0
    vbext_vm_Break = 1
    vbext_vm_Design = 2
End Enum

Public Enum vbext_WindowState
    vbext_ws_Normal = 0
    vbext_ws_Minimize = 1
    vbext_ws_Maximize = 2
End Enum

Public Enum vbext_WindowType
    vbext_wt_CodeWindow = 0
    vbext_wt_Designer = 1
    vbext_wt_Browser = 2
    vbext_wt_Watch = 3
    vbext_wt_Locals = 4
    vbext_wt_Immediate = 5
    vbext_wt_ProjectWindow = 6
    vbext_wt_PropertyWindow = 7
    vbext_wt_Find = 8
    vbext_wt_FindReplace = 9
    vbext_wt_Toolbox = 10
    vbext_wt_LinkedWindowFrame = 11
    vbext_wt_MainWindow = 12
    vbext_wt_ToolWindow = 15
End Enum

' FileSystemObject
Public Enum SpecialFolderConst
  WindowsFolder = 0
  SystemFolder = 1
  TemporaryFolder = 2
End Enum

Public Enum StandardStreamTypes
  StdIn = 0
  StdOut = 1
  StdErr = 2
End Enum

Public Enum IOMode
  ForReading = 1
  ForWriting = 2
  ForAppending = 8
End Enum

Public Enum Tristate
  TristateFalse = 0
  TristateMixed = -2
  TristateTrue = -1
  TristateUseDefault = -2
End Enum

Public Enum FileAttributes
  Normal = 0
  ReadOnly = 1
  Hidden = 2
  System = 4
  Volume = 8
  Directory = &H10
  Archive = &H20
  Alias = &H400
  Compressed = &H800
End Enum

Public Enum CutomErrorEnum
    ' 10000～: 引数・値
    ERR_INVALID_ARGUMENT = 10001
    ERR_INVALID_ENUM = 10002
    ERR_OUT_OF_RANGE = 10003
    ' 10100～: 状態
    ERR_INVALID_STATE = 10101
    ERR_NOT_SUPPORTED = 10102
    ERR_NOT_IMPLEMENTED = 10103
    ' 10200～: 検索
    ERR_NOT_FOUND = 10201
    ERR_DUPLICATED = 10202
    ' 10300～: ファイル
    ERR_FILE_NOT_FOUND = 10301
    ERR_FILE_ALREADY_EXISTS = 10302
    ERR_IO_ERROR = 10303
    ERR_INVALID_FILE = 10304
    ' 10400～: 権限
    ERR_PERMISSION_DENIED = 10401
    ' 10500～: 想定外
    ERR_INTERNAL_ERROR = 10501
End Enum

' 仕様を変更するときはここをメンテ
'   - リポジトリ用フォルダ名
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
' プロジェクト側のモジュールからコードを削除する
Public Sub DeleteCodeLines( _
            ByVal a_Component As Object)
    If a_Component.CodeModule.CountOfLines < 1 Then Exit Sub
    Call a_Component.CodeModule.DeleteLines( _
        1, a_Component.CodeModule.CountOfLines)
End Sub

' プロジェクトのモジュール（CodeModuleオブジェクト）からコードを抜き出す
Public Function ExtractProjectCode( _
            ByVal a_Component As Object) As String
    Dim ret As String
    ret = ""
    
    Dim cm As Object
    Set cm = a_Component.CodeModule
    If cm.CountOfLines < 1 Then GoTo Finally
    
    ret = cm.lines(1, cm.CountOfLines)
Finally:
    ExtractProjectCode = ret
End Function
    
' コードを比較する
Public Function HasChanged( _
            ByVal a_ProjectCode As String, _
            ByVal a_RepositoryCode As String) As Boolean
    Dim a As String, b As String
    a = NormalizeCode(a_ProjectCode)
    b = NormalizeCode(a_RepositoryCode)
    HasChanged = (a <> b)
End Function

' コードを正規化する
Public Function NormalizeCode( _
            ByVal a_Code As String) As String
    Dim ret As String
    
    ' 両端のスペースをトリム
    ret = Trim(a_Code)
    ' 改行を統一教会する
    ret = Replace(ret, vbCrLf, vbLf)
    ret = Replace(ret, vbCr, vbLf)
    ' 末尾の改行を削除
    Do While (Right$(ret, 1) = vbLf)
        ret = Left$(ret, Len(ret) - 1)
    Loop
    ' 両端のスペースをトリム
    ret = Trim(ret)
    
    NormalizeCode = ret
End Function

' モジュールにカスタム属性があるかどうか判定する
' モジュールから、コードの正味の部分のみ取り出す
Public Function HasCustomAttribute( _
            ByVal a_ComponentPath As String) As Boolean
    Dim ret As Boolean
    ret = False
    
    Dim conts As String, startLn As Long
    conts = LoadComponentCode(a_ComponentPath)
    startLn = FindCodeStartLine(a_ComponentPath)
    Dim codeArr As Variant
    codeArr = Split(conts, vbCrLf)
    
    Dim i As Long
    For i = startLn - 1 To UBound(codeArr)
        If Left$(codeArr(i), Len("Attribute ")) = "Attribute " Then
            ret = True
            GoTo Finally
        End If
    Next
Finally:
    HasCustomAttribute = ret
End Function

' モジュールから、コードの正味の部分のみ取り出す
'   - カスタム属性を除いた比較をするときは第2引数をTrueにする
Public Function ExtractCodeBody( _
            ByVal a_ComponentPath As String, _
   Optional ByVal a_SkipAttribute As Boolean = False) As String
    Dim ret As String
    
    Dim conts As String, startLn As Long
    conts = LoadComponentCode(a_ComponentPath)
    startLn = FindCodeStartLine(a_ComponentPath)
    Dim codeArr As Variant
    codeArr = Split(conts, vbCrLf)
    
    Dim i As Long, s As String
    For i = startLn - 1 To UBound(codeArr)
        ' カスタム属性スキップモード時は`Attribute `で始まる行をスキップ
        s = codeArr(i)
        If a_SkipAttribute _
            And Left$(s, Len("Attribute ")) = "Attribute " Then GoTo Continue
        If ret <> "" Then ret = ret & vbCrLf
        ret = ret & s
Continue:
    Next
    
    ExtractCodeBody = ret
End Function

' モジュールから取り出したコードの正味の開始位置（行）を取得する
'   - `Attribute`で始まる行の最後の行（とそれに続く空行）までが
'     ヘッダ情報部分
Public Function FindCodeStartLine( _
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
Public Function LoadComponentCode( _
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
Public Function FindComponent( _
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
Public Function ExtractModuleName( _
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

Public Function ComponentExtension( _
            ByVal a_ComponentType As vbext_ComponentType) As String
    ' モジュールの種類に応じた拡張子を返す
    Const ERR_SOURCE As String = "DevProj.GetExtension()"
    Dim ret As String
    Select Case a_ComponentType
        Case vbext_ct_StdModule: ret = ".bas"
        Case vbext_ct_ClassModule: ret = ".cls"
        Case vbext_ct_MSForm: ret = ".frm"
        Case vbext_ct_Document: ret = ".cls"
        Case vbext_ct_ActiveXDesigner
            Call RaiseError( _
                ERR_NOT_SUPPORTED, ERR_SOURCE, "ActiveX Designer はサポート対象外である。")
        Case Else
            Call RaiseError( _
                ERR_INVALID_ARGUMENT, ERR_SOURCE, "不正な引数が渡された。")
    End Select
    ComponentExtension = ret
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

Private Sub AA_Experiments(): End Sub
' =============================================================================
'   Experimental procedures
' =============================================================================
Private Sub Exp_HasCustomAttribute()
    Dim modPath As String
    modPath = ThisWorkbook.Path & "\.dev_helper\ForPullTest.bas"
    ' ↓こっちのパスは常にあるとは限らないので、動作確認後はコメントアウト
'    modPath = ThisWorkbook.Path & "\repo\cls_mod\Dictionary.cls"
'     modPath = ThisWorkbook.Path & "\repo\doc_mod\Sh01Data.cls"
    Debug.Print HasCustomAttribute(modPath)
End Sub

Private Sub Exp_HasChanged()
    Dim a As String, b As String
    a = "Private Sub Foo()" & vbCrLf & _
        vbTab & "Debug.Print ""ち～ん（笑）""" & vbCrLf & _
        "End Sub"
    b = "   Private Sub Foo()" & vbCrLf & _
        vbTab & "Debug.Print ""ち～ん（笑）""" & vbCrLf & _
        "End Sub  " & vbCrLf & vbCrLf & vbLf
    ' aとbは一致判定のはず
    Debug.Assert Not HasChanged(a, b)
    
    a = "Private Sub Foo()" & vbCrLf & _
        vbTab & "Debug.Print ""ち～ん（笑）""" & vbCrLf & _
        "End Sub"
    b = "   Private Sub Foo()" & vbCrLf & _
        vbTab & "Debug.Print ""ぢ～ん（笑）""" & vbCrLf & _
        "End Sub  " & vbCrLf & vbCrLf & vbLf
    ' aとbは不一致判定のはず
    Debug.Assert HasChanged(a, b)
    
    ' 空文字列の場合
    Debug.Assert Not HasChanged("", "")
    Debug.Assert HasChanged("", "Option Explicit")
    
    ' 次の検証用コードは開発用`dev_helper.xlsm`本体でしか再現しない
    '   -> 検証が終わったらコメントアウト
'    Dim modPath As String, modName As String
'    modPath = ThisWorkbook.Path & "\repo\cls_mod\SqlQueryDef.cls"
'    modName = ExtractModuleName(modPath)
'    Dim comp As Object
'    Set comp = FindComponent(modName)
'    Dim projCode As String, repoCode As String
'    projCode = ExtractProjectCode(comp)
'    repoCode = ExtractCodeBody(modPath) ' <- カスタムAttributeごと取得
'    ' 不一致のはず
'    Debug.Assert HasChanged(projCode, repoCode)
'    repoCode = ExtractCodeBody(modPath, True) ' <- カスタムAttribute除外
'    ' 一致するはず
'    Debug.Assert Not HasChanged(projCode, repoCode)
    
    Debug.Print "Done!!"
End Sub

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

