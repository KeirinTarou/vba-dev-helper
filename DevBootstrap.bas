Attribute VB_Name = "DevBootstrap"
Option Explicit

' =============================================================================
'   dev_helper関係モジュールpush・pull用モジュール
' =============================================================================

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
    ' dev_helperプロジェクト関係のモジュールを所定フォルダに全push
    Set m_fso = CreateObject("Scripting.FileSystemObject")
    
    Dim repoDir As String
    repoDir = ThisWorkbook.Path & DEV_HELPER_REPO_NAME
    ' `.dev_helper`フォルダがなかったら作る
    If Not m_fso.FolderExists(repoDir) Then
        Call m_fso.CreateFolder(repoDir)
    End If
    Dim comp As Object, cn As String, f As String, p As String
    ' 接頭辞`Dev`のもののみリポジトリに送り込む
    For Each comp In ThisWorkbook.VBProject.VBComponents
        cn = comp.Name
        If Not IsDevHelper(cn) Then GoTo Continue
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
    Set m_fso = CreateObject("Scripting.FileSystemObject")
    ' 一旦、dev_helper関係モジュールをポア
    Call RemoveDevModules
    Dim repoDir As String
    repoDir = ThisWorkbook.Path & DEV_HELPER_REPO_NAME
    Dim f As Object, ext As String, bs As String
    For Each f In m_fso.GetFolder(repoDir).Files
        ext = LCase$(m_fso.GetExtensionName(f.Path))
        bs = m_fso.GetBaseName(f.Path)
        ' `.bas`、`.cls`以外はスルー
        If Not (ext = "bas" Or ext = "cls") Then GoTo Continue
        ' dev_helper以外はスルー
        If Not IsDevHelper(bs) Then GoTo Continue
        ' このモジュールもスルー
        If bs = SELF_MOD_NAME Then GoTo Continue
        ' ここまで来たらインポート
        Call ThisWorkbook.VBProject.VBComponents.Import(f.Path)
        Debug.Print "Imported: " & f.Name
Continue:
    Next
End Sub

Private Sub AA_HelperFunctions(): End Sub
' =============================================================================
'   Helper Functions
' =============================================================================
Private Sub RemoveDevModules()
    ' dev_helper関係モジュール（本モジュール以外）をポアする
    '   - Documentモジュール、フォームモジュールは一旦除外
    Dim i As Long, comp As Object, cn As String
    With ThisWorkbook.VBProject.VBComponents
        ' コレクション要素の削除になるので逆順ループ
        For i = .Count To 1 Step -1
            Set comp = .Item(i)
            cn = comp.Name
            ' Documentモジュール、フォームモジュールはスキップ
            If comp.Type = vbext_ct_ActiveXDesigner Then GoTo Continue
            If comp.Type = vbext_ct_Document Then GoTo Continue
            If comp.Type = vbext_ct_MSForm Then GoTo Continue
            ' このモジュールはスキップ
            If cn = SELF_MOD_NAME Then GoTo Continue
            ' dev_helper以外はスキップ
            If Not IsDevHelper(cn) Then GoTo Continue
            ' ここまで来たらポアしても良い
            Call .Remove(comp)
Continue:
        Next
    End With
End Sub

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

