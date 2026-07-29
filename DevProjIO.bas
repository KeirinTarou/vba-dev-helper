Attribute VB_Name = "DevProjIO"
Option Explicit

Private m_fso As Object


Private Sub AA_DevHelpers(): End Sub
' 開発者向けヘルパ

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

Private Sub AA_HelperFunctions(): End Sub

Private Function ComponentExtension( _
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

Private Function ComponentFolderName( _
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

Private Function ComponentTypeName( _
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

Private Sub RaiseError( _
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
