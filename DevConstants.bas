Attribute VB_Name = "DevConstants"
Option Explicit

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
    ' 10000Å`: à¯êîÅEíl
    ERR_INVALID_ARGUMENT = 10001
    ERR_INVALID_ENUM = 10002
    ERR_OUT_OF_RANGE = 10003
    ' 10100Å`: èÛë‘
    ERR_INVALID_STATE = 10101
    ERR_NOT_SUPPORTED = 10102
    ERR_NOT_IMPLEMENTED = 10103
    ' 10200Å`: åüçı
    ERR_NOT_FOUMD = 10201
    ERR_DUPLICATED = 10202
    ' 10300Å`: ÉtÉ@ÉCÉã
    ERR_FILE_NOT_FOUND = 10301
    ERR_FILE_ALREADY_EXISTS = 10302
    ERR_IO_ERROR = 10303
    ' 10400Å`: å†å¿
    ERR_PERMISSION_DENIED = 10401
    ' 10500Å`: ëzíËäO
    ERR_INTERNAL_ERROR = 10501
End Enum
