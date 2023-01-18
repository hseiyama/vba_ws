Attribute VB_Name = "mdl_Main_RteModel"
Option Explicit

'外部定数
Public Const STR_RTE_FILE As String = "rte_model.c"

'内部定数
Private Const STR_PREFIX_BUS As String = "bus"

'内部変数
Private obj_SheetReader As cls_SheetReader
Private obj_TextWriter As cls_TextWriter

'コード生成処理
Public Sub Generate()
    '初期化処理
    Call Initialize
    'コード複製処理（ヘッダー部）
    Call CopyCodeHeader
    'RTE関数作成処理
    Call MakeRteFunc
    'コード複製処理（フッター部）
    Call CopyCodeFooter
    '終了処理
    Call Finalize
End Sub

'初期化処理
Private Sub Initialize()
    '前処理
    Set obj_SheetReader = New cls_SheetReader
    Set obj_TextWriter = New cls_TextWriter
    '各モジュールの初期化処理
    Call mdl_Input.Init
    Call obj_TextWriter.Init(mdl_Input.rng_RteCodePath.Value & "\" & STR_RTE_FILE, 1)
    Call obj_SheetReader.Init(sht_RteModel)
End Sub

'終了処理
Private Sub Finalize()
    '後処理
    Set obj_SheetReader = Nothing
    Set obj_TextWriter = Nothing
End Sub

'コード複製処理（ヘッダー部）
Private Sub CopyCodeHeader()
    'テキスト読込み処理
    Do While obj_SheetReader.ReadHeader
        'テキスト書込み処理
        Call obj_TextWriter.WriteText(obj_SheetReader.str_Text)
    Loop
End Sub

'コード複製処理（フッター部）
Private Sub CopyCodeFooter()
    'テキスト読込み処理
    Do While obj_SheetReader.ReadFooter
        'テキスト書込み処理
        Call obj_TextWriter.WriteText(obj_SheetReader.str_Text)
    Loop
End Sub

'RTE関数作成処理
Private Sub MakeRteFunc()
    Dim str_Comment As String
    Dim str_FnucText As String
    Dim str_MacroText As String
    'RTE関数の作成
    Do While mdl_Input.ReadText
        '各文字列の作成
        str_Comment = MakeComment
        str_FnucText = MakeFnucText
        str_MacroText = MakeMacroText
        'テキスト書込み処理
        Call obj_TextWriter.WriteText("/* " & str_Comment & " */")
        Call obj_TextWriter.WriteText("SdtType " & str_FnucText & " {")
        Call obj_TextWriter.WriteText("    " & str_MacroText & ";")
        Call obj_TextWriter.WriteText("    return STD_OK;")
        Call obj_TextWriter.WriteText("}")
        Call obj_TextWriter.WriteText("")
    Loop
End Sub

'コメント作成
Private Function MakeComment() As String
    MakeComment = mdl_Input.str_ModuleName & "(" & mdl_Input.str_DataName & ")"
End Function

'関数文字列作成
Private Function MakeFnucText() As String
    Dim str_Command As String
    Dim str_Param As String
    '各文字列の作成
    If mdl_Input.str_Attribute = mdl_Main_RteInfo.STR_ATTRIB_READ Then
        str_Command = "Read"
        str_Param = "*u"
    ElseIf mdl_Input.str_Attribute = mdl_Main_RteInfo.STR_ATTRIB_WRITE Then
        str_Command = "Write"
        If mdl_Input.str_Prefix = STR_PREFIX_BUS Then
            str_Param = "*u"
        Else
            str_Param = "u"
        End If
    Else
        str_Command = "Unknown"
        str_Param = "u"
    End If
    '関数文字列の作成
    MakeFnucText = mdl_Input.str_ModuleName & "_" & _
                str_Command & "_" & _
                mdl_Input.str_Prefix & "_g_" & mdl_Input.str_DataName & _
                mdl_Input.str_Prefix & "_g_" & mdl_Input.str_DataName & _
                "(" & mdl_Input.str_DataType & " " & str_Param & ")"
End Function

'マクロ文字列作成
Private Function MakeMacroText() As String
    Dim str_Command As String
    Dim str_Param As String
    '各文字列の作成
    If mdl_Input.str_Attribute = mdl_Main_RteInfo.STR_ATTRIB_READ Then
        str_Command = "read"
        str_Param = "u"
    ElseIf mdl_Input.str_Attribute = mdl_Main_RteInfo.STR_ATTRIB_WRITE Then
        str_Command = "write"
        If mdl_Input.str_Prefix = STR_PREFIX_BUS Then
            str_Param = "*u"
        Else
            str_Param = "u"
        End If
    Else
        str_Command = "unknown"
        str_Param = "u"
    End If
    'マクロ文字列の作成
    MakeMacroText = LCase(mdl_Input.str_ModuleName) & "_" & _
                str_Command & "_" & _
                mdl_Input.str_DataName & _
                "(" & str_Param & ")"
End Function
