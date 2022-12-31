Attribute VB_Name = "mdl_Main"
Option Explicit

'内部定数
Private Const STR_ATTRIB_READ As String = "SWC_I/F入力"
Private Const STR_ATTRIB_WRITE As String = "SWC_I/F出力"
Private Const STR_PREFIX_BUS As String = "bus"

'内部変数
Dim obj_Header As cls_Reader
Dim obj_Footer As cls_Reader

'コード生成処理
Public Sub GenerateCode()
    '前処理
    Set obj_Header = New cls_Reader
    Set obj_Footer = New cls_Reader
    '初期化処理
    Initialize
    'コードコピー処理（ヘッダー部）
    CopyCode obj_Header
    'RTE関数作成処理
    MakeRteFunc
    'コードコピー処理（フッター部）
    CopyCode obj_Footer
    '後処理
    Set obj_Header = Nothing
    Set obj_Footer = Nothing
End Sub

'初期化処理
Private Sub Initialize()
    '各モジュールの初期化処理
    mdl_Input.Init
    mdl_Output.Init
    obj_Header.Init sht_Header
    obj_Footer.Init sht_Footer
End Sub

'コードコピー処理
Private Sub CopyCode(ByRef obj_Reader As cls_Reader)
    'テキスト読込み処理
    Do While obj_Reader.ReadText
        'テキスト書込み処理
        mdl_Output.WriteText obj_Reader.str_Text
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
        mdl_Output.WriteText "/* " & str_Comment & " */"
        mdl_Output.WriteText "SdtType " & str_FnucText & " {"
        mdl_Output.WriteText "    " & str_MacroText & ";"
        mdl_Output.WriteText "    return STD_OK;"
        mdl_Output.WriteText "}"
        mdl_Output.WriteText ""
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
    If mdl_Input.str_Attribute = STR_ATTRIB_READ Then
        str_Command = "Read"
        str_Param = "*u"
    ElseIf mdl_Input.str_Attribute = STR_ATTRIB_WRITE Then
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
    If mdl_Input.str_Attribute = STR_ATTRIB_READ Then
        str_Command = "read"
        str_Param = "u"
    ElseIf mdl_Input.str_Attribute = STR_ATTRIB_WRITE Then
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
