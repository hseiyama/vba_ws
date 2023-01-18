Attribute VB_Name = "mdl_Main_RteCompiler"
Option Explicit

'外部定数
Public Const STR_RTE_FILE As String = "Rte_Compiler_Cfg.h"

'内部変数
Private obj_SheetReader As cls_SheetReader
Private obj_TextWriter As cls_TextWriter

'コード生成処理
Public Sub Generate()
    '初期化処理
    Call Initialize
    'コード複製処理（ヘッダー部）
    Call CopyCodeHeader
    'DEF作成処理
    Call MakeDefine
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
    Call obj_SheetReader.Init(sht_RteCompiler)
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

'DEF作成処理
Private Sub MakeDefine()
    Dim str_ModuleNamePre As String
    '変数の初期化
    str_ModuleNamePre = ""
    'RTE関数の作成
    Do While mdl_Input.ReadText
        If mdl_Input.str_ModuleName <> str_ModuleNamePre Then
            'テキスト書込み処理
            Call obj_TextWriter.WriteText("/* " & mdl_Input.str_ModuleName & "定義 */")
            Call obj_TextWriter.WriteText("#define " & mdl_Input.str_ModuleName & "_CODE")
        End If
        str_ModuleNamePre = mdl_Input.str_ModuleName
    Loop
End Sub
