Attribute VB_Name = "mdl_Main_RteInfo"
Option Explicit

'外部定数
Public Const STR_ATTRIB_READ As String = "SWC_I/F入力"
Public Const STR_ATTRIB_WRITE As String = "SWC_I/F出力"

'内部定数
Private Const INT_ROW_START As Integer = 3
Private Const INT_COL_ATTRIBUTE As Integer = 3
Private Const INT_COL_DATA_TYPE As Integer = 4
Private Const INT_COL_DATA_NAME As Integer = 5
Private Const INT_COL_DESCRIPTION As Integer = 6
Private Const STR_RNG_SHEET_CHECK As String = "E1"
Private Const STR_RNG_MODULE_NAME As String = "D2"
Private Const STR_SHEET_CHECK As String = "SWC_I/F情報"

'内部変数
Private int_Range_Row As Integer
Private int_Row As Integer
Private bln_First As Boolean

'RTE情報取得処理
Public Sub Collect()
    '初期化処理
    Call Initialize
    'RTE情報設定処理
    Call SetRteInfomation
End Sub

'初期化処理
Private Sub Initialize()
    '各モジュールの初期化処理
    Call mdl_Input.Init
    '変数の初期化
    int_Range_Row = 1
    int_Row = INT_ROW_START
    bln_First = False
End Sub

'RTE情報設定処理
Private Sub SetRteInfomation()
    Dim obj_Book As Workbook
    Dim obj_Sheet As Worksheet
    'シート範囲のクリア
    mdl_Input.rng_RteInfoList.ClearContents
    'RTE情報ファイルの全情報を検索
    Set obj_Book = Workbooks.Open(mdl_Input.rng_RteInfoFile.Value)
    For Each obj_Sheet In obj_Book.Worksheets
        '対象シートの判別
        If obj_Sheet.Range(STR_RNG_SHEET_CHECK) = STR_SHEET_CHECK Then
            int_Row = INT_ROW_START
            bln_First = False
            'RTE情報検索処理
            Do While SearchRteInfo(obj_Sheet, mdl_Input.rng_RteInfoList)
            Loop
        End If
    Next
    obj_Book.Close
End Sub

'RTE情報検索処理
Private Function SearchRteInfo(ByRef obj_Sheet As Worksheet, ByRef obj_Range As Range) As Boolean
    '各項目の読込み
    If obj_Sheet.Cells(int_Row, INT_COL_ATTRIBUTE).Value = STR_ATTRIB_READ _
    Or obj_Sheet.Cells(int_Row, INT_COL_ATTRIBUTE).Value = STR_ATTRIB_WRITE Then
        If bln_First Then
            obj_Range.Cells(int_Range_Row, 1).Value = "↑"
        Else
            obj_Range.Cells(int_Range_Row, 1).Value = obj_Sheet.Range(STR_RNG_MODULE_NAME).Value
            bln_First = True
        End If
        obj_Range.Cells(int_Range_Row, 2).Value = obj_Sheet.Cells(int_Row, INT_COL_ATTRIBUTE).Value
        obj_Range.Cells(int_Range_Row, 3).Value = obj_Sheet.Cells(int_Row, INT_COL_DATA_TYPE).Value
        obj_Range.Cells(int_Range_Row, 4).Value = obj_Sheet.Cells(int_Row, INT_COL_DATA_NAME).Value
        obj_Range.Cells(int_Range_Row, 5).Value = obj_Sheet.Cells(int_Row, INT_COL_DESCRIPTION).Value
        int_Range_Row = int_Range_Row + 1
    End If
    '検索終了の判定
    If obj_Sheet.Cells(int_Row, INT_COL_ATTRIBUTE).Value <> "END" Then
        '次の行へ移動
        int_Row = int_Row + 1
        '戻り値の設定
        SearchRteInfo = True
    Else
        '戻り値の設定
        SearchRteInfo = False
    End If
End Function
