Attribute VB_Name = "mdl_RteInfo"
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

'初期化処理
Public Sub Init()
    '変数の初期化
    int_Range_Row = 1
    int_Row = INT_ROW_START
    bln_First = False
End Sub

'RTE情報設定処理
Public Sub SetRteInfomation(ByRef str_FileName, ByRef obj_Range As Range)
    Dim obj_Book As Workbook
    Dim obj_Sheet As Worksheet
    'シート範囲のクリア
    obj_Range.ClearContents
    'RTE情報ファイルの全情報を検索
    Set obj_Book = Workbooks.Open(str_FileName)
    For Each obj_Sheet In obj_Book.Worksheets
        '対象シートの判別
        If obj_Sheet.Range(STR_RNG_SHEET_CHECK) = STR_SHEET_CHECK Then
            int_Row = INT_ROW_START
            bln_First = False
            'RTE情報検索処理
            Do While SearchRteInfo(obj_Range, obj_Sheet)
            Loop
        End If
    Next
    obj_Book.Close
End Sub

'RTE情報検索処理
Private Function SearchRteInfo(ByRef obj_Range As Range, ByRef obj_Sheet As Worksheet) As Boolean
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
        SearchRteInfo = True
        '次の行へ移動
        int_Row = int_Row + 1
    Else
        SearchRteInfo = False
    End If
End Function
