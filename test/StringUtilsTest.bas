Option Explicit

'--------------------------------------------------------------------------------
' StringUtils.basのテストをまとめた標準モジュール
'--------------------------------------------------------------------------------

Sub Test_Contains()
    Debug.Print "OKパターン：" & Contains("あいうえお", "あい")
    Debug.Print "OKパターン：" & Contains("あいうえお", "いうえ")
    Debug.Print "OKパターン：" & Contains("あいうえお", "えお")
    Debug.Print "OKパターン：" & Contains("あいうえお", "あいうえお")
    Debug.Print "OKパターン：" & Contains("あいうえお", "")
    Debug.Print "NGパターン：" & Contains("あいうえお", "あか")
End Sub

Sub Test_StartsWith()
    Debug.Print "OKパターン：" & StartsWith("あいうえお", "あい")
    Debug.Print "OKパターン：" & StartsWith("あいうえお", "あいうえお")
    Debug.Print "OKパターン：" & StartsWith("あいうえお", "")
    Debug.Print "NGパターン：" & StartsWith("あいうえお", "あか")
End Sub

Sub Test_EndsWith()
    Debug.Print "OKパターン：" & EndsWith("あいうえお", "えお")
    Debug.Print "OKパターン：" & EndsWith("あいうえお", "あいうえお")
    Debug.Print "OKパターン：" & EndsWith("あいうえお", "")
    Debug.Print "NGパターン：" & EndsWith("あいうえお", "えか")
End Sub

Sub Test_Compare()
    Debug.Print "OKパターン：" & Compare("アイウエオ", "ｱｲｳｴｵ")
    Debug.Print "OKパターン：" & Compare("ＡＢＣ", "abc")
    Debug.Print "OKパターン：" & Compare("123ＡＢＣｶﾞｷﾞｸﾞｹﾞｺﾞ", "１２３abcガギグゲゴ")
    Debug.Print "NGパターン：" & Compare("ｶﾞｷﾞｸﾞｹｺﾞ", "ガギグゲゴ")
End Sub
