Option Explicit

'--------------------------------------------------------------------------------
' StringUtilsのテストをまとめた標準モジュール
'--------------------------------------------------------------------------------

Sub Test_Contains()
    Debug.Print "OKパターン：" & StringUtils.Contains("あいうえお", "あい")
    Debug.Print "OKパターン：" & StringUtils.Contains("あいうえお", "いうえ")
    Debug.Print "OKパターン：" & StringUtils.Contains("あいうえお", "えお")
    Debug.Print "OKパターン：" & StringUtils.Contains("あいうえお", "あいうえお")
    Debug.Print "OKパターン：" & StringUtils.Contains("あいうえお", "")
    Debug.Print "NGパターン：" & StringUtils.Contains("あいうえお", "あか")
End Sub

Sub Test_StartsWith()
    Debug.Print "OKパターン：" & StringUtils.StartsWith("あいうえお", "あい")
    Debug.Print "OKパターン：" & StringUtils.StartsWith("あいうえお", "あいうえお")
    Debug.Print "OKパターン：" & StringUtils.StartsWith("あいうえお", "")
    Debug.Print "NGパターン：" & StringUtils.StartsWith("あいうえお", "あか")
End Sub

Sub Test_EndsWith()
    Debug.Print "OKパターン：" & StringUtils.EndsWith("あいうえお", "えお")
    Debug.Print "OKパターン：" & StringUtils.EndsWith("あいうえお", "あいうえお")
    Debug.Print "OKパターン：" & StringUtils.EndsWith("あいうえお", "")
    Debug.Print "NGパターン：" & StringUtils.EndsWith("あいうえお", "えか")
End Sub

Sub Test_Compare()
    Debug.Print "OKパターン：" & StringUtils.Compare("アイウエオ", "ｱｲｳｴｵ")
    Debug.Print "OKパターン：" & StringUtils.Compare("ＡＢＣ", "abc")
    Debug.Print "OKパターン：" & StringUtils.Compare("123ＡＢＣｶﾞｷﾞｸﾞｹﾞｺﾞ", "１２３abcガギグゲゴ")
    Debug.Print "NGパターン：" & StringUtils.Compare("ｶﾞｷﾞｸﾞｹｺﾞ", "ガギグゲゴ")
End Sub