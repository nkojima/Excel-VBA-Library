Option Explicit

'--------------------------------------------------------------------------------
' FileUtilsのテストをまとめた標準モジュール
'--------------------------------------------------------------------------------

Sub Test_ExistsFile()
    Debug.Print "----- ExistsFile() -----"
    Debug.Print "C:\Windows --> " & FileUtils.ExistsFile("C:\Windows")
    Debug.Print "C:\Windows2 --> " & FileUtils.ExistsFile("C:\Windows2")
    Debug.Print "C:\Windows\ --> " & FileUtils.ExistsFile("C:\Windows\")
    Debug.Print "C:\Windows\system.ini --> " & FileUtils.ExistsFile("C:\Windows\system.ini")
    Debug.Print "C:\Windows\system2.ini --> " & FileUtils.ExistsFile("C:\Windows\system2.ini")
End Sub

Sub Test_ExistsFolder()
    Debug.Print "----- ExistsFolder() -----"
    Debug.Print "C:\Windows --> " & FileUtils.ExistsFolder("C:\Windows")
    Debug.Print "C:\Windows2 --> " & FileUtils.ExistsFolder("C:\Windows2")
    Debug.Print "C:\Windows\ --> " & FileUtils.ExistsFolder("C:\Windows\")
    Debug.Print "C:\Windows\system.ini --> " & FileUtils.ExistsFolder("C:\Windows\system.ini")
    Debug.Print "C:\Windows\system2.ini --> " & FileUtils.ExistsFolder("C:\Windows\system2.ini")
End Sub

Sub Test_GetExtensionName()
    Debug.Print "----- GetExtensionName() -----"
    Debug.Print "C:\Windows --> " & FileUtils.GetExtensionName("C:\Windows")
    Debug.Print "C:\Windows2 --> " & FileUtils.GetExtensionName("C:\Windows2")
    Debug.Print "C:\Windows\ --> " & FileUtils.GetExtensionName("C:\Windows\")
    Debug.Print "C:\Windows\system.ini --> " & FileUtils.GetExtensionName("C:\Windows\system.ini")
    Debug.Print "C:\Windows\system2.ini --> " & FileUtils.GetExtensionName("C:\Windows\system2.ini")
End Sub