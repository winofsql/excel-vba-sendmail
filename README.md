# excel-vba-sendmail ( ロリポップ )

![image](https://user-images.githubusercontent.com/1501327/163671327-6e3e4965-ebc2-43f1-b476-6ec033b8505f.png)

![image](https://user-images.githubusercontent.com/1501327/163672133-ff8b7256-343e-49cc-8434-1a8b2b84f56b.png)

![image](https://user-images.githubusercontent.com/1501327/163672223-0885c45f-912a-4a2b-bd19-c25ee1d32e3f.png)

![image](https://user-images.githubusercontent.com/1501327/163672246-3e69edbc-9435-4a14-9e04-7756a17f202f.png)


```vba
    ' ***********************************************************
    ' Windows 標準オブジェクト
    ' ***********************************************************
    Set Cdo = CreateObject("CDO.Message")
    
    ' ***********************************************************
    ' 自分のアドレスと宛先
    ' ***********************************************************
    Cdo.From = "アカウントメールアドレス"
    Cdo.To = "宛先メールアドレス"

    ' ***********************************************************
    ' 件名と本文
    ' ***********************************************************
    Cdo.Subject = "件名の文字列 / " & Now()
    Cdo.Textbody = "テキスト本文" & vbCrLf & "改行は vbCrLf"

    ' ***********************************************************
    ' CC BCC HTMLメール( CC BCC はどちらか片方  )
    ' ※ 両方指定すると CC
    ' ***********************************************************
    'Cdo.Cc = "ユーザ名@ドメイン1,ユーザ名@ドメイン2"
    'Cdo.Bcc = "ユーザ名@ドメイン1,ユーザ名@ドメイン2"
    Cdo.Htmlbody = "<img src=""http://winofsql.jp/image/winofsql.png"">"

    ' ***********************************************************
    ' ファイル添付あり
    ' ***********************************************************
    Cdo.AddAttachment ("C:\Users\sworc\Desktop\画像\_img.jpg")

    ' ***********************************************************
    ' 設定
    ' ***********************************************************
    Cdo.Configuration.Fields.Item _
     ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
    Cdo.Configuration.Fields.Item _
     ("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.lolipop.jp"
    Cdo.Configuration.Fields.Item _
     ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 465
    Cdo.Configuration.Fields.Item _
     ("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
    
    Cdo.Configuration.Fields.Item _
     ("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
    Cdo.Configuration.Fields.Item _
     ("http://schemas.microsoft.com/cdo/configuration/sendusername") = "アカウント"
    Cdo.Configuration.Fields.Item _
     ("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "バスワード"

    ' ***********************************************************
    ' 設定の反映
    ' ***********************************************************
    Cdo.Configuration.Fields.Update

    ' ***********************************************************
    ' 送信
    ' ***********************************************************
    On Error Resume Next
    Cdo.Send
    If Err.Number <> 0 Then
        strMessage = Err.Description
    Else
        strMessage = "送信が完了しました"
    End If
    On Error GoTo 0

    MsgBox (strMessage)

```
