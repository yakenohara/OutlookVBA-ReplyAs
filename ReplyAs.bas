Attribute VB_Name = "ReplyAs"
Dim oMsgReply As Outlook.MailItem

'
'HTML形式で返信メールを作成する
'
Sub ReplyAsHTML()
    
    MakeHtmlReply
    
End Sub

'
'プレーンテキスト形式で返信メールを作成する
'
Sub ReplyAsPlainText()

    Dim lineColctn As New Collection
    Dim quotedLineColctn As New Collection
    Dim replyBody As String
    
    MakeHtmlReply
    
    If Not (oMsgReply Is Nothing) Then '返信メール作成成功の場合
        
        
        '1行毎に分解したcollectionを作る
        crlfSplitted = Split(oMsgReply.Body, vbCrLf)
        For Each crlfOneLine In crlfSplitted
        
            crSplitted = Split(crlfOneLine, vbCr)
            For Each crOneLine In crSplitted
                
                lfSplitted = Split(crOneLine, vbLf)
                For Each lfOneLine In lfSplitted
                    
                    lineColctn.Add Item:=lfOneLine
                    
                Next lfOneLine
            
            Next crOneLine
            
        Next crlfOneLine
        
        '引用"> "を付加したcollectionを作る
        For Each oneLine In lineColctn
            
            quotedLineColctn.Add Item:="> " & oneLine
            
        Next
        
        '本文を作る
        replyBody = String(3, vbCrLf) '先頭に空行を挿入する
        For Each oneLine In quotedLineColctn
            
            replyBody = replyBody & vbCrLf & oneLine
            
        Next
        
        oMsgReply.BodyFormat = olFormatPlain 'メールオブジェクトをプレーンテキスト形式に変更する
        oMsgReply.Body = replyBody '本文
        
    End If
    
End Sub

'
'選択状態のメールオブジェクトから
'HTML形式の返信メールを作成する
'
Private Sub MakeHtmlReply()

    Dim oSelection As Outlook.Selection
    Dim oItem As Object
    
    Set oMsgReply = Nothing

    'Get the selected item
    Select Case TypeName(Application.ActiveWindow)
        Case "Explorer"
            Set oSelection = Application.ActiveExplorer.Selection
            
            If oSelection.Count > 0 Then
                Set oItem = oSelection.Item(1)
            Else
                MsgBox "Please select an item first!", vbCritical, "Reply in HTML"
                Exit Sub
            End If
            
        Case "Inspector"
            Set oItem = Application.ActiveInspector.CurrentItem
        
        Case Else
            MsgBox "Unsupported Window type." & vbNewLine & "Please select or open an item first.", _
                   vbCritical, _
                   "Reply in HTML"
                   
            Exit Sub
            
    End Select
        
    Dim oMsg As Outlook.MailItem
    
    'Change the message format and reply
    If oItem.Class = olMail Then
        
        Set oMsg = oItem
        
        oMsg.BodyFormat = olFormatHTML '返信元メールをHTML形式に変換する
        
        'Set oMsgReply = oMsg.Reply '送信者に返信
        Set oMsgReply = oMsg.ReplyAll '全員に返信

        oMsg.Close (olDiscard)
        oMsgReply.Display
        
    Else 'Selected item isn't a mail item
        MsgBox "No message item selected. Please select a message first.", _
               vbCritical, _
               "Reply in HTML"
               
        Exit Sub
        
    End If

    'Cleanup
    Set oMsg = Nothing
    Set oItem = Nothing
    Set oSelection = Nothing

End Sub
