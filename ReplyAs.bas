Attribute VB_Name = "ReplyAs"
Dim oMsgReply As Outlook.MailItem

'
'HTML�`���ŕԐM���[�����쐬����
'
Sub ReplyAsHTML()
    
    MakeHtmlReply
    
End Sub

'
'�v���[���e�L�X�g�`���ŕԐM���[�����쐬����
'
Sub ReplyAsPlainText()

    Dim lineColctn As New Collection
    Dim quotedLineColctn As New Collection
    Dim replyBody As String
    
    MakeHtmlReply
    
    If Not (oMsgReply Is Nothing) Then '�ԐM���[���쐬�����̏ꍇ
        
        
        '1�s���ɕ�������collection�����
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
        
        '���p"> "��t������collection�����
        For Each oneLine In lineColctn
            
            quotedLineColctn.Add Item:="> " & oneLine
            
        Next
        
        '�{�������
        replyBody = String(3, vbCrLf) '�擪�ɋ�s��}������
        For Each oneLine In quotedLineColctn
            
            replyBody = replyBody & vbCrLf & oneLine
            
        Next
        
        oMsgReply.BodyFormat = olFormatPlain '���[���I�u�W�F�N�g���v���[���e�L�X�g�`���ɕύX����
        oMsgReply.Body = replyBody '�{��
        
    End If
    
End Sub

'
'�I����Ԃ̃��[���I�u�W�F�N�g����
'HTML�`���̕ԐM���[�����쐬����
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
        
        oMsg.BodyFormat = olFormatHTML '�ԐM�����[����HTML�`���ɕϊ�����
        
        'Set oMsgReply = oMsg.Reply '���M�҂ɕԐM
        Set oMsgReply = oMsg.ReplyAll '�S���ɕԐM

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
