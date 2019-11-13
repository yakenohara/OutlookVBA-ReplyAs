Attribute VB_Name = "ReplyAs"
' <Globals>---------------------------------------------
Dim obj_replyingMailItem As Outlook.MailItem
' --------------------------------------------</Globals>

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
    Dim qut As String: qut = "> "
    
    MakeHtmlReply
    
    If Not (obj_replyingMailItem Is Nothing) Then '�ԐM���[���쐬�����̏ꍇ
        
        
        '1�s���ɕ�������collection�����
        crlfSplitted = Split(obj_replyingMailItem.Body, vbCrLf)
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
            
            quotedLineColctn.Add Item:=qut & oneLine
            
        Next
        
        '�{�������
        replyBody = String(3, vbCrLf) '�擪�ɋ�s��}������
        replyBody = replyBody & qut & "-----Original Message-----"
        
        For Each oneLine In quotedLineColctn
            
            replyBody = replyBody & vbCrLf & oneLine
            
        Next
        
        obj_replyingMailItem.BodyFormat = olFormatPlain '���[���I�u�W�F�N�g���v���[���e�L�X�g�`���ɕύX����
        obj_replyingMailItem.Body = replyBody '�{��
        
    End If
    
End Sub

'
'�I����Ԃ̃��[���I�u�W�F�N�g����
'HTML�`���̕ԐM���[�����쐬����
'
Private Sub MakeHtmlReply()

    '�G���[�_�C�A���O�p�^�C�g��
    Const str_dialogTitle As String = "Reply in HTML"
    
    '�G���[�_�C�A���O�ŁA�I�u�W�F�N�g�N���X�ꗗ�̐���URL��\������ꍇ�ȉ����g��
    Const str_objClassEnumsUrl As String = "https://docs.microsoft.com/ja-jp/office/vba/api/outlook.olobjectclass"

    Dim obj_activeItem As Object '�\�����A�C�e��
    
    'Initialize
    Set obj_replyingMailItem = Nothing

    '<�I����ԃ`�F�b�N>--------------------------------------------------------------------------
    
    '��ʏ�Ԃɉ������A�C�e���擾
    Select Case TypeName(Application.ActiveWindow)
        
        Case "Explorer" 'Outlook ���C�����(���X�g�ꗗ, �{���E�B���h�E���)�̏ꍇ
        
            Set oSelection = Application.ActiveExplorer.Selection
            
            If (oSelection.Count = 0) Then ' �A�C�e�����I������Ă��Ȃ��ꍇ
                
                MsgBox "Please select an item first!", _
                       vbCritical, _
                       str_dialogTitle
                
                Exit Sub ' �I��
            
            ElseIf (1 < oSelection.Count) Then ' �A�C�e���������I������Ă���ꍇ
                
                MsgBox "Only one item can be replyed", _
                       vbCritical, _
                       str_dialogTitle
                
                Exit Sub ' �I��
                
            Else '�A�C�e���I���� 1 �݂̂̏ꍇ
                Set obj_activeItem = oSelection.Item(1) '�I���A�C�e�����L�^
            
            End If
            
        Case "Inspector" '�P�̕\��(�A�C�e�����_�u���N���b�N���ĊJ�������)�̏ꍇ
        
            Set obj_activeItem = Application.ActiveInspector.CurrentItem '�I���A�C�e�����L�^
        
        Case Else 'Unkown �ȏꍇ
        
            MsgBox "Unsupported Window type.", _
                   vbCritical, _
                   str_dialogTitle
                   
            Exit Sub ' �I��
            
    End Select
    
    '-------------------------------------------------------------------------</�I����ԃ`�F�b�N>
        
    '�ԐM���b�Z�[�W�̍쐬
    If (obj_activeItem.Class = olMail) Then ' MailItem �̏ꍇ
        
        obj_activeItem.BodyFormat = olFormatHTML '�ԐM�����[����HTML�`���ɕϊ�����
        Set obj_replyingMailItem = obj_activeItem.ReplyAll '�S���ԐM�̃��[���A�C�e�����쐬
        obj_activeItem.Close (olDiscard) '�ԐM�����[���ɑ΂���HTML�`���ϊ������j��
        
        obj_replyingMailItem.Display '�ԐM���[����\��
        
    ElseIf (obj_activeItem.Class = olAppointment) Then '�\��A�C�e���̏ꍇ
    
        '��c�o�Ȉ˗��`���̗\��ɕϊ����ꂽ copy item ���쐬����
        Set obj_tmpItem = obj_activeItem.Copy ' copy item ���쐬
        obj_tmpItem.MeetingStatus = olMeeting '��c�`���ɕϊ�
        obj_tmpItem.Save
        obj_tmpItem.Display '�ۑ����� copy item ��P�̕\��(�A�C�e�����_�u���N���b�N���ĊJ�������)�ɂ���
        
        Application.ActiveInspector.CommandBars.ExecuteMso ("ReplyAll") '�E�B���h�E�@�\�� `�S���ɕԐM` �����s
        Set obj_replyingMailItem = Application.ActiveInspector.CurrentItem '�J���ꂽ�V�K�E�B���h�E��ԐM MailItem �Ƃ���
        
        obj_tmpItem.Close (olDiscard) 'copy item �ɑ΂����c�`���ϊ������j��
        obj_tmpItem.Delete 'copy item ���폜
        
    ElseIf _
        (obj_activeItem.Class = olMeetingRequest) Or _
        (obj_activeItem.Class = olMeetingResponseNegative) Or _
        (obj_activeItem.Class = olMeetingResponsePositive) Or _
        (obj_activeItem.Class = olMeetingResponseTentative) _
    Then '��c�o�Ȉ˗��E�ԐM�̏ꍇ
        
        If (TypeName(Application.ActiveWindow) = "Explorer") Then 'Outlook ���C�����(���X�g�ꗗ, �{���E�B���h�E���)�̏ꍇ
        
            Application.ActiveExplorer.CommandBars.ExecuteMso ("ReplyAll") '�E�B���h�E�@�\�� `�S���ɕԐM` �����s
            
        Else ' `Inspector` ��� (= �P�̕\��(�A�C�e�����_�u���N���b�N���ĊJ�������))�̏ꍇ
        
            Application.ActiveInspector.CommandBars.ExecuteMso ("ReplyAll") '�E�B���h�E�@�\�� `�S���ɕԐM` �����s
        
        End If
        
        Set obj_replyingMailItem = Application.ActiveInspector.CurrentItem '�J���ꂽ�V�K�E�B���h�E��ԐM MailItem �Ƃ���
        
    Else 'Unkown �A�C�e���̏ꍇ
    
        int_retOfDialog = MsgBox( _
            "Unsupported item type `" & obj_activeItem.Class & "`." & vbCrLf & _
            "To check meaning of this enumeration value, please visit following URL." & vbCrLf & _
            "(select `OK` to copy URL)" & vbCrLf & _
            "" & vbCrLf & _
            str_objClassEnumsUrl, _
            vbCritical + vbOKCancel, _
            str_dialogTitle _
        )
        
        If (int_retOfDialog = vbOK) Then ' `OK` ���I�����ꂽ�ꍇ
            SetCB str_objClassEnumsUrl 'URL ���N���b�v�{�[�h�ɃR�s�[
        
        End If
               
        Exit Sub '�I��
        
    End If


End Sub


'<�N���b�v�{�[�h����>-------------------------------------------

'�N���b�v�{�[�h�ɕ�������i�[
Private Sub SetCB(ByVal str As String)
  With CreateObject("Forms.TextBox.1")
    .MultiLine = True
    .Text = str
    .SelStart = 0
    .SelLength = .TextLength
    .Copy
  End With
End Sub

'�N���b�v�{�[�h���當������擾
Private Sub GetCB(ByRef str As String)
  With CreateObject("Forms.TextBox.1")
    .MultiLine = True
    If .CanPaste = True Then .Paste
    str = .Text
  End With
End Sub

'------------------------------------------</�N���b�v�{�[�h����>

