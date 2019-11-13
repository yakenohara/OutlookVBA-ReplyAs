Attribute VB_Name = "ReplyAs"
' <Globals>---------------------------------------------
Dim obj_replyingMailItem As Outlook.MailItem
' --------------------------------------------</Globals>

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
    Dim qut As String: qut = "> "
    
    MakeHtmlReply
    
    If Not (obj_replyingMailItem Is Nothing) Then '返信メール作成成功の場合
        
        
        '1行毎に分解したcollectionを作る
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
        
        '引用"> "を付加したcollectionを作る
        For Each oneLine In lineColctn
            
            quotedLineColctn.Add Item:=qut & oneLine
            
        Next
        
        '本文を作る
        replyBody = String(3, vbCrLf) '先頭に空行を挿入する
        replyBody = replyBody & qut & "-----Original Message-----"
        
        For Each oneLine In quotedLineColctn
            
            replyBody = replyBody & vbCrLf & oneLine
            
        Next
        
        obj_replyingMailItem.BodyFormat = olFormatPlain 'メールオブジェクトをプレーンテキスト形式に変更する
        obj_replyingMailItem.Body = replyBody '本文
        
    End If
    
End Sub

'
'選択状態のメールオブジェクトから
'HTML形式の返信メールを作成する
'
Private Sub MakeHtmlReply()

    'エラーダイアログ用タイトル
    Const str_dialogTitle As String = "Reply in HTML"
    
    'エラーダイアログで、オブジェクトクラス一覧の説明URLを表示する場合以下を使う
    Const str_objClassEnumsUrl As String = "https://docs.microsoft.com/ja-jp/office/vba/api/outlook.olobjectclass"

    Dim obj_activeItem As Object '表示中アイテム
    
    'Initialize
    Set obj_replyingMailItem = Nothing

    '<選択状態チェック>--------------------------------------------------------------------------
    
    '画面状態に応じたアイテム取得
    Select Case TypeName(Application.ActiveWindow)
        
        Case "Explorer" 'Outlook メイン画面(リスト一覧, 閲覧ウィンドウ画面)の場合
        
            Set oSelection = Application.ActiveExplorer.Selection
            
            If (oSelection.Count = 0) Then ' アイテムが選択されていない場合
                
                MsgBox "Please select an item first!", _
                       vbCritical, _
                       str_dialogTitle
                
                Exit Sub ' 終了
            
            ElseIf (1 < oSelection.Count) Then ' アイテムが複数選択されている場合
                
                MsgBox "Only one item can be replyed", _
                       vbCritical, _
                       str_dialogTitle
                
                Exit Sub ' 終了
                
            Else 'アイテム選択は 1 のみの場合
                Set obj_activeItem = oSelection.Item(1) '選択アイテムを記録
            
            End If
            
        Case "Inspector" '単体表示(アイテムをダブルクリックして開いた画面)の場合
        
            Set obj_activeItem = Application.ActiveInspector.CurrentItem '選択アイテムを記録
        
        Case Else 'Unkown な場合
        
            MsgBox "Unsupported Window type.", _
                   vbCritical, _
                   str_dialogTitle
                   
            Exit Sub ' 終了
            
    End Select
    
    '-------------------------------------------------------------------------</選択状態チェック>
        
    '返信メッセージの作成
    If (obj_activeItem.Class = olMail) Then ' MailItem の場合
        
        obj_activeItem.BodyFormat = olFormatHTML '返信元メールをHTML形式に変換する
        Set obj_replyingMailItem = obj_activeItem.ReplyAll '全員返信のメールアイテムを作成
        obj_activeItem.Close (olDiscard) '返信元メールに対するHTML形式変換操作を破棄
        
        obj_replyingMailItem.Display '返信メールを表示
        
    ElseIf (obj_activeItem.Class = olAppointment) Then '予定アイテムの場合
    
        '会議出席依頼形式の予定に変換された copy item を作成する
        Set obj_tmpItem = obj_activeItem.Copy ' copy item を作成
        obj_tmpItem.MeetingStatus = olMeeting '会議形式に変換
        obj_tmpItem.Save
        obj_tmpItem.Display '保存した copy item を単体表示(アイテムをダブルクリックして開いた状態)にする
        
        Application.ActiveInspector.CommandBars.ExecuteMso ("ReplyAll") 'ウィンドウ機能の `全員に返信` を実行
        Set obj_replyingMailItem = Application.ActiveInspector.CurrentItem '開かれた新規ウィンドウを返信 MailItem とする
        
        obj_tmpItem.Close (olDiscard) 'copy item に対する会議形式変換操作を破棄
        obj_tmpItem.Delete 'copy item を削除
        
    ElseIf _
        (obj_activeItem.Class = olMeetingRequest) Or _
        (obj_activeItem.Class = olMeetingResponseNegative) Or _
        (obj_activeItem.Class = olMeetingResponsePositive) Or _
        (obj_activeItem.Class = olMeetingResponseTentative) _
    Then '会議出席依頼・返信の場合
        
        If (TypeName(Application.ActiveWindow) = "Explorer") Then 'Outlook メイン画面(リスト一覧, 閲覧ウィンドウ画面)の場合
        
            Application.ActiveExplorer.CommandBars.ExecuteMso ("ReplyAll") 'ウィンドウ機能の `全員に返信` を実行
            
        Else ' `Inspector` 状態 (= 単体表示(アイテムをダブルクリックして開いた画面))の場合
        
            Application.ActiveInspector.CommandBars.ExecuteMso ("ReplyAll") 'ウィンドウ機能の `全員に返信` を実行
        
        End If
        
        Set obj_replyingMailItem = Application.ActiveInspector.CurrentItem '開かれた新規ウィンドウを返信 MailItem とする
        
    Else 'Unkown アイテムの場合
    
        int_retOfDialog = MsgBox( _
            "Unsupported item type `" & obj_activeItem.Class & "`." & vbCrLf & _
            "To check meaning of this enumeration value, please visit following URL." & vbCrLf & _
            "(select `OK` to copy URL)" & vbCrLf & _
            "" & vbCrLf & _
            str_objClassEnumsUrl, _
            vbCritical + vbOKCancel, _
            str_dialogTitle _
        )
        
        If (int_retOfDialog = vbOK) Then ' `OK` が選択された場合
            SetCB str_objClassEnumsUrl 'URL をクリップボードにコピー
        
        End If
               
        Exit Sub '終了
        
    End If


End Sub


'<クリップボード操作>-------------------------------------------

'クリップボードに文字列を格納
Private Sub SetCB(ByVal str As String)
  With CreateObject("Forms.TextBox.1")
    .MultiLine = True
    .Text = str
    .SelStart = 0
    .SelLength = .TextLength
    .Copy
  End With
End Sub

'クリップボードから文字列を取得
Private Sub GetCB(ByRef str As String)
  With CreateObject("Forms.TextBox.1")
    .MultiLine = True
    If .CanPaste = True Then .Paste
    str = .Text
  End With
End Sub

'------------------------------------------</クリップボード操作>

