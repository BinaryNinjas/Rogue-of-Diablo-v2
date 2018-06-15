Attribute VB_Name = "Chat"
Public Sub AddChat(MSG As String, Color As String)

With frmRoD.rtb
.SelStart = Len(.Text)
'.SelAlignment = vbleft
.SelColor = Color
.SelText = " " & MSG & vbCrLf
.SelStart = Len(.Text)
End With

End Sub
