<div align="center">

## DELETE a user from ANY COMPUTER on the network


</div>

### Description

Simple API call to delete any user account on the network. I think you need to be admin to do so... infact, its very likely you have to be :)

You need to know:

1. Machine name (ie. \\jon) 2. Username (ie. the_cleaner)

Enjoy! :)
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Jon Barker](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/jon-barker.md)
**Level**          |Intermediate
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/jon-barker-delete-a-user-from-any-computer-on-the-network__1-30171/archive/master.zip)





### Source Code

```
'PASTE THE FOLLOWING INTO A FORM
'YOU NEED A COMMAND BUTTON NAMED
'COMMAND1
Option Explicit
Private Declare Function NetUserDel Lib "NETAPI32.DLL" (ByVal servername As String, ByVal userName As String) As Long
Private Sub Command1_Click()
  On Error GoTo error
  Dim r As Long
  Dim sServer As String
  Dim sUser As String
  sServer = StrConv("\\jon", vbUnicode) ' CHANGE THESE TO YOUR SELECTED USER AND SERVER
  sUser = StrConv("the_cleaner", vbUnicode)   ' CHANGE THESE TO YOUR SELECTED USER AND SERVER
  r = NetUserDel(sServer, sUser)
  If r <> 0 Then
    MsgBox "Delete user failed. Ensure: " & vbCrLf & vbCrLf & _
        "o The server name was correct and started with '\\'" & _
        "o You have admin rights for that server (I think :)" & _
        "o The username you specified was valid", vbCritical, "Error: " & r
  Else
    MsgBox "User deleted!", vbExclamation, "Success"
  End If
  Exit Sub
error:
  MsgBox "External error deleteing user: " & vbCrLf & vbCrLf & Err.Description, vbCritical, "Error: " & Err.Number
End Sub
```

