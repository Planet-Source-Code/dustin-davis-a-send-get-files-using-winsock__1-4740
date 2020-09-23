<div align="center">

## a Send / Get files using winsock


</div>

### Description

Want your program to be able to send and get files over the internet? Well, here you go! This is very easy to understand.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Dustin Davis](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/dustin-davis.md)
**Level**          |Advanced
**User Rating**    |4.4 (57 globes from 13 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/dustin-davis-a-send-get-files-using-winsock__1-4740/archive/master.zip)





### Source Code

```
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Note: This is a custom control for your applications. This will not properly
'Get files from the internet or from an ftp. Although the dataarival sub would
'be the same, I do not know how the transition would end so I just sent "xx"
'and that tells the sub that the transition has ended
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Public declarations
dim FileOpen as Boolean
public function Send_File(FileToSend as string)
'This is the function that sends a file
Dim Temp as string
Dim BlockSize as long
open filetosend for binary access read as #1 'Open the file to send
BlockSize = 2048 'Set the block size, if needed, set it higher
do while not EOF(1)
 temp = Space$(blockSize) 'Give temp some space to store the data
 Get 1, , temp 'Get first line from file
 Winsock1.SendData temp 'Send the data
 DoEvents
loop
winsock1.senddata "xx" 'This is a custom control that ends the transmition
close #1
end function
Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim temp As String
Dim data As String
'Check to see if the file is already open
If fileOpen = False Then
 Open "c:\somefile_here" For Binary Access Write As #2
 fileOpen = True
ElseIf fileOpen = True Then
 DoEvents
End If
Winsock1.GetData data 'Get the data
temp = data
'Check to see if it is the end of the transmition
If temp = "xx" Then
 Close #2
 fileOpen = False
Else
  Put 2, , temp 'Store the data to the file
End If
End Sub
```

