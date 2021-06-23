'Create an object for open/close the drive.
Set oWMP = CreateObject("WMPlayer.OCX.7" )

'Open the closed drive.
'If you have more drives, change the Item(0) to another value like 1,2,3
'Example oWMP.cdromCollection.Item(1).Eject will open 2nd drive.
oWMP.cdromCollection.Item(0).Eject

'Asking if you want to close the opened drive.
'At laptops its useless, so i put it as a comment.
'If you want to use it, just delete the <'> char from the below lines.
'answer = msgbox ("Close CD Tray?", vbOKCancel, "Erotisi xristi")
'if answer = vbOK then oWMP.cdromCollection.Item(0).Eject

'Just clear memory
set owmp = nothing

'Close script
wscript.quit