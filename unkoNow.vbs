option explicit 

' CREATE WSH-SHELL OBJECT
dim objWsh, i, j, k, tweetStr
set objWsh = createObject("wscript.shell")

const BROWSER = "chrome"
const URL = "https://twitter.com/home"

' objWsh.AppActivate "Google Chrome"
' WScript.Sleep 1000

function addStr(nth)
    for k = 0 to nth
        if k > 0 then
            objWsh.sendkeys("ÅI")
            WScript.Sleep 50
        end if
    next
end function

function tweet(nth)
    ' objWsh.sendkeys("{F5}")
    objWsh.run BROWSER & " " & URL
    WScript.Sleep 6000
    for i = 1 to 14
        objWsh.sendkeys("{TAB}")
        WScript.Sleep 500
    next
    objWsh.sendkeys("{ENTER}")
    WScript.Sleep 500
    objWsh.sendkeys("Çƒ")
    WScript.Sleep 500
    objWsh.sendkeys("Ç∑")
    WScript.Sleep 500
    objWsh.sendkeys("Ç∆")
    WScript.Sleep 500
    addStr(nth)
    ' objWsh.sendkeys("{F6}")
    ' WScript.Sleep 500
    objWsh.sendkeys("{ENTER}")
    WScript.Sleep 500
    objWsh.sendkeys("^{ENTER}")
    WScript.Sleep 500
end function

' tweet()

for j = 0 to 1
    tweet(j)
    ' WScript.Sleep 3600000
    WScript.Sleep 2000
next

Set objWsh = Nothing