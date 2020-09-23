option explicit

dim oIE
set oIE = createIEWindow("", "", 270, 80, "createIptFile by vOOs.........", "oIE_")

dim iptFile
set iptFile = createIptFile(oIE, "C:\", 10, 10, 240, 25)

do: wscript.sleep 1: loop until iptFile.value <> ""

oIE.visible = false

msgbox iptFile.value

oIE.quit

sub oIE_onQuit
  wscript.quit
end sub

function createIptFile(byVal Window, byVal sValue, byVal pxLeft, byVal pxTop, byVal pxWidth, byVal pxHeight)
  dim localObject: set localObject = Window.document.createElement("<INPUT TYPE=FILE>")
  localObject.style.position = "absolute"
  localObject.style.pixelLeft = pxLeft
  localObject.style.pixelTop = pxTop
  localObject.style.pixelWidth = pxWidth
  localObject.style.pixelHeight = pxHeight
  localObject.style.cursor = "hand"
  Window.document.body.appendChild localObject
  set createIptFile = localObject
end function

function createIEWindow(byVal pxLeft, byVal pxTop, byVal pxWidth, byVal pxHeight, byVal sTitle, byVal eventHandler)
  dim localObject
  set localObject = wscript.createObject("internetExplorer.Application", eventHandler)
  localObject.navigate "javascript: '<title>" & sTitle & "</title><body scroll=""no""></body>';"
  localObject.visible = false
  localObject.resizable = false
  localObject.toolbar = 0
  localObject.statusbar = 0
  do: wscript.sleep 1: loop until localObject.ReadyState = 4
  if isNumeric(pxWidth) then
    localObject.width = pxWidth
  else
    localObject.width = localObject.document.parentWindow.screen.availWidth
  end if
  if isNumeric(pxHeight) then
    localObject.height = pxHeight
  else
    localObject.height = localObject.document.parentWindow.screen.availHeight
  end if
  if isNumeric(pxTop) then 
    localObject.top = pxTop
  else
    localObject.top = (localObject.document.parentWindow.screen.availHeight / 2) - (localObject.height / 2)
  end if
  if isNumeric(pxLeft) then 
    localObject.Left = pxLeft
  else
    localObject.left = (localObject.document.parentWindow.screen.availWidth / 2) - (localObject.width / 2)
  end if
  localObject.document.parentWindow.focus
  localObject.visible = true
  set createIEWindow = localObject
end function