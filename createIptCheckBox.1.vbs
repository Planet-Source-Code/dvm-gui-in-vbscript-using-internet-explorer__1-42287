option explicit

dim oIE
set oIE = createIEWindow("", "", 300, 100, "createIptCheckBox 1 by vOOs.....", "oIE_")

dim checkbox
set checkbox = createIptCheckbox(oIE, 10, 10)

checkbox.checked = true

set checkbox.onClick = getRef("checkbox_onClick")

do: wscript.sleep 1: loop

sub checkbox_onClick
  msgbox checkbox.checked
end sub

sub oIE_onQuit
  wscript.quit 0
end sub

function createIptCheckbox(byVal Window, byVal pxLeft, byVal pxTop)
  dim localObject: set localObject = Window.document.createElement("<INPUT TYPE=CHECKBOX>")
  localObject.style.position = "absolute"
  localObject.style.pixelLeft = pxLeft
  localObject.style.pixelTop =  pxTop
  Window.document.body.appendChild localObject
  set createIptCheckbox = localObject
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