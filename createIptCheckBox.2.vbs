option explicit

dim oIE1
set oIE1 = createIEWindow("", "", 300, 100, "createIptCheckBox 2 by vOOs.....", "oIE1_")
dim oIE2
set oIE2 = createIEWindow("", "", 300, 100, "createIptCheckBox 2 by vOOs.....", "oIE2_")
oIE2.top = oIE1.top + oIE1.height
dim checkbox1
set checkbox1 = createIptCheckbox2(oIE1, "Dit is een label", "right", 10, 10, 270, 25)

dim checkbox2
set checkbox2 = createIptCheckbox2(oIE2, "Dit is een label", "left", 10, 10, 270, 25)

set checkbox1.all.checkbox.onClick = getRef("checkbox1_onClick")
set checkbox2.all.checkbox.onClick = getRef("checkbox2_onClick")

do: wscript.sleep 1: loop

sub checkbox1_onClick
  msgbox checkbox1.all.checkbox.checked
end sub

sub checkbox2_onClick
  if checkbox2.all.checkbox.checked then
    checkbox2.all.span.innerText = "I am true ;)"
  else
    checkbox2.all.span.innerText = "I am false :("
  end if
end sub

sub oIE1_onQuit
  oIE2.quit
  wscript.quit 0
end sub

sub oIE2_onQuit
  oIE1.quit
  wscript.quit 0
end sub

function createIptCheckbox2(byVal Window, sValue, sAlign, byVal pxLeft, byVal pxTop, pxWidth, pxHeight)
  dim localObjectDiv: set localObjectDiv = Window.document.createElement("<DIV>")
  dim localObjectCheckbox: set localObjectCheckbox = Window.document.createElement("<INPUT TYPE=CHECKBOX>")
  dim localObjectSpan: set localObjectSpan = Window.document.createElement("<SPAN>")
  localObjectDiv.style.position = "absolute"
  localObjectDiv.style.pixelLeft = pxLeft
  localObjectDiv.style.pixelTop = pxTop
  localObjectDiv.style.pixelWidth = pxWidth
  localObjectDiv.style.pixelHeight = pxHeight
  localObjectCheckbox.ID = "checkbox"
  localObjectCheckBox.style.position = "absolute"
  localObjectSpan.innerText = sValue
  localObjectSpan.ID = "span"
  localObjectSpan.style.position = "absolute"
  if uCase(cStr(sAlign)) = "RIGHT" then
    localObjectCheckbox.style.pixelLeft = pxWidth - (localObjectCheckbox.clientWidth + 20)
    localObjectSpan.style.pixelLeft = 0
    localObjectSpan.style.pixelTop = 0
    localObjectSpan.style.pixelWidth = pxWidth - (localObjectCheckbox.clientWidth + 30)
    localObjectSpan.style.pixelHeight = pxHeight
    localObjectSpan.style.textAlign = "right"
  else
    localObjectCheckbox.style.pixelLeft = 0
    localObjectSpan.style.pixelLeft = localObjectCheckbox.clientWidth + 20
    localObjectSpan.style.pixelTop = 0
    localObjectSpan.style.pixelWidth = pxWidth - (localObjectCheckbox.clientWidth + 20)
    localObjectSpan.style.pixelHeight = pxHeight
  end if
  localObjectCheckbox.style.pixelTop = 0
  localObjectSpan.style.pixelTop = 0
  localObjectDiv.appendChild localObjectCheckbox
  localObjectDiv.appendChild localObjectSpan
  Window.document.body.appendChild localObjectDiv
  set createIptCheckBox2 = localObjectDiv
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