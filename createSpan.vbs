option explicit

dim oIE
set oIE = createIEWindow("", "", 250, 100, "createSpan by vOOs.........", "oIE_")

dim span
set span = createSpan(oIE, "Hello world!", 10, 10, 220, 50)

span.style.textAlign = "center"
span.style.font = "24Pt Comic Sans MS"

function createSpan(byVal Window, byVal sValue, byVal pxLeft, byVal pxTop, byVal pxWidth, byVal pxHeight)  
  dim localObject: set localObject = Window.document.createElement("<SPAN>")
  localObject.innerText = sValue
  localObject.style.position = "absolute"
  localObject.style.pixelLeft = pxLeft
  localObject.style.pixelTop = pxTop
  localObject.style.pixelWidth = pxWidth
  localObject.style.pixelHeight = pxHeight
  Window.document.body.appendChild localObject
  set createSpan = localObject
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