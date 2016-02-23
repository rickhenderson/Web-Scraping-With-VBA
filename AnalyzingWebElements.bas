Option Explicit

Sub DeterminingWebElementIDNames()
' Uses the simple parsing from the WiseOwl web scraping tutorial.

'to refer to the running copy of Internet Explorer
Dim ie As InternetExplorer

'to refer to the HTML document returned
Dim html As HTMLDocument

'open Internet Explorer in memory, and go to website
Set ie = New InternetExplorer
ie.Visible = False
ie.Navigate Range("webURL").Value

'Wait until IE is done loading page
Do While ie.READYSTATE <> READYSTATE_COMPLETE
    Application.StatusBar = "Trying to go to " & Range("webURL").Value
    DoEvents
Loop

'Get text of HTML document returned
Set html = ie.Document

'close down IE and reset status bar
Set ie = Nothing
Application.StatusBar = ""

' RDH20160222: Disable application updates for slightly faster performance.
Application.ScreenUpdating = False

' Display some results

' innerHTML is all the tags between the <html> and </html> tags
MsgBox html.DocumentElement.innerHTML, vbOKOnly, "The .innerHTML property"

' List all the links
MsgBox ListAllLinks(html), vbOKOnly, "Links from the Links IE DOM object"

' Custom function to create a list of all tags of a specified name to display their ID and className
'Debug.Print ListAllElementsByTagName(html, "input")

' List all the Anchor tags
'MsgBox ListAllAnchorTags(html), vbOKOnly, "List of Anchor Tags"

Range("A12").Value = ListAllLinks(html)

Set html = Nothing

' RDH20160222: Re-enable screen updating to set things to normal.
Application.ScreenUpdating = True

End Sub

Private Function ListAllElementsByTagName(webHTML As HTMLDocument, strTagName As String) As String
' Returns a (potentially large) string which contains all the tags of a specific name from the current page
' Could also change this to a collection. Except it's already a collection

Dim pageElements As IHTMLElementCollection
Dim pageElement As IHTMLElement
Dim strElements As String

' Get a reference to the HTML elements we want
Set pageElements = webHTML.getElementsByTagName(strTagName)

For Each pageElement In pageElements
    strElements = pageElement.ID & ", " & pageElement.className & vbCrLf
Next

' Return the string
ListAllElementsByTagName = strElements

End Function

Private Function ListAllAnchorTags(webHTML As HTMLDocument) As String
' Returns a (potentially large) string which contains all the tags of a specific name from the current page

Dim pageAnchors As IHTMLElementCollection
Dim pageAnchor As IHTMLElement
Dim strLinks As String

' Get a reference to the HTML elements we want
Set pageAnchors = webHTML.getElementsByTagName("a")

For Each pageAnchor In pageAnchors
    strLinks = strLinks & pageAnchor.href & vbCrLf
Next

' Return the string
ListAllAnchorTags = strLinks

End Function

Private Function ListAllLinks(webHTML As HTMLDocument) As String
' The IE DOM has a Links object that contains all links in the HTMLDocument
' Avoids the use of getElementsByTagName

Dim pageLink As Object
Dim strLinks As String

For Each pageLink In webHTML.Links
    strLinks = strLinks & pageLink.href & vbCrLf
Next

' Return the string
ListAllLinks = strLinks

End Function
