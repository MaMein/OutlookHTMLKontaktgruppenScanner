Attribute VB_Name = "mdl_OGR_OrgaReader"
Option Explicit

Public Type Group
    Name As String
    MembersName As Collection
End Type



Public Sub main()

    Dim Checkgroup As Group
    
    Checkgroup = readOrgagram("https://XXXX.XXXX.com/XXX/8802029")
    Call DebugPrintGroup(Checkgroup)
    Call mdl_OGR_KontaktgruppenErstellen.NEUE_Kontaktgruppe_erstellen("ScriptTest_" & Checkgroup.Name, Checkgroup.MembersName)

    Checkgroup = readOrgagram("https://XXXX.XXXX.com/XXX/8900682")
    Call DebugPrintGroup(Checkgroup)
    Call mdl_OGR_KontaktgruppenErstellen.NEUE_Kontaktgruppe_erstellen("ScriptTest_" & Checkgroup.Name, Checkgroup.MembersName)

    Checkgroup = readOrgagram("https://XXXX.XXXX.com/XXX/8800757")
    Call DebugPrintGroup(Checkgroup)
    Call mdl_OGR_KontaktgruppenErstellen.NEUE_Kontaktgruppe_erstellen("ScriptTest_" & Checkgroup.Name, Checkgroup.MembersName)

End Sub


Private Sub DebugPrintGroup(grpOutput As Group)
    Debug.Print "===================="
    Debug.Print "OrgaName:" & vbTab & grpOutput.Name
    Debug.Print ""
    
    Dim varName As Variant
    For Each varName In grpOutput.MembersName
        Debug.Print "MemberName:" & vbTab & varName
    Next
End Sub

 
Public Function readOrgagram(strUrl As String) As Group

    Dim colNames As New Collection
    Dim strOrgaName As String
 
    Dim ReceivedHTML As String
    ReceivedHTML = getWebsiteText(strUrl)
     
    Dim doc As HTMLDocument
    Set doc = New HTMLDocument
    doc.Body.innerHTML = ReceivedHTML
    
    Dim el As IHTMLElement
    For Each el In doc.getElementsByClassName("org-title")
        strOrgaName = el.innerHTML
    Next
    
    Dim strPersonInfo As String
    For Each el In doc.getElementsByClassName("people-card")
        Dim docPerson As New HTMLDocument
        docPerson.Body.innerHTML = el.innerHTML
        
        strPersonInfo = ""
        Dim subEl As IHTMLElement
        For Each subEl In docPerson.getElementsByClassName("person-name")
            Debug.Print subEl.innerHTML
            ' Ignore Errors because of double Names
            On Error Resume Next
            Call colNames.add(subEl.innerText, subEl.innerText)
            On Error GoTo 0
        Next
    Next
    
    Dim retGroup As Group
    retGroup.Name = strOrgaName
    Set retGroup.MembersName = colNames

    readOrgagram = retGroup

End Function


Function getWebsiteText(strUrl As String) As String

    Dim XMLHttp As Object
    Dim strMethod As String, strUser As String
    Dim strPassword As String
    Dim bolAsync As Boolean
    Dim varMessage
    
    ' Microsoft XML HTTP Objekt erzeugen
    Set XMLHttp = CreateObject("MSXML2.XMLHTTP")
    
    ' Parameter für einen simplen POST-Request ohne Authentifizierung füllen
    strMethod = "POST"
    strUrl = strUrl
    bolAsync = False
    strUser = ""
    strPassword = ""
    varMessage = "CountryCode=DE"
    
    ' Request absetzen
    Call XMLHttp.Open(strMethod, strUrl, bolAsync, strUser, strPassword)
    Call XMLHttp.setRequestHeader("Content-Type", "application/x-www-form-urlencoded")
    Call XMLHttp.Send(varMessage)
    
    getWebsiteText = XMLHttp.responseText
    
End Function
