Attribute VB_Name = "mdl_OGR_KontaktgruppenErstellen"
Option Explicit


' neue Kontaktgruppe erzeugen
Public Sub NEUE_Kontaktgruppe_erstellen(strGruppennamen As String, colNames As Collection)

    Dim ContactsFolder As Folder
    Set ContactsFolder = Session.GetDefaultFolder(olFolderContacts)
    
    Dim myNamespace As Outlook.NameSpace
    Set myNamespace = Application.GetNamespace("MAPI")
    
    Dim DistList As DistListItem
    Set DistList = Application.CreateItem(olDistributionListItem)
    
    DistList.DLName = strGruppennamen
    
    Dim ad As Variant
    For Each ad In colNames
    
        Dim Rec1 As Recipient
        Set Rec1 = myNamespace.CreateRecipient(ad)
        Rec1.Resolve
        
        Call DistList.AddMember(Rec1)
    
    Next
        
    Call DistList.Move(ContactsFolder)
    
    
End Sub
