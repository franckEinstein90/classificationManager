Option Compare Database
Option Explicit


'Encapsulate information related
'to a section in a part in a prefix
'and manages the process of exporting that
'information to the XML format understood
'by GCDOCS
Private Type tTaxPartSection
    prefix As String
    part As String
    section As String
    
    primary_from As Long
    primary_to As Long
    
    title_E As String
    description_E As String
    
    title_F As String
    description_F As String
    
    primaries As Scripting.dictionary
    count As Long
End Type

Private m_section As tTaxPartSection

Private Function headTail(ByVal fileNumber As String) As Variant
'decomposes a file number into [head][delimiter][body]
'in format /[classCode][-/][body]?/
    Dim delimiterPos As Integer
    Dim V() As Variant
    
    delimiterPos = InStr(fileNumber, "-")
    If delimiterPos = 0 Then ' - not found
        delimiterPos = InStr(fileNumber, "/")
        GoTo composePair
    End If
    If InStr(fileNumber, "/") = 0 Then GoTo composePair ' / not found
    If InStr(fileNumber, "/") < delimiterPos Then delimiterPos = InStr(fileNumber, "/")

composePair:
    If delimiterPos = 0 Then
        Call Err.Raise(10015, , "Invalid Classification Number")
    End If
    ReDim V(1)
    V(0) = Left(fileNumber, delimiterPos - 1)
    V(1) = Right(fileNumber, Len(fileNumber) - delimiterPos + 1)
    headTail = V
End Function

Public Function add(ByVal fileNumber As String) As Long
    Dim FNHeadBody() As Variant
    FNHeadBody = headTail(fileNumber)
    Dim primaryNumber As String: primaryNumber = FNHeadBody(0)
    Debug.Assert (Len(primaryNumber) = 4) 'making sure pNumber is valid
    With m_section.primaries
        If Not .exists(primaryNumber) Then
            Dim newPrimeNumber As objClassificationNumber
            Set newPrimeNumber = New objClassificationNumber
            Call newPrimeNumber.initialize(primaryNumber)
            Call .add(primaryNumber, newPrimeNumber)
        End If
    End With
    With m_section
        add = .primaries(primaryNumber).add(FNHeadBody(1))
        .count = .count + add
    End With
End Function

Public Sub initialize(ByVal prefix As String, ByVal part As String, ByVal section As String)
    Dim rsSection As DAO.Recordset
    
    Set rsSection = CurrentDb.OpenRecordset("SELECT * FROM tblSections WHERE " & _
        "PREFIX ='" & prefix & "' AND " & _
        "PART ='" & part & "' AND " & _
        "SECTION ='" & section & "'")
    
    If rsSection.EOF Then GoTo errInitSectionObj
    With m_section
        .sectionID = section
        .primary_from = rsSection.fields("PRIMARY_FROM")
        .primary_to = rsSection.fields("PRIMARY_TO")
        
        '******English info
        .title_E = rsSection.fields("SECTION_TITLE_E")
        .description_E = rsSection.fields("SECTION_DESC_E")
        
        '******French info
        .title_F = rsSection.fields("SECTION_TITLE_F")
        .description_F = rsSection.fields("SECTION_DESC_F")
        Set .primaries = New Scripting.dictionary
        .count = 0
    End With
    rsSection.Close
    Exit Sub
errInitSectionObj:
End Sub
