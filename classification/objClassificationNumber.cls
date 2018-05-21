VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "objClassificationNumber"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Type TClassificationNumber
    fileNumber As String
    section As objTaxTreePartSection
    classificationHeadTag As String
    headDelimiter As String
    metadata As TbillingualClassificationMetadata
    subClassifications As Scripting.dictionary
    count As Long
    classifiesItems As Boolean 'true if classification actually classifies something
End Type

Private m_classificationNumber As TClassificationNumber

Public Property Get fileNumber()
    With m_classificationNumber
        fileNumber = .fileNumber
    End With
End Property


Private Function headTail(ByVal fileNumber As String) As Variant
'decomposes a file number into [head][delimiter][body]
'in format /[classCode][-/][body]?/
    Dim delimiterPos As Integer
    Dim V() As Variant
     
    delimiterPos = InStr(2, fileNumber, "-")
    If delimiterPos = 0 Then ' - not found
        delimiterPos = InStr(2, fileNumber, "/")
        GoTo composePair
    End If
    If InStr(2, fileNumber, "/") = 0 Then GoTo composePair ' / not found
    If InStr(2, fileNumber, "/") < delimiterPos Then delimiterPos = InStr(2, fileNumber, "/")
 
composePair:
    If delimiterPos = 0 Then
        ReDim V(0)
        V(0) = fileNumber
    Else
        ReDim V(1)
        V(0) = Left(fileNumber, delimiterPos - 1)
        V(1) = Right(fileNumber, Len(fileNumber) - delimiterPos + 1)
    End If
    headTail = V
End Function


Private Function getClassificationNumberInfo(ByVal classification As String) As TbillingualClassificationMetadata
    Dim rs As DAO.Recordset
    Dim classificationNumberElements() As String
    classificationNumberElements = Split(classification, " ")
    
    Dim sectionInfo() As String
    sectionInfo = Split(m_classificationNumber.section.fileNumber, ":")
    
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM tblFileNumberDescriptions WHERE [PREFIX]='" & sectionInfo(0) & "' AND [FILE_NUMBER] = '" & classification & "'")
    With rs
        If .EOF Then 'Classification doesn't exist
            With getClassificationNumberInfo
                .desc_English = "Intermediary Node"
                .desc_French = "Node intermediaire"
                .title_English = "Intermediary Node"
                .title_French = "Node intermediaire"
            End With
            Exit Function
        End If
        getClassificationNumberInfo.desc_English = Nz(rs.fields("FILE_DESC_E"), "No English Description")
        getClassificationNumberInfo.title_English = Nz(rs.fields("FILE_SUBJECT_E"), "No English Title")
        getClassificationNumberInfo.desc_French = Nz(rs.fields("FILE_DESC_F"), "No English Description")
        getClassificationNumberInfo.title_French = Nz(rs.fields("FILE_SUBJECT_F"), "Pas de titre Francais")
    End With
    
End Function

Public Function add(ByVal subClassification As String)
    Dim classDecomp As Variant
    classDecomp = headTail(subClassification)
    
    Dim delimiter As String: delimiter = Left(classDecomp(0), 1)
    Debug.Assert delimiter = "-" Or delimiter = "/"
    Dim subClassHeadTag As String: subClassHeadTag = Right(classDecomp(0), Len(classDecomp(0)) - 1)
    Dim createdNew As Boolean
    
    With m_classificationNumber.subClassifications
        If Not .exists(subClassHeadTag) Then
            Dim subClassificationNumber As objClassificationNumber
            Set subClassificationNumber = New objClassificationNumber
            Dim metadata As TbillingualClassificationMetadata
            
            metadata = getClassificationNumberInfo(fileNumber & delimiter & subClassHeadTag)

            Call subClassificationNumber.initialize(m_classificationNumber.section, subClassHeadTag, fileNumber & delimiter & subClassHeadTag, metadata, delimiter)
            .add subClassHeadTag, subClassificationNumber
            createdNew = True
        End If
    End With
    With m_classificationNumber
        If UBound(classDecomp) > 0 Then
            add = .subClassifications(subClassHeadTag).add(classDecomp(1))
            .count = .count + add
        Else
            If createdNew Then add = 1
            .classifiesItems = True
        End If
    End With
End Function


Public Sub exportToGCDOCSXML_Format(ofile As TextStream)
    Dim englishName As String
    
    With m_classificationNumber
        ofile.WriteLine "<RMClassification description=" & qu(.metadata.desc_English) & _
                    " managementtype=" & qu("manual") & " name=" & qu(.headDelimiter & .classificationHeadTag & " " & .metadata.title_English) & _
                    " selectable=" & qu("true") & " subtype=" & qu("551") & ">"
        Call multilingualToXML(ofile, .metadata)
        
        ofile.WriteLine "<filenumber>" & .headDelimiter & .classificationHeadTag & "</filenumber>"
        ofile.WriteLine "<filestatus>ACT</filestatus>"
        ofile.WriteLine "<createdate>20171101</createdate>"
        ofile.WriteLine "<filestatusdate>20171101</filestatusdate>"
        ofile.WriteLine "<essential>0</essential>"
        ofile.WriteLine "<storage>PHY</storage>"
        ofile.WriteLine "<cycleperiod>0</cycleperiod>"
        ofile.WriteLine "<closedflag>0</closedflag>"
        Dim subClassification As Variant
        For Each subClassification In .subClassifications
            Call .subClassifications(subClassification).exportToGCDOCSXML_Format(ofile)
        Next subClassification
        ofile.WriteLine "</RMClassification>"
    End With
End Sub


Public Sub initialize(section As objTaxTreePartSection, ByVal pN As String, _
    fileNumber As String, _
    metadata As TbillingualClassificationMetadata, _
    Optional ByVal delimiter As String = "")
    If glbDebugOn Then
        Stop
        Debug.Print "Initializing new subbranch for " & delimiter & pN
    End If
    With m_classificationNumber
        Set .section = section
        .fileNumber = fileNumber
        .classificationHeadTag = pN
        .headDelimiter = delimiter
        .metadata = metadata
        Set .subClassifications = New Scripting.dictionary
        .count = 0
        .classifiesItems = False
    End With
End Sub
