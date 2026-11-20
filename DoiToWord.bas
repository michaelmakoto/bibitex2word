Attribute VB_Name = "DoiToWord"
' DoiToWord - add Word bibliography sources from DOIs (Word for Mac 2016+ and Windows).
'
' Run the macro AddReferencesFromDOI. It asks for a single identifier or a .txt/.csv file,
' looks up each one at doi.org, and adds it to the active document's source list
' (References > Manage Sources). Accepted identifiers:
'   10.1021/acs.jpcc.8b06244, doi:10.1021/..., https://doi.org/10.1021/...,
'   publisher links that contain a DOI, URL-encoded DOIs,
'   arXiv:2301.01234v2, 2301.01234, https://arxiv.org/abs/hep-th/9711200,
'   PMID: 12345678, https://pubmed.ncbi.nlm.nih.gov/12345678/, PMC1234567.
' Duplicates (in the input, or a DOI already in the document) are skipped.
'
' Metadata: doi.org content negotiation (CSL-JSON). PMID/PMCID -> DOI: NCBI E-utilities.
' Windows uses WinHttp; Mac uses curl through the system shell.
'
' Optional progress bar: Insert > UserForm, set its (Name) to frmProgress, leave it empty,
' and paste this into the form's code window. Without it, progress shows in the status bar.
'
'   Private Sub UserForm_Activate()
'       DoiToWord.RunWithProgress Me
'   End Sub
'
'   Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
'       If CloseMode = 0 Then Cancel = True: Me.Tag = "stop"   ' closing the window = stop after the current item
'   End Sub

Option Explicit

Private Const BIB_NS As String = "http://schemas.openxmlformats.org/officeDocument/2006/bibliography"
Private Const CSL_JSON As String = "application/vnd.citationstyles.csl+json"
Private Const USER_AGENT As String = "doi2word-macro/1.0"
Private Const APP_TITLE As String = "DOI -> Word"
Private Const LCID_EN_US As String = "1033"   ' Language of every added source: English (US)

#If Mac Then
Private Declare PtrSafe Function popen Lib "/usr/lib/libc.dylib" (ByVal command As String, ByVal mode As String) As LongPtr
Private Declare PtrSafe Function pclose Lib "/usr/lib/libc.dylib" (ByVal file As LongPtr) As Long
Private Declare PtrSafe Function fread Lib "/usr/lib/libc.dylib" (ByVal outStr As String, ByVal size As LongPtr, ByVal items As LongPtr, ByVal stream As LongPtr) As Long
Private Declare PtrSafe Function feof Lib "/usr/lib/libc.dylib" (ByVal file As LongPtr) As LongPtr
#End If

Private Const BAR_WIDTH As Single = 330

' Import run state (shared with the progress window)
Private mIds As Collection, mExisting As Collection, mTags As Collection, mFailed As Collection
Private mAdded As Long, mDupes As Long, mDone As Long, mRan As Boolean
Private mForm As Object, mBar As Object, mText As Object

' JSON parser state
Private jS As String
Private jI As Long
Private jN As Long

' ============================================================ entry point

Public Sub AddReferencesFromDOI()
    Dim choice As VbMsgBoxResult, inputText As String, path As String

    If Documents.Count = 0 Then
        MsgBox "Open a document first.", vbExclamation, APP_TITLE
        Exit Sub
    End If

    choice = MsgBox("Add references to this document." & vbCr & vbCr & _
                    "Yes  =  enter a single DOI / arXiv ID / PMID" & vbCr & _
                    "No  =  read a list from a .txt or .csv file" & vbCr & _
                    "Cancel  =  stop", vbYesNoCancel + vbQuestion, APP_TITLE)
    Select Case choice
        Case vbYes
            inputText = InputBox("DOI, DOI link, arXiv ID, PMID or PMCID:", APP_TITLE)
        Case vbNo
            path = PickFile()
            If Len(path) = 0 Then Exit Sub
            inputText = ReadTextFile(path)
            If Len(inputText) = 0 Then
                MsgBox "Could not read the file, or it is empty:" & vbCr & path, vbExclamation, APP_TITLE
                Exit Sub
            End If
        Case Else
            Exit Sub
    End Select

    If Len(Trim$(inputText)) > 0 Then ImportIds inputText
End Sub

' ============================================================ main pipeline

Private Sub ImportIds(ByVal inputText As String)
    Dim unrecognised As New Collection, frm As Object, notRun As Long

    Set mIds = New Collection: Set mExisting = New Collection
    Set mTags = New Collection: Set mFailed = New Collection
    mAdded = 0: mDupes = 0: mDone = 0: mRan = False

    ExtractIds inputText, mIds, unrecognised, mDupes
    If mIds.Count = 0 Then
        MsgBox "No DOI, arXiv ID or PMID found in the input.", vbExclamation, APP_TITLE
        Exit Sub
    End If
    LoadExistingSources mExisting, mTags

    ' The progress window (frmProgress) is optional: its Activate event calls RunWithProgress.
    On Error Resume Next
    Set frm = VBA.UserForms.Add("frmProgress")
    On Error GoTo 0
    If Not frm Is Nothing Then frm.Show
    If Not mRan Then RunWithProgress Nothing          ' no progress window installed

    notRun = mIds.Count - mDone
    MsgBox "Added: " & mAdded & vbCr & _
           "Skipped duplicates: " & mDupes & vbCr & _
           "Failed: " & mFailed.Count & ListText(mFailed, 15) & vbCr & _
           IIf(notRun > 0, "Stopped early, not processed: " & notRun & vbCr, "") & _
           "Lines without an identifier: " & unrecognised.Count & ListText(unrecognised, 10) & vbCr & vbCr & _
           "Cite them with References > Insert Citation. Every source in the document's list " & _
           "appears in the bibliography, cited or not.", vbInformation, APP_TITLE
End Sub

' Looks up and adds every collected identifier. Called by frmProgress (frm) or directly (Nothing).
Public Sub RunWithProgress(ByVal frm As Object)
    Dim item As Variant, kind As String, value As String, doi As String, xml As String, p As Long

    If mRan Then Exit Sub
    mRan = True
    SetupProgress frm

    For Each item In mIds
        If StopRequested() Then Exit For
        p = InStr(item, "|")
        kind = Left$(item, p - 1)
        value = Mid$(item, p + 1)
        UpdateProgress "Looking up " & value

        Select Case kind
            Case "pmid": doi = DoiFromPubMed("pubmed", value)
            Case "pmcid": doi = DoiFromPubMed("pmc", value)
            Case Else: doi = value
        End Select

        If Len(doi) = 0 Then
            mFailed.Add value & "  (no DOI found)"
        ElseIf HasKey(mExisting, doi) Then
            mDupes = mDupes + 1
        Else
            xml = BuildSourceXml(HttpGet("https://doi.org/" & UrlEncodePath(doi), CSL_JSON), doi, mTags)
            If Len(xml) = 0 Then
                mFailed.Add doi & "  (not found or no connection)"
            Else
                On Error Resume Next
                ActiveDocument.Bibliography.Sources.Add xml
                If Err.Number <> 0 Then
                    mFailed.Add doi & "  (Word rejected it: " & Err.Description & ")"
                    Err.Clear
                Else
                    mAdded = mAdded + 1
                    AddKey mExisting, doi
                End If
                On Error GoTo 0
            End If
        End If
        mDone = mDone + 1
    Next

    UpdateProgress "Done"
    ShowStatus ""
    If Not frm Is Nothing Then Unload frm
    Set mForm = Nothing: Set mBar = Nothing: Set mText = Nothing
End Sub

' ============================================================ progress display

Private Sub SetupProgress(ByVal frm As Object)
    Set mForm = frm: Set mBar = Nothing: Set mText = Nothing
    If frm Is Nothing Then Exit Sub
    On Error Resume Next               ' if controls cannot be added, the window title still shows the count
    frm.Width = BAR_WIDTH + 30
    frm.Height = 100
    Set mText = frm.Controls.Add("Forms.Label.1", "lblText")
    mText.Left = 12: mText.Top = 12: mText.Width = BAR_WIDTH: mText.Height = 18
    With frm.Controls.Add("Forms.Label.1", "lblTrack")
        .Left = 12: .Top = 38: .Width = BAR_WIDTH: .Height = 14
        .BackColor = RGB(220, 220, 220)
    End With
    Set mBar = frm.Controls.Add("Forms.Label.1", "lblBar")
    mBar.Left = 12: mBar.Top = 38: mBar.Width = 0: mBar.Height = 14
    mBar.BackColor = RGB(0, 120, 212)
End Sub

Private Sub UpdateProgress(ByVal msg As String)
    Dim counter As String
    counter = mDone & " / " & mIds.Count
    On Error Resume Next
    If mForm Is Nothing Then
        ShowStatus APP_TITLE & "  " & counter & "  " & msg
        Exit Sub
    End If
    mForm.Caption = APP_TITLE & "  (" & counter & ")"
    If Not mText Is Nothing Then mText.Caption = Left$(msg, 70)
    If Not mBar Is Nothing Then mBar.Width = BAR_WIDTH * mDone / mIds.Count
    mForm.Repaint
    DoEvents
End Sub

' Closing the progress window sets its Tag to "stop" (see frmProgress code in the header).
Private Function StopRequested() As Boolean
    If mForm Is Nothing Then Exit Function
    StopRequested = (mForm.Tag = "stop")
End Function

Private Sub LoadExistingSources(existing As Collection, tags As Collection)
    Dim srcs As Object, k As Long, doi As String
    Set srcs = ActiveDocument.Bibliography.Sources
    For k = 1 To srcs.Count
        AddKey tags, srcs.item(k).Tag
        doi = XmlField(srcs.item(k).xml, "DOI")
        If Len(doi) > 0 Then AddKey existing, doi
    Next
End Sub

' ============================================================ identifier parsing

' Collects "kind|value" items (kind = doi, pmid, pmcid). arXiv IDs become DOIs.
Private Sub ExtractIds(ByVal inputText As String, ids As Collection, unrecognised As Collection, ByRef dupes As Long)
    Dim sep As Variant, ln As Variant, toks() As String, t As Long, id As String, lowTok As String, found As Boolean, s As String

    inputText = Replace(Replace(inputText, vbCrLf, vbLf), vbCr, vbLf)
    For Each sep In Array(vbTab, ",", ";", """", ChrW(160))
        inputText = Replace(inputText, sep, " ")
    Next

    For Each ln In Split(inputText, vbLf)
        s = Trim$(ln)
        Do While InStr(s, "  ") > 0
            s = Replace(s, "  ", " ")
        Loop
        If Len(s) > 0 Then
            found = False
            toks = Split(s, " ")
            For t = 0 To UBound(toks)
                id = ""
                lowTok = LCase$(toks(t))
                If (lowTok = "pmid" Or lowTok = "pmid:") And t < UBound(toks) Then
                    If IsDigits(toks(t + 1)) Then
                        id = "pmid|" & toks(t + 1)
                        t = t + 1
                    End If
                Else
                    id = IdFromToken(toks(t))
                End If
                If Len(id) > 0 Then
                    found = True
                    If HasKey(ids, id) Then
                        dupes = dupes + 1
                    Else
                        ids.Add id, id
                    End If
                End If
            Next
            If Not found Then unrecognised.Add Left$(s, 60)
        End If
    Next
End Sub

Private Function IdFromToken(ByVal tok As String) As String
    Dim lowTok As String, s As String
    If InStr(tok, "%") > 0 Then tok = UrlDecode(tok)
    lowTok = LCase$(tok)

    s = DoiFromToken(tok)
    If Len(s) > 0 Then IdFromToken = "doi|" & s: Exit Function
    s = PmidFromToken(lowTok)
    If Len(s) > 0 Then IdFromToken = "pmid|" & s: Exit Function
    s = PmcFromToken(lowTok)
    If Len(s) > 0 Then IdFromToken = "pmcid|" & s: Exit Function
    s = ArxivFromToken(lowTok)
    If Len(s) > 0 Then IdFromToken = "doi|10.48550/arXiv." & s
End Function

' Finds "10.<4+ digits>/<suffix>" and cleans URL and punctuation leftovers.
Private Function DoiFromToken(ByVal tok As String) As String
    Dim i As Long, j As Long, n As Long, d As String, prev As String, isUrl As Boolean
    Dim cut As Variant, sfx As Variant, changed As Boolean, ch As String, opener As String

    n = Len(tok)
    For i = 1 To n - 3
        If Mid$(tok, i, 3) = "10." Then
            If i > 1 Then prev = Mid$(tok, i - 1, 1) Else prev = " "
            If Not prev Like "[A-Za-z0-9]" Then
                j = i + 3
                Do While Mid$(tok, j, 1) Like "[0-9.]"
                    j = j + 1
                Loop
                If j - (i + 3) >= 4 And j < n And Mid$(tok, j, 1) = "/" Then
                    d = Mid$(tok, i)
                    isUrl = InStr(Left$(tok, i - 1), "/") > 0
                    Exit For
                End If
            End If
        End If
    Next
    If Len(d) = 0 Then Exit Function

    If isUrl Then
        For Each cut In Array("?", "#", "&")
            j = InStr(d, cut)
            If j > 0 Then d = Left$(d, j - 1)
        Next
        Do
            changed = False
            For Each sfx In Array("/full", "/abstract", "/pdf", "/epdf", "/html", "/fulltext", ".full", ".pdf", "/")
                If LCase$(Right$(d, Len(sfx))) = sfx Then
                    d = Left$(d, Len(d) - Len(sfx))
                    changed = True
                End If
            Next
        Loop While changed
    End If

    ' bioRxiv / medRxiv links carry a version (...979260v1) that is not part of the DOI
    If LCase$(Left$(d, 8)) = "10.1101/" Then
        j = InStrRev(LCase$(d), "v")
        If j > 8 Then
            If IsDigits(Mid$(d, j + 1)) Then d = Left$(d, j - 1)
        End If
    End If

    ' trailing punctuation; closing brackets only when unbalanced
    Do While Len(d) > 0
        ch = Right$(d, 1)
        Select Case ch
            Case ".", ",", ":", "'", """"
                d = Left$(d, Len(d) - 1)
            Case ")", "]", ">", "}"
                opener = Mid$("([<{", InStr(")]>}", ch), 1)
                If CountOf(d, opener) < CountOf(d, ch) Then d = Left$(d, Len(d) - 1) Else Exit Do
            Case Else
                Exit Do
        End Select
    Loop

    If InStr(d, "/") > 0 And InStr(d, "/") < Len(d) Then DoiFromToken = d
End Function

Private Function PmidFromToken(ByVal lowTok As String) As String
    Dim p As Long, s As String
    If Left$(lowTok, 5) = "pmid:" Then
        s = Mid$(lowTok, 6)
    Else
        p = InStr(lowTok, "pubmed.ncbi.nlm.nih.gov/")
        If p > 0 Then
            s = LeadingDigits(Mid$(lowTok, p + 24))
        Else
            p = InStr(lowTok, "/pubmed/")
            If p > 0 Then s = LeadingDigits(Mid$(lowTok, p + 8))
        End If
    End If
    If IsDigits(s) Then PmidFromToken = s
End Function

Private Function PmcFromToken(ByVal lowTok As String) As String
    Dim p As Long, s As String
    If Left$(lowTok, 6) = "pmcid:" Then lowTok = Mid$(lowTok, 7)
    If Left$(lowTok, 3) = "pmc" Then s = LeadingDigits(Mid$(lowTok, 4))
    If Len(s) = 0 Then
        p = InStrRev(lowTok, "/pmc")
        If p > 0 Then s = LeadingDigits(Mid$(lowTok, p + 4))
    End If
    If Len(s) > 0 Then PmcFromToken = s
End Function

' Returns the arXiv ID without version, or "".
Private Function ArxivFromToken(ByVal lowTok As String) As String
    Dim s As String, bare As Boolean, p As Long, cut As Variant

    If InStr(lowTok, "arxiv.org/abs/") > 0 Then
        s = Mid$(lowTok, InStr(lowTok, "arxiv.org/abs/") + 14)
    ElseIf InStr(lowTok, "arxiv.org/pdf/") > 0 Then
        s = Mid$(lowTok, InStr(lowTok, "arxiv.org/pdf/") + 14)
    ElseIf Left$(lowTok, 6) = "arxiv:" Then
        s = Mid$(lowTok, 7)
    Else
        s = lowTok
        bare = True
    End If

    For Each cut In Array("?", "#")
        p = InStr(s, cut)
        If p > 0 Then s = Left$(s, p - 1)
    Next
    If Right$(s, 1) = "/" Then s = Left$(s, Len(s) - 1)
    If Right$(s, 4) = ".pdf" Then s = Left$(s, Len(s) - 4)
    p = InStrRev(s, "v")
    If p > 1 Then
        If IsDigits(Mid$(s, p + 1)) Then s = Left$(s, p - 1)
    End If

    If s Like "####.####" Or s Like "####.#####" Then
        p = Val(Mid$(s, 3, 2))           ' YYMM: month must be 01-12
        If p >= 1 And p <= 12 Then ArxivFromToken = s
    ElseIf Not bare And s Like "[a-z]*/#######" Then
        ArxivFromToken = s               ' old style, e.g. hep-th/9711200
    End If
End Function

' ============================================================ lookups

Private Function DoiFromPubMed(ByVal db As String, ByVal num As String) As String
    Dim root As Collection, articleIds As Collection, a As Variant
    Set root = ParseJson(HttpGet("https://eutils.ncbi.nlm.nih.gov/entrez/eutils/esummary.fcgi?db=" & db & _
                                 "&id=" & num & "&retmode=json", "application/json"))
    Set articleIds = JObj(root, "result", num, "articleids")
    If articleIds Is Nothing Then Exit Function
    For Each a In articleIds
        If JText(a, "idtype") = "doi" Then
            DoiFromPubMed = JText(a, "value")
            Exit Function
        End If
    Next
End Function

' Turns CSL-JSON into a Word <b:Source> element. Returns "" when the JSON is unusable.
Private Function BuildSourceXml(ByVal json As String, ByVal doi As String, tags As Collection) As String
    Dim root As Collection, cslType As String, st As String, container As String, publisher As String
    Dim yr As String, firstFamily As String, authors As String, x As String

    Set root = ParseJson(json)
    If root Is Nothing Then Exit Function
    If Len(JText(root, "title")) = 0 Then Exit Function

    cslType = LCase$(JText(root, "type"))
    Select Case cslType
        Case "book", "monograph", "edited-book", "reference-book"
            st = "Book"
        Case "chapter", "book-chapter", "book-section", "book-part", "reference-entry"
            st = "BookSection"
        Case "paper-conference", "proceedings-article"
            st = "ConferenceProceedings"
        Case "report", "report-component"
            st = "Report"
        Case Else
            st = "JournalArticle"
    End Select

    container = JText(root, "container-title")
    publisher = JText(root, "publisher")
    yr = IssuedYear(root)
    authors = AuthorsXml(JObj(root, "author"), firstFamily)

    x = "<b:Source xmlns:b=""" & BIB_NS & """>" & _
        El("Tag", MakeTag(firstFamily, yr, tags)) & El("SourceType", st) & El("LCID", LCID_EN_US) & authors & _
        El("Title", JText(root, "title")) & El("Year", yr)
    Select Case st
        Case "JournalArticle"
            If Len(container) = 0 Then container = publisher     ' preprints: "arXiv", "openRxiv"
            x = x & El("JournalName", container)
        Case "BookSection"
            x = x & El("BookTitle", container)
        Case "ConferenceProceedings"
            x = x & El("ConferenceName", container)
    End Select
    x = x & El("Volume", JText(root, "volume")) & El("Issue", JText(root, "issue")) & _
        El("Pages", JText(root, "page")) & El("Publisher", publisher) & _
        El("City", JText(root, "publisher-place")) & El("DOI", doi) & "</b:Source>"
    BuildSourceXml = x
End Function

Private Function IssuedYear(root As Collection) As String
    Dim k As Variant, y As String
    For Each k In Array("issued", "published-print", "published-online", "published", "created")
        y = JText(root, CStr(k), "date-parts", 1, 1)
        If Len(y) > 0 Then IssuedYear = y: Exit Function
    Next
End Function

Private Function AuthorsXml(authors As Collection, ByRef firstFamily As String) As String
    Dim a As Variant, fam As String, giv As String, lit As String, persons As String, corporate As String
    If authors Is Nothing Then Exit Function
    For Each a In authors
        fam = JText(a, "family")
        giv = JText(a, "given")
        lit = JText(a, "literal")
        If Len(lit) = 0 Then lit = JText(a, "name")
        If Len(fam) > 0 Then
            If Len(firstFamily) = 0 Then firstFamily = fam
            persons = persons & "<b:Person><b:Last>" & XmlEscape(fam) & "</b:Last>" & El("First", giv) & "</b:Person>"
        ElseIf Len(lit) > 0 And Len(corporate) = 0 Then
            corporate = lit
        End If
    Next
    If Len(persons) > 0 Then
        AuthorsXml = "<b:Author><b:Author><b:NameList>" & persons & "</b:NameList></b:Author></b:Author>"
    ElseIf Len(corporate) > 0 Then
        firstFamily = corporate
        AuthorsXml = "<b:Author><b:Author><b:Corporate>" & XmlEscape(corporate) & "</b:Corporate></b:Author></b:Author>"
    End If
End Function

' Tag like Smith2020, Smith2020a, Smith2020b ... unique within the document.
Private Function MakeTag(ByVal family As String, ByVal yr As String, tags As Collection) As String
    Dim i As Long, base As String, tag As String
    For i = 1 To Len(family)
        If Mid$(family, i, 1) Like "[A-Za-z]" Then base = base & Mid$(family, i, 1)
    Next
    If Len(base) = 0 Then base = "Ref"
    base = base & yr
    tag = base
    i = 0
    Do While HasKey(tags, tag)
        i = i + 1
        If i <= 26 Then tag = base & Chr$(96 + i) Else tag = base & "_" & i
    Loop
    AddKey tags, tag
    MakeTag = tag
End Function

' ============================================================ HTTP and files

' GET a URL and return the body as text (UTF-8), or "" on any failure.
Private Function HttpGet(ByVal url As String, ByVal accept As String) As String
    Dim b() As Byte
#If Mac Then
    ' url only ever contains percent-encoded / digit-only parts, so single quotes are safe
    b = HexToBytes(ShellRead("curl -sfL --max-time 30 -A '" & USER_AGENT & "' -H 'Accept: " & accept & _
                             "' '" & url & "' | od -An -v -tx1"))
#Else
    Dim h As Object
    On Error GoTo Fail
    Set h = CreateObject("WinHttp.WinHttpRequest.5.1")
    h.Open "GET", url, False
    h.SetTimeouts 10000, 10000, 15000, 30000
    h.SetRequestHeader "Accept", accept
    h.SetRequestHeader "User-Agent", USER_AGENT
    h.Send
    If h.Status <> 200 Then Exit Function
    b = h.ResponseBody
    On Error GoTo 0
#End If
    HttpGet = Utf8Decode(b)
    Exit Function
Fail:
End Function

#If Mac Then
Private Function ShellRead(ByVal cmd As String) As String
    Dim f As LongPtr, chunk As String, n As Long, out As String
    f = popen(cmd, "r")
    If f = 0 Then Exit Function
    Do While feof(f) = 0
        chunk = Space$(4096)
        n = fread(chunk, 1, Len(chunk) - 1, f)
        If n > 0 Then out = out & Left$(chunk, n)
    Loop
    pclose f
    ShellRead = out
End Function

Private Function HexToBytes(ByVal h As String) As Byte()
    Dim b() As Byte, i As Long, n As Long
    h = UCase$(Replace(Replace(Replace(h, " ", ""), vbLf, ""), vbCr, ""))
    n = Len(h) \ 2
    If n = 0 Then Exit Function
    ReDim b(0 To n - 1)
    For i = 0 To n - 1
        b(i) = CByte(Val("&H" & Mid$(h, 2 * i + 1, 2)) And &HFF)
    Next
    HexToBytes = b
End Function
#End If

Private Function PickFile() As String
    Dim p As String
#If Mac Then
    On Error Resume Next
    p = MacScript("try" & vbLf & _
                  "POSIX path of (choose file with prompt ""Choose a .txt or .csv file with DOIs"")" & vbLf & _
                  "on error" & vbLf & "return """"" & vbLf & "end try")
    If Err.Number <> 0 Then
        Err.Clear
        p = InputBox("The file picker is not available." & vbCr & _
                     "Paste the full path of the file (Finder: select the file, then Option+Cmd+C):", APP_TITLE)
    End If
    On Error GoTo 0
#Else
    With Application.FileDialog(3)            ' msoFileDialogFilePicker
        .Title = "Choose a .txt or .csv file with DOIs"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "DOI lists", "*.txt;*.csv"
        .Filters.Add "All files", "*.*"
        If .Show = -1 Then p = .SelectedItems(1)
    End With
#End If
    PickFile = Trim$(p)
End Function

Private Function ReadTextFile(ByVal path As String) As String
    Dim f As Integer, b() As Byte, s As String
#If Mac Then
    If Not GrantAccessToMultipleFiles(Array(path)) Then Exit Function
#End If
    On Error GoTo Fail
    f = FreeFile
    Open path For Binary Access Read As #f
    If LOF(f) = 0 Then Close #f: Exit Function
    ReDim b(0 To LOF(f) - 1)
    Get #f, , b
    Close #f
    On Error GoTo 0

    s = ""
    If UBound(b) >= 1 Then
        If b(0) = &HFF And b(1) = &HFE Then s = b    ' UTF-16 LE ("Unicode text" from Excel)
    End If
    If Len(s) = 0 Then s = Utf8Decode(b)
    If Left$(s, 1) = ChrW(&HFEFF&) Then s = Mid$(s, 2)
    ReadTextFile = s
    Exit Function
Fail:
    On Error Resume Next
    If f > 0 Then Close #f
End Function

' ============================================================ minimal JSON parser
' Objects and arrays become Collections (objects keyed by name). Values: String, Double, Boolean, Null.

Private Function ParseJson(ByVal s As String) As Collection
    jS = s: jN = Len(s): jI = 1
    On Error GoTo Bad
    JSkip
    Select Case Mid$(jS, jI, 1)
        Case "{": Set ParseJson = JContainer(True)
        Case "[": Set ParseJson = JContainer(False)
    End Select
Bad:
End Function

Private Function JValue() As Variant
    JSkip
    Select Case Mid$(jS, jI, 1)
        Case "{": Set JValue = JContainer(True)
        Case "[": Set JValue = JContainer(False)
        Case """": JValue = JString()
        Case "t": JValue = True: jI = jI + 4
        Case "f": JValue = False: jI = jI + 5
        Case "n": JValue = Null: jI = jI + 4
        Case Else: JValue = JNumber()
    End Select
End Function

Private Function JContainer(ByVal isObj As Boolean) As Collection
    Dim c As New Collection, k As String, closer As String
    If isObj Then closer = "}" Else closer = "]"
    jI = jI + 1
    JSkip
    If Mid$(jS, jI, 1) = closer Then
        jI = jI + 1
    Else
        Do
            If isObj Then
                JSkip
                k = JString()
                JSkip
                jI = jI + 1                       ' ":"
                On Error Resume Next              ' repeated or empty keys are ignored
                c.Add JValue(), k
                On Error GoTo 0
            Else
                c.Add JValue()
            End If
            JSkip
            If Mid$(jS, jI, 1) <> "," Then jI = jI + 1: Exit Do
            jI = jI + 1
        Loop
    End If
    Set JContainer = c
End Function

Private Function JString() As String
    Dim r As String, q As Long, b As Long, c As String
    jI = jI + 1
    Do
        q = InStr(jI, jS, """")
        If q = 0 Then jI = jN + 1: Exit Do
        b = InStr(jI, jS, "\")
        If b = 0 Or b > q Then
            r = r & Mid$(jS, jI, q - jI)
            jI = q + 1
            Exit Do
        End If
        r = r & Mid$(jS, jI, b - jI)
        c = Mid$(jS, b + 1, 1)
        Select Case c
            Case "n": r = r & vbLf
            Case "r": r = r & vbCr
            Case "t": r = r & vbTab
            Case "b", "f"
            Case "u"
                r = r & ChrW(Val("&H" & UCase$(Mid$(jS, b + 2, 4))) And &HFFFF&)
                b = b + 4
            Case Else: r = r & c
        End Select
        jI = b + 2
    Loop
    JString = r
End Function

Private Function JNumber() As Double
    Dim p As Long
    p = jI
    Do While jI <= jN
        If InStr("+-0123456789.eE", Mid$(jS, jI, 1)) = 0 Then Exit Do
        jI = jI + 1
    Loop
    JNumber = Val(Mid$(jS, p, jI - p))
End Function

Private Sub JSkip()
    Do While jI <= jN
        Select Case Mid$(jS, jI, 1)
            Case " ", vbTab, vbCr, vbLf: jI = jI + 1
            Case Else: Exit Do
        End Select
    Loop
End Sub

' Walks keys (names for objects, 1-based numbers for arrays). False if a step is missing.
Private Function JWalk(ByVal node As Variant, ByVal keys As Variant, ByRef out As Variant) As Boolean
    Dim i As Long
    If Not IsObject(node) Then Exit Function
    Set out = node
    For i = LBound(keys) To UBound(keys)
        If Not IsObject(out) Then Exit Function
        If Not TryItem(out, keys(i), out) Then Exit Function
    Next
    JWalk = True
End Function

Private Function TryItem(ByVal c As Collection, ByVal k As Variant, ByRef out As Variant) As Boolean
    On Error GoTo Missing
    If IsObject(c(k)) Then Set out = c(k) Else out = c(k)
    TryItem = True
Missing:
End Function

Private Function JObj(ByVal node As Variant, ParamArray keys() As Variant) As Collection
    Dim v As Variant
    If JWalk(node, keys, v) Then
        If IsObject(v) Then Set JObj = v
    End If
End Function

' Value as cleaned text; for an array (e.g. container-title) the first element.
Private Function JText(ByVal node As Variant, ParamArray keys() As Variant) As String
    Dim v As Variant, first As Variant
    If Not JWalk(node, keys, v) Then Exit Function
    If IsObject(v) Then
        If v Is Nothing Then Exit Function
        If v.Count = 0 Then Exit Function
        If IsObject(v(1)) Then Exit Function
        first = v(1)
        JText = CleanText(ScalarText(first))
    Else
        JText = CleanText(ScalarText(v))
    End If
End Function

Private Function ScalarText(ByVal v As Variant) As String
    Select Case VarType(v)
        Case vbString
            ScalarText = v
        Case vbDouble, vbSingle, vbLong, vbInteger, vbCurrency, vbDecimal
            If v = Int(v) And Abs(v) < 2000000000# Then ScalarText = CStr(CLng(v)) Else ScalarText = CStr(v)
    End Select
End Function

' ============================================================ text helpers

' Removes HTML/JATS tags (<i>, <sub>, <jats:p>), decodes basic entities, collapses whitespace.
Private Function CleanText(ByVal s As String) As String
    Dim p As Long, q As Long, i As Long, nxt As String
    p = InStr(s, "<")
    Do While p > 0
        nxt = Mid$(s, p + 1, 1)
        q = InStr(p, s, ">")
        If q > 0 And (nxt Like "[A-Za-z]" Or nxt = "/") Then
            s = Left$(s, p - 1) & Mid$(s, q + 1)
            p = InStr(p, s, "<")
        Else
            p = InStr(p + 1, s, "<")
        End If
    Loop
    s = Replace(s, "&lt;", "<"): s = Replace(s, "&gt;", ">")
    s = Replace(s, "&quot;", """"): s = Replace(s, "&#39;", "'"): s = Replace(s, "&amp;", "&")
    For i = 1 To Len(s)
        If AscW(Mid$(s, i, 1)) >= 0 And AscW(Mid$(s, i, 1)) < 32 Then Mid(s, i, 1) = " "
    Next
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    CleanText = Trim$(s)
End Function

Private Function El(ByVal name As String, ByVal value As String) As String
    If Len(value) > 0 Then El = "<b:" & name & ">" & XmlEscape(value) & "</b:" & name & ">"
End Function

Private Function XmlEscape(ByVal s As String) As String
    s = Replace(s, "&", "&amp;")
    s = Replace(s, "<", "&lt;")
    s = Replace(s, ">", "&gt;")
    XmlEscape = Replace(s, """", "&quot;")
End Function

' Text of the first <b:NAME> (or <NAME>) element in a Source.XML string.
Private Function XmlField(ByVal xml As String, ByVal name As String) As String
    Dim p As Long, q As Long, s As String
    p = InStr(1, xml, ":" & name & ">", vbTextCompare)
    If p = 0 Then p = InStr(1, xml, "<" & name & ">", vbTextCompare)
    If p = 0 Then Exit Function
    p = p + Len(name) + 2
    q = InStr(p, xml, "<")
    If q = 0 Then Exit Function
    s = Mid$(xml, p, q - p)
    s = Replace(s, "&lt;", "<"): s = Replace(s, "&gt;", ">"): s = Replace(s, "&quot;", """")
    XmlField = Trim$(Replace(s, "&amp;", "&"))
End Function

Private Function Utf8Decode(b() As Byte) As String
    Dim lo As Long, hi As Long, i As Long, k As Long, o As Long, c As Long, cp As Long, need As Long, out As String
    On Error Resume Next
    lo = LBound(b): hi = UBound(b)
    If Err.Number <> 0 Then Exit Function     ' empty array
    On Error GoTo 0
    out = Space$(hi - lo + 1)
    i = lo
    Do While i <= hi
        c = b(i)
        If c < &H80 Then
            cp = c: need = 0
        ElseIf c >= &HF0 Then
            cp = c And &H7: need = 3
        ElseIf c >= &HE0 Then
            cp = c And &HF: need = 2
        ElseIf c >= &HC0 Then
            cp = c And &H1F: need = 1
        Else
            cp = &HFFFD&: need = 0                ' stray continuation byte
        End If
        If i + need > hi Then Exit Do
        For k = 1 To need
            cp = cp * &H40 Or (b(i + k) And &H3F)
        Next
        i = i + need + 1
        If cp >= &H10000 Then
            cp = cp - &H10000
            o = o + 1: Mid(out, o, 1) = ChrW(&HD800& + cp \ &H400)
            o = o + 1: Mid(out, o, 1) = ChrW(&HDC00& + (cp And &H3FF))
        Else
            o = o + 1: Mid(out, o, 1) = ChrW(cp)
        End If
    Loop
    Utf8Decode = Left$(out, o)
End Function

Private Function UrlEncodePath(ByVal s As String) As String
    Dim i As Long, ch As String, c As Long, r As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        c = AscW(ch) And &HFFFF&
        If ch Like "[A-Za-z0-9]" Or InStr("-._~/", ch) > 0 Then
            r = r & ch
        ElseIf c < &H80 Then
            r = r & "%" & Right$("0" & Hex$(c), 2)
        ElseIf c < &H800 Then
            r = r & "%" & Hex$(&HC0 Or (c \ &H40)) & "%" & Hex$(&H80 Or (c And &H3F))
        Else
            r = r & "%" & Hex$(&HE0 Or (c \ &H1000)) & "%" & Hex$(&H80 Or ((c \ &H40) And &H3F)) & _
                "%" & Hex$(&H80 Or (c And &H3F))
        End If
    Next
    UrlEncodePath = r
End Function

' Decodes ASCII %XX escapes only (enough for DOIs in links).
Private Function UrlDecode(ByVal s As String) As String
    Dim i As Long, r As String, h As String
    i = 1
    Do While i <= Len(s)
        h = UCase$(Mid$(s, i + 1, 2))
        If Mid$(s, i, 1) = "%" And h Like "[0-7][0-9A-F]" Then
            r = r & Chr$(Val("&H" & h))
            i = i + 3
        Else
            r = r & Mid$(s, i, 1)
            i = i + 1
        End If
    Loop
    UrlDecode = r
End Function

Private Function LeadingDigits(ByVal s As String) As String
    Dim i As Long
    i = 1
    Do While Mid$(s, i, 1) Like "#"
        i = i + 1
    Loop
    LeadingDigits = Left$(s, i - 1)
End Function

Private Function IsDigits(ByVal s As String) As Boolean
    IsDigits = Len(s) > 0 And Not s Like "*[!0-9]*"
End Function

Private Function CountOf(ByVal s As String, ByVal ch As String) As Long
    CountOf = Len(s) - Len(Replace(s, ch, ""))
End Function

Private Function HasKey(c As Collection, ByVal key As String) As Boolean
    Dim v As Variant
    On Error GoTo Missing
    v = IsObject(c(key))
    HasKey = True
Missing:
End Function

Private Sub AddKey(c As Collection, ByVal key As String)
    On Error Resume Next
    c.Add key, key
End Sub

Private Function ListText(c As Collection, ByVal maxItems As Long) As String
    Dim i As Long
    For i = 1 To c.Count
        If i > maxItems Then ListText = ListText & vbCr & "   ... and " & (c.Count - maxItems) & " more": Exit For
        ListText = ListText & vbCr & "   " & c(i)
    Next
End Function

Private Sub ShowStatus(ByVal msg As String)
    On Error Resume Next
    Application.StatusBar = msg
    DoEvents
End Sub
