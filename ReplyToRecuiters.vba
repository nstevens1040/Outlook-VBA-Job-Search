Sub ReplyAllWithAttachments()
    Dim fso As Object
    Dim ts As Object
    Dim du As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.GetFile(Environ("appdata") & "\Microsoft\Signatures\nanicksig.htm").OpenAsTextStream(1, -2)
    Signature = ts.readall
    ts.Close
    Set du = fso.GetFile(".\commuter_map_data_uri_.txt").OpenAsTextStream(1, -2)
    DataUri = du.readall
    du.Close
    Dim oReply As Outlook.MailItem
    Dim oItem As Object
    Set oItem = GetCurrentItem()
    sender = GetSenderFName(oItem)
    dateString = GetDateFriday()
    olDefaultReply = "<!DOCTYPE html>" & vbCrLf & "<html>" & vbCrLf & "    <head>" & vbCrLf & "        <meta charset=" & Chr(34) & "utf-8" & Chr(34) & "/>" & vbCrLf & "        <style type=" & Chr(34) & "text/css" & Chr(34) & ">" & vbCrLf & "            .pmain {" & vbCrLf & "                font-size: 12.0pt;" & vbCrLf & "                color: black;" & vbCrLf & "                margin-left: 41.05pt;" & vbCrLf & "            }" & vbCrLf & "            .pq {" & vbCrLf & "                margin-left: 41.05pt;" & vbCrLf & "            }" & vbCrLf & "            .bq {" & vbCrLf & "                font-size: 12.0pt;" & vbCrLf & "                color: black;" & vbCrLf & "            }" & vbCrLf & "            .mapimg {" & vbCrLf & "                margin-right: 0in;" & vbCrLf & "                margin-bottom: .25in;" & vbCrLf & "                margin-left: .5in;" & vbCrLf & "                line-height: 105%;" & vbCrLf & "            }" & vbCrLf & "        </style>" & vbCrLf & _
    "    </head>" & vbCrLf & "    <body>" & vbCrLf & "        <div>" & vbCrLf & "            <div>" & vbCrLf & "                <p class=" & Chr(34) & "pmain" & Chr(34) & ">Hello " & sender & ", </p>" & vbCrLf & "                <p class=" & Chr(34) & "pmain" & Chr(34) & ">This is Nicholas Stevens and I" & Chr(39) & "m emailing to let you know that I am still currently looking for work. Here I will attempt to answer questions that I am commonly asked by recruiters. </p>" & vbCrLf & "                <p class=" & Chr(34) & "pq" & Chr(34) & "><b class=" & Chr(34) & "bq" & Chr(34) & ">What is your minimum salary requirement?</b></p>" & vbCrLf & "                <p class=" & Chr(34) & "pmain" & Chr(34) & ">I" & Chr(39) & "m asking for $55,000 annually with medical benefits (PPO), dental, vision, and 401k available and $29 per hour for contract jobs that do not provide a healthcare benefits package. </p>" & vbCrLf & _
    "                <p class=" & Chr(34) & "pq" & Chr(34) & "><b class=" & Chr(34) & "bq" & Chr(34) & ">Where do you live and to where do you have the ability to commute?</b></p>" & vbCrLf & "                <p class=" & Chr(34) & "pmain" & Chr(34) & ">I live in Chicago, nine miles North of the loop (downtown) near Loyola. CTA accessibility is a factor in my job search as I do not own a vehicle. The area in green on the map below represents a radius of 60 minutes commuting via CTA from my home. For positions that require me to report to an office everyday I" & Chr(39) & "m hoping the map will quickly tell you how feasible my commute would be. That being said, I am foremost considering future employers that offer a remote work option or roles that are not dependent on my ability to appear anywhere in person. </p>" & vbCrLf & _
    "                <p style=" & Chr(34) & "margin-left: 41.05pt;" & Chr(34) & ">" & vbCrLf & _
    "                    <img class=" & Chr(34) & "mapimg" & Chr(34) & " width=" & Chr(34) & "454" & Chr(34) & " height=" & Chr(34) & "574" & Chr(34) & " src=" & Chr(34) & DataUri & Chr(34) & " alt=" & Chr(34) & "one" & Chr(34) & ">" & vbCrLf & "                <p class=" & Chr(34) & "pq" & Chr(34) & "><b class=" & Chr(34) & "bq" & Chr(34) & ">Are you willing to take on contract or contract to hire roles?</b></p>" & vbCrLf & "                <p class=" & Chr(34) & "pmain" & Chr(34) & ">I will look at full-time, direct hire opportunities first, but I will also evaluate contract to hire roles to see if they meet my conditions for taking on that type of work. I am looking to make a long-term commitment to a firm as a technical support engineer. As far as contract roles are concerned, I am more willing to look at openings that will expand and build on my talent rather than openings that don" & Chr(39) & "t add anything new to my resume. </p>" & vbCrLf & _
    "                </p>" & vbCrLf & _
    "                <p class=" & Chr(34) & "pq" & Chr(34) & "><b class=" & Chr(34) & "bq" & Chr(34) & ">When will you be available to work?</b></p>" & vbCrLf & "                <p class=" & Chr(34) & "pmain" & Chr(34) & ">Immediately, and to give you an idea of my timeline; my goal is to have interviews set up for this week as well as next week and then to accept an offer by " & dateString & ". </p>" & vbCrLf & "                <p class=" & Chr(34) & "pq" & Chr(34) & "><b class=" & Chr(34) & "bq" & Chr(34) & ">What kind of work are you looking for?</b></p>" & vbCrLf & _
    "                <p class=" & Chr(34) & "pmain" & Chr(34) & ">I am looking for openings in technical support as it is my tried and true talent that I have seven years of experience with. Full time, direct-hire positions are preferred as I am looking to make a very long-term commitment to a firm, but I will also give contract positions consideration as well. The distinction I draw relates to the scope of work. As I stated earlier, I am more willing to take on contract roles that will expand and build on my talent rather than contract roles that don" & Chr(39) & "t add anything new to my resume. I tend to do my best work as a primary contact point for users who encounter technical obstacles. Contextually, I do prefer to be working with a firm where technical support is an essential, yet custodial role within the firm" & Chr(39) & "s business model and distinct from production staff. </p>" & vbCrLf & _
    "                <p class=" & Chr(34) & "pmain" & Chr(34) & ">That being said, I do also have experience coding, debugging, and deploying PowerShell scripts into production that perform infrastructure dependent tasks, generate reports, send automated emails, configure laptops for new users, etc. Building on this, I am also looking for a way to gain experience developing in C# using .NET Framework. I have a few finished projects on my <a href=" & Chr(34) & "https://github.com/nstevens1040" & Chr(34) & " target=" & Chr(34) & "_blank" & Chr(34) & ">GitHub</a> page if you would like to get a better idea of my abilities as a developer.</p>" & vbCrLf & _
    "                <p class=" & Chr(34) & "pmain" & Chr(34) & ">Additionally, I have a lot of exposure to the methods used to collect, organize, and query for specific information within large datasets using STATA, R, Python, and SAS to interoperate with SQL databases. I" & Chr(39) & "ve also taken a personal interest in the various means of data collection via web scraping with Python and PowerShell. </p>" & vbCrLf & _
    "                <p class=" & Chr(34) & "pmain" & Chr(34) & ">Please refer to my resume for any general questions regarding my work history. I" & Chr(39) & "m also more than happy to answer any specific questions you might have. You can view or download my CV as <a href=" & Chr(34) & "https://nanick.hopto.org/resume" & Chr(34) & " target=" & Chr(34) & "_blank" & Chr(34) & ">HTML</a>, <a href=" & Chr(34) & "https://nanick.hopto.org/resumedocx" & Chr(34) & " target=" & Chr(34) & "_blank" & Chr(34) & ">DOCX</a>, or <a href=" & Chr(34) & "https://nanick.hopto.org/resumepdf" & Chr(34) & " target=" & Chr(34) & "_blank" & Chr(34) & ">PDF</a>. </p>" & vbCrLf & _
    "                <p class=" & Chr(34) & "pmain" & Chr(34) & ">I appreciate your taking the time to read this and for considering my candidacy for any open positions. I am eager to further discuss how my knowledge and experience will meet the IT needs of my future employer. Please do not hesitate to reach back out to me at <a href=" & Chr(34) & "tel:+12242232299" & Chr(34) & " target=" & Chr(34) & "_blank" & Chr(34) & ">(224) 223-2299</a> or by replying to this email directly. </p>" & vbCrLf & "                <p class=" & Chr(34) & "pmain" & Chr(34) & ">Thank you, </p>" & vbCrLf & "            </div>" & vbCrLf & "        </div>" & vbCrLf & "    </body>" & vbCrLf & "</html><br>" & Signature
    If Not oItem Is Nothing Then
        Set oReply = oItem.ReplyAll
        CopyAttachments oReply
        oReply.BodyFormat = olFormatHTML
        oReply.HTMLBody = olDefaultReply
        oReply.Display
    End If
    Set oReply = Nothing
    Set oItem = Nothing
End Sub
Function GetSenderFName(olMailItem) As String
    fullName = olMailItem.SenderName
    If InStr(fullName, ",") <> 0 Then
        fName = Replace(Split(fullName, ",")(1), " ", "")
    Else
        fName = Replace(Split(fullName, " ")(0), " ", "")
    End If
    GetSenderFName = StrConv(fName, vbProperCase)
End Function
Function GetCurrentItem() As Object
    Dim objApp As Outlook.Application
    Set objApp = Application
    On Error Resume Next
    Select Case TypeName(objApp.ActiveWindow)
        Case "Explorer"
            Set GetCurrentItem = objApp.ActiveExplorer.Selection.Item(1)
        Case "Inspector"
            Set GetCurrentItem = objApp.ActiveInspector.CurrentItem
    End Select
    Set objApp = Nothing
End Function
 
Sub CopyAttachments(objTargetItem)
   strPDF = ".\resume\Nicholas Stevens CV 2020 .pdf"
   strDOCX = ".\resume\Nicholas Stevens CV 2020 .docx"
   objTargetItem.Attachments.Add strPDF
   objTargetItem.Attachments.Add strDOCX
End Sub
Function GetDateFriday() As String
    Dim dArray(3)
    Dim now
    now = Date
    Dim count As Integer
    count = 0
    For i = 1 To 19
        dayIter = DateAdd("d", i, now)
        If Weekday(dayIter) = 6 Then
            dStr = Left(Replace(Format(dayIter, "Long Date"), Split(Format(dayIter, "Long Date"), ",")(2), ""), Len(Replace(Format(dayIter, "Long Date"), Split(Format(dayIter, "Long Date"), ",")(2), "")) - 1)
            dArray(count) = dStr
            count = count + 1
        End If
    Next i
    GetDateFriday = dArray(count - 1)
End Function
Sub RunReply()
    ReplyAllWithAttachments
End Sub
