Attribute VB_Name = "Module11"
Sub SendEmails()
    LastRow = Range("A" & Rows.Count).End(xlUp).Row
    i = 2
    Do Until i > LastRow
        If Range("K" & i) = "" Then
            Set mApp = CreateObject("Outlook.Application")
            Set mMailStudent = mApp.CreateItem(0)
            Set mMailTeacher = mApp.CreateItem(0)
            rowEmail = 0
            lRowEmails = Worksheets("ClassList").Range("A" & Rows.Count).End(xlUp).Row
            j = 2
            Do Until j > lRowEmails
                If Worksheets("ClassList").Range("B" & j) = Range("E" & i) And Worksheets("ClassList").Range("C" & j) = Range("F" & i) Then
                    answer = MsgBox("Is " & Worksheets("ClassList").Range("D" & j) & " the correct name?", vbQuestion + vbYesNo + vbDefaultButton2, "Confirmation")
                    If answer = vbNo Then
                        MsgBox ("Program killed")
                        Exit Sub
                    End If
                    mailTo = Worksheets("ClassList").Range("E" & j)
                    rowEmail = j
                End If
                If Worksheets("ClassList").Range("D" & j) = Range("A" & i) Then
                    prefectEmail = Worksheets("ClassList").Range("E" & j)
                End If
                If Worksheets("ClassList").Range("F" & j) = Range("A" & i) Then
                    prefectEmail = Worksheets("ClassList").Range("G" & j)
                End If
                If Worksheets("ClassList").Range("H" & j) = Range("A" & i) Then
                    prefectEmail = Worksheets("ClassList").Range("I" & j)
                End If
                j = j + 1
            Loop
            
            ' STUDENT MAIL
            With mMailStudent
            '.To = mailTo
            .To = "25YANGJ254C@student.ri.edu.sg"
            '.Bcc = prefectEmail & "; 26YTANJ353F@student.ri.edu.sg; 26YLAUJ662F@student.ri.edu.sg; 27YJOSI038H@student.ri.edu.sg"
            .Subject = "Booking registered on " & Range("B" & i)
            .Body = "Dear " & Range("D" & i) & " of Class " & Range("E" & i) & " Index Number " & Range("F" & i) & "," & Chr(10) & Chr(10) & "This is to notify you that you were booked for committing the offence of """ & Range("H" & i) & """ on " & Range("B" & i) & " (" & Range("C" & i) & " hrs)." & Chr(10) & Chr(10) & "Your Form Teachers and Year Head have also been informed of your offence, as well as any suggested consequences, via email." & Chr(10) & Chr(10) & "We hope that you understand the rationale behind this booking, and will endeavour to carry yourself appropriately as a Rafflesian with self-discipline. You are warned that accumulation of multiple bookings will lead to further consequences, including but not limited to detentions, conduct slips, meetings with your year head/discipline master etc." & Chr(10) & Chr(10) & Chr(10) & "This is an automated message, please do not reply."
            .Send
            End With
            Application.Wait (Now + TimeValue("0:00:5"))
            
            lRowOffences = Worksheets("Offenders List").Range("A" & Rows.Count).End(xlUp).Row
            k = 2
            Do Until k > lRowOffences
                If Worksheets("Offenders List").Range("C" & k) = Range("D" & 2) Then
                    noOfOffences = Worksheets("Offenders List").Range("D" & k)
                End If
            k = k + 1
            Loop
            
            ' TEACHER MAIL
            With mMailTeacher
             '.To = Worksheets("ClassList").Range("G" & rowEmail) & "; " & Worksheets("ClassList").Range("I" & rowEmail)
              .To = "25YANGJ254C@student.ri.edu.sg"
             '.Bcc = prefectEmail & "; 26YTANJ353F@student.ri.edu.sg; 26YLAUJ662F@student.ri.edu.sg; 27YJOSI038H@student.ri.edu.sg"
             .Subject = "Booking of Student from Class " & Range("E" & i)
            .Body = "Dear " & Worksheets("ClassList").Range("F" & rowEmail) & " and " & Worksheets("ClassList").Range("H" & rowEmail) & "," & Chr(10) & Chr(10) & "Your student, " & Range("D" & i) & ", Class " & Range("E" & i) & ", Index Number " & Range("F" & i) & " committed an offence of """ & Range("H" & i) & """ on " & Range("B" & i) & " (" & Range("C" & i) & " hrs)." & Chr(10) & Chr(10) & "This booking was made by prefect/teacher " & Range("A" & i) & "." & Chr(10) & Chr(10) & "You may wish to consider the number of bookings this student has accumulated and whether he has demonstrated recalcitrant behaviour. This student has committed " & noOfOffences & " offences in the past year. The consolidated year-to-date bookings for your own form class can be requested from Mr. Phua Zhengjie if required." & Chr(10) & Chr(10) & "Thank you for your attention and efforts in upholding the discipline standards of Rafflesians." & Chr(10) & Chr(10) & Chr(10) & "This is an automated message, please do not reply."
            .Send
            End With
            Application.Wait (Now + TimeValue("0:00:05"))
            Range("K" & i) = "sent"
        End If
        i = i + 1
    Loop
End Sub
