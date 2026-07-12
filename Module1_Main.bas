Attribute VB_Name = "Module1_Main"
Option Explicit

'==========================================================
' ENTER YOUR AI API KEY HERE
' Works with Gemini by default
' See AIRequest function below to switch to OpenAI or others
'==========================================================
Const API_KEY As String = "YOUR_API_KEY_HERE"

'==========================================================
' MAIN MACRO
' Only processes rows where Column E is empty
'==========================================================
Sub first()

    Dim ws As Worksheet
    Dim LastRow As Long
    Dim i As Long
    Dim SkippedCount As Long
    Dim ProcessedCount As Long

    Dim ExpenseDate As String
    Dim Description As String
    Dim Amount As String
    Dim PaymentMethod As String

    Dim Result1 As String
    Dim Result2 As String

    Dim arr1() As String
    Dim arr2() As String

    Dim Category As String
    Dim Necessity As String
    Dim Budget As String

    Set ws = ActiveSheet
    LastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    SkippedCount   = 0
    ProcessedCount = 0

    Debug.Print "=========================================="
    Debug.Print "START - " & Now()
    Debug.Print "Last Row: " & LastRow
    Debug.Print "=========================================="

    If LastRow < 2 Then
        MsgBox "No data found!", vbCritical
        Exit Sub
    End If

    Application.ScreenUpdating = False

    For i = 2 To LastRow

        '==================================================
        ' SKIP ROW IF ALREADY PROCESSED
        '==================================================
        If Trim(ws.Cells(i, 5).Value) <> "" And _
           ws.Cells(i, 5).Value <> "PARSE ERROR" And _
           ws.Cells(i, 5).Value <> "API ERROR" Then

            Debug.Print "--- ROW " & i & " SKIPPED (already processed) ---"
            SkippedCount = SkippedCount + 1
            GoTo NextRow

        End If

        ExpenseDate   = ws.Cells(i, 1).Text
        Description   = ws.Cells(i, 2).Value
        Amount        = CStr(ws.Cells(i, 3).Value)
        PaymentMethod = ws.Cells(i, 4).Value

        '=== Skip if row has no description ===
        If Trim(Description) = "" Then
            Debug.Print "--- ROW " & i & " SKIPPED (empty description) ---"
            SkippedCount = SkippedCount + 1
            GoTo NextRow
        End If

        Debug.Print "--- ROW " & i & " PROCESSING ---"
        Debug.Print "Input: " & ExpenseDate & " | " & Description & " | " & Amount & " | " & PaymentMethod

        '==================================================
        ' FIRST AI CALL — Classification
        '==================================================
        Result1 = AIRequestWithRetry(BuildClassificationPrompt( _
                    ExpenseDate, Description, Amount, PaymentMethod))

        Debug.Print "CALL 1 Result: [" & Result1 & "]"

        arr1 = Split(Result1, "|")

        If UBound(arr1) >= 2 Then
            Category  = Trim(arr1(0))
            Necessity = Trim(arr1(1))
            Budget    = Trim(arr1(2))
        Else
            Category  = "PARSE ERROR"
            Necessity = "PARSE ERROR"
            Budget    = "PARSE ERROR"
            Debug.Print "WARNING: CALL 1 parse failed. Raw=[" & Result1 & "]"
        End If

        ws.Cells(i, 5).Value = Category
        ws.Cells(i, 6).Value = Necessity
        ws.Cells(i, 7).Value = Budget

        Call WaitSeconds(2)

        '==================================================
        ' SECOND AI CALL — Advice
        '==================================================
        Result2 = AIRequestWithRetry(BuildAdvicePrompt( _
                    ExpenseDate, Description, Amount, PaymentMethod, _
                    Category, Necessity, Budget))

        Debug.Print "CALL 2 Result: [" & Result2 & "]"

        arr2 = Split(Result2, "|")

        If UBound(arr2) >= 3 Then
            ws.Cells(i, 8).Value  = Trim(arr2(0))
            ws.Cells(i, 9).Value  = Trim(arr2(1))
            ws.Cells(i, 10).Value = Trim(arr2(2))
            ws.Cells(i, 11).Value = Trim(arr2(3))
        Else
            ws.Cells(i, 8).Value  = "PARSE ERROR"
            ws.Cells(i, 9).Value  = "PARSE ERROR"
            ws.Cells(i, 10).Value = "PARSE ERROR"
            ws.Cells(i, 11).Value = "PARSE ERROR"
            Debug.Print "WARNING: CALL 2 parse failed. Raw=[" & Result2 & "]"
        End If

        ProcessedCount = ProcessedCount + 1

        Call WaitSeconds(2)

        DoEvents

NextRow:

    Next i

    Application.ScreenUpdating = True

    Debug.Print "=========================================="
    Debug.Print "DONE - " & Now()
    Debug.Print "Processed: " & ProcessedCount & " | Skipped: " & SkippedCount
    Debug.Print "=========================================="

    If ProcessedCount = 0 Then
        MsgBox "No new rows found to process!" & vbCrLf & vbCrLf & _
               "All " & SkippedCount & " rows are already analysed." & vbCrLf & vbCrLf & _
               "Add a new row and run again.", _
               vbInformation, "Nothing to Process"
    Else
        MsgBox "Analysis completed!" & vbCrLf & vbCrLf & _
               "✅ New rows processed : " & ProcessedCount & vbCrLf & _
               "⏭️ Rows skipped       : " & SkippedCount, _
               vbInformation, "Done"
    End If

End Sub

'==========================================================
' REPROCESS SELECTED ROW ONLY
'==========================================================
Sub ReprocessSelectedRow()

    Dim ws As Worksheet
    Dim i As Long

    Set ws = ActiveSheet
    i = ActiveCell.Row

    If i < 2 Then
        MsgBox "Please select a data row (Row 2 or below).", vbExclamation
        Exit Sub
    End If

    If Trim(ws.Cells(i, 1).Value) = "" Then
        MsgBox "Selected row appears empty. Please select a valid data row.", vbExclamation
        Exit Sub
    End If

    Dim answer As Integer
    answer = MsgBox("Reprocess Row " & i & "?" & vbCrLf & vbCrLf & _
                    "Description : " & ws.Cells(i, 2).Value & vbCrLf & _
                    "Amount      : " & ws.Cells(i, 3).Value, _
                    vbYesNo + vbQuestion, "Confirm Reprocess")

    If answer = vbNo Then Exit Sub

    ws.Range(ws.Cells(i, 5), ws.Cells(i, 11)).ClearContents

    Debug.Print "Row " & i & " cleared — reprocessing..."

    Call first()

End Sub

'==========================================================
' AI REQUEST WITH AUTO RETRY ON 429
'==========================================================
Function AIRequestWithRetry(Prompt As String) As String

    Dim attempt As Integer
    Dim result As String

    For attempt = 1 To 3

        result = AIRequest(Prompt)

        If Left(result, 9) <> "429_ERROR" Then
            AIRequestWithRetry = result
            Exit Function
        End If

        Debug.Print "  429 Rate Limit hit. Attempt " & attempt & " of 3. Waiting " & (attempt * 10) & "s..."
        Call WaitSeconds(attempt * 10)

    Next attempt

    Debug.Print "  All retries failed."
    AIRequestWithRetry = "API ERROR|RATE LIMIT|429|RETRY FAILED"

End Function

'==========================================================
' AI API CALL
' Default: Gemini 2.5 Flash
' To switch provider: update URL and JSON body below
'==========================================================
Function AIRequest(Prompt As String) As String

    Dim Http As Object
    Dim JSON As String
    Dim URL As String
    Dim Response As String

    '==========================================================
    ' GEMINI — currently active
    '==========================================================
    URL = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=" & API_KEY
    JSON = "{""contents"":[{""parts"":[{""text"":""" & EscapeJSON(Prompt) & """}]}]}"

    '==========================================================
    ' OPENAI — uncomment below and comment out Gemini above
    ' URL = "https://api.openai.com/v1/chat/completions"
    ' JSON = "{""model"":""gpt-4o-mini"",""messages"":[{""role"":""user"",""content"":""" & EscapeJSON(Prompt) & """}]}"
    '==========================================================

    Set Http = CreateObject("MSXML2.XMLHTTP")
    Http.Open "POST", URL, False
    Http.setRequestHeader "Content-Type", "application/json"

    '==========================================================
    ' For OpenAI — uncomment this line:
    ' Http.setRequestHeader "Authorization", "Bearer " & API_KEY
    '==========================================================

    Http.send JSON

    Debug.Print "  HTTP Status: " & Http.Status & " " & Http.statusText

    Response = Http.responseText

    If Http.Status = 429 Then
        Debug.Print "  429 Detected."
        AIRequest = "429_ERROR"
        Exit Function
    End If

    If Http.Status <> 200 Then
        Debug.Print "  ERROR - Status: " & Http.Status
        Debug.Print "  Response: " & Left(Response, 300)
        AIRequest = "API ERROR|CHECK DEBUG|STATUS " & Http.Status & "|N/A"
        Exit Function
    End If

    Debug.Print "  Raw Response (first 400): " & Left(Response, 400)

    AIRequest = ExtractText(Response)

End Function

'==========================================================
' EXTRACT TEXT FROM AI JSON RESPONSE
'==========================================================
Function ExtractText(JSON As String) As String

    Dim s As Long
    Dim extracted As String
    Dim k As Long
    Dim char As String
    Dim prevChar As String

    s = InStr(JSON, """text"": """)

    If s = 0 Then
        s = InStr(JSON, """text"":""")
        If s > 0 Then
            s = s + 8
        End If
    Else
        s = s + 9
    End If

    If s = 0 Then
        Debug.Print "  [ExtractText] ERROR: 'text' key not found!"
        Debug.Print "  [ExtractText] JSON dump: " & Left(JSON, 500)
        ExtractText = ""
        Exit Function
    End If

    extracted = ""
    prevChar  = ""

    For k = s To Len(JSON)
        char = Mid(JSON, k, 1)
        If char = """" And prevChar <> "\" Then Exit For
        extracted = extracted & char
        prevChar  = char
    Next k

    extracted = Replace(extracted, "\n", " ")
    extracted = Replace(extracted, "\r", "")
    extracted = Replace(extracted, "\\", "\")
    extracted = Replace(extracted, "\""", """")

    Debug.Print "  [ExtractText] Extracted: [" & extracted & "]"

    ExtractText = Trim(extracted)

End Function

'==========================================================
' WAIT HELPER
'==========================================================
Sub WaitSeconds(seconds As Integer)
    Dim endTime As Double
    endTime = Timer + seconds
    Do While Timer < endTime
        DoEvents
    Loop
End Sub

'==========================================================
' PROMPT 1 — CLASSIFICATION
'==========================================================
Function BuildClassificationPrompt(ExpenseDate As String, _
                                   Description As String, _
                                   Amount As String, _
                                   PaymentMethod As String) As String

    BuildClassificationPrompt = _
        "You are a finance assistant." & vbCrLf & _
        "Classify this expense." & vbCrLf & _
        "Return ONLY one line with no extra text, no explanation." & vbCrLf & _
        "Format EXACTLY: Category|Necessity|Budget Status" & vbCrLf & _
        "Example: Food|Essential|Within Budget" & vbCrLf & _
        vbCrLf & _
        "Date: " & ExpenseDate & vbCrLf & _
        "Description: " & Description & vbCrLf & _
        "Amount: " & Amount & vbCrLf & _
        "Payment Method: " & PaymentMethod

End Function

'==========================================================
' PROMPT 2 — ADVICE
'==========================================================
Function BuildAdvicePrompt(ExpenseDate As String, _
                           Description As String, _
                           Amount As String, _
                           PaymentMethod As String, _
                           Category As String, _
                           Necessity As String, _
                           Budget As String) As String

    BuildAdvicePrompt = _
        "You are an expert personal finance advisor." & vbCrLf & _
        "Give practical advice." & vbCrLf & _
        "Return ONLY one line with no extra text, no explanation." & vbCrLf & _
        "Format EXACTLY: Saving Suggestion|Monthly Trend|Financial Advice|Savings Potential" & vbCrLf & _
        "Example: Cook at home|Increasing|Reduce dining out|15%" & vbCrLf & _
        vbCrLf & _
        "Date: " & ExpenseDate & vbCrLf & _
        "Description: " & Description & vbCrLf & _
        "Amount: " & Amount & vbCrLf & _
        "Payment Method: " & PaymentMethod & vbCrLf & _
        "Category: " & Category & vbCrLf & _
        "Necessity: " & Necessity & vbCrLf & _
        "Budget Status: " & Budget

End Function

'==========================================================
' ESCAPE JSON SPECIAL CHARACTERS
'==========================================================
Function EscapeJSON(txt As String) As String

    txt = Replace(txt, "\", "\\")
    txt = Replace(txt, """", "\""")
    txt = Replace(txt, vbCrLf, "\n")
    txt = Replace(txt, vbCr, "\n")
    txt = Replace(txt, vbLf, "\n")

    EscapeJSON = txt

End Function