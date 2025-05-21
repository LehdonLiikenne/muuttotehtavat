VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTextParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'---------------------------------------------------------------------------------------
' Class: clsTextParser
' Description: Parses a string based on a specified delimiter.
'              Provides properties to access the first two parts and indicates if split occurred.
'---------------------------------------------------------------------------------------

Option Explicit

'--- Private Member Variables ---
Private m_sInputText As String      ' The original text to parse
Private m_sDelimiter As String      ' The delimiter character(s)
Private m_sPart1 As String          ' Stores the first part after parsing
Private m_sPart2 As String          ' Stores the second part after parsing
Private m_vParts As Variant         ' Stores all parts as an array (optional)
Private m_bIsSplit As Boolean       ' Flag indicating if the delimiter was found

'--- Public Properties ---

' InputText: The string to be parsed
Public Property Get InputText() As String
    InputText = m_sInputText
End Property
Public Property Let InputText(ByVal value As String)
    m_sInputText = value
    ' Reset parsed values when input changes
    ClearResults
End Property

' Delimiter: The character(s) to split by
Public Property Get Delimiter() As String
    Delimiter = m_sDelimiter
End Property
Public Property Let Delimiter(ByVal value As String)
    m_sDelimiter = value
    ClearResults ' Reset if delimiter changes
End Property

' Part1: Read-only property for the first part
Public Property Get Part1() As String
    Part1 = m_sPart1
End Property

' Part2: Read-only property for the second part
Public Property Get Part2() As String
    Part2 = m_sPart2
End Property

' Parts: Read-only property for all parts as an array
Public Property Get Parts() As Variant
    Parts = m_vParts
End Property

' IsSplit: Read-only property indicating if the split occurred
Public Property Get IsSplit() As Boolean
    IsSplit = m_bIsSplit
End Property

'--- Public Methods ---

' Parse: Performs the parsing operation based on current InputText and Delimiter
Public Sub Parse()
    Dim tempInput As String
    Dim delimiterPos As Integer

    ' Clear previous results before parsing
    ClearResults

    ' Ensure we have input and a delimiter
    If Len(m_sInputText) = 0 Or Len(m_sDelimiter) = 0 Then
        Exit Sub ' Nothing to parse or no delimiter set
    End If

    ' Trim input for consistency
    tempInput = Trim(m_sInputText)

    ' Check if the delimiter exists
    delimiterPos = InStr(tempInput, m_sDelimiter)

    If delimiterPos > 0 Then
        ' Delimiter found - Split the string
        m_bIsSplit = True
        m_vParts = Split(tempInput, m_sDelimiter)

        ' Assign Part1 (always the first element if split occurs)
        If UBound(m_vParts) >= 0 Then
            m_sPart1 = Trim(m_vParts(0))
        End If

        ' Assign Part2 (the second element, if it exists)
        If UBound(m_vParts) >= 1 Then
            m_sPart2 = Trim(m_vParts(1))
        End If
        ' Note: If there are more than two parts, they are available in the m_vParts array
        ' but Part1 and Part2 only hold the first two.

    Else
        ' Delimiter not found
        m_bIsSplit = False
        m_sPart1 = tempInput ' Assign the whole string to Part1
        m_sPart2 = ""        ' Part2 remains empty as no split occurred
        m_vParts = Array(tempInput) ' Parts array contains the single, unsplit string
    End If
End Sub

'--- Private Helper Methods ---

' ClearResults: Resets the internal state (parsed values)
Private Sub ClearResults()
    m_sPart1 = ""
    m_sPart2 = ""
    m_vParts = Empty ' Reset variant array
    m_bIsSplit = False
End Sub

'--- Standard Class Events ---

Private Sub Class_Initialize()
    ' Set default values if needed when the object is created
    ' For example: m_sDelimiter = ","
    ClearResults ' Ensure clean state on creation
End Sub

Private Sub Class_Terminate()
    ' Clean up resources if needed when the object is destroyed
    ' (Not usually necessary for simple string manipulation)
End Sub

