Attribute VB_Name = "VB6"
Option Explicit

'This is placed here because I have VB5 and I need VB6 functions
'You cant just delete this, because:
'1. I call some functions by referencing the module
'2. My Split and Join functions are KUSTOMIZED a bit :-D

'Visit VBSpeed for more functions like this : http://www.xbeat.net/vbspeed/

Public Function Replace(Expression As String, sOld As String, sNew As String, Optional ByVal Start As Long = 1, Optional ByVal Count As Long = -1, Optional Compare As VbCompareMethod = vbBinaryCompare) As String
' by Donald, donald@xbeat.net, 20001201 (rev 002)
' partly based on Replace05 by Jost Schwider

  Dim cntReplace As Long
  Dim lenOld As Long
  Dim lenNew As Long
  Dim lenCopy As Long
  Dim posSource As Long
  Dim posTarget As Long
  
  lenOld = Len(sOld)
  If lenOld Then
    
    ' get first match
    Start = InStr(Start, Expression, sOld, Compare)
    If Start Then
    
      ' set Count to "+inf" to get them all
      If Count < 0 Then Count = &H7FFFFFFF
      
      lenNew = Len(sNew)
      If lenOld = lenNew Then
        ' easy: simply overwrite Old with New
        Replace = Expression
        For Count = 1 To Count
          ' replace
          Mid$(Replace, Start) = sNew
          ' find next
          Start = InStr(Start + lenOld, Expression, sOld, Compare)
          If Start = 0 Then Exit Function
        Next
      
      Else
        ' not so easy: gotta rebuild target from scratch
        '
        ' determine the number of actual replacements
        ' and calculate exact space needed for returned string
        posSource = Start
        For Count = 1 To Count
          cntReplace = cntReplace + 1
          ' find next
          posSource = InStr(posSource + lenOld, Expression, sOld, Compare)
          If posSource = 0 Then Exit For
        Next
        ' allocate space
        Replace = Space$(Len(Expression) + cntReplace * (lenNew - lenOld))
        '
        ' do the rebuild
        posSource = 1
        posTarget = 1
        For cntReplace = 1 To cntReplace
          lenCopy = Start - posSource
          ' insert source
          If lenCopy Then
            Mid$(Replace, posTarget) = Mid$(Expression, posSource, lenCopy)
          End If
          ' insert new
          If lenNew Then
            Mid$(Replace, posTarget + lenCopy) = sNew
          End If
          ' next pos in target
          posTarget = posTarget + lenCopy + lenNew
          ' find next
          posSource = Start + lenOld
          Start = InStr(posSource, Expression, sOld, Compare)
        Next
        ' insert source remainder
        If posTarget <= Len(Replace) Then
          Mid$(Replace, posTarget) = Mid$(Expression, posSource)
        End If
      End If
    
    Else
      ' no match
      Replace = Expression
    End If
  
  Else
    ' find string is zero-length
    Replace = Expression
  End If
  
  DoEvents
  
End Function

Public Function Join(sArray() As String, Optional lLo As Long = -1, Optional lHi As Long = -1, Optional sDelimiter As String = " ") As String
 
   ' by G.Beckmann   eMail: G.Beckmann@NikoCity.de
   ' modified by Keith, kmatzen@ispchannel.com
    
   Dim lNdx      As Long
   Dim lJoinLen  As Long
   Dim lCurPos   As Long
   Dim lDelimLen As Long
   Dim lLastStr  As Long
   
   If lLo = -1 Then lLo = LBound(sArray)
   If lHi = -1 Then lHi = UBound(sArray)
   lDelimLen = Len(sDelimiter)
   
   '/ Calculate the size of the new string
   lJoinLen = (lHi - lLo) * lDelimLen
   For lNdx = lHi To lLo Step -1
      If Len(sArray(lNdx)) > 0 Then
         lJoinLen = lJoinLen + Len(sArray(lNdx))
         lLastStr = lNdx          'Position of last non-empty string
         Exit For
      End If
   Next lNdx
   For lNdx = lLo To lLastStr - 1
      lJoinLen = lJoinLen + Len(sArray(lNdx))
   Next lNdx
 
   '/ Fill the new string
   If lJoinLen > 0 Then
      Join = Space$(lJoinLen)
            
      Mid$(Join, 1) = sArray(lLo)
      lCurPos = Len(sArray(lLo)) + 1
      If lDelimLen > 0 Then
         For lNdx = lLo + 1 To lLastStr
            Mid$(Join, lCurPos, lDelimLen) = sDelimiter
            lCurPos = lCurPos + lDelimLen
            
            Mid$(Join, lCurPos, lJoinLen) = sArray(lNdx)
            lCurPos = lCurPos + Len(sArray(lNdx))
         Next lNdx
         For lNdx = lCurPos To lJoinLen Step lDelimLen
            Mid$(Join, lNdx, lDelimLen) = sDelimiter
         Next lNdx
      Else
         For lNdx = lLo + 1 To lLastStr
            Mid$(Join, lCurPos, lJoinLen) = sArray(lNdx)
            lCurPos = lCurPos + Len(sArray(lNdx))
         Next lNdx
      End If
   End If
 
End Function

Public Function Filter(sSourceArray() As String, sMatch As String, sTargetArray() As String, Optional bInclude As Boolean = True, Optional lCompare As VbCompareMethod = vbBinaryCompare) As Long
 
' by Donald, donald@xbeat.net, 20000918
' Modified by Keith, kmatzen@ispchannel.com
' returns Ubound(sTargetArray), or -1 if sTargetArray is not bound (empty array)
    
   Dim lNdx      As Long
   Dim lLo       As Long
   Dim lHi       As Long
   Dim lLenMatch As Long
   
   lLenMatch = Len(sMatch)
   lLo = LBound(sSourceArray)
   lHi = UBound(sSourceArray)
   ReDim sTargetArray(lHi - lLo) 'make maximal space
   
   Filter = -1
   
   If lLenMatch Then
      If bInclude Then              'Need a match
         For lNdx = lLo To lHi
            If Len(sSourceArray(lNdx)) >= lLenMatch Then
               If InStr(1, sSourceArray(lNdx), sMatch, lCompare) Then
                  Filter = Filter + 1
                  sTargetArray(Filter) = sSourceArray(lNdx)
               End If
            End If
         Next
      Else                          'Need a mismatch
         For lNdx = lLo To lHi
            Select Case Len(sSourceArray(lNdx))
               Case Is < lLenMatch 'Can't match
                  Filter = Filter + 1
                  sTargetArray(Filter) = sSourceArray(lNdx)
               Case Else
                  If InStr(1, sSourceArray(lNdx), sMatch, lCompare) = 0 Then
                     Filter = Filter + 1
                     sTargetArray(Filter) = sSourceArray(lNdx)
                  End If
            End Select
         Next
      End If
   ElseIf bInclude Then             'Include all
      For lNdx = lLo To lHi
         Filter = Filter + 1
         sTargetArray(Filter) = sSourceArray(lNdx)
      Next
   End If
   
   ' erase or shrink
   If Filter = -1 Then
      Erase sTargetArray
   Else
      ReDim Preserve sTargetArray(Filter)
   End If
    
End Function

Public Function InStrRev(sCheck As String, sMatch As String, Optional ByVal Start As Long = -1, Optional Compare As VbCompareMethod = vbBinaryCompare) As Long
' by Donald, donald@xbeat.net, 20000923
' strategy: search right to left until found
  Dim lenSearchFor As Long
  Dim posFound As Long
  Dim posSearch As Long
 
  lenSearchFor = Len(sMatch)
  
  If lenSearchFor Then
    If Start <= 0 Then
      Start = Len(sCheck)
    End If
    For posSearch = Start - lenSearchFor + 1 To 1 Step -1
      posFound = InStr(posSearch, sCheck, sMatch, Compare)
      If posFound Then
        If (posFound + lenSearchFor - 1) <= Start Then
          InStrRev = posFound
          Exit Function
        End If
      End If
    Next
  Else
    ' as VB6 InStrRev
    If Start <= Len(sCheck) Then
      InStrRev = Start
    End If
  End If
  
End Function

Public Sub Split(ByVal Expression$, ResultSplit$(), Optional Delimiter$ = " ")
' by G.Beckmann, G.Beckmann@NikoCity.de
 
    Dim c&, iLen&, iLast&, iCur&
    
    iLen = Len(Delimiter)
    
    If iLen Then
        
        '/ count delimiters
        iCur = InStr(Expression, Delimiter)
        Do While iCur
            iCur = InStr(iCur + iLen, Expression, Delimiter)
            c = c + 1
        Loop
        
        '/ initalization
        ReDim Preserve ResultSplit(0 To c)
        c = 0: iLast = 1
        
        '/ search again...
        iCur = InStr(Expression, Delimiter)
        Do While iCur
            ResultSplit(c) = Mid$(Expression, iLast, iCur - iLast)
            iLast = iCur + iLen
            iCur = InStr(iLast, Expression, Delimiter)
            c = c + 1
        Loop
        ResultSplit(c) = Mid$(Expression, iLast)
        
    Else
        ReDim Preserve ResultSplit(0 To 0)
        ResultSplit(0) = Expression
    End If
 
End Sub

Public Function InStrCount(Expression As String, Find As String, Optional Start As Long = 1, Optional Compare As VbCompareMethod = vbBinaryCompare) As Long

' by Peter Nierop, pnierop.pnc@inter.nl.net, 20001226

  Dim aOrg() As Byte, lMaxOrg&, lCurOrg&
  Dim aFind() As Byte, lMaxFind&, lCurFind&, lFind&, lComp&

  Dim lFindCount&

  '=========== check op input ========================================
  lMaxOrg = Len(Expression)
  lMaxFind = Len(Find)


  ' preload the first character to find
  If lMaxOrg = 0 Or lMaxFind = 0 Or Start > lMaxOrg Then
    InStrCount = 0
    Exit Function
  End If

  If Start < 1 Then
    Err.Raise 5, "InStrCount Function", "Start can't be smaller than 1"
    Exit Function
  ElseIf Start > 1 Then
    lCurOrg = Start * 2 - 2
  End If



  '=========== prepare buffers =======================================
  aOrg = Expression
  lMaxOrg = UBound(aOrg)



  '==========  With one character to find -> shorter loop =====================
  If lMaxFind = 1 Then

    lFind = Asc(Find)
    If Compare = vbBinaryCompare Then
      For lCurOrg = lCurOrg To lMaxOrg Step 2

        If lFind = aOrg(lCurOrg) Then
          lFindCount = lFindCount + 1
        End If

      Next

    Else
      lComp = &HDF   'to uppercase
      lFind = lFind And lComp

      For lCurOrg = lCurOrg To lMaxOrg Step 2

        If lFind = (aOrg(lCurOrg) And lComp) Then
          lFindCount = lFindCount + 1
        End If

      Next
    End If

  Else
  '============ Longer loop if multiple characters to find ======================

    aFind = Find
    lMaxFind = UBound(aFind)
    lFind = aFind(0)

    If Compare = vbBinaryCompare Then
      For lCurOrg = lCurOrg To lMaxOrg Step 2

        If lFind = aOrg(lCurOrg) Then

          lCurFind = lCurFind + 2
          ' if no more characters to test -> match with string happened
          If lCurFind >= lMaxFind Then
            lFindCount = lFindCount + 1
            lCurFind = 0  'and start over
          End If
          ' now load next character from string to find
          lFind = aFind(lCurFind)

        Else
          ' no match so back to next character after first match
          lCurOrg = lCurOrg - lCurFind
          lCurFind = 0
          lFind = aFind(0)
        End If

      Next

    Else

      ' modify find array to uppercase
      For lCurFind = 0 To lMaxFind Step 2
        aFind(lCurFind) = aFind(lCurFind) And &HDF
      Next
      lCurFind = 0
      lFind = aFind(0)
      lComp = &HDF

      For lCurOrg = lCurOrg To lMaxOrg Step 2

        If lFind = (aOrg(lCurOrg) And lComp) Then

          lCurFind = lCurFind + 2
          ' if no more characters to test -> match with string happened
          If lCurFind >= lMaxFind Then
            lFindCount = lFindCount + 1
            lCurFind = 0  'and start over
          End If
          ' now load next character from string to find
          lFind = aFind(lCurFind)

        Else
          ' no match so back to next character after first match
          lCurOrg = lCurOrg - lCurFind
          lCurFind = 0
          lFind = aFind(0)
        End If

      Next

    End If

  End If

  InStrCount = lFindCount
End Function



