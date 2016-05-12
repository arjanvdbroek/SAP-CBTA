Function ReadINIString (IniSection, IniKey, IniFile, DefaultValue, Options)
  On Error Resume Next
  EventComponentBegin()
  If ConditionsManager().CheckConditions() Then
    ReportDebugLog "CUSTOM Function - ReadINIString " & _ 
                 vbCrLf & " - IniSection " & IniSection & _
                 vbCrLf & " - IniKey " & IniKey & _
                 vbCrLf & " - IniFile " & IniFile & _
                 vbCrLf & " - DefaultValue " & DefaultValue & _
                 vbCrLf & " - Options " & Options                                                                     
    ReadINIString = ReadINIString_Impl (IniSection, IniKey, IniFile, DefaultValue)
  End If
  EventComponentEnd()
End Function

Function ReadINIString_Impl(Section, KeyName, FileName, DefaultValue)   'internal
  Dim INIContents, PosSection, PosEndSection, sContents, Value, Found, bSuccess
  bSuccess = True
  'Get contents of the INI file As a string
  INIContents = GetFile(FileName)

  'Find section
  PosSection = InStr(1, INIContents, "[" & Section & "]", vbTextCompare)
  If PosSection>0 Then
    'Section exists. Find end of section
    PosEndSection = InStr(PosSection, INIContents, vbCrLf & "[")
    '?Is this last section?
    If PosEndSection = 0 Then PosEndSection = Len(INIContents)+1
    
    'Separate section contents
    sContents = Mid(INIContents, PosSection, PosEndSection - PosSection)

    If InStr(1, sContents, vbCrLf & KeyName & "=", vbTextCompare)>0 Then
      Found = True
      'Separate value of a key.
      Value = SeparateField(sContents, vbCrLf & KeyName & "=", vbCrLf)
    End If
  End If
  If isempty(Found) Then 
    If IsNull(DefaultValue) Then
      bSuccess = False
    Else
      Value = DefaultValue
    End If
  End If
  If bSuccess then
    ReadINIString_Impl = Value
    ReportLog "DONE", "ReadINIString", "Setting Read = [" & Section & "] - " & KeyName & " = " & Value
  Else
    ReportLog "FAILED", "ReadINIString", "Setting or file not found"
  End If
End Function

Function WriteINIString (IniSection, IniKey, IniFile, IniValue, Options)
  On Error Resume Next
  EventComponentBegin()
  If ConditionsManager().CheckConditions() Then
    ReportDebugLog "CUSTOM Function - WriteINIString " & _
                 vbCrLf & " - IniSection " & IniSection & _
                 vbCrLf & " - IniKey " & IniKey & _
                 vbCrLf & " - IniFile " & IniFile & _
                 vbCrLf & " - IniValue " & IniValue & _
                 vbCrLf & " - Options " & Options                                                                     
    WriteINIString = WriteINIString_Impl (IniSection, IniKey, IniFile, IniValue)
  End If
  EventComponentEnd()
End Function

Function WriteINIString_Impl(Section, KeyName, FileName, Value)   'internal
  Dim INIContents, PosSection, PosEndSection
  
  'Get contents of the INI file As a string
  INIContents = GetFile(FileName)

  'Find section
  PosSection = InStr(1, INIContents, "[" & Section & "]", vbTextCompare)
  If PosSection>0 Then
    'Section exists. Find end of section
    PosEndSection = InStr(PosSection, INIContents, vbCrLf & "[")
    '?Is this last section?
    If PosEndSection = 0 Then PosEndSection = Len(INIContents)+1
    
    'Separate section contents
    Dim OldsContents, NewsContents, Line
    Dim sKeyName, Found
    OldsContents = Mid(INIContents, PosSection, PosEndSection - PosSection)
    OldsContents = split(OldsContents, vbCrLf)

    'Temp variable To find a Key
    sKeyName = LCase(KeyName & "=")

    'Enumerate section lines
    For Each Line In OldsContents
      If LCase(Left(Line, Len(sKeyName))) = sKeyName Then
        Line = KeyName & "=" & Value
        Found = True
      End If
      NewsContents = NewsContents & Line & vbCrLf
    Next

    If isempty(Found) Then
      'key Not found - add it at the end of section
      NewsContents = NewsContents & KeyName & "=" & Value
    Else
      'remove last vbCrLf - the vbCrLf is at PosEndSection
      NewsContents = Left(NewsContents, Len(NewsContents) - 2)
    End If

    'Combine pre-section, new section And post-section data.
    INIContents = Left(INIContents, PosSection-1) & _
      NewsContents & Mid(INIContents, PosEndSection)
  else'if PosSection>0 Then
    'Section Not found. Add section data at the end of file contents.
    If Right(INIContents, 2) <> vbCrLf And Len(INIContents)>0 Then 
      INIContents = INIContents & vbCrLf 
    End If
    INIContents = INIContents & "[" & Section & "]" & vbCrLf & _
      KeyName & "=" & Value
  end if'if PosSection>0 Then
  WriteFile FileName, INIContents
  ReportLog "DONE", "WriteINIString", "Setting Written = [" & Section & "] - " & KeyName & " = " & Value
  WriteINIString_Impl = True
End Function

Function DeleteINIString (IniSection, IniKey, IniFile, Par4, Options)
  On Error Resume Next
  EventComponentBegin()
  If ConditionsManager().CheckConditions() Then
    ReportDebugLog "CUSTOM Function - DeleteINIString " & _
                 vbCrLf & " - IniSection " & IniSection & _
                 vbCrLf & " - IniKey " & IniKey & _
                 vbCrLf & " - IniFile " & IniFile & _
                 vbCrLf & " - Par4 " & Par4 & _
                 vbCrLf & " - Options " & Options                                                                     
    DeleteINIString = DeleteINIString_Impl (IniSection, IniKey, IniFile)
  End If
  EventComponentEnd()
End Function

Function DeleteINIString_Impl(Section, KeyName, FileName)   'internal
  Dim INIContents, PosSection, PosEndSection
  
  'Get contents of the INI file As a string
  INIContents = GetFile(FileName)

  'Find section
  PosSection = InStr(1, INIContents, "[" & Section & "]", vbTextCompare)
  If PosSection>0 Then
    'Section exists. Find end of section
    PosEndSection = InStr(PosSection, INIContents, vbCrLf & "[")
    '?Is this last section?
    If PosEndSection = 0 Then PosEndSection = Len(INIContents)+1
    
    'Separate section contents
    Dim OldsContents, NewsContents, Line
    Dim sKeyName, Found
    OldsContents = Mid(INIContents, PosSection, PosEndSection - PosSection)
    OldsContents = split(OldsContents, vbCrLf)

    'Temp variable To find a Key
    sKeyName = LCase(KeyName & "=")

    'Enumerate section lines
    For Each Line In OldsContents
      If LCase(Left(Line, Len(sKeyName))) = sKeyName Then
        Found = True
      Else
        NewsContents = NewsContents & Line & vbCrLf
      End If
    Next

    If isempty(Found) Then
      'key Not found
      ReportLog "INFO", "DeleteINIString", "Setting not found, [" & Section & "] - " & KeyName
    Else
      'remove last vbCrLf - the vbCrLf is at PosEndSection
      NewsContents = Left(NewsContents, Len(NewsContents) - 2)
    End If

    'Combine pre-section, new section And post-section data.
    INIContents = Left(INIContents, PosSection-1) & _
      NewsContents & Mid(INIContents, PosEndSection)
  else'if PosSection>0 Then
    ReportLog "INFO", "DeleteINIString", "Setting not found = [" & Section & "] - " & KeyName
  end if'if PosSection>0 Then
  WriteFile FileName, INIContents
  ReportLog "DONE", "DeleteINIString", "Setting Removed = [" & Section & "] - " & KeyName
  DeleteINIString_Impl = True
End Function


Function DeleteINISection (IniSection, IniFile, Par3, Par4, Options)
  On Error Resume Next
  EventComponentBegin()
  If ConditionsManager().CheckConditions() Then
    ReportDebugLog "CUSTOM Function - DeleteINISection " & _
                 vbCrLf & " - IniSection " & IniSection & _
                 vbCrLf & " - IniFile " & IniFile & _
                 vbCrLf & " - Par3 " & Par3 & _
                 vbCrLf & " - Par4 " & Par4 & _
                 vbCrLf & " - Options " & Options                                                                     
    DeleteINISection = DeleteINISection_Impl (IniSection, IniFile)
  End If
  EventComponentEnd()
End Function


Function DeleteINISection_Impl(Section, FileName)   'internal
  Dim INIContents, PosSection, PosEndSection
  
  'Get contents of the INI file As a string
  INIContents = GetFile(FileName)

  'Find section
  PosSection = InStr(1, INIContents, "[" & Section & "]", vbTextCompare)
  If PosSection>0 Then
    'Section exists. Find end of section
    PosEndSection = InStr(PosSection, INIContents, vbCrLf & "[")
    '?Is this last section?
    If PosEndSection = 0 Then PosEndSection = Len(INIContents)+1
   
    'Combine pre-section, And post-section data.
    INIContents = Left(INIContents, PosSection-1) & Mid(INIContents, PosEndSection)
  else'if PosSection>0 Then
    ReportLog "INFO", "DeleteINISection", "Section not found = [" & Section & "]"
  end if'if PosSection>0 Then
  WriteFile FileName, INIContents
  ReportLog "DONE", "DeleteINISection", "Deleted Section [" & Section & "]"
  DeleteINISection_Impl = True
End Function

'Separates one field between sStart And sEnd
Function SeparateField(ByVal sFrom, ByVal sStart, ByVal sEnd)   'internal
  Dim PosB: PosB = InStr(1, sFrom, sStart, 1)
  If PosB > 0 Then
    PosB = PosB + Len(sStart)
    Dim PosE: PosE = InStr(PosB, sFrom, sEnd, 1)
    If PosE = 0 Then PosE = InStr(PosB, sFrom, vbCrLf, 1)
    If PosE = 0 Then PosE = Len(sFrom) + 1
    SeparateField = Mid(sFrom, PosB, PosE - PosB)
  End If
End Function


'File functions
Function GetFile(ByVal FileName)   'internal
  Dim FS: Set FS = CreateObject("Scripting.FileSystemObject")
  'Go To windows folder If full path Not specified.
  If InStr(FileName, ":\") = 0 And Left (FileName,2)<>"\\" Then 
    FileName = FS.GetSpecialFolder(0) & "\" & FileName
  End If
  On Error Resume Next

  GetFile = FS.OpenTextFile(FileName).ReadAll
End Function

Function WriteFile(ByVal FileName, ByVal Contents)   'internal
  
  Dim FS: Set FS = CreateObject("Scripting.FileSystemObject")
  'On Error Resume Next

  'Go To windows folder If full path Not specified.
  If InStr(FileName, ":\") = 0 And Left (FileName,2)<>"\\" Then 
    FileName = FS.GetSpecialFolder(0) & "\" & FileName
  End If

  Dim OutStream: Set OutStream = FS.OpenTextFile(FileName, 2, True)
  OutStream.Write Contents
End Function

'==================================================================
If Not IsEmpty(Environment) Or CBASE_BOOTSTRAP Then
  'This line prevents from loading the library twice.
  ExecutionContext().DeclareLibrary "INI_File.vbs"
End If
