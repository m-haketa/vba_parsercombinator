Attribute VB_Name = "F"
Option Explicit

Function ParserOutput(Success As Boolean, Result As Variant, NewPosition As Long) As ParserOutput
  Dim PO As ParserOutput
  Set PO = New ParserOutput
  
  PO.Success = Success
  
  If IsObject(Result) Then
    Set PO.Result = Result
  Else
    PO.Result = Result
  End If
  
  PO.NewPosition = NewPosition
  
  Set ParserOutput = PO
End Function


Function Parse(Parser As iParser, Target As String, Optional Position As Long = 1) As ParserOutput
  Set Parse = Parser.Parse(Target, Position)
End Function


Function Token(Str As String) As Token
  Set Token = New Token
  Call Token.Init(Str)
End Function

Function Many(iParser As iParser) As Many
  Set Many = New Many
  Call Many.Init(iParser)
End Function


Function Choice(ParamArray iParsers()) As Choice
  Set Choice = New Choice
  
  Dim Parsers() As iParser
  ReDim Parsers(LBound(iParsers) To UBound(iParsers))
  
  Dim No As Long
  For No = LBound(iParsers) To UBound(iParsers)
    Set Parsers(No) = iParsers(No)
  Next
  
  Call Choice.Init(Parsers)
End Function

Function Seq(ParamArray iParsers()) As Seq
  Set Seq = New Seq
  
  Dim Parsers() As iParser
  ReDim Parsers(LBound(iParsers) To UBound(iParsers))
  
  Dim No As Long
  For No = LBound(iParsers) To UBound(iParsers)
    Set Parsers(No) = iParsers(No)
  Next
  
  Call Seq.Init(Parsers)
End Function

Function Opt(iParser As iParser) As Opt
  Set Opt = New Opt
  Call Opt.Init(iParser)
End Function

Function RegEx(Pattern As String, Optional IgnoreCase As Boolean = False, Optional RegGlobal As Boolean = False, Optional Multiline As Boolean = False) As RegEx
  Set RegEx = New RegEx
  
  If Left(Pattern, 1) <> "^" Then
    Pattern = "^(?:" & Pattern & ")"
  End If
  
  Call RegEx.Init(getRegExp(Pattern, IgnoreCase))
End Function


Function getRegExp(Pattern As String, Optional IgnoreCase As Boolean = False, Optional RegGlobal As Boolean = False, Optional Multiline As Boolean = False) As VBScript_RegExp_55.RegExp
  Set getRegExp = New VBScript_RegExp_55.RegExp
  
  getRegExp.Pattern = Pattern
  getRegExp.IgnoreCase = IgnoreCase
  getRegExp.Global = RegGlobal
  getRegExp.Multiline = Multiline

End Function


Function Map(iParser As iParser, iMapCallback As iMapCallback) As Map
  Set Map = New Map
  Call Map.Init(iParser, iMapCallback)
End Function


Function Char(Chars As String) As Char
  Set Char = New Char
  Call Char.Init(Chars)
End Function


'Function Lazy(Callback As iParser) As Lazy
'  Set Lazy = New Lazy
'  Call Lazy.Init(Callback)
'End Function

