Attribute VB_Name = "Test"
Option Explicit

Sub testFormula()
  Dim Number As iParser
  Set Number = F.Map(F.RegEx("[1-9][0-9]*|[0-9]"), New MapParseInt)

  Dim Operator As iParser
  Set Operator = F.Char("+-")
  
  Dim Parenthesis As Lazy
  Set Parenthesis = New Lazy
  
  Dim Atom As iParser
  Set Atom = F.Choice(Number, Parenthesis)
  
  Dim Expression As iParser
  Set Expression = F.Map(F.Seq(Atom, F.Many(F.Seq(Operator, Atom))), New MapFlattenSeqManySeq)
  
  Call Parenthesis.SetCallback(F.Map(F.Seq(F.Token("("), Expression, F.Token(")")), New MapDeleteParenthesis))
  
  Debug.Print F.Parse(Expression, "1+2+3+4")
  
  Debug.Print F.Parse(Expression, "1+(2-3)+4")

  Debug.Print F.Parse(Expression, "2-10+5")
  Debug.Print F.Parse(Expression, "10+(0-30)+4")
  Debug.Print F.Parse(Expression, "12")
  Debug.Print F.Parse(Expression, "1+1")
  
  
End Sub


Sub Test()
  
  Debug.Print "ParseHoge"
  Debug.Print F.Parse(New ParseHoge, "hoge")
  Debug.Print F.Parse(New ParseHoge, "ahoge", 2)
  Debug.Print F.Parse(New ParseHoge, "aaa")

  Debug.Print "Token"
  Debug.Print F.Parse(F.Token("foobar"), "foobar")
  Debug.Print F.Parse(F.Token("foobar"), "foobar", 2)

  Debug.Print "Many"
  Debug.Print F.Parse(F.Many(F.Token("hoge")), "hogehoge")
  Debug.Print F.Parse(F.Many(F.Token("hoge")), "")
  Debug.Print F.Parse(F.Many(F.Token("foobar")), "foo")

  Debug.Print "Many,Choice,Token"
  Dim PMany As iParser
  Set PMany = F.Many(F.Choice(F.Token("hoge"), F.Token("fuga")))
  Debug.Print F.Parse(PMany, "")
  Debug.Print F.Parse(PMany, "hogehoge")
  Debug.Print F.Parse(PMany, "fugahoge")
  Debug.Print F.Parse(PMany, "fugafoo")

  Debug.Print "Seq"
  Dim PSeq As iParser
  Set PSeq = F.Seq(F.Token("foo"), F.Choice(F.Token("bar"), F.Token("baz")))
  Debug.Print F.Parse(PSeq, "foobar")
  Debug.Print F.Parse(PSeq, "foobaz")
  Debug.Print F.Parse(PSeq, "foo")

  Debug.Print "Opt"
  Debug.Print F.Parse(F.Opt(F.Token("hoge")), "hoge")
  Debug.Print F.Parse(F.Opt(F.Token("fuga")), "hoge")

  Debug.Print "RegEx"
  Debug.Print F.Parse(F.RegEx("hoge"), "hoge")
  Debug.Print F.Parse(F.RegEx("([1-9][0-9]*)"), "2014")
  Debug.Print F.Parse(F.RegEx("([1-9][0-9]*)"), "01")
  
    
  Debug.Print "Lazy"
  Dim pLazy As Lazy     'Lazyå^Ç…ÇµÇƒÇ®Ç©Ç»Ç¢Ç∆SetCallbackÇ™åƒÇ—èoÇπÇ»Ç¢
  Set pLazy = New Lazy

  Dim pLazyBase As iParser
  Set pLazyBase = F.Opt(F.Seq(F.Token("hoge"), pLazy))

  'íxâÑï]âøÇµÇΩÇ¢èÍçáÇÕÅAå„Ç©ÇÁÉZÉbÉgÇ∑ÇÈ
  Call pLazy.SetCallback(pLazyBase)

  Debug.Print F.Parse(pLazyBase, "hoge")
  Debug.Print F.Parse(pLazyBase, "hogehoge")
  Debug.Print F.Parse(pLazyBase, "hogehogehoge")

  Debug.Print "Map"
  Debug.Print F.Parse(F.Map(F.Token("hello"), New MapParsed), "hello")
  Debug.Print F.Parse(F.Map(F.Token("hello"), New MapParsed), "foobar")
  
  Debug.Print "Char"
  Debug.Print F.Parse(F.Char("abcdef"), "a")
  Debug.Print F.Parse(F.Char("abcdef"), "b")
  Debug.Print F.Parse(F.Char("abcdef"), "g")
  Debug.Print F.Parse(F.Char("abcdef"), "")
  
  
  Debug.Print "Many-Seq"
  Debug.Print F.Parse(F.Seq(F.Char("123456789"), F.Many(F.Seq(F.Char("+"), F.Char("123456789")))), "1+2+3+4")
  


End Sub


Function CreateCollection(ParamArray Vals()) As Collection
  Dim Ret As Collection
  Set Ret = New Collection
  
  Dim Val
  For Each Val In Vals
    Ret.Add Val
  Next
  
  Set CreateCollection = Ret
End Function


