Attribute VB_Name = "Test"
Option Explicit

Sub Test()
  
  Debug.Print F.Parse(New ParseHoge, "hoge", 1).toString
  Debug.Print F.Parse(New ParseHoge, "ahoge", 2).toString
  Debug.Print F.Parse(New ParseHoge, "aaa", 1).toString

  Debug.Print F.Parse(F.Token("foobar"), "foobar", 1).toString
  Debug.Print F.Parse(F.Token("foobar"), "foobar", 2).toString

  Debug.Print F.Parse(F.Many(F.Token("hoge")), "hogehoge", 1).toString
  Debug.Print F.Parse(F.Many(F.Token("hoge")), "", 1).toString
  Debug.Print F.Parse(F.Many(F.Token("foobar")), "foo", 1).toString

  Dim PMany As iParser
  Set PMany = F.Many(F.Choice(F.Token("hoge"), F.Token("fuga")))
  Debug.Print F.Parse(PMany, "", 1).toString
  Debug.Print F.Parse(PMany, "hogehoge", 1).toString
  Debug.Print F.Parse(PMany, "fugahoge", 1).toString
  Debug.Print F.Parse(PMany, "fugafoo", 1).toString

  Dim PSeq As iParser
  Set PSeq = F.Seq(F.Token("foo"), F.Choice(F.Token("bar"), F.Token("baz")))
  Debug.Print F.Parse(PSeq, "foobar", 1).toString
  Debug.Print F.Parse(PSeq, "foobaz", 1).toString
  Debug.Print F.Parse(PSeq, "foo", 1).toString
  
  Debug.Print F.Parse(F.Opt(F.Token("hoge")), "hoge", 1).toString
  Debug.Print F.Parse(F.Opt(F.Token("fuga")), "hoge", 1).toString

  Debug.Print F.Parse(F.RegEx("hoge"), "hoge", 1).toString
  Debug.Print F.Parse(F.RegEx("([1-9][0-9]*)"), "2014", 1).toString
  Debug.Print F.Parse(F.RegEx("([1-9][0-9]*)"), "01", 1).toString


End Sub



