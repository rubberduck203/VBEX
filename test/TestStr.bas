Attribute VB_Name = "TestStr"
'@TestModule
Option Explicit
Option Private Module
Private Assert As New Rubberduck.AssertClass
'
'
' Constructors
' ------------
'
'@TestMethod
Public Sub TestStrJoin()

    Dim s As Str
    Set s = Str.Join(List.Create("Hello", "World"), ", ")
    
    Assert.AreEqual "Hello, World", s.Show

End Sub
'@TestMethod
Public Sub TestStrMake()

    Dim s As Str
    Set s = Str.Make("Hello, World")
    
    Assert.AreEqual "Hello, World", s.Show

End Sub
'@TestMethod
Public Sub TestStrRepeat()

    Dim s As Str
    Set s = Str.Repeat("Spam", 3)
    
    Assert.AreEqual "SpamSpamSpam", s.Show

End Sub
'@TestMethod
Public Sub TestStrFormat()

    Dim s As Str
    Set s = Str.Format("{0}, {2}, {1}", "a", 2, 4.5)
    
    Assert.AreEqual "a, 4.5, 2", s.Show

End Sub
'@TestMethod
Public Sub TestStrEscape()

    Dim s As Str
    Set s = Str.Escape("Phil's parrot said ""I'm not dead""")
    
    Assert.AreEqual "Phil`s` parrot` said` `""I`'m` not` dead`""", s.Show

End Sub
'@TestMethod
Public Sub StrIterable()

    BatteryIterable.Battery Str.Make("Hello, World")

End Sub

