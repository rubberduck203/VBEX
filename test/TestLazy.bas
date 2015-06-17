Attribute VB_Name = "TestLazy"
'@TestModule
Option Explicit
Option Private Module
Private Assert As New Rubberduck.AssertClass


'@TestMethod
Public Sub LazyMakeTest()

    Dim op As OnArgs
    Set op = OnArgs.Make("Contains", VbMethod, SortedSet.Create(1, 2, 3))
    
    Dim lzy As Lazy
    Set lzy = Lazy.Make(ByName.Create(op, 2))
    
    Assert.IsTrue lzy.IsDelayed
    Assert.IsTrue lzy
    Assert.IsFalse lzy.IsDelayed

End Sub
'@TestMethod
Public Sub LazyMonadicTest()
' lazy is monadic?
' set y = x.Map(op) evaluates x but not y
' set y = x.Bind(op) evaluates neither x or y
    
    Dim cont As OnArgs
    Set cont = OnArgs.Make("Contains", VbMethod, SortedSet.Create(1, 2, 3))
    
    Dim x As Lazy
    Set x = Lazy.Make(ByName.Create(cont, 2))
    
    Assert.IsTrue x.IsDelayed
    
    Dim negate As Lambda
    Set negate = Lambda.FromShort("Not _ ")
    
    Dim y As Lazy
    Set y = x.Bind(negate)
    
    Assert.IsTrue x.IsDelayed
    Assert.IsTrue y.IsDelayed
    
    Dim z As Lazy
    Set z = x.Map(negate)
    
    Assert.IsTrue x.IsEvaluated
    Assert.IsTrue y.IsDelayed
    Assert.IsTrue z.IsDelayed
    
    
End Sub

