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
Public Sub LazyMapTest()

    Dim lazyFour As Lazy
    Set lazyFour = Lazy.Make(ByName.Create(Lambda.FromShort("_ + _"), 2, 2)) ' 2 + 2
    Assert.IsTrue lazyFour.IsDelayed
    
    Dim lazyFourty As Lazy
    Set lazyFourty = lazyFour.Map(Lambda.FromShort("10 * _ "))
    Assert.IsTrue lazyFourty.IsDelayed
    Assert.IsTrue lazyFour.IsDelayed
    
    Dim forty As Integer
    forty = lazyFourty.Evaluate
    Assert.AreEqual 40, forty, "(2 + 2) * 10 == 40"
    
    Assert.IsTrue lazyFourty.IsEvaluated ' Obviously I just called evaluate
    Assert.IsTrue lazyFour.IsEvaluated ' Not obvious but optimal.
    
End Sub

