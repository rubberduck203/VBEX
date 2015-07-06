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
    Set lzy = Lazy.Create(op, 2)
    
    Assert.IsTrue lzy.IsDelayed
    Assert.IsTrue lzy
    Assert.IsFalse lzy.IsDelayed

End Sub
'@TestMethod
Public Sub LazyMapTest()

    Dim lazyFour As Lazy
    Set lazyFour = Lazy.Create(Lambda.FromShort("_ + _"), 2, 2)
    Assert.IsTrue lazyFour.IsDelayed
    Debug.Assert lazyFour.IsDelayed
    
    Dim lazyFourty As Lazy
    Set lazyFourty = lazyFour.Map(Lambda.FromShort("10 * _ "))
    Assert.IsTrue lazyFourty.IsDelayed
    Assert.IsTrue lazyFour.IsDelayed
    Debug.Assert lazyFourty.IsDelayed
    Debug.Assert lazyFour.IsDelayed
    
    Dim forty As Integer
    forty = lazyFourty.Evaluate
    Assert.AreEqual 40, forty, "(2 + 2) * 10 == 40"
    Debug.Assert forty = 40
    
    Assert.IsTrue lazyFourty.IsEvaluated ' Obviously I just called evaluate
    Assert.IsTrue lazyFour.IsEvaluated ' Not obvious but optimal.
    Debug.Assert lazyFourty.IsEvaluated
    Debug.Assert lazyFour.IsEvaluated
    
End Sub
'@TestMethod
Public Sub LazyMapTestWithOnArgs()

    Dim is123 As OnArgs
    Set is123 = OnArgs.Make("Contains", VbMethod, HashSet.Create(1, 2, 3))

    Dim lazyYes As Lazy
    Set lazyYes = Lazy.Create(is123, 2)
    Assert.IsTrue lazyYes.IsDelayed
    
    Dim somer As OnArgs
    Set somer = OnArgs.Make("Some", VbMethod, Maybe)
    
    Dim maybeYes As Lazy
    Set maybeYes = lazyYes.Map(somer)
    Assert.IsTrue lazyYes.IsDelayed
    Assert.IsTrue maybeYes.IsDelayed
    
End Sub
