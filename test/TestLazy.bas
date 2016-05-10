Attribute VB_Name = "TestLazy"
'@TestModule
Option Explicit
Option Private Module
Private Assert As New Rubberduck.AssertClass

'@TestMethod
Public Sub LazyMakeTest()

    Dim op As OnArgs
    Set op = OnArgs.Make("Contains", VbMethod, SortedSet.Create(2))
    
    Dim lzy As Lazy
    Set lzy = Lazy.Create(op, 2)
    
    Assert.IsTrue lzy.IsDelayed
    Assert.IsTrue lzy
    Assert.IsFalse lzy.IsDelayed

End Sub
'
'@TestMethod
Public Sub LazyTestWithHashSet()

    Dim lz As Lazy
    Set lz = Lazy.Create(InternalDelegate.Make("BaseName"), "Hello\World")
    
    Dim hs As HashSet
    Set hs = HashSet.Create("World")
    
    Assert.IsTrue hs.Contains(lz) ' WILL FAIL, but shouldn't
    Assert.IsTrue hs.Contains(lz.Evaluate) ' doesn't fail
    Assert.IsTrue hs.Contains(lz) ' still does...
    
End Sub
'
' Various problems with Map...
' ----------------------------
'
'@TestMethod
Public Sub LazyLambdaMapTest()

    Dim lazyFour As Lazy
    Set lazyFour = Lazy.Create(Lambda.FromShort("_ + _"), 2, 2)
    
    Dim multTen As Lambda
    Set multTen = Lambda.FromShort("10 * _ ")
    
    LazyMapTest lazyFour, multTen, 40
    
End Sub
'@TestMethod
Public Sub LazyMapTestWithInternalDelegate()

    Dim root As Lazy
    Set root = Lazy.Create(InternalDelegate.Make("RootName"), "C:\Some\Path\Yo.txt")

    Dim baser As Applicable
    Set baser = InternalDelegate.Make("BaseName")

    LazyMapTest root, baser, "Path"

End Sub
Private Sub LazyMapTest(ByVal initialLazy As Lazy, ByVal toMapWith As Applicable, ByVal mappedResult)

    Debug.Assert initialLazy.IsDelayed

    Dim mapped As Lazy
    Set mapped = initialLazy.Map(toMapWith)

    Assert.IsTrue initialLazy.IsDelayed
    Assert.IsTrue mapped.IsDelayed
    
    Assert.IsTrue Equals(mapped.Evaluate, mappedResult)
    Assert.IsTrue initialLazy.IsEvaluated
    
End Sub
