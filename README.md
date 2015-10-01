VBEX
====

_VBA Extension Library_

Ease production of VBA code with the VBEX library of rich idiomatic containers and some functional programing 
capabilities to bring VBA into the new millenium. With VBEX you can:

  1. Use Class Constructors for immutable classes.
  1. Print meaningful debug methods that reveal a datastructures contents

        Console.PrintLine List.Create(1, 2, 3, 4) ' Note usage of class constructors
        List(1, 2, 3, 4)

  1. Create functional objects to use with higher order functions.  With those we have created some monadic classes _(List, Maybe, Try)_ that implement the traditonal `Map`, `FlatMap` or `Bind` methods.
  1. Access a growing library of Containers.
  1. Perform file-system operations.
      - These will later be replaced or enhanced with an object-oriented model
  1. Later there will be APIs for ADODB (SQL) and Windows Scripting Host


<!--
    Debug.Print Show(List.Create(1, 2, 3))
    List(1, 2, 3)

    Console.PrintLine xs
    List(1, a, Collection(&289234581))
    
    Console.PrintLine s
    SortedSet(1, 2, 3)
    
    Console.PrintLine d
    Dict(Parrot -> Dead, Spam -> Yum)

    Dim xs As List
    Set xs = List.Create(1,"a", New Collection)
    
    Dim s As SortedSet
    Set s = SortedSet.Create(1, 2, 2, 2, 3, 2, 1, 2, 1, 3)
    
    Dim d As Dict
    Set d = Dict.Create( _
        Assoc.Make("Parrot", "Dead"), _
        Assoc.Make("Spam", "Yum") _
    )

    Dim getRow As OnObject
    Set getRow = OnObject.Make("Row", vbGet)
    
    Dim offsetRow As Lambda
    Set offsetRow = Lambda.FromShort(" _ + 3 ")
    
    Dim tableRows As List
    Set tableRows = List.Copy(Selection.Rows)
    
    Dim rowIndexes As List
    Set rowIndexes = tableRows.Map(getRow.AndThen(offsetRow))
-->

Intro
-----

Install
-------

Once you acquire the source by either cloning this repo or downloading zip and extracting

  1. Run the _Make.ps1_ script to build _VBEXsrc.xlam_ and _VBEXtest.xlam_.
      - _VBEXtest.xlam_ contains unit-testing code and is only relevant to development.
  1. Reference _VBEXsrc.xlam_ in projects to use VBEX
      - From the VBE from the menu _tools >> References >> Browse_.
  1. Enable ["Programmatic access to Office VBA project"](https://support.microsoft.com/en-us/kb/282830)
      -  This is required for the Lambda class.

Usage
-----

VBEX is not a normal VBA library, before you start using you should understand the following aspects about VBEX.

  1. All public classes have a predeclared instance of that class called the "predeclared object".
       - The predeclared object has the same name as the class, _e.g._

               Dim xs As List ' word "List" as a type
               Set xs = List.Create(1, 2, 3) ' word "List" here is the predeclared object

       - All creatable classes are created from the predeclared object.
       - Predeclared objects of mutable classes can be mutated.
           + Dont do that.
  2. VBEX utilizes a broad interface system.  
       - Interfaces are minimally defined.
       - Many classes implement more than one interface (are polymorphic).
  3. Implementations of the _IApplicable_ interface allow methods and functions to be treated as objects.
       - All IApplicable objects are immutable.
       - The _Lambda_ class writes functions to a VBEX modules and allows you to execute that code.
           + Using the Lambda class will sometimes disable the debugger.
           + A lambda has no reference to environment in it was created.
               * `Lambda.FromShort("_ + x")` will always error even if `x` is in the current scope.
       - _OnArgs_ and _OnObject_ are complementary.
           + `OnArgs.Make(myObject, "method", vbMethod)` is `(x) => myObject.method(x)`
           + `OnObject.Make("method", vbMethod, myArgs)` is `(o) => o.method(myArgs)`
           + These are the _only_ applicable objects that have references to the current environment.

<!--
### Predeclared Objects

All public classes have a predeclared instace of that class.
The predeclared object is referenced using the name of the class and is
most often that instance is used for constructing new instaces of that object.

For example the class `List` has a predeclared instance:

    Debug.Print IsObject(List) ' outputs true
    Dim xs As List ' word "List" as a type
    Set xs = List.Create(1, 2, 3) ' word "List" as the predeclared object

The predeclared object is just another instance of a list and since our implementation of lists are mutable the predeclared object can be mutated just like any other list.
Please don't mutate predeclared objects.
It should be impossible but there currently isn't any elegant implementation preventing it.

### Interfaces

You will notice that some methods or functions require an interface instead of a specific class.
For example the `srch.MaxIndex` function has the signature `Public Function MaxIndex(ByVal iterable As IIterable) As Long`.
This means any IIterable

### Functional Objects
-->
