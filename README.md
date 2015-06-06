VBEX
====

VBA Extension Library

Ease production of VBA code with the VBEX library of rich idiomatic containers and some functional programing 
capabilities. With VBEX you can:

Use in-lined constructors:

    Dim xs As List
    Set xs = List.Create(1,"a", New Collection)
    
    Dim s As SortedSet
    Set s = SortedSet.Create(1, 2, 2, 2, 3, 2, 1, 2, 1, 3)
    
    Dim d As Dict
    Set d = Dict.Create( _
        Assoc.Make("Parrot", "Dead"), _
        Assoc.Make("Spam", "Yum") _
    )

Print meaningful debug messages

    Debug.Print Show(List.Create(1, 2, 3))
    List(1, 2, 3)

    Console.PrintLine xs
    List(1, a, Collection(&289234581))
    
    Console.PrintLine s
    SortedSet(1, 2, 3)
    
    Console.PrintLine d
    Dict(Parrot -> Dead, Spam -> Yum)

Create functional objects

    Dim getRow As OnObject
    Set getRow = OnObject.Make("Row", vbGet)
    
    Dim offsetRow As Lambda
    Set offsetRow = Lambda.FromShort(" _ + 3 ")
    
    Dim tableRows As List
    Set tableRows = List.Copy(Selection.Rows)
    
    Dim rowIndexes As List
    Set rowIndexes = tableRows.Map(getRow.AndThen(offsetRow))

