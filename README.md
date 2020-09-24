<div align="center">

## vbaObjSetAddref \- The native undocumented visual basic run\-time APIs


</div>

### Description

For more than six years, i've been using visual basic, i've discovered many secrets, more than you can imagine..

now i decided to giveaway some of what i found out..

If you look in visual basic's run-time dll using depends, you'll see many functions, all undocumented, and from the names, it makes you want to grab them and use them.

Some of the functions include:

__vbaObjSetAddref

__vbaObjSet

They look so useful, but yet impossible to use, without knowing the correct declaration statement for each function.. (including byval/byref , data types, and the number of the parameters)

if searched google, i've never found any result, only ordinals dump listing, cracking sites &amp; such.

i've tried and tried, and finally, i figured some of them out..

----

now to the description of those functions:

__vbaObjSetAddref

__vbaObjSet

Both of these functions can be passed 2 parameters

a long value(object pointer), and an object variable

both functions will map(reference) the object pointer to the object variable, but in two different ways:

__vbaObjSet: works pretty much as using CopyMemory to write the object pointer into the object variable, can crash your program at ease.

__vbaObjSetAddref: This is the magic wand, will do the same as vbaObjSet, and increase the reference count (if you have any idea about COM interfaces, you'll understand).. thus preventing your application from crashing.

----

now you must be wondering, why would i want to do that?!

in some cases, like when using the timer APIs, you may send an object pointer, and recive it back later, what can you do with the pointer now?!

instead of looking through a collection for the object using its pointer as the key(most vb programmers use this approach) you can simply map the pointer to an object variable, do what you want with the object, and then set it to nothing!!!

you can experiment with other situations (like passing the pointer through DDE, multi-threading)..

another approach is to link two object variables and

and please tell me the results..

----

below is a simple code example of the usage.

I've discovered more functions, but i'm planning on releasing them all as one big (module/tlb), didn't decide yet :)

any comments/suggestions are welcome.
 
### More Info
 
ObjPtr(object), Long

Object

If you use vbaObjSet, it can crash your program if the object is released before you release the new variable..


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[voodooattack](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/voodooattack.md)
**Level**          |Advanced
**User Rating**    |5.0 (30 globes from 6 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/voodooattack-vbaobjsetaddref-the-native-undocumented-visual-basic-run-time-apis__1-62247/archive/master.zip)

### API Declarations

```
Private Declare Function vbaObjSetAddref Lib "MSVBVM60.DLL" Alias "__vbaObjSetAddref" (dstObject As Any, ByVal srcObjPtr As Long) As Long
Private Declare Function vbaObjSet Lib "MSVBVM60.DLL" Alias "__vbaObjSet" (dstObject As Any, ByVal srcObjPtr As Long) As Long
```


### Source Code

```
Private Sub Form_Load()
 Dim x As New Collection  'create a new collection
 Dim y As Object    'empty object
 x.Add "test value 1", "test" 'add some data to the 1st collection to identify it.
 'Now use the function:
 ' note: the first parameter is the target objet variable, which is empty at the moment.
 '  the second is the pointer to the source object (ObjPtr) will return the pointer(DO NOT USE VatPtr).
 Call vbaObjSetAddref(y, ObjPtr(x))  'use the magic wand ;)
 'Now test the second variable for identification
 MsgBox y("test")
 'Done, the second variable refers to the same object in memory.
 'this is not a copy of the 1st object, they are still connected!!!
 'for example, you can add a second value to the (y), and you can read it back from (x)
 y.Add "test value 2", "test2"
 MsgBox x("test2")
 'Imagine the possibilities now (using events perhaps) ;)
End Sub
```

