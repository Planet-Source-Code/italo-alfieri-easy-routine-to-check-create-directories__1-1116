<div align="center">

## Easy routine to check/create directories


</div>

### Description

This very simple routine avoid checking if a correct path already

exist before using it and, if not, create it exactly as you want.

Imagine you wont to write a log file in a path defined as:

C:\Myapplic\Services\logs\LOG.TXT

you must check before if the directory Myapplic exist and

then check all other subdirectory (Service,logs) before opening the

file For Output. Probably you will use a lot of Error Resume Next, Mkdir(...),

Error GoTo 0, dir(....) and so.

Instead you can use this routine as described below:

Myfile="C:\Myapplic\Services\logs\LOG.TXT"

Call CheckDir(Myfile)

nf=FreeFile()

Open Myfile For Output As #nf

.

.

.

Close #nf

and including the following .bas module:

Public Sub CheckDir(file)

Ix = 4 'Initial index

KSlash = InStr(1, file, "\", 1) 'Search for first "\"

For Cnt = 1 To Len(file) 'Run until discover

'other directories

KSlash = InStr((KSlash + 1), file, "\", 1)

If KSlash = 0 Then Exit For 'Last slash

dir1 = Left(file, (KSlash - 1))

cdir1 = Mid(dir1, Ix)

Ix = Ix + Len(cdir1) + 1

hh = Dir(dir1, vbDirectory)

'If Directory doesn't exist, create it

If StrComp(hh, cdir1, 1) <> 0 Then

MkDir (dir1)

End If

Next Cnt

End Sub
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Italo ALFIERI](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/italo-alfieri.md)
**Level**          |Unknown
**User Rating**    |4.2 (159 globes from 38 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/italo-alfieri-easy-routine-to-check-create-directories__1-1116/archive/master.zip)





### Source Code

```
Public Sub CheckDir(file)
		Ix = 4 'Initial index
		KSlash = InStr(1, file, "\", 1) 'Search for first "\"
  		For Cnt = 1 To Len(file) 'Run until discover
               	 	 'other directories
    		KSlash = InStr((KSlash + 1), file, "\", 1)
    		If KSlash = 0 Then Exit For 'Last slash
    		dir1 = Left(file, (KSlash - 1))
    		cdir1 = Mid(dir1, Ix)
    		Ix = Ix + Len(cdir1) + 1
    		hh = Dir(dir1, vbDirectory)
    		'If Directory doesn't exist, create it
    		If StrComp(hh, cdir1, 1) <> 0 Then
      			MkDir (dir1)
    		End If
   		Next Cnt
	End Sub
```

