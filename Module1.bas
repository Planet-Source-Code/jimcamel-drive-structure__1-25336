Attribute VB_Name = "Module1"
Option Explicit
'declare the variables needed for the File Number and a FileSystemObject
Public Filenum As Integer
Public FS As FileSystemObject

Public Function logtofile(recursionlevel As Integer, Path) As Boolean
'Right, here is where the hard work comes in. This example uses
'The FileSystemObject and recursion (shudder) to do the hard work.
'Recursion is very handy, as it makes code very small, but it's also
'quite hard to understand.
'I will try my best to explain it, but it's quite complicated

'I use recursionlevel to know how many spaces to put in the textfile
'It's not needed, but otherwise the output file would be near impossible
'to understand

Dim Fldr As Folder, SubFldr As Folder, fil As File
'create the FileSystemObject
Set FS = CreateObject("Scripting.FileSystemObject")
'Set the Folder to the path passed in the path
Set Fldr = FS.GetFolder(Path)
    DoEvents
    'Go through each Subfolder in the Folders Path
    For Each SubFldr In Fldr.SubFolders
        'Print the Folder name to the file
        Print #Filenum, Spc(recursionlevel * 2); "|--" & SubFldr
        'Call this function with the sub folder, and one more recursion number
        logtofile recursionlevel + 1, SubFldr
    Next
    
    'When we are at the top folder, go through all the files
    For Each fil In Fldr.Files
        'and print them out
        Print #Filenum, Spc(recursionlevel * 2); "|--" & fil
    Next
'Then we go back a folder

'If you want any more help understanding this, email me on jimcamel@jimcamel.8m.com
'or ICQ me at 25282667
End Function
