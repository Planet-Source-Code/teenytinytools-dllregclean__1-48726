<div align="center">

## DLLRegClean


</div>

### Description

DLLRegClean searches Windows Registry and cleans up all references to DLL file appointed by the user.

Every VB developer faced situations when same DLL is listed more than one time in the list of the registered ActiveX DLL's. Broken or multiple installtions in different folders, bugs in installer, broken binary compatibilty during developing of ActiveX are only most common causes. This is a real pain for developer to identify which version should be added to the project, and the best way to avoid DLL conflict is clean up the registry and registered only last version of the DLL.
 
### More Info
 
Name of the DLL.

User may either enter it in inputbox or use a command-line interface.

Utility doesn't create Undo file with deleted Registry entries.

Please remember that editing the registry is a risky operation. Before you edit the registry, make sure you understand how to restore it if a problem occurs.


<span>             |<span>
---                |---
**Submitted On**   |2003-09-19 19:57:44
**By**             |[TeenyTinyTools](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/teenytinytools.md)
**Level**          |Intermediate
**User Rating**    |4.8 (19 globes from 4 users)
**Compatibility**  |VB 6\.0
**Category**       |[Registry](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/registry__1-36.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[DLLRegClea1649289232003\.zip](https://github.com/Planet-Source-Code/teenytinytools-dllregclean__1-48726/archive/master.zip)

### API Declarations

```
RegQueryValueEx
RegCloseKey
RegDeleteKey
RegEnumKey
RegEnumKeyEx
RegOpenKey
RegOpenKeyEx
RegEnumValue
SHDeleteKey
```





