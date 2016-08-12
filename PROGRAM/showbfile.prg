FUNCTION  ShowBFile(bbb)
 tbname = ALLTRIM(bbb)
 wshell.ShellExecute('notepad', tbname)
RETURN 