PROCEDURE DelSpareFiles

 pComDel= fso.GetParentFolderName(pcommon) + '\COMMON.DEL'

 CREATE CURSOR curFiles (flname c(12))
 SELECT curFiles
 INDEX ON flname TAG flname 
 SET ORDER TO flname 
 
 INSERT INTO curFiles (flname) VALUES ('admokrxx.dbf')
 INSERT INTO curFiles (flname) VALUES ('errors.dbf')
 INSERT INTO curFiles (flname) VALUES ('osoerzxx.dbf')
 INSERT INTO curFiles (flname) VALUES ('smo.dbf')
 INSERT INTO curFiles (flname) VALUES ('smo.cdx')
 INSERT INTO curFiles (flname) VALUES ('spraboxx.dbf')
 INSERT INTO curFiles (flname) VALUES ('sprlpuxx.dbf')
 INSERT INTO curFiles (flname) VALUES ('spr_ulxx.dbf')
 INSERT INTO curFiles (flname) VALUES ('users.cdx')
 INSERT INTO curFiles (flname) VALUES ('users.dbf')
 INSERT INTO curFiles (flname) VALUES ('usrlpu.dbf')
 INSERT INTO curFiles (flname) VALUES ('tiplpu.dbf')
 
 oDir        = fso.GetFolder(pCommon)
 DirName     = oDir.Path
 oFilesInDir = oDir.Files
 nFilesInDir = oFilesInDir.Count
 
 FOR EACH oFileInDir IN oFilesInDir
  m.bname = LOWER(ALLTRIM(oFileInDir.Name))
  
  IF !SEEK(PADR(m.bname,12), 'curFiles')
   IF !fso.FolderExists(pComDel)
    fso.CreateFolder(pComDel)
   ENDIF 
   fso.CopyFile(pcommon+'\'+m.bname, pcomdel+'\'+m.bname, .t.)
   fso.DeleteFile(pcommon+'\'+m.bname)
  ENDIF 
  
 NEXT 
 
 USE IN curFiles
 
RETURN 