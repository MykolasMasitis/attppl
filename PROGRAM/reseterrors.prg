PROCEDURE ResetErrors(lcDir, mcod, IsOK)
 PRIVATE lcdir,mcod,isok

 IF !fso.FolderExists(lcDir)
  RETURN .F.
 ENDIF 

 IF !fso.FileExists(lcDir + '\People.dbf')
  RETURN .F.
 ENDIF 

 IF OpenFile("&lcDir\People", "People", "SHARE")>0
  RETURN .F.
 ENDIF 
 
 IF RECCOUNT('people')<=0
  USE IN people 
  RETURN .F.
 ENDIF 
 
 WAIT mcod WINDOW NOWAIT 
 
 SELECT people
 REPLACE ALL c_err WITH '', lpu_id WITH 0, date_in2 WITH {}, spos2 WITH ''
* REPLACE ALL c_err WITH ''
 USE 
 
 WAIT CLEAR 

RETURN 