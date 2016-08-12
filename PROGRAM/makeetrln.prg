PROCEDURE MakeETRLN(lcDir, mcod)
 PRIVATE lcdir,mcod
 
 m.thisid = lpuid
 m.mmy   = SUBSTR(m.gcPeriod,5,2)+SUBSTR(m.gcPeriod,4,1)
 
 IF !fso.FolderExists(lcDir)
  RETURN 
 ENDIF 
 IF !fso.FileExists(lcdir+'\people.dbf')
  RETURN 
 ENDIF 
 IF !fso.FileExists(lcdir+'\efoms.dbf')
  RETURN 
 ENDIF 
 IF OpenFile(lcdir+'\people', 'people', 'shar', 'recid_lpu')>0
  IF USED('people')
   USE IN people
  ENDIF 
  SELECT aisoms 
  RETURN 
 ENDIF 
 IF OpenFile(lcdir+'\efoms', 'efoms', 'shar')>0
  USE IN people
  IF USED('efoms')
   USE IN efoms
  ENDIF 
  SELECT aisoms 
  RETURN 
 ENDIF 
 
 IF fso.FileExists(lcdir+'\etrl'+m.qcod+'.dbf')
  IF !fso.FileExists(lcdir+'\etrl'+m.qcod+m.qcod+'.dbf')
   fso.CopyFile(lcdir+'\etrl'+m.qcod+'.dbf', lcdir+'\etrl'+m.qcod+m.qcod+'.dbf')
  ENDIF 
  fso.DeleteFile(lcdir+'\etrl'+m.qcod+'.dbf')
 ENDIF 
 fso.CopyFile(ptempl+'\etrlqq.dbf', lcdir+'\etrl'+m.qcod+'.dbf', .t.)
 IF !fso.FileExists(lcdir+'\etrl'+m.qcod+'.dbf')
  USE IN people
  USE IN efoms
  SELECT aisoms
  RETURN 
 ENDIF 
 
 IF OpenFile(lcdir+'\etrl'+m.qcod, 'ctrl', 'shar')>0
  IF USED('ctrl')
   USE IN ctrl
  ENDIF 
  USE IN people
  USE IN efoms
  SELECT aisoms
  RETURN 
 ENDIF 
 
 SELECT efoms
 SET RELATION TO rec_mo INTO people
 SCAN 
  m.recid = PADL(RECNO(),6,'0')
  m.rec_mo = rec_mo
  m.date_in = IIF(!EMPTY(people.date_in2), people.date_in2, people.date_in)
  m.date_out = IIF(!EMPTY(people.date_out), people.date_out, {})
  m.spos  = IIF(!EMPTY(people.spos2), people.spos2, people.spos)
  m.s_pol = IIF(!EMPTY(people.s_pol2), people.s_pol2, people.s_pol)
  m.n_pol = IIF(!EMPTY(people.n_pol2), people.n_pol2, people.n_pol)
  m.tip_d = IIF(!EMPTY(people.tip_d2), people.tip_d2, people.tip_d)
  m.q     = people.q
  m.erc = erc
  m.lpu_id = IIF(people.c_err='F9', people.lpu_id, 0)
  m.name_err = ALLTRIM(name_err)
  
  INSERT INTO ctrl FROM MEMVAR 
  
 ENDSCAN 
 SET RELATION OFF INTO people
 
 USE IN people
 USE IN ctrl
 USE IN efoms

 SELECT aisoms
RETURN 
