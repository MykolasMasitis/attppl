PROCEDURE MakeETRL(lcDir, mcod)
 PRIVATE lcdir,mcod
 
 m.thisid = lpuid
 m.mmy   = SUBSTR(m.gcPeriod,5,2)+SUBSTR(m.gcPeriod,4,1)
 
 IF !fso.FolderExists(lcDir)
  RETURN 
 ENDIF 
 IF !fso.FileExists(lcdir+'\people.dbf')
  RETURN 
 ENDIF 
 IF OpenFile(lcdir+'\people', 'people', 'shar')>0
  IF USED('people')
   USE IN people
  ENDIF 
  SELECT aisoms 
  RETURN 
 ENDIF 
 
 IF fso.FileExists(lcdir+'\etrl'+m.qcod+'.dbf')
  fso.DeleteFile(lcdir+'\etrl'+m.qcod+'.dbf')
 ENDIF 
 fso.CopyFile(ptempl+'\etrlqq.dbf', lcdir+'\etrl'+m.qcod+'.dbf', .t.)
 IF !fso.FileExists(lcdir+'\etrl'+m.qcod+'.dbf')
  USE IN people
  SELECT aisoms
  RETURN 
 ENDIF 
 
 m.lWasUsedError = .T.
 IF !USED('errors')
  m.lWasUsedError = .F.
  IF OpenFile("&pBase\&gcPeriod\NSI\errors", "errors", "shar", "code") > 0
   IF USED('errors')
    USE IN errors
   ENDIF 
   IF USED('people')
    USE IN people
   ENDIF 
   SELECT aisoms
   RETURN 
  ENDIF 
 ENDIF 
 
 rItem    = 'PR'+STR(m.thisid,4)+m.qcod + '.'+m.mmy
 m.lIsRItem = .F.
 IF fso.FileExists(lcdir+'\'+rItem)
  m.lIsRItem = .T.
 ENDIF 
 IF m.lIsRItem = .F.
  IF m.lWasUsedError=.F.
   IF USED('errors')
    USE IN errors
   ENDIF 
  ENDIF 
  IF USED('people')
   USE IN people
  ENDIF 
  SELECT aisoms
  RETURN 
 ENDIF 
 
 =OpenFile(lcdir+'\'+rItem, 'lpuip', 'excl')
 SELECT lpuip
 INDEX ON recid+n_pol TAG n_pol
 SET ORDER TO n_pol

 =OpenFile(lcdir+'\etrl'+m.qcod, 'ctrl', 'shar')
 
 SELECT people 
 SET RELATION TO c_err INTO errors
 SET RELATION TO recid_lpu+n_pol INTO lpuip ADDITIVE 
 SCAN 
  m.recid = PADL(recid,6,'0')
  m.rec_mo = recid_lpu
  m.date_in = IIF(!EMPTY(date_in2), date_in2, date_in)
  m.date_out = IIF(!EMPTY(date_out), date_out, {})
  m.spos  = IIF(!EMPTY(spos2), spos2, spos)
  m.s_pol = IIF(!EMPTY(s_pol2), s_pol2, s_pol)
  m.n_pol = IIF(!EMPTY(n_pol2), n_pol2, n_pol)
  m.tip_d = IIF(!EMPTY(tip_d2), tip_d2, tip_d)
  m.q     = q
  m.erc = IIF(EMPTY(c_err) AND !EMPTY(date_out), 'OK', c_err)
  m.erc = IIF(m.erc='OK' AND (s_pol!=s_pol2 OR n_pol!=n_pol2), 'FF', m.erc)
  m.lpu_id = IIF(m.erc='F9', lpu_id, 0)
  m.name_err = ALLTRIM(errors.name)
  IF EMPTY(m.name_err) AND !EMPTY(m.date_out)
   m.name_err = 'Открепление подтверждено (сообщение)'
  ENDIF 
  m.reserv = reserv
  
  IF !EMPTY(lpuip.recid)
   INSERT INTO ctrl FROM MEMVAR 
   IF USED('ctrlall')
    INSERT INTO ctrlall FROM MEMVAR 
   ENDIF 
  ENDIF 
  
 ENDSCAN 
 SET RELATION OFF INTO errors
 SET RELATION OFF INTO lpuip
 
 USE IN people
 USE IN ctrl
 SELECT lpuip
 SET ORDER TO 
 DELETE TAG ALL 
 USE IN lpuip 

 IF m.lWasUsedError=.F.
  IF USED('errors')
   USE IN errors
  ENDIF 
 ENDIF 

 SELECT aisoms
RETURN 
