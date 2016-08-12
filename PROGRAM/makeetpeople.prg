PROCEDURE MakeEtPeople
 IF MESSAGEBOX(CHR(13)+CHR(10)+'’Œ“»“≈ —‘Œ–Ã»–Œ¬¿“‹ ›“¿ÀŒÕÕ€… –≈√»—“–?'+CHR(13)+CHR(10),4+32,'')==7
  RETURN 
 ENDIF 
 IF !fso.FileExists(pbase+'\'+gcperiod+'\people.dbf')
  MESSAGEBOX(CHR(13)+CHR(10)+'—¬ŒƒÕ€… –≈√»—“– Õ≈ —Œ¡–¿Õ!'+CHR(13)+CHR(10),0+16,'')
  RETURN 
 ENDIF 
 FomsErrFile = 'er'+m.qcod+PADL(tmonth,2,'0')+RIGHT(STR(tyear,4),2)
 IF !fso.FileExists(pbase+'\'+gcperiod+'\'+FomsErrFile+'.dbf')
  MESSAGEBOX(CHR(13)+CHR(10)+'‘¿…À Œÿ»¡Œ  ‘ŒÕƒ¿ Õ≈ «¿√–”∆≈Õ!'+CHR(13)+CHR(10),0+16,'')
  RETURN 
 ENDIF 

 IF OpenFile(pbase+'\'+gcperiod+'\people', 'people', 'excl')>0
  IF USED('people')
   USE IN people
  ENDIF 
  RETURN 
 ENDIF 

 tFile = pbase+'\'+gcperiod+'\'+FomsErrFile+'.dbf'
 oSettings.CodePage('&tFile', 866, .t.)

 IF OpenFile(pbase+'\'+gcperiod+'\'+FomsErrFile, 'errors', 'excl')>0
  IF USED('people')
   USE IN people
  ENDIF 
  IF USED('errors')
   USE IN errors
  ENDIF 
  RETURN 
 ENDIF 
 
 WAIT "‘Œ–Ã»–Œ¬¿Õ»≈ ›“¿ÀŒÕÕŒ√Œ –≈√»—“–¿..." WINDOW NOWAIT 
 
* IF fso.FileExists(pcommon+'\people.dbf')
*  fso.DeleteFile(pcommon+'\people.dbf')
* ENDIF 
 
 IF !fso.FileExists(pcommon+'\people.dbf')
  ppeople = pcommon+'\people'
  CREATE TABLE (PPeople) ;
   (lpu_id n(6), mcod c(7), period c(6), recid i, spos c(1), date_in d, date_out d, ;
    tip_d c(1), s_pol c(6), n_pol c(16), fam c(25), im c(20), ot c(20), dr d, w n(1))
  INDEX on n_pol TAG n_pol 
  USE 
 ENDIF 
 
 IF OpenFile(pcommon+'\people', 'svppl', 'excl', 'n_pol')>0
  IF USED('people')
   USE IN people
  ENDIF 
  IF USED('errors')
   USE IN errors
  ENDIF 
  IF USED('svppl')
   USE IN svppl 
  ENDIF 
  RETURN 
 ENDIF 
 
 SELECT errors
 INDEX ON recid TAG recid 
 SET ORDER TO recid
 
 SELECT people
 SET RELATION TO PADL(recid,7,'0') INTO errors
 
 m.nAdded = 0
 m.nModed = 0 
 
 SCAN 
  IF !EMPTY(errors.recid)
   LOOP
  ENDIF 

  m.n_pol    = n_pol2
  m.date_out = date_out

  IF SEEK(m.n_pol, 'svppl')
   IF EMPTY(m.date_out)
    LOOP 
   ELSE 
    IF EMPTY(svppl.date_out)
     REPLACE svppl.date_out WITH m.date_out IN svppl
     m.nModed = m.nModed + 1
    ENDIF 
    LOOP 
   ENDIF 
  ENDIF 
  
  m.lpu_id   = lpu_id
  m.mcod     = mcod
  m.period   = m.gcperiod
  m.recid    = recid
  m.date_in  = date_in
  m.spos     = spos
  m.s_pol    = s_pol2
  m.tip_d    = IIF(tip_d2='—', '1', '3')
  m.fam      = fam
  m.im       = im
  m.ot       = ot
  m.dr       = dr
  m.w        = w
  
  m.nAdded = m.nAdded + 1
  INSERT INTO svppl FROM MEMVAR 
  
 ENDSCAN 
 
 SET RELATION OFF INTO errors
 USE 
 SELECT errors
 SET ORDER TO 
 DELETE TAG ALL 
 USE 
 
 USE IN svppl
 
 WAIT CLEAR 
 
 MESSAGEBOX(CHR(13)+CHR(10)+'›“¿ÀŒÕÕ€… –≈√»—“– —‘Œ–Ã»–Œ¬¿Õ!'+CHR(13)+CHR(10)+;
 'ƒŒ¡¿¬À≈ÕŒ '+TRANSFORM(m.nAdded,'999999')+' «¿œ»—≈…, '+CHR(13)+CHR(10)+;
 'œŒ√¿ÿ≈ÕŒ '+TRANSFORM(m.nModed,'999999')+' «¿œ»—≈…, '+CHR(13)+CHR(10), 0+64,'')

RETURN 