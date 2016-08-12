PROCEDURE MakeETRLs2
 IF !fso.FileExists(ptempl+'\etrlqq.dbf')
  MESSAGEBOX(CHR(13)+CHR(10)+'Œ“—”“—“¬”≈“ ÿ¿¡ÀŒÕ ‘¿…À¿ Œÿ»¡Œ  ETRLQQ.DBF'+CHR(13)+CHR(10),0+16,'')
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
 
 SELECT people
 INDEX ON PADL(RECNO(),7,'0') TAG recid 
 SET ORDER TO recid
 SELECT errors
 IF UPPER(FIELD('mcod')) !=  'MCOD'
  ALTER TABLE errors ADD COLUMN mcod c(7)
 ENDIF 
 SET RELATION TO recid INTO people
 m.nRecsInErrors = RECCOUNT('errors')
 m.nRelatedRecs = 0 
 COUNT FOR !EMPTY(people.recid) TO m.nRelatedRecs
 
 IF m.nRelatedRecs != m.nRecsInErrors
  SET RELATION OFF INTO people
  USE IN errors
  SELECT people 
  SET ORDER TO 
  DELETE TAG ALL 
  USE 
  MESSAGEBOX(CHR(13)+CHR(10)+'‘¿…À€ Õ≈ —¬ﬂ«¿À»—‹!'+CHR(13)+CHR(10),0+16,'')
  RETURN 
 ENDIF 

 REPLACE ALL mcod WITH people.mcod
 
 SELECT mcod, COUNT(*) FROM errors GROUP BY mcod INTO CURSOR grmcod
 
 SELECT grmcod
 SCAN 
  m.mcod = mcod
  IF EMPTY(m.mcod)
   LOOP 
  ENDIF 
  IF !fso.FolderExists(pbase+'\'+gcperiod+'\'+m.mcod)
   LOOP 
  ENDIF 
  IF fso.FileExists(pbase+'\'+gcperiod+'\'+m.mcod+'\efoms.dbf')
   fso.DeleteFile(pbase+'\'+gcperiod+'\'+m.mcod+'\efoms.dbf')
  ENDIF 
  efoms = pbase+'\'+gcperiod+'\'+m.mcod+'\efoms'
  CREATE TABLE (efoms) (rec_mo c(6), erc c(2), name_err c(50))
*  INDEX on rec_mo TAG rec_mo
  USE 
 ENDSCAN 
 USE IN grmcod
 
 SELECT errors
 SCAN 
  m.mcod = people.mcod
  IF EMPTY(m.mcod)
   LOOP 
  ENDIF 
  IF !fso.FolderExists(pbase+'\'+gcperiod+'\'+m.mcod)
   LOOP 
  ENDIF 
  WAIT m.mcod+'...' WINDOW NOWAIT 
*  MESSAGEBOX(CHR(13)+CHR(10)+m.mcod+CHR(13)+CHR(10),0+64,'')
  IF !fso.FileExists(pbase+'\'+gcperiod+'\'+m.mcod+'\people.dbf')
   LOOP 
  ENDIF 
  IF OpenFile(pbase+'\'+gcperiod+'\'+m.mcod+'\people', 'lpppl', 'shar', 'recid')>0
   IF USED('lpppl')
    USE IN lpppl
   ENDIF 
   SELECT errors
   LOOP 
  ENDIF 
*  IF fso.FileExists(pbase+'\'+gcperiod+'\'+m.mcod+'\efoms.dbf')
*   fso.DeleteFile(pbase+'\'+gcperiod+'\'+m.mcod+'\efoms.dbf')
*  ENDIF 
*  efoms = pbase+'\'+gcperiod+'\'+m.mcod+'\efoms'
*  CREATE TABLE (efoms) (rec_mo c(6), erc c(2), name_err c(50))
*  USE 
*  =OpenFile(pbase+'\'+gcperiod+'\'+m.mcod+'\efoms', 'efoms', 'shar', 'rec_mo')
  =OpenFile(pbase+'\'+gcperiod+'\'+m.mcod+'\efoms', 'efoms', 'shar')
  SELECT errors 
  m.rec_mo   = IIF(SEEK(people.recid, 'lpppl'), lpppl.recid_lpu, '')
  m.erc      = etr
  m.name_err = comment
  
*  IF !SEEK(m.rec_mo, 'efoms')
   INSERT INTO efoms FROM MEMVAR 
*  ENDIF 
  
  USE IN efoms
  USE IN lpppl
 
 ENDSCAN 
 WAIT CLEAR 

 SET RELATION OFF INTO people
 USE IN errors
 SELECT people 
 SET ORDER TO 
 DELETE TAG ALL 
 USE 

 WAIT CLEAR 
 MESSAGEBOX('Œ¡–¿¡Œ“ ¿ «¿ ŒÕ◊≈Õ¿!', 0+64, '')

RETURN 

