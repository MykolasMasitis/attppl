PROCEDURE MakeFOMS
 IF MESSAGEBOX(CHR(13)+CHR(10)+'—‘Œ–Ã»–Œ¬¿“‹ Œ“◊≈“ ¬ Ã√‘ŒÃ—?'+CHR(13)+CHR(10),4+32,'')==7
  RETURN 
 ENDIF 
 IF !fso.FileExists(ptempl+'\prqqmmyy.dbf')
  MESSAGEBOX(CHR(13)+CHR(10)+'Œ“—”“—“¬”≈“ ‘¿…À ÿ¿¡ÀŒÕ¿ '+'PRQQMMYY.DBF'+CHR(13)+CHR(10),0+16,'')
  RETURN 
 ENDIF 
 IF !fso.FileExists(pbase+'\'+gcperiod+'\people.dbf')
  MESSAGEBOX(CHR(13)+CHR(10)+'—¬ŒƒÕ€… –≈√»—“– Õ≈ —Œ¡–¿Õ!'+CHR(13)+CHR(10),0+16,'')
  RETURN 
 ENDIF 
 IF OpenFile(pbase+'\'+gcperiod+'\people', 'people', 'excl')>0
  RETURN 
 ENDIF 
 
 WAIT "«¿œ”—  —–≈ƒ—“¬¿ —¿ÃŒƒ»¿√ÕŒ—“» »..." WINDOW NOWAIT 
  m.nDubls = 0
  m.nCross = 0
  SELECT * FROM people WHERE n_pol2 in ;
   (SELECT n_pol2 FROM people GROUP BY n_pol2 HAVING COUNT(*)>1) INTO CURSOR dubls
  SELECT dubls
  m.nDubls = RECCOUNT('dubls')
  USE IN dubls
  IF tdat1>={01.01.2013} && Ì‡‰Ó ÔÓ‚ÂˇÚ¸ ÔÓ ˝Ú‡ÎÓÌÌÓÏÛ Â„ËÒÚÛ!
   IF fso.FileExists(pcommon+'\people.dbf')
    IF OpenFile(pcommon+'\people', 'ppl', 'shar')>0
     IF USED('ppl')
      USE IN ppl 
     ENDIF 
    ELSE 
     SELECT * FROM people WHERE n_pol2 in (SELECT n_pol FROM ppl) INTO CURSOR cross
     SELECT cross
     m.nCross = RECCOUNT('cross')
     IF m.nCross>0
      COPY TO pbase+'\'+gcperiod+'\dbls' 
     ENDIF 
     USE IN cross
    ENDIF 
   ELSE 
   ENDIF 
  ENDIF 
  IF m.nDubls<=0
   MESSAGEBOX(CHR(13)+CHR(10)+'œ–Œ¡À≈Ã Õ≈ Œ¡Õ¿–”∆≈ÕŒ!'+CHR(13)+CHR(10),0+64,'')
  ELSE 
   MESSAGEBOX(CHR(13)+CHR(10)+'Œ¡Õ¿–”∆≈ÕŒ '+ALLTRIM(STR(m.nDubls))+' ƒ”¡À≈…!'+;
   CHR(13)+CHR(10),0+64,'')
  ENDIF 
  IF m.nCross<=0
   MESSAGEBOX(CHR(13)+CHR(10)+'œ–Œ¡À≈Ã Õ≈ Œ¡Õ¿–”∆≈ÕŒ!'+CHR(13)+CHR(10),0+64,'')
  ELSE 
   MESSAGEBOX(CHR(13)+CHR(10)+'Œ¡Õ¿–”∆≈ÕŒ '+ALLTRIM(STR(m.nCross))+' ƒ”¡À≈…!'+;
   CHR(13)+CHR(10),0+64,'')
  ENDIF 
 WAIT CLEAR 
 
 prfile = 'pr'+LOWER(m.qcod)+PADL(tMonth,2,'0')+RIGHT(STR(tYear,4),2)
 fso.CopyFile(ptempl+'\prqqmmyy.dbf', ;
  pbase+'\'+gcperiod+'\'+prfile+'.dbf')

 IF OpenFile(pbase+'\'+gcperiod+'\'+prfile, 'prfile', 'shar')>0
  IF USED('prfile')
   USE IN prfile
  ENDIF 
  fso.DeleteFile(pbase+'\'+gcperiod+'\'+prfile+'.dbf')
  RETURN 
 ENDIF 
 
 WAIT "Œ¡–¿¡Œ“ ¿..." WINDOW NOWAIT 
 SELECT people
 INDEX ON lpu_id TAG lpu_id
 SET ORDER TO lpu_id
 SCAN 
  m.recid   = PADL(RECNO(),7,'0')
  m.lpu_id  = lpu_id
  m.date_in = date_in
  m.spos    = spos
  m.s_pol   = s_pol2
  m.n_pol   = n_pol2
  m.tip_d   = IIF(tip_d2='—', '1', '3')
  m.q       = m.qcod

  INSERT INTO prfile FROM MEMVAR 
 ENDSCAN 
 SET ORDER TO 
 DELETE TAG ALL 
 USE IN people 
 USE IN prfile
 WAIT CLEAR 

RETURN 