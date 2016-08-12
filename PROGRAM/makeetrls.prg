PROCEDURE MakeETRLs
 IF MESSAGEBOX(CHR(13)+CHR(10)+'¬€ ’Œ“»“≈ —‘Œ–Ã»–Œ¬¿“‹'+CHR(13)+CHR(10)+'‘¿…À€ Œÿ»¡Œ  ƒÀﬂ Àœ”?'+CHR(13)+CHR(10),4+32,'')=7
  RETURN 
 ENDIF 

 IF !fso.FileExists(ptempl+'\etrlqq.dbf')
  MESSAGEBOX(CHR(13)+CHR(10)+'Œ“—”“—“¬”≈“ ÿ¿¡ÀŒÕ ‘¿…À¿ Œÿ»¡Œ  ETRLQQ.DBF'+CHR(13)+CHR(10),0+16,'')
  RETURN 
 ENDIF 
 IF OpenFile("&pBase\&gcPeriod\aisoms", "aisoms", "shar") > 0
  IF USED('aisoms')
   USE IN aisoms
  ENDIF 
  RETURN 
 ENDIF 
 IF OpenFile("&pBase\&gcPeriod\NSI\errors", "errors", "shar", "code") > 0
  IF USED('errors')
   USE IN errors
  ENDIF 
  IF USED('aisoms')
   USE IN aisoms
  ENDIF 
  RETURN 
 ENDIF 

 IF fso.FileExists(pBase+'\'+m.gcperiod+'\etrl'+m.qcod+'.dbf')
  fso.DeleteFile(pBase+'\'+m.gcperiod+'\etrl'+m.qcod+'.dbf')
 ENDIF 
 fso.CopyFile(ptempl+'\etrlqq.dbf', pBase+'\'+m.gcperiod+'\etrl'+m.qcod+'.dbf', .t.)
 
 =OpenFile(pBase+'\'+m.gcperiod+'\etrl'+m.qcod, 'ctrlall', 'shar')

 OldEscStatus = SET("Escape")
 SET ESCAPE OFF 
 CLEAR TYPEAHEAD 
 
 SELECT AisOms
 SCAN
  WAIT mcod WINDOW NOWAIT 

  =MakeETRL(pbase+'\'+gcPeriod+'\'+mcod, mcod)
  
  SELECT AisOms

  IF CHRSAW(0) == .T.
   IF INKEY() == 27
    IF MESSAGEBOX('¬€ ’Œ“»“≈ œ–≈–¬¿“‹ Œ¡–¿¡Œ“ ”?',4+32,'') == 6
     EXIT 
    ENDIF 
   ENDIF 
  ENDIF 

 ENDSCAN 

 USE IN aisoms
 USE IN errors
 USE IN ctrlall

 WAIT CLEAR 

 SET ESCAPE &OldEscStatus

 MESSAGEBOX('Œ¡–¿¡Œ“ ¿ «¿ ŒÕ◊≈Õ¿!', 0+64, '')

RETURN 

