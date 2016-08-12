PROCEDURE MakeETRLsN
 IF MESSAGEBOX(CHR(13)+CHR(10)+'бш унрхре ятнплхпнбюрэ'+CHR(13)+CHR(10)+'тюикш ньханй дкъ кос?'+CHR(13)+CHR(10),4+32,'')=7
  RETURN 
 ENDIF 

 IF !fso.FileExists(ptempl+'\etrlqq.dbf')
  MESSAGEBOX(CHR(13)+CHR(10)+'нрясрярбсер ьюакнм тюикю ньханй ETRLQQ.DBF'+CHR(13)+CHR(10),0+16,'')
  RETURN 
 ENDIF 
 IF OpenFile("&pBase\&gcPeriod\aisoms", "aisoms", "shar") > 0
  IF USED('aisoms')
   USE IN aisoms
  ENDIF 
  RETURN 
 ENDIF 

 OldEscStatus = SET("Escape")
 SET ESCAPE OFF 
 CLEAR TYPEAHEAD 
 
 SELECT AisOms
 SCAN
  WAIT mcod WINDOW NOWAIT 

  =MakeETRLN(pbase+'\'+gcPeriod+'\'+mcod, mcod)
  
  SELECT AisOms

  IF CHRSAW(0) == .T.
   IF INKEY() == 27
    IF MESSAGEBOX('бш унрхре опепбюрэ напюанрйс?',4+32,'') == 6
     EXIT 
    ENDIF 
   ENDIF 
  ENDIF 

 ENDSCAN 

 USE IN aisoms

 WAIT CLEAR 

 SET ESCAPE &OldEscStatus

 MESSAGEBOX('напюанрйю гюйнмвемю!', 0+64, '')

RETURN 

