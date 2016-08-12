PROCEDURE AllResetErrors

 OldEscStatus = SET("Escape")
 SET ESCAPE OFF 
 CLEAR TYPEAHEAD 
 
 SCAN
  WAIT mcod WINDOW NOWAIT 
  MailView.refresh

  =reseterrors(pbase+'\'+gcPeriod+'\'+mcod, mcod, .t.)  

  SELECT AisOms
  MailView.pazbad = MailView.pazbad - pazbad
  MailView.pazdbl = MailView.pazdbl - pazdbl
  MailView.adults = MailView.adults - adults
  MailView.childs = MailView.childs - childs

  MailView.refresh
  
  REPLACE pazbad WITH 0, pazdbl WITH 0, adults WITH 0, childs WITH 0

  IF CHRSAW(0) == .T.
   IF INKEY() == 27
    IF MESSAGEBOX('¬€ ’Œ“»“≈ œ–≈–¬¿“‹ Œ¡–¿¡Œ“ ”?',4+32,'') == 6
     EXIT 
    ENDIF 
   ENDIF 
  ENDIF 

 ENDSCAN 

 WAIT CLEAR 

 SET ESCAPE &OldEscStatus

 MESSAGEBOX('Œ¡–¿¡Œ“ ¿ «¿ ŒÕ◊≈Õ¿!', 0+64, '')

RETURN 

