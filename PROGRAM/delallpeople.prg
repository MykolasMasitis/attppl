PROCEDURE DelAllPeople
 IF MESSAGEBOX('ÁÓÄÓÒ ÓÄÀËÅÍÛ ÂÑÅ ALLPEOPLE-ÔÀÉËÛ?'+CHR(13)+CHR(10)+;
               'ÝÒÎ ÒÎ, ×ÒÎ ÂÛ ÄÅÉÑÒÂÈÒÅËÜÍÎ ÕÎÒÈÒÅ ÑÄÅËÀÒÜ?',4+48,'') != 6
  RETURN 
 ENDIF 

 IF MESSAGEBOX('ÂÛ ÀÁÑÎËÞÒÍÎ ÓÂÅÐÅÍÛ Â ÑÂÎÈÕ ÄÅÉÑÒÂÈßÕ?',4+48,'') != 6
  RETURN 
 ENDIF 
 
 IF OpenFile("&pBase\&gcPeriod\aisoms", "aisoms", "shar", "mcod") > 0
  RETURN
 ENDIF 
 
 SELECT AisOms
 
 SCAN
  m.mcod = mcod

  WAIT m.mcod WINDOW NOWAIT 

  IF fso.FileExists(pBase+'\'+m.gcperiod+'\'+mcod+'\allpeople.dbf')
   fso.DeleteFile(pBase+'\'+m.gcperiod+'\'+mcod+'\allpeople.dbf')
  ENDIF 
  IF fso.FileExists(pBase+'\'+m.gcperiod+'\'+mcod+'\allpeople.cdx')
   fso.DeleteFile(pBase+'\'+m.gcperiod+'\'+mcod+'\allpeople.cdx')
  ENDIF 
  IF fso.FileExists(pBase+'\'+m.gcperiod+'\'+mcod+'\allpeople.bak')
   fso.DeleteFile(pBase+'\'+m.gcperiod+'\'+mcod+'\allpeople.bak')
  ENDIF 

 ENDSCAN 
 WAIT CLEAR 
 USE 

RETURN 