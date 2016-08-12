PROCEDURE DelAllZapros
 IF MESSAGEBOX('ÁÓÄÓÒ ÓÄÀËÅÍÛ ÂÑÅ ÑÔÎÐÌÈÐÎÂÀÍÍÍÛÅ ÐÀÍÅÅ '+CHR(13)+CHR(10)+;
               'ÔÀÉËÛ ÇÀÏÐÎÑÎÂ Ê ÅÐÇ (Zapros.dbf)!'+CHR(13)+CHR(10)+;
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
  
  REPLACE erzsv4_id WITH '', sv4st WITH 0

 ENDSCAN 
 WAIT CLEAR 
 USE 

RETURN 