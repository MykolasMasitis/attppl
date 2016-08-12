PROCEDURE fill_s_pred
SELECT AisOms 
SET ORDER TO 
GO TOP 

SCAN
 lcDir = pBase + '\' + priod + '\' + mcod
 IF fso.FileExists(lcDir+'\talon.dbf')
  tn_result = 0
  tn_result = tn_result + OpenFile("&lcDir\Talon", "Talon", "SHARE")
  IF tn_result == 0
   SELECT Talon
   m.s_pred  = 0
   m.s_predz = 0
   SCAN 
    m.d_type = d_type
    m.s_all  = s_all
    IF INLIST(m.d_type, 'z', 'h')
     m.s_pred = m.s_pred - m.s_all
     m.s_predz = m.s_predz + m.s_all
    ELSE 
     m.s_pred = m.s_pred + m.s_all
    ENDIF 
   ENDSCAN 
   USE
   SELECT AisOms
   REPLACE s_pred WITH m.s_pred, s_predz WITH m.s_predz
  ENDIF 
 ENDIF 
 MailView.refresh
ENDSCAN 

