PROCEDURE SndToLpu
 IF MESSAGEBOX(CHR(13)+CHR(10)+'������ ��������� ����� � ���?'+CHR(13)+CHR(10),4+32, '')=7
  RETURN 
 ENDIF 
 IF !fso.FileExists(pbase+'\'+gcperiod++'\aisoms.dbf')
  MESSAGEBOX(CHR(13)+CHR(10)+'����������� ���� AISOMS.DBF'+CHR(13)+CHR(10),0+16,'')
  RETURN 
 ENDIF 
 IF OpenFile(pbase+'\'+gcperiod++'\aisoms', 'aisoms', 'shar', 'mcod')>0
  IF USED('aisoms')
   USE IN aisoms
  ENDIF 
  RETURN 
 ENDIF 
 IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\spraboxx', "sprabo", "shar", "lpu_id") > 0
  IF USED('aisoms')
   USE IN aisoms
  ENDIF 
  IF USED('sprabo')
   USE IN sprabo
  ENDIF 
  RETURN
 ENDIF 

 SELECT aisoms
 SCAN 
  IF EMPTY(bname)
*   LOOP 
  ENDIF 
  
  m.lcPath = pbase+'\'+gcPeriod+'\'+mcod
  m.cmcod  = mcod
  m.clpuid = STR(lpuid,4)

  =SendToLpu(lcPath, cmcod, clpuid)

 ENDSCAN 
 USE 
 USE IN sprabo

 MESSAGEBOX(CHR(13)+CHR(10)+'��������� ���������!'+CHR(13)+CHR(10),0+64,'')

RETURN 