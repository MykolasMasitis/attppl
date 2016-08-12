PROCEDURE MakeFOMS2
 IF MESSAGEBOX(CHR(13)+CHR(10)+'������������ ����� � ������?'+CHR(13)+CHR(10),4+32,'����� ������')==7
  RETURN 
 ENDIF 
 IF !fso.FileExists(ptempl+'\prqqmmyy.dbf')
  MESSAGEBOX(CHR(13)+CHR(10)+'����������� ���� ������� '+'PRQQMMYY.DBF'+CHR(13)+CHR(10),0+16,'')
  RETURN 
 ENDIF 
 IF !fso.FileExists(pbase+'\'+gcperiod+'\aisoms.dbf')
  MESSAGEBOX(CHR(13)+CHR(10)+'����������� ���� AISOMS.DBF!'+CHR(13)+CHR(10),0+16,'')
  RETURN 
 ENDIF 
 IF OpenFile(pbase+'\'+gcperiod+'\aisoms', 'aisoms', 'shar')>0
  IF USED('aisoms')
   USE IN aisoms
  ENDIF 
  RETURN 
 ENDIF 
 
 prfile = 'pr'+LOWER(m.qcod)+PADL(tMonth,2,'0')+RIGHT(STR(tYear,4),2)
 fso.CopyFile(ptempl+'\prqqmmyy.dbf', ;
  pbase+'\'+gcperiod+'\'+prfile+'.dbf')

 IF OpenFile(pbase+'\'+gcperiod+'\'+prfile, 'prfile', 'shar')>0
  IF USED('prfile')
   USE IN prfile
  ENDIF 
  fso.DeleteFile(pbase+'\'+gcperiod+'\'+prfile+'.dbf')
  USE IN aisoms
  RETURN 
 ENDIF 
 
 SELECT aisoms 
 WAIT "���������..." WINDOW NOWAIT 
 SCAN 
  m.paztot = paztot
  m.lpu_id  = lpuid
  IF m.paztot<=0
   LOOP 
  ENDIF 
  m.mcod = mcod 
  IF !fso.FolderExists(pbase+'\'+m.gcperiod+'\'+m.mcod)
   LOOP 
  ENDIF 
  IF !fso.FileExists(pbase+'\'+m.gcperiod+'\'+m.mcod+'\people.dbf')
   LOOP 
  ENDIF 
  IF OpenFile(pbase+'\'+m.gcperiod+'\'+m.mcod+'\people', 'people', 'shar')>0
   IF USED('people')
    USE IN people 
   ENDIF 
   SELECT aisoms 
   LOOP 
  ENDIF 
  
  SELECT people
  SCAN 
*   m.recid   = recid_lpu
*   m.lpu_id  = lpu_id
*   m.recid = PADL(recid ,7,'0')
*   m.recid   = PADL(RECNO(),7,'0')
   m.date_in = date_in
   m.spos    = spos
   m.s_pol   = s_pol
   m.n_pol   = n_pol
   m.tip_d   = tip_d
   m.q       = m.qcod
   m.reserv  = reserv

   INSERT INTO prfile FROM MEMVAR 
  ENDSCAN 
  USE IN people 
  
  SELECT aisoms 
  
 ENDSCAN 
 
 USE IN aisoms
 
 SELECT prfile 
 REPLACE ALL recid WITH PADL(RECNO(),7,'0')
 USE IN prfile
 
 WAIT CLEAR 

RETURN 