PROCEDURE PackPrnSv
 IF MESSAGEBOX('�� ������ ����������� ���� ������?',4+32,'')=7
  RETURN 
 ENDIF 
 IF !fso.FileExists(pbase+'\'+m.gcperiod+'\aisoms.dbf')
  RETURN 
 ENDIF 
 IF OpenFile(pbase+'\'+m.gcperiod+'\aisoms', 'aisoms', 'shar')>0
  IF USED('aisoms')
   USE IN aisoms
  ENDIF 
  RETURN 
 ENDIF 

 PUBLIC oWord AS Word.Application
 WAIT "������ MS Word..." WINDOW NOWAIT 
 TRY 
  oWord=GETOBJECT(,"Word.Application")
 CATCH 
  oWord=CREATEOBJECT("Word.Application")
 ENDTRY 
 WAIT CLEAR 

 SELECT * FROM aisoms ORDER BY cokr, mcod INTO CURSOR curais
 USE IN aisoms
 
 SELECT curais
 idocs = 0 
 SCAN 
  m.mcod  = mcod 
  m.lpuid = lpuid
  IF !fso.FolderExists(pbase+'\'+m.gcperiod+'\'+m.mcod)
   LOOP 
  ENDIF 
  m.docname = '\Sv' + LOWER(m.qcod) + PADL(tMonth,2,'0')+RIGHT(STR(tYear,4),2)
  IF !fso.FileExists(pbase+'\'+m.gcperiod+'\'+m.mcod+'\'+m.docname+'.doc')
   LOOP 
  ENDIF 
  
  idocs = idocs+1
  IF idocs=50
   IF MESSAGEBOX('���������� �� ������ 50 �������.'+CHR(13)+CHR(10)+;
    '����������?',4+32,'')=7
    EXIT 
   ELSE 
    idocs = 0 
   ENDIF 
  ENDIF 

  oDoc = oWord.Documents.Add(pbase+'\'+m.gcperiod+'\'+m.mcod+'\'+m.docname+'.doc')
  oDoc.PrintOut(,,3,,"1","1")
  oDoc.Close

 ENDSCAN 
 
 USE IN curais

 oWord.Quit 
 MESSAGEBOX('������ ���������!',0+64,'')

RETURN 