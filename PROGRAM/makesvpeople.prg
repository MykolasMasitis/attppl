PROCEDURE MakeSvPeople
 IF fso.FileExists(pbase+'\'+gcperiod+'\people.dbf')
  IF MESSAGEBOX(CHR(13)+CHR(10)+'������� ������� ��� ������.'+CHR(13)+CHR(10)+;
   '������ ��������������� ���?',4+32,'')=7
   RETURN 
  ENDIF 
 ENDIF 

 IF MESSAGEBOX(CHR(13)+CHR(10)+'�� ������ �������'+CHR(13)+CHR(10)+;
  '������� �������?'+CHR(13)+CHR(10),4+32,'')=7
  RETURN 
 ENDIF 

 IF MESSAGEBOX(CHR(13)+CHR(10)+'����� �������� �������� ��������'+CHR(13)+CHR(10)+;
  '���������� ����� ������� ������ ����������!'+CHR(13)+CHR(10)+'����������?'+;
  CHR(13)+CHR(10),4+32,'') = 7
  RETURN 
 ENDIF 
  
 IF !fso.FolderExists(pbase)
  =MESSAGEBOX(CHR(13)+CHR(10)+'���������� '+UPPER(ALLTRIM(PBASE))+ '�����������!'+CHR(13)+CHR(10),0+64,'')
  RETURN 
 ENDIF 
 
 IF !fso.FolderExists(pbase+'\'+gcperiod)
  =MESSAGEBOX(CHR(13)+CHR(10)+'���������� '+UPPER(ALLTRIM(pbase+'\'+gcperiod))+ '�����������!'+CHR(13)+CHR(10),0+64,'')
  RETURN 
 ENDIF 

 IF !fso.FileExists(pbase+'\'+gcperiod+'\aisoms.dbf')
  =MESSAGEBOX(CHR(13)+CHR(10)+'����������� ���� '+UPPER(ALLTRIM(pbase+'\'+gcperiod+'\aisoms.dbf'))+ '!'+;
   CHR(13)+CHR(10),0+64,'')
  RETURN 
 ENDIF 
 
 IF OpenFile(pBase+'\'+gcPeriod+'\aisoms', 'aisoms', 'shar') > 0
  IF USED('aisoms')
   USE IN aisoms
  ENDIF 
  RETURN 
 ENDIF 
 IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\sprlpuxx', "sprlpu", "shar", "mcod") > 0
  IF USED('sprlpu')
   USE IN sprlpu
  ENDIF 
  IF USED('aisoms')
   USE IN aisoms
  ENDIF 
  RETURN
 ENDIF 

 CREATE CURSOR svppl ;
  (lpu_id n(6), mcod c(7), RecId i, date_in d, date_in2 d, date_out d, ;
   spos c(1), spos2 c(1), s_pol c(6), s_pol2 c(6), n_pol c(16), n_pol2 c(16), ;
   tip_d c(1), tip_d2 c(1), q c(2), fam c(25), im c(20), ot c(20), dr d, w n(1), ;
   ans_r c(3), reserv c(20))
 SELECT svppl
 INDEX on n_pol2 TAG n_pol2
 SET ORDER TO n_pol2
 
 m.ldat1 = GOMONTH(m.tdat1,-1)
 m.ldat2 = GOMONTH(m.ldat1,1)-1
 IF m.ldat1<{01.12.2012}
  m.lIsNeedperiod = .F.
 ELSE 
  m.lIsNeedperiod = .T.
 ENDIF 
 
 IF m.lIsNeedperiod
  lcperiod = STR(YEAR(m.ldat1),4)+PADL(MONTH(m.ldat1),2,'0')
  IF !fso.FolderExists(pbase+'\'+lcperiod)
   m.lIsNeedperiod = .F.
  ENDIF 
  IF m.lIsNeedperiod
   IF !fso.FileExists(pbase+'\'+lcperiod+'\aisoms.dbf')
    m.lIsNeedperiod = .F.
   ENDIF 
  ENDIF 
  IF m.lIsNeedperiod
   IF OpenFile(pbase+'\'+lcperiod+'\aisoms', 'lcais', 'shar', 'lpuid')>0
    IF USED('lcais')
     USE IN lcais
    ENDIF 
    m.lIsNeedperiod = .F.
   ELSE 
    m.lIsNeedperiod = .T.
   ENDIF 
  ENDIF 
  
 ENDIF 
 
 IF m.lIsNeedperiod
  MESSAGEBOX(CHR(13)+CHR(10)+'��������� ����������� ������� ��������!'+CHR(13)+CHR(10),0+64,'')
 ENDIF 
 
 SELECT aisoms
 SCAN 
  m.mcod  = mcod 
  m.lpuid = lpuid
  IF EMPTY(bname) AND m.lIsNeedperiod
   IF SEEK(m.lpuid, 'lcais') AND !EMPTY(lcais.bname)
    m.bname     = lcais.bname
    m.dname     = lcais.dname
    m.sent      = lcais.sent
    m.recieved  = lcais.recieved
    m.processed = lcais.processed
    m.cfrom     = lcais.cfrom
    m.cmessage  = lcais.cmessage
    m.paztot    = lcais.paztot
    m.pazin     = lcais.pazin
    m.pazbad    = lcais.pazbad
    m.pazdbl    = lcais.pazdbl
    m.erzsv4_id = lcais.erzsv4_id
    m.sv4st     = lcais.sv4st
    m.erzlpu_id = lcais.erzlpu_id
    m.lpust     = lcais.lpust
    m.issent    = lcais.issent
    IF FIELD('lcais.childs')='CHILDS'
     m.childs    = lcais.childs
    ELSE 
     m.childs = 0
    ENDIF 
    IF FIELD('lcais.adults')='ADULTS'
     m.adults    = lcais.adults
    ELSE 
     m.adults = 0
    ENDIF 
    
    IF !fso.FolderExists(pbase+'\'+lcperiod+'\'+m.mcod)
     LOOP 
    ENDIF 
    IF !fso.FileExists(pbase+'\'+lcperiod+'\'+m.mcod+'\people.dbf')
     LOOP 
    ENDIF 
    
    IF MESSAGEBOX(CHR(13)+CHR(10)+'��� �� ���������� ����� � ������� ������,'+;
     CHR(13)+CHR(10)+'�� ���������� � ����������.'+;
     CHR(13)+CHR(10)+'��������� ��� ���������� � ������� �����?'+CHR(13)+CHR(10),32+4, '')=6
    
     IF fso.FileExists(pbase+'\'+lcperiod+'\'+m.mcod+'\'+ALLTRIM(lcais.bname))
      fso.CopyFile(pbase+'\'+lcperiod+'\'+m.mcod+'\'+ALLTRIM(lcais.bname), ;
       pbase+'\'+gcperiod+'\'+m.mcod+'\'+ALLTRIM(lcais.bname))
     ENDIF      
     IF fso.FileExists(pbase+'\'+lcperiod+'\'+m.mcod+'\people.dbf')
      fso.CopyFile(pbase+'\'+lcperiod+'\'+m.mcod+'\people.dbf', ;
       pbase+'\'+gcperiod+'\'+m.mcod+'\people.dbf')
     ENDIF      
     IF fso.FileExists(pbase+'\'+lcperiod+'\'+m.mcod+'\people.cdx')
      fso.CopyFile(pbase+'\'+lcperiod+'\'+m.mcod+'\people.cdx', ;
       pbase+'\'+gcperiod+'\'+m.mcod+'\people.cdx')
     ENDIF      
     IF fso.FileExists(pbase+'\'+lcperiod+'\'+m.mcod+'\pr'+STR(m.lpuid,4)+m.qcod+'.zip')
      fso.CopyFile(pbase+'\'+lcperiod+'\'+m.mcod+'\'+'\pr'+STR(m.lpuid,4)+m.qcod+'.zip', ;
       pbase+'\'+gcperiod+'\'+m.mcod+'\'+'\pr'+STR(m.lpuid,4)+m.qcod+'.zip')
     ENDIF      
     
     REPLACE bname WITH m.bname, dname WITH m.dname, sent WITH m.sent,;
      recieved  WITH m.recieved, processed WITH m.processed, ; 
      cfrom WITH m.cfrom, cmessage WITH m.cmessage, paztot WITH m.paztot 
      
     REPLACE pazin WITH m.pazin, pazbad WITH m.pazbad, pazdbl WITH m.pazdbl,;
      erzsv4_id WITH m.erzsv4_id, sv4st WITH m.sv4st, erzlpu_id WITH m.erzlpu_id,;
      lpust WITH m.lpust, issent WITH m.issent, childs WITH m.childs, ;
      adults  WITH m.adults

    ENDIF 
    
   ELSE 
   LOOP 
   ENDIF 
  ENDIF 

  IF !fso.FolderExists(pbase+'\'+gcperiod+'\'+m.mcod)
   LOOP 
  ENDIF 
  IF !fso.FileExists(pbase+'\'+gcperiod+'\'+m.mcod+'\people.dbf')
   LOOP 
  ENDIF 
  IF OpenFile(pBase+'\'+gcPeriod+'\'+m.mcod+'\people.dbf', 'people', 'shar') > 0
   IF USED('people')
    USE IN people
    SELECT aisoms
   ENDIF 
   LOOP 
  ENDIF 
  IF RECCOUNT('people')=0
   USE IN people 
   SELECT aisoms
   LOOP 
  ENDIF 
  
  WAIT m.mcod+'...' WINDOW NOWAIT 
  
  SELECT people
  m.pazbad = 0
  m.pazdbl = 0

  SCAN 
   IF !EMPTY(date_out)
    LOOP 
   ENDIF 
   IF c_err!='OK'
    LOOP 
   ENDIF 
   

   SCATTER MEMVAR
   RELEASE c_err
   m.lpu_id = IIF(SEEK(m.mcod, 'sprlpu'), sprlpu.lpu_id, 0)
   
   IF m.n_pol2 = '7797799791000721'
    SET STEP ON 
   ENDIF 
   
   IF SEEK(m.n_pol2, 'svppl')
    DO CASE 
     CASE spos = '1' AND svppl.spos = '1' && ��� �� ����������
*      IF date_in<=svppl.date_in && �������� � ��������� �� �������� 06.03.2014 �� ������� ���� ���������!
      IF date_in>svppl.date_in
       m.dmcod = svppl.mcod
       IF m.mcod=m.dmcod
*        SET STEP ON 
       ENDIF 
       m.drecid = svppl.recid
       IF fso.FolderExists(pbase+'\'+gcperiod+'\'+m.dmcod)
        IF fso.FileExists(pbase+'\'+gcperiod+'\'+m.dmcod+'\people.dbf')
         IF OpenFile(pbase+'\'+gcperiod+'\'+m.dmcod+'\people', 'ppeople', 'shar', 'recid')>0 
          IF USED('ppeople')
           USE IN ppeople
          ENDIF 
         ELSE 
          m.pazbad = m.pazbad + 1
          m.pazdbl = m.pazdbl + 1
          SELECT ppeople
          =SEEK(svppl.recid)
          REPLACE c_err WITH 'FP', ;
           lpu_id WITH m.lpu_id, spos2 WITH m.spos, date_in2 WITH m.date_in
          USE IN ppeople
          DELETE IN svppl
          INSERT INTO svppl FROM MEMVAR 
          SELECT people
         ENDIF 
        ENDIF 
       ENDIF 
      ELSE 
       m.pazbad = m.pazbad + 1
       m.pazdbl = m.pazdbl + 1
       REPLACE c_err WITH 'FP', ;
        lpu_id WITH svppl.lpu_id, spos2 WITH svppl.spos, date_in2 WITH svppl.date_in
      ENDIF 

     CASE spos = '1' AND svppl.spos = '2'
      m.pazbad = m.pazbad + 1
      m.pazdbl = m.pazdbl + 1
      REPLACE c_err WITH 'FP', ;
       lpu_id WITH svppl.lpu_id, spos2 WITH svppl.spos, date_in2 WITH svppl.date_in

     CASE spos = '2' AND svppl.spos = '1'
       m.dmcod = svppl.mcod
       IF m.mcod=m.dmcod
*        SET STEP ON 
       ENDIF 
       m.drecid = svppl.recid
       IF fso.FolderExists(pbase+'\'+gcperiod+'\'+m.dmcod)
        IF fso.FileExists(pbase+'\'+gcperiod+'\'+m.dmcod+'\people.dbf')
         IF OpenFile(pbase+'\'+gcperiod+'\'+m.dmcod+'\people', 'ppeople', 'shar', 'recid')>0
          IF USED('ppeople')
           USE IN ppeople
          ENDIF 
         ELSE 
          m.pazbad = m.pazbad + 1
          m.pazdbl = m.pazdbl + 1
          SELECT ppeople
          =SEEK(svppl.recid)
          REPLACE c_err WITH 'FP', ;
           lpu_id WITH m.lpu_id, spos2 WITH m.spos, date_in2 WITH m.date_in
          USE IN ppeople
          SELECT people
          DELETE IN svppl
          INSERT INTO svppl FROM MEMVAR 
         ENDIF 
        ENDIF 
       ENDIF 

     CASE spos = '2' AND svppl.spos = '2' && ��� �� ���������
      IF date_in<=svppl.date_in
       IF YEAR(date_in)=YEAR(svppl.date_in)
       m.dmcod = svppl.mcod
       m.drecid = svppl.recid
       IF fso.FolderExists(pbase+'\'+gcperiod+'\'+m.dmcod)
        IF fso.FileExists(pbase+'\'+gcperiod+'\'+m.dmcod+'\people.dbf')
         IF OpenFile(pbase+'\'+gcperiod+'\'+m.dmcod+'\people', 'ppeople', 'shar', 'recid')>0
          IF USED('ppeople')
           USE IN ppeople
          ENDIF 
         ELSE 
          m.pazbad = m.pazbad + 1
          m.pazdbl = m.pazdbl + 1
          SELECT ppeople
          =SEEK(svppl.recid)
          REPLACE c_err WITH 'FP', ;
           lpu_id WITH m.lpu_id, spos2 WITH m.spos, date_in2 WITH m.date_in
          USE IN ppeople
          SELECT people
          DELETE IN svppl
          INSERT INTO svppl FROM MEMVAR 
         ENDIF 
        ENDIF 
       ENDIF 
       ELSE 
        m.pazbad = m.pazbad + 1
        m.pazdbl = m.pazdbl + 1
        REPLACE c_err WITH 'FP', ;
         lpu_id WITH svppl.lpu_id, spos2 WITH svppl.spos, date_in2 WITH svppl.date_in
       ENDIF 
      ELSE 
       IF YEAR(date_in)=YEAR(svppl.date_in)
        m.pazbad = m.pazbad + 1
        m.pazdbl = m.pazdbl + 1
        REPLACE c_err WITH 'FP', ;
         lpu_id WITH svppl.lpu_id, spos2 WITH svppl.spos, date_in2 WITH svppl.date_in
       ELSE 
        m.dmcod = svppl.mcod
        m.drecid = svppl.recid
        IF fso.FolderExists(pbase+'\'+gcperiod+'\'+m.dmcod)
         IF fso.FileExists(pbase+'\'+gcperiod+'\'+m.dmcod+'\people.dbf')
          IF OpenFile(pbase+'\'+gcperiod+'\'+m.dmcod+'\people', 'ppeople', 'shar', 'recid')>0
           IF USED('ppeople')
            USE IN ppeople
           ENDIF 
          ELSE 
           m.pazbad = m.pazbad + 1
           m.pazdbl = m.pazdbl + 1
           SELECT ppeople
           =SEEK(svppl.recid)
           REPLACE c_err WITH 'FP', ;
            lpu_id WITH m.lpu_id, spos2 WITH m.spos, date_in2 WITH m.date_in
           USE IN ppeople
           SELECT people
           DELETE IN svppl
           INSERT INTO svppl FROM MEMVAR 
          ENDIF 
         ENDIF 
        ENDIF 
       ENDIF 
      ENDIF 
     OTHERWISE 
    ENDCASE 
   ELSE
    INSERT INTO svppl FROM MEMVAR 
   ENDIF 
   
  ENDSCAN 
  USE IN people 
  
  WAIT CLEAR 
  
  SELECT aisoms
  m.opazbad = pazbad
  m.opazdbl = pazdbl
  REPLACE pazbad WITH m.opazbad+m.pazbad, pazdbl WITH m.opazdbl+m.pazdbl

 ENDSCAN 
 USE 
 USE IN sprlpu
 
 SELECT svppl 
 IF RECCOUNT('svppl')<=0
  USE IN svppl
  =MESSAGEBOX(CHR(13)+CHR(10)+'�� ���������� �� ������ ��������!'+CHR(13)+CHR(10),0+64,'')
  IF USED('lcais')
   USE IN lcais
  ENDIF 
  RETURN 
 ENDIF  
 
 svpeople = pbase+'\'+gcperiod+'\people'
 IF fso.FileExists(svpeople+'.dbf')
  fso.DeleteFile(svpeople+'.dbf')
 ENDIF 
 CREATE TABLE (svpeople) ;
  (lpu_id n(6), mcod c(7), RecId i, date_in d, date_in2 d, date_out d, ;
   spos c(1), spos2 c(1), s_pol c(6), s_pol2 c(6), n_pol c(16), n_pol2 c(16), ;
   tip_d c(1), tip_d2 c(1), q c(2), fam c(25), im c(20), ot c(20), dr d, w n(1), ;
   ans_r c(3), reserv c(20))
 
 SELECT svppl 
 SCAN FOR !DELETED()
  SCATTER MEMVAR 
  INSERT INTO people FROM MEMVAR 
 ENDSCAN 
 USE IN svppl
 USE IN people

 IF USED('lcais')
  USE IN lcais
 ENDIF 

 =MESSAGEBOX(CHR(13)+CHR(10)+'������� ������� ������!'+CHR(13)+CHR(10),0+64,'')
 
RETURN 