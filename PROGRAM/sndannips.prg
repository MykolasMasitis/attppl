PROCEDURE SndAnnIpS
 IF MESSAGEBOX('��������� ��������������'+CHR(13)+CHR(10)+'�������������� �������?',4+32,'')=7
  RETURN 
 ENDIF
 IF !fso.FolderExists(pAisOms+'\OMS\OUTPUT')
  MESSAGEBOX('���������� '+CHR(13)+CHR(10)+;
   UPPER(pAisOms)+'\OMS\OUTPUT'+CHR(13)+CHR(10)+;
   '�� ��������!',0+64,'')
 ENDIF 
 
 pAnnDir = fso.GetParentFolderName(pBase)+'\ANNDIR'
 IF !fso.FileExists(pAnnDir+'\ips.dbf')
  RETURN 
 ENDIF 
 IF OpenFile(pAnnDir+'\ips', 'ips', 'shar')>0
  IF USED('ips')
   USE IN ips
  ENDIF 
  RETURN 
 ENDIF 
 
 SELECT ips
 m.iptosend = 0 
 SCAN 
  IF status=1 && ���������. ���� ��?
   m.un_id = un_id
   m.bfile = 'b'+m.un_id
   IF !fso.FileExists(pAnnDir+'\OUTPUT\'+m.bfile)
    REPLACE status WITH 2
   ENDIF 
   LOOP 
  ENDIF 
  IF status!=0
   LOOP 
  ENDIF 

  m.lcname  = ALLTRIM(name)
  m.flname  = 'D'+UPPER(m.qcod)+m.lcname
  m.annpath = pAnnDir+'\OUTPUT\'

  m.zipfile = m.annpath+m.flname+'.zip'
  IF !fso.FileExists(m.zipfile)
   LOOP 
  ENDIF 

  m.un_id = SYS(3)
  m.bfile = 'b' + m.un_id
  m.tfile = 't' + m.un_id
  m.dfile = 'd' + m.un_id
  m.mmid  = m.un_id+'.OMS@'+m.qmail
  
  m.period = '20'+SUBSTR(period,3,2)+SUBSTR(period,1,2)
  
  m.subject = 'OMS#'+m.period+'###gn'
  
  m.cTo   = 'oms@mgf.msk.oms'

  poi = fso.CreateTextFile(pAisOms+'\OMS\OUTPUT\'+m.tfile)

  poi.WriteLine('To: '+m.cTO)
  poi.WriteLine('Message-Id: ' + m.mmid)
  poi.WriteLine('Subject: ' + m.subject)
  poi.WriteLine('Content-Type: multipart/mixed')
  poi.WriteLine('Attachment: '+m.dfile+' '+STRTRAN(m.flname,'D','PR')+'.ZIP')

  poi.Close

  fso.CopyFile(m.annpath+m.flname+'.zip', pAisOms+'\oms\output\'+m.dfile)
  fso.CopyFile(pAisOms+'\OMS\OUTPUT\'+m.tfile, pAisOms+'\OMS\OUTPUT\'+m.bfile)
  fso.DeleteFile(pAisOms+'\OMS\OUTPUT\'+m.tfile)
  
  m.iptosend = m.iptosend + 1
  
  REPLACE status WITH 1, un_id WITH m.un_id, formed WITH DATETIME()
  
 ENDSCAN 
 USE IN ips 
 
 MESSAGEBOX('��������� ���������!',0+64,'')
 
RETURN 