FUNCTION SendToLpu(lcPath, cmcod, clpuid)

 ctrl='etrl'+m.qcod
 IF !fso.FileExists(lcpath+'\'+ctrl+'.dbf')
  MESSAGEBOX(CHR(13)+CHR(10)+'тюик ньханй ме ятнплхпнбюм!'+CHR(13)+CHR(10),0+16,cmcod)
  RETURN 
 ENDIF 

 actsv = 'sv'+LOWER(m.qcod)+PADL(tMonth,2,'0')+RIGHT(STR(tYear,4),2)
 IF !fso.FileExists(lcpath+'\'+actsv+'.pdf')
  MESSAGEBOX(CHR(13)+CHR(10)+'юйр ябепйх ме ятнплхпнбюм!'+CHR(13)+CHR(10),0+16,cmcod)
  RETURN 
 ENDIF 
 
 ZipFile = 'AR'+clpuid+m.qcod+'.zip'
 IF fso.FileExists(lcPath+'\'+ZipFile)
  fso.DeleteFile(lcPath+'\'+ZipFile)
 ENDIF 
 
 ZipOpen(lcPath+'\'+ZipFile)
 ZipFile(lcPath+'\'+ctrl+'.dbf')
 ZipFile(lcPath+'\'+actsv+'.pdf')
 ZipClose()
 
 IF !fso.FileExists(lcPath+'\'+ZipFile)
*  MESSAGEBOX(CHR(13)+CHR(10)+'мебнглнфмн янгдюрэ юпухб!'+CHR(13)+CHR(10),0+16,'')
  RETURN 
 ENDIF 
 
 m.IsEmiac = IsEmiac
 IF m.IsEmiac
  m.cto = 'oms@spuemias.msk.oms'
 ELSE 
  m.cto      = IIF(SEEK(INT(VAL(clpuid)), 'sprabo', 'lpu_id'), 'oms@'+sprabo.abn_name, '')
 ENDIF 

 m.un_id    = SYS(3)
 m.bansfile = 'b' + m.un_id
 m.tansfile = 't' + m.un_id
 m.dfile    = 'd' + m.un_id
 m.mmid     = m.un_id+'.'+m.gcUser+'@'+m.qmail
 m.csubj    = 'OMS#'+gcPeriod+'###gp'

 poi = fso.CreateTextFile(lcPath + '\' + m.tansfile)

 poi.WriteLine('To: '+m.cTO)
 poi.WriteLine('Message-Id: ' + m.mmid)
 poi.WriteLine('Subject: ' + m.csubj)
 poi.WriteLine('Content-Type: multipart/mixed')
 poi.WriteLine('Resent-Message-Id: '+ALLTRIM(cmessage))
 poi.WriteLine('Attachment: '+m.dfile+' '+ZipFile)
 
 poi.Close
 
 fso.CopyFile(lcPath+'\'+ZipFile, pAisOms+'\oms\output\'+m.dfile)
 fso.CopyFile(lcPath+'\'+m.tansfile, pAisOms+'\oms\output\'+m.bansfile)

 REPLACE issent WITH .t. 
* MESSAGEBOX('тюикш нропюбкемш!',0+64,'')

RETURN  