FUNCTION SendToFoms
 prfile = 'pr'+LOWER(m.qcod)+PADL(tMonth,2,'0')+RIGHT(STR(tYear,4),2)+'.dbf'
 IF !fso.FileExists(pbase+'\'+gcperiod+'\'+prfile)
  MESSAGEBOX(CHR(13)+CHR(10)+'тюик'++' ме ятнплхпнбюм!'+CHR(13)+CHR(10),0+16,'')
  RETURN 
 ENDIF 

 actname = '\vv'+m.qcod+PADL(tMonth,2,'0')+RIGHT(STR(tYear,4),2)+'.pdf'
 IF !fso.FileExists(pbase+'\'+gcperiod+'\'+actname)
  MESSAGEBOX(CHR(13)+CHR(10)+'ябндмши юйр ябепйх ме ятнплхпнбюм!'+CHR(13)+CHR(10),0+16,'')
  RETURN 
 ENDIF 
 
 m.mmyy = PADL(tmonth,2,'0')+RIGHT(STR(tyear,4),2)
 ZipFile = 'pr'+m.qcod+m.mmyy+'.zip'
 IF fso.FileExists(pbase+'\'+gcperiod+'\'+ZipFile)
  fso.DeleteFile(pbase+'\'+gcperiod+'\'+ZipFile)
 ENDIF 
 
 ZipOpen(pbase+'\'+gcperiod+'\'+ZipFile)
 ZipFile(pbase+'\'+gcperiod+'\'+prfile)
 ZipFile(pbase+'\'+gcperiod+'\'+actname)
 ZipClose()
 
 IF !fso.FileExists(pbase+'\'+gcperiod+'\'+ZipFile)
  MESSAGEBOX(CHR(13)+CHR(10)+'мебнглнфмн янгдюрэ юпухб!'+CHR(13)+CHR(10),0+16,'')
  RETURN 
 ENDIF 
 
 m.un_id    = SYS(3)
 m.bansfile = 'b' + m.un_id
 m.tansfile = 't' + m.un_id
 m.dfile    = 'd' + m.un_id
 m.mmid     = m.un_id+'.'+m.gcUser+'@'+m.qmail
 m.csubj    = 'OMS#'+gcPeriod+'###gp'

 poi = fso.CreateTextFile(pbase+'\'+gcperiod + '\' + m.tansfile)

 poi.WriteLine('To: oms@mgf.msk.oms')
 poi.WriteLine('Message-Id: ' + m.mmid)
 poi.WriteLine('Subject: ' + m.csubj)
 poi.WriteLine('Content-Type: multipart/mixed')
 poi.WriteLine('Attachment: '+m.dfile+' '+ZipFile)
 
 poi.Close
 
 fso.CopyFile(pbase+'\'+gcperiod+'\'+ZipFile, pAisOms+'\oms\output\'+m.dfile)
 fso.CopyFile(pbase+'\'+gcperiod+'\'+m.tansfile, pAisOms+'\oms\output\'+m.bansfile)

 MESSAGEBOX('тюикш нропюбкемш!',0+64,'')

RETURN  