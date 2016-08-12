PROCEDURE SendToLpu2(lcmcod, lnlpuid)

 m.lcpath = pbase+'\'+m.gcperiod+'\'+m.lcmcod
 
 IF !fso.FileExists(m.lcpath+'\spos2.xls')
  MESSAGEBOX('По выбранному ЛПУ'+CHR(13)+CHR(10)+;
   'список не сформирован!'+CHR(13)+CHR(10),0+16,'spos2.xls')
  RETURN 
 ENDIF 
 
 ZipReq = 'rq2'+m.lcmcod
 IF fso.FileExists(m.lcpath+'\'+ZipReq+'.zip')
  fso.DeleteFile(m.lcpath+'\'+ZipReq+'.zip')
 ENDIF 
 
 ZipOpen(m.lcpath+'\'+ZipReq+'.zip')
 ZipFile(m.lcpath+'\spos2.xls')
 ZipClose()
 
* m.cTo   = cfrom
 m.cTo   = 'usr010@'+IIF(SEEK(m.lnlpuid, 'sprabo', 'lpu_id'), ALLTRIM(sprabo.abn_name), cfrom)
 m.un_id = SYS(3)
 m.bfile = 'brq2'+m.un_id
 m.tfile = 'trq2'+m.un_id
 m.dfile = 'drq2' + m.un_id
 m.mmid  = m.un_id+'.OMS@'+m.qmail
 m.csubj = 'Запрос заявлений на прикрепление ('+ALLTRIM(m.qname)+')'

 poi = fso.CreateTextFile(m.lcpath + '\' + m.tfile)

 poi.WriteLine('To: '+m.cTO)
 poi.WriteLine('Message-Id: ' + m.mmid)
 poi.WriteLine('Subject: ' + m.csubj)
 poi.WriteLine('Content-Type: multipart/mixed')
 poi.WriteLine('Attachment: '+m.dfile+' '+ZipReq+'.zip')

 poi.Close

 fso.CopyFile(m.lcpath+'\'+ZipReq+'.zip', pAisOms+'\usr010\output\'+m.dfile)
 fso.CopyFile(m.lcpath+'\'+m.tfile, pAisOms+'\usr010\output\'+m.bfile)
 
 fso.CopyFile(m.lcpath+'\'+m.tfile, m.lcpath+'\'+m.bfile)
 fso.DeleteFile(m.lcpath+'\'+m.tfile)
 
 MESSAGEBOX('ЗАПРОС ОТПРАВЛЕН!',0+64, '')

RETURN 
 
