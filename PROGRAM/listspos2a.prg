PROCEDURE ListSpos2a(para1, para2, para3)
 m.lcmcod   = para1
 m.lnlpuid  = para2
 m.IsVisible = para3
 
* IF MESSAGEBOX('ÑÔÎÐÌÈÐÎÂÀÒÜ ÑÏÈÑÎÊ ÏÐÈÊÐÅÏËÅÍÍÛÕ?'+CHR(13)+CHR(10)+'(SPOS=2)', 4+32, m.lcmcod)=7
*  RETURN 
* ENDIF 
 IF !fso.FileExists(ptempl+'\spos2.xls')
  MESSAGEBOX('ÎÒÑÓÒÑÒÂÓÅÒ ØÀÁËÎÍ ÄÎÊÓÌÅÍÒÀ SPOS2.XLS!',0+15,'')
  RETURN 
 ENDIF 
 
 IF !fso.FolderExists(pbase+'\'+m.gcperiod+'\'+m.lcmcod)
  RETURN 
 ENDIF 
 IF !fso.FileExists(pbase+'\'+m.gcperiod+'\'+m.lcmcod+'\allpeople.dbf')
  RETURN 
 ENDIF 
 
 IF OpenFile(pbase+'\'+m.gcperiod+'\'+m.lcmcod+'\allpeople', 'people', 'shar')>0
  IF USED('people')
   USE IN people 
  ENDIF 
  SELECT aisoms
  RETURN 
 ENDIF 

 SELECT 000000 as nrec, date_in, n_pol as enp, fam, im,ot,dr FROM people INTO CURSOR curdata READWRITE 
 USE IN people
 SELECT curdata
 REPLACE ALL nrec WITH RECNO()

 m.lcTmpName = pTempl + '\spos2.xls'
 m.lcRepName = pOut+'\sp2_'+m.lcmcod+'_'+m.gcperiod+'.xls'
 m.lpuname = IIF(SEEK(m.lnlpuid, 'sprlpu'), sprlpu.fullname, '')
 m.recs = RECCOUNT('curdata')
 
 m.llResult = X_Report(m.lcTmpName, m.lcRepName, m.IsVisible)
 
 USE IN curdata 
 
 SELECT aisoms

RETURN 