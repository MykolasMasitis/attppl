PROCEDURE ActMeeA(para1, para2, para3)
 m.lcmcod  = para1
 m.lnlpuid = para2
 m.IsVisible = para3
 
* IF MESSAGEBOX('—‘Œ–Ã»–Œ¬¿“‹ ¡À¿Õ  ¿ “¿ Ã››?'+CHR(13)+CHR(10)+'(SPOS=2)', 4+32, m.lcmcod)=7
*  RETURN 
* ENDIF 
 IF !fso.FileExists(ptempl+'\actmeesv.xls')
  MESSAGEBOX('Œ“—”“—“¬”≈“ ÿ¿¡ÀŒÕ ƒŒ ”Ã≈Õ“¿ ACTMEESV.XLS!',0+15,'')
  RETURN 
 ENDIF 
 
 IF !fso.FolderExists(pbase+'\'+m.gcperiod+'\'+m.lcmcod)
  RETURN 
 ENDIF 
 IF !fso.FileExists(pbase+'\'+m.gcperiod+'\'+m.lcmcod+'\allpeople.dbf')
  RETURN 
 ENDIF 
 
 IF m.tdat1 <= {01.04.2015}
  m.exp_dat1 = m.tdat1
  m.lcperiod = m.gcperiod
 ELSE 
  m.exp_dat1 = {01.04.2015}
  m.lcperiod = '201504'  && {01.04.2015}
 ENDIF 

 IF OpenFile(pbase+'\'+m.gcperiod+'\'+m.lcmcod+'\allpeople', 'people', 'shar')>0
  IF USED('people')
   USE IN people 
  ENDIF 
  SELECT aisoms
  RETURN 
 ENDIF 

 m.ynorm = 0 
 IF OpenFile(pCommon+'\pnyear', 'pnyear', 'shar', 'period')>0
  IF USED('pnyear')
   USE IN pnyear
  ENDIF 
 ELSE 
  SELECT pnyear
  IF SEEK(STR(tYear,4), 'pnyear')
   m.ynorm = pnyear.pnorm
  ENDIF 
  IF USED('pnyear')
   USE IN pnyear
  ENDIF 
 ENDIF 

 SELECT 000000 as nrec, date_in, n_pol as enp, fam, im,ot,dr, isdoc, 00000.00 as straf FROM people ;
  INTO CURSOR curdata READWRITE 
 USE IN people
 SELECT curdata
 REPLACE ALL nrec WITH RECNO()
 REPLACE ALL straf WITH IIF(IsDoc=.t., 0, m.ynorm)
 SUM straf TO m.totstraf

 m.lcTmpName = pTempl + "\actmeesv.xls"
 m.lcRepName = pOut+'\Act_'+m.lcmcod+'_'+m.gcperiod+'.xls'
 m.lpuname = IIF(SEEK(m.lnlpuid, 'sprlpu'), sprlpu.fullname, '')
 m.recs = RECCOUNT('curdata')
 m.d_akt = DTOC(DATE())
 m.exp_dat2 = m.tdat2
 
 
 m.llResult = X_Report(m.lcTmpName, m.lcRepName, m.IsVisible)
 
 USE IN curdata 
 
 SELECT aisoms

RETURN 