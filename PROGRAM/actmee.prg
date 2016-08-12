PROCEDURE ActMee(para1, para2)
 m.lcmcod  = para1
 m.lnlpuid = para2
 
 IF MESSAGEBOX('—‘Œ–Ã»–Œ¬¿“‹ ¡À¿Õ  ¿ “¿ Ã››?'+CHR(13)+CHR(10)+'(SPOS=2)', 4+32, m.lcmcod)=7
  RETURN 
 ENDIF 
 IF !fso.FileExists(ptempl+'\spos2.xls')
  MESSAGEBOX('Œ“—”“—“¬”≈“ ÿ¿¡ÀŒÕ ƒŒ ”Ã≈Õ“¿ SPOS2.XLS!',0+15,'')
  RETURN 
 ENDIF 
 
 IF m.tdat1 <= {01.04.2015}
  m.exp_dat1 = m.tdat1
  m.lcperiod = m.gcperiod
 ELSE 
  m.exp_dat1 = {01.04.2015}
  m.lcperiod = '201504'  && {01.04.2015}
 ENDIF 
 
 IF !fso.FolderExists(m.pbase+'\'+m.lcperiod)
  RETURN 
 ENDIF 
 IF !fso.FolderExists(m.pbase+'\'+m.lcperiod+'\'+m.lcmcod)
  RETURN
 ENDIF 
 IF !fso.FileExists(m.pbase+'\'+m.lcperiod+'\'+m.lcmcod+'\people.dbf')
  RETURN 
 ENDIF 
 IF OpenFile(m.pbase+'\'+m.lcperiod+'\'+m.lcmcod+'\people', 'people', 'shar')>0
  IF USED('people')
   USE IN people 
  ENDIF 
  SELECT aisoms
  RETURN 
 ENDIF 
  
 CREATE CURSOR curdata (nrec n(6), date_in d, enp c(16), fam c(25), im c(20), ot c(20), dr d)
  
 SELECT people 
 SCAN 
  m.c_err = c_err
  m.spos  = spos
  IF !(m.c_err='OK' AND spos='2')
   LOOP 
  ENDIF 
   
  m.date_in = date_in
  m.enp     = n_pol
  m.fam     = fam 
  m.im      = im
  m.ot      = ot
  m.dr      = dr
   
  INSERT INTO curdata FROM MEMVAR 
 ENDSCAN 
 USE IN people 
  
 
 IF m.tdat1>{01.04.2015}
  m.ldat1 = {01.04.2015}
  DO WHILE m.ldat1<=m.tdat1
   m.ldat1 = GOMONTH(m.ldat1,1)
   m.lcperiod = STR(YEAR(m.ldat1),4)+PADL(MONTH(m.ldat1),2,'0')
   IF !fso.FolderExists(m.pbase+'\'+m.lcperiod)
    LOOP 
   ENDIF 
   IF !fso.FolderExists(m.pbase+'\'+m.lcperiod+'\'+m.lcmcod)
    LOOP 
   ENDIF 
   IF !fso.FileExists(m.pbase+'\'+m.lcperiod+'\'+m.lcmcod+'\people.dbf')
    LOOP 
   ENDIF 
   IF OpenFile(m.pbase+'\'+m.lcperiod+'\'+m.lcmcod+'\people', 'people', 'shar')>0
    IF USED('people')
     USE IN people 
    ENDIF 
    LOOP 
   ENDIF 

   SELECT people 
   SCAN 
    m.c_err = c_err
    m.spos  = spos
    IF !(m.c_err='OK' AND spos='2')
     LOOP 
    ENDIF 
   
    m.date_in = date_in
    m.enp     = n_pol
    m.fam     = fam 
    m.im      = im
    m.ot      = ot
    m.dr      = dr
   
    INSERT INTO curdata FROM MEMVAR 
   ENDSCAN 
   USE IN people 
   
  ENDDO 
  
 ENDIF 
 
 SELECT curdata
 REPLACE ALL nrec WITH RECNO()

 m.lcTmpName = pTempl + "\actmeesv.xls"
 m.lcRepName = pbase+'\'+m.gcperiod+'\'+m.lcmcod+'\actmee.xls'
 m.lpuname = IIF(SEEK(m.lnlpuid, 'sprlpu'), sprlpu.fullname, '')
 m.recs = RECCOUNT('curdata')
 m.d_akt = DTOC(DATE())
 m.exp_dat2 = m.tdat2
 
 
 m.llResult = X_Report(m.lcTmpName, m.lcRepName, .t.)
 
 USE IN curdata 
 
 SELECT aisoms

RETURN 