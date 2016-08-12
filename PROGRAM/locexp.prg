PROCEDURE locexp(lcDir, mcod, IsOK)
 PRIVATE lcdir,mcod,isok, pazbad, pazdbl, adults, childs
 
 m.lpuid = lpuid
 m.lpust = lpust
 
 m.pazbad = 0
 m.pazdbl = 0
 m.adults = 0
 m.childs = 0

 IF !fso.FolderExists(lcDir)
  RETURN .F.
 ENDIF 

 IF !fso.FileExists(lcDir + '\People.dbf')
  RETURN .F.
 ENDIF 
 IF !fso.FileExists(lcDir + '\ans_sv4.dbf')
  RETURN .F.
 ENDIF 

 IF OpenFile("&lcDir\People", "People", "SHARE")>0
  RETURN .F.
 ENDIF 
 IF OpenFile("&lcDir\ans_sv4", "answer", "excl")>0
  USE IN people 
  RETURN .F.
 ENDIF 

 SELECT Answer 
 DELETE TAG ALL 
 INDEX ON RecId TAG RecId
 SET ORDER TO RecId
 
 IF RECCOUNT('people')<=0
  USE IN people 
  USE IN answer
  RETURN .F.
 ENDIF 
 
 m.mcod = mcod
 WAIT m.mcod WINDOW NOWAIT 
 
 m.VzGr = 0
* m.VzGr = INT(VAL(SUBSTR(mcod,2,1)))
 m.tippr = IIF(SEEK(m.lpuid, 'tiplpu'), tiplpu.pr, 'В')
 DO CASE 
  CASE m.tippr='В'
   m.VzGr = 1
  CASE m.tippr='Д'
   m.VzGr = 2
  CASE m.tippr='С'
   m.VzGr = 3
  OTHERWISE 
 ENDCASE 
 
 SELECT n_pol, date_out DISTINCT FROM people WHERE !EMPTY(date_out) ORDER BY n_pol ;
  INTO CURSOR outppl

 m.isfff = 0
 IF RECCOUNT('outppl')>0
  SET ORDER TO n_pol IN people
  SELECT outppl
  SCAN 
   m.nnn = n_pol
   m.date_out = date_out
   IF SEEK(m.nnn, 'people')
    SELECT people
    DO WHILE n_pol=m.nnn
     IF !EMPTY(date_out)
      SKIP 
      LOOP 
     ENDIF 
     m.isfff = m.isfff + 1
     REPLACE date_out WITH m.date_out
     SKIP 
    ENDDO 
    SELECT outppl
   ENDIF 
  ENDSCAN 
  IF m.isfff>0
*   MESSAGEBOX(STR(m.isfff,2),0+64,m.mcod)
  ENDIF 
 ENDIF 

 USE IN outppl
 SELECT people
 SET ORDER TO 
 SET RELATION TO PADL(recid,6,'0') INTO Answer
 
 SCAN FOR EMPTY(date_out) AND EMPTY(c_err)
  m.dr      = dr
  m.date_in = date_in
  m.date_pr = IIF(!EMPTY(answer.date_pr), answer.date_pr, {01.01.1980})
  m.reserv  = LOWER(ALLTRIM(reserv))
*  m.vozr = m.tyear - YEAR(m.dr)
  m.vozr = (m.date_in - m.dr)/365.25
  DO CASE 
   CASE spos='1'
    m.pazbad = m.pazbad + 1
    REPLACE c_err WITH 'FL'

   CASE spos='2' AND m.date_in<GOMONTH(m.tdat1,-1)
    m.pazbad = m.pazbad + 1
    REPLACE c_err WITH 'F9'

   CASE m.date_in - m.date_pr<365 AND (answer.tip_pr!=1 AND m.reserv!='z')
    m.pazbad = m.pazbad + 1
    REPLACE c_err WITH 'F9'

*   CASE (m.VzGr=2 AND m.vozr>20) OR (m.VzGr=1 AND m.vozr<15) && 
   CASE (m.VzGr=2 AND m.vozr>=18) OR (m.VzGr=1 AND m.vozr<18) && 
    m.pazbad = m.pazbad + 1
    REPLACE c_err WITH 'F0'
   
   CASE INLIST(ans_r,'***','311','331','411','431') && Некорректный запрос
    m.pazbad = m.pazbad + 1
    REPLACE c_err WITH 'F1'
   
   CASE (tip_d='1' AND tip_d2='С' AND n_pol!=n_pol2) OR ;
    (tip_d='3' AND tip_d2='П' AND n_pol!=n_pol2)
    m.pazbad = m.pazbad + 1
    REPLACE c_err WITH 'F4'

   CASE INLIST(ans_r,'13*','23*','33*','43*','130','231','331','431','210','310') && Полис недействителен
    m.pazbad = m.pazbad + 1
    REPLACE c_err WITH 'F5'

   CASE INLIST(ans_r,'0*0','000') && Полис не зарегистрирован
    m.pazbad = m.pazbad + 1
    REPLACE c_err WITH 'F6'
   
   CASE !EMPTY(ans_r) AND q!=m.qcod && Не так компашка
    m.pazbad = m.pazbad + 1
    REPLACE c_err WITH 'F7'
    
   CASE !EMPTY(ans_r) AND lpu_id = m.lpuid && Повтор прикрепления!
    m.childs = m.childs + IIF(m.vozr<18,1,0)
    m.adults = m.adults + IIF(m.vozr>=18,1,0)
    REPLACE c_err WITH 'OK'

   OTHERWISE 
    IF !EMPTY(ans_r)
     m.childs = m.childs + IIF(m.vozr<18,1,0)
     m.adults = m.adults + IIF(m.vozr>=18,1,0)
     REPLACE c_err WITH 'OK'
    ENDIF 
  ENDCASE 
 ENDSCAN 

 SELECT * FROM people WHERE n_pol2 in ;
 (SELECT n_pol2 FROM people WHERE !EMPTY(n_pol2) AND EMPTY(date_out) GROUP BY n_pol2 HAVING ;
  COUNT(*)>1) ORDER BY n_pol2, tip_d desc, s_pol DESC  INTO CURSOR ddbl READWRITE 
 
 IF _tally=0
  USE IN ddbl
 ELSE 
  SELECT ddbl
  GO TOP 
  m.npol2 = n_pol2
  SKIP 
  DO WHILE !EOF()
   IF n_pol2!=m.npol2
    m.npol2 = n_pol2
   ELSE 
    m.recid = recid 
    IF c_err=='OK'
     m.pazbad = m.pazbad + 1
     m.pazdbl = m.pazdbl + 1
     REPLACE c_err WITH 'F8' IN people FOR recid = m.recid
     m.childs = m.childs - IIF(m.vozr<18,1,0)
     m.adults = m.adults - IIF(m.vozr>=18,1,0)
    ENDIF 
   ENDIF 
   SKIP 
  ENDDO 
  USE
 ENDIF 

 SELECT people 
 SET RELATION OFF INTO answer
 USE 
 SELECT Answer
 SET ORDER TO 
 DELETE TAG ALL 
 USE IN Answer

 SELECT aisoms
 
 REPLACE pazbad WITH m.pazbad, pazdbl WITH m.pazdbl, ;
  adults WITH m.adults, childs WITH m.childs
 
 MailView.pazbad = MailView.pazbad + m.pazbad
 MailView.pazdbl = MailView.pazdbl + m.pazdbl
 MailView.adults = MailView.adults + m.adults
 MailView.childs = MailView.childs + m.childs
 
 MailView.Refresh
 
 WAIT CLEAR 

RETURN 