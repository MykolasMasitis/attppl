PROCEDURE CorStruct
 IF MESSAGEBOX('ÂÛ ÕÎÒÈÒÅ ÏÐÎÂÅÑÒÈ '+CHR(13)+CHR(10)+;
               'ÊÎÐÐÅÊÒÈÐÎÂÊÓ ÑÒÐÓÊÒÓÐÛ ÁÄ?!'+CHR(13)+CHR(10)+;
               '',4+48,'') != 6
  RETURN 
 ENDIF 

 IF MESSAGEBOX('ÂÛ ÀÁÑÎËÞÒÍÎ ÓÂÅÐÅÍÛ Â ÑÂÎÈÕ ÄÅÉÑÒÂÈßÕ?',4+48,'') != 6
  RETURN 
 ENDIF 

 ppriod = STR(tYear,4)+PADL(tMonth,2,'0')
 spriod = PADL(tMonth,2,'0')+RIGHT(STR(tYear,4),1)

 ppdir  = pbase+'\'+ppriod
 IF !fso.FolderExists(ppdir)
  MESSAGEBOX('ÎÒÑÓÒÑÒÂÓÅÒ ÄÈÐÅÊÒÎÐÈß '+ppdir,0+16,'')  
  RETURN
 ENDIF 
 
 aisfile = ppdir+'\AisOms'
 IF !fso.FileExists(aisfile+'.dbf')
  MESSAGEBOX('ÎÒÑÓÒÑÒÂÓÅÒ ÔÀÉË '+aisfile,0+16,'')  
  RETURN
 ENDIF 
 
 IF OpenFile(aisfile, 'AisOms', 'shared', 'mcod')>0
  RETURN 
 ENDIF 
 IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\UsrLpu', "UsrLpu", "shar", "mcod") > 0
  USE IN aisoms
  RETURN
 ENDIF 

 SELECT AisOms
 SCAN
  m.mcod = mcod
  m.lpuid = STR(lpuid,4)
  m.nvfile = 'nv'+m.lpuid
  m.usr  = IIF(SEEK(m.mcod, "usrlpu"), 'USR'+PADL(usrlpu.usr,3,'0'), "")
  IF m.usr != m.gcUser AND m.gcUser!='OMS'
   LOOP 
  ENDIF 

  WAIT m.mcod WINDOW NOWAIT 

  IF !fso.FolderExists(ppdir+'\'+m.mcod)
   LOOP 
  ENDIF 
  IF !fso.FileExists(ppdir+'\'+m.mcod+'\people.dbf')
   LOOP 
  ENDIF 
  IF !fso.FileExists(ppdir+'\'+m.mcod+'\talon.dbf')
   LOOP 
  ENDIF 
  IF OpenFile(ppdir+'\'+m.mcod+'\people', 'people', 'excl')>0
   LOOP 
  ENDIF 
  IF OpenFile(ppdir+'\'+m.mcod+'\talon', 'talon', 'excl')>0
   USE IN People
   LOOP 
  ENDIF 
  
  SELECT talon 
  IF FSIZE('c_i')!=25
   ALTER TABLE Talon ALTER COLUMN c_i c(25)
  ENDIF 
  IF FIELD('IsPr')!='ISPR'
   ALTER TABLE Talon ADD COLUMN IsPr L
  ENDIF 
  IF FIELD('e_cod')!='E_COD'
   ALTER TABLE Talon ADD COLUMN e_cod n(6)
  ENDIF 
  IF FIELD('e_ku')!='E_KU'
   ALTER TABLE Talon ADD COLUMN e_ku n(3)
  ENDIF 
  IF FIELD('e_tip')!='E_TIP'
   ALTER TABLE Talon ADD COLUMN e_tip c(1)
  ENDIF 
  IF FIELD('err_mee')!='ERR_MEE'
   ALTER TABLE Talon ADD COLUMN err_mee c(3)
  ENDIF 
  IF FIELD('e_period')!='E_PERIOD'
   ALTER TABLE Talon ADD COLUMN e_period c(6)
   REPLACE FOR !EMPTY(err_mee) e_period WITH '201204'
  ENDIF 
  IF FIELD('et')!='ET'
   ALTER TABLE Talon ADD COLUMN et c(1)
  ENDIF 
  IF FIELD('e_ds')!='E_DS'
   ALTER TABLE Talon ADD COLUMN e_ds c(6)
  ENDIF 
  IF FIELD('e_dtype')!='E_DTYPE'
   ALTER TABLE Talon ADD COLUMN e_dtype c(1)
  ENDIF 
  IF FIELD('e_sall')!='E_SALL'
   ALTER TABLE Talon ADD COLUMN e_sall n(11,2)
  ENDIF 
  IF FIELD('koeff')!='KOEFF'
   ALTER TABLE Talon ADD COLUMN koeff n(4,2)
  ENDIF 
  SCAN 
   m.e_period = e_period
   m.et = et 
   IF EMPTY(m.et)
    REPLACE et WITH '2'
   ENDIF 
   IF EMPTY(m.e_period)
    m.new_period = gcPeriod
    REPLACE e_period WITH m.new_period
   ENDIF 
    
   IF OCCURS('.', m.e_period)>0
    m.pos = at('.',e_period)
    m.part1 = SUBSTR(e_period,1, at('.',e_period)-2)
    m.part2 = PADL(SUBSTR(e_period, at('.',e_period)-1, 1),2,'0')
    m.new_period = m.part1 + m.part2
    
    REPLACE e_period WITH m.new_period
    
   ENDIF 
   
*   mefile = 'ME'+m.lpuid
*   IF !fso.FileExists(ppdir+'\'+m.mcod+'\'+mefile+'.dbf')
*    CREATE TABLE &ppdir\&mcod\&mefile ;
     (recid i, er_c c(2), et c(1), osn230 c(9), ds c(6), tip c(1), cod n(6), ;
      k_u n(6), d_type c(1), s_opl n(12,2), s_sank n(12,2))
*     INDEX on recid TAG recid CANDIDATE 
*   ENDIF 
  ENDSCAN 
  USE 
  
  SELECT people
  IF FIELD('IsPr')!='ISPR'
   ALTER TABLE People ADD COLUMN IsPr L
  ENDIF 
  IF FIELD('s_all')!='S_ALL'
   ALTER TABLE People ADD COLUMN s_all n(11,2)
  ENDIF 
  USE 

  IF !fso.FileExists(ppdir+'\'+m.mcod+'\'+m.nvfile+'.dbf')
   IF fso.FileExists(ppdir+'\'+m.mcod+'\'+m.nvfile+'.'+spriod)
    fso.CopyFile(ppdir+'\'+m.mcod+'\'+m.nvfile+'.'+spriod, ppdir+'\'+m.mcod+'\'+m.nvfile+'.dbf')
    oSettings.CodePage(ppdir+'\'+m.mcod+'\'+m.nvfile+'.dbf', 866, .t.)
    IF OpenFile(ppdir+'\'+m.mcod+'\'+nvfile, 'nvfile', 'excl') == 0
     SELECT nvfile 
     INDEX ON pcod TAG pcod 
     USE 
    ENDIF 
   ENDIF 
  ELSE 
*   fso.DeleteFile(ppdir+'\'+m.mcod+'\'+m.nvfile+'.dbf')
*   fso.DeleteFile(ppdir+'\'+m.mcod+'\'+m.nvfile+'.cdx')
  ENDIF 

 ENDSCAN 
 USE 
 USE IN UsrLpu
 
 WAIT CLEAR 

 MESSAGEBOX('OK!', 0+64, '')

RETURN 