FUNCTION erz_lpu(lcDir, mcod, IsOK)
 PRIVATE lcdir,mcod,isok

 IF !fso.FolderExists(lcDir)
  RETURN .F.
 ENDIF 

 IF !fso.FileExists(lcDir + '\People.dbf')
  RETURN .F.
 ENDIF 

 IF OpenFile("&lcDir\People", "People", "SHARE")>0
  RETURN .F.
 ENDIF 
 
 IF RECCOUNT('people')<=0
  USE IN people 
  RETURN .F.
 ENDIF 
 
 fso.CopyFile(ptempl+'\erz_lpu.dbf', lcDir+'\erz_lpu.dbf', .t.)
 oSettings.CodePage(lcDir+'\erz_lpu.dbf', 866, .t.)
 
 IF OpenFile("&lcDir\erz_lpu", "Zapros", "SHARE")>0
  USE IN people
  RETURN .F.
 ENDIF 

 SELECT people
 SCAN 
  m.recid = PADL(recid,6,'0')
  m.s_pol = s_pol
  m.n_pol = n_pol
  m.tip_d = tip_d

  INSERT INTO Zapros FROM MEMVAR 
    
 ENDSCAN 
 USE IN Zapros
 USE IN people 

 SELECT AisOms
   
 ChVal  = SYS(3)
 ID     = ALLTRIM(ChVal+'.'+m.gcUser+'@'+m.qmail)
  
 TFile  = 'terz_' + mcod
 BFile  = 'berz_' + mcod
 DFile  = 'derz_' + mcod

 iii = 1
 DO WHILE fso.FileExists(pAisOms+'\'+m.gcUser+'\OUTPUT\'+m.bfile)
  m.tfile  = 'terz_' + m.mcod + '_' + PADL(iii,2,'0')
  m.bfile  = 'berz_' + m.mcod + '_' + PADL(iii,2,'0')
  m.dfile  = 'derz_' + m.mcod + '_' + PADL(iii,2,'0')
  iii = iii + 1
 ENDDO 
   
 fso.CopyFile(lcDir+'\erz_lpu.dbf', PAisOms+'\'+m.gcUser+'\OutPut\'+DFile)
 fso.DeleteFile(lcDir+'\erz_lpu.dbf')
   
 poi = FCREATE('&PAisOms\&gcUser\OutPut\&TFile')
 IF poi != -1
  =FPUTS(poi,'To: erz@mgf.msk.oms')
  =FPUTS(poi,'Message-Id: &ID')
  =FPUTS(poi,'Subject: erz_lpu')
  fzz = 'q_' + PADL(MONTH(DATE()),2,'0')+RIGHT(ALLTRIM(STR(YEAR(DATE()))),1)+'.dbf'
  =FPUTS(poi,'Attachment: &DFile &Fzz')
 ENDIF 
 =FCLOSE(poi)
 
 oTFile = fso.GetFile('&PAisOms\&gcUser\OutPut\&TFile')
 oTFile.Move('&PAisOms\&gcUser\OutPut\&BFile')

 REPLACE erzlpu_id WITH m.id, lpust WITH 1

 IF IsOk==.t.
  MESSAGEBOX('«¿œ–Œ— Œ“œ–¿¬À≈Õ!', 0+64, '')
 ENDIF 
RETURN .T.