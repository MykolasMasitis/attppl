FUNCTION ActOneMee(para1, para2, para3) && Ïëàíîâàÿ ÌÝÝ

 m.lcmcod    = para1
 m.lnlpuid   = para2
 m.IsVisible = para3
 
 m.lpuname = IIF(SEEK(m.lcmcod, 'splpu'), ALLTRIM(splpu.name)+', '+m.lcmcod, '')

 DotName = 'Àêò_ÌÝÝ_ÑÑ.dot'
 IF !fso.FileExists(pTempl+'\'+DotName)
  MESSAGEBOX('ÎÒÑÓÒÑÒÂÓÅÒ ÔÀÉË ØÀÁËÎÍ ÎÒ×ÅÒÀ'+CHR(13)+CHR(10)+;
   'Àêò_ÌÝÝ_ÑÑ.dot',0+32,'')
  RETURN 
 ENDIF 

 m.exp_dat1 = '01.'+PADL(tMonth,2,'0')+'.'+STR(tYear,4)
 m.exp_dat2 = DTOC(GOMONTH(CTOD(m.exp_dat1),1)-1)

 m.sn_pol = n_pol
 m.ppolis = STRTRAN(ALLTRIM(m.sn_pol),' ','') && Äëÿ íàçâàíèÿ Àêòà
 m.DocName = m.lcmcod+'_'+m.gcperiod+'_'+m.ppolis
  
 IF m.IsVisible = .t.
 WAIT "ÇÀÏÓÑÊ WORD..." WINDOW NOWAIT 
 TRY 
  oWord = GETOBJECT(,"Word.Application")
 CATCH 
  oWord = CREATEOBJECT("Word.Application")
 ENDTRY 
 WAIT CLEAR 
 ENDIF 

 oDoc = oWord.Documents.Add(pTempl+'\'+DotName)
 
 m.d_akt = DTOC(DATE())
 
 m.straf = 0
 m.totstraf = 0

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
 SELECT people 
 m.IsDoc = IsDoc
 m.straf = IIF(m.IsDoc,0,m.ynorm)

 oDoc.Bookmarks('d_exp').Select  
 oWord.Selection.TypeText(DTOC(DATE()))
 oDoc.Bookmarks('d_akt').Select  
 oWord.Selection.TypeText(m.d_akt)
 oDoc.Bookmarks('d_akt2').Select  
 oWord.Selection.TypeText(m.d_akt)
 oDoc.Bookmarks('smo_name').Select  
 oWord.Selection.TypeText(m.qname)
 oDoc.Bookmarks('lpu_name').Select  
 oWord.Selection.TypeText(m.lpuname)
 oDoc.Bookmarks('sn_pol').Select  
 oWord.Selection.TypeText(ALLTRIM(m.sn_pol))
 oDoc.Bookmarks('straf').Select  
 oWord.Selection.TypeText(TRANSFORM(m.straf, '999999.99'))
 oDoc.Bookmarks('tot_straf').Select  
 oWord.Selection.TypeText(TRANSFORM(m.straf, '999999.99'))

 oDoc.SaveAs(pout+'\'+DocName,0)
 
 IF m.IsVisible=.t. 
  oWord.Visible = .t.
 ELSE 
  oDoc.Close
 ENDIF 

RETURN 

