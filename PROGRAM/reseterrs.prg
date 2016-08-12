FUNCTION ResetErrs(ppolis, tipexp)

 PRIVATE TipOfExp
 m.TipOfExp = tipexp

 m.addsumplk = 0
 m.addsumst  = 0
 m.addsumdst = 0

 m.ambbadmee = 0
 m.stbadmee  = 0
 m.dstbadmee = 0
 
 m.ambchkdmee = 0
 m.stchkdmee  = 0
 m.dstchkdmee = 0

 orecp = RECNO()
 SELECT talon 
 
 IF m.TipOfExp='4'
  selparam = "INLIST(et,'4','5','6')"
 ELSE 
  selparam = "INLIST(et,'2','3')"
 ENDIF 

 SCAN FOR sn_pol = ppolis
  IF !EMPTY(err_mee) AND &selparam
   IF LEFT(err_mee,2) != 'W0'
    DO CASE 
     CASE IsUsl(cod)
      m.addsumplk = m.addsumplk + (s_all-fsumm(e_cod, e_tip, e_ku, IIF(LEFT(mcod,1) == '0', .F., .T.)))
     CASE IsKD(cod)
      m.addsumdst = m.addsumdst + (s_all-fsumm(e_cod, e_tip, e_ku, IIF(LEFT(mcod,1) == '0', .F., .T.)))
     CASE IsMes(cod)
      m.addsumst = m.addsumst + (s_all-fsumm(e_cod, e_tip, e_ku, IIF(LEFT(mcod,1) == '0', .F., .T.)))
     OTHERWISE 
     ENDCASE 
   ELSE && W0
    DO CASE 
     CASE IsUsl(cod)
      m.ambchkdmee = m.ambchkdmee + 1
     CASE IsKD(cod)
      m.dstchkdmee = m.dstchkdmee + 1
     CASE IsMes(cod)
      m.stchkdmee = m.stchkdmee + 1
     OTHERWISE 
     ENDCASE 
   ENDIF 
  ENDIF 
  REPLACE err_mee WITH '', et WITH '', e_period WITH '', ;
   e_cod WITH 0, e_ku WITH 0, e_tip WITH '', e_sall WITH 0
 ENDSCAN 

 GO (orecp)

 IF m.addsumplk > 0
  goApp.ambmeesum = goApp.ambmeesum - m.addsumplk
  goApp.ambbadmee = goApp.ambbadmee - IIF(m.addsumplk>0,1,0)
 ENDIF 
 goApp.ambchkdmee = goApp.ambchkdmee - IIF(m.addsumplk>0 or m.ambchkdmee>0,1,0)
 IF m.addsumst  > 0
  goApp.stmeesum = goApp.stmeesum - m.addsumst
  goApp.stbadmee = goApp.stbadmee - IIF(m.addsumst>0,1,0)
 ENDIF 
  goApp.stchkdmee = goApp.stchkdmee - IIF(m.addsumst>0 or m.stchkdmee>0,1,0)
 IF m.addsumdst > 0
  goApp.dstmeesum = goApp.dstmeesum - m.addsumdst
  goApp.dstbadmee = goApp.dstbadmee - IIF(m.addsumdst>0,1,0)
 ENDIF 
 goApp.dstchkdmee = goApp.dstchkdmee - IIF(m.addsumdst>0 or m.dstchkdmee>0,1,0)

RETURN 