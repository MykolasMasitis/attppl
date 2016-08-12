
 

 
 IF 3=2

 List = 1
 Acti Wind Oms2

 If stc_d>0
  Stroka1 = 'Страховая медицинская организация  : ' + Allt(qnam)
  Stroka2 = 'Лечебно-профилактическое учреждение: ' + Allt(lp(num_lpu,2))
  Stroka3 = 'СЧЕТ-ФАКТУРА N '+Allt(Str(NMonth))+' от '+DToC(date())+iif(InfTip=1,' (Основная)','( Дополнительная)')+' за '+Allt(RMon(NMonth,1))+' '+Year+' г.'+' (c '+DToC(Dat1)+ ' по '+DToc(Dat2)+')'
  Stroka4 = 'за оказанные медицинские услуги по программе ОМС '+stroka6
  
  \\&l1O
  \<<PadR(Stroka1,112)>>
  \<<PadR(Stroka2,112)>>
  \
  \<<PadC(Stroka3,112)>>
  \<<PadC(Stroka4,112)>>
  \
   
  For iii = 1 To stc_d
   inOtd = STC(iii,13)
   ggg = ASca(STE,Stc(iii,13))
   \Отделение: <<iif(Seek(Stc(iii,13),'Otdel'), Allt(Otdel.IOtd)+' ('+Allt(Otdel.Otd)+'), '+Allt(Otdel.Name), Space(40))>>
   \Количество койко-дней всего   : <<iif(ggg>0, STE(ggg+1), 0)>>
   \Количество пролеченных больных: <<iif(ggg>0, STE(ggg+2)+STE(ggg+3), 0)>>, в т.ч.: по "законченному" случаю : <<iif(ggg>0, STE(ggg+2), 0)>>, по "прерванному" случаю : <<iif(ggg>0, STE(ggg+3), 0)>>
   \+--------------------------------------------------------------------------------------------------------------+
   \¦ Код  ¦Код по¦   Наименование заболевания     ¦  "Законченные" случаи  ¦"Прерванные"случаи лечения¦ Всего  из ¦
   \¦ МЭС  ¦МКБ-X ¦                                ¦------------------------+--------------------------¦  средств  ¦
   \¦      ¦      ¦                                ¦Кол.¦  Тариф  ¦  Сумма  ¦Кол.¦ МЭС ¦Факт ¦  Сумма  ¦   Фонда   ¦
   \¦------+------+--------------------------------+----+---------+---------+----+-----+-----+---------+-----------¦
   Dime STD(7)
   STD = 0
   Do While STC(iii,13) = inOtd And iii<=stc_d
    std(1) = Std(1) + stc(iii,5)
    std(2) = Std(2) + stc(iii,7)
    std(3) = Std(3) + stc(iii,8)
    std(4) = Std(4) + stc(iii,9)
    std(5) = Std(5) + stc(iii,10)
    std(6) = Std(6) + stc(iii,11)
    std(7) = Std(7) + stc(iii,12)
    \¦<<Stc(iii,2)>>¦<<Stc(iii,3)>>¦<<Stc(iii,4)>>¦<<PadL(Stc(iii,5),4))>>¦<<Trans(Stc(iii,6),'999999.99')>>¦
    \\<<Trans(Stc(iii,7),'999999.99')>>¦<<PadL(Stc(iii,8),4)>>¦<<PadL(Stc(iii,9),5)>>¦<<PadL(Stc(iii,10),5)>>¦
    \\<<Trans(Stc(iii,11),'999999.99')>>¦<<Trans(Stc(iii,12),'99999999.99')>>¦
    iii = iii+1
    if iii > stc_d
     Exit
    EndI
   EndD
   iii = iii-1
   \¦----------------------------------------------+----+---------+---------+----+-----+-----+---------+-----------¦
   \¦ Итого:                                       ¦<<PadL(Std(1),4))>>¦         ¦<<Trans(Std(2),'999999.99')>>¦<<PadL(Std(3),4)>>
   \\¦<<PadL(Std(4),5)>>¦<<PadL(Std(5),5)>>¦<<Trans(Std(6),'999999.99')>>
   \\¦<<Trans(Std(7),'99999999.99')>>¦
   \+--------------------------------------------------------------------------------------------------------------+
   if iii < stc_d
*   \
   Else
   \Итого по профильным отделениям:
   \ Количество койко-дней всего         :
   \ Количество пролеченных больных всего: <<k_obrs>>
   \  В том числе:
   \   По "законченному" случаю лечения  : <<k_z>>
   \   По "прерванному" случаю лечения   : <<k_nz>>
   EndI
  EndF

  If sta_d>0
  \
  \&l0O
  For iii = 1 To sta_d
   i=1
   \                                                                # <<alltrim(str(list))>>
   If list=1
    \                                                                Ф.N ОМС - 2
    \+--------------------------------------------------------------------------+
    \¦Услуга¦Наименование оказанной медицинской услуги¦  Кол ¦ Тариф ¦ Cтоим.   ¦
    \¦------+-----------------------------------------+------+-------+----------¦
    \¦  1   ¦                    2                    ¦   3  ¦   4   ¦    5     ¦
    \¦------+-----------------------------------------+------+-------+----------¦
   Else
    \                                                                 Ф.N ОМС - 2
    \+--------------------------------------------------------------------------+
    \¦  1   ¦                   2                     ¦  3   ¦   4   ¦    5     ¦
    \¦------+-----------------------------------------+------+-------+----------¦
   EndI

   Do Whil i <= iif(list=1, PageLen-8, PageLen) And iii <= sta_d
    x = Len(Allt(Sta(iii,5)))
    If x <= 40
     \¦<<Sta(iii,1)>>¦<<PadL(Sta(iii,5),41)>>¦<<PadL(Sta(iii,2),6)>>¦<<Tran(Sta(iii,3), '9999.99')>>¦<<Tran(Sta(iii,4), '9999999.99')>>¦
     i=i+1
    Else
     \¦<<Sta(iii,1)>>¦<<PadL(Sta(iii,5),40)>>-¦<<PadL(Sta(iii,2),6)>>¦<<Tran(Sta(iii,3), '9999.99')>>¦<<Tran(Sta(iii,4),'9999999.99')>>¦
     \¦<<Repl(' ',6)>>¦<<PadR(Subs(Allt(Sta(iii,5)),41,40),41)>>¦<<RepL(' ',6)>>¦<<RepL(' ',7)>>¦<<RepL(' ',10)>>¦
     i=i+2
    EndI
    iii = iii + 1
    was_time=seco()
    Do Whil Seco()-was_time <= 0.05
    EndD
   EndD
   iii = iii-1

   If iii < sta_d
    \+--------------------------------------------------------------------------+
    \
   Else
    \¦------+-----------------------------------------+------+-------+----------¦
    \¦ИТОГО:¦                                         ¦<<PadL(Allt(Str(TotalUsl)),6)>>¦   х   ¦<<Tran(TotalSumP, '9999999.99'))>>¦
    \+--------------------------------------------------------------------------+
   EndI
   list=list+1
  EndF
  EndI
  TotalSum = TotalSumS + TotalSumP
  \
  \Итого представлено к оплате: <<trans(totalsum,'99999999.99')>> руб.
  \<<cpr(INT(totalsum))+str(int((totalsum-int(totalsum))*100),2)+' КОП.'>>
  \
  \Страховая стоимость услуг заявленных для исполнения в другом ЛПУ:
  \... <<trans(ZSum,'99999999.99')>> руб.
  \<<cpr(INT(ZSum))+str(int((ZSum-int(ZSum))*100),2)+' КОП.'>>
  \
  \ От медучреждения <<allt(lp(num_lpu,2))>>
  \  Руководитель медучреждения                  <<allt(lp(num_lpu,3))>>
  \
  \  Главный бухгалтер                           <<allt(lp(num_lpu,4))>>
  \
  \&l0O
  
 Else
 For iii = 1 To sta_d
  i=1
  \                                                                # <<alltrim(str(list))>>
  If list=1
   \                                                                Ф.N ОМС - 2
   \ Страховая организация <<qnam>>
   \ Медицинское учреждение : <<allt(lp(num_lpu,2))>>
   \
   \                 СЧЕТ-ФАКТУРА N <<NMonth>> от <<date()>> <<iif(InfTip=1,'(Основная)','(Дополнительная)')>>
   \                             за <<allt(RMon(NMonth,1))>> <<Year>> г.
   \                 за оказанные медицинские услуги по программе ОМС
   \ Количество обратившихся пациентов : <<kol_obr>>
   \
   \+--------------------------------------------------------------------------+
   \¦Услуга¦Наименование оказанной медицинской услуги¦  Кол ¦ Тариф ¦ Cтоим.   ¦
   \¦------+-----------------------------------------+------+-------+----------¦
   \¦  1   ¦                    2                    ¦   3  ¦   4   ¦    5     ¦
   \¦------+-----------------------------------------+------+-------+----------¦
  Else
   \                                                                 Ф.N ОМС - 2
   \ Страховая организация : <<qnam>>
   \ Медицинское учреждение: <<allt(lp(num_lpu,2))>>
   \
   \+--------------------------------------------------------------------------+
   \¦  1   ¦                   2                     ¦  3   ¦   4   ¦    5     ¦
   \¦------+-----------------------------------------+------+-------+----------¦
  EndI

  Do Whil i <= iif(list=1, PageLen-8, PageLen) And iii <= sta_d
   x = Len(Allt(Sta(iii,5)))
   If x <= 40
    \¦<<Sta(iii,1)>>¦<<PadL(Sta(iii,5),41)>>¦<<PadL(Sta(iii,2),6)>>¦<<Tran(Sta(iii,3), '9999.99')>>¦<<Tran(Sta(iii,4), '9999999.99')>>¦
    i=i+1
   Else
    \¦<<Sta(iii,1)>>¦<<PadL(Sta(iii,5),40)>>-¦<<PadL(Sta(iii,2),6)>>¦<<Tran(Sta(iii,3), '9999.99')>>¦<<Tran(Sta(iii,4),'9999999.99')>>¦
    \¦<<Repl(' ',6)>>¦<<PadR(Subs(Allt(Sta(iii,5)),41,40),41)>>¦<<RepL(' ',6)>>¦<<RepL(' ',7)>>¦<<RepL(' ',10)>>¦
    i=i+2
   EndI
   iii = iii + 1
   was_time=seco()
   Do Whil Seco()-was_time <= 0.05
   EndD
  EndD
  iii = iii-1

  if iii < sta_d
   \+--------------------------------------------------------------------------+
  else
   \¦------+-----------------------------------------+------+-------+----------¦
   \¦ИТОГО:¦                                         ¦<<PadL(Allt(Str(TotalUsl)),6)>>¦   х   ¦<<Tran(TotalSumP, '9999999.99'))>>¦
   \+--------------------------------------------------------------------------+
   \
   \ Итого представлено к оплате: <<trans(TotalSumP,'99999999.99')>> руб.
   \   <<cpr(INT(TotalSumP))+str(int((TotalSumP-int(TotalSumP))*100),2)+' КОП.'>>
   \
   \ От медучреждения <<allt(lp(num_lpu,2))>>
   \  <<PadR(Allt(lp(num_lpu,32)),40)>>    <<allt(lp(num_lpu,3))>>
   \
   \  <<PadR(Allt(lp(num_lpu,33)),40)>>    <<allt(lp(num_lpu,4))>>
   \
  EndI
  \
  list=list+1
 EndF
 EndI

  Deac Wind Oms2
  Do BeepOk
  set text to
  set text noshow
  TxFile=iif(Polz=1, '&PLocal\oms2.txt', '&PBase\&lp(num_lpu,1)\Oms2.txt')

  If Kol_Lpu = 1
   copy file &TxFile to &PLocal\temp.txt
   on key label Ctrl+P copy file &PLocal\temp.txt to prn
   Keyb [{Ctrl+End}]
   modi comm &TxFile wind Oms2 noed
   on key label Ctrl+P
   dele file &PLocal\temp.txt
  EndI

 Use In Otdel

 Pop Key
*lnSumtarz = 0
*SUM s_all TO lnSumtarz FOR !INLIST(LOWER(d_type),'z','h')
*IF lnSumtarz > 0
*	lcDocZ = 'Страховая стоимость услуг, заявленных для исполнения в другом ЛПУ ' +;
*		ALLTRIM(STR(INT(lnSumtarz))) + ' руб. ' + ALLTRIM(STR(mod(lnSumtarz,1)*100)) + ' коп.'
*ELSE
*	lcDocZ = ''
*endif

*SELECT * FROM cSfile WHERE !INLIST(LOWER(d_type),'z','h') INTO CURSOR cSfilez


*if lcProf = 'p'
*	do case
*		case oSettings.PrintDevice
*			report form sfp.frx preview
*			REPORT FORMAT sfp.frx TO file d:\sfp.txt ascii
*		otherwise
*			PRINTJOB
*				REPORT FORMAT sfp.frx TO PRINT NOCONS
*			ENDPRINTJOB
*	endcase

*else
*	if lcProf = 's'
*		lnNumPage = 0
*		if reccount() > 0
*			do case
*				case oSettings.PrintDevice
*					report form sfs1.frx preview
*				otherwise
*					PRINTJOB
*						REPORT FORMAT sfs1.frx TO PRINT NOCONS
*					ENDPRINTJOB
*			endcase
*			lnNumPage = lnNumPage + _pageno
*		endif
*		use in cSFprof

*		select cSFpol
*		if reccount() > 0
*			do case
*				case oSettings.PrintDevice
*					report form sfs2.frx preview
*				otherwise
*					PRINTJOB
*						REPORT FORMAT sfs2.frx TO PRINT NOCONS
*					ENDPRINTJOB
*			endcase
*			lnNumPage = lnNumPage + _pageno
*		endif
*		use in cSFpol

*		select cSFsvod
*		if reccount() > 0
*			do case
*				case oSettings.PrintDevice
*					report form sfs3.frx preview
*				otherwise
*					PRINTJOB
*						REPORT FORMAT sfs3.frx TO PRINT NOCONS
*					ENDPRINTJOB
*			endcase
*		endif
*		use in cSFsvod
*	else
*		messagebox('Черт знает какое ЛПУ',0,'')	
*	endif
*endif

*loTherm.release

ENDIF 