PROCEDURE m_menu
SET SYSMENU TO

DEFINE PAD mmenu_1 OF _MSYSMENU PROMPT '\<ÈÍÔÎÐÌÀÖÈß ÎÒ ËÏÓ' COLOR SCHEME 3 KEY ALT+A, ""
DEFINE PAD mmenu_2 OF _MSYSMENU PROMPT '\<ÝÊÑÏÅÐÒÈÇÀ' COLOR SCHEME 3 KEY ALT+A, ""
DEFINE PAD mmenu_5 OF _MSYSMENU PROMPT '\<ÑÅÐÂÈÑ' COLOR SCHEME 3 KEY ALT+F , ""
ON PAD mmenu_1 OF _MSYSMENU ACTIVATE POPUP popInfFrLpu
ON PAD mmenu_2 OF _MSYSMENU ACTIVATE POPUP popExp
ON PAD mmenu_5 OF _MSYSMENU ACTIVATE POPUP popTuneUp

DEFINE POPUP popInfFrLpu MARGIN RELATIVE shadow COLOR SCHEME 4
DEFINE BAR 01 OF popInfFrLpu PROMPT 'ÏÐÈÅÌ ÈÍÎÐÌÀÖÈÈ ÎÒ ËÏÓ (ÀÈÑ ÎÌÑ)'
DEFINE BAR 02 OF popInfFrLpu PROMPT '\-'
DEFINE BAR 03 OF popInfFrLpu PROMPT 'ÄÅÔÅÊÒÍÛÅ ÈÍÔÎÐÌÀÖÈÎÍÍÛÅ ÏÎÑÛËÊÈ'
DEFINE BAR 04 OF popInfFrLpu PROMPT 'ÏÎÂÒÎÐÍÛÅ ÈÍÔÎÐÌÀÖÈÎÍÍÛÅ ÏÎÑÛËÊÈ'
DEFINE BAR 05 OF popInfFrLpu PROMPT '\-'
DEFINE BAR 06 OF popInfFrLpu PROMPT 'ÑÎÁÐÀÒÜ ÑÂÎÄÍÛÉ ÐÅÃÈÑÒÐ'
DEFINE BAR 07 OF popInfFrLpu PROMPT 'ÑÔÎÐÌÈÐÎÂÀÒÜ ÔÀÉËÛ ÎØÈÁÎÊ ÄËß ËÏÓ' SKIP FOR !fso.FileExists(pbase+'\'+gcperiod+'\people.dbf')
DEFINE BAR 08 OF popInfFrLpu PROMPT 'ÑÔÎÐÌÈÐÎÂÀÒÜ ÀÊÒÛ ÑÂÅÐÊÈ Ñ ËÏÓ' SKIP FOR !fso.FileExists(pbase+'\'+gcperiod+'\people.dbf')
DEFINE BAR 09 OF popInfFrLpu PROMPT 'ÎÒÏÐÀÂÈÒÜ ÎÒ×ÅÒ Â ËÏÓ' SKIP FOR !fso.FileExists(pbase+'\'+gcperiod+'\people.dbf')
DEFINE BAR 10 OF popInfFrLpu PROMPT '\-'
DEFINE BAR 11 OF popInfFrLpu PROMPT 'ÑÔÎÐÌÈÐÎÂÀÒÜ ÎÒ×ÅÒ Â ÌÃÔÎÌÑ' SKIP FOR !fso.FileExists(pbase+'\'+gcperiod+'\people.dbf')
DEFINE BAR 12 OF popInfFrLpu PROMPT 'ÑÔÎÐÌÈÐÎÂÀÒÜ ÎÒ×ÅÒ Â ÌÃÔÎÌÑ (ÍÎÂÀß ÂÅÐÑÈß)'
DEFINE BAR 13 OF popInfFrLpu PROMPT 'ÑÔÎÐÌÈÐÎÂÀÒÜ ÑÂÎÄÍÛÉ ÀÊÒ Â ÌÃÔÎÌÑ' SKIP FOR !fso.FileExists(pbase+'\'+gcperiod+'\people.dbf')
DEFINE BAR 14 OF popInfFrLpu PROMPT 'ÎÒÏÐÀÂÈÒÜ ÎÒ×ÅÒ Â ÌÃÔÎÌÑ' SKIP FOR !fso.FileExists(pbase+'\'+gcperiod+'\people.dbf')
DEFINE BAR 15 OF popInfFrLpu PROMPT '\-'
DEFINE BAR 16 OF popInfFrLpu PROMPT 'ÇÀÃÐÓÇÈÒÜ ÎØÈÁÊÈ ÌÃÔÎÌÑ'
DEFINE BAR 17 OF popInfFrLpu PROMPT 'ÑÔÎÐÌÈÐÎÂÀÒÜ ÎØÈÁÊÈ ÄËß ËÏÓ (2-ÎÉ ÝÒÀÏ)' ;
 SKIP FOR !fso.FileExists(pbase+'\'+gcperiod+'\'+'er'+m.qcod+PADL(tmonth,2,'0')+RIGHT(STR(tyear,4),2)+'.dbf')
DEFINE BAR 18 OF popInfFrLpu PROMPT 'ÑÔÎÐÌÈÐÎÂÀÒÜ ÔÀÉËÛ ÎØÈÁÎÊ ÄËß ÎÒÏÐÀÂÊÈ Â ËÏÓ' SKIP FOR !fso.FileExists(pbase+'\'+gcperiod+'\people.dbf')
DEFINE BAR 19 OF popInfFrLpu PROMPT 'ÑÎÁÐÀÒÜ ÝÒÀËÎÍÍÛÉ ÐÅÃÈÑÒÐ' ;
 SKIP FOR !fso.FileExists(pbase+'\'+gcperiod+'\'+'er'+m.qcod+PADL(tmonth,2,'0')+RIGHT(STR(tyear,4),2)+'.dbf')
DEFINE BAR 20 OF popInfFrLpu PROMPT '\-'
DEFINE BAR 21 OF popInfFrLpu PROMPT 'ÏÀÊÅÒÍÀß ÏÅ×ÀÒÜ'
DEFINE BAR 22 OF popInfFrLpu PROMPT '\-'
DEFINE BAR 23 OF popInfFrLpu PROMPT 'ÂÛÕÎÄ'

ON SELECTION BAR 01 OF popInfFrLpu DO FORM MailView
ON SELECTION BAR 03 OF popInfFrLpu DO FORM MailTrash
ON SELECTION BAR 04 OF popInfFrLpu DO FORM MailDouble
ON SELECTION BAR 06 OF popInfFrLpu DO MakeSvPeople
ON SELECTION BAR 07 OF popInfFrLpu DO MakeETRLs
ON SELECTION BAR 08 OF popInfFrLpu DO MakeActsLpu
ON SELECTION BAR 09 OF popInfFrLpu DO SndToLpu
ON SELECTION BAR 11 OF popInfFrLpu DO MakeFOMS
ON SELECTION BAR 12 OF popInfFrLpu DO MakeFOMS2
ON SELECTION BAR 13 OF popInfFrLpu DO MakeActFOMS
ON SELECTION BAR 14 OF popInfFrLpu DO SendToFoms
ON SELECTION BAR 16	OF popInfFrLpu DO LoadErrFomsN
ON SELECTION BAR 17 OF popInfFrLpu DO MakeETRLs22
ON SELECTION BAR 18 OF popInfFrLpu DO MakeETRLsN
ON SELECTION BAR 19 OF popInfFrLpu DO MakeEtPeople
ON BAR 21 OF popInfFrLpu ACTIVATE POPUP popPrn
ON SELECTION BAR 23 OF popInfFrLpu CLEAR EVENTS 

DEFINE POPUP popPrn MARGIN RELATIVE shadow COLOR SCHEME 4
DEFINE BAR 1 OF popPrn PROMPT 'ÏÅ×ÀÒÜ ÀÊÒÀ ÑÂÅÐÊÈ'
ON SELECTION BAR 1 OF popPrn DO PackPrnSv

DEFINE POPUP popExp MARGIN RELATIVE shadow COLOR SCHEME 4
DEFINE BAR 01 OF popExp PROMPT 'ÝÊÑÏÅÐÒÈÇÀ ÒÅÊÓÙÅÃÎ ÏÅÐÈÎÄÀ'
DEFINE BAR 02 OF popExp PROMPT 'ÑÔÎÐÌÈÐÎÂÀÒÜ ÑÏÈÑÊÈ ÏÎ ÂÑÅÌ ÌÎ'
DEFINE BAR 03 OF popExp PROMPT 'ÎÒÏÐÀÂÈÒÜ ÑÏÈÑÊÈ ÂÑÅÌ ÌÎ'
DEFINE BAR 04 OF popExp PROMPT 'ÀÍÍÓËÈÐÎÂÀÍÈÅ ÏÐÈÊÐÅÏËÅÍÈß'
ON SELECTION BAR 01 OF popExp DO FORM ExpView
ON SELECTION BAR 02 OF popExp DO PlWait01
ON SELECTION BAR 03 OF popExp DO PlWait02
ON BAR 04 OF popExp ACTIVATE POPUP AnnPpl

DEFINE POPUP AnnPpl MARGIN RELATIVE SHADOW COLOR SCHEME 4
DEFINE BAR 01 OF AnnPpl PROMPT 'ÏÐÎÂÅÐÊÀ ÔÀÉËÎÂ ÄËß ÈÏ'
DEFINE BAR 02 OF AnnPpl PROMPT 'ÎÒÏÐÀÂÊÀ ÈÏ Â ÌÃÔÎÌÑ'
ON SELECTION BAR 01 OF AnnPpl do ChkAnnIpS
ON SELECTION BAR 02 OF AnnPpl do SndAnnIpS


DEFINE POPUP popTuneUp MARGIN RELATIVE SHADOW COLOR SCHEME 4
DEFINE BAR 01 OF popTuneUp PROMPT 'ÂÛÁÎÐ ÎÒ×ÅÒÍÎÃÎ ÏÅÐÈÎÄÀ' 
DEFINE BAR 02 OF popTuneUp PROMPT 'ÍÀÑÒÐÎÉÊÀ ÐÀÁÎ×ÈÕ ÄÈÐÅÊÒÎÐÈÉ'
DEFINE BAR 03 OF popTuneUp PROMPT 'ÏÀÐÀÌÅÒÐÛ ÏÅ×ÀÒÈ' 
DEFINE BAR 04 OF popTuneUp PROMPT '\-'
DEFINE BAR 05 OF popTuneUp PROMPT 'ÏÅÐÅÈÍÄÅÊÑÀÖÈß ÁÄ ÍÑÈ'
DEFINE BAR 06 OF popTuneUp PROMPT 'ÏÅÐÅÈÍÄÅÊÑÀÖÈß ÐÀÁÎ×ÈÕ ÁÀÇ'
DEFINE BAR 07 OF popTuneUp PROMPT '\-'
DEFINE BAR 08 OF popTuneUp PROMPT 'ÓÏÀÊÎÂÀÒÜ ÔÀÉËÛ ÎØÈÁÎÊ'
DEFINE BAR 09 OF popTuneUp PROMPT 'ÓÏÀÊÎÂÀÒÜ ÑÂÎÄÍÛÅ ÁÀÇÛ'
DEFINE BAR 10 OF popTuneUp PROMPT '\-'
DEFINE BAR 11 OF popTuneUp PROMPT 'ÎÁÍÓËÈÒÜ ÔÀÉËÛ ÎØÈÁÎÊ'
DEFINE BAR 12 OF popTuneUp PROMPT 'ÓÄÀËÈÒÜ CTRL-êè'
DEFINE BAR 13 OF popTuneUp PROMPT 'ÓÄÀËÈÒÜ ÔÀÉËÛ ÇÀÏÐÎÑÎÂ'
DEFINE BAR 14 OF popTuneUp PROMPT 'ÓÄÀËÈÒÜ BAK-ÔÀÉËÛ'
DEFINE BAR 15 OF popTuneUp PROMPT 'ÓÄÀËÈÒÜ ALLPEOPLE'
DEFINE BAR 16 OF popTuneUp PROMPT '\-'
DEFINE BAR 17 OF popTuneUp PROMPT 'ÊÎÐÐÅÊÒÈÐÎÂÊÀ ÑÒÐÓÊÒÓÐÛ ÐÀÁÎ×ÈÕ ÁÄ'
DEFINE BAR 18 OF popTuneUp PROMPT '\-'
DEFINE BAR 19 OF popTuneUp PROMPT 'ÇÀÃÐÓÇÈÒÜ ÄÀÍÍÛÅ Î ÏÐÈÊÐÅÏËÅÍÈÈ'
DEFINE BAR 20 OF popTuneUp PROMPT 'ÑÒÀÒÈÑÒÈÊÀ ÏÎ ÎØÈÁÊÀÌ'

ON SELECTION BAR 01 OF popTuneUp DO FORM SetPeriod
ON SELECTION BAR 02 OF popTuneUp DO FORM TuneBase
ON SELECTION BAR 03 OF popTuneUp goApp.doForm('set_print','settings')
ON SELECTION BAR 05 OF popTuneUp DO ComReind
ON SELECTION BAR 06 OF popTuneUp DO BasReind
ON SELECTION BAR 08 OF popTuneUp DO PackBD
ON SELECTION BAR 09 OF popTuneUp DO PackBDSv
ON SELECTION BAR 11 OF popTuneUp DO ZapEFiles
ON SELECTION BAR 12 OF popTuneUp DO DelAllCtrl
ON SELECTION BAR 13 OF popTuneUp DO DelAllZapros
ON SELECTION BAR 14 OF popTuneUp DO DelBakFiles
ON SELECTION BAR 15 OF popTuneUp DO DelAllPeople
ON SELECTION BAR 17 OF popTuneUp DO CorStruct
ON SELECTION BAR 19 OF popTuneUp DO AppMgfData
ON SELECTION BAR 20 OF popTuneUp DO ErrStat

SET SYSMENU AUTOMATIC
SET SYSMENU ON