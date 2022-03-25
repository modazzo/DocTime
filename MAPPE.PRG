//Programm zum Erzeugen für Termintagesordner
#include 'fivewin.ch' 
//----------------------------------------------------------------------------//

REQUEST DBFCDX
REQUEST DBFFPT
//----------------------------------------------------------------------------//


function Main() 
	local I := 0
	local dStart 	:= date()
	local dEnd   	:= date() + 365
	local cDir  	:= "c:\Termine_nachfassen"
	local cTempDay := ""
	
	? "c:\Termine_nachfassen für 1 Jahr im Voraus"
	
	if lIsDir( cDir ) = .f.
		lMKDir( cDir )
	endif
	
	for I := 0  to ( dEnd - dStart )
		
		cTempDay := substr( ALLTRIM( dtos( dStart + I ) ),3, 2 ) + "-" +substr( ALLTRIM( dtos( dStart + I ) ), 5,2 ) + "-" + substr( ALLTRIM( dtos( dStart + I ) ), 7,2 )   +"_"+ fWH_KURZDATUM( dStart + I )
		
		if lIsDir( cDir + "\" + cTempDay ) = .f.
			lMKDir( cDir + "\" + cTempDay )
		endif
	
	next

? "Ende"

RETURN NIL
//----------------------------------------------------------------------------//

INIT PROCEDURE PrgInit

   SET CENTURY ON
   SET EPOCH TO YEAR(DATE())-98

   SET DELETED ON
   SET EXCLUSIVE OFF

   REQUEST HB_Lang_DE
  

   HB_LangSelect("DE")
  

   SET DATE TO GERMAN

   SetHandleCount(205)
   rddsetdefault( "DBFCDX" )
   SetGetColorFocus()

   EXTERN DESCEND

   SetBalloon( .F. )
   SetKinetic( .T. )

   fwNumFormat( 'E', .t. )

RETURN

//----------------------------------------------------------------------------//



function fWH_KURZDATUM( dDatum )
	local cKurzDatum := ""

	cKurzDatum := num2day( dow( dDatum) )


return( cKurzDatum )
//----------------------------------------------------------------------------//



FUNCTION num2day(nTag)
   local cTag     := ""
   local cTITLE   := " Zimmerstammdaten "
   *----------------------------------------------------------
   IF nTag = 1
      cTag ="SONNTAG"
   elseif nTag = 2
      cTag ="Montag"
   elseif nTag = 3
      cTag ="Dienstag"
   elseif nTag = 4
      cTag="Mittwoch"
   elseif nTag = 5
      cTag="Donnerstag"
   elseif nTag = 6
      cTag="Freitag"
   elseif nTag = 7
      cTag="Samstag"
   ENDIF
return (cTag)
//----------------------------------------------------------------------------//
