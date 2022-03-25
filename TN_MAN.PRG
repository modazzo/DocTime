// C:\DOCTIME\TN_MAN.PRG - 25/03/2022 - Joao
// Programm zum erzeugen der Tagesordner C:\FWH\samples\mappe.prg

#Include "FiveWin.Ch"

#define Primary            RGB(18, 102, 241)
#define Secondary          RGB(178, 60, 253)
#define Success            RGB(0, 183, 74)
#define Danger             RGB(249, 49, 84)
#define Warning            RGB(255, 169, 0)
#define Info               RGB( 57, 192, 237)
#define Light              RGB(251, 251, 251)
#define Dark               RGB(38, 38, 38)
#define CLR_COBALT         RGB( 0, 80, 239 )
#define CLR_VIOLET         RGB( 170,   0, 255 )
#define CLR_PINK           RGB( 244, 114, 208 )
#define CLR_CRIMSON        RGB( 162,   0,  37 )
#define CLR_AMBER          RGB( 240, 163,  10 )
#define CLR_OLIVE          RGB( 109, 135, 100 )
#define CLR_STEEL          RGB( 100, 118, 135 )
#define CLR_MAUVE          RGB( 118,  96, 138 )
#define CLR_TAUPE          RGB( 135, 121,  78 )
#define CLR_BLAUMETRO      RGB(   0, 174, 219 )
#define CLR_ROTMETRO       RGB( 191,  30,  75 )
#define CLR_PURPLE         RGB( 111,  66, 193 )
#Define CLR_MSPURPLE       RGB( 0,   120, 215 )
#Define CLR_MSRED          RGB( 232,  17,  35 )
#Define CLR_MSGRAY         RGB( 229, 229, 229 )
#Define CLR_LGRAY          RGB( 230, 230, 230 )


#xcommand TEXT INTO <v> => #pragma __cstream|<v>:=%s
#xcommand TEXT INTO <v> ADDITIVE => #pragma __cstream|<v>+=%s
#xcommand TEXT TO VAR <var> => #pragma __stream|<var>:=%s
#xcommand ENDTEXT => #pragma __endtext

//----------------------------------------------------------------------------//
REQUEST DBFCDX
//----------------------------------------------------------------------------//
static cDir := "c:\Termine_nachfassen"
static aData :={}
static aAppointments := {}
static oTimer, oDlg, oBrw, oClp


function Main(cAufruf)

   local oFont, n, oFontBrw, oBTNSave
   local nWidth   := 800                        //GetSysMetrics( 0 ) * 1/2
   local nHeight  := GetSysMetrics(1 ) * 1/2
   local cTitle   := "appointment management"
   local nOffset  := 0
   local oBtn, oIcon
   local oFontFLATBTN
   local I := 0

   local aDir := directory( cDir + "\*.*", "DHS" )
   local aDirSub := {}
   *----------------------------------------------------------

   DEFINE FONT oFontFLATBTN NAME 'TAHOMA' SIZE 0,-18
   *----------------------------------------------------------
   //xbrowse(adir)
   FOR I = 1 to len(aDir)
      if aDir[I,1] = "."
      else
         aDirSub := directory( cDir + "\" + aDir[I,1]  +"\*.*","DHS")
         if len(aDirSub)>2
            AADD( aData, { aDir[I,1], "Termine" } )
         else
            AADD( aData, { aDir[I,1], "-" })
         endif

      endif
   NEXT

   AADD(aData, { "[A Meaningful Page Title]",  "Revenue Report - WINHOTEL" } )

   FOR I := 1 to len( aData )
      adata[I,2] += space(300)
   NEXT

   DEFINE ICON oIcon RESOURCE "1" // TN_MAN.RES

   DEFINE FONT oFontFLATBTN NAME 'TAHOMA' SIZE 0, - 18
   DEFINE FONT oFont        NAME 'TAHOMA' SIZE 0, - 18
   DEFINE FONT oFontBrw     NAME 'TAHOMA' SIZE 0, - 26 ITALIC

   DEFINE DIALOG oDlg TITLE cTitle SIZE nWidth, nHeight PIXEL FONT oFont     ;
      ICON oIcon TRUEPIXEL COLORS CLR_MSPURPLE, CLR_WHITE TRANSPARENT

   oDlg:lHelpIcon := .F.

   @ 20, 48 SAY cTitle OF oDlg PIXEL FONT oFontBrw SIZE 800, 40 UPDATE TRANSPARENT

   @ 80, 50 XBROWSE oBrw OF oDlg PIXEL SIZE nWidth - 100, nHeight - 250      ;
      DATASOURCE aData                                                       ;
      HEADERS "Key", "Value"                                                 ;
      COLSIZES 250, 200                                                      ;
      ON DBLCLICK OpenTotalCommander( oBrw )                                 ;
      ON RIGHT CLICK OpenWindowsExplorer( oBrw )                             ;
      AUTOCOLS FASTEDIT LINES NOBORDER

   xBrwColors( oBrw ) // Colores en el xBrowse.

   WITH OBJECT oBrw

      :nStretchCol      := 1

      WITH OBJECT :oCol( 1 )

         :oDatafont := oFont

      END

      WITH OBJECT :oCol( 2 )
         :oDatafont := oFontBrw
         :nEditType := EDIT_GET
      END

      :l2007 := .T.
      :lExcelCellwise   := .F.
      :lFreezeLikeExcel := .F.

   END

   oBrw:CreateFromCode()

   nOffSet := If( oBtn == NIL, 25, nOffSet + 15 + oBtn:nWidth )

   nPixPerChar := GetTextWidth( oDlg:hWnd, "TotalCommander",  oFontFLATBTN:hFont )  + 30

   @ oDlg:nHeight - 70, nOffset BTNBMP oBtn PROMPT "Open TotalCommander"     ;
      SIZE nPixPerChar, 52 PIXEL OF oDlg FLAT COLOR  CLR_WHITE,              ;
      Warning ACTION OpenTotalCommander( oBrw ) // oDlg:End()

   oBtn:cToolTip := "Open TotalCommander"

   nOffSet := If( oBtn == NIL, 25, nOffSet + 15 + oBtn:nWidth )

   nPixPerChar := GetTextWidth( oDlg:hWnd, "Windows Explorer",  oFontFLATBTN:hFont )  + 30

   @ oDlg:nHeight - 70, nOffset BTNBMP oBtn PROMPT "Windows Explorer"        ;
      SIZE nPixPerChar, 52 PIXEL OF oDlg FLAT COLOR  CLR_WHITE, CLR_OLIVE    ;
      ACTION OpenWindowsExplorer( oBrw ) // oDlg:End()

   oBtn:cToolTip := "Windows Explorer"

   nOffSet := If( oBtn == NIL, 25, nOffSet + 15 + oBtn:nWidth )

   nPixPerChar := GetTextWidth( oDlg:hWnd, "Show Details",  oFontFLATBTN:hFont )  + 30

   @ oDlg:nHeight - 70, nOffset BTNBMP oBtn PROMPT "Clipboard"               ;
      SIZE nPixPerChar, 52 PIXEL OF oDlg FLAT COLOR CLR_WHITE, CLR_ROTMETRO  ;
      ACTION Report( oBrw )

   oBtn:cToolTip := "Copy to Clipboard"

   nOffSet := If( oBtn == NIL, 25, nOffSet + 15 + oBtn:nWidth )

   nPixPerChar := GetTextWidth( oDlg:hWnd, "Fixe Termine",  oFontFLATBTN:hFont )  + 30

   @ oDlg:nHeight - 70, nOffset BTNBMP oBtn PROMPT "Fixe Termine"            ;
      SIZE nPixPerChar, 52 PIXEL OF oDlg FLAT COLOR CLR_WHITE, CLR_BLAUMETRO ;
      ACTION FixeTermine( oBrw )

   oBtn:cToolTip := "Fixe and Termine job"

   nPixPerChar := GetTextWidth( oDlg:hWnd, "Cancel", oFontFLATBTN:hFont ) + 30

   nOffSet := 25

   nPixPerChar := GetTextWidth( oDlg:hWnd, "Terminnotiz",  oFontFLATBTN:hFont )  + 30

   @ oDlg:nHeight - 130, nOffset BTNBMP oBtn PROMPT "Terminnotiz"            ;
      SIZE nPixPerChar, 52 PIXEL OF oDlg FLAT COLOR CLR_WHITE, CLR_STEEL     ;
      ACTION Terminnote( oBrw )

   oBtn:cToolTip := "Terminnotiz"

   nOffSet := nWidth - 135 - nPixPerChar

   @ oDlg:nHeight - 130, nOffset BTNBMP oBtn PROMPT "End"                    ;
      SIZE nPixPerChar, 52 PIXEL OF oDlg FLAT COLOR CLR_MSGRAY ,             ;
      Success ACTION( oDlg:End() ) CANCEL

   oBtn:cToolTip := "Exit - Ende - Cancel"

   // SET FONT OF oBtn TO oFontFLATBTN

   ACTIVATE DIALOG oDlg CENTERED ;
      ON INIT (  StartTimer(), oBrw:SetFocus(), .F. )

   RELEASE FONT oFont
   RELEASE FONT oFontBrw

RETURN NIL //( 0 )

STATIC FUNCTION xBrwColors( oBrw, lFoot )

   LOCAL cClrBack

   DEFAULT lFoot          := .F.

   // oBrw:l2007             := .T.

   // oBrw:nRowHeight        := 24
   // oBrw:nHeaderHeight     := 24
   oBrw:lFooter           := lFoot
   
   // oBrw:lRecordSelector     := .F.
   // oBrw:lColDividerComplete := .F.
   // oBrw:lRowDividerComplete := .F.
   
   oBrw:lFlatStyle        := .F.
   oBrw:bClrHeader        := { || { CLR_BLACK, CLR_WHITE, CLR_WHITE } } //RGB( 232, 255, 232 ), RGB( 232, 255, 232 ) }}
   oBrw:lFullGrid         := .F.

   // tira as linhas, nao gostei.
   // oBrw:nRowDividerStyle  := LINESTYLE_NOLINES //DARKGRAY  //LINESTYLE_LIGHTGRAY      //
   // oBrw:nColDividerStyle  := LINESTYLE_NOLINES //LIGHTGRAY // LINESTYLE_NOLINES

   oBrw:nMarqueeStyle := MARQSTYLE_HIGHLROW

   oBrw:bClrStd = { || If( oBrw:KeyNo() % 2 == 0, ;
                         { If( ( oBrw:cAlias )->( Deleted() ), CLR_HRED, CLR_HBLUE ),;
                           RGB( 198, 255, 198 ) }, ;
                         { If( ( oBrw:cAlias )->( Deleted() ), CLR_HRED, CLR_HBLUE ),;
                           RGB( 232, 255, 232 ) } ) }

   oBrw:bClrSel = { || { If( ( oBrw:cAlias )->( Deleted() ), CLR_HRED, CLR_WHITE ),;
                           RGB( 34, 177, 76 ) } }
   
   cClrBack = Eval( oBrw:bClrSelFocus )[ 2 ]

   oBrw:bClrSelFocus  := { || { If( ( oBrw:cAlias )->( Deleted() ), CLR_HRED, CLR_WHITE ),;
                              cClrBack } }

   oBrw:SetColor( CLR_CYAN, RGB( 232, 255, 232 ) )

   // oBrw:SetFont( oFont )

RETURN NIL

//----------------------------------------------------------------------------//


STATIC FUNCTION Terminnote( oBrw )

   LOCAL aDirSub := Directory( cDir + "\" + oBrw:oCol( 1 ):Value  + "\*.*", "DHS" )
   LOCAL cCode :=  cDir + "\" + oBrw:oCol( 1 ):Value  + "\"
   LOCAL cTerminnote := ""

   // xbrowse(aDirSub)

   cTerminnote := cCode + "Termin_" + AllTrim( Str( Len( aDirSub ) ) ) + ".rtf"

   MemoWrit( cTerminnote, '{\rtf1\ansi\ansicpg1252\deff0\nouicompat\deflang1031{\fonttbl{\f0\fnil\fcharset0 Calibri;}}{\*\generator Riched20 6.3.9600}\viewkind4\uc1\pard\sa200\sl276\slmult1\f0\fs22\lang7\par}' )

   Winexec( "write.exe " + cTerminnote )

RETURN NIL

STATIC FUNCTION Report( oBrw )

   LOCAL aDirSub := Directory( cDir + "\" + oBrw:oCol( 1 ):Value  + "\*.*", "DHS" )
   LOCAL cCode :=  cDir + "\" + oBrw:oCol( 1 ):Value  + "\"

   oClp:Clear()
   oClp:SetText( cCode )

RETURN NIL

STATIC FUNCTION OpenWindowsExplorer( oBrw )

   LOCAL cDirSub := "Start " + cDir + "\" + oBrw:oCol( 1 ):Value

   MemoWrit( "EXPL.bat", cDirSub )

   // WINEXEC( "EXPL.bat"  )
   // WINEXEC( "EXPL.bat" , 0 )

   MsgRun( "Wait a Moment, Please! Opening Windows Explorer.",               ;
           "Wait a Moment, Please! Opening Windows Explorer.",               ;
           { || WinExec( "EXPL.bat" ), 3 } )

RETURN NIL

STATIC FUNCTION OpenTotalCommander( oBrw )

   LOCAL cDirSub := "C:\totalcmd\totalcmd.exe /O /T /O= " + cDir + "  " + cDir + "\" + oBrw:oCol( 1 ):Value

   IF .NOT. FILE( "C:\totalcmd\totalcmd.exe" )

      MsgStop( "TOTALCMD BEFINDET SICH NICHT IN ORDNER C:\TOTALCMD", ;
               "TOTALCMD BEFINDET SICH NICHT IN ORDNER C:\TOTALCMD" )

   ELSE

      MemoWrit( "TC.bat", cDirSub )

      // WINEXEC( "TC.bat", 0 )
      MsgRun( "Wait a Moment, Please! Opening TotalCmd.",                    ;
              "Wait a Moment, Please! Opening TotalCmd.",                    ;
              { || WinExec( "TC.bat" ), 3 } )

   ENDIF

RETURN NIL


init procedure PrgInit

   SET DATE ITALIAN
   SET CENTURY ON
   SET TIME FORMAT TO "HH:MM:SS"
   SET EPOCH TO YEAR(DATE())-50

   SET DELETED ON
   SET EXCLUSIVE OFF

   RDDSETDEFAULT( "DBFCDX" )

   XbrNumFormat( 'E', .t. )
   SetKinetic( .f. )
   SetGetColorFocus()

return

//----------------------------------------------------------------------------//

EXIT PROCEDURE PrgExit // IN: 25/03/2022 - Joao.

   DbCommitAll()
   DbUnLockAll()
   DbCloseAll()
   FreeResources()
   Release All
   SysRefresh()
   HB_GCALL( .T. )

   CLEAR MEMORY

   PostQuitMessage( 0 )

   QUIT

RETURN NIL

FUNCTION StartTimer()

   DEFINE CLIPBOARD oClp  OF oDlg

   DEFINE TIMER oTimer OF oDlg ;
      INTERVAL 1000 ;
      ACTION ( RefreshoBrw() )

   ACTIVATE TIMER oTimer

RETURN NIL

FUNCTION FixeTermine( oBrw )

   oTimer:Deactivate()

   xBrowse( aAppointments )

   oTimer:Activate()

RETURN NIL

FUNCTION RefreshoBrw()

   LOCAL I := 0
   LOCAL aDirSub := {}
   LOCAL aDir := Directory( cDir + "\*.*", "DHS" )
   LOCAL cTempDay := ""
   LOCAL cTempheute := ""
   LOCAL dDate := Date() - 1
   LOCAL aDirTemp := {}
   LOCAL cSrc := ""
   LOCAL cDst := ""
   LOCAL nIdx := 0

   oTimer:Deactivate()

   dDate      := Date() - 1
   cTempDay   := SubStr( AllTrim( DToS( dDate ) ), 3, 2 ) + "-" + SubStr( AllTrim( DToS( dDate ) ), 5, 2 ) + "-" + SubStr( AllTrim( DToS( dDate ) ), 7, 2 )   + "_" + fWH_KURZDATUM( dDate )
   dDate      := Date()
   cTempheute := SubStr( AllTrim( DToS( dDate ) ), 3, 2 ) + "-" + SubStr( AllTrim( DToS( dDate ) ), 5, 2 ) + "-" + SubStr( AllTrim( DToS( dDate ) ), 7, 2 )   + "_" + fWH_KURZDATUM( dDate )

   IF lIsDir( cDir + "\" + cTempDay ) = .T.

      // Dateien vom Vortag in den heutigen Tag verschieben und Verzeichnis löschen
      aDirTemp := Directory( cDir + "\" + cTempDay + "\*.*", "DHS" )

      IF Len( aDirTemp ) = 2

         IF .NOT. LRMDIR( cDir + "\" + cTempDay )

            ? "Ordner kann nicht gelöscht werden  " + cDir + "."

         ENDIF

      ENDIF

      FOR I := 1 TO Len( aDirTemp )

         IF aDirTemp[ I, 1 ] <> "."

            cSrc  := cDir + "\" + cTempDay + "\" + aDirTemp[ I, 1 ]

            cDst  := cDir + "\" + cTempheute + "\" + aDirTemp[ I, 1 ]

            ? cSrc + " ----> " + cDst

            IF  ( CopyFile( cSrc, cDst,  .T. ) ) = .T.

               FErase( cSrc )

            ELSE  // ??

            ENDIF

         ENDIF

      NEXT

   ENDIF

   // angezeigte Termine aktualisieren
   aAppointments := {}

   aData := {}

   FOR I = 1 TO Len( aDir )

      IF aDir[ I, 1 ] = "."

      ELSE

         aDirSub := Directory( cDir + "\" + aDir[ I, 1 ]  + "\*.*", "DHS" )

         IF Len( aDirSub ) > 2

            FOR nIdx := 3 TO Len( aDirSub )

               IF Left( aDirSub[ nIdx, 1 ], 1 ) = "#"

                  AAdd( aAppointments,   { aDir[ I, 1 ], aDirSub[ nIdx, 1 ] }  )

               ENDIF

            NEXT

            AAdd( aData, { aDir[ I, 1 ], "Termine" } )

         ELSE

            AAdd( aData, { aDir[ I, 1 ], "-" } )

         ENDIF

      ENDIF

   NEXT

   oBrw:aArraydata := aData
   oBrw:refresh()

   oTimer:Activate()

RETURN NIL

FUNCTION fWH_KURZDATUM( dDatum )

   LOCAL cKurzDatum := ""

   cKurzDatum := num2day( DoW( dDatum ) )

RETURN( cKurzDatum )

FUNCTION num2day( nTag )

   LOCAL cTag     := ""
   LOCAL cTITLE   := " Zimmerstammdaten "

   IF nTag = 1
      cTag = "SONNTAG"
   ELSEIF nTag = 2
      cTag = "Montag"
   ELSEIF nTag = 3
      cTag = "Dienstag"
   ELSEIF nTag = 4
      cTag = "Mittwoch"
   ELSEIF nTag = 5
      cTag = "Donnerstag"
   ELSEIF nTag = 6
      cTag = "Freitag"
   ELSEIF nTag = 7
      cTag = "Samstag"
   ENDIF

RETURN( cTag )
// fin / ende / end - kapiabafwh@gmail.com

