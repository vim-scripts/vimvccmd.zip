/* vi:set ts=8 sts=4 sw=4:
 *
 * VIM - Vi IMproved		by Bram Moolenaar
 *
 * if_vccmd.c Interface between Vim and Visual C++ 
 *
 */

#define USE_VCCMDS
#ifdef USE_VCCMDS

#include <windows.h>
#include "if_vccmd.h"

#include <ObjModel\appauto.h>
#include <ObjModel\dbgauto.h>
#include <ObjModel\textauto.h>

#include <initguid.h>
#include <ObjModel\appguid.h>
#include <ObjModel\dbgguid.h>
#include <ObjModel\textguid.h>

#define _ATL_APARTMENT_THREADED
#define _ATL_DEBUG_QI

#include <atlbase.h>
extern CComModule _Module;

#include <atlcom.h>

static CComPtr< IApplication > s_pApplication;
static BOOL s_bComLibsInited = FALSE;

/////////////////////////////////////////////////////////////////////

BOOL vc_connect( void )
{
    if ( !s_bComLibsInited )
    {

#if _WIN32_WINNT >= 0x0400 & defined( _ATL_FREE_THREADED )
        HRESULT hr = CoInitializeEx( NULL, COINIT_MULTITHREADED );
#else
        HRESULT hr = CoInitialize( NULL );
#endif
        _ASSERTE( SUCCEEDED( hr ) );

        s_bComLibsInited = TRUE;
    }

    CLSID clsid;
    BOOL  bSuccess = FALSE;
    CComPtr< IUnknown > pUnk;

    //  If s_pApplication is already init'ed the user may have closed
    //  the instance of VC they were using for previous vc_connect
    //  try it again

    if ( s_pApplication != NULL )
    {
        s_pApplication.Release( );
    }

    HRESULT hr = CLSIDFromProgID( OLESTR( "MSDEV.APPLICATION" ), &clsid );

    if ( SUCCEEDED( hr ) ) 
    {
        //  This function retrives the first running instance of VC it comes across

        hr = GetActiveObject( clsid, NULL, &pUnk );

        if ( SUCCEEDED( hr ) )
        {
            hr = pUnk->QueryInterface( IID_IApplication, ( LPVOID * )&s_pApplication );

            if ( SUCCEEDED( hr ) )
            {
                bSuccess = TRUE;
            }
        }

        //  Start VC up...

        else
        {
            hr = CoCreateInstance( clsid, NULL, CLSCTX_LOCAL_SERVER, IID_IApplication, ( LPVOID* )&s_pApplication );

            if ( SUCCEEDED( hr ) )
            {
                bSuccess = TRUE;
            }
        }
    }

    return bSuccess;
}

/////////////////////////////////////////////////////////////////////

void vc_disconnect( void )
{
    if ( s_bComLibsInited )
    {
        //  Have to explicitly invoke Release of the ATL wrapper for IApplication
        //  before calling CoUninitialize( ) for global objects...

        if ( s_pApplication != NULL )
        {
            s_pApplication.Release( );
        }

        CoUninitialize( );
    }
}

/////////////////////////////////////////////////////////////////////
    
void do_vccmd( char_u* pszVCCommand )
{
    //  CComPtr has overloaded operator T*

    if ( s_pApplication == NULL )
    {
        if ( !vc_connect( ) )
        {
	    EMSG( "Couldn't start/find Visual C++" );
            return;
        }
    }

    ASSERT( s_pApplication );

    /*
    if ( bBringToTop )
    {
        HRESULT hr = s_pApplication->put_Visible( VARIANT_TRUE );

        if ( FAILED ( hr ) )
        {
            //  User may have closed MSDEV and re-opened it

            if ( !vc_connect( ) )
            {
                EMSG( "Couldn't start/find Visual C++" );
                return;
            }

            //  Try it again, it it fails here something is WRONG

            hr = s_pApplication->put_Visible( VARIANT_TRUE );

            if ( FAILED ( hr ) )
            {
                EMSG( "Couldn't bring Visual C++ to top" );
            }
        }
    }
    */

    CComBSTR bstrCmd = ( char* )pszVCCommand;
    
    //  NOTE: This command will 'block' VIM if it involves
    //  showing a modal dialog within Visual C++

    HRESULT hr = s_pApplication->ExecuteCommand( bstrCmd );

    if ( FAILED ( hr ) )
    {
        //  User may have closed MSDEV and re-opened it

        if ( !vc_connect( ) )
        {
	    EMSG( "Couldn't start/find Visual C++" );
            return;
        }

        //  Try it again, it it fails here something is WRONG

        hr = s_pApplication->ExecuteCommand( bstrCmd );

        if ( FAILED ( hr ) )
        {
            EMSG2( "Couldn't send command %s to Visual C++", ( char_u* )pszVCCommand );
        }
    }
}

/////////////////////////////////////////////////////////////////////

void vc_open_file( char_u* pFileName, int nStartingLine, int nStartingColumn )
{
    if ( s_pApplication == NULL )
    {
        if ( !vc_connect( ) )
        {
	    EMSG( "Couldn't start/find Visual C++" );
            return;
        }
    }

    ASSERT( s_pApplication );

    CComQIPtr<IDocuments, &IID_IDocuments> pDocs;
    CComPtr<IDispatch> pDispDocs;

    HRESULT hr = s_pApplication->get_Documents( &pDispDocs );

    if ( FAILED ( hr ) )
    {
        //  User may have closed MSDEV and re-opened it

        if ( !vc_connect( ) )
        {
	    EMSG( "Couldn't start/find Visual C++" );
            return;
        }

        //  Try it again, it it fails here something is WRONG

        HRESULT hr = s_pApplication->get_Documents( &pDispDocs );

        if ( FAILED ( hr ) )
        {
            EMSG( "Couldn't get Visual C++ documents" );
            return;
        }
    }

    ASSERT( s_pApplication );

    //  All calls from this point shouldn't fail because of
    //  a previously closed MSDEV.  They still can fail 
    //  for other reasons however.

    pDocs = pDispDocs;

    if ( pDocs )
    {
        CComVariant vtDocType       = "Auto";
        CComVariant vtBoolReadOnly  = FALSE;

        CComBSTR bstrFile = ( char* )pFileName;
        CComPtr<IDispatch> pDispDoc;

        hr = pDocs->Open( bstrFile, vtDocType, vtBoolReadOnly, &pDispDoc );

        if ( FAILED ( hr ) )
        {
	    EMSG2( "Couldn't open file %s in Visual C++", ( char_u* ) pFileName );
        }

        //  TODO: Set the line # and col # of the opened file
    }
}

/////////////////////////////////////////////////////////////////////

void vc_set_breakpoint( char_u* pFileName, int nLineNum )
{
    if ( s_pApplication == NULL )
    {
        if ( !vc_connect( ) )
        {
	    EMSG( "Couldn't start/find Visual C++" );
            return;
        }
    }

    ASSERT( s_pApplication );

    CComQIPtr<IDebugger, &IID_IDebugger> pDebugger;
    CComQIPtr<IBreakpoints, &IID_IBreakpoints> pBrkPnts;

    CComPtr<IDispatch> pDispDebugger;
    CComPtr<IDispatch> pDispBrkPnts;

    HRESULT hr  = s_pApplication->get_Debugger( &pDispDebugger );

    if ( FAILED ( hr ) )
    {
        //  User may have closed MSDEV and re-opened it

        if ( !vc_connect( ) )
        {
	    EMSG( "Couldn't start/find Visual C++" );
            return;
        }

        //  Try it again, it it fails here something is WRONG

        HRESULT hr  = s_pApplication->get_Debugger( &pDispDebugger );

        if ( FAILED ( hr ) )
        {
            EMSG( "Couldn't connect to Visual C++'s debugger" );
            return;
        }
    }

    ASSERT( s_pApplication );

    //  All calls from this point shouldn't fail because of
    //  a previously closed MSDEV.  They still can fail 
    //  for other reasons however.

    pDebugger   = pDispDebugger;

    //  Open the file requested.  It will come to 
    //  the top if already opened

    vc_open_file( pFileName, nLineNum, 0 );

    CComPtr<IDispatch> pDispDoc;
    CComPtr<IGenericDocument> pGenericDoc;

    s_pApplication->get_ActiveDocument( &pDispDoc );

    if ( pDispDoc )
    {
        pDispDoc.QueryInterface( &pGenericDoc );
    }

    pDebugger->get_Breakpoints( &pDispBrkPnts );
    pBrkPnts = pDispBrkPnts;

    if ( pBrkPnts && pGenericDoc )
    {
        CComQIPtr<ITextDocument, &IID_ITextDocument> pTextDoc;
        CComPtr<IDispatch> pDispBrkPnt;
        CComPtr<IDispatch> pDispSel;
        CComQIPtr<ITextSelection, &IID_ITextSelection> pSel;
        hr = pGenericDoc->put_Active( VARIANT_TRUE );
        pTextDoc = pGenericDoc;

        if ( pTextDoc )
        {
            hr = pTextDoc->get_Selection( &pDispSel );
            if ( SUCCEEDED( hr ) && pDispSel )
            {
                    CComVariant varSelMode = dsMove;
                    CComVariant varSel;
                    pSel = pDispSel;
                    hr = pSel->GoToLine( nLineNum, varSelMode );

                    varSel = nLineNum;
                    hr = pBrkPnts->AddBreakpointAtLine( varSel, &pDispBrkPnt );
            }
        }
    }
}

#endif
