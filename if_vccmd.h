/*
 * if_vccmd.h Interface between Vim and Visual C++
 *
 * Code changes neccessary for this to work are minimal within Vim:
 *
 *  1.      gui_w32.c      :633        OnDestroy() calls vc_disconnect()
 *
 *  2.      ex_docmd.c	   :2399       case CMD_vccmd:
 *		                                    do_vccmd(ea.arg);
 *		                                    break;
 *
 *  3.      ex_cmds.h      :346        EXCMD(CMD_vccmd,	"vc",	    WORD1|NEEDARG), 
 *
 *  4.      vim.h          :1001       #include "if_vccmd.h"
 *
 *  TODO: If submitted add preprocessor defs!!!
 */

#ifndef _if_vccmd_h_
#define _if_vccmd_h_

#ifdef __cplusplus
extern "C"{
#endif 

#include "vim.h"

void do_vccmd __ARGS( (char_u *arg ) ); 
static BOOL vc_connect  ( void );
void vc_disconnect  ( void );

#ifdef __cplusplus
}
#endif

#endif // _if_vccmd_h_
