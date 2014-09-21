/* 
Code here comes from Chicken Scheme, git describe: 4.5.0rc1-2450-g627ea21
;
; Copyright (c) 2008-2014, The CHICKEN Team
; Copyright (c) 2000-2007, Felix L. Winkelmann
; All rights reserved.
;
; Redistribution and use in source and binary forms, with or without modification, are permitted provided that the following
; conditions are met:
;
;   Redistributions of source code must retain the above copyright notice, this list of conditions and the following
;     disclaimer.
;   Redistributions in binary form must reproduce the above copyright notice, this list of conditions and the following
;     disclaimer in the documentation and/or other materials provided with the distribution.
;   Neither the name of the author nor the names of its contributors may be used to endorse or promote
;     products derived from this software without specific prior written permission.
;
; THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND ANY EXPRESS
; OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY
; AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDERS OR
; CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR
; CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR
; SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY
; THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR
; OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE
; POSSIBILITY OF SUCH DAMAGE.
*/
#ifndef __EXPY_CHUTILS_H__
#define __EXPY_CHUTILS_H__

/* These are wrappers around the following idiom:
 *    assert(SOME_PRED(obj));
 *    do_something_with(obj);
 * This works around the fact obj may be an expression with side-effects.
 *
 * To make this work with nested expansions, we need semantics like
 * (let ((x 1)) (let ((x x)) x)) => 1, but in C, int x = x; results in
 * undefined behaviour because x refers to itself.  As a workaround,
 * we keep around a reference to the previous level (one scope up).
 * After initialisation, "previous" is redefined to mean "current".
 */
# define C_VAL1(x)                 C__PREV_TMPST.n1
# define C_VAL2(x)                 C__PREV_TMPST.n2
# define C__STR(x)                 #x
# define C__CHECK_panic(a,s,f,l)                                       \
  ((a) ? (void)0 :                                                     \
   C_panic_hook("Low-level type assertion " s " failed at " f ":" C__STR(l)))
# define C__CHECK_core(v,a,s,x)                                         \
  ({ struct {                                                           \
      typeof(v) n1;                                                     \
  } C__TMPST = { .n1 = (v) };                                           \
    typeof(C__TMPST) C__PREV_TMPST=C__TMPST;                            \
    C__CHECK_panic(a,s,__FILE__,__LINE__);                              \
    x; })
# define C__CHECK2_core(v1,v2,a,s,x)                                    \
  ({ struct {                                                           \
      typeof(v1) n1;                                                    \
      typeof(v2) n2;                                                    \
  } C__TMPST = { .n1 = (v1), .n2 = (v2) };                              \
    typeof(C__TMPST) C__PREV_TMPST=C__TMPST;                            \
    C__CHECK_panic(a,s,__FILE__,__LINE__);                              \
    x; })
# define C_CHECK(v,a,x)            C__CHECK_core(v,a,#a,x)
# define C_CHECK2(v1,v2,a,x)       C__CHECK2_core(v1,v2,a,#a,x)

#endif
