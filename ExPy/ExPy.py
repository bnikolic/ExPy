# Copyright (c) Bojan Nikolic <bojan@bnikolic.co.uk> 2010, 2013
# License: See the LICENSE file distributed with this git repository
"""
Registration of functions with Excel and their dispatch back into Python
"""

import __main__
import types
import ctypes
import xlltypes

__version__ = 0.4

# Holds ancillary data about all registred functions
fnregstry = {}

# This is the list of all results, used to prevent garbage collection
# of results passed back to Excel. This will eventually cause memory
# problems. See github issue #2
reslist = []


def nArgs(formula):
    """
    Number of arguments that an excel-available function takes
    """
    fname=formula[1:].split("(")[0]
    return (fnregstry[fname][0],)

def dispatch(formula,
             firstarg=None,
             restargs=None):
    """
    Dispatch an excel-available function to its Python implementation
    """
    fname=formula[1:].split("(")[0]
    arglist=[]
    if firstarg is not None:
        arglist.append(xlltypes.XLOper.from_address(firstarg))
    if restargs is not None:
        arglist.extend([xlltypes.XLOper.from_address(x) for x in restargs])
    try:
        res=getattr(__main__, fname)(*arglist)
    except BaseException, e:
        res=str(e)
    res=xlltypes.mkXLOper(res)
    if res:
        reslist.append(res)
        return (ctypes.addressof(res),)
    else:
        return (0,)

def register(fnname, nargs):
    """
    Register a function with excel
    """
    if type(fnname) == types.FunctionType:
        fnname=fnname.func_name
    r=RegisterFnXLL(str(getXLLDLL()),
                    "ExPyDispatch",
                    "P"+"P"*nargs+"#",
                    fnname)
    fnregstry[fnname]=(nargs,)
    return r    
    

def RegisterFnXLL(*fargs):
    """
    Register a function with Excel
    """
    args=(ctypes.POINTER(xlltypes.XLOper)*4)()
    xargs=[xlltypes.mkXLOper(x) for x in fargs]
    for i in range(4):
        args[i]=ctypes.POINTER(xlltypes.XLOper)(xargs[i])
    rr=xlltypes.XLOper()
    r=ctypes.windll.xlcall32.Excel4v(149, #This is xlfRegister
                                     ctypes.byref(rr),
                                     4,
                                     args)
    fnregstry[fargs[3]]=(len(fargs[2])-2,)
    return r

def mkEx4(n=0):
    p=ctypes.CFUNCTYPE(ctypes.c_int,
                       ctypes.c_int,
                       ctypes.POINTER(xlltypes.XLOper),
                       ctypes.c_int,
                       *([ctypes.POINTER(xlltypes.XLOper)]*n))
    return p(ctypes.windll.xlcall32.Excel4)


def getXLLDLL():
    rr=xlltypes.XLOper()
    fn=ctypes.windll.xlcall32.Excel4v
    r=fn(9|0x4000, #xlGetName
         ctypes.byref(rr),
         0,
         None)
    return rr

def xlStack():
    rr=xlltypes.XLOper()
    fn=ctypes.windll.xlcall32.Excel4v
    r=fn(1|0x4000, #xlStack
         ctypes.byref(rr),
         0,
         None)
    return rr


def Splash():
    title="BN Algortihms Excel/Python add-in "
    blurb= """
    This is the ExPy Python add-in for Microsoft Excel 

    Version: %g

    Visit us at: http://www.bnikolic.co.uk/expy
    Inquiries/bug reports to: webs@bnikolic.co.uk

    """% (__version__,)
    lic="""
Copyright (C) Copyright (c) Bojan Nikolic <bojan@bnikolic.co.uk> 2010, 2013, 2014

Redistribution and use in source and binary forms, with or without
modification, are permitted provided that the following conditions are
met:

    * Redistributions of source code must retain the above copyright
       notice, this list of conditions and the following disclaimer.

    * Redistributions in binary form must reproduce the above
       copyright notice, this list of conditions and the following
       disclaimer in the documentation and/or other materials provided
       with the distribution.

    * Neither the name of the ExPy Developers nor the names of any
       contributors may be used to endorse or promote products derived
       from this software without specific prior written permission.

THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS
"AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT
LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR
A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT
OWNER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL,
SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT
LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE,
DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY
THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
(INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE
OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
    
    """
    ctypes.windll.user32.MessageBoxA(0,
                                     blurb+lic, 
                                     title, 
                                     0x40)
