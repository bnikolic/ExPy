# Copyright (c) Bojan Nikolic <bojan@bnikolic.co.uk> 2010, 2013
# License: See the LICENSE file distributed with this git repository
"""

"""

import types
import ctypes
import xlltypes

fnregstry = {}

#This is the list of all results
reslist = []

def Splash():
    title="BN Algortihms Excel/Python add-in "
    blurb= """
    This is the Python add-in for Excel 

    Version: 0.3

    Visit us at: http://www.bnikolic.co.uk/expy
    Inquiries/bug reports to: webs@bnikolic.co.uk

    """
    ctypes.windll.user32.MessageBoxA(0,
                                     blurb, 
                                     title, 
                                     0x40)

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

