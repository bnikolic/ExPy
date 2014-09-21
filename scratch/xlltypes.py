# Copyright (c) Bojan Nikolic <bojan@bnikolic.co.uk> 2010, 2013
# License: See the LICENSE file distributed with this git repository
"""
Definition of types for interaction with excel
"""

import ctypes
#import numpy

class XLRef(ctypes.Structure):
    _fields_=[("rwFirst", ctypes.c_int16),
              ("rwLast", ctypes.c_int16),
              ("colFirst", ctypes.c_int8),
              ("colLast", ctypes.c_int8)]

class XLOper(ctypes.Structure):

    def float(self):
        if self.xltype==xltypeNum:
            return self.val.double
        else:
            raise BaseException("XLOper is not of type double")

    def array(self):
        if self.xltype==xltypeMulti:
            r=[]
            for i in range(self.val.array.rows):
                rr=[]
                for j in range(self.val.array.columns):
                    rr.append(self.val.array.lparray[i*self.val.array.columns+j].float())
                r.append(rr)
            return r
        else:
            raise BaseException("XLOper is not of array type")

    def __str__(self):
        if self.xltype==xltypeStr:
            slen=ord(self.val.xstr[0])
            return str(self.val.xstr[1:1+slen])
        if self.xltype==xltypeMulti:
            return "Array %iR X %iC" (self.val.array.rows,
                                      self.val.array.columns)
        else:
            return "No simple text representation"

class SRef(ctypes.Structure):
    """
    Single reference
    """
    _fields_=[("count", ctypes.c_int16),
              ("ref", XLRef)]

class XLMRef(ctypes.Structure):
    _fields_=[("count", ctypes.c_int16),
              ("reftbl", ctypes.POINTER(XLRef))]

class MRef(ctypes.Structure):
    """
    Multiple references
    """
    _fields_=[("lpmref", ctypes.POINTER(XLMRef)),
              ("idSheet", ctypes.c_int32)]

class Array(ctypes.Structure):
    _fields_=[("lparray",  ctypes.POINTER(XLOper)),
              ("rows", ctypes.c_int16),
              ("columns", ctypes.c_int16)]


class ValFlow(ctypes.Union):
    _fields_=[("level",  ctypes.c_int16),
              ("tbctrl",  ctypes.c_int16),
              ("idSheet", ctypes.c_int32),]

class Flow(ctypes.Structure):
    _fields_=[("valflow", ValFlow),
              ("rw", ctypes.c_int16),
              ("col", ctypes.c_int8),
              ("xlflow", ctypes.c_int8)]


class H(ctypes.Union):
    _fields_=[("lpbData", ctypes.POINTER(ctypes.c_int8)),
              ("hdata",  ctypes.c_void_p)]

class BigData(ctypes.Structure):
    _fields_=[("h", H),
              ("cbData",  ctypes.c_int32)]

class Val(ctypes.Union):
    _fields_=[("double", ctypes.c_double),
              ("xstr", ctypes.c_char_p),
              ("xbool", ctypes.c_int16),
              ("err",   ctypes.c_int16),
              ("xint",   ctypes.c_int8),
              ("sref",  SRef),
              ("mref",  MRef),
              ("array", Array),
              ("flow", Flow),
              ("bigdata", BigData)]

XLOper._fields_=[("val", Val),
                 ("xltype", ctypes.c_int16)]


xltypeNum=0x0001
xltypeStr=0x0002 
xltypeBool=0x0004
xltypeRef=0x0008
xltypeErr=0x0010
xltypeFlow=0x0020
xltypeMulti=0x0040
xltypeMissing=0x0080
xltypeNil=0x0100
xltypeSRef=0x0400
xltypeInt=0x0800
xlbitXLFree=0x1000
xlbitDLLFree=0x4000
xltypeBigData=  (xltypeStr | xltypeInt)    
    

def fillXLOper(r, d):
    """
    Make an XLOper object
    """
    if type(d) == str:
        if len(d)>255:
            raise BaseException("Strings longer than 255 characters can not be passed into XLOPER")
        r._s=ctypes.create_string_buffer(len(d)+1)
        r._s[1:len(d)+1]=d[:]
        r._s[0]=chr(len(d))
        r.xltype=xltypeStr
        r.val.xstr=ctypes.c_char_p(r._s.raw)
    elif (type(d) == int) or (type(d) == float ) or type(d) == numpy.float64:
        r.xltype=xltypeNum
        r.val.double=float(d)
    elif type(d)==numpy.ndarray:
        nn=len(d.shape)
        if nn==1:
            nrows=d.shape[0]
            ncolumns=1
        elif len(d.shape)==2:
            nrows,ncolumns=d.shape
        else:
            raise BaseException("Can't convert more than 2d array:" +str(d.shape))            
        r.xltype=xltypeMulti
        r.val.array.rows=nrows
        r.val.array.columns=ncolumns
        print "Got an array" , nrows, ncolumns
        atype=XLOper*(nrows*ncolumns)
        r.val.array.lparray=atype()
        for i in range(nrows):
            for j in range(ncolumns):
                if nn==1:
                    fillXLOper(r.val.array.lparray[i*ncolumns+j], d[i])
                else:
                    fillXLOper(r.val.array.lparray[i*ncolumns+j], d[i,j])
    else:
        raise BaseException("Don't know how to convert objects of type:" +str(type(d)))
    return r

def mkXLOper(d):
    """
    Make an XLOper object
    """
    r=XLOper()
    try:
        fillXLOper(r, d)
        return r
    except BaseException, e:
        print e
        return None
    

    
