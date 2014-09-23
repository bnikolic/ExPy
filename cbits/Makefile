# Copyright (c) Bojan Nikolic <bojan@bnikolic.co.uk> 2014
#
# License: BSD-3, see the LICENSE file distributed at the root of this
# git repository
# 
# Quick makefile for the c parts of ExPy. Note: I build in Linux but
# using wine-hosted MinGW. YMMV. 

WINEP = ${HOME}/.wine-expy32
WEXEC = WINEPREFIX=${WINEP}  wine

CC =  ${WEXEC} C:\\MinGW\\bin\\gcc.exe

CCFLAGS = -IC:\\Python27\\include

LDFLAGS = -mdll -Wl,--add-stdcall-alias 

CFILES = expy.c

default: expy.xll

clean:
	rm ${CFILES:.c=.o} expy.xll

%.o : %.c
	${CC} ${CCFLAGS} -c  $< -o $@

expy.xll : ${CFILES:.c=.o}
	${CC} ${LDFLAGS}    $<   -o $@ ${WINEP}/drive_c/windows/syswow64/python27.dll ./xlcall32.lib


