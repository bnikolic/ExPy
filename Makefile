# Copyright (c) Bojan Nikolic <bojan@bnikolic.co.uk> 2014
#
# License: BSD-3, see the LICENSE file distributed at the root of this
# git repository
# 
# Top-level makefile for ExPy

GITREV := $(shell git describe --dirty)

PACK := ExPy-${GITREV}.zip

default: ${PACK}

PYFILES  := ExPy/ExPy.py ExPy/xlltypes.py
XLLFILES := cbits/expy.xll

${PACK}: ${XLLFILES} ${PYFILES}
	cd cbits && $(MAKE)
	rm -f $@
	zip $@ -j ${XLLFILES} ${PYFILES}
