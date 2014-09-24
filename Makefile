# Copyright (c) Bojan Nikolic <bojan@bnikolic.co.uk> 2014
#
# License: BSD-3, see the LICENSE file distributed at the root of this
# git repository
# 
# Top-level makefile for ExPy

default: ExPy.zip

PYFILES= ExPy/ExPy.py ExPy/xlltypes.py
XLLFILES= cbits/expy.xll

ExPy.zip: ${XLLFILES} ${PYFILES}
	cd cbits && $(MAKE)
	rm -f ExPy.zip
	zip ExPy.zip -j ${XLLFILES} ${PYFILES}
