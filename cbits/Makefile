WINEP = WINEPREFIX=/home/bnikolic/.wine-expy32  wine

CC =  ${WINEP} C:\\MinGW\\bin\\gcc.exe

CCFLAGS = -IC:\\Python27\\include

LDFLAGS = -mdll -Wl,--add-stdcall-alias 

expy.o : expy.c
	${CC} ${CCFLAGS} -c  $< -o $@

expy.dll : expy.o
	${CC} ${LDFLAGS}    $<   -o $@ /home/bnikolic/.wine-expy32/drive_c/windows/syswow64/python27.dll /home/bnikolic/oss/ExPy/cbits/xlcall32.lib


