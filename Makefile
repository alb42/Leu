FPCOPT=-XX -CX
FPCFLAGS=-FUunits -Fu../MUIClass/src -Fuspread $(FPCOPT) -Xs
FPCFLAGSR=-FUunits -Fu../MUIClass/src -Fuspread $(FPCOPT)

all:
	fpc4aros.sh $(FPCFLAGS) Leu.pas
	
amiga: 
	fpc4amiga.sh $(FPCFLAGS) Leu.pas
amigafpu: 
	fpc4amigafpu.sh $(FPCFLAGS) Leu.pas

os4: 
	fpc4os4.sh $(FPCFLAGS) Leu.pas
	
mos: 
	fpc4mos.sh $(FPCFLAGS) Leu.pas
	
arosarm:
	fpc4arosarm.sh $(FPCFLAGS) Leu.pas
	
aros64: 
	fpc4aros64.sh $(FPCFLAGS) Leu.pas
	
clean:
	rm -f Leu units/* *.lha

release:
	make clean all
	cp Leu Release/i386-aros
	make clean amiga
	cp Leu Release/m68k-amiga
	make clean os4
	cp Leu Release/powerpc-amiga
	make clean mos
	cp Leu Release/powerpc-morphos
	make clean arosarm
	cp Leu Release/arm-aros
	make clean aros64
	cp Leu Release/x86_64-aros
	make clean
