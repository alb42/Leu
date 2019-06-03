all:
	fpc4aros.sh -FUunits -Fu../MUIClass/src -Fuspread Leu.pas
	
amiga: 
	fpc4amiga.sh -FUunits -Fu../MUIClass/src -Fuspread Leu.pas

os4: 
	fpc4os4.sh -FUunits -Fu../MUIClass/src -Fuspread Leu.pas
	
mos: 
	fpc4amiga.sh -FUunits -Fu../MUIClass/src -Fuspread Leu.pas
	
arosarm:
	fpc4arosarm.sh -FUunits -Fu../MUIClass/src -Fuspread Leu.pas
	
aros64: 
	fpc4amiga.sh -FUunits -Fu../MUIClass/src -Fuspread Leu.pas