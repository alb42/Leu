all:
	fpc4aros.sh -FUunits -Fu../MUIClass/src -Fuspread Leu.pas
	
amiga: 
	fpc4amiga.sh -FUunits -Fu../MUIClass/src -Fuspread Leu.pas