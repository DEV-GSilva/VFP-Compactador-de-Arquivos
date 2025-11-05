* Compactar um diretório
cNovoZip = "C:\Temp\Compactado.Zip"
cDirOrig = "C:\Temp\VFP-Compactador-de-Arquivos"

Strtofile(Chr(0x50)+Chr(0x4B)+Chr(0x05)+Chr(0x06)+Replicate(Chr(0),18),cNovoZip)

oShell = Createobject("Shell.Application")

For Each oArchi In oShell.NameSpace(cDirOrig).Items
   oShell.NameSpace(cNovoZip).CopyHere(oArchi)
Endfor


* Descompactar um diretório
cArquivoZip = "C:\Temp\Compactado.Zip"
cDirDestino = "C:\Temp\Descompactado"

oShell = Createobject("Shell.Application")

For Each oArchi In oShell.NameSpace(cArquivoZip).Items

   oShell.NameSpace(cDirDestino).CopyHere(oArchi)
EndFor