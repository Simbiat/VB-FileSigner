# VB-FileSigner
Function to sign/encrypt files using Russian Cryptographic Software "Signatura". It allows signing, unsigning, encrypting and decrypting of files by mask using different crypto-keys in a batch mode. Requiers to be started with elevated rights and will not cache eToken password.

The setting is referencing a line in .ini file which should follow these rules:
spb chooser should be used in way like this `spbaddress:::regexmask;;;spdaddress2:::regexmask2` and so on. `;;;` is used as delimiter for addresses, `:::` is used as delimiter for address and filemask. If filemask is " " (space) - only folder check is done for the address. Do not use `*.*` or similar masks unless absolutely necessary: in this case it's recommended for the mask to be the last one in the list

Example:

`spbchooser=ias:::^(.*ez.{3}765\..{1}30\.xml)|(tps.*\.xml)|(F0409350.*\.xml)|(.*ez.*765\..{1}30\..*\.xml)|(.*\..*765.*\.xml)|(.*\..{1}30)$;;;otzi:::^(q.*\.zip)|(765.*\.zip)|(otzi.*\.zip)|(.*\.txt)$;;;expl:::^(expl.*\.zip)|(.*\.xls)$;;;grkc:::^(grkc.*\.zip)$;;;mou:::^(mou.*\.zip)$;;;ueko:::^(ueko.*\.zip)$;;;ubzi:::^(ubzi.*\.zip)$;;;upsir::: ;;;kcntr:::^(kcntr.*\.zip)$
`

Uses [HTA Logger](https://github.com/Simbiat/HTA-Logging), [FileMover](https://github.com/Simbiat/VBS-Filemover), ReadIni (available in the logger) and FileDel function which can be replaced with regular Delete. 
