


Get-MsolAccountSku | Where {$_.AccountSkuId -eq "sdcountycagov:ENTERPRISEPACK_GOV"} | Out-File ./output.txt