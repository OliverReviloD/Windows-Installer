
Function Main {
    cls

    $MsiFile = "C:\.....\Dell-Dock-Firmware-Updater_v1.0.0\Source\DockWrapper_v1.0.0.msi"
    $MsiFile = "C:\__B\Downloads\Intune-Apps\Dell-Dock-Firmware-Updater_v1.0.0\Source\DockWrapper_v1.0.0.msi"
    $MsiFileReturn =  Get-MsiProperties -MsiFilePath $MsiFile -MyDebug $false
    $MsiFileReturn | FL | Out-String | Write-Host
    <#
            $MsiFileReturn | FL | Out-String | Write-Host
 
            MsiFilePath               : C:\Dell\Dell-Dock-Firmware-Updater_v1.0.0\Source\DockWrapper_v1.0.0.msi
            MsiPropertyManufacturer   : Dell
            MsiPropertyProductName    : Dell Dock Firmware Update Package
            MsiPropertyProductVersion : 1.0.2
            MsiPropertyProductCode    : {92E40D32-4403-466D-89A3-DA7D504A3A41}
    #>
    remove-variable MsiFileReturn
    remove-variable MsiFile
    # get-variable -Scope Script
    }

function Invoke-Method ($Object, $MethodName, $ArgumentList) {
  return $Object.GetType().InvokeMember($MethodName, 'Public, Instance, InvokeMethod', $null, $Object, $ArgumentList)
}

function Get-Property ($Object, $PropertyName, [object[]]$ArgumentList) {
  return $Object.GetType().InvokeMember($PropertyName, 'Public, Instance, GetProperty', $null, $Object, $ArgumentList)
}


Function Get-MsiProperties {
    Param(  [parameter(Mandatory = $true)][String]$MsiFilePath, 
            [parameter(Mandatory = $false)][bool]$MyDebug )
    
    <#
        example

            $MsiFileReturn = Get-MSI_ProductCode_and_Version -MsiName "C:\Dell-Dock-Firmware-Updater_v1.0.0\Source\DockWrapper_v1.0.0.msi"
            $MsiFileReturn | FL | Out-String | Write-Host
        
        Return - PS-Custom-Object
            MsiPropertyManufacturer   : Dell
            MsiPropertyProductName    : Dell Dock Firmware Update Package
            MsiPropertyProductVersion : 1.0.2
            MsiPropertyProductCode    : {92E40D32-4403-466D-89A3-DA7D504A3A41}
    #>
    $FktName = "Get-MsiProperties() -"

    If ( !(Test-Path ($msiFile) )) {
        Write-Host "`r`n$FktName '$msiFile' does not exist!`n`r$FktName Returning  PSObject MsiDetails"
        $MsiDetails = New-Object -TypeName PSObject
        $MsiDetails | Add-Member -Name 'MsiFilePath'               -MemberType Noteproperty -Value "missing file at '$msiFile'"
        $MsiDetails | Add-Member -Name 'MsiPropertyManufacturer'   -MemberType Noteproperty -Value "missing file"
        $MsiDetails | Add-Member -Name 'MsiPropertyProductName'    -MemberType Noteproperty -Value "missing file"
        $MsiDetails | Add-Member -Name 'MsiPropertyProductVersion' -MemberType Noteproperty -Value "missing file"
        $MsiDetails | Add-Member -Name 'MsiPropertyProductCode '   -MemberType Noteproperty -Value "missing file"
        $MsiDetails | FL | Out-String | Write-Host
        Write-Warning "File does not exist : '$msiFile' !" 
        Return $MsiDetails 
        }
    if ($MyDebug -eq $True) { write-host  "`r`n$FktName '$msiFile' detected in filesystem" }
    
    # ########################################################
    # ########################################################
    
    $Installer = New-Object -ComObject WindowsInstaller.Installer
    try {
        write-host  "`r`n$FktName Opening database '$msiFile' - ReadOnly"
        $msiOpenDatabaseModeReadOnly = 0
       # $ErrorActionPreference = "Stop"
        $Database = Invoke-Method $Installer OpenDatabase @($MsiFilePath , $msiOpenDatabaseModeReadOnly) -erroraction Stop
        }
    catch {
        $Error[0].Exception.GetType().FullName 
        Write-Error "$FktName script cannot open existing MSI - maybe MSI is currently in use " -Category SecurityError
        write-host  "$FktName $_.ScriptStackTrace"
        Return 5 
        }

    # ########################################################
    #
    #   check if table   'Property' exists in MSI - more a kind of MSI.DLL test  -  the used APIs have to get knowledge about the table properties ( rows, columns, headers, dataformat, ..)
    #
    $TableName = "Property"
    $MsiQuery1 = "SELECT Name FROM _Tables WHERE Name='$TableName'"
    if ($MyDebug -eq $True) { Write-host "`r`n$FktName ------    check if table '$TableName' exists in MSI   ---- '$MsiQuery1'"}
    if ($MyDebug -eq $True) { Write-host "$FktName Table - '_Tables' - OPENVIEW" }
    $ViewTableProperty = Invoke-Method $Database OpenView @($MsiQuery1)
    if ($ViewTableProperty) { if ($MyDebug -eq $True) {  Write-host "$FktName OPENVIEW succeeded"  } }
    
    if ($MyDebug -eq $True) { Write-host "$FktName Table - '_Tables' - Execute" }
    Invoke-Method $ViewTableProperty Execute
    if ($MyDebug -eq $True) { Write-host "$FktName Execute done"  }

    if ($MyDebug -eq $True) { Write-host "$FktName Table - '_Tables' - FETCH" }
    $TableProperty = Invoke-Method $ViewTableProperty Fetch
    
    if ($MyDebug -eq $True) { if ($TableProperty) { Write-host "$FktName FETCH succeeded"  } }
    if ($MyDebug -eq $True) { write-host  "$FktName Table - '_Tables' - Close-VIEW"  }
    Invoke-Method $ViewTableProperty Close @()
    if ($MyDebug -eq $True) {  Write-host "$FktName Close-VIEW done" } 

    
    # ########################################################
    #
    #   query table 'Property' for ProductName, ProductVersion, ProductCode and Manufacturer
    #
    $TableName = "Property"
    Write-host "`r`n$FktName  ------    query table '$TableName' for ProuctName and ProductVersion and ProductCode   ------    "

    $ProductName    = ""
    $ProductName    = ""
    $ProductCode    = ""

    if ($TableProperty) {
        # https://learn.microsoft.com/en-us/windows/win32/msi/record-readstream
        $msiReadStreamAnsi    = 2
        $MsiQuery2 = "SELECT Property,Value FROM $TableName"
        if ($MyDebug -eq $True) {  write-host  "$FktName Table  - '$TableName' - '$MsiQuery2' - OPENVIEW" }
        $ViewProperty = Invoke-Method $Database OpenView @($MsiQuery2)
        if ($MyDebug -eq $True) {  if ($ViewProperty) {  Write-host "$FktName OPENVIEW succeeded"  } }

        if ($MyDebug -eq $True) {  write-host  "$FktName Table  - '$TableName' - Execute" }
        Invoke-Method $ViewProperty Execute
        if ($MyDebug -eq $True) {  Write-host "$FktName Execute done"  } 

        if ($MyDebug -eq $True) {  Write-host "$FktName Do-While (Property) FETCH start"  }
        Do {
            $Property = Invoke-Method $ViewProperty Fetch
            # if ($Property) {  Write-Host "$FktName Table - '$TableName' - FETCH table '$TableName' succeeded"  }

            if ($Property) {
                $PropName = Get-Property $Property StringData 1
                If ($PropName -eq "Manufacturer" )   {  $Manufacturer   = Get-Property $Property StringData 2 ; Write-Host "--- Manufacturer   = $Manufacturer "}
                If ($PropName -eq "ProductName" )    {  $ProductName    = Get-Property $Property StringData 2 ; Write-Host "--- ProductName    = $ProductName "}
                If ($PropName -eq "ProductCode" )    {  $ProductCode    = Get-Property $Property StringData 2 ; Write-Host "--- ProductCode    = $ProductCode  "}
                If ($PropName -eq "ProductVersion" ) {  $ProductVersion = Get-Property $Property StringData 2 ; Write-Host "--- ProductVersion = $ProductVersion "}
                }
            }
        While ($Property)

    if ($MyDebug -eq $True) {  write-host  "$FktName Table - '$TableName' - Close-VIEW" }
    Invoke-Method $ViewProperty Close @()
    if ($MyDebug -eq $True) {  Write-host "$FktName Close-VIEW done"  }

    Remove-Variable -Name Property,ViewProperty
    }  
    Remove-Variable -Name TableProperty, ViewTableProperty
    Remove-Variable -Name Database, Installer

   


    $MsiDetails = New-Object -TypeName PSObject
    $MsiDetails | Add-Member -Name 'MsiFilePath'               -MemberType Noteproperty -Value $msiFile
    $MsiDetails | Add-Member -Name 'MsiPropertyManufacturer'   -MemberType Noteproperty -Value $Manufacturer
    $MsiDetails | Add-Member -Name 'MsiPropertyProductName'    -MemberType Noteproperty -Value $ProductName
    $MsiDetails | Add-Member -Name 'MsiPropertyProductVersion' -MemberType Noteproperty -Value $ProductVersion
    $MsiDetails | Add-Member -Name 'MsiPropertyProductCode '   -MemberType Noteproperty -Value $ProductCode
    
    Write-Host "`r`n$FktName Returning  PSObject MsiDetails "
    Return $MsiDetails
    }
    
 
return Main
    
