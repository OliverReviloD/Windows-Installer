# https://  stackoverflow  .com/  questions  /  29937568  /  how-can-i-find-the-product-guid-of-an-installed-msi-setup

# Pro-tip: Change the 7 to a 3 to find programs lurking as per-user installs rather than machine wide.

#  no ELEVATION required

# much faster than and 
#   get-wmiobject Win32_Product # .... like      | Sort-Object -Property Name |Format-Table IdentifyingNumber, Name, LocalPackage -AutoSize


<#
    ProductCode                            LocalPackage                     VersionString    ProductName                                                              
    -----------                            ------------                     -------------    -----------                                                              
    {C788B026-20BD-4E96-B698-533F1D6C5013} C:\Windows\Installer\5cd360.msi  7.2.4            64 Bit HP CIO Components Installer                                       
    {13DA9C7C-EBFB-40D0-94A1-55B42883DF21} C:\Windows\Installer\28de847.msi 21.2.1           64 Bit HP CIO Components Installer                                       
    {23170F69-40C1-2702-2409-000001000000} C:\Windows\Installer\3d42ef.msi  24.09.00.0       7-Zip 24.09 (x64 edition)                                                
    {1CD74565-DF61-48AD-AE55-C15C9EF7683B} C:\Windows\Installer\1c01b0.msi  1.38             ARM                                                                      
    {AFC4B68C-BB41-77BC-96ED-18716A6DAF68} C:\Windows\Installer\b29a1b.msi  10.1.25398.1     BCD and Boot                                                             
    {4F9D175F-8BFA-4A70-A374-9599773A8C15} C:\Windows\Installer\94791.msi   25.2.485         BeyondTrust Privilege Management Cloud Adapter (x64)    
    ----
#>

$Installer = New-Object -ComObject WindowsInstaller.Installer
$InstallerProducts = $Installer.ProductsEx("", "", 7)
$InstalledProducts = ForEach($Product in $InstallerProducts)
    {
    $Product.InstallProperty
    [PSCustomObject]@{
        ProductCode = $Product.ProductCode(); 
        LocalPackage = $Product.InstallProperty("LocalPackage"); 
        VersionString = $Product.InstallProperty("VersionString"); 
        ProductName = $Product.InstallProperty("ProductName");
        }
    } 
$InstalledProducts | sort ProductName


