# https://  stackoverflow  .com/  questions  /  29937568  /  how-can-i-find-the-product-guid-of-an-installed-msi-setup

# Pro-tip: Change the 7 to a 3 to find programs lurking as per-user installs rather than machine wide.

#  no ELEVATION required

# much fast than and 
#   get-wmiobject Win32_Product # .... like      | Sort-Object -Property Name |Format-Table IdentifyingNumber, Name, LocalPackage -AutoSize

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
