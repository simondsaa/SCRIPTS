#############################################################
#						     #
#	Listing Certificates in Cert Store located in      #
#	remote machine			              #
#						     #
#############################################################

[int] $CERT_STORE_PROV_SYSTEM = 10
[int] $CERT_SYSTEM_STORE_LOCAL_MACHINE = 0x20000

$certs = @()
$computer = xlwuw-b0mld5

$signature = @'
[DllImport("CRYPT32.DLL", EntryPoint="CertEnumCertificatesInStore", CharSet=CharSet.Auto, SetLastError=true)]
public static extern IntPtr CertEnumCertificatesInStore( 
	IntPtr storeProvider, 
	IntPtr prevCertContext);
	
[DllImport("CRYPT32.DLL", EntryPoint="CertOpenStore", CharSet=CharSet.Auto, SetLastError=true)]
public static extern IntPtr CertOpenStoreStringPara( 
	int storeProvider,
	int encodingType,
	IntPtr hcryptProv,
	int flags,
	String pvPara);
	
[DllImport("CRYPT32.DLL", EntryPoint="CertCloseStore", CharSet=CharSet.Auto, SetLastError=true)]
[return : MarshalAs(UnmanagedType.Bool)]
public static extern bool CertCloseStore(
	IntPtr storeProvider, 
	int flags);
'@
$type = Add-Type -MemberDefinition $signature `
		-Name Win32Utils -Namespace CertStore `
		-PassThru

$store = $type::CertOpenStoreStringPara($CERT_STORE_PROV_SYSTEM, 0, 0, $CERT_SYSTEM_STORE_LOCAL_MACHINE, $computer)
$certID = ($type::CertEnumCertificatesInStore($store,0))
While ($certID -ne 0) {
	$certs += ([System.Security.Cryptography.X509Certificates.X509Certificate2]($certID))
	$certID = ($type::CertEnumCertificatesInStore($store,$certID))
}
$type::CertCloseStore($store,$null)

#Just for testing
foreach ($cert in $certs) {
	$cert.Subject
}
