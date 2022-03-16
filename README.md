<h1 align="center">
  <a href="https://blackseals.net">
    <img src="https://blackseals.net/features/blackseals.png" width=66% alt="BlackSeals">
  </a>
</h1>

> Promoted development by BlackSeals.net Technology.
> Written by Andyt for BlackSeals.net.
> Copyright 2017-2022 by BlackSeals Network.

## Description

**change_wmsvc-certificate.ps1** is a Windows PowerShell script that will allow changing private certificate for WMSVC (Web Management Service) for IIS. If WMSVC isn't installed, the script will install and enable it. After that the script will add the certificate for you.

 
## Quick Start

Download the script and copy **change_wmsvc-certificate.ps1** to a folder which contain also the certificate and server list. 


## Syntax

#Syntax: .\change_wmsvc-certificate.ps1 *Text file with servernames* *path and private certificate* *secret key for private certificate*
`.\change_wmsvc-certificate.ps1 *Text file with server* *path to private certificate* *Secret Key for private certificate*`
* **Text file with server** is a txt-file which contain FQDN or NetBIOS name for IIS web servers.
* **path to private certificate** is the path including file name for a pfx-file, which should be used for all WMSVC services on different IIS web servers. The certificate should be include all NetBIOS names as Subject Alternative Name (SAN).
* **secret key for private certificate** is the secret key to access and install the pfx-file on every IIS web server. Be carefull with the secret key. Some special signs may be not usable.


## Examples

`.\change_wmsvc-certificate.ps1 Liste.txt "C:\privat.pfx" cert_password`
* Using "Liste.txt", which is in the same folder like the script. The private certificate is on the root of the same computer which is running the script.

`.\change_wmsvc-certificate.ps1 C:\Liste.txt "C:\privat.pfx" "c€rt_p@$$w0rd"`
* Using "Liste.txt" and the private certificate is on the root of the same computer which is running the script.
* The password for the private certificate use special signs. Therefore it is between quotation marks.
