#See main_example() voor de echte code

clear-host

Set-ExecutionPolicy unrestricted -scope currentuser
# SSL uit/trust
add-type @"
    using System.Net;
    using System.Security.Cryptography.X509Certificates;
    public class TrustAllCertsPolicy : ICertificatePolicy {
        public bool CheckValidationResult(
            ServicePoint srvPoint, X509Certificate certificate,
            WebRequest request, int certificateProblem) {
            return true;
        }
    }
"@
[System.Net.ServicePointManager]::CertificatePolicy = New-Object TrustAllCertsPolicy

function main_example() {

    $dllFullFileName="$psscriptroot\FileNet.Api.dll"

    $tmpUser="cpadmin"
    #$password="NoSecrectsInScript"

    $Using = (
     "FileNet.Api.Authentication",
     "FileNet.Api.Core",
     "FileNet.Api.Util",
     "FileNet.Api.Constants",
     "FileNet.Api.Admin"
    )

    try 
    {
        #$asmFileNet=Add-Type -Path $dllFullFileName -ReferencedAssemblies $Using
        $asmFileNet=[System.Reflection.Assembly]::LoadFrom($dllFullFileName)
    } 
      catch [System.Reflection.ReflectionTypeLoadException]
    {
       Write-Host "INFO: Multiple warnings when loading which can be ignored"
       #Write-Host "It seems to load more then needed"
       #Write-Host "Make sure WSE 3.0 is installed from MSI"
       Write-Host "Message: $($_.Exception.Message)"
       Write-Host "StackTrace: $($_.Exception.StackTrace)"
       Write-Host "LoaderExceptions: $($_.Exception.LoaderExceptions)"
    }  

    $cred=get-FNP8Credentials $tmpUser

    $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($cred.password)
    $passWord = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)

    Authenticate-User $cred.UserName $passWord

    #$credentials = new-UsernameCredentials($tmpUser, $password)
    #$credentials = new-object System.Management.Automation.PSCredential($tmpUser, $password)
    
#    $uri="https://localhost:9443/wsi/FNCEWS40MTOM/"
    $uri="https://127.0.0.1:9443/wsi/FNCEWS40MTOM/"
    $uri="https://cpe-cp4ba-starter.apps.itz-v7bmzs.hub04-lb.techzone.ibm.com/wsi/FNCEWS40MTOM/"

    $global:FNP8_Connection=get-Connection $uri
    $global:FNP8_Domain=get-Domain $FNP8_Connection
    $global:FNP8_ObjectStore = get-ObjectStore $FNP8_domain.ObjectStores.getlist()[0].symbolicname
}

#Create a new filenet user credentials object which is used in FNP8 clientcontext

function new-UsernameCredentials($userName, $passWord){
    return new-object FileNet.Api.Authentication.UserNameCredentials($userName, $passWord)
}

function get-Connection($uri){
    $conn = [FileNet.Api.Core.Factory+Connection]::GetConnection($uri)
    return $conn 
}

function Authenticate-User($userName, $passWord){
    $creds=new-UserNameCredentials $userName $passWord
    $result=[FileNet.Api.Util.ClientContext]::SetProcessCredentials($creds)
}

function get-Domain($conn){
    write-host "connection " $conn
    [FileNet.Api.Core.IDomain] $domain = [FileNet.Api.Core.Factory+Domain]::FetchInstance($conn, $null, $null)
    return $domain
}

function get-ObjectStore($os){
    write-host "os " $os
    return $FNP8_Domain.ObjectStores.where({$_.name -eq $os})
}

function get-FNP8Credentials()
{
param($forUser)
    if (test-path $env:homedrive) {
        $homeDir=$env:homedrive + $env:homepath
        If ($homeDir[-1] -notmatch "\\") { $homeDir+="\\" } 

        $tmpUser=$forUser
        if ($tmpUser -eq $null) { $tmpUser=$env:username}
        
        $pwdFNP8File=$homeDir + "FNP8" + $tmpUser + "DefaultPWD.sec"

        write-host "Credentials in " $pwdFNP8File
        write-host "*** Delete that file for resetting password or when you feel its time to clean ***"

        #ask it at least every 2 hours

        if (test-path $pwdFNP8File) {
            if (((gci $pwdFNP8File).LastWriteTime | new-timespan).hours -gt 2) {
                write-host "removing"
                remove-item $pwdFNP8File
            }
        }
  

        if (!(Test-Path -path $pwdFNP8File)) {
            write-host "asking again " $pwdFNP8File
            read-host -prompt ("Password for " + $tmpUser) -assecurestring | convertfrom-securestring | out-file $pwdFNP8File
        }

      
        $password = get-content $pwdFNP8File | convertto-securestring
       
        $credentials = new-object System.Management.Automation.PSCredential($tmpUser, $password)
    }
    return $credentials
}

#run the main function
main_example

#query mbv powershell
$FNP8_Domain.ObjectStores | select name, datecreated, datelastmodified, creator, lastmodifier | format-table

#stop de dos objectstore in een variable
$dos=$FNP8_Domain.ObjectStores.where({$_.name -eq  $FNP8_domain.ObjectStores.getlist()[0].symbolicname})

#DUMP ALLE PROPERTIES DIE JE VIA ACCE KAN ZIEN OP HET SCHERM (KAN EVEN DUREN)
$dos

function get-os( $name) {
    $FNP8_Domain.ObjectStores.where({$_.name -eq $name})
}

function New-DynamicFunctionsForFNP8()
{

    foreach($os in $FNP8_Domain.ObjectStores) {

         # creates a new function dynamically
         $Name = $os.name.tostring()
         $code= ""
         $Code += "write-host 'Switching to objectstore:' $name "
         $Code += "`r`nset-variable fnp8os -value (get-os $name) -scope global | out-null"
   #      $Code += "`r`n`$fnp8os "
         $Code += ""
        # create new function in function: drive and set scope to "script:"
        if (test-path -path function:$name) {
            remove-item function:$name
        }
        $null = New-Item -Path function: -Name "script:$Name" -Value $Code
   }
}
New-DynamicFunctionsForFNP8



#Generate some documentation

#Helper function
function documentSection($summary, $textOutput, $fileName) 
{
    "<details>"                               | out-file $fileName -Encoding utf8 -Append -Width 2048
    $summary                                  | out-file $fileName -Encoding utf8 -Append -Width 2048
    '```'                                     | out-file $fileName -Encoding utf8 -Append -Width 2048
    $textOutput                               | out-file $fileName -Encoding utf8 -Append -Width 2048
    '```'                                     | out-file $fileName -Encoding utf8 -Append -Width 2048
    "</details>"                              | out-file $fileName -Encoding utf8 -Append -Width 2048

}

#Initial documentation for an unknown domain to get quick oversight
function letsDocument() {
    $fileName= "$env:USERPROFILE\OneDrive\Documenten\FNP8Tool2\Domain documentation.md"

    remove-item $fileName

    $textOutput = $FNP8_Domain.ObjectStores | select name, symbolicname,displayname, id,DatabaseTableStorageLocation, DatabaseSchemaName, datecreated, creator, datelastmodified, lastmodifier | ft -AutoSize
    $summary="<summary>All objectstores</summary>`r`n"
    documentSection $summary $textOutput $fileName

    foreach($os in $FNP8_Domain.ObjectStores) {
    "<details>"                               | out-file $fileName -Encoding utf8 -Append
    $summary="<summary>Details for $($os.Displayname) with rootfolder $($os.rootfolder)</summary>`r`n"

    $summary                                  | out-file $fileName -Encoding utf8 -Append
    
        $textOutput=$os.AddOnInstallationRecords | select addonname, installer, installationdate | sort addonname | ft -AutoSize
        $summary="<summary>AddOnInstallationRecords for $($os.Displayname)</summary>`r`n"
        documentSection $summary $textOutput $fileName

        $textOutput=$os.ChoiceLists | select id, name, @{name="AddonName"; expression={$_.installedbyaddon.displayname}} | ft -AutoSize
        $summary="<summary>Choicelists for $($os.Displayname)</summary>`r`n"
        documentSection $summary $textOutput $fileName
                
        $textOutput=$os.ClassDescriptions | select id, symbolicname, name, displayname, descriptiveText | sort symbolicname | ft -AutoSize
        $summary="<summary>Classdescriptions for $($os.Displayname)</summary>`r`n"
        documentSection $summary $textOutput $fileName

        $textOutput=$os.RootClassDefinitions | select id, name, symbolicname, descriptivetext, datecreated, creator, datelastmodified, modifier, issystemowned | sort symbolicname | ft -AutoSize 
        $summary="<summary>RootClassDefinitions for $($os.Displayname)</summary>`r`n"
        documentSection $summary $textOutput $fileName


    "</details>"                              | out-file $fileName -Encoding utf8 -Append

    }
}
function Add-FNHelloWorldDocuments {
    param (
        [int]$Count = 100,
        [string]$FolderPath = "/Temp",
        [string]$DocumentClass = "Document",
        [string]$CustomDateCreatedProperty = "DateCreated",
        [string]$CustomDateModifiedProperty = "DateLastModified"
    )

    if (-not $global:FNP8_ObjectStore) {
        Write-Error "FNP8_ObjectStore is not defined. Call get-ObjectStore first."
        return
    }

    # Optional folder filing
    $folder = $null
    if ($FolderPath) {
        try {
            $folder = [Filenet.Api.Core.Factory+Folder]::FetchInstance($global:FNP8_ObjectStore, $FolderPath, $null)
        } catch {
            Write-Warning "⚠️ Folder not found: $FolderPath. Documents will not be filed."
        }
    }

    for ($i = 1; $i -le $Count; $i++) {
        $title = "HelloWorldDoc_$i"

        # Create document
        $doc = [Filenet.Api.Core.Factory+Document]::CreateInstance($global:FNP8_ObjectStore, $DocumentClass)
        $props = $doc.Properties

        $props.PutValue("DocumentTitle", $title)

        $created = (Get-Date)
        $rndHour=get-random -minimum 0 -Maximum 24
        $rndMinute=get-random -minimum 0 -Maximum 60
        $modified = $created.AddHours($rndHour)
        $modified = $created.AddMinutes($rndMinute)

        # Optional: set custom datetime properties (assuming they are allowed)
        $props.PutValue($CustomDateCreatedProperty, $created)
        $props.PutValue($CustomDateModifiedProperty, $modified)

        # Create content transfer element
        $stream = New-Object System.IO.MemoryStream
        $writer = New-Object System.IO.StreamWriter($stream)
        $writer.Write("hello world")
        $writer.Flush()
        $stream.Position = 0

        $ct = [Filenet.Api.Core.Factory+ContentTransfer]::CreateInstance()
        $ct.SetCaptureSource($stream)
        $ct.RetrievalName = "$title.txt"

# Create the content element list and add content transfer
        #$contentList = [Filenet.Api.Core.Factory+ContentElementList]::new()
       # $contentList.Add($ct)

        # Add content to document
      #  $doc.ContentElements = $contentList
        $doc.MimeType = "image/tiff"
        $doc.MimeType = "text/plain"

        # File to folder (optional)
        if ($folder) {
            $folder.File($doc, [Filenet.Api.Constants.AutoUniqueName]::GENERATE, $null)
        }

        $doc.Save(1)
        #Write-Host "✅ Created: $title"
    }
}


