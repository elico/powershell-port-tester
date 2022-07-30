
function Get-RandomCharacters($length, $characters) {
    $random = 1..$length | ForEach-Object { Get-Random -Maximum $characters.length }
    $private:ofs=""
    return [String]$characters[$random]
}
 
function Scramble-String([string]$inputString){     
    $characterArray = $inputString.ToCharArray()   
    $scrambledStringArray = $characterArray | Get-Random -Count $characterArray.Length     
    $outputString = -join $scrambledStringArray
    return $outputString 
}
 
$password = Get-RandomCharacters -length 6 -characters 'abcdefghiklmnoprstuvwxyz'
$password += Get-RandomCharacters -length 3 -characters 'ABCDEFGHKLMNOPRSTUVWXYZ'
$password += Get-RandomCharacters -length 2 -characters '1234567890'
#$password += Get-RandomCharacters -length 1 -characters '!"Â§$%&/()=?}][{@#*+'

$password = Scramble-String $password

Write-Host $password

$tempFolderPath = Join-Path $Env:Temp $(New-Guid); New-Item -Type Directory -Path $tempFolderPath | Out-Null

$File = $tempFolderPath + "\" + $password + ".msrcIncident"

$ProcName  = "msra.exe"
$Arguments = ("/saveasfile "+ $File + " " + $password)

$ProcessStartInfo           = New-Object System.Diagnostics.ProcessStartInfo $ProcName
$ProcessStartInfo.Arguments = $Arguments
$Process                    = [System.Diagnostics.Process]::Start($ProcessStartInfo)

$Computer = $env:COMPUTERNAME

Write-Host $Computer

$Users = query user /server:$Computer 2>&1

$Users = $Users | ForEach-Object {
    (($_.trim() -replace ">" -replace "(?m)^([A-Za-z0-9]{3,})\s+(\d{1,2}\s+\w+)", '$1  none  $2' -replace "\s{2,}", "," -replace "none", $null))
} | ConvertFrom-Csv

foreach ($User in $Users)
{
    [PSCustomObject]@{
        ComputerName = $Computer
        Username = $User.USERNAME
        SessionState = $User.STATE.Replace("Disc", "Disconnected")
        SessionType = $($User.SESSIONNAME -Replace '#', '' -Replace "[0-9]+", "")
    } 
}

Write-Host $Users
$UsersStr = Out-String -InputObject $Users

while (!(Test-Path $File)) {
    Start-Sleep 1 
}

$ol = New-Object -comObject Outlook.Application

$mail = $ol.CreateItem(0)
$mail.Subject = $Computer + " Invited you for a Remote Assistence Session"
$mail.Body =  "Password`n-----`n" + $password + "`n-----`n" + $UsersStr
$mail.Attachments.Add($File);

$mail.save()

$inspector = $mail.GetInspector
$inspector.Display()

Read-Host -Prompt "Press Enter to continue"
