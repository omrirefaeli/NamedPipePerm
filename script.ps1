$pipelist_path = "C:\Users\omri.refaeli\Downloads\SysinternalsSuite\pipelist64.exe"
$accesschk_path = "C:\Users\omri.refaeli\Downloads\SysinternalsSuite\accesschk64.exe"
$output_path = "C:\users\public\pipes.csv"
$delimeter = ","

$pipe_cont =  & $pipelist_path /accepteula| select -Skip 7
$pipes = New-Object System.Collections.ArrayList
$pipe_cont[0..9] | %{
$_ -match "(.+?)\s+(.+?)\s+?(.*)\s" | out-null
$single = New-Object System.Object
$single | Add-Member -MemberType NoteProperty -Name "Pipe_Name" -Value $Matches[1].Trim()
$single | Add-Member -MemberType NoteProperty -Name "Instances" -Value $Matches[2].Trim()
# -1 means unlimited instances possible
$single | Add-Member -MemberType NoteProperty -Name "Max_Instances" -Value $Matches[3].Trim()
$temp_acc = $accesschk_path + " \pipe\" + $matches[1].Trim() + " /accepteula"
$perm = & $accesschk_path \pipe\$($matches[1].trim()) /accepteula | select -Skip 6
$res_perm = $null
if ($perm -match "Error")
{
    $res_perm = "There was an error reading permissions, could be that instances are maxed out"
}
else
{
   $res_perm = $($perm -split ",") -join " $delimeter"
}
$single | Add-Member -MemberType NoteProperty -Name "permissions" -Value $res_perm
$pipes.Add($single)| Out-Null
}

$pipes | export-csv -NoTypeInformation $output_path



