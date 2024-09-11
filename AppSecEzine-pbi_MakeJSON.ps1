#requires -version 7.4
<#
.SYNOPSIS
  Makes JSON file from AppSecEzine
.DESCRIPTION
  This script makes a JSON file that is used to build a PowerBI report for AppSecEzine. Original content location: https://github.com/Simpsonpt/AppSecEzine
.INPUTS
  Reads AppSecEzine from .\AppSecEzine
.OUTPUTS
  JSON file stored in .\AppSecEzine-pbi.json
.NOTES
  Version:        1.0
  Author:         vcudachi
  Creation Date:  2024-0909@1900
  License:        Creative Commons Attribution NonCommercial ShareAlike (CC-NC-SA: https://creativecommons.org/licenses/by-nc-sa/3.0/)
  GitHub:         https://github.com/vcudachi/AppSecEzine-pbi
  
.EXAMPLE
  #Powershell:
  git clone https://github.com/vcudachi/AppSecEzine-pbi.git
  cd AppSecEzine-pbi
  git clone 'https://github.com/Simpsonpt/AppSecEzine.git'
  . .\AppSecEzine-pbi_MakeJSON.ps1
#>

#-----------------------------------------------------------[Execution]------------------------------------------------------------

If (!(Test-Path .\AppSecEzine -PathType Container -ErrorAction SilentlyContinue)) {
    Write-Host 'AppSecEzine is not found. Read README.md for details.' -ForegroundColor RED
    Start-Sleep -s 5
}
Else {
    $Ezines = Get-ChildItem -LiteralPath .\AppSecEzine\Ezines -File -ErrorAction SilentlyContinue | Sort-Object { Try { [int]$_.Name.Substring(0, $_.Name.IndexOf('-')) }Catch {}; }
    $EzineObjects = [System.Collections.Generic.List[PSCustomObject]]::New()
    
    #Extract metadata
    $Pattern = '^#+\s+Week:\s+(?<Week>\d+)\s+\|\s+Month:\s+(?<Month>\w+)\s+\|\s+Year:\s+(?<Year>\d+)\s+\|\s+Release\s+Date:\s+(?<ReleaseDate>\S+)\s+\|\s+Edition:\s+#*(?<Edition>\d+)º*\s+#+$'
    ForEach ($Ezine in $Ezines) {
        $EzineContent = Get-Content $Ezine.FullName -Encoding UTF8
        If ($EzineContent[6] -match $Pattern) {
            $Week = [int]$matches.Week
            $Month = $matches.Month
            $Year = [int]$matches.Year
            $ReleaseDate = $matches.ReleaseDate
            $Edition = [int]$matches.Edition
        }
        Else {
            #This script needs to be adjusted
            Write-Error -Message "Ezine content not match pattern. Skipping '$($Ezine.Name)'"
            Continue
        }
        $Section = $null

        #Parsing
        $SmallPattern = '^(?<Prop>\w+):\s+(?<Val>.+)$'
        $EzineContent | ForEach-Object {
            Switch ($_) {
                '''  ╔╦╗┬ ┬┌─┐┌┬┐  ╔═╗┌─┐┌─┐' { $Section = 'MustSee' }
                '''  ╦ ╦┌─┐┌─┐┬┌─' { $Section = 'Hack' }
                '''  ╔═╗┌─┐┌─┐┬ ┬┬─┐┬┌┬┐┬ ┬' { $Section = 'Security' }
                '''  ╔═╗┬ ┬┌┐┌' { $Section = 'Fun' }
            }
            $uri = $null
            If ($Section) {
                If ($_ -eq [String]::Empty) {
                    If ($Flag) {
                        $EzineObjects.Add($EzineObject)
                    }
                    $EzineObject = [PSCustomObject]@{
                        Week        = $Week
                        Month       = $Month
                        Year        = $Year
                        ReleaseDate = $ReleaseDate
                        Edition     = $Edition
                        Section     = $Section
                    }
                    $Flag = 0
                }
                ElseIf ($_ -match $SmallPattern) {
                    If ($matches.Prop -in 'URL', 'Description') {
                        $Prop = $matches.Prop
                        $Val = $matches.Val
                        If ($Prop -eq 'URL') {
                            $Val = $Val -replace '\s.+$', ''  #repair bad URLs
                            Try { $uri = [uri]$Val } Catch { $uri = $null }
                        }
                    }
                    Else {
                        $Prop = 'Other'
                        $Val = $matches.Val
                    }
                    #The following is not good, but it works
                    If ($Prop -in $EzineObject.psobject.Members.Name) {
                        $EzineObject.$Prop = $EzineObject.$Prop, $Val -join "`n"
                        If ($uri) {
                            $EzineObject.WebSite = $EzineObject.WebSite, $uri.Host -join "`n"
                        }
                    }
                    Else {
                        $EzineObject | Add-Member -MemberType NoteProperty -Name $Prop -Value $Val
                        If ($uri) {
                            $EzineObject | Add-Member -MemberType NoteProperty -Name 'WebSite' -Value $uri.Host
                        }
                    }
                    $Flag = 1
                }
            }
        }
    }
}

#Output JSON file
$EzineObjects | ConvertTo-Json | Set-Content -LiteralPath .\AppSecEzine-pbi.json -Force -Encoding UTF8