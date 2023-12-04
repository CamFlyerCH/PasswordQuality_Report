# 04.12.2023 - Create_PasswordQuality_Report.ps1 by https://github.com/CamFlyerCH

# Init
Import-Module ImportExcel
$ScriptPath = $MyInvocation.MyCommand.Path
$ScriptDir = Split-Path -Parent $ScriptPath
$DaysOld = 365

$DataFiles = Get-ChildItem -Path $ScriptDir -Filter "PWTEST-Result_*.xml"

ForEach($DataFile in $DataFiles){
    $Domain = $DataFile.BaseName.Replace("PWTEST-Result_","")

    Write-Host ("Start working on domain $Domain")

    $QualityData = Import-Clixml -LiteralPath $DataFile.FullName
    $AccountData = Import-Clixml -Path ($ScriptDir + "\PWTEST-Accounts_" + $Domain + ".xml")

    Write-Host ("Imported quality data and " + $AccountData.Count + " accounts")

    # Create Excel
    $ExcelFilePath = $ScriptDir + "\PWTEST-Report  " + $Domain + "  " + $(Get-Date $DataFile.LastWriteTime -Format "yyyy.MM.dd HH.mm") + "_.xlsx"
    If($ExcelFilePath){
        If (Test-Path $ExcelFilePath) { Remove-Item $ExcelFilePath -Force }
        $ExcelPack = Open-ExcelPackage -Path $ExcelFilePath -Create
        Write-Host("Excel package defined : $ExcelFilePath")
    }

    ForEach($Property in $QualityData.PSObject.Properties.Name) {

        # Special for accounts with the same password hash (else case below)
        If($Property -ne "DuplicatePasswordGroups"){
            Write-Host ("Working on " + $QualityData.$Property.Count + " accounts with $Property")
            $AccountList = @()
            ForEach($User in $QualityData.$Property){
                $ShortUser = $User.Split("\")[1]
                ForEach($Account in $AccountData){
                    If($Account.sAMAccountName -eq $ShortUser){
                        $AccountList += $Account
                        Break
                    } # End if account found in account data
                } # End for each account in all account data
            } # End for each account of this quality type

            # Output to additional Excel sheet
            $LengthStyle = {
                param(
                    $workSheet,
                    $totalRows,
                    $lastColumn
                )    

                2..$totalRows | ForEach-Object{

                    # Mark old dates
                    $Type = $WorkSheet.Cells["M$($_)"]
                    $PWNeverExp = $WorkSheet.Cells["F$($_)"]
                    $Created = $WorkSheet.Cells["B$($_)"]

                    # Coloring for LastLogon
                    $Cell = $WorkSheet.Cells["D$($_)"]
                    If(($Cell.Value -AND $Type.Value -eq "Computer") -or ($Cell.Value -AND $PWNeverExp.Value -eq "TRUE")){
                        $TimeDiff = New-TimeSpan -Start $Cell.Value -End $DataFile.LastWriteTime
                        If ($TimeDiff.TotalDays -gt $DaysOld){
                            $Cell.Style.Fill.PatternType='Solid'
                            $Cell.Style.Fill.BackgroundColor.SetColor('LightPink')
                        }
                    }

                    If([string]::IsNullOrEmpty($Cell.Value) -AND $PWNeverExp.Value -eq "TRUE"){
                        $TimeDiff = New-TimeSpan -Start $Created.Value -End $DataFile.LastWriteTime
                        If ($TimeDiff.TotalDays -gt $DaysOld){
                            $Cell.Style.Fill.PatternType='Solid'
                            $Cell.Style.Fill.BackgroundColor.SetColor('LightGoldenrodYellow')
                        }
                    }

                    # Coloring PwLastSet
                    $Cell=$WorkSheet.Cells["E$($_)"]
                    If($Cell.Value -AND $PWNeverExp.Value -ne "TRUE"){
                        $TimeDiff = New-TimeSpan -Start $Cell.Value -End $DataFile.LastWriteTime
                        If ($TimeDiff.TotalDays -gt $DaysOld){
                            $Cell.Style.Fill.PatternType='Solid'
                            $Cell.Style.Fill.BackgroundColor.SetColor('LightGoldenrodYellow')
                        }
                    }

                    If($Cell.Value){
                        $TimeDiff = New-TimeSpan -Start $Cell.Value -End $DataFile.LastWriteTime
                        If ($TimeDiff.TotalDays -gt ($DaysOld * 10)){
                            $Cell.Style.Fill.PatternType='Solid'
                            $Cell.Style.Fill.BackgroundColor.SetColor('LightPink')
                        }
                    }
                }
            }

            If($Script:ExcelPack -AND $AccountList.Count -gt 0){
                # Add details to Excel
                $Script:ExcelPack = $AccountList | Export-Excel -ExcelPackage $Script:ExcelPack -PassThru -WorksheetName $Property -BoldTopRow -AutoFilter -FreezeTopRowFirstColumn -AutoSize -CellStyleSB $LengthStyle
                $Script:ExcelPack.$Property.View.ZoomScale = 80
                $Columns = 2..5
                ForEach($Column in $Columns){
                    Set-ExcelColumn -ExcelPackage $Script:ExcelPack -Worksheetname $Property -Column $Column -VerticalAlignment Top -HorizontalAlignment Left -Width ([Int]($Script:ExcelPack.$Property.Column($Column).Width * 1.3))
                }
                $Columns = 7..($Script:ExcelPack.$Property.Dimension.Columns)
                ForEach($Column in $Columns){
                    Set-ExcelColumn -ExcelPackage $Script:ExcelPack -Worksheetname $Property -Column $Column -VerticalAlignment Top -HorizontalAlignment Left -Width ([Int]($Script:ExcelPack.$Property.Column($Column).Width * 1.1))
                }
            }


        } Else { 
            $PasswordGroupsUsers = [ordered] @{}
            $Count = 0
            ForEach($Group in $QualityData.$Property) {
                $Count++
                Write-Host ("Working on " + $QualityData.$Property.Count + " groups with " + $Group.Count + " accounts in group $Count with $Property")
                ForEach($User in $Group) { 
                    #Write-Host $User
                    $PasswordGroupsUsers[$User] = "Group $Count"
                    $ShortUser = $User.Split("\")[1]
                    ForEach($Account in $AccountData){
                        If($Account.sAMAccountName -eq $ShortUser){
                            $AccountList += $Account | Select-Object @{n="Group";e={$Count}},*
                            Break
                        } # End if account found in account data
                    } # End for each account in all account data
                } # End for each account in this group of DuplicatePasswordGroups
            } # End for each group of DuplicatePasswordGroups

            # Output to additional Excel sheet
            $LengthStyle = {
                param(
                    $workSheet,
                    $totalRows,
                    $lastColumn
                )    

                2..$totalRows | ForEach-Object{

                    # Mark old dates
                    $Type = $WorkSheet.Cells["N$($_)"]
                    $PWNeverExp = $WorkSheet.Cells["G$($_)"]
                    $Created = $WorkSheet.Cells["C$($_)"]

                    # Coloring for LastLogon
                    $Cell = $WorkSheet.Cells["E$($_)"]
                    If(($Cell.Value -AND $Type.Value -eq "Computer") -or ($Cell.Value -AND $PWNeverExp.Value -eq "TRUE")){
                        $TimeDiff = New-TimeSpan -Start $Cell.Value -End $DataFile.LastWriteTime
                        If ($TimeDiff.TotalDays -gt $DaysOld){
                            $Cell.Style.Fill.PatternType='Solid'
                            $Cell.Style.Fill.BackgroundColor.SetColor('LightPink')
                        }
                    }

                    If([string]::IsNullOrEmpty($Cell.Value) -AND $PWNeverExp.Value -eq "TRUE"){
                        $TimeDiff = New-TimeSpan -Start $Created.Value -End $DataFile.LastWriteTime
                        If ($TimeDiff.TotalDays -gt $DaysOld){
                            $Cell.Style.Fill.PatternType='Solid'
                            $Cell.Style.Fill.BackgroundColor.SetColor('LightGoldenrodYellow')
                        }
                    }

                    # Coloring PwLastSet
                    $Cell=$WorkSheet.Cells["F$($_)"]
                    If($Cell.Value -AND $PWNeverExp.Value -ne "TRUE"){
                        $TimeDiff = New-TimeSpan -Start $Cell.Value -End $DataFile.LastWriteTime
                        If ($TimeDiff.TotalDays -gt $DaysOld){
                            $Cell.Style.Fill.PatternType='Solid'
                            $Cell.Style.Fill.BackgroundColor.SetColor('LightGoldenrodYellow')
                        }
                    }

                    If($Cell.Value){
                        $TimeDiff = New-TimeSpan -Start $Cell.Value -End $DataFile.LastWriteTime
                        If ($TimeDiff.TotalDays -gt ($DaysOld * 10)){
                            $Cell.Style.Fill.PatternType='Solid'
                            $Cell.Style.Fill.BackgroundColor.SetColor('LightPink')
                        }
                    }
                }
            }

            If($Script:ExcelPack -AND $AccountList.Count -gt 0){
                # Add details to Excel
                $Script:ExcelPack = $AccountList | Export-Excel -ExcelPackage $Script:ExcelPack -PassThru -WorksheetName $Property -BoldTopRow -AutoFilter -FreezePane 2,3 -AutoSize -CellStyleSB $LengthStyle
                $Script:ExcelPack.$Property.View.ZoomScale = 80
                $Columns = 3..6
                ForEach($Column in $Columns){
                    Set-ExcelColumn -ExcelPackage $Script:ExcelPack -Worksheetname $Property -Column $Column -VerticalAlignment Top -HorizontalAlignment Left -Width ([Int]($Script:ExcelPack.$Property.Column($Column).Width * 1.3))
                }
                $Columns = 8..($Script:ExcelPack.$Property.Dimension.Columns)
                ForEach($Column in $Columns){
                    Set-ExcelColumn -ExcelPackage $Script:ExcelPack -Worksheetname $Property -Column $Column -VerticalAlignment Top -HorizontalAlignment Left -Width ([Int]($Script:ExcelPack.$Property.Column($Column).Width * 1.1))
                }
            }

        } # End if DuplicatePasswordGroups

    } # End for each quality type

    # Close Excel
    Close-ExcelPackage -ExcelPackage $ExcelPack -Show
    Write-Host("Excel file saved : $ExcelFilePath")

} # End for each domain file

