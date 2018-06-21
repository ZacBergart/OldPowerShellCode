[datetime]$startTime = Get-Date -Format s

$Error.Clear()

#region Functions

#region MyProgBarFunctions
Function stringOverlay( [string]$topLayer, [string]$bottomLayer )
{
	$objLineComponents = New-Object System.Object
	$LeadBlank = 0
	$TrailBlank = 0
	For($a = 0; $a -le $topLayer.Length - 1; $a++)
	{
		If($topLayer[$a] -notlike " ")
		{
			$LeadBlank = $a - 1
			Break
		}
	}
	For($a = $topLayer.Length - 1; $a -ge 0; $a--)
	{
		If($topLayer[$a] -notlike " ")
		{
			$TrailBlank = $a + 1
			Break
		}
	}
	If($bottomLayer.Length -gt ($LeadBlank + 1 + $topLayer.Trim().Length))
	{
		$contBottom = $LeadBlank + 1 + $topLayer.Trim().Length
		$objLineComponents | Add-Member -type NoteProperty -name LeadFilled -value "$($bottomLayer.Substring(0,$LeadBlank + 1))"
		$objLineComponents | Add-Member -type NoteProperty -name LeadEmpty -value ""
		$objLineComponents | Add-Member -type NoteProperty -name Text -value "$($topLayer.Substring($LeadBlank+1,$topLayer.Trim().Length))"
		$objLineComponents | Add-Member -type NoteProperty -name Trail -value "$($bottomLayer.Substring($contBottom,$bottomLayer.Length - $contBottom))"
	}
	Else
	{
		If($bottomLayer.Length -lt $LeadBlank)
		{
			$objLineComponents | Add-Member -type NoteProperty -name LeadFilled -value "$($bottomLayer.Substring(0,$bottomLayer.Length))"
			$numTemp = $LeadBlank - $bottomLayer.Length
			$objLineComponents | Add-Member -type NoteProperty -name LeadEmpty -value "$($topLayer.Substring($bottomLayer.Length,$numTemp + 1))"
			$objLineComponents | Add-Member -type NoteProperty -name Text -value "$($topLayer.Substring($LeadBlank+1,$topLayer.Trim().Length))"
			$objLineComponents | Add-Member -type NoteProperty -name Trail -value ""

		}
		Else
		{
			$objLineComponents | Add-Member -type NoteProperty -name LeadFilled -value "$($bottomLayer.Substring(0,$LeadBlank))" #here
			$objLineComponents | Add-Member -type NoteProperty -name LeadEmpty -value ""
			$objLineComponents | Add-Member -type NoteProperty -name Text -value "$($topLayer.Substring($LeadBlank+1,$topLayer.Trim().Length))"
			$contBottom = $LeadBlank + 1 + $topLayer.Trim().Length
			$objLineComponents | Add-Member -type NoteProperty -name Trail -value ""
		}
	}
	
	Return $objLineComponents
}
Function EstTimeComplete()
{
	Param(
	[int]$TotalIterations,
	[int]$CompletedIterations,
	[int]$secondsSpentSoFar
	)
	$strETA = "Calculating"
	$ETA_minute = 0
	$ETA_hour = 0
	$ETA_day = 0
	$minute = 60
	$hour = $minute * 60
	$day = $hour * 24
	$testminute = $minute - 1
	$testhour = $hour - 1
	$testday = $day - 1
	$TimeToComplete = [Math]::Round(($TotalIterations - $CompletedIterations) * ($secondsSpentSoFar / $CompletedIterations),0)
	If($TimeToComplete -gt $testday)
	{
		$ETA_day = [math]::floor($TimeToComplete / $day)
		$TimeToComplete = $TimeToComplete - ($day * $ETA_day)
	}
	If($TimeToComplete -gt $testhour)
	{
		$ETA_hour = [math]::floor($TimeToComplete / $hour)
		$TimeToComplete = $TimeToComplete - ($hour * $ETA_hour)
	}
	If($TimeToComplete -gt $testminute)
	{
		$ETA_minute = [math]::floor($TimeToComplete / $minute)
		$TimeToComplete = $TimeToComplete - ($minute * $ETA_minute)
	}
	$strSecond = ""
	$strMinute = ""
	$strHour = ""
	$strDay = ""
	If($TimeToComplete -ge 10 -and ($ETA_minute -gt 0 -or $ETA_hour -gt 0 -or $ETA_day -gt 0)){$strSecond = ":$($TimeToComplete)"}
	ElseIf($TimeToComplete -ge 0 -and ($ETA_minute -gt 0 -or $ETA_hour -gt 0 -or $ETA_day -gt 0)){$strSecond = ":0$($TimeToComplete)"}
	Else{$strSecond = "$($TimeToComplete) seconds"}
	If($ETA_minute -ge 10 -and ($ETA_hour -gt 0 -or $ETA_day -gt 0)){$strMinute = ":$($ETA_minute)"}
	ElseIf($ETA_minute -ge 0 -and ($ETA_hour -gt 0 -or $ETA_day -gt 0)){$strMinute = ":0$($ETA_minute)"}
	ElseIf($ETA_minute -gt 0){$strMinute = "$($ETA_minute)"}
	If($ETA_hour -ge 10 -and $ETA_day -gt 0){$strHour = ":$($ETA_hour)"}
	ElseIf($ETA_hour -ge 0 -and $ETA_day -gt 0){$strHour = ":0$($ETA_hour)"}
	ElseIf($ETA_hour -gt 0){$strHour = "$($ETA_hour)"}
	If($ETA_day -gt 0){$strDay = "$($ETA_day)"}
	$strETA = "$($strDay)$($strHour)$($strMinute)$($strSecond)"
	Return $strETA
}
Function MyProgBar()
{
	Param(
	[int]$Numerator = 1,
	[int]$Denominator = 100,
	[string]$Message = "Test message line1`nTest message line2",
	[string]$PBColour = "Red"
	)
	$BGColour = "Black"
	$LinesInMsg = $Message.Split("`n")
	$LinesInUse = $LinesInMsg.Count
	$Percentage = ($Numerator / $Denominator)*100
	$Percentage = "[ {0:#.000}% Completed" -f $Percentage
	$WindowWidth = $Host.UI.RawUI.WindowSize.Width - 3
	$WindowPB = [Math]::Round(($WindowWidth * ($Numerator / $Denominator)),0)
	$MiddlePos = [Math]::Round(($WindowWidth / 2),0)
	[string]$blankLine = ""
	For($tmp = 1; $tmp -le $WindowWidth; $tmp++)
	{
		$blankLine = "$($blankLine) "
	}
	[string]$PBar = ""
	For($tmp = 1; $tmp -le $WindowPB; $tmp++)
	{
		$PBar = "$($PBar) "
	}
	$OldCursorPos = $Host.UI.RawUI.CursorPosition
	$PBLocation = $Host.UI.RawUI.CursorPosition
	$PBLocation.X = 0
	If($OldCursorPos.Y -le $Host.UI.RawUI.WindowSize.get_Height()){$PBLocation.Y = 0}
	Else{$PBLocation.Y = $OldCursorPos.Y - $Host.UI.RawUI.WindowSize.get_Height() + 1}
	$Host.UI.RawUI.CursorPosition = $PBLocation
	For($clnLine = 0; $clnLine -le $LinesInUse; $clnLine++)
	{
		Write-Host -BackgroundColor $Host.UI.RawUI.get_BackgroundColor() "$($blankLine)"
	}
	$Host.UI.RawUI.CursorPosition = $PBLocation
	Write-Host "$($Message)"
	$seconds = (Get-Date).Subtract($startTime)
	$ETA = EstTimeComplete -T:$Denominator -C:$Numerator -s:$seconds.TotalSeconds
	$strETA = " `| Time Left: $($ETA) ]"
	$centeredText = [Math]::Round((($WindowWidth - $Percentage.Length - $strETA.Length) / 2),0)
	[string]$leadETA = ""
	For($tmp = 1; $tmp -le $centeredText; $tmp++)
	{
		$leadETA = "$($leadETA) "
	}
	$rtnObject = stringOverlay "$($leadETA)$($Percentage)$($strETA)" "$($PBar)"
	Write-Host -BackgroundColor $PBColour $rtnObject.LeadFilled -NoNewline
	Write-Host $rtnObject.LeadEmpty -NoNewline
	Write-Host -ForegroundColor White -BackgroundColor Black $rtnObject.Text -NoNewline
	Write-Host -BackgroundColor $PBColour $rtnObject.Trail
	$Host.UI.RawUI.CursorPosition = $OldCursorPos
}
#endregion MyProgBarFunctions

Function DisplayProg ( $ActivityText, $CurrentOperation, $CurrentStatus, $numerator, $denominator )
{
	$tmpPercent = ($numerator / $denominator)*100
	$wrkPercent = "{0:#.000}" -f $tmpPercent
	$tmpRTime = (get-date).Subtract($startTime)
	$hour = $tmpRTime.hours
	$minute = $tmpRTime.minutes
	$second = $tmpRTime.seconds
	If ($minute -le 9){$minute = "0$($minute)"}
	If ($second -le 9){$second = "0$($second)"}
	$runTime = "$($hour):$($minute):$($second)"
	$avgSecPerIteration = [Math]::Round(($denominator - $numerator) * ($tmpRTime.TotalSeconds / $numerator),0)
	$secondsLeft = [Math]::Round(($denominator-$numerator)*$avgSecPerIteration,0)
	$CurrentOperation = "$($secondsLeft)"
	Write-Progress -activity "[$(Get-Date -UFormat %T)] <$($ActivityText)>  $($numerator) of $($denominator) ($($wrkPercent)%)  [RunTime: $($runTime)]  [StartTime: $($startTime)]" -CurrentOperation "$($CurrentOperation)" -status "$($CurrentStatus)" -PercentComplete (($numerator / $denominator)  * 100) -SecondsRemaining $secondsLeft
}
#endregion Functions

For($a = 1; $a -le $Host.UI.RawUI.WindowSize.Width; $a++)
{
	Sleep -Milliseconds 100
	MyProgBar -M:"Testing`nOne`nTwo`tThree" -N:$a -D:$Host.UI.RawUI.WindowSize.Width -PBC:"Blue"
}


#For($a = 1; $a -le $Host.UI.RawUI.WindowSize.Width; $a++)
#{
#	Sleep -Milliseconds 1000
#	DisplayProg "ActivityText" "CurrentOperation" "CurrentStatus" $a $Host.UI.RawUI.WindowSize.Width
#}


