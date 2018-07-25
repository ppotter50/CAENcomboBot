#load base variables and assemblies
. .\variables\variables.ps1
. .\variables\token.ps1
Add-Type -AssemblyName System.Web

#load help text
$help = Get-Content .\help.txt -Raw
$helpencoded = [System.Web.HttpUtility]::UrlEncode($help)

#generate file with channel variables from API request then remove '-' from variable names
Clear-Content ".\variables\channelvar.ps1",".\variables\channelvar-fixed.ps1"

$list = Invoke-WebRequest -Uri "https://slack.com/api/channels.list?token=$token" -Method "GET"
$listob = $list.Content | ConvertFrom-Json

for ($n=0; $n -lt $listob.channels.id.Length; $n++) {
	Add-Content -Path .\variables\channelvar.ps1 -Value "`$$($listob.channels.name[$n]) = '$($listob.channels.id[$n])'"
}

Get-Content .\variables\channelvar.ps1 | ForEach-Object {$_ -replace '-',''} | Out-File .\variables\channelvar-fixed.ps1

#generate file with user variables from API request then remove '-' from variable names
Clear-Content ".\variables\uservar.ps1",".\variables\uservar-fixed.ps1"

$user = Invoke-WebRequest -Uri "https://slack.com/api/users.list?token=$token" -Method "GET"
$userob = $user.Content | ConvertFrom-Json

for ($n=0; $n -lt $userob.members.id.Length; $n++) {
	Add-Content -Path .\variables\uservar.ps1 -Value "`$$($userob.members.name[$n]) = '$($userob.members.id[$n])'"
}

Get-Content .\variables\uservar.ps1 | ForEach-Object {$_ -replace '-',''} | Out-File .\variables\uservar-fixed.ps1

#load variables from generated files
$paulstesting = 'GBS051B28' #for testing
. .\variables\channelvar-fixed.ps1
. .\variables\uservar-fixed.ps1

#function to reset the loop after a message has been output or the loop has completed
function done {
	$checkDoub = $mesob.messages.text
	Start-Sleep 1
	continue
}

#function to determine whether counting failed or successful loads
function setType {
	if ($mesob.messages.text.Contains("successful") -and -not $mesob.messages.text.Contains("failed")) {
		$type = "successful"
	}
	elseif ($mesob.messages.text.Contains("failed") -and -not $mesob.messages.text.Contains("successful")) {
		$type = "failed"
	}
	else{
		$badType = [System.Web.HttpUtility]::UrlEncode("Please specify if you would like to know how many were successful or how many failed by including either 'successful' or 'failed' in your message`nFor help text say 'help'")

		Invoke-WebRequest -Uri "https://slack.com/api/chat.postMessage?token=$token&channel=$windowslogs&text=$badType" -Method "POST"
		done
	}
	$type
}

#function that actually counts the loads
function howmanyCount {
	if ($mesob.messages.text.Contains("successful")) {
		for ($n=0;$n -lt $histob.messages.attachments.color.Length; $n++) {
			if ($histob.messages.attachments.color[$n] -eq $green) {

				$count++

			}
		}
	}
	elseif ($mesob.messages.text.Contains("failed")) {
		for ($n=0;$n -lt $histob.messages.attachments.color.Length; $n++) {
			if ($histob.messages.attachments.color[$n] -eq $red) {

				$count++

			}
		}
	}

	$count
}

#function that generates the list of loaded computers
function loadList {
	if ($mesob.messages.text.Contains("failed")) {
		$Script:loadtype = " unsuccessfully"
		for ($eye=0;$eye -lt $histob.messages.attachments.color.Length;$eye++){
			if ($histob.messages.attachments.color[$eye] -eq $red) {
				$fullText = $histob.messages.attachments.text[$eye].Split(' ')
				$simpleText = $fullText[0]
				[void]$loads.Add($simpleText)
				if ($mesob.messages.text.Contains("remove doubles")) {
					for ($nose=$eye+1;$nose -lt $histob.messages.attachments.color.Length;$nose++) {
						if ($histob.messages[$nose].attachments[0].text.Contains($simpleText)) {
							$histob.messages[$nose].attachments[0].color = "000000"
						}
					}
				}
			}
		}
	}
	elseif ($mesob.messages.text.Contains("all")) {
		$Script:loadtype = ""
		for ($eye=0;$eye -lt $histob.messages.attachments.color.Length;$eye++){
			if ($histob.messages.attachments.color[$eye] -eq $blue) {
				$fullText = $histob.messages.attachments.text[$eye].Split(' ')
				$simpleText = $fullText[0]
				[void]$loads.Add($simpleText)
				if ($mesob.messages.text.Contains("remove doubles")) {
					for ($nose=$eye+1;$nose -lt $histob.messages.attachments.color.Length;$nose++) {
						if ($histob.messages[$nose].attachments[0].text.Contains($simpleText)) {
							$histob.messages[$nose].attachments[0].color = "000000"
						}
					}
				}
			}
		}
	}
	else {
		$Script:loadtype = " successfully"
		for ($eye=0;$eye -lt $histob.messages.attachments.color.Length;$eye++){
			if ($histob.messages.attachments.color[$eye] -eq $green) {
				$fullText = $histob.messages.attachments.text[$eye].Split(' ')
				$simpleText = $fullText[0]
				[void]$loads.Add($simpleText)
				if ($mesob.messages.text.Contains("remove doubles")) {
					for ($nose=$eye+1;$nose -lt $histob.messages.attachments.color.Length;$nose++) {
						if ($histob.messages[$nose].attachments[0].text.Contains($simpleText)) {
							$histob.messages[$nose].attachments[0].color = "000000"
						}
					}
				}
			}
		}
	}

	if ($mesob.messages.text.Contains("alphabetical")) {
		$loads = $loads | Sort-Object
	}

	$loadsout = $loads -join "\n"
	$loadsout
}


while ($infinite) {

	#get message from windows-logs channel
	$mes = Invoke-WebRequest -Uri "https://slack.com/api/groups.history?token=$token&channel=$windowslogs&count=1&inclusive=true" -Method "GET"
	$mesob = $mes.Content | ConvertFrom-Json

	#if statements check if the message is the previous message, if the message is directed to the bot, and determines what the user wants to do
	if ($checkDoub -ne $mesob.messages.text -and $checkDoub -ne ""){
		if ($mesob.messages.text.Contains("<@$combobot>")){

			#'help' command
			if ($mesob.messages.text.Contains("help")){

				Invoke-WebRequest -Uri "https://slack.com/api/chat.postMessage?token=$token&channel=$windowslogs&text=$helpencoded" -Method "POST"
				done

			}

			#'how many' command
			elseif ($mesob.messages.text.Contains("how many")){

				#initialize counter
				$count = 0

				#'since' sub command of 'how many'
				if ($mesob.messages.text.Contains("since")) {
					#first ensure a date was entered then ensure the proper amount of dates are present
					if ($mesob.messages.text -match "(\d\d|\d)(\/|-)(\d\d|\d)(\/|-)(\d\d\d\d|\d\d)") {

						$SplitMatches = Select-String "(\d\d|\d)(\/|-)(\d\d|\d)(\/|-)(\d\d\d\d|\d\d)" -input $mesob.messages.text -AllMatches | ForEach-Object{$_.matches.value}

						if (@($SplitMatches).Length -ne 1) {

							$wrongQuan = [System.Web.HttpUtility]::UrlEncode("The wrong number of dates has been entered`nFor the 'how many since' command one date is required`nFor help text say 'help'")

							Invoke-WebRequest -Uri "https://slack.com/api/chat.postMessage?token=$token&channel=$windowslogs&text=$wrongQuan" -Method "POST"
							done

						}

						#convert date to unix time
						$date = Get-Date "$(@($SplitMatches)[0]) 00:01" -UFormat %s

						#pull channel history from Slack channel 'windows-logs'
						$hist = Invoke-WebRequest "https://slack.com/api/channels.history?token=$token&channel=$windowslogs&count=1000&oldest=$date&inclusive=true" -Method "GET"
						$histob = $hist.Content | ConvertFrom-Json

						$type = setType

						$count = howmanyCount

						#output result to Slack channel
						$countResult = [System.Web.HttpUtility]::UrlEncode("There have been $count $type loads since $(@($SplitMatches)[0])")
						Invoke-WebRequest -Uri "https://slack.com/api/chat.postMessage?token=$token&channel=$windowslogs&text=$countResult" -Method "POST"
						done

					}
					#response if no date was entered
					elseif ($mesob.messages.text -NotMatch "(\d\d|\d)(\/|-)(\d\d|\d)(\/|-)(\d\d\d\d|\d\d)") {

						$badDate = [System.Web.HttpUtility]::UrlEncode("The date entered did not have the proper formatting or there was no date entered, please enter a date in mm/dd/yyyy or mm-dd-yyyy format`nFor help text say 'help'")

						Invoke-WebRequest -Uri "https://slack.com/api/chat.postMessage?token=$token&channel=$windowslogs&text=$badDate" -Method "POST"
						done

					}
				}

				#'between' sub command of 'how many'
				elseif ($mesob.messages.text.Contains("between")) {
					#first ensure a date was entered then ensure the proper amount of dates are present
					if ($mesob.messages.text -match "(\d\d|\d)(\/|-)(\d\d|\d)(\/|-)(\d\d\d\d|\d\d)") {

						$SplitMatches = Select-String "(\d\d|\d)(\/|-)(\d\d|\d)(\/|-)(\d\d\d\d|\d\d)" -input $mesob.messages.text -AllMatches | ForEach-Object{$_.matches.value}

						if (@($SplitMatches).Length -ne 2) {

							$wrongQuan = [System.Web.HttpUtility]::UrlEncode("The wrong number of dates has been entered`nFor the 'how many between' command two dates are required`nFor help text say 'help'")

							Invoke-WebRequest -Uri "https://slack.com/api/chat.postMessage?token=$token&channel=$windowslogs&text=$wrongQuan" -Method "POST"
							done

						}

						#convert dates to unix time
						$startdate = Get-Date "$(@($SplitMatches)[0]) 00:01" -UFormat %s
						$enddate = Get-Date "$(@($SplitMatches)[1]) 23:59" -UFormat %s

						#pull channel history from Slack channel 'windows-logs'
						$hist = Invoke-WebRequest "https://slack.com/api/channels.history?token=$token&channel=$windowslogs&count=1000&oldest=$startdate&latest=$enddate&inclusive=true" -Method "GET"
						$histob = $hist.Content | ConvertFrom-Json

						$type = setType

						$count = howmanyCount

						#output result to Slack channel
						$countResult = [System.Web.HttpUtility]::UrlEncode("There were $count $type loads between $(@($SplitMatches)[0]) and $(@($SplitMatches)[1])")
						Invoke-WebRequest -Uri "https://slack.com/api/chat.postMessage?token=$token&channel=$windowslogs&text=$countResult" -Method "POST"
						done
					}
					#response if no date was entered
					elseif ($mesob.messages.text -NotMatch "(\d\d|\d)(\/|-)(\d\d|\d)(\/|-)(\d\d\d\d|\d\d)") {

						$badDate = [System.Web.HttpUtility]::UrlEncode("The date entered did not have the proper formatting or there was no date entered, please enter a date in mm/dd/yyyy or mm-dd-yyyy format`nFor help text say 'help'")

						Invoke-WebRequest -Uri "https://slack.com/api/chat.postMessage?token=$token&channel=$windowslogs&text=$badDate" -Method "POST"
						done

					}

				}

				#'on' sub command of 'how many'
				elseif ($mesob.messages.text.Contains("on")) {
					#first ensure a date was entered then ensure the proper amount of dates are present
					if ($mesob.messages.text -match "(\d\d|\d)(\/|-)(\d\d|\d)(\/|-)(\d\d\d\d|\d\d)") {

						$SplitMatches = Select-String "(\d\d|\d)(\/|-)(\d\d|\d)(\/|-)(\d\d\d\d|\d\d)" -input $mesob.messages.text -AllMatches | ForEach-Object{$_.matches.value}

						if (@($SplitMatches).Length -ne 1) {

							$wrongQuan = [System.Web.HttpUtility]::UrlEncode("The wrong number of dates has been entered`nFor the 'how many on' command one date is required`nFor help text say 'help'")

							Invoke-WebRequest -Uri "https://slack.com/api/chat.postMessage?token=$token&channel=$windowslogs&text=$wrongQuan" -Method "POST"
							done

						}

						#convert dates to unix time (need to be start and end of day)
						$startdate = Get-Date "$(@($SplitMatches)[0]) 00:01" -UFormat %s
						$enddate = Get-Date "$(@($SplitMatches)[0]) 23:59" -UFormat %s

						#pull channel history from Slack channel 'windows-logs'
						$hist = Invoke-WebRequest "https://slack.com/api/channels.history?token=$token&channel=$windowslogs&count=1000&oldest=$startdate&latest=$enddate&inclusive=true" -Method "GET"
						$histob = $hist.Content | ConvertFrom-Json

						$type = setType

						$count = howmanyCount

						#output result to Slack channel
						$countResult = [System.Web.HttpUtility]::UrlEncode("There were $count $type loads on $(@($SplitMatches)[0])")
						Invoke-WebRequest -Uri "https://slack.com/api/chat.postMessage?token=$token&channel=$windowslogs&text=$countResult" -Method "POST"
						done
					}
					#response if no date was entered
					elseif ($mesob.messages.text -NotMatch "(\d\d|\d)(\/|-)(\d\d|\d)(\/|-)(\d\d\d\d|\d\d)") {

						$badDate = [System.Web.HttpUtility]::UrlEncode("The date entered did not have the proper formatting or there was no date entered, please enter a date in mm/dd/yyyy or mm-dd-yyyy format`nFor help text say 'help'")

						Invoke-WebRequest -Uri "https://slack.com/api/chat.postMessage?token=$token&channel=$windowslogs&text=$badDate" -Method "POST"
						done

					}

				}

				#default 'how many' behavior
				else {
					#send message if the wrong number of dates was entered
					if ($mesob.messages.text -match "(\d\d|\d)(\/|-)(\d\d|\d)(\/|-)(\d\d\d\d|\d\d)") {
						$hasdateencode = [System.Web.HttpUtility]::UrlEncode("It appears you have included a date but no sub command, if you would like to specify a date please include a sub command.`nFor help text say 'help'")
						Invoke-WebRequest -Uri "https://slack.com/api/chat.postMessage?token=$token&channel=$windowslogs&text=$hasdateencode" -Method "POST"
						done
					}

					else {
						#get date in UNIX time
						$startdate = Get-Date -Hour 00 -Minute 01 -UFormat %s
						$enddate = Get-Date -Hour 23 -Minute 59 -Uformat %s

						#pull history
						$hist = Invoke-WebRequest "https://slack.com/api/channels.history?token=$token&channel=$windowslogs&count=1000&oldest=$enddate&latest=$startdate&inclusive=true" -Method "GET"
						$histob = $hist.Content | ConvertFrom-Json

						$type = setType

						$count = howmanyCount

						#display message
						$countResult = [System.Web.HttpUtility]::UrlEncode("There have been $count $type loads today")
						Invoke-WebRequest -Uri "https://slack.com/api/chat.postMessage?token=$token&channel=$windowslogs&text=$countResult" -Method "POST"
						done
					}
				}


			}

			#'active' command
			elseif ($mesob.messages.text.Contains("active")) {

				#initialize counter and array
				$count = 0
				$active = New-Object System.Collections.ArrayList

				#get the date from two days ago
				$predate = (Get-Date).AddDays(-2)
				$date = Get-Date $predate -UFormat %s

				$hist = Invoke-WebRequest -Uri "https://slack.com/api/channels.history?token=$token&channel=$windowslogs&oldest=$date&count=1000&inclusive=true"
				$histob = $hist.Content | ConvertFrom-Json

				:skipCount for ($n=0; $n -lt $histob.length; $n++) {
					if ($histob.messages.text[$n] -match '\w{8}-\w{4}-\w{4}-\w{4}-\w{12}') {

						$GUID = $histob.messages.text[$n]
						for ($g=$n+1;$g -lt $histob.length; $g++) {
							if ($histob.messages.text[$g] -eq $GUID) {

								$histob.messages.text[$g] = ''
								continue skipCount

							}
						}

						$count++

						$messyname = $histob.messages.attachments.text[$n].Split(' ')
						$name = $messyname[0]
						[void]$active.Add($name)
					}
				}

				$nameString = $active -join '\n'
				$encodedActive = [System.Web.HttpUtility]::UrlEncode("There are currently $count active loads. The hostnames of these loads are as follows")
				Invoke-WebRequest -Uri "https://slack.com/api/chat.postMessage?token=$token&channel=$windowslogs&text=$encodedActive&attachments[{`"color`":`"000000`",`"text`":`"$nameString`"}]"
				done

			}

			#'what was loaded' command
			elseif ($mesob.messages.text.Contains("what was loaded")) {

				#initialize array
				$loads = New-Object System.Collections.ArrayList

				if ($mesob.messages.text.Contains("since")) {
					#first ensure a date was entered then ensure the proper amount of dates are present
					if ($mesob.messages.text -match "(\d\d|\d)(\/|-)(\d\d|\d)(\/|-)(\d\d\d\d|\d\d)") {

						$SplitMatches = Select-String "(\d\d|\d)(\/|-)(\d\d|\d)(\/|-)(\d\d\d\d|\d\d)" -input $mesob.messages.text -AllMatches | ForEach-Object{$_.matches.value}

						#send message if the wrong number of dates was entered
						if (@($SplitMatches).Length -ne 1) {

							$wrongQuan = [System.Web.HttpUtility]::UrlEncode("The wrong number of dates has been entered`nFor the 'how many since' command one date is required`nFor help text say 'help'")

							Invoke-WebRequest -Uri "https://slack.com/api/chat.postMessage?token=$token&channel=$windowslogs&text=$wrongQuan" -Method "POST"
							done

						}

						#convert date to unix time
						$date = Get-Date "$(@($SplitMatches)[0]) 00:01" -UFormat %s

						#pull channel history from Slack channel 'windows-logs'
						$hist = Invoke-WebRequest "https://slack.com/api/channels.history?token=$token&channel=$windowslogs&count=1000&oldest=$date&inclusive=true" -Method "GET"
						$histob = $hist.Content | ConvertFrom-Json

						$loadlist = loadList

						#encode and output message
						$loadencode = [System.Web.HttpUtility]::UrlEncode("The following computers have been$loadtype loaded since $(@($SplitMatches)[0])")
						Invoke-WebRequest -Uri "https://slack.com/api/chat.postMessage?token=$token&channel=$windowslogs&text=$loadencode&attachments=[{`"color`":`"$purple`",`"text`":`"$loadlist`"}]" -Method 'POST'
						done

					}

					#message if no date is entered
					elseif ($mesob.messages.text -NotMatch "(\d\d|\d)(\/|-)(\d\d|\d)(\/|-)(\d\d\d\d|\d\d)") {

						$badDate = [System.Web.HttpUtility]::UrlEncode("The date entered did not have the proper formatting or there was no date entered, please enter a date in mm/dd/yyyy or mm-dd-yyyy format`nFor help text say 'help'")

						Invoke-WebRequest -Uri "https://slack.com/api/chat.postMessage?token=$token&channel=$windowslogs&text=$badDate" -Method "POST"
						done

					}
				}

				elseif ($mesob.messages.text.Contains("between")) {
					#first ensure a date was entered then ensure the proper amount of dates are present
					if ($mesob.messages.text -match "(\d\d|\d)(\/|-)(\d\d|\d)(\/|-)(\d\d\d\d|\d\d)") {

						$SplitMatches = Select-String "(\d\d|\d)(\/|-)(\d\d|\d)(\/|-)(\d\d\d\d|\d\d)" -input $mesob.messages.text -AllMatches | ForEach-Object{$_.matches.value}

						#send message if the wrong number of dates was entered
						if (@($SplitMatches).Length -ne 2) {

							$wrongQuan = [System.Web.HttpUtility]::UrlEncode("The wrong number of dates has been entered`nFor the 'how many between' command two dates are required`nFor help text say 'help'")

							Invoke-WebRequest -Uri "https://slack.com/api/chat.postMessage?token=$token&channel=$windowslogs&text=$wrongQuan" -Method "POST"
							done

						}

						#convert dates to unix time
						$startdate = Get-Date "$(@($SplitMatches)[0]) 00:01" -UFormat %s
						$enddate = Get-Date "$(@($SplitMatches)[1]) 23:59" -UFormat %s

						#pull channel history from Slack channel 'windows-logs'
						$hist = Invoke-WebRequest "https://slack.com/api/channels.history?token=$token&channel=$windowslogs&count=1000&oldest=$startdate&latest=$enddate&inclusive=true" -Method "GET"
						$histob = $hist.Content | ConvertFrom-Json

						$loadlist = loadList

						#encode and output message
						$loadencode = [System.Web.HttpUtility]::UrlEncode("The following computers were$loadtype loaded between $(@($SplitMatches)[0]) and $(@($SplitMatches)[1])")
						Invoke-WebRequest -Uri "https://slack.com/api/chat.postMessage?token=$token&channel=$windowslogs&text=$loadencode&attachments=[{`"color`":`"$purple`",`"text`":`"$loadlist`"}]" -Method 'POST'
						done
					}

					#message if no date is entered
					elseif ($mesob.messages.text -NotMatch "(\d\d|\d)(\/|-)(\d\d|\d)(\/|-)(\d\d\d\d|\d\d)") {

						$badDate = [System.Web.HttpUtility]::UrlEncode("The date entered did not have the proper formatting or there was no date entered, please enter a date in mm/dd/yyyy or mm-dd-yyyy format`nFor help text say 'help'")

						Invoke-WebRequest -Uri "https://slack.com/api/chat.postMessage?token=$token&channel=$windowslogs&text=$badDate" -Method "POST"
						done

					}
				}

				elseif ($mesob.messages.text.Contains("on")) {
					#first ensure a date was entered then ensure the proper amount of dates are present
					if ($mesob.messages.text -match "(\d\d|\d)(\/|-)(\d\d|\d)(\/|-)(\d\d\d\d|\d\d)") {

						$SplitMatches = Select-String "(\d\d|\d)(\/|-)(\d\d|\d)(\/|-)(\d\d\d\d|\d\d)" -input $mesob.messages.text -AllMatches | ForEach-Object{$_.matches.value}

						#send message if the wrong number of dates was entered
						if (@($SplitMatches).Length -ne 1) {

							$wrongQuan = [System.Web.HttpUtility]::UrlEncode("The wrong number of dates has been entered`nFor the 'how many on' command one date is required`nFor help text say 'help'")

							Invoke-WebRequest -Uri "https://slack.com/api/chat.postMessage?token=$token&channel=$windowslogs&text=$wrongQuan" -Method "POST"
							done

						}

						#convert dates to unix time (need to be start and end of day)
						$startdate = Get-Date "$(@($SplitMatches)[0]) 00:01" -UFormat %s
						$enddate = Get-Date "$(@($SplitMatches)[0]) 23:59" -UFormat %s

						#pull channel history from Slack channel 'windows-logs'
						$hist = Invoke-WebRequest "https://slack.com/api/channels.history?token=$token&channel=$windowslogs&count=1000&oldest=$startdate&latest=$enddate&inclusive=true" -Method "GET"
						$histob = $hist.Content | ConvertFrom-Json

						$loadlist = loadList

						#encode and output message
						$loadencode = [System.Web.HttpUtility]::UrlEncode("The following computers were$loadtype loaded on $(@($SplitMatches)[0])")
						Invoke-WebRequest -Uri "https://slack.com/api/chat.postMessage?token=$token&channel=$windowslogs&text=$loadencode&attachments=[{`"color`":`"$purple`",`"text`":`"$loadlist`"}]" -Method 'POST'
						done
					}

					#message if no date is entered
					elseif ($mesob.messages.text -NotMatch "(\d\d|\d)(\/|-)(\d\d|\d)(\/|-)(\d\d\d\d|\d\d)") {

						$badDate = [System.Web.HttpUtility]::UrlEncode("The date entered did not have the proper formatting or there was no date entered, please enter a date in mm/dd/yyyy or mm-dd-yyyy format`nFor help text say 'help'")

						Invoke-WebRequest -Uri "https://slack.com/api/chat.postMessage?token=$token&channel=$windowslogs&text=$badDate" -Method "POST"
						done

					}
				}

				else {
					#send message if the wrong number of dates was entered
					if ($mesob.messages.text -match "(\d\d|\d)(\/|-)(\d\d|\d)(\/|-)(\d\d\d\d|\d\d)") {
						$hasdateencode = [System.Web.HttpUtility]::UrlEncode("It appears you have included a date but no sub command, if you would like to specify a date please include a sub command.`nFor help text say 'help'")
						Invoke-WebRequest -Uri "https://slack.com/api/chat.postMessage?token=$token&channel=$windowslogs&text=$hasdateencode" -Method "POST"
						done
					}
					else {
						#get UNIX dates
						$startdate = Get-Date -Hour 00 -Minute 01 -UFormat %s
						$enddate = Get-Date -Hour 23 -Minute 59 -UFormat %s

						#pull history
						$hist = Invoke-WebRequest "https://slack.com/api/channels.history?token=$token&channel=$windowslogs&count=1000&oldest=$startdate&latest=$enddate&inclusive=true" -Method "GET"
						$histob = $hist.Content | ConvertFrom-Json
						$histob.messages.attachments.color | Out-File "C:\Users\Paul Potter\Downloads\DeleteThis\testing.txt"

						$loadlist = loadList

						#encode and output message
						$loadencode = [System.Web.HttpUtility]::UrlEncode("The following computers were$loadtype loaded today")
						Invoke-WebRequest -Uri "https://slack.com/api/chat.postMessage?token=$token&channel=$windowslogs&text=$loadencode&attachments=[{`"color`":`"$purple`",`"text`":`"$loadlist`"}]" -Method 'POST'
						done
					}
				}
			}

			#no valid command entered
			else {
				$nocommand = [System.Web.HttpUtility]::UrlEncode("No known command was entered.`nFor help text enter the command 'help'`nTo request that the command you tried be added email pauljp@umich.edu and I will do my best to add that functionality")
				Invoke-WebRequest -Uri "https://slack.com/api/chat.postMessage?token=$token&channel=$windowslogs&text=$nocommand" -Method 'POST'
				done
			}
		}
	}

	done

}