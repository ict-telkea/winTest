<#
    Author: jpallavicini@telkea.com
#>

### Déclaration des variables ###
$client = "Core Capital"
$serverListPath = "C:\Scripts\ServerList.txt"
$serverList = Get-Content $serverListPath
$date = Get-Date -UFormat "%Y%m%d-%H%M%S"
$computerName = $env:computername
$global:origin = "$computerName@corecapital.eu"
$global:path = "C:\Scripts\missingUpdates_$date.txt"
$global:smtpServer = "Mail.corecapital.eu"
$ExchangeServer = "10.200.23.2"

### Fonction d'envoie de report par mail ###
function send-mail {

	param(
		[string] $subject,
		[string] $attachments
	)

	Send-MailMessage -subject $subject -BodyAsHtml -body "$bootDate $PatchInfo" -from $origin -to "jpallavicini@telkea.com" -attachments $attachments -SmtpServer $smtpServer

}

### Fonction de création du fichier de report ###
function createReportFile {

    if(!(get-item $path -ErrorAction SilentlyContinue)) {
        New-Item -Path $path -ItemType "File" | Out-Null
    }

}

### Création du fichier de report si il n'existe pas ###
createReportFile


### Création du report et ajout des information dans le fichier créé précédemment ###
foreach($server in $serverList) {
	$pingStatus = Test-NetConnection -ComputerName $server -ErrorAction SilentlyContinue

	if($pingStatus.PingSucceeded -eq $false) {

		Add-Content -Path $path -Value "Error - $server : Unreachable"

	} else {
        import-module pswindowsupdate
		$output = Get-WUList -ComputerName $server
		$Output | Select ComputerName, KB, Title, LastDeploymentChangeTime | Export-Csv $path -NoTypeInformation -Append
        Add-Content -Path $path -Value '"","","",""'
	}

}

### Récupération de la version Exchange

$ExchangeVersion = (get-command "\\$ExchangeServer\c$\Program Files\Microsoft\Exchange Server\V15\Bin\exsetup.exe" | ForEach {$_.FileVersionInfo}).FileVersion

### Génération du contenu HTML pour le corp du mail à envoyer ###
$reportTitle = "<h2><span>$client</span> - Report - $(Get-Date -Format g)</h2>"
$style = "<style>BODY{font-family: Arial; font-size: 16px;}"
$style = $style + ".underlined {text-decoration: underline; color: blue}; span {text-decoration: underline;} h4{font-weight: normal}"
$style = $style + "th { text-align: center; background: #0046c3; color: #fff; max-width: 400px; padding: 5px 10px; }"
$style = $style + "td { font-size: 11px; padding: 5px 20px; color: #000; } tr { background: #b8d1f3; }"
$style = $style + "</style>"
$style = $style + "<h4>Exchange Version installed : <span class='underlined'>$ExchangeVersion</span></h4>"
$style = $style + "<h3>Missing patches per servers</h3>"

$style2 = "<style>BODY{font-family: Arial; font-size: 16px;}"
$style2 = $style2 + ".underlined {text-decoration: underline; color: blue}; span {text-decoration: underline;} h4{font-weight: normal}"
$style2 = $style2 + "th { text-align: center; background: #0046c3; color: #fff; max-width: 400px; padding: 5px 10px; }"
$style2 = $style2 + "td { font-size: 11px; padding: 5px 20px; color: #000; } tr { background: #b8d1f3; }"
$style2 = $style2 + "</style>"
$style2 = $style2 + "$reportTitle"
$style2 = $style2 + "<h3>Last boot up time per servers</h3>"


$PatchInfo = Import-Csv $path | select computername, kb, Title, @{Name = "PublishedDate"; Expression = {$_.LastDeploymentChangeTime}} | ConvertTo-Html -Head $style
$global:message = "$PatchInfo"

$bootDate = Get-Content .\ServerList.txt | %{Get-WmiObject win32_operatingSystem -computerName $PSITEM | select @{Label = "ComputerName"; Expression = {$_.csname}}, @{Label = 'LastBootUpTime'; Expression={$_.ConvertToDateTime($_.lastbootuptime)}}}  | ConvertTo-Html -Head $style2


### Envoi du report par mail
    #$content = (Get-Content $path) -join '<br>'
    #$global:message = "Missing updates for <strong>$client</strong> servers : <br><br> $content"
send-mail -subject "$client - Missing patches status" -attachments $path
