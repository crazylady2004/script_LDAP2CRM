# Initialiser la variable $arg avec le premier argument passé au script
if ($args.Count -ne 1) {
    Write-Output "You must provide exactly one argument."
    break
}

$arg = $args[0]

if ($arg -cne $arg.ToLower() -or ($arg -ne "test" -and $arg -ne "prod" -and $arg -ne "preprod")) {
    Write-Output "The argument you wrote $arg does not match with the accepted arguments (test,prod,Preprod) "
	break
}

# Définir les variables
$encoding = [System.Text.Encoding]::UTF8
$from = "noreply+" + $env:computername + "@epfl.ch"
$to = "emilie.dorer@epfl.ch"
$Server = "mail.epfl.ch"
$SMTPclient = new-object System.Net.Mail.SmtpClient($Server)
$Subject =  "CRM de $arg - Un problème est survenu lors de la synchnonisation entre LDAP et le CRM " 
$body = @"
La synchronisation entre LDAP et le CRM n'a donc pas pu être effectuée aujourd'hui sur l'environnement de $arg du CRM. <br> 
Pour plus d'informations, veuillez vous connecter à la machine $env:computername et consulter les logs au chemin suivant : C:\Scripts\ELCA.CRM.EPFL.LDAPSynchro\Logs ou les tâches planifiées ainsi que directement dans le script dans C:\Scripts\ELCA.CRM.EPFL.LDAPSynchro\launch_ELCA.CRM.EPFL.LDAPSynchro.ps1
"@
$filepath = "C:\Scripts\ELCA.CRM.EPFL.LDAPSynchro\ELCA.CRM.EPFL.LDAPSynchro_$arg.exe"
# Débloquer le fichier
Unblock-File -Path $filepath

# Exécuter le fichier et rediriger la sortie standard vers un fichier de log
Start-Process -FilePath $filepath -NoNewWindow -Wait -RedirectStandardOutput "log_$arg.txt"

# Lire tout le contenu du fichier de log
$content = Get-Content -Path .\log_$arg.txt
$lineCount = $content.Count


# Définir la fonction send-Email
function send-Email() {
    $Message = New-Object System.Net.Mail.MailMessage($from, $to, $Subject, $body)
    $Message.IsBodyHtml = $true
	$Message.SubjectEncoding = $encoding
    $SMTPclient.Send($Message)
}

# Vérifier la condition
if ($lineCount -gt 4 -and $content[4] -match "Dynamics connection successful with user :") {
    # Ne rien faire si la condition est remplie
} else {
    # Envoyer un email si la condition n'est pas remplie
    send-Email 
}


