param(
    [Parameter(Mandatory=$true)]
    [string]$desktopPath,
    [string]$csvFileName
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing



# Obtention du chemin du bureau de l'utilisateur et specification du dossier et du fichier CSV

$csvDirectory = ""  # Nom du dossier où le fichier CSV sera stocke
$csvPath = Join-Path -Path $desktopPath -ChildPath "$csvDirectory\$csvFileName"



# Verification et creation du repertoire si necessaire
if (-not (Test-Path -Path (Split-Path -Path $csvPath -Parent))) {
    New-Item -ItemType Directory -Path (Split-Path -Path $csvPath -Parent) | Out-Null
}

# Fonction pour recuperer le statut du chariot
function Get-ChariotStatut {
    param([string]$numeroChariot)
    $csv = Import-Csv -Path $csvPath
    $chariot = $csv | Where-Object { $_.NumeroChariot -eq $numeroChariot }
    return $chariot
}

function WriteInstructionsToFile {
    param([string]$numeroChariot, [string]$statut, [string]$details)
    $filePath = Join-Path -Path $desktopPath -ChildPath "instructions.txt"
    $text = @"
Numero de Chariot: $numeroChariot
Statut: $statut
$details
Date: $(Get-Date -Format "dd-MM-yyyy HH:mm:ss")
"@
    $text | Out-File -FilePath $filePath -Encoding UTF8 -Force
}

function WriteChariotDetailsToFile {
    param([string]$numeroChariot, [string]$statut, [string]$details)
    $chariotFilePath = Join-Path -Path $desktopPath -ChildPath "Chariot_$numeroChariot.txt"
    $text = @"
Date: $(Get-Date -Format "dd-MM-yyyy HH:mm:ss")
$details
______________________________________________________________________________________________
"@

    $text | Out-File -FilePath $chariotFilePath -Append -Encoding UTF8
}


# Fonction pour mettre a jour le statut du chariot dans le CSV
function Update-ChariotStatut {
    param([string]$numeroChariot, [string]$newStatut, [string]$duree, [string]$info, [bool]$Retrait )
    $csv = Import-Csv -Path $csvPath
    $updated = $false
    $currentTime = Get-Date -Format "dd-MM-yyyy HH:mm:ss"
    foreach ($item in $csv) {
        if ($item.NumeroChariot -eq $numeroChariot) {
            $item.Statut = $newStatut
            $item.LastUpdated = $currentTime
	    $item.info = $info
	    	

		if ($Retrait -eq $True){ $item.NbreRetrait = [int]$item.NbreRetrait + 1 }
		
		
	    $durationTimeSpan = [TimeSpan]::Parse($duree) + [TimeSpan]::Parse($($item.duree))
	    $item.LautreDuree = "$($durationTimeSpan.Days)j $($durationTimeSpan.Hours)h $($durationTimeSpan.Minutes)m"
	    $item.Duree = $durationTimeSpan

            $updated = $true
	    }
    }
    if ($updated) {
        $csv | Export-Csv -Path $csvPath -NoTypeInformation -Force
        #[System.Windows.Forms.MessageBox]::Show("Le statut du chariot a ete mis a jour a '$newStatut'.", "Mise a jour reussie", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    }
}

function CalculateServiceDuration {
    param([string]$lastUpdated)
    $lastUpdateDate = [datetime]::ParseExact($lastUpdated, "dd-MM-yyyy HH:mm:ss", $null)
    $currentDate = Get-Date
    $duration = $currentDate - $lastUpdateDate
    return $duration
}

# Fonction pour demarrer le processus de mise a jour
function Start-SecondScript {
    param($chariot)

    $form.Size = New-Object System.Drawing.Size(600,540)
    $form.Text = "Action sur le chariot $($chariot.NumeroChariot)"
    $label.Visible = $false
    $textBox.Visible = $false
    $okButton.Visible = $false

    $global:label2 = New-Object System.Windows.Forms.Label
    $global:label2.Location = New-Object System.Drawing.Point(20,40)
    $global:label2.Size = New-Object System.Drawing.Size(560,80)
    $global:label2.Text = "Chariot $($chariot.NumeroChariot) est actuellement $($chariot.Statut). Choisissez une action :"
    $global:label2.Font = New-Object System.Drawing.Font('Arial',16,[System.Drawing.FontStyle]::Bold)
    $form.Controls.Add($global:label2)

    if ($chariot.Statut -eq 'en circulation'){
	$global:checkBox1 = New-Object System.Windows.Forms.CheckBox
	$global:checkBox1.Location = New-Object System.Drawing.Point(10,120)
	$global:checkBox1.Size = New-Object System.Drawing.Size(520,60)
	$global:checkBox1.Font = New-Object System.Drawing.Font('Arial',15)
	$global:checkBox1.Text = 'Retrait car le chariot est defaillant'
	$form.Controls.Add($global:checkBox1)

	
	$global:checkBox2 = New-Object System.Windows.Forms.CheckBox
	$global:checkBox2.Location = New-Object System.Drawing.Point(10,180)
	$global:checkBox2.Size = New-Object System.Drawing.Size(520,60)
	$global:checkBox2.Font = New-Object System.Drawing.Font('Arial',15)
	$global:checkBox2.Text = "Retrait pour tester d'autre chariot"
	$form.Controls.Add($global:checkBox2)

	

	$global:checkBox3 = New-Object System.Windows.Forms.CheckBox
	$global:checkBox3.Location = New-Object System.Drawing.Point(10,240)
	$global:checkBox3.Size = New-Object System.Drawing.Size(520,60)
	$global:checkBox3.Font = New-Object System.Drawing.Font('Arial',15)
	$global:checkBox3.Text = "Retrait suite au test du chariot"
	$form.Controls.Add($global:checkBox3)
	

	$checkBox1.Add_CheckedChanged({
		$global:checkBox2.Checked = $false
		$global:checkBox3.Checked = $false
	})
	
	$checkBox2.Add_CheckedChanged({
		$global:checkBox1.Checked = $false
		$global:checkBox3.Checked = $false
	})
	
	$checkBox3.Add_CheckedChanged({
		$global:checkBox1.Checked = $false
		$global:checkBox2.Checked = $false
	})
    }
    else 
    {
        $global:checkBox4 = New-Object System.Windows.Forms.CheckBox
	$global:checkBox4.Location = New-Object System.Drawing.Point(10,120)
	$global:checkBox4.Size = New-Object System.Drawing.Size(520,60)
	$global:checkBox4.Font = New-Object System.Drawing.Font('Arial',15)
	$global:checkBox4.Text = 'Ajout car trop peu de chariots'
	$form.Controls.Add($global:checkBox4)

	
	$global:checkBox5 = New-Object System.Windows.Forms.CheckBox
	$global:checkBox5.Location = New-Object System.Drawing.Point(10,180)
	$global:checkBox5.Size = New-Object System.Drawing.Size(520,60)
	$global:checkBox5.Font = New-Object System.Drawing.Font('Arial',15)
	$global:checkBox5.Text = "Ajout pour test suite a une maintenance"
	$form.Controls.Add($global:checkBox5)

	

	$global:checkBox6 = New-Object System.Windows.Forms.CheckBox
	$global:checkBox6.Location = New-Object System.Drawing.Point(10,240)
	$global:checkBox6.Size = New-Object System.Drawing.Size(520,60)
	$global:checkBox6.Font = New-Object System.Drawing.Font('Arial',15)
	$global:checkBox6.Text = "Remise en route suite a une periode de test"
	$form.Controls.Add($global:checkBox6)
	

	$checkBox4.Add_CheckedChanged({
		$global:checkBox5.Checked = $false
		$global:checkBox6.Checked = $false
	})
	
	$checkBox5.Add_CheckedChanged({
		$global:checkBox4.Checked = $false
		$global:checkBox6.Checked = $false
	})
	
	$checkBox6.Add_CheckedChanged({
		$global:checkBox4.Checked = $false
		$global:checkBox5.Checked = $false
	})
     }


	$global:checkBox7 = New-Object System.Windows.Forms.CheckBox
	$global:checkBox7.Location = New-Object System.Drawing.Point(10,120)
	$global:checkBox7.Size = New-Object System.Drawing.Size(520,60)
	$global:checkBox7.Font = New-Object System.Drawing.Font('Arial',15)
	$global:checkBox7.Text = "Mercure"
	$global:checkBox7.Visible = $false 
	$form.Controls.Add($global:checkBox7)

	$global:checkBox8 = New-Object System.Windows.Forms.CheckBox
	$global:checkBox8.Location = New-Object System.Drawing.Point(10,180)
	$global:checkBox8.Size = New-Object System.Drawing.Size(520,60)
	$global:checkBox8.Font = New-Object System.Drawing.Font('Arial',15)
	$global:checkBox8.Text = "Roulements"
	$global:checkBox8.Visible = $false 
	$form.Controls.Add($global:checkBox8)

	$global:checkBox9 = New-Object System.Windows.Forms.CheckBox
	$global:checkBox9.Location = New-Object System.Drawing.Point(10,240)
	$global:checkBox9.Size = New-Object System.Drawing.Size(520,60)
	$global:checkBox9.Font = New-Object System.Drawing.Font('Arial',15)
	$global:checkBox9.Text = "Moteurs"
	$global:checkBox9.Visible = $false 
	$form.Controls.Add($global:checkBox9)

	$global:checkBox10 = New-Object System.Windows.Forms.CheckBox
	$global:checkBox10.Location = New-Object System.Drawing.Point(10,300)
	$global:checkBox10.Size = New-Object System.Drawing.Size(520,60)
	$global:checkBox10.Font = New-Object System.Drawing.Font('Arial',15)
	$global:checkBox10.Text = "Pignons"
	$global:checkBox10.Visible = $false 
	$form.Controls.Add($global:checkBox10)

	$global:checkBox11 = New-Object System.Windows.Forms.CheckBox
	$global:checkBox11.Location = New-Object System.Drawing.Point(11,360)
	$global:checkBox11.Size = New-Object System.Drawing.Size(520,60)
	$global:checkBox11.Font = New-Object System.Drawing.Font('Arial',15)
	$global:checkBox11.Text = "Roues"
	$global:checkBox11.Visible = $false 
	$form.Controls.Add($global:checkBox11)


    $global:actionButton = New-Object System.Windows.Forms.Button
    $global:actionButton.Location = New-Object System.Drawing.Point(60,420)
    $global:actionButton.Size = New-Object System.Drawing.Size(200,60)
    $global:actionButton.Text = if ($chariot.Statut -eq 'en circulation'){'Retirer'} else  {'Ajouter'}
    $global:actionButton.Font = New-Object System.Drawing.Font('Arial',20)
    $global:actionButton.BackColor = [System.Drawing.Color]::FromArgb(255, 200, 200)
    $global:actionButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat

    $global:var = $chariot.NumeroChariot
    $global:var2 = $global:actionButton.Text

    $global:siRetrait = $false
    $global:actionButton.Add_Click({
	# Déterminer le nouveau statut basé sur l'action choisie
	$newStatut = if ($global:var2 -eq 'Ajouter') {'en circulation'} else {'arrete'}
	# Initialiser les détails selon le statut

	# Ajouter les raisons de retrait seulement si le chariot est retiré
	if ($newStatut -eq 'arrete') {
	    $global:siRetrait = $True
	    $details = 'Chariot retire pour la raison suivante:' 
	    if ($global:checkBox1.Checked) { $details += "Defaillance technique. " }
	    if ($global:checkBox2.Checked) { $details += "Test d'autres chariots. " }
	    if ($global:checkBox3.Checked) { $details += "Retrait suite au test du chariot. " }

	    # Calculer la durée de service et l'ajouter aux détails
	    $info = $details
	    $chariot = Get-ChariotStatut -numeroChariot $global:var
	    $duration = CalculateServiceDuration -lastUpdated $chariot.LastUpdated
 	    $details += "`nDuree de service: $($duration.Days) jours, $($duration.Hours) heures, $($duration.Minutes) minutes."
	}   
        if ($newStatut -eq 'en circulation') {
	    $details = "Chariot ajoute pour la raison suivante: "
	    if ($global:checkBox4.Checked) { $details += "Manque de chariot. " }
	    if ($global:checkBox5.Checked) { $details += "Test du chariot. " }
	    if ($global:checkBox6.Checked) { $details += "Remise en route suite a une periode de test. " }
	    $info = $details
	    $duration = 0}
	
	if ($global:actionButton.Text -eq 'Confirmer'){
	    $newStatut = 'arrete'
	    $details = "Les travaux suivant ont ete realise :"
	    if ($global:checkBox7.Checked) { $details += "Mercure; " }
	    if ($global:checkBox8.Checked) { $details += "Roulements; " }
	    if ($global:checkBox9.Checked) { $details += "Moteurs; " }
	    if ($global:checkBox10.Checked) { $details += "Pignons; " }
	    if ($global:checkBox11.Checked) { $details += "Roues ; " }
	    $info = $details
	    $duration = 0}

	# Mettre à jour le statut du chariot dans le CSV et écrire les instructions dans le fichier
	Update-ChariotStatut -numeroChariot $global:var -newStatut $newStatut -duree $duration -info $info -Retrait $global:siRetrait
	WriteInstructionsToFile -numeroChariot $global:var -statut $newStatut -details $details
	WriteChariotDetailsToFile -numeroChariot $global:var -statut $newStatut -details $details
	# Fermer le formulaire
	$form.Close()
    })

    $form.Controls.Add($global:actionButton)

    $global:reparerButton = New-Object System.Windows.Forms.Button
    $global:reparerButton.Location = New-Object System.Drawing.Point(200,340)
    $global:reparerButton.Size = New-Object System.Drawing.Size(200,60)
    $global:reparerButton.Text = 'Reparer'
    $global:reparerButton.Font = New-Object System.Drawing.Font('Arial',20)
    $global:reparerButton.BackColor = [System.Drawing.Color]::FromArgb(255, 200, 200)
    $global:reparerButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $global:reparerButton.Add_Click({
	$global:actionButton.Text = 'Confirmer'
	$global:reparerButton.Visible = $false 
	$global:checkBox4.Visible = $false
	$global:checkBox5.Visible = $false
	$global:checkBox6.Visible = $false
	$global:checkBox7.Visible = $true
	$global:checkBox8.Visible = $true
	$global:checkBox9.Visible = $true
	$global:checkBox10.Visible = $true
	$global:checkBox11.Visible = $true
	$global:label2.Text = "Choisissez une ou des reparations :"
    })
    $form.Controls.Add($global:reparerButton)
    if ($chariot.Statut -eq 'en circulation') { $global:reparerButton.Visible = $false }
 
	
    $actionButtonAn = New-Object System.Windows.Forms.Button
    $actionButtonAn.Location = New-Object System.Drawing.Point(320,420)
    $actionButtonAn.Size = New-Object System.Drawing.Size(200,60)
    $actionButtonAn.Text = 'Annuler'
    $actionButtonAn.Font = New-Object System.Drawing.Font('Arial',20)
    $actionButtonAn.BackColor = [System.Drawing.Color]::FromArgb(255, 200, 200)
    $actionButtonAn.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $actionButtonAn.Add_Click({
	$form.Close()
    })
    $form.Controls.Add($actionButtonAn)

}

# Creation du formulaire principal
$form = New-Object System.Windows.Forms.Form
$form.Text = 'Verification du numero de chariot'
$form.Size = New-Object System.Drawing.Size(600,280)
$form.StartPosition = 'CenterScreen'
$form.BackColor = [System.Drawing.Color]::LightGray
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
$form.TopMost = $true

# Création et configuration du Timer
$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 30000  # 30 sec en millisecondes
$timer.Add_Tick({
    $form.Close()  # Ferme le formulaire lorsque le timer expire
})

# Démarrage du Timer lorsque le formulaire est chargé
$form.Add_Load({
    $timer.Start()
})

# Assurez-vous d'arrêter le Timer lorsque le formulaire se ferme pour nettoyer les ressources
$form.Add_FormClosing({
    $timer.Stop()
})


$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(20,40)
$label.Size = New-Object System.Drawing.Size(560,40)
$label.Text = 'Veuillez entrer le numero de chariot :'
$label.Font = New-Object System.Drawing.Font('Arial',20,[System.Drawing.FontStyle]::Bold)
$form.Controls.Add($label)

$textBox = New-Object System.Windows.Forms.TextBox
$textBox.Location = New-Object System.Drawing.Point(20,100)
$textBox.Size = New-Object System.Drawing.Size(520,40)
$textBox.Font = New-Object System.Drawing.Font('Arial',20)
$form.Controls.Add($textBox)

$okButton = New-Object System.Windows.Forms.Button
$okButton.Location = New-Object System.Drawing.Point(200,160)
$okButton.Size = New-Object System.Drawing.Size(200,60)
$okButton.Text = 'Suivant'
$okButton.Font = New-Object System.Drawing.Font('Arial',20)
$okButton.BackColor = [System.Drawing.Color]::FromArgb(255, 200, 200)
$okButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat

$okButton.Add_Click({
    $chariot = Get-ChariotStatut -numeroChariot $textBox.Text
    if ($chariot) {
        Start-SecondScript -chariot $chariot
    } else {
        [System.Windows.Forms.MessageBox]::Show("Numero de chariot non trouve dans la base de donnees.", "Erreur", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})

$form.Controls.Add($okButton)

$form.ShowDialog()