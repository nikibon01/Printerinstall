CLS

Write-Host "Stelle sicher, dass das WLAN BFSTG ist." -ForegroundColor White -BackgroundColor Red
pause

Set-ExecutionPolicy Unrestricted -Force -ErrorAction SilentlyContinue





#Treiberlinks für den Download
$TreiberFKOF2E05 = "https://support.ricoh.com/bb/pub_e/dr_ut_e/0001311/0001311160/V1700/z94580L16.exe"

$TreiberKOF2E05 = "https://support.ricoh.com/bb/pub_e/dr_ut_e/0001318/0001318367/V1300/z91790L16.exe"

$TreiberFKOF1022A = "https://support.ricoh.com/bb/pub_e/dr_ut_e/0001332/0001332746/V11500/z97066L16.exe"
                                         
$TreiberKOF1022A = "https://support.ricoh.com/bb/pub_e/dr_ut_e/0001318/0001318367/V1300/z91790L16.exe"

$TreiberFKOFC101 = "https://support.ricoh.com/bb/pub_e/dr_ut_e/0001330/0001330637/V1500/z96722L16.exe"

$TreiberFKOFC201 = "https://support.ricoh.com/bb/pub_e/dr_ut_e/0001311/0001311159/V1600/z91862L16.exe"
#_________________________________________________________________________________________________________


$infFileNamePath = "disk1\oemsetup.inf"
$getprinter = Get-Printer

function cleandoubleprinters{

#Kopien die der Mobility Print erstellt werden mit diesen beiden Befehlen gelöscht.
$delprinter = Get-Printer | Where-Object {$_.Name -match "kopie"}
Remove-Printer -InputObject $delprinter -ErrorAction SilentlyContinue
cls

}

Function ClearTempFiles{

Remove-Item "C:\t3mp" -Recurse -Force

}

Function PrintInstall{

New-Item -Path C:\t3mp -ItemType Directory -force
$label.Text = "Created Folder C:\t3mp"


if($getprinter.name -eq "$druckername"){

$label.Text = "$druckername ist bereits installiert oder der Name ist vergeben"
$progressBar.Value = 100
}

else
{




if ($getprinter.name -eq $OldPrinter){

function installingdriver{
# Treiber werden installiert
$label.Text = "Treiber wird heruntergeladen"
Invoke-WebRequest -Uri $LinkTreiber -OutFile "C:\t3mp\$druckername.zip"
$label.Text = "Treiber wurde heruntergeladen"
$progressBar.Value = 15
# Temporärer Ordner für die Druckertreiber
New-Item -Path "C:\t3mp" -Name "$druckername" -ItemType Directory -Force
$label.Text = "Temporäre Ordner sind erstellt worden"
$progressBar.Value = 30
#Druckertreiber werden auf den Temp Ordner hinzugefügt
Expand-Archive "C:\t3mp\$druckername.zip" -DestinationPath "C:\t3mp\$druckername" -Force
$label.Text = "ZIP wurde extrahiert"
$progressBar.Value = 40
# Hinzufügen des Treiber im Treiberpool
pnputil.exe -i -a "C:\t3mp\$druckername\$infFileNamePath"
$label.Text = "Treiber ist installiert"
$progressBar.Value = 50
# Hinzufügen vom Treiber
Add-PrinterDriver $Treibername
$label.Text = "Drucker wurde hinzugefügt"
$progressBar.Value = 65
}


Function AddingNewPrinter{
    
# Hier hole ich den Portname von dem Drucker
$portn = Get-Printer "*$druckername *"
$portname = $portn.PortName
$progressBar.Value = 80

# Hier füge ich den neuen Drucker mit dem richtigem Treiber hinzu.
Add-Printer -Name $druckername -DriverName $Treibername -PortName $portname
$label.Text = "Drucker wurde konfiguriert."
$progressBar.Value = 95
}






Function RemovingOldPrinter{

do{
#Hier Frage ich den User ob er den alten Drucker löschen will. j oder n
$read = "j"


switch($read)
    {


    #Falls der User Ja nimmt löscht er den nicht gebrauchten Drucker. Ansonsten macht er nichts.
    j {
            $delprinter = Get-Printer | Where-Object {$_.Name -eq $OldPrinter}
            Remove-Printer -InputObject $delprinter
            
      }

    n {exit}
    default {cls}


}

}until($read -eq "j")

}

installingdriver
AddingNewPrinter
RemovingOldPrinter
}

else{
Write-Host "Installiere den Basic-Treiber für den Drucker $druckername im Portal"
Write-Host "Mobility Print wird geöffnet!"
sleep -Seconds 1
Start-Process .\MobilityPrint.exe
}

}
}

Function Forms{

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Import the System.Windows.Forms assembly
$comboBox = New-Object System.Windows.Forms.ComboBox
$comboBox.Width = 300
$comboBox.Height = 250
$comboBox.Location = New-Object System.Drawing.Point(10,20)
$comboBox.Font = New-Object System.Drawing.Font("Arial", 14)
Add-Type -AssemblyName System.Windows.Forms

# Create a new form
$form = New-Object System.Windows.Forms.Form
$form.text = "Erweiterte Treiberinstallation"
$form.AutoSize = $True
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $False

#Set the form Icon
$iconpath = Resolve-Path -Path ".\printer2-add.ico"
$form.Icon = New-Object system.drawing.icon ($iconpath)

# Create a new ComboBox and add it to the form

$form.Controls.Add($comboBox)

# Add the options to the ComboBox
$comboBox.Items.Add("FKOF2E05")
$comboBox.Items.Add("KOF2E05")
$comboBox.Items.Add("FKOF1022A")
$comboBox.Items.Add("KOF1022A")
$comboBox.Items.Add("FKOFC101")
$comboBox.Items.Add("FKOFC201")

# Create a button and add it to the form
$button = New-Object System.Windows.Forms.Button
$button.Text = "Install"
$button.Width = 100
$button.Height = 50
$button.Location = New-Object System.Drawing.Point(350,10)
$button.Font = New-Object System.Drawing.Font("Arial", 11)
$form.Controls.Add($button)

$button2 = New-Object System.Windows.Forms.Button
$button2.Location = New-Object System.Drawing.Size(350,70)
$button2.Size = New-Object System.Drawing.Size(75,30)
$button2.Text = "Cleanup"
$button2.Width = 100
$button2.Height = 50
$button2.Font = New-Object System.Drawing.Font("Arial", 11)
$form.Controls.Add($button2)

$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Size(10,50)
$progressBar.Size = New-Object System.Drawing.Size(300,20)
$form.Controls.Add($progressBar)

$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Size(10,150)
$label.Size = New-Object System.Drawing.Size(340,100)
$label.Text = "Starte die Installation vom Drucker"
$form.Controls.Add($label)



# Add an event handler for the button click event
$button.Add_Click({
    switch($comboBox.SelectedItem)
    {
        "FKOF2E05"
        {
            cls
            $progressBar.Value = 0
            $getprinter = Get-Printer
            $Treibername = "RICOH MP C6004 PCL 6"
            $OldPrinter = "FKOF2E05 [Frauenfeld  Bau 2  E05](Mobility)"
            $druckername = "FKOF2E05"
            $LinkTreiber = $TreiberFKOF2E05
            PrintInstall
	        $label.Text = "Drucker $druckername wurde erfolgreich installiert"
            $progressBar.Value = 100
        }
        "KOF2E05"
        {
            cls
            $progressBar.Value = 0
            $getprinter = Get-Printer
            $Treibername = "RICOH MP 7503 PCL 6"
            $Finisher = "Finisher SR4130"
            $OldPrinter = "KOF2E05 [Frauenfeld  Bau 2  E05](Mobility)"
            $druckername = "KOF2E05"
            $LinkTreiber = $TreiberKOF2E05
            PrintInstall
	        $label.Text = "Drucker $druckername wurde erfolgreich installiert"
            $progressBar.Value = 100
        }
        "FKOF1022A"
        {
            cls
            $progressBar.Value = 0
            $getprinter = Get-Printer
            $Treibername = "RICOH MP C3003 PCL 6"
            $OldPrinter = "FKOF1022A [Frauenfeld  Bau 1  022A](Mobility)"
            $druckername = "FKOF1022A"
            $LinkTreiber = $TreiberFKOF1022A
            PrintInstall
	        $label.Text = "Drucker $druckername wurde erfolgreich installiert"
            $progressBar.Value = 100
        }
        "KOF1022A"
        {
            cls
            $progressBar.Value = 0
            $getprinter = Get-Printer
            $Treibername = "RICOH MP 6503 PCL 6"
            $OldPrinter = "KOF1022A [Frauenfeld  Bau 1  022A](Mobility)"
            $druckername = "KOF1022A"
            $LinkTreiber = $TreiberKOF1022A
	        PrintInstall
	        $label.Text = "Drucker $druckername wurde erfolgreich installiert"            $progressBar.Value = 100
        }
        "FKOFC101"
        {
            cls
            $progressBar.Value = 0
            $getprinter = Get-Printer
            $Treibername = "RICOH MP C6004ex PCL 6"
            $OldPrinter = "FKOFC101 [Frauenfeld  Bau C  Kor 2.OG](Mobility)"
            $druckername = "FKOFC101"
            $LinkTreiber = $TreiberFKOFC101
            PrintInstall
	        $label.Text = "Drucker $druckername wurde erfolgreich installiert"
            $progressBar.Value = 100
        }
        "FKOFC201"
        {
            cls
            $progressBar.Value = 0
            $getprinter = Get-Printer
            $Treibername = "RICOH MP C3004 PCL 6"
            $OldPrinter = "FKOFC201 [Frauenfeld  Bau C  Kor 3.OG](Mobility)"
            $druckername = "FKOFC201"
            $LinkTreiber = $TreiberFKOFC201
            PrintInstall
	        $label.Text = "Drucker $druckername wurde erfolgreich installiert"
            $progressBar.Value = 100
        }
    }
})
$button2.Add_Click({
    cleandoubleprinters
    ClearTempFiles
})
# Show the form
$form.ShowDialog()

}

Forms
