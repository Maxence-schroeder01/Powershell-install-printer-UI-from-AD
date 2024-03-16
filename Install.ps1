Add-Type -AssemblyName System.Windows.Forms

# Fonction pour afficher un message dans une boîte de dialogue modale
function Show-ModalMessageBox {
    param(
        [string]$Message,
        [string]$Title,
        [System.Windows.Forms.MessageBoxButtons]$Buttons = [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]$Icon = [System.Windows.Forms.MessageBoxIcon]::Information
    )
    [System.Windows.Forms.MessageBox]::Show($Message, $Title, $Buttons, $Icon, [System.Windows.Forms.MessageBoxDefaultButton]::Button1)
}

# Fonction pour mettre à jour la liste des imprimantes dans la ListBox
function RefreshPrinterList {
    # Effacer la ListBox
    $listBox.Items.Clear()
    
    # Récupération de la liste des imprimantes installées localement
    $installedPrinters = [System.Drawing.Printing.PrinterSettings]::InstalledPrinters

    # Récupération de la liste des imprimantes de l'Active Directory
    $printers = Get-ADObject -Filter { objectClass -eq "printQueue" } -Property uNCName

    # Ajout des imprimantes à la ListBox, en filtrant celles qui sont déjà installées localement
    foreach ($printer in $printers) {
        $uncName = $printer.uNCName
        # Vérifier si l'imprimante n'est pas déjà installée localement
        if ($uncName -notin $installedPrinters) {
            # Utiliser une expression régulière pour extraire le nom du serveur et le nom de l'imprimante
            $printerName = $uncName -replace '^\\\\([^\\]+)\\(.+)$', '\\$1\$2'
            $listBox.Items.Add($printerName)
        }
    }
}

# Création d'une nouvelle instance de Form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Liste des Imprimantes"
$form.Size = New-Object System.Drawing.Size(800,700)

# Création d'une ListBox pour afficher la liste des imprimantes
$listBox = New-Object System.Windows.Forms.ListBox
$listBox.Size = New-Object System.Drawing.Size(600, 450)
$listBox.Location = New-Object System.Drawing.Point(50, 20)

# Appeler la fonction pour initialiser la ListBox
RefreshPrinterList

# Ajout de la ListBox au formulaire
$form.Controls.Add($listBox)

# Création d'un bouton pour installer l'imprimante
$installButton = New-Object System.Windows.Forms.Button
$installButton.Location = New-Object System.Drawing.Point(50, 500)
$installButton.Size = New-Object System.Drawing.Size(150, 30)
$installButton.Text = "Installer l'imprimante"

# Définition de l'événement Click pour le bouton Installer
$installButton.Add_Click({
    # Récupérer l'élément sélectionné dans la ListBox
    $selectedPrinter = $listBox.SelectedItem
    # Installer l'imprimante
    if ($selectedPrinter -ne $null) {
        $message = "Installation de l'imprimante : $selectedPrinter"
        Show-ModalMessageBox -Message $message -Title "Installation d'imprimante"
        Add-Printer -ConnectionName $selectedPrinter
        # Rafraîchir la liste des imprimantes après l'installation
        RefreshPrinterList
        $message = "L'imprimante $selectedPrinter a été installée avec succès."
        Show-ModalMessageBox -Message $message -Title "Installation d'imprimante"
    } else {
        Show-ModalMessageBox -Message "Veuillez sélectionner une imprimante à installer." -Title "Erreur"
    }
})

# Ajout du bouton à la formulaire
$form.Controls.Add($installButton)

# Affichage du formulaire
$form.ShowDialog() | Out-Null
