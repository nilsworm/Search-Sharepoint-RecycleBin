#Abfrage zur Url der Sharepoint Seite
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$UrlForm = New-Object Windows.Forms.Form -Property @{
    StartPosition = [Windows.Forms.FormStartPosition]::CenterScreen
    Size          = New-Object Drawing.Size 300, 150
    Text          = 'Sharepoint Url'
    Topmost       = $true
}

$UrlLabel = New-Object System.Windows.Forms.Label -Property @{
    Location = New-Object System.Drawing.Point(30,10)
    Size = New-Object System.Drawing.Size(200,20)
    Text = 'Wer hat die Daten gelöscht?'
 }
 $UrlForm.Controls.Add($UrlLabel)

 $UrlTextBox = New-Object System.Windows.Forms.TextBox -Property @{
    Location = New-Object System.Drawing.Point(30,35)
    Size = New-Object System.Drawing.Size(200,20)
}
$UrlForm.Controls.Add($UrlTextBox)

$UrlConfirmButton = New-Object Windows.Forms.Button -Property @{
    Location     = New-Object Drawing.Point 80, 60
    Size         = New-Object Drawing.Size 75, 23
    Text         = 'Weiter'
    DialogResult = [Windows.Forms.DialogResult]::OK
    BackColor ="Green"
}
$UrlForm.CancelButton = $UrlConfirmButton
$UrlForm.Controls.Add($UrlConfirmButton)

$UrlFormResult = $UrlForm.ShowDialog()

#Genehmigung zum verbinden mit der Sharepoint Seite
if ($UrlFormResult -eq [Windows.Forms.DialogResult]::OK) {
    
#Link der Sharepoint Seite 
[string]$Url = $UrlTextBox.Text

#Verbinden mit der Sharepoint Seite
Connect-PnPOnline -Url $Url -UseWebLogin

#Abfrageform ob ein Datenbereich ausgewählt werden soll
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$DatepickerForm = New-Object Windows.Forms.Form -Property @{
    StartPosition = [Windows.Forms.FormStartPosition]::CenterScreen
    Size          = New-Object Drawing.Size 243, 400
    Text          = 'Daten auswählen'
    Topmost       = $true
}

$LabelDatepicker = New-Object System.Windows.Forms.Label -Property @{
   Location = New-Object System.Drawing.Point(0,0)
   Size = New-Object System.Drawing.Size(243,20)
   Text = 'Datenbereich angeben'
}
$DatepickerForm.Controls.Add($LabelDatepicker)

#Kalender Form
$calendar = New-Object Windows.Forms.MonthCalendar -Property @{
    ShowTodayCircle   = $false
    MaxSelectionCount = 100
    Location = New-Object Drawing.Point 5,25
}
$DatepickerForm.Controls.Add($calendar)

$DatepickerOkButton = New-Object Windows.Forms.Button -Property @{
    Location     = New-Object Drawing.Point 0, 200
    Size         = New-Object Drawing.Size 75, 23
    Text         = 'Weiter'
    DialogResult = [Windows.Forms.DialogResult]::OK
    BackColor = "Green"
}
$DatepickerForm.AcceptButton = $DatepickerOkButton
$DatepickerForm.Controls.Add($DatepickerOkButton)

$DatepickerCancelButton = New-Object Windows.Forms.Button -Property @{
    Location     = New-Object Drawing.Point 80, 200
    Size         = New-Object Drawing.Size 130, 23
    Text         = 'Ohne Datum suchen'
    DialogResult = [Windows.Forms.DialogResult]::Cancel
    BackColor ="Red"
}
$DatepickerForm.CancelButton = $DatepickerCancelButton
$DatepickerForm.Controls.Add($DatepickerCancelButton)

$DatepickerResult = $DatepickerForm.ShowDialog()

#Genehmigung zur weiteren Suche mit einem Datenbereich
If($DatepickerResult -eq [Windows.Forms.DialogResult]::OK){
   $StartDate = $calendar.SelectionStart
   $Enddate = $calendar.SelectionEnd
   $ConvertedStartDate = $StartDate.ToString("dd.MM.yy")
   $ConvertedEndDate = $Enddate.ToString("dd.MM.yy")

#Abfrageform für die weiteren Daten 
 Add-Type -AssemblyName System.Windows.Forms
 Add-Type -AssemblyName System.Drawing

 $form = New-Object Windows.Forms.Form -Property @{
     StartPosition = [Windows.Forms.FormStartPosition]::CenterScreen
     Size          = New-Object Drawing.Size 250, 300
     Text          = 'Papierkorb durchsuchen'
     Topmost       = $true
 }

 $LabelSelectedDate = New-Object System.Windows.Forms.Label -Property @{
    Location = New-Object System.Drawing.Point(0,0)
    Size = New-Object System.Drawing.Size(200,20)
    Text = If($StartDate -eq $Enddate){
        "Papierkorb am $ConvertedStartDate" 
    } else {"Papierkorb vom $ConvertedStartDate bis $ConvertedEndDate"}
 }
 $form.Controls.Add($LabelSelectedDate)

 $LabelDeleted = New-Object System.Windows.Forms.Label -Property @{
    Location = New-Object System.Drawing.Point(0,25)
    Size = New-Object System.Drawing.Size(200,20)
    Text = 'Wer hat die Daten gelöscht?'
 }
 $form.Controls.Add($LabelDeleted)

#Wer hat die Dateien gelöscht?
 $Deletedtextbox = New-Object System.Windows.Forms.TextBox -Property @{
    Location = New-Object System.Drawing.Point(0,50)
    Size = New-Object System.Drawing.Size(200,20)
}
$form.Controls.Add($Deletedtextbox)

#Wer hat die Dateien erstellt?
$Addedtextbox = New-Object System.Windows.Forms.TextBox -Property @{
    Location = New-Object System.Drawing.Point(0,100)
    Size = New-Object System.Drawing.Size(200,20)
}
$form.Controls.Add($Addedtextbox)

$LabelAdded = New-Object System.Windows.Forms.Label -Property @{
    Location = New-Object System.Drawing.Point(0,75)
    Size = New-Object System.Drawing.Size(200,20)
    Text = 'Wer hat die Daten erstellt?'
 }
 $form.Controls.Add($LabelAdded)
 
#Wo liegen die Dateien?
$Pathtextbox = New-Object System.Windows.Forms.TextBox -Property @{
    Location = New-Object System.Drawing.Point(0,150)
    Size = New-Object System.Drawing.Size(200,20)
}
$form.Controls.Add($Pathtextbox)

$LabelPath = New-Object System.Windows.Forms.Label -Property @{
    Location = New-Object System.Drawing.Point(0,125)
    Size = New-Object System.Drawing.Size(200,20)
    Text = 'Wo liegen die Daten?'
 }
 $form.Controls.Add($LabelPath)

 $LabelName = New-Object System.Windows.Forms.Label -Property @{
    Location = New-Object System.Drawing.Point(0,175)
    Size = New-Object System.Drawing.Size(200,20)
    Text = 'Wie heißen die Daten?'
 }
 $form.Controls.Add($LabelName)

 $Nametextbox = New-Object System.Windows.Forms.TextBox -Property @{
    Location = New-Object System.Drawing.Point(0,200)
    Size = New-Object System.Drawing.Size(200,20)
}
$form.Controls.Add($Nametextbox)

 #Bestätigung der Suche
 $okButton = New-Object Windows.Forms.Button -Property @{
     Location     = New-Object Drawing.Point 0, 225
     Size         = New-Object Drawing.Size 75, 23
     Text         = 'Suchen'
     DialogResult = [Windows.Forms.DialogResult]::OK
     BackColor = "Green"
 }
 $form.AcceptButton = $okButton
 $form.Controls.Add($okButton)

 #Abbrechen der Suche
 $cancelButton = New-Object Windows.Forms.Button -Property @{
     Location     = New-Object Drawing.Point 80, 225
     Size         = New-Object Drawing.Size 75, 23
     Text         = 'Abbruch'
     DialogResult = [Windows.Forms.DialogResult]::Cancel
     BackColor = "Red"
 }
 $form.CancelButton = $cancelButton
 $form.Controls.Add($cancelButton)
 
 $result = $form.ShowDialog()

 #Eingegebene Daten empfangen
 if ($result -eq [Windows.Forms.DialogResult]::OK) {
     $Startdate
     $EndDate
     $DeletedInput = $Deletedtextbox.Text
     $AddedInput = $Addedtextbox.Text
     $PathInput = $Pathtextbox.Text
     $NameInput = $Nametextbox.Text

     If($DeletedInput.Length -eq 0 -and $AddedInput.Length -eq 0 -and $PathInput.Length -eq 0 -and $NameInput.Length -ne 0) {
         $RecycleBinItems = Get-PnPRecycleBinItem -RowLimit 500000| Select-Object -Property DeletedDate, DeletedByName,AuthorName, DirName, Title | Where-Object { (Get-Date($_.DeletedDate)) -ge (Get-Date($StartDate)) -and (Get-Date($_.DeletedDate)) -le (Get-Date($Enddate.AddDays(1))) -and $_.Title -like "*$NameInput*"} | Sort-Object DeletedDate -Descending
         $RecycleBinItems | Out-GridView -PassThru | Restore-PnPRecycleBinItem      
     }

     elseif ($DeletedInput.Length -eq 0 -and $AddedInput.Length -eq 0 -and $PathInput.Length -ne 0 -and $NameInput.Length -eq 0) {
         $RecycleBinItems = Get-PnPRecycleBinItem -RowLimit 500000| Select-Object -Property DeletedDate, DeletedByName,AuthorName, DirName, Title | Where-Object { (Get-Date($_.DeletedDate)) -ge (Get-Date($StartDate)) -and (Get-Date($_.DeletedDate)) -le (Get-Date($Enddate.AddDays(1))) -and $_.DirName -like "*$PathInput*"} | Sort-Object DeletedDate -Descending
         $RecycleBinItems | Out-GridView -PassThru | Restore-PnPRecycleBinItem
     }
       
     elseif ($DeletedInput.Length -eq 0 -and $AddedInput.Length -eq 0 -and $PathInput.Length -ne 0 -and $NameInput.Length -ne 0) {
         $RecycleBinItems = Get-PnPRecycleBinItem -RowLimit 500000| Select-Object -Property DeletedDate, DeletedByName,AuthorName, DirName, Title | Where-Object { (Get-Date($_.DeletedDate)) -ge (Get-Date($StartDate)) -and (Get-Date($_.DeletedDate)) -le (Get-Date($Enddate.AddDays(1))) -and $_.DirName -like "*$PathInput*" -and $_.Title -like "*$NameInput*"} | Sort-Object DeletedDate -Descending
         $RecycleBinItems | Out-GridView -PassThru | Restore-PnPRecycleBinItem
     }

     elseif ($DeletedInput.Length -eq 0 -and $AddedInput.Length -ne 0 -and $PathInput.Length -eq 0 -and $NameInput.Length -eq 0) {
         $RecycleBinItems = Get-PnPRecycleBinItem -RowLimit 500000| Select-Object -Property DeletedDate, DeletedByName,AuthorName, DirName, Title | Where-Object { (Get-Date($_.DeletedDate)) -ge (Get-Date($StartDate)) -and (Get-Date($_.DeletedDate)) -le (Get-Date($Enddate.AddDays(1))) -and $_.AuthorName -like "*$AddedInput*"} | Sort-Object DeletedDate -Descending
         $RecycleBinItems | Out-GridView -PassThru | Restore-PnPRecycleBinItem
     }

     elseif ($DeletedInput.Length -eq 0 -and $AddedInput.Length -ne 0 -and $PathInput.Length -eq 0 -and $NameInput.Length -ne 0) {
         $RecycleBinItems = Get-PnPRecycleBinItem -RowLimit 500000| Select-Object -Property DeletedDate, DeletedByName,AuthorName, DirName, Title | Where-Object { (Get-Date($_.DeletedDate)) -ge (Get-Date($StartDate)) -and (Get-Date($_.DeletedDate)) -le (Get-Date($Enddate.AddDays(1))) -and $_.AuthorName -like "*$AddedInput*" -and $_.Title -like "*$NameInput*"} | Sort-Object DeletedDate -Descending
         $RecycleBinItems | Out-GridView -PassThru | Restore-PnPRecycleBinItem
     }

     elseif ($DeletedInput.Length -eq 0 -and $AddedInput.Length -ne 0 -and $PathInput.Length -ne 0 -and $NameInput.Length -eq 0) {
         $RecycleBinItems = Get-PnPRecycleBinItem -RowLimit 500000| Select-Object -Property DeletedDate, DeletedByName,AuthorName, DirName, Title | Where-Object { (Get-Date($_.DeletedDate)) -ge (Get-Date($StartDate)) -and (Get-Date($_.DeletedDate)) -le (Get-Date($Enddate.AddDays(1))) -and $_.AuthorName -like "*$AddedInput*" -and $_.DirName -like "*$PathInput*"} | Sort-Object DeletedDate -Descending
         $RecycleBinItems | Out-GridView -PassThru | Restore-PnPRecycleBinItem
     }

     elseif ($DeletedInput.Length -eq 0 -and $AddedInput.Length -ne 0 -and $PathInput.Length -ne 0 -and $NameInput.Length -ne 0) {
         $RecycleBinItems = Get-PnPRecycleBinItem -RowLimit 500000| Select-Object -Property DeletedDate, DeletedByName,AuthorName, DirName, Title | Where-Object { (Get-Date($_.DeletedDate)) -ge (Get-Date($StartDate)) -and (Get-Date($_.DeletedDate)) -le (Get-Date($Enddate.AddDays(1))) -and $_.AuthorName -like "*$AddedInput*" -and $_.DirName -like "*$PathInput*" -and $_.Title -like "*$NameInput*"} | Sort-Object DeletedDate -Descending
         $RecycleBinItems | Out-GridView -PassThru | Restore-PnPRecycleBinItem
     }

     elseif ($DeletedInput.Length -ne 0 -and $AddedInput.Length -eq 0 -and $PathInput.Length -eq 0 -and $NameInput.Length -eq 0) {
         $RecycleBinItems = Get-PnPRecycleBinItem -RowLimit 500000| Select-Object -Property DeletedDate, DeletedByName,AuthorName, DirName, Title | Where-Object { (Get-Date($_.DeletedDate)) -ge (Get-Date($StartDate)) -and (Get-Date($_.DeletedDate)) -le (Get-Date($Enddate.AddDays(1))) -and $_.DeletedByName -like "*$DeletedInput*"} | Sort-Object DeletedDate -Descending
         $RecycleBinItems | Out-GridView -PassThru | Restore-PnPRecycleBinItem
     }

     elseif ($DeletedInput.Length -ne 0 -and $AddedInput.Length -eq 0 -and $PathInput.Length -eq 0 -and $NameInput.Length -ne 0) {
         $RecycleBinItems = Get-PnPRecycleBinItem -RowLimit 500000| Select-Object -Property DeletedDate, DeletedByName,AuthorName, DirName, Title | Where-Object { (Get-Date($_.DeletedDate)) -ge (Get-Date($StartDate)) -and (Get-Date($_.DeletedDate)) -le (Get-Date($Enddate.AddDays(1))) -and $_.DeletedByName -like "*$DeletedInput*" -and $_.Title -like "*$NameInput*"} | Sort-Object DeletedDate -Descending
         $RecycleBinItems | Out-GridView -PassThru | Restore-PnPRecycleBinItem
     }

     elseif ($DeletedInput.Length -ne 0 -and $AddedInput.Length -eq 0 -and $PathInput.Length -ne 0 -and $NameInput.Length -eq 0) {
         $RecycleBinItems = Get-PnPRecycleBinItem -RowLimit 500000| Select-Object -Property DeletedDate, DeletedByName,AuthorName, DirName, Title | Where-Object { (Get-Date($_.DeletedDate)) -ge (Get-Date($StartDate)) -and (Get-Date($_.DeletedDate)) -le (Get-Date($Enddate.AddDays(1))) -and $_.DeletedByName -like "*$DeletedInput*" -and $_.DirName -like "*$PathInput*"} | Sort-Object DeletedDate -Descending
         $RecycleBinItems | Out-GridView -PassThru | Restore-PnPRecycleBinItem
     }

     elseif ($DeletedInput.Length -ne 0 -and $AddedInput.Length -eq 0 -and $PathInput.Length -ne 0 -and $NameInput.Length -ne 0) {
         $RecycleBinItems = Get-PnPRecycleBinItem -RowLimit 500000| Select-Object -Property DeletedDate, DeletedByName,AuthorName, DirName, Title | Where-Object { (Get-Date($_.DeletedDate)) -ge (Get-Date($StartDate)) -and (Get-Date($_.DeletedDate)) -le (Get-Date($Enddate.AddDays(1))) -and $_.DeletedByName -like "*$DeletedInput*" -and $_.DirName -like "*$PathInput*" -and $_.Title -like "*$NameInput*"} | Sort-Object DeletedDate -Descending
         $RecycleBinItems | Out-GridView -PassThru | Restore-PnPRecycleBinItem
     }

     elseif ($DeletedInput.Length -ne 0 -and $AddedInput.Length -ne 0 -and $PathInput.Length -eq 0 -and $NameInput.Length -eq 0) {
         $RecycleBinItems = Get-PnPRecycleBinItem -RowLimit 500000| Select-Object -Property DeletedDate, DeletedByName,AuthorName, DirName, Title | Where-Object { (Get-Date($_.DeletedDate)) -ge (Get-Date($StartDate)) -and (Get-Date($_.DeletedDate)) -le (Get-Date($Enddate.AddDays(1))) -and $_.DeletedByName -like "*$DeletedInput*" -and $_.AuthorName -like "*$AddedInput*"} | Sort-Object DeletedDate -Descending
         $RecycleBinItems | Out-GridView -PassThru | Restore-PnPRecycleBinItem
     }

     elseif ($DeletedInput.Length -ne 0 -and $AddedInput.Length -ne 0 -and $PathInput.Length -eq 0 -and $NameInput.Length -ne 0) {
         $RecycleBinItems = Get-PnPRecycleBinItem -RowLimit 500000| Select-Object -Property DeletedDate, DeletedByName,AuthorName, DirName, Title | Where-Object { (Get-Date($_.DeletedDate)) -ge (Get-Date($StartDate)) -and (Get-Date($_.DeletedDate)) -le (Get-Date($Enddate.AddDays(1))) -and $_.DeletedByName -like "*$DeletedInput*" -and $_.AuthorName -like "*$AddedInput*" -and $_.Title -like "*$NameInput*"} | Sort-Object DeletedDate -Descending
         $RecycleBinItems | Out-GridView -PassThru | Restore-PnPRecycleBinItem
     }
     elseif ($DeletedInput.Length -ne 0 -and $AddedInput.Length -ne 0 -and $PathInput.Length -eq 0 -and $NameInput.Length -ne 0) {
         $RecycleBinItems = Get-PnPRecycleBinItem -RowLimit 500000| Select-Object -Property DeletedDate, DeletedByName,AuthorName, DirName, Title | Where-Object { (Get-Date($_.DeletedDate)) -ge (Get-Date($StartDate)) -and (Get-Date($_.DeletedDate)) -le (Get-Date($Enddate.AddDays(1))) -and $_.DeletedByName -like "*$DeletedInput*" -and $_.AuthorName -like "*$AddedInput*" -and $_.Title -like "*$NameInput*"} | Sort-Object DeletedDate -Descending
         $RecycleBinItems | Out-GridView -PassThru | Restore-PnPRecycleBinItem
     }

     elseif ($DeletedInput.Length -ne 0 -and $AddedInput.Length -ne 0 -and $PathInput.Length -ne 0 -and $NameInput.Length -eq 0) {
         $RecycleBinItems = Get-PnPRecycleBinItem -RowLimit 500000| Select-Object -Property DeletedDate, DeletedByName,AuthorName, DirName, Title | Where-Object { (Get-Date($_.DeletedDate)) -ge (Get-Date($StartDate)) -and (Get-Date($_.DeletedDate)) -le (Get-Date($Enddate.AddDays(1))) -and $_.DeletedByName -like "*$DeletedInput*" -and $_.AuthorName -like "*$AddedInput*" -and $_.DirName -like "*$PathInput*"} | Sort-Object DeletedDate -Descending
         $RecycleBinItems | Out-GridView -PassThru | Restore-PnPRecycleBinItem
     }

     elseif ($DeletedInput.Length -ne 0 -and $AddedInput.Length -ne 0 -and $PathInput.Length -ne 0 -and $NameInput.Length -ne 0) {
         $RecycleBinItems = Get-PnPRecycleBinItem -RowLimit 500000| Select-Object -Property DeletedDate, DeletedByName,AuthorName, DirName, Title | Where-Object { (Get-Date($_.DeletedDate)) -ge (Get-Date($StartDate)) -and (Get-Date($_.DeletedDate)) -le (Get-Date($Enddate.AddDays(1))) -and $_.DeletedByName -like "*$DeletedInput*" -and $_.AuthorName -like "*$AddedInput*" -and $_.DirName -like "*$PathInput*" -and $_.Title -like "*$NameInput*"} | Sort-Object DeletedDate -Descending
         $RecycleBinItems | Out-GridView -PassThru | Restore-PnPRecycleBinItem
     }
}
} elseif ($DatepickerResult -eq [Windows.Forms.DialogResult]::Cancel) {

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
   
    $form = New-Object Windows.Forms.Form -Property @{
        StartPosition = [Windows.Forms.FormStartPosition]::CenterScreen
        Size          = New-Object Drawing.Size 250, 275
        Text          = 'Papierkorb durchsuchen'
        Topmost       = $true
    }
   
    $LabelDeleted = New-Object System.Windows.Forms.Label -Property @{
       Location = New-Object System.Drawing.Point(0,0)
       Size = New-Object System.Drawing.Size(200,20)
       Text = 'Wer hat die Daten gelöscht?'
    }
    $form.Controls.Add($LabelDeleted)
   
   #Wer hat die Dateien gelöscht?
    $Deletedtextbox = New-Object System.Windows.Forms.TextBox -Property @{
       Location = New-Object System.Drawing.Point(0,25)
       Size = New-Object System.Drawing.Size(200,20)
   }
   $form.Controls.Add($Deletedtextbox)
   
   #Wer hat die Dateien erstellt?
   $Addedtextbox = New-Object System.Windows.Forms.TextBox -Property @{
       Location = New-Object System.Drawing.Point(0,75)
       Size = New-Object System.Drawing.Size(200,20)
   }
   $form.Controls.Add($Addedtextbox)
   
   $LabelAdded = New-Object System.Windows.Forms.Label -Property @{
       Location = New-Object System.Drawing.Point(0,50)
       Size = New-Object System.Drawing.Size(200,20)
       Text = 'Wer hat die Daten erstellt?'
    }
    $form.Controls.Add($LabelAdded)
    
   #Wo liegen die Dateien?
   $Pathtextbox = New-Object System.Windows.Forms.TextBox -Property @{
       Location = New-Object System.Drawing.Point(0,125)
       Size = New-Object System.Drawing.Size(200,20)
   }
   $form.Controls.Add($Pathtextbox)
   
   $LabelPath = New-Object System.Windows.Forms.Label -Property @{
       Location = New-Object System.Drawing.Point(0,100)
       Size = New-Object System.Drawing.Size(200,20)
       Text = 'Wo liegen die Daten?'
    }
    $form.Controls.Add($LabelPath)

    $LabelName = New-Object System.Windows.Forms.Label -Property @{
        Location = New-Object System.Drawing.Point(0,150)
        Size = New-Object System.Drawing.Size(200,20)
        Text = 'Wie heißen die Daten?'
     }
     $form.Controls.Add($LabelName)

     $Nametextbox = New-Object System.Windows.Forms.TextBox -Property @{
        Location = New-Object System.Drawing.Point(0,175)
        Size = New-Object System.Drawing.Size(200,20)
    }
    $form.Controls.Add($Nametextbox)

   
    #Bestätigung der Suche
    $okButton = New-Object Windows.Forms.Button -Property @{
        Location     = New-Object Drawing.Point 0, 200
        Size         = New-Object Drawing.Size 75, 23
        Text         = 'Suchen'
        DialogResult = [Windows.Forms.DialogResult]::OK
        BackColor = "Green"
    }
    $form.AcceptButton = $okButton
    $form.Controls.Add($okButton)
   
    #Abbrechen der Suche
    $cancelButton = New-Object Windows.Forms.Button -Property @{
        Location     = New-Object Drawing.Point 80, 200
        Size         = New-Object Drawing.Size 75, 23
        Text         = 'Abbruch'
        DialogResult = [Windows.Forms.DialogResult]::Cancel
        BackColor = "Red"
    }
    $form.CancelButton = $cancelButton
    $form.Controls.Add($cancelButton)
    
    $result = $form.ShowDialog()

    if ($result -eq [Windows.Forms.DialogResult]::OK) {
        $DeletedInput = $Deletedtextbox.Text
        $AddedInput = $Addedtextbox.Text
        $PathInput = $Pathtextbox.Text
        $NameInput = $Nametextbox.Text

        If($DeletedInput.Length -eq 0 -and $AddedInput.Length -eq 0 -and $PathInput.Length -eq 0 -and $NameInput.Length -ne 0) {
            $RecycleBinItems = Get-PnPRecycleBinItem -RowLimit 500000| Select-Object -Property DeletedDate, DeletedByName,AuthorName, DirName, Title | Where-Object { $_.Title -like "*$NameInput*"} | Sort-Object DeletedDate -Descending
            $RecycleBinItems | Out-GridView -PassThru | Restore-PnPRecycleBinItem      
        }

        elseif ($DeletedInput.Length -eq 0 -and $AddedInput.Length -eq 0 -and $PathInput.Length -ne 0 -and $NameInput.Length -eq 0) {
            $RecycleBinItems = Get-PnPRecycleBinItem -RowLimit 500000| Select-Object -Property DeletedDate, DeletedByName,AuthorName, DirName, Title | Where-Object { $_.DirName -like "*$PathInput*"} | Sort-Object DeletedDate -Descending
            $RecycleBinItems | Out-GridView -PassThru | Restore-PnPRecycleBinItem
        }
          
        elseif ($DeletedInput.Length -eq 0 -and $AddedInput.Length -eq 0 -and $PathInput.Length -ne 0 -and $NameInput.Length -ne 0) {
            $RecycleBinItems = Get-PnPRecycleBinItem -RowLimit 500000| Select-Object -Property DeletedDate, DeletedByName,AuthorName, DirName, Title | Where-Object { $_.DirName -like "*$PathInput*" -and $_.Title -like "*$NameInput*"} | Sort-Object DeletedDate -Descending
            $RecycleBinItems | Out-GridView -PassThru | Restore-PnPRecycleBinItem
        }

        elseif ($DeletedInput.Length -eq 0 -and $AddedInput.Length -ne 0 -and $PathInput.Length -eq 0 -and $NameInput.Length -eq 0) {
            $RecycleBinItems = Get-PnPRecycleBinItem -RowLimit 500000| Select-Object -Property DeletedDate, DeletedByName,AuthorName, DirName, Title | Where-Object { $_.AuthorName -like "*$AddedInput*"} | Sort-Object DeletedDate -Descending
            $RecycleBinItems | Out-GridView -PassThru | Restore-PnPRecycleBinItem
        }

        elseif ($DeletedInput.Length -eq 0 -and $AddedInput.Length -ne 0 -and $PathInput.Length -eq 0 -and $NameInput.Length -ne 0) {
            $RecycleBinItems = Get-PnPRecycleBinItem -RowLimit 500000| Select-Object -Property DeletedDate, DeletedByName,AuthorName, DirName, Title | Where-Object { $_.AuthorName -like "*$AddedInput*" -and $_.Title -like "*$NameInput*"} | Sort-Object DeletedDate -Descending
            $RecycleBinItems | Out-GridView -PassThru | Restore-PnPRecycleBinItem
        }

        elseif ($DeletedInput.Length -eq 0 -and $AddedInput.Length -ne 0 -and $PathInput.Length -ne 0 -and $NameInput.Length -eq 0) {
            $RecycleBinItems = Get-PnPRecycleBinItem -RowLimit 500000| Select-Object -Property DeletedDate, DeletedByName,AuthorName, DirName, Title | Where-Object { $_.AuthorName -like "*$AddedInput*" -and $_.DirName -like "*$PathInput*"} | Sort-Object DeletedDate -Descending
            $RecycleBinItems | Out-GridView -PassThru | Restore-PnPRecycleBinItem
        }

        elseif ($DeletedInput.Length -eq 0 -and $AddedInput.Length -ne 0 -and $PathInput.Length -ne 0 -and $NameInput.Length -ne 0) {
            $RecycleBinItems = Get-PnPRecycleBinItem -RowLimit 500000| Select-Object -Property DeletedDate, DeletedByName,AuthorName, DirName, Title | Where-Object { $_.AuthorName -like "*$AddedInput*" -and $_.DirName -like "*$PathInput*" -and $_.Title -like "*$NameInput*"} | Sort-Object DeletedDate -Descending
            $RecycleBinItems | Out-GridView -PassThru | Restore-PnPRecycleBinItem
        }

        elseif ($DeletedInput.Length -ne 0 -and $AddedInput.Length -eq 0 -and $PathInput.Length -eq 0 -and $NameInput.Length -eq 0) {
            $RecycleBinItems = Get-PnPRecycleBinItem -RowLimit 500000| Select-Object -Property DeletedDate, DeletedByName,AuthorName, DirName, Title | Where-Object { $_.DeletedByName -like "*$DeletedInput*"} | Sort-Object DeletedDate -Descending
            $RecycleBinItems | Out-GridView -PassThru | Restore-PnPRecycleBinItem
        }

        elseif ($DeletedInput.Length -ne 0 -and $AddedInput.Length -eq 0 -and $PathInput.Length -eq 0 -and $NameInput.Length -ne 0) {
            $RecycleBinItems = Get-PnPRecycleBinItem -RowLimit 500000| Select-Object -Property DeletedDate, DeletedByName,AuthorName, DirName, Title | Where-Object { $_.DeletedByName -like "*$DeletedInput*" -and $_.Title -like "*$NameInput*"} | Sort-Object DeletedDate -Descending
            $RecycleBinItems | Out-GridView -PassThru | Restore-PnPRecycleBinItem
        }

        elseif ($DeletedInput.Length -ne 0 -and $AddedInput.Length -eq 0 -and $PathInput.Length -ne 0 -and $NameInput.Length -eq 0) {
            $RecycleBinItems = Get-PnPRecycleBinItem -RowLimit 500000| Select-Object -Property DeletedDate, DeletedByName,AuthorName, DirName, Title | Where-Object { $_.DeletedByName -like "*$DeletedInput*" -and $_.DirName -like "*$PathInput*"} | Sort-Object DeletedDate -Descending
            $RecycleBinItems | Out-GridView -PassThru | Restore-PnPRecycleBinItem
        }

        elseif ($DeletedInput.Length -ne 0 -and $AddedInput.Length -eq 0 -and $PathInput.Length -ne 0 -and $NameInput.Length -ne 0) {
            $RecycleBinItems = Get-PnPRecycleBinItem -RowLimit 500000| Select-Object -Property DeletedDate, DeletedByName,AuthorName, DirName, Title | Where-Object { $_.DeletedByName -like "*$DeletedInput*" -and $_.DirName -like "*$PathInput*" -and $_.Title -like "*$NameInput*"} | Sort-Object DeletedDate -Descending
            $RecycleBinItems | Out-GridView -PassThru | Restore-PnPRecycleBinItem
        }

        elseif ($DeletedInput.Length -ne 0 -and $AddedInput.Length -ne 0 -and $PathInput.Length -eq 0 -and $NameInput.Length -eq 0) {
            $RecycleBinItems = Get-PnPRecycleBinItem -RowLimit 500000| Select-Object -Property DeletedDate, DeletedByName,AuthorName, DirName, Title | Where-Object { $_.DeletedByName -like "*$DeletedInput*" -and $_.AuthorName -like "*$AddedInput*"} | Sort-Object DeletedDate -Descending
            $RecycleBinItems | Out-GridView -PassThru | Restore-PnPRecycleBinItem
        }

        elseif ($DeletedInput.Length -ne 0 -and $AddedInput.Length -ne 0 -and $PathInput.Length -eq 0 -and $NameInput.Length -ne 0) {
            $RecycleBinItems = Get-PnPRecycleBinItem -RowLimit 500000| Select-Object -Property DeletedDate, DeletedByName,AuthorName, DirName, Title | Where-Object { $_.DeletedByName -like "*$DeletedInput*" -and $_.AuthorName -like "*$AddedInput*" -and $_.Title -like "*$NameInput*"} | Sort-Object DeletedDate -Descending
            $RecycleBinItems | Out-GridView -PassThru | Restore-PnPRecycleBinItem
        }
        elseif ($DeletedInput.Length -ne 0 -and $AddedInput.Length -ne 0 -and $PathInput.Length -eq 0 -and $NameInput.Length -ne 0) {
            $RecycleBinItems = Get-PnPRecycleBinItem -RowLimit 500000| Select-Object -Property DeletedDate, DeletedByName,AuthorName, DirName, Title | Where-Object { $_.DeletedByName -like "*$DeletedInput*" -and $_.AuthorName -like "*$AddedInput*" -and $_.Title -like "*$NameInput*"} | Sort-Object DeletedDate -Descending
            $RecycleBinItems | Out-GridView -PassThru | Restore-PnPRecycleBinItem
        }

        elseif ($DeletedInput.Length -ne 0 -and $AddedInput.Length -ne 0 -and $PathInput.Length -ne 0 -and $NameInput.Length -eq 0) {
            $RecycleBinItems = Get-PnPRecycleBinItem -RowLimit 500000| Select-Object -Property DeletedDate, DeletedByName,AuthorName, DirName, Title | Where-Object { $_.DeletedByName -like "*$DeletedInput*" -and $_.AuthorName -like "*$AddedInput*" -and $_.DirName -like "*$PathInput*"} | Sort-Object DeletedDate -Descending
            $RecycleBinItems | Out-GridView -PassThru | Restore-PnPRecycleBinItem
        }

        elseif ($DeletedInput.Length -ne 0 -and $AddedInput.Length -ne 0 -and $PathInput.Length -ne 0 -and $NameInput.Length -ne 0) {
            $RecycleBinItems = Get-PnPRecycleBinItem -RowLimit 500000| Select-Object -Property DeletedDate, DeletedByName,AuthorName, DirName, Title | Where-Object { $_.DeletedByName -like "*$DeletedInput*" -and $_.AuthorName -like "*$AddedInput*" -and $_.DirName -like "*$PathInput*" -and $_.Title -like "*$NameInput*"} | Sort-Object DeletedDate -Descending
            $RecycleBinItems | Out-GridView -PassThru | Restore-PnPRecycleBinItem
        }
    } 
}
}


 
 

 



