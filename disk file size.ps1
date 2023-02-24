Add-Type -AssemblyName System.Windows.Forms

# Create the GUI form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Get Folder Report"
$form.Width = 300
$form.Height = 200

# Create a button and add it to the form
$button = New-Object System.Windows.Forms.Button
$button.Location = New-Object System.Drawing.Point(50,50)
$button.Size = New-Object System.Drawing.Size(200,50)
$button.Text = "Select Folder"
$form.Controls.Add($button)

# Define the button click event
$button.Add_Click({
    $global:browser = New-Object System.Windows.Forms.FolderBrowserDialog
    $global:browser.Description = "Select a folder or drive to create a report of"
    $global:browser.RootFolder = [System.Environment+SpecialFolder]::MyComputer
   

 if($global:browser.ShowDialog() -eq 'OK'){
        $Path = $global:browser.SelectedPath
        $global:browser.Dispose()

        # Create a SaveFileDialog to allow the user to select the location where the report will be generated
        $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
        $saveDialog.Title = "Save report file"
        $saveDialog.Filter = "Text files (*.txt)|*.txt"
        $saveDialog.InitialDirectory = [System.IO.Path]::GetDirectoryName($Path)
        $saveDialog.FileName = "report.txt"

 if ($saveDialog.ShowDialog() -eq 'OK') {
            $reportPath = $saveDialog.FileName
        

        Get-ChildItem $Path -Recurse -ErrorAction SilentlyContinue |
        Sort-Object Length -Descending |
        Select-Object @{Name='Filename';Expression={$_.Name}}, @{Name='Size(MB)';Expression={$_.Length / 1MB}} |
        Sort-Object 'Size(MB)' -Descending|
        Format-Table -AutoSize |
        Out-String -Width 4096 |
        Out-File $reportPath -Force

        # Show a message box when the report is generated
        [System.Windows.Forms.MessageBox]::Show("Report generated successfully!")
        $global:browser.Dispose()
        $form.Dispose()
    }
   }

})
 

# Show the form
$form.ShowDialog()



