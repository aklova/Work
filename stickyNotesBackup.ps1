# Checking all subfolders in C:\Users
$users = Get-ChildItem -Path C:\Users

foreach ($user in $users)
{
	# Going through all users Sticky Notes sqlite file one by one
    $DataSourse = "C:\Users\$user\AppData\Local\Packages\Microsoft.MicrosoftStickyNotes_8wekyb3d8bbwe\LocalState\plum.sqlite"

    try
    {
	# Trying to extraxt data from sqlite file
    $text = Invoke-SqliteQuery -DataSource $DataSourse -Query "SELECT Text FROM Note"
    Write-Host "Saving $user's data"
	# Putting everything in C:\temp\username txt file, with '-replace "id.*?' removing not necessary ids from sqlite file
    $newText = $text.Text -replace "id.*? " | Out-File -FilePath (New-Item -Path "C:\temp\Sticky Notes Backup\$user\stickybackup.txt" -Force)
    }
    Catch
    {
	# If there were no plum.sqlite file in user, going to the next user
    Write-Host "$user doesn't have sticky notes data"
    continue
    }
}
