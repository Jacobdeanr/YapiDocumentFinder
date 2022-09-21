#---------------------------------------------------------------------
# GUI requirements
#---------------------------------------------------------------------
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 

Add-Type -AssemblyName PresentationCore,PresentationFramework
Add-Type -assembly System.Windows.Forms

#---------------------------------------------------------------------
# Create a basic textbox and button.
#---------------------------------------------------------------------
function Generate-Form{

	$main_form = New-Object System.Windows.Forms.Form
	$main_form.Text ='YAPI Document Finder'
	$main_form.Width = 600
	$main_form.Height = 100
	$main_form.AutoSize = $true

	$Label = New-Object System.Windows.Forms.Label
	$Label.Text = "GUID"
	$Label.Font = 'Microsoft Sans Serif,12'
	$Label.Location  = New-Object System.Drawing.Point(10,10)
	$Label.AutoSize = $true
	$main_form.Controls.Add($Label)

	$textbox = New-Object System.Windows.Forms.TextBox
	$textbox.Location = New-Object System.Drawing.Point(80,10)
	$textbox.Text = ""
	$textbox.Font = 'Microsoft Sans Serif,12'
	$textbox.Height = "24"
	$textbox.Width = "380"
	$textbox.MaxLength = "38" # GUIDs are 38 characters long.
	$textbox.AutoSize = $false
	$main_form.Controls.Add($textbox)

	$button = New-Object System.Windows.Forms.Button
	$Button.Location = New-Object System.Drawing.Point(470,10)
	$Button.Size = New-Object System.Drawing.Size(80, 24)
	$Button.Text = "Search"
	$Button.Font = "Microsoft Sans Serif,12"
	$Button.Add_Click({ReadForms})
	$main_form.Controls.Add($Button)

	$main_form.ShowDialog()| Out-Null

}

#---------------------------------------------------------------------
# This function name is awful. We basically are creating a few
# variables and concecrating them into several arrays.
#---------------------------------------------------------------------
function ReadForms {
    # Combine the GUID with the website URL to grab a list of forms
    $path = "https://yapi.me/f/p/list.php?guid="
    
    $GUID = $textbox.Text
    $GUID = $GUID -replace '\s','' # cleanup any invalid space values.
    $domain = $path+$GUID
    
    # Test if the GUID domain is good.
    if(($xml = Invoke-RestMethod -uri $domain) -like "*G-Invalid GUID*") {
		[System.Windows.MessageBox]::Show("Nothing Found, make sure the GUID was entered correctly.",'Error','OK','Error')  # Error handling for invalid entries.
    }
    else{

        $duplicates = $null #TODO: Implement Duplicates
        $str_FormNames_Array = @() #Initialize array used to build the list of forms.
        $uniqueNamesArray = @() #Initialize array for unique names
        $urlArray = @() #Initialize array for form URLs

        # Build the resulting domain using a ton of concatenate string variables. I'm sure theres a better way for this but it works.
        $first = "https://yapi.me/f/"
        $practice = "&guid="
        $tag = "&tag=[[TAGID]]"

        # For each file that has an @oxygen attribute loop over it and assign variables.
        $lists = $xml.SelectNodes("//list/form[@oxygen='1']") | ForEach-Object { 
            $formName = $_.name
            $path = $_.url
            $id = $_.id
            
            $formNameSeparator = $formName + ':'
            $form =  $first + "$path" + $practice + $GUID + $tag

            $urlArray += $form; # Array of the form urls

            $str_FormNames_Array += $formNameSeparator = $formName + ': ' + $id; #Add each Form Name into the Array
        }

        # TODO: Implment Color the GUI list with unique and duplicate colors based on $uniqueNamesArray
        $uniqueNamesArray = $str_FormNames_Array | select -unique #unique values in the array
        
        #TODO: Implement Duplicates.
        Try{
            $duplicates = Compare-Object -ReferenceObject $uniqueNamesArray -DifferenceObject $str_FormNames_Array | ForEach-Object{
                $dupeNames = $_.InputObject 
                #Write-Host $dupeNames -ForegroundColor Red -BackgroundColor Black
            }
            if($dupeNames -eq $null){
                #TODO: Implement Duplicates
                #Write-Host `n"No Duplicates :)"`n -ForegroundColor Green
                
            }
        }
        Catch{
         #TODO: Implement Duplicates Error Catch
         #Write-Host "An Error Occured"`n -ForegroundColor Red
        }
    
    Generate-Links($str_FormNames_Array)
    }
}

#---------------------------------------------------------------------
# Create a new window with a list of forms that can be selected.
#---------------------------------------------------------------------
function Generate-Links($str_FormNames_Array){

    #Basic Container
    $FormListMenu = New-Object System.Windows.Forms.Form
    $FormListMenu.Text ='Available Documents'
    $FormListMenu.Width = 800
    $FormListMenu.Height = 400
    $FormListMenu.AutoSize = $true
    
    #TODO: Implement Duplicates color the array based on Duplicates or not.
    $nodecount = $str_FormNames_Array.Length #Size of the Array
    $listBox = New-Object System.Windows.Forms.ListBox
    $listBox.Location = New-Object System.Drawing.Point(10,40)
    $listBox.Size = New-Object System.Drawing.Size(500,300)
    $listBox.Height = 300
    $listBox.Width = 600
    $listBox.Font = New-Object System.Drawing.Font("Lucida Console",12,[System.Drawing.FontStyle]::Regular)
    $listBox.HorizontalScrollbar = $true
    
    # For each node, add an item from the array.
    for($i=0;$i -lt $nodecount; $i++){
        
        [void] $listBox.Items.Add($str_FormNames_Array[$i]) 
    }

    $FormListMenu.Controls.Add($listBox) # Add the listBox to the FormListMenu

    # Form Button
    $openButton                  = New-Object System.Windows.Forms.Button
    $openButton.Font             = New-Object System.Drawing.Font("arial",12,[System.Drawing.FontStyle]::Regular)
    $openButton.Location         = New-Object System.Drawing.Point(650),(300) # This is not aligned :(
    $openButton.Size             = New-Object System.Drawing.Size(100,30)
    $openButton.Text             = "Open"
    $openButton.Add_Click({
	    $listIndex = $listBox.SelectedIndex
	    Start-Process $urlArray[$listIndex]
    }) # Find the selected item from the listBox then open it in the default browser.

    $FormListMenu.Controls.Add($openButton) #Add the button to the Menu    
    $FormListMenu.ShowDialog() | Out-Null # Present the FormListMenu
}

Generate-Form
