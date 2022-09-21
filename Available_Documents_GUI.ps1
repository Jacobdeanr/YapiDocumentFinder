#---------------------------------------------------------------------
# GUI stuff
#---------------------------------------------------------------------
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 

Add-Type -AssemblyName PresentationCore,PresentationFramework
Add-Type -assembly System.Windows.Forms

function Generate-Form{
clear

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
$textbox.MaxLength = "38"
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
function ReadForms {
clear
#---------------------------------------------------------------------
# Combine the GUID with the website URL to grab a list of forms
#---------------------------------------------------------------------
    $path = "https://yapi.me/f/p/list.php?guid="
    
    $GUID = $textbox.Text
    #$GUID = "{1DBFAF39-2943-4607-8353-189749D20A81}"
    #$GUID = Read-Host "Enter Practice GUID"
    $GUID = $GUID -replace '\s','' #cleanup any invalid characters when copy/pasting from the YMC
    
    $domain = $path+$GUID
    
    if(($xml = Invoke-RestMethod -uri $domain) -like "*G-Invalid GUID*") {
		[System.Windows.MessageBox]::Show("Nothing Found, make sure the GUID was entered correctly.",'Error','OK','Error')

    ##Write-Host "Nothing Found, make sure the GUID was entered correctly." -ForegroundColor Red
    }
    else{
    #$duplicates = $null
    $str_FormNames_Array = @() #Empty array used to build the list of forms.
    $uniqueNamesArray = @() #Empty array for
    $urlArray = @()


#---------------------------------------------------------------------
# Build the resulting domain using Variables
#---------------------------------------------------------------------
    $first = "https://yapi.me/f/"
    $practice = "&guid="
    $tag = "&tag=[[TAGID]]"
    
    
    
#---------------------------------------------------------------------
# for each file that has an @oxygen attribute loop over it
#---------------------------------------------------------------------
    $a = 0
    
    $lists = $xml.SelectNodes("//list/form[@oxygen='1']") | ForEach-Object { 
    $formName = $_.name
    $path = $_.url
    $id = $_.id
    
    $formNameSeparator = $formName + ':'
    $form =  $first + "$path" + $practice + $GUID + $tag
    
    #Write-Host Array Entry $a is $str_FormNames_Array[$a]

    # If "a", once its been divided by 2 has a remainder, switch text color.
    if ($a % 2  -eq 0) {
            #Write-Host $formNameSeparator $id `n$form`n -ForegroundColor Gray
        }
        else{
            #Write-Host $formNameSeparator $id `n$form`n -ForegroundColor DarkGreen
        }

    $urlArray += $form;
    #Write-Host URL Entry $a is $urlArray[$a]

    $str_FormNames_Array += $formNameSeparator = $formName + ': ' + $id; #Add each Form Name into the Array
    ##Write-Host New Entry $a is $str_FormNames_Array[$a];
    
    $a++ #increment to change color
   
    #Write-Host Array Length is $str_FormNames_Array.Length
    #Write-Host URL Array Length is $urlArray.Length

    }
    
#---------------------------------------------------------------------
# Find Duplicate Form Names
#---------------------------------------------------------------------
    $i = 0
    
    $uniqueNamesArray = $str_FormNames_Array | select -unique #unique values in the array
    
    #Write-Host "Duplicates:"
    
    Try{
        $duplicates = Compare-Object -ReferenceObject $uniqueNamesArray -DifferenceObject $str_FormNames_Array | ForEach-Object{
            $dupeNames = $_.InputObject
            #Write-Host $dupeNames -ForegroundColor Red -BackgroundColor Black
            }
        if($dupeNames -eq $null){
            #Write-Host `n"No Duplicates :)"`n -ForegroundColor Green
            
        }
    }
    
    Catch{
     #Write-Host "An Error Occured"`n -ForegroundColor Red
    }
    
    Finally{
    
        }
        

#---------------------------------------------------------------------
# Final instructions before closing.
#---------------------------------------------------------------------
    #Write-Host `n"Highlight and right click to copy the text to the clipboard"`n -ForegroundColor Yellow -BackgroundColor Black;
    Generate-Links($str_FormNames_Array)
    }
}
function Generate-Links($str_FormNames_Array){

#Basic Container
$FormListMenu = New-Object System.Windows.Forms.Form
$FormListMenu.Text ='Available Documents'
$FormListMenu.Width = 800
$FormListMenu.Height = 400
$FormListMenu.AutoSize = $true
#

#------------------------------------------------------------------------------
#Populate the FlowLayoutPanel
#------------------------------------------------------------------------------
$flowlayoutpanel = New-Object System.Windows.Forms.FlowLayoutPanel
$flowlayoutpanel.AutoScroll=$true
$flowlayoutpanel.Width = 300
$flowlayoutpanel.Height= 250
$flowlayoutpanel.FlowDirection = "TopDown"
$flowlayoutpanel.WrapContents = $false;
$flowlayoutpanel.AutoSize = $fase
$flowlayoutpanel.Anchor  = 'top'
$flowlayoutpanel.Padding = 0
$flowlayoutpanel.BackColor='white'

#$nodecount = 1
$nodecount = $str_FormNames_Array.Length
#Write-Host Unique Names Array length is $nodecount

##Flow Layout
#for($i=0;$i -lt $nodecount; $i++){
#    
#    #Write-Host Normal Array entry [$i] is $str_FormNames_Array[$i]
#
#    $label=New-Object System.Windows.Forms.Label
#    $label.Text= $str_FormNames_Array[$i]
#    $label.Height = 30
#    $label.Width = "500"
#
#    ##Write-Host Creating Label $uniqueNamesArray[$i]
#
#    $flowlayoutpanel.Controls.Add($label)  
#}

#table layout
$listBox = New-Object System.Windows.Forms.ListBox
$listBox.Location = New-Object System.Drawing.Point(10,40)
$listBox.Size = New-Object System.Drawing.Size(500,300)
$listBox.Height = 300
$listBox.Width = 600
$listBox.Font = New-Object System.Drawing.Font("Lucida Console",12,[System.Drawing.FontStyle]::Regular)
$listBox.HorizontalScrollbar = $true

#for each node, add an item from the array.
for($i=0;$i -lt $nodecount; $i++){
    
    [void] $listBox.Items.Add($str_FormNames_Array[$i])

    #$listBox.Controls.Add($label)  
}

$FormListMenu.Controls.Add($listBox)

##----------------------------------------------------------
##Form Button
##----------------------------------------------------------
$openButton                  = New-Object System.Windows.Forms.Button
$openButton.Font             = New-Object System.Drawing.Font("arial",12,[System.Drawing.FontStyle]::Regular)
$openButton.Location         = New-Object System.Drawing.Point(650),(300)
$openButton.Size             = New-Object System.Drawing.Size(100,30)
$openButton.Text             = "Open"
$openButton.Add_Click({

$listIndex = $listBox.SelectedIndex

Start-Process $urlArray[$listIndex]

})

$FormListMenu.Controls.Add($openButton)
##----------------------------------------------------------

#Clean up arrays...
$duplicates = ""
[String[]]$str_FormNames_Array = @()#empty the array
[String[]]$uniqueNamesArray = @()#empty the array

$FormListMenu.ShowDialog() | Out-Null
}
Generate-Form