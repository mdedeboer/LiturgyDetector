<#
    .SYNOPSIS
    The idea of this script is to take a humanly structured order of worship and convert into an object oriented ordered order of worship.
    Why?  
        -  The benefit of having an object oriented is that we can take this and generate a predictable PPTX file or pre-create live stream via YoutTube etc...
        -  This script is an aggregate of functions that can be customized to your specific needs...that being said, it has been written specifically for my needs...
    Functions 
        - Read-MultilineInputBoxDialog - Multiline unformated text dialog box...This has primarily been repurposed from Daniel Schroeder (https://www.powershellgallery.com/packages/PowerShellFrame/0.0.0.20/Content/Public%5CRead-MultiLineInputBoxDialog.ps1)
        - Validate-Liturgy - This is the main function and tries to detect the type of liturgy for each specific row.  This is largely based on the settings variables, although some prediction is based on the standard order of worship of our church.  I've added some notes to make this somewhat editable
        - Decode-Reading - The other liturgy types I kept in the main 'validate-liturgy' function.  I found that there are too many different ways of writing the readings.  
            For example Romans 1:1-2:4 or Romans 1:1-2:4 (page 512), Romans 1:1-2:4a, Romans 1:1-31;2:1-4a etc...
            Due to the challenges of the different ways to write this, the script strips any non-digit or : characters that it encounters which are not detectable as a valid reading.  
            As a result, one issue with the below script is that Romains 1:1-2:4a will be detected as Romans 1:1:2:4 (the a will be stripped).  At this point, I find this acceptable as it isn't common and is possible to be user edited after detection
        - UserValidateDataTable - This is a grid view that is user editable in order to quickly make corrections or additions to output from the script
        - Generate-SongBoard - Edits given Powerpoint file with detected liturgy.  Only lists Reading and Singing entries.  Reading entries will be tabbed in and bolded
                   
    Settings
        - $SongBoardTemplatePPTX - This is the template used by function Generate-SongBoard.  All data in this pptx will be replaced...the format (font and font size) will be kept.  Songs will be regular type, while readings will be indented and bold.  If this doesn't work for you, you'll want to edit the function Generate-SongBoard
        - $ValidSongs - Array of entries which will be detected as "Singing"
        - $ValidReading - Hashtable of valid reading entries.  The matches on this hashtable are wildcard matches.  As a result, 'P'='Psalm' is the same as an entry 'Ps'='Psalm' and 'Prov'='Proverbs'.  In other words, ambiguous or duplicate abbreviations must be avoided.  Otherwise the script will detect more than one reading in a line and won't understand how to split the reading appropriately
        - $script:LiturgyTypes - Types of valid liturgy types to populate drop menu in function UserValidateDataTable 
        - $boolGenerateSongBoard - Control whether or not to generate a songboard
        - $boolGenerateMainPPTX - Control whether or not to generate main powerpoint file
        - $boolMainPPTXIncludeLyrics - Control whether or not to include lyrics for singings (currently useless...we don't utilize this in our church, so this feature isn't written yet)
        - $AMMainTemplatePPTX - This is the main powerpoint file.  Our church uses a display text, theme and points.  The function Generate-MainPPTX replaces the tags [Display Text], [Theme] and [Points], so these tags need to be present in the PPTX file if you want this data populated
        - $PMMainTemplatePPTX - Opportunity for two seperate templates.  Our church utilizes a seperate benediction/blessing etc...this allows this to be static in the template
        - $CrossWayV3APIKey - Set this to your api key created here: https://api.esv.org/account
    Non-Settings
        - $NewItem - Template for object of liturgy items
        - $PreviousItem - Keeps track of what was detected on the previous row...this helps predict what the next row will be. 
        - $script:UserValidatedDataTable - Script scoped variable populated when user clicks ok in the UserValidateDataTable function
       
    Issues:
        - Readings with non-digit or : or ; characters will have these characters stripped.  For example: Romains 1:1-2:4a will be detected as Romans 1:1:2:4 (the a will be stripped).  At this point, I find this acceptable as it isn't common and is possible to be user edited after detection
        - Readings at the beginning will be detected as "Reading" rather than "Display text".  This is expected, and user editable
        - SongBoard Template is dependent on a PPTX file with only 1 slide.  This would need to be heavily edited if you have more than one slide


        1 Samuel 2 detected as unknown 12: etc... see Dec 15 liturgy
        Theme cut off....see Dec 15 liturgy
#>

#******************************
#  Settings 
#******************************
$pptxpath = 'C:\Users\mdede\Documents\Church\Sound'
$SongBoardTemplatePPTX = Join-path $pptxpath -ChildPath 'SongBoardTemplate.pptx'
$AMMainTemplatePPTX = Join-Path $pptxpath -ChildPath 'MainPPTXTemplate-AM.pptx'
$PMMainTemplatePPTX = Join-Path $pptxpath -ChildPath 'MainPPTXTemplate-PM.pptx'
<#
$SongBoardTemplatePPTX = Join-path $PSScriptRoot -ChildPath 'SongBoardTemplate.pptx'
$AMMainTemplatePPTX = Join-Path $PSScriptRoot -ChildPath 'MainPPTXTemplate-AM.pptx'
$PMMainTemplatePPTX = Join-Path $PSScriptRoot -ChildPath 'MainPPTXTemplate-PM.pptx'
#>
$validSongs = @{'Psalm'='P.';'Hymn'='H.';'P'='P.';'Ps'='P.';'H'='H.';'Hy.'='H.'}
$script:LiturgyTypes = @('Reading','Singing','Display Text','Theme','Points')
$boolGenerateSongBoard = $true
$BoolGenerateMainPPTX = $true
$boolMainPPTXIncludeLyrics = $false
$CrossWayV3APIKey = '62d004c9b0ad0cb3dbd877f1121027d0b8cfb9dd'

#For purposes of the valid readings list, ps is the same as ps.  do not include both as ps is a substring of ps. and causes ambiguity.
#$validReadings = @{'Genesis'='Gen.';'Exodus'='Ex.';'Leviticus'='Lev.';'Numbers'='Num.';'Deuteronomy'='Deut.';'Joshua'='Josh.';'Judges'='Judg.';'Ruth'='Ruth';'1 Sam'='1 Sam.';'2 Sam'='2 Sam.';'1 Kings'='1 Kings';'2 Kings'='2 Kings';'1 Chron'='1 Chron.';'2 Chron'='2 Chron.';'Ezra'='Ezra';'Nehemiah'='Neh.';'Esther'='Est.';'Job'='Job';'Ps'='Ps.';'Prov.'='Prov.';'Proverbs'='Prov.';'Ecclesiastes'='Eccles.';'Eccl'='Eccles.';'Song of Solomon'='Song.';'Isaiah'='Isa.';'Jeremiah'='Jer.';'Lamentations'='Lam.';'Ezekiel'='Ezek.';'Daniel'='Dan.';'Hosea'='Hos.';'Joel'='Joel';'Amos'='Amos';'Obadiah'='Obad.';'Jonah'='Jonah';'Michah'='Mic.';'Nahum'='Nah.';'Habakkuk'='Hab.';'Zephaniah'='Zeph.';'Haggai'='Hag.';'Zechariah'='Zech.';'Malachi'='Mal.';'Matthew'='Matt.';'Mt'='Matt.';'Mark'='Mark';'Luke'='Luke';'John'='John';'Acts'='Acts';'Romans'='Rom.';'1 Cor'='1 Cor.';'2 Cor'='2 Cor.';'Galatians'='Gal.';'Ephesians'='Eph.';'Philippians'='Phil.';'Colossians'='Col.';'1 Thessalonians'='1 Thess.';'2 Thessalonians'='2 Thess.';'1 Timothy'='1 Tim.';'2 Timothy'='2 Tim.';'Titus'='Titus';'Philemon'='Philem.';'Hebrews'='Heb.';'Heb'='Heb.';'James'='James';'1 Peter'='1 Pet.';'2 Peter'='2 Pet.';'1 John'='1 John';'2 John'='2 John';'3 John'='John';'Jude'='Jude';'Revelation'='Rev.';'Belgic Confession'='B.C.';'Heidelberg Catechism'='L.D.';'Canons of Dort'='C.D.';"Lord's Day"='L.D.'; "Lord’s Day"='L.D.'; 'Lords Day'='L.D.'; "Apostle's Creed"="Apostle's Creed";'Apostles Creed'="Apostles Creed"; 'Nicene Creed'='Nicene Creed';'Athanasian Creed'='Athanasian Creed'}
$validReadings = @{'Gen'='Gen.';'Exodus'='Exodus';'Leviticus'='Lev.';'Numbers'='Num.';'Deuteronomy'='Deut.';'Joshua'='Joshua';'Judges'='Judges';'Ruth'='Ruth';'1 Sam'='1 Sam.';'2 Sam'='2 Sam.';'1 Kings'='1 Kings';'2 Kings'='2 Kings';'1 Chron'='1 Chron.';'2 Chron'='2 Chron.';'Ezra'='Ezra';'Nehemiah'='Neh.';'Esther'='Esther';'Job'='Job';'Psalm'='Psalms';'Prov'='Prov.';'Ecclesiastes'='Eccles.';'Eccl'='Eccles.';'Song of Solomon'='Song.';'Isaiah'='Isaiah';'Jeremiah'='Jer.';'Lamentations'='Lam.';'Ezekiel'='Ezek.';'Daniel'='Daniel';'Hosea'='Hosea';'Joel'='Joel';'Amos'='Amos';'Obadiah'='Obad.';'Jonah'='Jonah';'Michah'='Michah';'Nahum'='Nahum';'Habakkuk'='Hab.';'Zephaniah'='Zeph.';'Haggai'='Haggai';'Zechariah'='Zech.';'Malachi'='Mal.';'Matt'='Matt.';'Mt'='Matt.';'Mark'='Mark';'Luke'='Luke';'John'='John';'Acts'='Acts';'Romans'='Rom.';'1 Cor'='1 Cor.';'2 Cor'='2 Cor.';'Galatians'='Gal.';'Ephesians'='Eph.';'Philippians'='Phil.';'Colossians'='Col.';'1 Thessalonians'='1 Thess.';'2 Thessalonians'='2 Thess.';'1 Timothy'='1 Tim.';'2 Timothy'='2 Tim.';'Titus'='Titus';'Philemon'='Philem.';'Heb'='Heb.';'James'='James';'1 Peter'='1 Pet.';'2 Peter'='2 Pet.';'1 John'='1 John';'2 John'='2 John';'3 John'='John';'Jude'='Jude';'Rev'='Rev.';'Belgic Confession'='B.C.';'Heidelberg Catechism'='L.D.';'Canons of Dort'='C.D.';"Lord's Day"='L.D.'; "Lord’s Day"='L.D.'; 'Lords Day'='L.D.'; "Apostle's Creed"="Apostle's Creed";'Apostles Creed'="Apostles Creed"; 'Nicene Creed'='Nicene Creed';'Athanasian Creed'='Athanasian Creed'}


#******************************
# Non-Settings 
#******************************
$newItem = [pscustomobject]([ordered]@{'Type'='';'Title'='';'Value'=''})
$previousItem = $newItem.psobject.copy()

#******************************
#  Functions 
#******************************
function Read-MultiLineInputBoxDialog([string]$Message, [string]$WindowTitle, [string]$DefaultText)
{
<#
    .SYNOPSIS
    Prompts the user with a multi-line input box and returns the text they enter, or null if they cancelled the prompt.

    .DESCRIPTION
    Prompts the user with a multi-line input box and returns the text they enter, or null if they cancelled the prompt.

    .PARAMETER Message
    The message to display to the user explaining what text we are asking them to enter.

    .PARAMETER WindowTitle
    The text to display on the prompt window's title.

    .PARAMETER DefaultText
    The default text to show in the input box.

    .EXAMPLE
    $userText = Read-MultiLineInputDialog "Input some text please:" "Get User's Input"

    Shows how to create a simple prompt to get mutli-line input from a user.

    .EXAMPLE
    # Setup the default multi-line address to fill the input box with.
    $defaultAddress = @'
    John Doe
    123 St.
    Some Town, SK, Canada
    A1B 2C3
    '@

    $address = Read-MultiLineInputDialog "Please enter your full address, including name, street, city, and postal code:" "Get User's Address" $defaultAddress
    if ($address -eq $null)
    {
        Write-Error "You pressed the Cancel button on the multi-line input box."
    }

    Prompts the user for their address and stores it in a variable, pre-filling the input box with a default multi-line address.
    If the user pressed the Cancel button an error is written to the console.

    .EXAMPLE
    $inputText = Read-MultiLineInputDialog -Message "If you have a really long message you can break it apart`nover two lines with the powershell newline character:" -WindowTitle "Window Title" -DefaultText "Default text for the input box."

    Shows how to break the second parameter (Message) up onto two lines using the powershell newline character (`n).
    If you break the message up into more than two lines the extra lines will be hidden behind or show ontop of the TextBox.

    .NOTES
    Name: Show-MultiLineInputDialog
    Author: Daniel Schroeder (originally based on the code shown at http://technet.microsoft.com/en-us/library/ff730941.aspx)
    Version: 1.0
#>
    Add-Type -AssemblyName System.Drawing
    Add-Type -AssemblyName System.Windows.Forms

    # Create the Label.
    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Size(10,10)
    $label.Size = New-Object System.Drawing.Size(280,20)
    $label.AutoSize = $true
    $label.Text = $Message

    # Create the TextBox used to capture the user's text.
    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Location = New-Object System.Drawing.Size(10,40)
    $textBox.Size = New-Object System.Drawing.Size(575,200)
    $textBox.AcceptsReturn = $true
    $textBox.AcceptsTab = $false
    $textBox.Multiline = $true
    $textBox.ScrollBars = 'Both'
    $textBox.Text = $DefaultText

    # Create the OK button.
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Location = New-Object System.Drawing.Size(415,250)
    $okButton.Size = New-Object System.Drawing.Size(75,25)
    $okButton.Text = "OK"
    $okButton.Add_Click({ $form.Tag = $textBox.Text; $form.Close() })

    # Create the Cancel button.
    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Location = New-Object System.Drawing.Size(510,250)
    $cancelButton.Size = New-Object System.Drawing.Size(75,25)
    $cancelButton.Text = "Cancel"
    $cancelButton.Add_Click({ $form.Tag = $null; $form.Close() })

    # Create the form.
    $form = New-Object System.Windows.Forms.Form
    $form.Text = $WindowTitle
    $form.Size = New-Object System.Drawing.Size(610,320)
    $form.FormBorderStyle = 'FixedSingle'
    $form.StartPosition = "CenterScreen"
    $form.AutoSizeMode = 'GrowAndShrink'
    $form.Topmost = $True
    $form.AcceptButton = $okButton
    $form.CancelButton = $cancelButton
    $form.ShowInTaskbar = $true

    # Add all of the controls to the form.
    $form.Controls.Add($label)
    $form.Controls.Add($textBox)
    $form.Controls.Add($okButton)
    $form.Controls.Add($cancelButton)

    # Initialize and show the form.
    $form.Add_Shown({$form.Activate()})
    $form.ShowDialog() > $null  # Trash the text of the button that was clicked.

    # Return the text that the user entered.
    return $form.Tag
}

function Prompt-MessageBox{
    Param(
    [string[]]$buttons,
    [string]$message,
    [string]$title
    )

    [void][reflection.assembly]::Load("System.Windows.Forms, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
    [void][reflection.assembly]::Load("System.Drawing, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a")
    [void][reflection.assembly]::Load("mscorlib, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
    [void][reflection.assembly]::Load("System, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
    
    [System.Windows.Forms.Application]::EnableVisualStyles()
    $form = New-Object System.Windows.Forms.Form
    $form.Text = $title
    $formheight = 100
    $formwidth = ($buttons.count * 150)
    $form.Size = New-Object System.Drawing.Size($formwidth,$formheight)
    $form.FormBorderStyle = 'FixedSingle'
    $form.StartPosition = "CenterScreen"
    $form.AutoSizeMode = 'GrowAndShrink'
    $form.Topmost = $True
    $form.Tag = $null

    $txtmessage = New-object System.Windows.Forms.Label
    $txtmessage.location = New-object system.drawing.size(50,5)
    $txtmessagewidth = $formwidth - 25
    $txtmessage.size = New-object system.drawing.size($txtmessagewidth,11)
    $txtMessage.Text = $message
    $form.Controls.Add($txtMessage)
    
    
    $i = 0
    ForEach($buttontext in $buttons){
        $newButton = New-Object System.Windows.Forms.Button
        $xaxislocation = (($i*100))+50
        $newButton.Location = New-Object System.Drawing.Size($xaxislocation,25)
        $newButton.Size = New-Object System.Drawing.Size(75,25)
        $newButton.Text =  $buttonText
        $newButton.Add_Click({$form.tag = $this.Text; $form.Close() })
        $form.Controls.Add($newButton)
        $i++
    }
    $form.Add_Shown({$form.Activate()})
    $form.ShowDialog() > $null
    

    Return $form.tag
    
}

function UserValidateDataTable {
Param(
    [array]$DataTable
)

    [void][reflection.assembly]::Load("System.Windows.Forms, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
    [void][reflection.assembly]::Load("System.Drawing, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a")
    [void][reflection.assembly]::Load("mscorlib, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
    [void][reflection.assembly]::Load("System, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
    [System.Windows.Forms.Application]::EnableVisualStyles()
    $form = New-Object System.Windows.Forms.Form

    $OKButton = New-Object System.Windows.Forms.Button
    $InsertButton = New-Object System.Windows.Forms.Button
    $DeleteButton = New-Object System.Windows.Forms.Button
    $CopyButton = New-Object System.Windows.Forms.Button
    $PasteButton = New-Object System.Windows.Forms.Button
    $CancelButton = New-Object System.Windows.Forms.Button

    $dgv = New-Object System.Windows.Forms.DataGridView
    $InitialFormWindowState = New-Object System.Windows.Forms.FormWindowState
    $script:UserValidatedDataTable = @()

    $formEvent_Load={
    }
    $handler_button_Click={
        $script:UserValidatedDataTable = @()
        #Gather DataGrid View
        $dgvkeys = @()
        ForEach($column in $dgv.columns){
            $dgvKeys += $column.Name
        }
    
        $customHash = [ordered]@{$dgvkeys[0]='';}
        Foreach($i in 1..($dgvkeys.count-1)){
            $CustomHash.Add($dgvkeys[$i],'')
        }
        $customObject = [pscustomobject]$customHash
    
        ForEach($row in $dgv.Rows){
            $RowObject = $customObject.psobject.Copy()
            ForEach($key in $dgvKeys){
                $rowObject.$key = $row.Cells[$Key].Value
            }
        
            $Script:UserValidatedDataTable += $RowObject
        }
        $script:UserValidatedDataTable
        $form.Close() | out-null

    }

    $handler_cancelbutton_Click={
        $form.Close() | out-null
    }

    $handler_Insertbutton_Click={
        $rowIndex = $dgv.CurrentCell.RowIndex
        $dgv.Rows.Insert($rowIndex,1)
    }
    $handler_DeleteButton_Click={
        $rowIndex = $dgv.CurrentCell.RowIndex
        $dgv.Rows.RemoveAt($rowIndex)
    }
    $handler_CopyButton_Click={
        #$rowIndex = $dgv.CurrentCell.RowIndex
        $script:CopiedDGVRow = $dgv.currentRow.Clone()
        ForEach($cell in $dgv.CurrentRow.Cells){
            $script:copieddgvrow.cells[$cell.ColumnIndex].Value = $cell.Value
        }
        #$script:CopiedDGVRow = $rowIndex
        $PasteButton.Enabled = $true
    }
    $handler_Pastebutton_Click={
        $rowIndex = $dgv.CurrentCell.RowIndex
        #$dgv.Rows.InsertCopy($script:CopiedDGVRow,$rowIndex)
        $dgv.Rows.Insert($rowIndex,$script:CopiedDGVRow)
    }

    $form_StateCorrection_Load=
    {
    $form.WindowState = $InitialFormWindowState
    }

    $form.Controls.Add($OKButton)
    $form.controls.Add($InsertButton)
    $form.Controls.Add($PasteButton)
    $form.Controls.Add($CopyButton)
    $form.Controls.Add($CancelButton)
    $form.Controls.Add($DeleteButton)
    $form.Controls.Add($dgv)
    $form.Text = "Verify Data"
    $form.Name = "VerifyData"
    $form.DataBindings.DefaultDataSourceUpdateMode = [System.Windows.Forms.DataSourceUpdateMode]::OnValidation 
    $form.ClientSize = New-Object System.Drawing.Size(800,400)

    $form.BackgroundImageLayout = "None"
    $form.add_Load($formEvent_Load)

    $OKButton.TabIndex = 3
    $OKButton.Name = "Update"
    $OKButton.Size = New-Object System.Drawing.Size(75,23)
    $OKButton.UseVisualStyleBackColor = $True
    $OKButton.Text = "OK"
    $OKButton.Location = New-Object System.Drawing.Point(300,375)
    $OKButton.DataBindings.DefaultDataSourceUpdateMode = [System.Windows.Forms.DataSourceUpdateMode]::OnValidation 
    $OKButton.add_Click($handler_button_Click)

    $InsertButton.Name = 'Insert'
    $InsertButton.Size = New-Object System.Drawing.Size(75,23)
    $InsertButton.UseVisualStyleBackColor = $True
    $InsertButton.Text = 'Insert Row'
    $InsertButton.Location = New-Object System.Drawing.Point(50,5)
    $InsertButton.DataBindings.DefaultDataSourceUpdateMode = [System.Windows.Forms.DataSourceUpdateMode]::OnValidation 
    $InsertButton.add_Click($handler_Insertbutton_Click)

    $DeleteButton.Name = 'Delete'
    $DeleteButton.Size = New-Object System.Drawing.Size(75,23)
    $DeleteButton.UseVisualStyleBackColor = $True
    $DeleteButton.Text = 'Delete Row'
    $DeleteButton.Location = New-Object System.Drawing.Point(150,5)
    $DeleteButton.DataBindings.DefaultDataSourceUpdateMode = [System.Windows.Forms.DataSourceUpdateMode]::OnValidation 
    $DeleteButton.add_Click($handler_deletebutton_Click)

    $CopyButton.Name = 'Copy Row'
    $CopyButton.Size = New-Object System.Drawing.Size(75,23)
    $CopyButton.UseVisualStyleBackColor = $True
    $CopyButton.Text = 'Copy'
    $CopyButton.Location = New-Object System.Drawing.Point(250,5)
    $CopyButton.DataBindings.DefaultDataSourceUpdateMode = [System.Windows.Forms.DataSourceUpdateMode]::OnValidation 
    $CopyButton.add_Click($handler_Copybutton_Click)

    $PasteButton.Name = 'Paste Row'
    $PasteButton.Size = New-Object System.Drawing.Size(75,23)
    $PasteButton.UseVisualStyleBackColor = $True
    $PasteButton.Text = 'Paste'
    $PasteButton.Location = New-Object System.Drawing.Point(350,5)
    $PasteButton.DataBindings.DefaultDataSourceUpdateMode = [System.Windows.Forms.DataSourceUpdateMode]::OnValidation 
    $PasteButton.add_Click($handler_Pastebutton_Click)
    $PasteButton.Enabled = $false

    $CancelButton.Name = 'Cancel'
    $CancelButton.Size = New-Object System.Drawing.Size(75,23)
    $CancelButton.UseVisualStyleBackColor = $True
    $CancelButton.Text = 'Cancel'
    $CancelButton.Location = New-Object System.Drawing.Point(400,375)
    $CancelButton.DataBindings.DefaultDataSourceUpdateMode = [System.Windows.Forms.DataSourceUpdateMode]::OnValidation 
    $CancelButton.add_Click($handler_cancelbutton_Click)
    
    
    
    
    
    

    $dgv.Name = "VerifiedData"
    $dgv.DataBindings.DefaultDataSourceUpdateMode = [System.Windows.Forms.DataSourceUpdateMode]::OnValidation 
    $dgv.Location = New-Object System.Drawing.Point(5,35)
    $dgv.Size = New-Object System.Drawing.Size(785,320) 
    $dgv.font = "Calabri"
    $dgv.MultiSelect = $false

    $keys = ($Datatable | select -first 1).psobject.Properties.Name
    ForEach($key in $keys) {
    
        #Custom to create dropdown in 'Type column'
        if($key -eq 'Type'){
            $newColumn = new-object System.Windows.Forms.DataGridViewComboBoxColumn
            ForEach($liturgytype in $script:LiturgyTypes){
                $newColumn.Items.Add($Liturgytype) | out-null
            }  
        }else{
            #Standard Text box
            $newcolumn = New-object System.Windows.Forms.DataGridViewColumn
            $newColumn.CellTemplate = new-object System.Windows.Forms.DataGridViewTextBoxCell
        }
        if($key -eq 'Value'){
            $newColumn.Width = '400'
        }else{
            $newColumn.Width = '150'
        }
        $newColumn.Name = $key
        $newColumn.HeaderText = $key

        $dgv.columns.add($newColumn) | out-null
    }

    ForEach($row in $dataTable){
        $rowdata = @()
        ForEach($key in $keys){$rowdata += $row.$key}
        $dgv.Rows.Add($rowData) | out-null
    }

    $InitialFormWindowState = $form.WindowState
    $form.add_Load($form_StateCorrection_Load) | out-null
    $form.ShowDialog() | out-null

 } 

<#
Function Decode-Reading{
                Param($row,$LiturgyItem,$previousItem)
                   #Could be Confession/Creed/Bible reading
                if($validReadings.keys | where{$row -match $_ -and ($row -notmatch 'Sing' -and $row -notmatch 'opening song')}){
                    #Bible Reading
                    $validBoB = $validReadings.keys | where{$row -match "(\b$_)" -and $_ -notmatch 'John'}

                    #Find reading of John/1 John/2 John/3 John
                    $TestForJohn = [regex]::Match($row,'(?:[1-3]\s)*(John)')
                    if($TestForJohn.Success){
                        $validBobJohn = Do{$TestForJohn.value; $TestForJohn = $TestForJohn.NextMatch();}While($TestForJohn.value)
                    }
                    $validBob = $validBob + ($validReadings.keys | where{$_ -in $validBobJohn})

                    if($validBob.count -le 1){
                        $sanitizedrow = $row -replace 'verse',':'
                        #$ChapterandVerse = ($row.substring($row.indexof($validBob)+$validBob.length) -replace '\(.+\)|[^\d:-]') #-split ($validBoB) #Take all digit : or - after the matched book of the Bible, unless it is contained in parentheses ()
                        $ChapterandVerse = ($sanitizedrow -split ($validBoB) | select -last 1) -replace '\(.+\)|[^\d,:-]'
                        $LiturgyItem.Title = $validReadings.$validBoB
                        
                        $liturgyItem.Value = $ChapterandVerse | select -first 1
                        $liturgyItem.Value = $liturgyItem.Value  -replace 'verse',':' -replace ':',': ' -replace ',',', ' -replace '-',' - '
                        #Is there a previous reading...does it matter?
                        $LiturgyItem.Type = 'Reading'
                        if($liturgyItem.Value -match '\d'){
                             $LiturgyItem
                        }else{
                            write-warning "Unable to decipher reading: `"${row}`"...assuming theme/points"
                        }
                    }else{
                        #More than one reading match on this line
                        #how is it delimited? (probably ; or ,)
                        $readings = $row -replace 'verse',':' -split ';'
                        if($readings.count -eq 1){
                            $readings = $row -split ','
                        }
                        if($readings.count -eq 1){
                            write-warning "Unable to parse reading ${row}  There appears to be more than 1, but not delimited by an expected character"
                        }
                        if(($readings).count -ge 2){
                            ForEach($reading in $readings | Where{$_ -replace "\s+" -ne ""}){
                                $validBobList = $validBob
                                $validBoB = $validBobList | where{$reading -match $_}
                                $liturgyItem = $newItem.psobject.copy()

                                #Is there a previous reading...does it matter?
                                $RowType = 'Reading'

                                if(!$validBoB){
                                    write-warning "Reading: $reading is ambiguous...assuming this is part of previous reading"
                                    $LiturgyItem.Type = $RowType
                                    $liturgyItem.Title = ''
                                    
                                    $liturgyItem.Value = $reading -replace '\(.+\)|[^\d:-]' #-split ($validBoB) #Take all digit : or - after the matched book of the Bible, unless it is contained in parentheses ()
                                    if($liturgyItem.Value -match '\d'){
                                        $liturgyItem
                                    }else{
                                        write-warning "Unable to decipher reading: `"${row}`"...assuming theme/points"
                                    }
                                }else{
                                    #could add if here for Title -eq L.D. to take into account Q&A.... IE. L.D. 37 Q&A 42
                                    #Improvement for later?
                                    $ChapterandVerse = ($reading.substring($reading.indexof($validBob)+$validBob.length)  -replace '\(.+\)|[^\d:-]') #-split ($validBoB) #Take all digit : or - after the matched book of the Bible, unless it is contained in parentheses ()
                                    
                                    if($ChapterandVerse -gt 1){
                                        ForEach($item in $chapterandverse){
                                            $liturgyItem = $newItem.psobject.copy()
                                            $LiturgyItem.Type = $RowType
                                            $liturgyItem.Title = $validReadings.$validBoB
                                            $liturgyItem.Value = $item  -replace 'verse',':' -replace ':',': ' -replace ',',', ' -replace '-',' - '
                                            if($liturgyItem.Value -match '\d'){
                                                $liturgyItem
                                            }else{
                                                write-warning "Unable to decipher reading: `"${item}`"...assuming theme/points"
                                            }
                                        }
                                    }else{
                                        #Only on match on book of bible in this line
                                        $LiturgyItem.Type = $RowType
                                        $liturgyItem.Title = $validReadings.$validBoB
                                        $liturgyitem.Value = $ChapterandVerse -replace ':',': ' -replace ',',', ' -replace '-',' - '
                                        if($liturgyItem.Value -match '\d'){
                                            $liturgyItem
                                        }else{
                                            write-warning "Unable to decipher reading: `"${item}`"...assuming theme/points"
                                        }
                                    }
                                }
                                                                
                            }
                        }
                    }  
                    $previousItem = $liturgyItem                  
                }else{
                    #What are we reading?
                    write-warning "Unable to decipher reading: `"${row}`"...assuming theme/points"
                    #Call function again?
                }
}
#>

Function Decode-Reading{
Param($row)
                $validReadings = @{'Gen'='Gen.';'Ex'='Exodus';'Lev'='Lev.';'Num'='Num.';'Deut'='Deut.';'Josh'='Joshua';'Judges'='Judges';'Ruth'='Ruth';'1 Sam'='1 Sam.';'2 Sam'='2 Sam.';'1 Kings'='1 Kings';'2 Kings'='2 Kings';'1 Chron'='1 Chron.';'2 Chron'='2 Chron.';'Ezra'='Ezra';'Neh'='Neh.';'Esther'='Esther';'Job'='Job';'Psalm'='Psalms';'Prov'='Prov.';'Eccl'='Eccles.';'Song of Solomon'='Song.';'Isaiah'='Isaiah';'Jer'='Jer.';'Lam'='Lam.';'Ezek'='Ezek.';'Dan'='Daniel';'Hosea'='Hosea';'Joel'='Joel';'Amos'='Amos';'Obad'='Obad.';'Jonah'='Jonah';'Micah'='Micah';'Nahum'='Nahum';'Hab'='Hab.';'Zeph'='Zeph.';'Haggai'='Haggai';'Zech'='Zech.';'Malachi'='Mal.';'Matt'='Matt.';'Mt'='Matt.';'Mark'='Mark';'Luke'='Luke';'John'='John';'Acts'='Acts';'Rom'='Rom.';'1 Cor'='1 Cor.';'2 Cor'='2 Cor.';'Gal'='Gal.';'Eph'='Eph.';'Philip'='Phil.';'Col'='Col.';'1 Thess'='1 Thess.';'2 Thess'='2 Thess.';'1 Tim'='1 Tim.';'2 Tim'='2 Tim.';'Titus'='Titus';'Philemon'='Philem.';'Heb'='Heb.';'James'='James';'1 Pet'='1 Pet.';'2 Pet'='2 Pet.';'1 John'='1 John';'2 John'='2 John';'3 John'='John';'Jude'='Jude';'Rev'='Rev.';'BC'='B.C.';'Belgic Confession'='B.C.';'Heidelberg Catechism'='L.D.';'Canons of Dort'='C.D.';"Lord's Day"='L.D.'; "Lord’s Day"='L.D.'; 'Lords Day'='L.D.'; 'LD'='L.D.'; "Apostle's Creed"="Apostle's Creed";'Apostles Creed'="Apostles Creed"; 'Nicene Creed'='Nicene Creed';'Athanasian Creed'='Athanasian Creed'}
                $newitem = [pscustomobject]([ordered]@{'Type'='';'Title'='';'Value'=''})
                $LiturgyItem = $newItem.psobject.copy()

                   #Could be Confession/Creed/Bible reading
                if($validReadings.keys | where{$row -match $_ -and ($row -notmatch 'Sing' -and $row -notmatch 'opening song')}){
                    #Bible Reading
                    $validBoBlist = [array]($validReadings.keys | where{$row -match "(\b$_)" -and $_ -notmatch 'John'})

                    #Find reading of John/1 John/2 John/3 John
                    $TestForJohn = [regex]::Match($row,'(?:[1-3]\s)*(John)')
                    if($TestForJohn.Success){
                        $validBobJohn = Do{$TestForJohn.value; $TestForJohn = $TestForJohn.NextMatch();}While($TestForJohn.value)
                    }else{
                        $validBobJohn = $null
                    }
                    $validBoblist = $validBoblist + ($validReadings.keys | where{$_ -in $validBobJohn})

                    if($validBobList.count -eq 0){
                        write-warning "Unable to decipher reading: `"${row}`"....Does not seem to properly match a valid reading"
                    }elseif($validBoblist.count -eq 1){
                        $sanitizedrow = $row -replace 'verse',':'
                        #$ChapterandVerse = ($row.substring($row.indexof($validBob)+$validBob.length) -replace '\(.+\)|[^\d:-]') #-split ($validBoB) #Take all digit : or - after the matched book of the Bible, unless it is contained in parentheses ()
                        $ChapterandVerse = ($sanitizedrow -split ($validBoBlist | select -first 1) | select -last 1) -replace '\(.+\)|[^\d,:-]'
                        $LiturgyItem.Title = $validReadings.($validBoBlist  | select -first 1)
                        
                        $liturgyItem.Value = $ChapterandVerse | select -first 1
                        $liturgyItem.Value = $liturgyItem.Value  -replace 'verse',':' -replace ':',': ' -replace ',',', ' -replace '-',' - '
                        #Is there a previous reading...does it matter?
                        $LiturgyItem.Type = 'Reading'
                        if($liturgyItem.Value -match '\d'){
                             $LiturgyItem
                        }else{
                            write-warning "Unable to decipher reading: `"${row}`"...assuming theme/points"
                        }
                    }else{
                        #More than one reading match on this line
                        #how is it delimited? (probably ; or ,)
                        $readings = $row -replace 'verse',':' -split ';'
                        if($readings.count -eq 1){
                            $readings = $row -split ','
                        }
                        if($readings.count -eq 1){
                            write-warning "Unable to parse reading ${row}  There appears to be more than 1, but not delimited by an expected character"
                        }
                        if(($readings).count -ge 2){
                            ForEach($reading in $readings | Where{$_ -replace "\s+" -ne ""}){
                                $validBoB = $validBobList | where{$reading -match $_}
                                $liturgyItem = $newItem.psobject.copy()

                                #Is there a previous reading...does it matter?
                                $RowType = 'Reading'

                                if(!$validBoB){
                                    write-warning "Reading: $reading is ambiguous...assuming this is part of previous reading"
                                    $LiturgyItem.Type = $RowType
                                    $liturgyItem.Title = ''
                                    
                                    $liturgyItem.Value = $reading -replace '\(.+\)|[^\d:-]' #-split ($validBoB) #Take all digit : or - after the matched book of the Bible, unless it is contained in parentheses ()
                                    if($liturgyItem.Value -match '\d'){
                                        $liturgyItem
                                    }else{
                                        write-warning "Unable to decipher reading: `"${row}`"...assuming theme/points"
                                    }
                                }else{
                                    #could add if here for Title -eq L.D. to take into account Q&A.... IE. L.D. 37 Q&A 42
                                    #Improvement for later?
                                    $ChapterandVerse = ($reading.substring($reading.indexof($validBob)+$validBob.length)  -replace '\(.+\)|[^\d:-]') #-split ($validBoB) #Take all digit : or - after the matched book of the Bible, unless it is contained in parentheses ()
                                    
                                    if($ChapterandVerse -gt 1){
                                        ForEach($item in $chapterandverse){
                                            $liturgyItem = $newItem.psobject.copy()
                                            $LiturgyItem.Type = $RowType
                                            $liturgyItem.Title = $validReadings.$validBoB
                                            $liturgyItem.Value = $item  -replace 'verse',':' -replace ':',': ' -replace ',',', ' -replace '-',' - '
                                            if($liturgyItem.Value -match '\d'){
                                                $liturgyItem
                                            }else{
                                                write-warning "Unable to decipher reading: `"${item}`"...assuming theme/points"
                                            }
                                        }
                                    }else{
                                        #Only on match on book of bible in this line
                                        $LiturgyItem.Type = $RowType
                                        $liturgyItem.Title = $validReadings.$validBoB
                                        $liturgyitem.Value = $ChapterandVerse -replace ':',': ' -replace ',',', ' -replace '-',' - '
                                        if($liturgyItem.Value -match '\d'){
                                            $liturgyItem
                                        }else{
                                            write-warning "Unable to decipher reading: `"${item}`"...assuming theme/points"
                                        }
                                    }
                                }
                                                                
                            }
                        }
                    }  
                    $previousItem = $liturgyItem                  
                }else{
                    #What are we reading?
                    write-warning "Unable to decipher reading: `"${row}`"...assuming theme/points"
                    #Call function again?
                }
}

Function ValidateLiturgy{
    #Arrays to hold Data
    $liturgyUnStructuredText = Read-MultiLineInputBoxDialog -Message "Please input liturgy for individual service in order" -WindowTitle 'Liturgy'
    
    if(!$liturgyUnStructuredText){
        #Cancel clicked
        Return
    }else{
        #Data filled
        $arrLiturgy = $liturgyUnStructuredText.Split("`n")
        $arrLiturgy = $arrLiturgy | where{($_ -replace '\s+')} #Remove whitespace lines
        $arrValidatedLiturgy = @()
        $i = 0
        $previousItem = $newItem.psobject.copy()
        ForEach($row in $arrLiturgy){
            $matchedrow = $false
            
            $LiturgyItem = $newItem.psobject.copy()
            $row = $row -replace '\s+',' ' 
        
            $rowWordSplit = $row -split '\W'
            if($i -eq '0' -and ($row -match 'Display' -or ($row -match 'Text')) -or ($row -match 'Display')){
                

                #If Display Text Found
                $liturgyItem.Type = 'Display Text'

                #Check if any word matches book of Bible
                
                $validBOB = $validReadings.keys | where{$row -match $_}

                if($validBoB){
                    #Valid Book of Bible found.  Set abbreviated title and value 
                    $LiturgyItem.Title = $validReadings.$validBoB
                    
                    $LiturgyItem.Value = ($row -split ($validBoB) | select -last 1) -replace 'verse',':' -replace '\(.+\)|[^\d:-]' #-split ($validBoB) #Take all digit : or - after the matched book of the Bible, unless it is contained in parentheses ()
                    $matchedrow = $true
                }else{
                    Write-warning 'Invalid Display text found'
                    #Call function again?
                }
                $LiturgyItem
            }
            

            if(($matchedrow -eq $false) -and ((($previousItem.Type -eq 'Reading') -and $row -match "^\s+" -and ($rowWordSplit | Where{$_ -in ($validReadings.keys)})) -or $row -match 'Reading' -or $row -match 'Scripture' -or $row -match 'Sermon' -or $row -match 'Text' -or $row -match 'Read' -or $row -match 'Cat' -or ($rowWordSplit | Where{$_ -in ($validReadings.keys | where{$validReadings.$_ -ne 'Ps.' -and ($_ -notmatch 'sing')})}))){ #($rowWordSplit | Where{$_ -in $validReadings.keys}) this may be problematic if the theme contains the word 'Psalm' or songs?
            #Reading
                if($row -notmatch '(?<!.)((\*|\s|Psalm)+)'){ #If the word Psalm appears and it has no preceding text (except white space or *) then it is probably singing not reading
                    Decode-Reading -row $row -LiturgyItem $LiturgyItem -previousItem $previousItem | ForEach{
                        $matchedrow = $true
                        $_                   
                    }
                }
                
            }
            if(($matchedrow -eq $false) -and (($rowWordSplit | where{$_ -in $validSongs.keys}))){
                if($row.substring($row.indexOf(($rowWordSplit | where{$_ -in $validSongs.keys}))+(($rowWordSplit | where{$_ -in $validSongs.keys}).length)) -match "\d"){  #A song in ValidSongs that is followed by numbers of some sort.  This filters out the possibility that the word "Psalm" in a theme is interpreted as a singing
                    #Singing
                    #Could this be a reading of Psalm X ?  Could this be a theme that has the word "Psalm" in it? How to tell the difference?
                    $LiturgyItem.Type = 'Singing'
                    $SongBook = ($rowWordSplit | where{$_ -in $validSongs.keys})
                    $LiturgyItem.Title = $validSongs.$SongBook
                    $LiturgyItem.Value = ($row -split $SongBook | select -last 1) -replace 'and',',' -replace '\s+' -replace 'stanza',':' -replace '\.' -replace ':',': ' -replace ',',', ' -replace '-',' - '
                    $LiturgyItem
                    $matchedrow = $true
                }
            }

            
            if(!($matchedrow)){
                 #this could be or another text/reading for the sermon,  but should be caught by reading, unless it is a psalm that is on a line that does not contain another reading or the key words used for readings
                #Could be singing of a Psalm, but this would be caught by the validsongs
                #Could be Theme/Points/Benediction/Ten words of the Covenant/Blessing/Benediction/Votum/Salutation/Prayer/Offering/Announcements
                
                #The next line after a detected theme is likely points....same with the next line after a detected point...unless it has other keywords
                #Look back detection
                if(!$matchedrow -and $previousItem.Type -match 'Theme' -or $previousItem.Type -match 'Points'){
                    $DetectedType = 'Points'
                    $LiturgyItem.Type = $DetectedType
                    $LiturgyItem.Value = $row -replace "^(\s+)+"
                    $liturgyItem.Title = ''
                    $LiturgyItem  
                    $matchedrow = $true
                }

                #How to tell the difference between an unlabeled theme and points vs other parts of worship service?
                $DetectedType = $null
                if(!$matchedrow -and ($row -match 'Theme' -or $row -match 'Sermon' -or $previousItem.Type -eq 'Reading') -and $PreviousItem.Type -ne 'Theme'){
                    $DetectedType = 'Theme'
                    $LiturgyItem.Type = $DetectedType
                    $LiturgyItem.Value = $row -replace "^(Theme|Sermon|:|\s+)+"
                    $liturgyItem.Title = ''
                    $LiturgyItem  
                    $matchedrow = $true   
                }

                
                #If look ahead next row is not reading or singing or prayer or benedition or votum or salutation or offering or collection or announcement or law or covenant
                #If we get this far, we haven't had a theme yet (would have matched earlier)
                
                $nextrow = $arrLiturgy[$i+1]
                $nextrowwordsplit = $nextrow -split "\W"
                $keywordstomatch = @('Reading','Sermon','Text','Pray','Read','Sing','Benediction','votum','salutation','offering','collection','offeratory','announcement','law','covenant','announcement','cat')
                if(!$matchedrow -and !($keywordstomatch | where{$row -match $_}) -and !($nextrowWordSplit | where{$_ -in $validSongs.keys}) -and ($previousItem.Type -ne 'Theme')){
                    $DetectedType = 'Theme'
                    $LiturgyItem.Type = $DetectedType
                    $LiturgyItem.Value = $row -replace "^(Theme|Sermon|:|\s+)+"
                    $liturgyItem.Title = ''
                    $LiturgyItem 
                    $matchedrow = $true
                }
            }
            $i++
            $PreviousItem = $LiturgyItem

        }

 
    }
}

function Generate-SongBoard{
    Param(
        [string]$TemplatePPTX,
        [array]$validatedLiturgy,
        [string]$OutputDir = $PSScriptRoot
    )

    #Open PPTX template
    if(Test-Path $templatepptx){
        $app = New-Object -ComObject powerpoint.application
        $pres = $app.Presentations.open($templatepptx)
        $app.visible = "msoTrue"
        $slides = $pres.slides
        start-sleep 2
        if($slides.count -gt 1){
            write-warning "Unable to identify appropriate format for songboards"
        }else{
            ForEach($slide in $slides){
                if($slide.shapes.count -gt 1){
                    write-warning "Unable to identify appropriate format for songboards"
                }else{
                    Foreach($shape in $slide.shapes){
                        #This shape is where we want to put our text
                        $i=0
                        $shape.TextFrame.TextRange.Text = ''
                        ForEach($Item in $validatedLiturgy){
                            
                            if($i -gt 0 -and $item.Type -in @('Reading','Singing') -and $item.Title){$range = $shape.textframe.textrange.insertafter("`n")}
                            if($Item.Type -eq "Reading"){
                                if(!($Item.Title)){
                                    
                                    $range = $shape.textframe.textrange.insertafter(", " + $Item.Value)
                                }else{
                                    #Bold and indented
                                    $range = $shape.textframe.textrange.InsertAfter("`t" + $Item.Title + " " + $Item.Value)
                                    $range.font.bold = $true
                                }
                                $i++
                            }
                            if($Item.Type -eq "Singing"){
                                #Not bold...not indented
                                $range = $shape.textframe.textrange.insertafter($Item.Title + " " + $Item.Value)
                                $range.font.bold = $false
                                $i++
                            }

                            
                        }
                    }
                }
            }
        }
        $app.PresentationClose
        
        
    }else{
        Write-Error "Unable to find $templatepptx"
    }


}

Function Generate-MainPPTX{
    Param($IncludeLyrics = $false,$TemplateFile,$CrosswayV3APIkey)

    #Attempt to open pptx
     if(Test-Path $Templatefile){
        $app = New-Object -ComObject powerpoint.application
        $pres = $app.Presentations.open($templatefile)
        $app.visible = "msoTrue"
        $slides = $pres.slides
        start-sleep 2
    }

    if($slides){  #If we found powerpoint presentation and was able to find slides
        ForEach($liturgyItem in $script:UserValidatedDataTable){
            if($liturgyItem.Type -eq 'Display Text'){
                #Look up display text on an online service
                $TextToLookUp = ($liturgyItem.Title + ' ' + $liturgyItem.Value)
                $queryParams = @{
                    'q'=$TextToLookUp
                    'include-headings'=$False
                    'include-footnotes'=$False
                    'include-verse-numbers'=$False
                    'include-short-copyright'=$false
                    'include-passage-references'=$False
                }

                $url = 'https://api.esv.org/v3/passage/text/'
                
                 $headers = @{ 
                    "Accept"="application/json"
                    "Authorization" = "$CrosswayV3APIkey"
                }
                $jsonResponse = Invoke-RestMethod -Method Get -Uri $url -Headers $headers -Body $queryParams -UseBasicParsing
                $DisplayText = ([string]($jsonResponse.passages) -Replace "(\r|\n|\t|\s{2,})+", ' ').TrimStart(' ').TrimEnd(' ')
                
                if($DisplayText){
                    ForEach($slide in $slides){
                        ForEach($shape in $slide.Shapes){
                            $shape.TextFrame.TextRange.Replace('[Display Text]',$DisplayText) | out-null
                            
                            <# When replacing text, a new paragraph is created below the inserted text...we need to delete this.
                            $replacedText = $shape.TextFrame.TextRange.Replace('[Display Text]','[Display Text2][SomethingToDelete]')
                            $replacedText = $shape.TextFrame.TextRange.Replace('[Display Text2]',($DisplayText))
                        
                        
                            if($replacedText){ #When Powerpoint replaces text, it adds a new paragraph...we need to delete it
                                $paragraphtoDelete = $shape.TextFrame.textrange.Paragraphs() | where{$_.Text -match 'SomethingToDelete'}
                                $paragraphtoDelete | %{$_.Delete()}          
                            }
                            #>

                            $shape.TextFrame.TextRange.Replace('[Display Text Reference]',$TextToLookup) | out-null
                           

                        }
                    }

                }else{
                    write-warning "Failed to lookup display text"
                }
            }

            if($liturgyItem.Type -eq 'Reading'){
                #Nothing to do here
            }

            if($liturgyItem.Type -eq 'Singing' -and $includeLyrics){
                #What do we do here?
            }

            if($liturgyItem.Type -eq 'Theme'){
                #Replace [THEME] tag
                $theme = $liturgyItem.Value
                ForEach($slide in $slides){
                    ForEach($shape in $slide.Shapes){
                        $replacedText = $null
                        $fontsize = $shape.TextFrame.TextRange.Font.Size
                        $replacedText = $shape.TextFrame.TextRange.Replace('[Theme]',$theme)
                        if($replacedText){
                            $shape.textframe.TextRange.Text = $shape.TextFrame.TextRange.Text -replace ("(\r|\n)$", '')
                        }
                        
                        

                    }
                }
            }

            if($liturgyItem.Type -eq 'Points' -and !$FoundPoints){
                #Need to get all points
                [array]$points = $script:UserValidatedDataTable | where{$_.Type -eq 'Points'} | select -expand value
                $strPoints = $points -join "`n"

                #Replace [Points] tag
                ForEach($slide in $slides){
                    ForEach($shape in $slide.Shapes){
                        
                        $shape.TextFrame.TextRange.Replace('[Points]',$strpoints) | out-null
                        #This returns an error....not sure why....it still works!
                    }
                }

                $foundPoints = $true
            }
        }
    }
    else{
        write-warning "This PPTX file doesn't seem to have any content"
    }
}

#******************************
#  Main 
#******************************
$validatedLiturgy = ValidateLiturgy
UserValidateDataTable -DataTable $validatedLiturgy

if($script:UserValidatedDataTable.count -ge 1 -and $boolGenerateSongBoard){
    Generate-SongBoard -TemplatePPTX $SongBoardTemplatePPTX -validatedLiturgy $script:UserValidatedDataTable
}

if($script:UserValidatedDataTable.count -ge 1 -and $boolGenerateMainPPTX){
    $PromptAMPM =  Prompt-Messagebox -buttons @('AM','PM') -message 'Is this an AM or PM service?' -Title 'AM or PM'
    $TemplateFile = Switch ($PromptAMPM){
        'AM' {$AMMainTemplatePPTX}
        'PM' {$PMMainTemplatePPTX}
    }
    Generate-MainPPTX -IncludeLyrics $boolMainPPTXIncludeLyrics -TemplateFile $TemplateFile -CrosswayV3APIkey $CrossWayV3APIKey
}
