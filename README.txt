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

