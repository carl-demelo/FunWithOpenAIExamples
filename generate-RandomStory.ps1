<#
.SYNOPSIS
    Silly program that I wrote to illustrate the power of OpenAI in PowerShell.
    It writes short stories from a random generated input seed sentence.

    Many thanks to Doug Finke for the awesome PowerShellAI module!!!!
    https://github.com/dfinke/PowerShellAI
#>

param(
    [string]$APIKey
    , [string]$DestinationPath
    , [int]$MaxStoryLength = 1024
)
# We will use write-information versus write-host since we do not want to kill any puppies
# We should set our inormationPreference at the beginning of the script
# http://jeffwouters.nl/index.php/2014/01/write-host-the-gremlin-of-powershell/#:~:text=%E2%80%9CEvery%20time%20you%20use%20Write,nutshell%3A%20It%20interferes%20with%20automation.
$InformationPreference = 'continue'

Import-Module PowerShellAI

function set-MyOpenAIKey {
    <#
    .SYNOPSIS
        Function sets the OpenAI API key for the current session
    #>
    param(
        [string]$APIKey
    )
    $SecureAPIKey = $(ConvertTo-SecureString -String $APIKey -AsPlainText -Force)
    Set-OpenAIKey -Key $SecureAPIKey
}

function get-randomword {
    <#
    .SYNOPSIS
        Function selects a random word from an array of words
        If $WordType is passed in a word is selected from https://www.randomlists.com
    #>
   
    param(
        [Parameter(ValueFromPipeline = $true, ParameterSetName = 'ByWordType')]
        [string]$WordType
        , [Parameter(ParameterSetName = 'ByWordArray')] [string[]]$words
    )
    $random = New-Object System.Random

    # Test the parameter set name that is used
    if ($PSCmdlet.ParameterSetName -eq 'ByWordType') {
        $url = "https://www.randomlists.com/data/$($WordType).json"
        $words = $(Invoke-RestMethod -Uri $url).data
    }
    # test that we have an array of words
    if (-not $words) {
        throw 'No words were passed in'
    }
    
    $arraymax = $words.count - 1
    $word = $words[$random.Next(0, $arraymax)]
    return $word
}

function get-captilized {
    <#
    .SYNOPSIS
        Inline function capitilizes a word passed in
    #>
    
    param(
        [Parameter(ValueFromPipeline = $true)]
        [string]$word
    )
    return $($word.Substring(0, 1).ToUpper() + $word.Substring(1))
}

function get-superpowers {
    <#
    .SYNOPSIS
        Function selects a random superpower from https://www.randomlists.com
    #>

    $url = 'https://www.randomlists.com/data/superpowers.json'
    $words = $(Invoke-RestMethod -Uri $url).RandL.items
    $arraymax = $words.count - 1
    $random = New-Object System.Random
    $word = $words[$random.Next(0, $arraymax)]
    return $word
}

function Open-HTMLStringInBrowser {
    <#
    .SYNOPSIS
        Function opens a HTML string in the user's default browser and saves the HTML to a file
    #>
    param(
        [string]$HTML
        , [string]$DestinationPath
    )
    # Generate a base filename
    $basefilename = "Story_$(Get-Date -Format 'yyyyMMdd_HHmmss')"

    # Create a HTML file in the $DESTINATIONPATH
    $tempFile = "$($DestinationPath)$($basefilename).html"

    # Write the HTML string to the file
    Set-Content -Path $tempFile -Value $HTML

    # Launch the file in the user's default browser
    Start-Process $tempFile
}
Function Convert-PngToJpg {
    <#
    .SYNOPSIS
        Function converts a PNG file to a JPG file
    #>

    Param (
        [Parameter(Mandatory = $true)]
        [string]$SourceFile,
        [Parameter(Mandatory = $true)]
        [string]$DestinationFile
    )
    Add-Type -AssemblyName system.drawing
    $imageFormat = 'System.Drawing.Imaging.ImageFormat' -as [type]
    $image = [drawing.image]::FromFile($sourceFile)
    $image.Save($DestinationFile, $imageFormat::jpeg)
    $image.Dispose()

    # Test that the jpg file exists before removing the $SourceFile file
    if (-not (Test-Path -Path $DestinationFile)) {
        throw "Convert-PngToJpg: The file $($DestinationFile) does not exist"
    }
    Remove-Item $SourceFile
}

function get-AIJpeg {
    <#
    .SYNOPSIS
        Function generates a an image based upon the passed in description and saves it as a JPG file
    #>
    param(
        [string]$ImageDescription
        , [string]$DestinationPath
    )
    # Create a unique basefilename
    $basefilename = "AIImage_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
    $pngFile = "$($DestinationPath)$($basefilename).png"
    $jpgFile = "$($DestinationPath)$($basefilename).jpg"

    # Generate the image and save it as a PNG file
    Get-DalleImage -Description $ImageDescription -Size 1024 -DestinationPath $pngFile | Out-Null

    # test that the png image file exists
    if (-not (Test-Path -Path $pngFile)) {
        throw "get-AIJpeg: The file $($pngFile) does not exist"
    }
    Convert-PngToJpg -SourceFile $pngFile -DestinationFile $jpgFile
    return $jpgFile
}

# Set the API key for OpenAI
# This is a free API key, but it is limited to 5,000 calls per month
set-MyOpenAIKey -APIKey $APIKey

# Generate a random topic
$Topic = get-randomword -WordType 'topics'

# Create a crazy user name
$Name = get-randomword -WordType 'adjectives' | get-captilized
$Name += get-randomword -WordType 'nouns' | get-captilized

# Our story will be about a person with a superpower, let's generate one!
$Powers = get-superpowers

# Generate a random sentence about the person

# Generate a random modifier and verb
$Modifier = get-randomword -WordType 'adverbs'
$Verb = get-randomword -WordType 'verbs'

# Generate a random compound word. A word made up of two words
$CompundWord = get-randomword -WordType 'compound-words'

# Generate a random story style and media type (romantic comedy, action adventure, etc.)
$StoryStyle = get-randomword -words @('Romantic', 'Comedy', 'Action', 'Adventure', 'Fantasy', 'Thriller', 'Horror', 'Mysterious', 'Sci-Fi', 'Dramatic', 'Sphagetti Western', 'Uplifting', 'Sad', 'scary')

# Generate a random story media type (song, play, movie, novel, etc.)
$StoryMedia = get-randomword -words @('Limerick', 'Poem', 'Movie script dialog', 'Song', 'Play', 'Movie', 'Novel', 'TV Series', 'Reality TV Show')

# Generate a random influence.  An  existing story, movie, play, etc.
$Influence = "A existing $($StoryStyle) $($StoryMedia)" | ai 'select a title from a list online'

if ($null -eq $Influence) {
    $Influence = "The new and upcoming $($StoryStyle) $($StoryMedia)"
}

# Generate something to manipulate resembling the Compound Word
$Thing = ai "randomly select the name of an object like $($CompundWord)"

# About me
$AboutMe = "I like to $($Modifier) $($Verb) my $($Thing)"

# Tell OpenAi to make construct an action using the object
$AIAction = $AboutMe | ai 'write a sentence that describes an action you would like to do with an object'

# Set the maximum length of the story
$instruction = "Write a $($StoryMedia) using the topic $($Topic) with maximum $($MaxStoryLength) characters.  " 
$intro = "My name is $($Name) and $($AboutMe) $($AIAction). I have the superpower of $($Powers.name), $($Powers.detail)"
$jpgFile = get-AIJpeg -ImageDescription $intro -DestinationPath $DestinationPath
$Opener = "Here is a short $($StoryMedia) (maximum $($MaxStoryLength) characters) about me using the topic `"$($Topic)`" in the style of the $($StoryMedia) `"$($Influence)`".  I hope you enjoy it."

# Create a Here string with the instructions for OpenAI to generate a story
$HereString = @"
Hello, $($intro)

$($Opener)
"@

# Set the story to the Here string
#$Story = $HereString

# Generate the story
$StoryBody = $HereString | ai $instruction -max_tokens $MaxStoryLength

# Replace carriage return \ line feeds in $StoryBody with the html tag <br>
$StoryBody = $StoryBody -replace "`r`n", "<br>" -replace "`n", "<br>"
<#
 Create a powershell here string containing HTML to display the story and image. Put the $name in the title and the $intro as a subtitle.  add a background color to the title and subtitle.  put the image in the top left corner of the page.   Put a border around the image. Put the story in a scrollable div so that it can be scrolled if it is too long.  Put the text "Carl Demelo - GitHub:https://github.com/carl-demelo/FunWithOpenAIExamples - https://www.linkedin.com/in/carl-demelo" in italics in the footer.  Place a horizontal bar over the footer.
#>
$FinalStory = @"
<!DOCTYPE html>
<html>
<head>
    <title>$($Name)</title>
    <style>
        body {
            background-color: #F7F9F0;
            color: #000000;
        }
        h1 {
            background-color: #F7F9F0;
            color: #000000;
        }
        h2 {
            background-color: #ffffff;
            color: #000000;
            border: 1px solid #000000;
        }
        h3 {
            background-color: #ffffff;
            color: #000000;
            border: 1px solid #000000;
        }
        .story {
            background-color: #BCBCBC;
            color: #000000;
            border: 1px solid #000000;
        }
        .image {
            border: 3px solid #000000;
        }
    </style>
</head>
<body>
    <h1><img class="image" src="$($jpgFile)" alt="$($Name)" width="200" height="200"> $($Name)</h1>
    <br>
    <p>$($intro)</p>
    <p>$($Opener)</p>
    <br>
    <div class="story">
        <p>$($StoryBody)</p>
    </div>
</body>
<footer>
    <hr>
    <p><i>Carl Demelo - GitHub: https://github.com/carl-demelo/FunWithOpenAIExamples - https://www.linkedin.com/in/carl-demelo</i></p>
</footer>
</html>
"@

# Write the story to a file and display it in the default browser
Open-HTMLStringInBrowser -HTML $FinalStory -DestinationPath $DestinationPath


