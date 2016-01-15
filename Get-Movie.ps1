function Get-Movie {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true, Position=0, ParameterSetName='Title')]
        [String]$Title,
        [Parameter(Mandatory=$true, Position=0, ParameterSetName='imdbid')]
        [String]$imdbID,
        [String]$Plot = "Short",
        [String]$Type = "Movie",
        [Switch]$Tomatoes        
    )
    if($Title) {
        $request = "t=$Title&type=$Type&plot=$Plot&r=json&tomatoes=$Tomatoes"
        Write-Verbose $request
    } elseif ($imdbID) {
        $request = "i=$imdbID&type=$Type&plot=$Plot&r=json&tomatoes=$Tomatoes"
        Write-Verbose $request
    }

    $web_result = Invoke-WebRequest -Uri "http://www.omdbapi.com/?$request"
    
    ConvertFrom-Json -InputObject $web_result.Content
}

function Search-Movie {
    param (
        [Parameter(Mandatory=$true)]
        [String]$Title
    )

    $request = "s=$Title&type=movie&r=json"
    Write-Verbose $request

    $web_result = Invoke-WebRequest -Uri "http://www.omdbapi.com/?$request"
    
    ConvertFrom-Json -InputObject $web_result.Content
    

}

function Generate-NMPSInventory {
    param (
        [Parameter(Mandatory=$true)]
        [String]$src,
        [Int]$limit = -1,
        [Switch]$reprocess = $false
    )
    Write-Verbose $src
    $source_csv = Import-Csv -Path $src

    $output = @()

    #for testing
    $counter = 0
    foreach($row in $source_csv) {
        if($counter -eq $limit) {
            Write-Warning ("Limit reached - halting processing")
            break
        }
        if($counter % 50 -eq 0) {
            #Write status message every 50 records
            Write-Warning ("Processing Record $counter")
        }
        $title = $row.Title
        $proc = 'true'

        if(-Not $reprocess -and $row.PROCESSED -eq 'true')
        {
            #Assume we have already processed this title and skip the lookup
            $output += $row
            continue
            
            #$metadata = $row
        }
        elseif ($row.imdbID) {
            #Check for a manually entered imdbID (for manual fix of bad titles)
            Write-Verbose ("Performing imdbID search for $title")
            $metadata = Get-Movie -imdbID $row.imdbID
        }
        else {
            #Else this is a brand new entered/unfixed movie Title
            #Fetch movie data from omdbapi.com by Title search
            $metadata = Get-Movie -Title $title
        }

        if($metadata.Error) {
            $proc = 'Lookup Error'
            Write-Warning ("Error looking up $title")
        }
        elseif($row.Title -ne $metadata.Title) {
            $warn_title = $metadata.Title
            Write-Verbose ("Title mismatch $Title (original) >> $warn_title (lookup)") 
            $title = $metadata.Title
        }
     
        $properties = [Ordered] @{
            "ID" = $row.ID;
            "imdbID" = $metadata.imdbID;
            "IMDB Link" = '';
            "TITLE" = $title;
            "RATED" = $metadata.Rated;
            "EXPIRATION" = '="' + $row.EXPIRATION + '"'; #Preserve the NMPS Expiration code
            "IMDB RATING" = $metadata.imdbRating;
            "GENRE" = $metadata.Genre;
            "PLOT" = $metadata.Plot;
            "PROCESSED" = $proc
        }
        $output += New-Object -TypeName psobject -Property $properties
        $counter++
    }
    #Push it to the pipeline
    $output
}

#Generate-NMPSInventory -Input 'C:\Dev\Powershell\NMPSDecember2015Inventory.csv' -Verbose | Export-Csv -Path 'C:\Dev\Powershell\test_export.csv' -Force -NoTypeInformation
