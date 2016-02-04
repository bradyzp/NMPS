<#

#>
function Get-Movie {
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
        $request = "?t=$Title&type=$Type&plot=$Plot&r=json&tomatoes=$Tomatoes"
        Write-Verbose $request
    } elseif ($imdbID) {
        $request = "?i=$imdbID&type=$Type&plot=$Plot&r=json&tomatoes=$Tomatoes"
        Write-Verbose $request
    }

    $uri = "http://www.omdbapi.com"
    $web_result = Invoke-WebRequest -Uri "$uri/$request"
    
    #Encoutered an error performing a title lookup, try to search with API
    if($web_result.Error -and $Title) {
        $request = "?s=$Title&type=movie&r=json"
        Write-Verbose $request
        $web_result = Invoke-WebRequest -Uri "$uri/$request"
    }
    ConvertFrom-Json -InputObject $web_result.Content
}


<#

#>
function Generate-NMPSInventory {
    param (
        [Parameter(Mandatory=$true)]
        [String]$src,
        [Int]$limit = -1,
        [Switch]$Reprocess = $false
    )
    Write-Verbose $src
    $source_csv = Import-Csv -Path $src

    $output = @()

    #for testing
    $counter = 0
    foreach($row in $source_csv) {
        #####
        #TESTING BLOCK
        #####
        if($counter -eq $limit) {
            Write-Warning ("Limit reached - halting processing")
            break
        }
        if($counter % 50 -eq 0) {
            #Write status message every 50 records
            Write-Warning ("Processing Record $counter")
        }

        ######
        #Actual Logic
        ######
        $Title     = $row.Title
        $Processed = 'true'

        #Check if this row has already been processed (and reprocess flag isn't set)
        if(-Not $Reprocess -and ($row.PROCESSED -eq 'true')) {
            $metadata = $row
            #continue
        }
        #Check for a manually entered imdb ID# and perform a lookup
        elseif ($row.imdbID -and ($row.PROCESSED -ne 'true')) {
            Write-Verbose ("Performing imdbID search for $title")
            $metadata = Get-Movie -imdbID $row.imdbID
        }
        #Else we assume this is a new entry, perform a regular lookup
        else {
            $metadata = Get-Movie -Title $Title
        }

        if($metadata.Error) {
            $Processed = 'Lookup Error'
            Write-Warning ("Error looking up $title")
        }
        elseif($row.Title -ne $metadata.Title) {
            $warn_title = $metadata.Title
            Write-Verbose ("Title mismatch $Title (original) >> $warn_title (lookup)") 
            #TODO: Write relevent info out to an error.log

            $title = $metadata.Title
        }
     
        $properties = [Ordered] @{
            "ID" = $row.ID;
            "imdbID" = $metadata.imdbID;
            "IMDB Link" = '';
            "TITLE" = $Title;
            "RATED" = $metadata.Rated;
            "EXPIRATION" = '="' + $row.EXPIRATION + '"'; #Preserve the NMPS Expiration code
            "IMDB RATING" = $metadata.imdbRating;
            "GENRE" = $metadata.Genre;
            "PLOT" = $metadata.Plot;
            "PROCESSED" = $Processed
        }
        $output += New-Object -TypeName psobject -Property $properties
        $counter++
    }
    $output
}

Export-ModuleMember -Function 'Generate-NMPSInventory'
#Generate-NMPSInventory -Input 'C:\Dev\Powershell\NMPSDecember2015Inventory.csv' -Verbose | Export-Csv -Path 'C:\Dev\Powershell\test_export.csv' -Force -NoTypeInformation
