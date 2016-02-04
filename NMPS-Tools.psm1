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
function Export-NMPSInventory {
    param (
        [Parameter(Mandatory=$true,Position=0)]
        [String]$Source,
        [Parameter(Mandatory=$true,Position=1)]
        [String]$Destination,
        [Switch]$Append,
        [Switch]$Reprocess,
        [Switch]$Overwrite = $True,
        [Switch]$WriteErrors = $True,
        [Int]$limit = -1
    )
    Write-Verbose $Source
    $source_csv = Import-Csv -Path $Source
    $records = $source_csv.length

    $output = @()

    if($WriteErrors) {
        $log = New-Item -Path (Split-Path $Destination) -Name 'nmpserror.log' -ItemType File -Force
        "NMPS Error Log" > $log
    }

    $counter = 0
    Write-Verbose "Init: Counter = $counter"
    foreach($row in $source_csv) {
        #Artificially limit #of lookups to perform
        if($counter -eq $limit) {
            Write-Warning ("Limit reached - halting processing Counter=$counter")
            break
        }
                
        Write-progress -Activity "Movie Lookup" -Status "Processing record $counter of $records" -PercentComplete ([int]($counter/$records * 100))

        ######
        #Actual Logic
        ######
        $Title     = $row.Title
        $Processed = 'true'

        #Check if this row has already been processed (and reprocess flag isn't set)
        if(-Not $Reprocess -and ($row.PROCESSED -eq 'true')) {
            Write-Verbose "Already processed record $($row.Title)"
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
            $Message = "LOOKUP::Error looking up Title: {0}, NMPS ID#: {1}" -f $row.Title,$row.ID
            if($log) {
                $Message >> $log
            }
            Write-Warning $Message
        }
        elseif($row.Title -ne $metadata.Title) {
            $Message = "MISMATCH::Title mismatch between original:{0} and API lookup:{1}. NMPS ID#{2}" -f $Title,$($metadata.title),$row.ID
            if($log) {
                $Message >> $log
            }
            
            Write-Warning $Message
            $Title = $metadata.Title
        }
     
        $properties = [Ordered] @{
            "ID" = $row.ID;
            "imdbID" = $metadata.imdbID;
            "IMDBLink" = '';
            "TITLE" = $Title;
            "RATED" = $metadata.Rated;
            "EXPIRATION" = '="' + $row.EXPIRATION + '"'; #Preserve the NMPS Expiration code
            "IMDBRATING" = $metadata.imdbRating;
            "GENRE" = $metadata.Genre;
            "PLOT" = $metadata.Plot;
            "PROCESSED" = $Processed
        }
        $output_obj = New-Object -TypeName psobject -Property $properties
        #Append on the fly, or wait til processing is complete to export
        if($Append) {
            $output_obj | Export-CSv -path $Destination -Append -NoTypeInformation -Force:$Overwrite
        } 
        $output += $output_obj
        $counter++
    }
    if(-not $Append) {
        $output | Export-CSV -Path $Destination -NoTypeInformation -Force:$Overwrite
    }
}

Export-ModuleMember -Function 'Export-NMPSInventory'
