function Get-Movie {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [String]$Title,
        [String]$Plot = "Short",
        [String]$Type = "Movie",
        [Switch]$Tomatoes        
    )

    $request = "t=$Title&type=movie&plot=$Plot&r=json&tomatoes=$Tomatoes"
    Write-Verbose $request

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
        [String]$Path
    )

    $source = Import-Csv -Path $Path

    $output = @()
    foreach($row in $source) {
        $metadata = Get-Movie -Title $row.TITLE

        $properties = [Ordered] @{
            "ID" = $row.ID;
            "imdbID" = $metadata.imdbID;
            "IMDB Link" = '';
            "TITLE" = $row.TITLE;
            "RATED" = $metadata.Rated;
            "EXPIRATION" = '="' + $row.EXPIRATION + '"';
            "IMDB RATING" = $metadata.imdbRating;
            "GENRE" = $metadata.Genre;
            "PLOT" = $metadata.Plot
        }
        $output += New-Object -TypeName psobject -Property $properties
    }
    #Push it to the pipeline
    $output
}

Generate-NMPSInventory -Path 'C:\temp\test_inventory.csv' | Export-Csv -Path 'C:\temp\test_export.csv' -Force
