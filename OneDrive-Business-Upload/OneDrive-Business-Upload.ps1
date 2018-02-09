try {
    # SharePoint Online Client Components SDK
    # https://www.microsoft.com/en-us/download/details.aspx?id=42038
    Add-Type -Path 'C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll'
    Add-Type -Path 'C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll'
} catch {
    Write-Host $_.Exception.Message
    Write-Host "Error reading sharepoint dll's!" -ForegroundColor Red
    break
}

function UploadFileToOneDriveBusiness ($SourcePath, $DestinationPath, [Microsoft.SharePoint.Client.ClientContext]$SharePointContext, $SharePointLibrary) {

    # Set slice size for the large file upload approach
    # Files smaller than this value getting uploaded with the regular approach
    $FileChunkSizeInMB = 9

    # Format destination path
    if (!$DestinationPath.StartsWith('/')) {
        $DestinationPath = '/' + $DestinationPath
    }

    # Each sliced upload requires a unique ID.
    $UploadId = [GUID]::NewGuid()

    # Get the folder to upload into. 
    try {
        Write-Host "Loading sharepoint library '$SharePointLibrary'..."
        $List = $SharePointContext.Web.Lists.GetByTitle($SharePointLibrary)
        $SharePointContext.Load($List)
        $SharePointContext.Load($List.RootFolder)
        ExecuteQueryWithIncrementalRetry -SharePointContext $SharePointContext
        Write-Host "  -> Loading finished successfully" -ForegroundColor Green
    } catch {
        Write-Host $_.Exception.Message -ForegroundColor Red
        Write-Host "Error reading sharepoint document library '$SharePointLibrary'!" -ForegroundColor Red
        break
    }

    # Get the information about the folder that will hold the file.
    $ServerRelativeUrlOfRootFolder = $List.RootFolder.ServerRelativeUrl

    # Set the relative URL for the destination file
    $DestinationURL = $ServerRelativeUrlOfRootFolder + $DestinationPath

    # Calculate block size in bytes.
    $BlockSize = $fileChunkSizeInMB * 1024 * 1024

    # Get the size of the file.
    $FileSize = (Get-Item $SourcePath).length

    # Get the number of slices
    $SlicesCount = [math]::Round($FileSize / $BlockSize)

    if ($FileSize -le $FileChunkSizeInMB) {

        # Use regular file upload approach / without slicing

        try {
            Write-Debug "Opening file stream..."
            $FileStream = New-Object IO.FileStream($SourcePath,[System.IO.FileMode]::Open)
            Write-Debug "  -> File stream opened successfully"
            $FileCreationInfo = New-Object Microsoft.SharePoint.Client.FileCreationInformation
            $FileCreationInfo.Overwrite = $true
            $FileCreationInfo.ContentStream = $FileStream
            $FileCreationInfo.URL = $DestinationURL
            $Upload = $List.RootFolder.Files.Add($FileCreationInfo)
            Write-Host "Uploading file '$SourcePath' to '$DestinationPath'..."
            $SharePointContext.Load($Upload)
            $SharePointContext.Load($List.ContentTypes)
            ExecuteQueryWithIncrementalRetry -SharePointContext $SharePointContext
            
            Write-Host "  -> Uploading finished successfully" -ForegroundColor Green
        } catch {
            Write-Host $_.Exception.Message -ForegroundColor Red
            Write-Host "Error uploading '$SourcePath' to '$DestinationPath'!" -ForegroundColor Red
            break
        } finally {
            DisposeFileStream -FileStream $FileStream
        }

    } else {

        # Use large file upload approach / with slicing

        $BytesUploaded = $null
        $FileStream = $null

        try {
            
            $FileStream = [System.IO.File]::Open($SourcePath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)

            $BinaryReader = New-Object System.IO.BinaryReader($FileStream)

            $Buffer = New-Object System.Byte[]($BlockSize)
            $lastBuffer = $null
            $fileoffset = 0
            $totalBytesRead = 0
            $bytesRead
            $first = $true
            $last = $false

            # Read data from file system in blocks. 
            $i = 0
            while(($bytesRead = $BinaryReader.Read($Buffer, 0, $Buffer.Length)) -gt 0) {

                $i++
                $totalBytesRead = $totalBytesRead + $bytesRead

                # You've reached the end of the file.
                if($totalBytesRead -eq $FileSize) {
                    $last = $true
                    # Copy to a new buffer that has the correct size.
                    $lastBuffer = New-Object System.Byte[]($bytesRead)
                    [array]::Copy($Buffer, 0, $lastBuffer, 0, $bytesRead)
                }

                If($first) {
                    $ContentStream = New-Object System.IO.MemoryStream

                    # Add an empty file.
                    $fileInfo = New-Object Microsoft.SharePoint.Client.FileCreationInformation
                    $fileInfo.ContentStream = $ContentStream
                    $fileInfo.Url = $DestinationURL
                    $fileInfo.Overwrite = $true
                    $Upload = $List.RootFolder.Files.Add($fileInfo)
                    $SharePointContext.Load($Upload)

                    # Start upload by uploading the first slice.
                    
                    $s = [System.IO.MemoryStream]::new($Buffer) 
                    
                    # Call the start upload method on the first slice.
                    Write-Host "Uploading sliced file '$SourcePath' to '$DestinationPath'..."
                    Write-Host "  -> Uploading slice nr 1 of $SlicesCount..."
                    $BytesUploaded = $Upload.StartUpload($UploadId, $s)
                    ExecuteQueryWithIncrementalRetry -SharePointContext $SharePointContext
                    
                    # fileoffset is the pointer where the next slice will be added.
                    $fileoffset = $BytesUploaded.Value
                    
                    # You can only start the upload once.
                    $first = $false

                } else {
                    # Get a reference to your file.
                    $Upload = $SharePointContext.Web.GetFileByServerRelativeUrl($DestinationURL);
                    
                    If($last) {
                        
                        # Is this the last slice of data?
                        $s = [System.IO.MemoryStream]::new($lastBuffer)
                        

                        # End sliced upload by calling FinishUpload.
                        $Upload = $Upload.FinishUpload($UploadId, $fileoffset, $s)
                        ExecuteQueryWithIncrementalRetry -SharePointContext $SharePointContext

                        Write-Host "  -> Uploading finished successfully" -ForegroundColor Green
                        
                    } else {
                        
                        $s = [System.IO.MemoryStream]::new($Buffer)
                        
                        # Continue sliced upload.
                        Write-Host "  -> Uploading slice nr $i of $SlicesCount..."
                        $BytesUploaded = $Upload.ContinueUpload($UploadId, $fileoffset, $s)
                        ExecuteQueryWithIncrementalRetry -SharePointContext $SharePointContext
                        
                        # Update fileoffset for the next slice.
                        $fileoffset = $BytesUploaded.Value
                        
                    }
                }
                
            }
        } catch {
            Write-Host $_.Exception.Message -ForegroundColor Red
            Write-Host "Error uploading '$SourcePath' to '$DestinationPath'!" -ForegroundColor Red
            break
        } finally {
            DisposeFileStream -FileStream $FileStream
        }

    }

}

function ExecuteQueryWithIncrementalRetry([Microsoft.SharePoint.Client.ClientContext]$SharePointContext, $retryCount, $delay) {

    if ($retryCount -eq $null) {
        $retryCount = 5 # default to 5
    }

    if ($delay -eq $null) {
        $delay = 500 # default to 500
    }

    $retryAttempts = 0
    $backoffInterval = $delay

    if ($retryCount -le 0) {
        throw New-Object ArgumentException("Provide a retry count greater than zero.")
    }

    if ($delay -le 0) {
        throw New-Object ArgumentException("Provide a delay greater than zero.")
    }

    # Do while retry attempt is less than retry count
    while ($retryAttempts -lt $retryCount) {
        try {
            $SharePointContext.ExecuteQuery()
            return
        } catch [System.Net.WebException] {

            $response = [System.Net.HttpWebResponse]$_.Exception.Response

            # Check if request was throttled - http status code 429
            # Check is request failed due to server unavailable - http status code 503
            if ($response -ne $null -and ($response.StatusCode -eq 429 -or $response.StatusCode -eq 503)) {
                # Output status to console. Should be changed as Debug.WriteLine for production usage.
                Write-Host " -> Request frequency exceeded usage limits. Sleeping for $backoffInterval seconds before retrying..."

                # Add delay for retry
                Start-Sleep -m $backoffInterval

                # Add to retry count and increase delay.
                $retryAttempts++
                $backoffInterval = $backoffInterval * 2
            }
            else {
                Write-Host $_.Exception.Message -ForegroundColor Red
                Write-Host "Unknown HTTP Response!" -ForegroundColor Red
                break
            }
        }
    }
    Write-Host "Maximum retry attempts ($retryCount) exceeded!" -ForegroundColor Red
    break
}

function DisposeFileStream ($FileStream) {
    if ($FileStream) {
        Write-Debug "Closing file stream..."
        try {
            $FileStream.Dispose()
            Write-Debug " -> File stream closed successfully"
        } catch {
            Write-Host $_.Exception.Message -ForegroundColor Red
            Write-Host "Error closing file stream!" -ForegroundColor Red
        }
    }
}

$User = 'xxxx@xxxxxx.com'
$Password  = ConvertTo-SecureString ‘xxxxxxxx’ -AsPlainText -Force

$SharePointUrl = 'https://xxxxxxxx-my.sharepoint.com/personal/xxxxxxxxxxx'
$SharePointLibraryShownName = 'Dokumente'


$SPContext  = New-Object Microsoft.SharePoint.Client.ClientContext($SharePointUrl)
$SPCreds = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($User,$Password)
$SPContext.Credentials = $SPCreds

UploadFileToOneDriveBusiness -SourcePath "C:\test.txt" -DestinationPath "test.txt" -SharePointContext $SPContext -SharePointLibrary "$SharePointLibraryShownName"
UploadFileToOneDriveBusiness -SourcePath "C:\test2.txt" -DestinationPath "/test2.txt" -SharePointContext $SPContext -SharePointLibrary "$SharePointLibraryShownName"
UploadFileToOneDriveBusiness -SourcePath "C:\test3.txt" -DestinationPath "/testfolder/test3.txt" -SharePointContext $SPContext -SharePointLibrary "$SharePointLibraryShownName"
