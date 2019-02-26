# 
# Merge-MailboxFolder.ps1 
# 
# By David Barrett, Microsoft Ltd. 2015-2016. Use at your own risk.  No warranties are given. 
# 
#  DISCLAIMER: 
# THIS CODE IS SAMPLE CODE. THESE SAMPLES ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND. 
# MICROSOFT FURTHER DISCLAIMS ALL IMPLIED WARRANTIES INCLUDING WITHOUT LIMITATION ANY IMPLIED WARRANTIES OF MERCHANTABILITY OR OF FITNESS FOR 
# A PARTICULAR PURPOSE. THE ENTIRE RISK ARISING OUT OF THE USE OR PERFORMANCE OF THE SAMPLES REMAINS WITH YOU. IN NO EVENT SHALL 
# MICROSOFT OR ITS SUPPLIERS BE LIABLE FOR ANY DAMAGES WHATSOEVER (INCLUDING, WITHOUT LIMITATION, DAMAGES FOR LOSS OF BUSINESS PROFITS, 
# BUSINESS INTERRUPTION, LOSS OF BUSINESS INFORMATION, OR OTHER PECUNIARY LOSS) ARISING OUT OF THE USE OF OR INABILITY TO USE THE 
# SAMPLES, EVEN IF MICROSOFT HAS BEEN ADVISED OF THE POSSIBILITY OF SUCH DAMAGES. BECAUSE SOME STATES DO NOT ALLOW THE EXCLUSION OR LIMITATION 
# OF LIABILITY FOR CONSEQUENTIAL OR INCIDENTAL DAMAGES, THE ABOVE LIMITATION MAY NOT APPLY TO YOU. 
 
param ( 
    [Parameter(Position=0,Mandatory=$False,HelpMessage="Specifies the source mailbox (from which items will be moved/copied)")] 
    [ValidateNotNullOrEmpty()] 
    [string]$SourceMailbox, 
     
    [Parameter(Position=1,Mandatory=$False,HelpMessage="Specifies the target mailbox (if not specified, the source mailbox is also the target)")] 
    [ValidateNotNullOrEmpty()] 
    [string]$TargetMailbox, 
     
    [Parameter(Position=2,Mandatory=$False,HelpMessage="Specifies the folder(s) to be merged")] 
    [ValidateNotNullOrEmpty()] 
    $MergeFolderList, 
         
    [Parameter(Mandatory=$False,HelpMessage="If specified, only items that match the given AQS filter will be moved `r`n(see https://msdn.microsoft.com/EN-US/library/dn579420(v=exchg.150).aspx)")] 
    [string]$SearchFilter, 
 
    [Parameter(Mandatory=$False,HelpMessage="When specified, the folders in MergeFolderList are identified by EwsId (not path)")] 
    [switch]$ByFolderId, 
 
    [Parameter(Mandatory=$False,HelpMessage="When specified, the folders in MergeFolderList are identified by EntryId (not path)")] 
    [switch]$ByEntryId, 
 
    [Parameter(Mandatory=$False,HelpMessage="When specified, subfolders will also be processed")] 
    [switch]$ProcessSubfolders, 
         
    [Parameter(Mandatory=$False,HelpMessage="When specified, items in subfolders will all be moved to specified target folder (hierarchy will NOT be maintained)")] 
    [switch]$CombineSubfolders, 
         
    [Parameter(Mandatory=$False,HelpMessage="When specified, if the target folder doesn't exist, then it will be created (if possible)")] 
    [switch]$CreateTargetFolder, 
         
    [Parameter(Mandatory=$False,HelpMessage="When specified, the source mailbox being accessed will be the archive mailbox")] 
    [switch]$SourceArchive, 
         
    [Parameter(Mandatory=$False,HelpMessage="When specified, the target mailbox being accessed will be the archive mailbox")] 
    [switch]$TargetArchive, 
         
    [Parameter(Mandatory=$False,HelpMessage="When specified, hidden (associated) items of the folder are processed (normal items are ignored)")] 
    [switch]$AssociatedItems, 
         
    [Parameter(Mandatory=$False,HelpMessage="When specified, the source folder will be deleted after the move (can't be used with -Copy)")] 
    [switch]$Delete, 
         
    [Parameter(Mandatory=$False,HelpMessage="When specified, items are copied rather than moved (can't be used with -Delete)")] 
    [switch]$Copy, 
         
    [Parameter(Mandatory=$False,HelpMessage="Credentials used to authenticate with EWS")] 
    [System.Management.Automation.PSCredential]$Credentials, 
                 
    [Parameter(Mandatory=$False,HelpMessage="Username used to authenticate with EWS")] 
    [string]$Username, 
     
    [Parameter(Mandatory=$False,HelpMessage="Password used to authenticate with EWS")] 
    [string]$Password, 
     
    [Parameter(Mandatory=$False,HelpMessage="Domain used to authenticate with EWS")] 
    [string]$Domain, 
     
    [Parameter(Mandatory=$False,HelpMessage="Whether we are using impersonation to access the mailbox")] 
    [switch]$Impersonate, 
     
    [Parameter(Mandatory=$False,HelpMessage="EWS Url (if omitted, then autodiscover is used)")]     
    [string]$EwsUrl, 
     
    [Parameter(Mandatory=$False,HelpMessage="Path to managed API (if omitted, a search of standard paths is performed)")]     
    [string]$EWSManagedApiPath = "", 
     
    [Parameter(Mandatory=$False,HelpMessage="Whether to ignore any SSL errors (e.g. invalid certificate)")]     
    [switch]$IgnoreSSLCertificate, 
     
    [Parameter(Mandatory=$False,HelpMessage="Whether to allow insecure redirects when performing autodiscover")]     
    [switch]$AllowInsecureRedirection, 
     
    [Parameter(Mandatory=$False,HelpMessage="Log file - activity is logged to this file if specified")]     
    [string]$LogFile = "", 
 
    [Parameter(Mandatory=$False,HelpMessage="Trace file - if specified, EWS tracing information is written to this file")]     
    [string]$TraceFile, 
 
    [Parameter(Mandatory=$False,HelpMessage="Throttling delay (time paused between sending EWS requests) - note that this will be increased automatically if throttling is detected")]     
    [int]$ThrottlingDelay = 0, 
 
    [Parameter(Mandatory=$False,HelpMessage="Batch size (number of items batched into one EWS request) - this will be decreased if throttling is detected")]     
    [int]$BatchSize = 150 
) 
 
 
# Define our functions 
 
Function Log([string]$Details, [ConsoleColor]$Colour) 
{ 
    if ($Colour -eq $null) 
    { 
        $Colour = [ConsoleColor]::White 
    } 
    Write-Host $Details -ForegroundColor $Colour 
    if ( $LogFile -eq "" ) { return } 
    "$([DateTime]::Now.ToShortDateString()) $([DateTime]::Now.ToLongTimeString())   $Details" | Out-File $LogFile -Append 
} 
 
Function LogVerbose([string]$Details) 
{ 
    Write-Verbose $Details 
    if ( $LogFile -eq "" ) { return } 
    "$([DateTime]::Now.ToShortDateString()) $([DateTime]::Now.ToLongTimeString())   $Details" | Out-File $LogFile -Append 
} 
 
Function LoadEWSManagedAPI() 
{ 
    # Find and load the managed API 
     
    if ( ![string]::IsNullOrEmpty($EWSManagedApiPath) ) 
    { 
        if ( { Test-Path $EWSManagedApiPath } ) 
        { 
            Add-Type -Path $EWSManagedApiPath 
            return $true 
        } 
        Write-Host ( [string]::Format("Managed API not found at specified location: {0}", $EWSManagedApiPath) ) -ForegroundColor Yellow 
    } 
     
    $a = Get-ChildItem -Recurse "C:\Program Files (x86)\Microsoft\Exchange\Web Services" -ErrorAction SilentlyContinue | Where-Object { ($_.PSIsContainer -eq $false) -and ( $_.Name -eq "Microsoft.Exchange.WebServices.dll" ) } 
    if (!$a) 
    { 
        $a = Get-ChildItem -Recurse "C:\Program Files\Microsoft\Exchange\Web Services" -ErrorAction SilentlyContinue | Where-Object { ($_.PSIsContainer -eq $false) -and ( $_.Name -eq "Microsoft.Exchange.WebServices.dll" ) } 
    } 
     
    if ($a)     
    { 
        # Load EWS Managed API 
        Write-Host ([string]::Format("Using managed API {0} found at: {1}", $a.VersionInfo.FileVersion, $a.VersionInfo.FileName)) -ForegroundColor Gray 
        Add-Type -Path $a.VersionInfo.FileName 
        $script:EWSManagedApiPath = $a.VersionInfo.FileName 
        return $true 
    } 
    return $false 
} 
 
Function CurrentUserPrimarySmtpAddress() 
{ 
    # Attempt to retrieve the current user's primary SMTP address 
    $searcher = [adsisearcher]"(samaccountname=$env:USERNAME)" 
    $result = $searcher.FindOne() 
 
    if ($result -ne $null) 
    { 
        $mail = $result.Properties["mail"] 
        return $mail 
    } 
    return $null 
} 
 
Function TrustAllCerts() 
{ 
    # Implement call-back to override certificate handling (and accept all) 
    $Provider=New-Object Microsoft.CSharp.CSharpCodeProvider 
    $Compiler=$Provider.CreateCompiler() 
    $Params=New-Object System.CodeDom.Compiler.CompilerParameters 
    $Params.GenerateExecutable=$False 
    $Params.GenerateInMemory=$True 
    $Params.IncludeDebugInformation=$False 
    $Params.ReferencedAssemblies.Add("System.DLL") | Out-Null 
 
    $TASource=@' 
        namespace Local.ToolkitExtensions.Net.CertificatePolicy { 
        public class TrustAll : System.Net.ICertificatePolicy { 
            public TrustAll() 
            { 
            } 
            public bool CheckValidationResult(System.Net.ServicePoint sp, 
                                                System.Security.Cryptography.X509Certificates.X509Certificate cert,  
                                                System.Net.WebRequest req, int problem) 
            { 
                return true; 
            } 
        } 
        } 
'@  
    $TAResults=$Provider.CompileAssemblyFromSource($Params,$TASource) 
    $TAAssembly=$TAResults.CompiledAssembly 
 
    ## We now create an instance of the TrustAll and attach it to the ServicePointManager 
    $TrustAll=$TAAssembly.CreateInstance("Local.ToolkitExtensions.Net.CertificatePolicy.TrustAll") 
    [System.Net.ServicePointManager]::CertificatePolicy=$TrustAll 
} 
 
Function CreateTraceListener($service) 
{ 
    # Create trace listener to capture EWS conversation (useful for debugging) 
    if ($script:Tracer -eq $null) 
    { 
        $Provider=New-Object Microsoft.CSharp.CSharpCodeProvider 
        $Params=New-Object System.CodeDom.Compiler.CompilerParameters 
        $Params.GenerateExecutable=$False 
        $Params.GenerateInMemory=$True 
        $Params.IncludeDebugInformation=$False 
        $Params.ReferencedAssemblies.Add("System.dll") | Out-Null 
        $Params.ReferencedAssemblies.Add($EWSManagedApiPath) | Out-Null 
 
        $traceFileForCode = $traceFile.Replace("\", "\\") 
 
        if (![String]::IsNullOrEmpty($TraceFile)) 
        { 
            LogVerbose "Tracing to: $TraceFile" 
        } 
 
        $TraceListenerClass = @" 
            using System; 
            using System.Text; 
            using System.IO; 
            using System.Threading; 
            using Microsoft.Exchange.WebServices.Data; 
         
            namespace TraceListener { 
                class EWSTracer: Microsoft.Exchange.WebServices.Data.ITraceListener 
                { 
                    private StreamWriter _traceStream = null; 
                    private string _lastResponse = String.Empty; 
 
                    public EWSTracer() 
                    { 
                        try 
                        { 
                            _traceStream = File.AppendText("$traceFileForCode"); 
                        } 
                        catch { } 
                    } 
 
                    ~EWSTracer() 
                    { 
                        Close(); 
                    } 
 
                    public void Close() 
                    { 
                        try 
                        { 
                            _traceStream.Flush(); 
                            _traceStream.Close(); 
                        } 
                        catch { } 
                    } 
 
 
                    public void Trace(string traceType, string traceMessage) 
                    { 
                        if ( traceType.Equals("EwsResponse") ) 
                            _lastResponse = traceMessage; 
 
                        if ( traceType.Equals("EwsRequest") ) 
                            _lastResponse = String.Empty; 
 
                        if (_traceStream == null) 
                            return; 
 
                        lock (this) 
                        { 
                            try 
                            { 
                                _traceStream.WriteLine(traceMessage); 
                                _traceStream.Flush(); 
                            } 
                            catch { } 
                        } 
                    } 
 
                    public string LastResponse 
                    { 
                        get { return _lastResponse; } 
                    } 
                } 
            } 
"@ 
 
        $TraceCompilation=$Provider.CompileAssemblyFromSource($Params,$TraceListenerClass) 
        $TraceAssembly=$TraceCompilation.CompiledAssembly 
        $script:Tracer=$TraceAssembly.CreateInstance("TraceListener.EWSTracer") 
    } 
 
    # Attach the trace listener to the Exchange service 
    $service.TraceListener = $script:Tracer 
} 
 
Function DecreaseBatchSize() 
{ 
    param ( 
        $DecreaseMultiplier = 0.8 
    ) 
 
    $script:currentBatchSize = [int]($script:currentBatchSize * $DecreaseMultiplier) 
    if ($script:currentBatchSize -lt 50) { $script:currentBatchSize = 50 } 
    LogVerbose "Retrying with smaller batch size of $($script:currentBatchSize)" 
} 
 
Function IncreaseThrottlingDelay() 
{ 
    # Increase our throttling delay to try and avoid throttling (we only increase to a maximum delay of 10 seconds between requests) 
    if ( $script:currentThrottlingDelay -lt 10000) 
    { 
        if ($script:currentThrottlingDelay -lt 1) 
        { 
            $script:currentThrottlingDelay = $ThrottlingDelay 
            if ($script:currentThrottlingDelay -lt 1)  
            { 
                # In case throttling delay parameter is set to 0, or a silly value 
                $script:currentThrottlingDelay = 100 
            } 
        } 
        else 
        { 
            $script:currentThrottlingDelay = $script:currentThrottlingDelay * 2 
        } 
        if ( $script:currentThrottlingDelay -gt 10000) 
        { 
            $script:currentThrottlingDelay = 10000 
        } 
    } 
    LogVerbose "Updated throttling delay to $($script:currentThrottlingDelay)ms" 
} 
 
Function Throttled() 
{ 
    # Checks if we've been throttled.  If we have, we wait for the specified number of BackOffMilliSeconds before returning 
 
    if ([String]::IsNullOrEmpty($script:Tracer.LastResponse)) 
    { 
        return $false # Throttling does return a response, if we don't have one, then throttling probably isn't the issue (though sometimes throttling just results in a timeout) 
    } 
 
    $lastResponse = $script:Tracer.LastResponse.Replace("<?xml version=`"1.0`" encoding=`"utf-8`"?>", "") 
    $lastResponse = "<?xml version=`"1.0`" encoding=`"utf-8`"?>$lastResponse" 
    $responseXml = [xml]$lastResponse 
 
    if ($responseXml.Trace.Envelope.Body.Fault.detail.MessageXml.Value.Name -eq "BackOffMilliseconds") 
    { 
        # We are throttled, and the server has told us how long to back off for 
        IncreaseThrottlingDelay 
 
        # Now back off for the time given by the server 
        Log "Throttling detected, server requested back off for $($responseXml.Trace.Envelope.Body.Fault.detail.MessageXml.Value."#text") milliseconds" Yellow 
        Sleep -Milliseconds $responseXml.Trace.Envelope.Body.Fault.detail.MessageXml.Value."#text" 
        Log "Throttling budget should now be reset, resuming operations" Gray 
        return $true 
    } 
    return $false 
} 
 
function ThrottledFolderBind() 
{ 
    param ( 
        [Microsoft.Exchange.WebServices.Data.FolderId]$folderId, 
        $propset = $null) 
 
    LogVerbose "Attempting to bind to folder $folderId" 
    $folder = $null 
    try 
    { 
        if ($propset -eq $null) 
        { 
            $folder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($script:service, $folderId) 
        } 
        else 
        { 
            $folder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($script:service, $folderId, $propset) 
        } 
        Sleep -Milliseconds $script:currentThrottlingDelay 
        if (!($folder -eq $null)) 
        { 
            LogVerbose "Successfully bound to folder $folderId" 
        } 
        return $folder 
    } 
    catch {} 
 
    if (Throttled) 
    { 
        try 
        { 
            if ($propset -eq $null) 
            { 
                $folder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($script:service, $folderId) 
            } 
            else 
            { 
                $folder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($script:service, $folderId, $propset) 
            } 
            if (!($folder -eq $null)) 
            { 
                LogVerbose "Successfully bound to folder $folderId" 
            } 
            return $folder 
        } 
        catch {} 
    } 
 
    # If we get to this point, we have been unable to bind to the folder 
    LogVerbose "FAILED to bind to folder $folderId" 
    return $null 
} 
 
Function RemoveProcessedItemsFromList() 
{ 
    # Process the results of a batch move/copy and remove any items that were successfully moved from our list of items to move 
    param ( 
        $requestedItems, 
        $results, 
        $Items 
    ) 
 
    if ($results -ne $null) 
    { 
        $failed = 0 
        for ($i = 0; $i -lt $requestedItems.Count; $i++) 
        { 
            if ($results[$i].ErrorCode -eq "NoError") 
            { 
                $Items.Remove($requestedItems[$i]) 
            } 
            else 
            { 
                if ( ($results[$i].ErrorCode -eq "ErrorMoveCopyFailed") -or ($results[$i].ErrorCode -eq "ErrorInvalidOperation") ) 
                { 
                    # This is a permanent error, so we remove the item from the list 
                    $Items.Remove($requestedItems[$i]) 
                } 
                LogVerbose("Error $($results[$i].ErrorCode) reported for item: $($requestedItems[$i].UniqueId)") 
                $failed++ 
            }  
        } 
    } 
    if ( $failed -gt 0 ) 
    { 
        Log "$failed items reported error during batch request (if throttled, this is expected)" Yellow 
    } 
} 
 
Function ThrottledBatchMove() 
{ 
    # Send request to move/copy items, allowing for throttling (which in this case is likely to manifest as time-out errors) 
    param ( 
        $ItemsToMove, 
        $TargetFolderId, 
        $Copy 
    ) 
 
    if ($script:currentBatchSize -lt 1) { $script:currentBatchSize = $BatchSize } 
 
    $itemId = New-Object Microsoft.Exchange.WebServices.Data.ItemId("xx") 
    $itemIdType = [Type] $itemId.GetType() 
    #$baseList = [System.Collections.Generic.List``1] 
    $genericItemIdList = [System.Collections.Generic.List``1].MakeGenericType(@($itemIdType)) 
     
    $finished = $false 
    if ($Copy) 
    { 
        $progressActivity = "Copying items" 
    } 
    else 
    { 
        $progressActivity = "Moving items" 
    } 
    $totalItems = $ItemsToMove.Count 
    Write-Progress -Activity $progressActivity -Status "0% complete" -PercentComplete 0 
 
    if ( $totalItems -gt 10000 ) 
    { 
        if ( $script:currentThrottlingDelay -lt 1000 ) 
        { 
            $script:currentThrottlingDelay = 1000 
            LogVerbose "Large number of items will be processed, so throttling delay set to 1000ms" 
        } 
    } 
 
    while ( !$finished ) 
    { 
        $global:moveIds = [Activator]::CreateInstance($genericItemIdList) 
         
        for ([int]$i=0; $i -lt $BatchSize; $i++) 
        { 
            if ($ItemsToMove[$i] -ne $null) 
            { 
                $moveIds.Add($ItemsToMove[$i]) 
            } 
            if ($i -ge $ItemsToMove.Count) 
                { break } 
        } 
 
        $results = $null 
        try 
        { 
            if ( $Copy ) 
            { 
                LogVerbose "Sending batch request to copy $($moveIds.Count) items ($($ItemsToMove.Count) remaining)" 
                $results = $script:service.CopyItems( $moveIds, $TargetFolderId, $false ) 
            } 
            else 
            { 
                LogVerbose "Sending batch request to move $($moveIds.Count) items ($($ItemsToMove.Count) remaining)" 
                $results = $script:service.MoveItems( $moveIds, $TargetFolderId, $false) 
            } 
            Sleep -Milliseconds $script:currentThrottlingDelay 
        } 
        catch 
        { 
            if ( Throttled ) 
            { 
                # We've been throttled, so we reduce batch size (to a minimum size of 50) and try again 
                if ($BatchSize -gt 50) 
                { 
                    DecreaseBatchSize 
                } 
                else 
                { 
                    $finished = $true 
                } 
            } 
            elseif ($Error[0].Exception.InnerException.ToString().Contains("The operation has timed out")) 
            { 
                # We've probably been throttled, so we'll reduce the batch size and try again 
                if ($script:currentBatchSize -gt 50) 
                { 
                    LogVerbose "Timeout error received" 
                    DecreaseBatchSize 
                } 
                else 
                { 
                    $finished = $true 
                } 
            } 
            else 
            { 
                $lastResponse = $script:Tracer.LastResponse.Replace("<?xml version=`"1.0`" encoding=`"utf-8`"?>", "") 
                $lastResponse = "<?xml version=`"1.0`" encoding=`"utf-8`"?>$lastResponse" 
                $responseXml = [xml]$lastResponse 
                if ($responseXml.Trace.Envelope.Body.Fault.detail.ResponseCode.Value -eq "ErrorNoRespondingCASInDestinationSite") 
                { 
                    # We get this error if the destination CAS (request was proxied) hasn't returned any data within the timeout (usually 60 seconds) 
                    # Reducing the batch size should help here, and we want to reduce it quite aggressively 
                    if ($BatchSize -gt 50) 
                    { 
                        LogVerbose "ErrorNoRespondingCASInDestinationSite error received" 
                        DecreaseBatchSize 0.7 
                    } 
                    else 
                    { 
                        $finished = $true 
                    } 
                } 
                else 
                { 
                    Log "Unexpected error: $($Error[0].Exception.InnerException.ToString())" Red 
                    $finished = $true 
                } 
            } 
        } 
 
        RemoveProcessedItemsFromList $moveIds $results $ItemsToMove 
 
        $percentComplete = ( ($totalItems - $ItemsToMove.Count) / $totalItems ) * 100 
        Write-Progress -Activity $progressActivity -Status "$percentComplete% complete" -PercentComplete $percentComplete 
 
        if ($ItemsToMove.Count -eq 0) 
        { 
            $finished = $True 
            Write-Progress -Activity $progressActivity -Status "100% complete" -Completed 
        } 
    } 
} 
 
Function MoveItems() 
{ 
    # Process all the items in the given source folder, and move (or copy) them to the target 
     
    if ($args -eq $null) 
    { 
        throw "No folders specified for MoveItems" 
    } 
    $SourceFolderObject, $TargetFolderObject = $args[0] 
     
    if ($SourceFolderObject.Id -eq $TargetFolderObject.Id) 
    { 
        Log "Cannot move or copy from/to the same folder (source folder Id and target folder Id are the same)" Red 
        return 
    } 
     
    if ($Copy) 
    { 
        $action = "Copy" 
        $actioning = "Copying" 
    } 
    else 
    { 
        $action = "Move" 
        $actioning = "Moving" 
    } 
    Log "$actioning from $($SourceMailbox):$(GetFolderPath($SourceFolderObject)) to $($TargetMailbox):$(GetFolderPath($TargetFolderObject))" White 
     
    # Set parameters - we will process in batches of 500 for the FindItems call 
    $Offset = 0 
    $PageSize = 1000 # We're only querying Ids, so 1000 items at a time is reasonable 
    $MoreItems = $true 
    $moveCountSuccess = 0 
    $moveCountFail = 0 
 
    # We create a list of all the items we need to move, and then batch move them later (much faster than doing it one at a time) 
    $itemsToMove = New-Object System.Collections.ArrayList 
    $i = 0 
     
    $progressActivity = "Reading items in folder $($SourceMailbox):$(GetFolderPath($SourceFolderObject))" 
    LogVerbose "Building list of items to $($action.ToLower())" 
    Write-Progress -Activity $progressActivity -Status "0 items found" -PercentComplete -1 
 
    if (![String]::IsNullOrEmpty($SearchFilter)) 
    { 
        LogVerbose "Search query being applied: $SearchFilter" 
    } 
 
    while ($MoreItems) 
    { 
        $View = New-Object Microsoft.Exchange.WebServices.Data.ItemView($PageSize, $Offset, [Microsoft.Exchange.Webservices.Data.OffsetBasePoint]::Beginning) 
        $View.PropertySet = New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly) 
        if ($AssociatedItems) 
        { 
            $View.Traversal = [Microsoft.Exchange.WebServices.Data.ItemTraversal]::Associated 
        } 
 
        $FindResults = $null 
        try 
        { 
            if (![String]::IsNullOrEmpty($SearchFilter)) 
            { 
                # We have a search filter, so need to apply this 
                $FindResults = $SourceFolderObject.FindItems($SearchFilter, $View) 
            } 
            else 
            { 
                # No search filter, we want everything 
                $FindResults = $SourceFolderObject.FindItems($View) 
            } 
            Sleep -Milliseconds $script:currentThrottlingDelay 
        } 
        catch 
        { 
            # We have an error, so check if we are being throttled 
            if (Throttled) 
            { 
                $FindResults = $null # We do this to retry the request 
            } 
            else 
            { 
                Log "Error when querying items: $($Error[0])" Red 
                $MoreItems = $false 
            } 
        } 
         
        if ($FindResults) 
        { 
            ForEach ($item in $FindResults.Items) 
            { 
                [void]$itemsToMove.Add($item.Id) 
            } 
            $MoreItems = $FindResults.MoreAvailable 
            if ($MoreItems) 
            { 
                LogVerbose "$($itemsToMove.Count) items found so far, more available" 
            } 
            $Offset += $PageSize 
        } 
        Write-Progress -Activity $progressActivity -Status "$($itemsToMove.Count) items found" -PercentComplete -1 
    } 
    Write-Progress -Activity $progressActivity -Status "$($itemsToMove.Count) items found" -Completed 
 
    if ( $itemsToMove.Count -gt 0 ) 
    { 
        Log "$($itemsToMove.Count) items found; attempting to $($action.ToLower())" Green 
        ThrottledBatchMove $itemsToMove $TargetFolderObject.Id $Copy 
    } 
    else 
    { 
        Log "No matching items were found" Green 
    } 
 
    # Now process any subfolders 
    if ($ProcessSubFolders) 
    { 
        if ($SourceFolderObject.ChildFolderCount -gt 0) 
        { 
            # Deal with any subfolders first 
            $FolderView = New-Object Microsoft.Exchange.WebServices.Data.FolderView(1000) 
            $SourceFindFolderResults = $SourceFolderObject.FindFolders($FolderView) 
            Sleep -Milliseconds $script:currentThrottlingDelay 
            ForEach ($SourceSubFolderObject in $SourceFindFolderResults.Folders) 
            { 
                if ($CombineSubfolders) 
                { 
                    # We are moving all subfolder items into the target folder (ignoring hierarchy) 
                     MoveItems($SourceSubFolderObject, $TargetFolderObject) 
                } 
                else 
                { 
                    # We need to recreate folder hierarchy in target folder 
                    $Filter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName, $SourceSubFolderObject.DisplayName) 
                    $FolderView = New-Object Microsoft.Exchange.WebServices.Data.FolderView(2) 
                    $FindFolderResults = $TargetFolderObject.FindFolders($Filter, $FolderView) 
                    Sleep -Milliseconds $script:currentThrottlingDelay 
                    if ($FindFolderResults.TotalCount -eq 0) 
                    { 
                        $TargetSubFolderObject = New-Object Microsoft.Exchange.WebServices.Data.Folder($script:service) 
                        $TargetSubFolderObject.DisplayName = $SourceSubFolderObject.DisplayName 
                        $TargetSubFolderObject.Save($TargetFolderObject.Id) 
                    } 
                    else 
                    { 
                        $TargetSubFolderObject = $FindFolderResults.Folders[0] 
                    } 
                    MoveItems($SourceSubFolderObject, $TargetSubFolderObject) 
                } 
            } 
        } 
    } 
 
    # If delete parameter is set, check if the source folder is now empty (and if so, delete it) 
    if ($Delete) 
    { 
        $SourceFolderObject.Load() 
        Sleep -Milliseconds $script:currentThrottlingDelay 
        if (($SourceFolderObject.TotalCount -eq 0) -And ($SourceFolderObject.ChildFolderCount -eq 0)) 
        { 
            # Folder is empty, so can be safely deleted 
            try 
            { 
                $SourceFolderObject.Delete([Microsoft.Exchange.Webservices.Data.DeleteMode]::SoftDelete) 
                Sleep -Milliseconds $script:currentThrottlingDelay 
                Log "$($SourceFolderObject.DisplayName) successfully deleted" Green 
            } 
            catch 
            { 
                Log "Failed to delete $($SourceFolderObject.DisplayName)" Red 
            } 
        } 
        else 
        { 
            # Folder is not empty 
            Log "$($SourceFolderObject.DisplayName) could not be deleted as it is not empty." Red 
        } 
    } 
} 
 
Function GetFolder() 
{ 
    # Return a reference to a folder specified by path 
     
    $RootFolder, $FolderPath, $Create = $args[0] 
     
    if ( $RootFolder -eq $null ) 
    { 
        LogVerbose "GetFolder called with null root folder" 
        return $null 
    } 
 
    $Folder = $RootFolder 
    if ($FolderPath -ne '\') 
    { 
        $PathElements = $FolderPath -split '\\' 
        For ($i=0; $i -lt $PathElements.Count; $i++) 
        { 
            if ($PathElements[$i]) 
            { 
                $View = New-Object  Microsoft.Exchange.WebServices.Data.FolderView(2,0) 
                $View.PropertySet = [Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly 
                         
                $SearchFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName, $PathElements[$i]) 
                 
                $FolderResults = $Null 
                try 
                { 
                    $FolderResults = $Folder.FindFolders($SearchFilter, $View) 
                    Sleep -Milliseconds $script:currentThrottlingDelay 
                } 
                catch {} 
                if ($FolderResults -eq $Null) 
                { 
                    if (Throttled) 
                    { 
                    try 
                    { 
                        $FolderResults = $Folder.FindFolders($SearchFilter, $View) 
                    } 
                    catch {} 
                    } 
                } 
                if ($FolderResults -eq $null) 
                { 
                    return $null 
                } 
 
                if ($FolderResults.TotalCount -gt 1) 
                { 
                    # We have more than one folder returned... We shouldn't ever get this, as it means we have duplicate folders 
                    $Folder = $null 
                    Write-Host "Duplicate folders ($($PathElements[$i])) found in path $FolderPath" -ForegroundColor Red 
                    break 
                } 
                elseif ( $FolderResults.TotalCount -eq 0 ) 
                { 
                    if ($Create) 
                    { 
                        # Folder not found, so attempt to create it 
                        $subfolder = New-Object Microsoft.Exchange.WebServices.Data.Folder($script:service) 
                        $subfolder.DisplayName = $PathElements[$i] 
                        try 
                        { 
                            $subfolder.Save($Folder.Id) 
                            LogVerbose "Created folder $($PathElements[$i])" 
                        } 
                        catch 
                        { 
                            # Failed to create the subfolder 
                            $Folder = $null 
                            Log "Failed to create folder $($PathElements[$i]) in path $FolderPath" Red 
                            break 
                        } 
                        $Folder = $subfolder 
                    } 
                    else 
                    { 
                        # Folder doesn't exist 
                        $Folder = $null 
                        Log "Folder $($PathElements[$i]) doesn't exist in path $FolderPath" Red 
                        break 
                    } 
                } 
                else 
                { 
                    $Folder = ThrottledFolderBind $FolderResults.Folders[0].Id 
                } 
            } 
        } 
    } 
     
    $Folder 
} 
 
function GetFolderPath($Folder) 
{ 
    # Return the full path for the given folder 
 
    # We cache our folder lookups for this script 
    if (!$script:folderCache) 
    { 
        # Note that we can't use a PowerShell hash table to build a list of folder Ids, as the hash table is case-insensitive 
        # We use a .Net Dictionary object instead 
        $script:folderCache = New-Object 'System.Collections.Generic.Dictionary[System.String,System.Object]' 
    } 
 
    $propset = New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly, [Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName, [Microsoft.Exchange.WebServices.Data.FolderSchema]::ParentFolderId) 
    $parentFolder = ThrottledFolderBind $Folder.Id $propset 
    $folderPath = $Folder.DisplayName 
    $parentFolderId = $Folder.Id 
    while ($parentFolder.ParentFolderId -ne $parentFolderId) 
    { 
        if ($script:folderCache.ContainsKey($parentFolder.ParentFolderId.UniqueId)) 
        { 
            $parentFolder = $script:folderCache[$parentFolder.ParentFolderId.UniqueId] 
        } 
        else 
        { 
            $parentFolder = ThrottledFolderBind $parentFolder.ParentFolderId $propset 
            $script:FolderCache.Add($parentFolder.Id.UniqueId, $parentFolder) 
        } 
        $folderPath = $parentFolder.DisplayName + "\" + $folderPath 
        $parentFolderId = $parentFolder.Id 
    } 
    return $folderPath 
} 
 
function ConvertId($entryId) 
{ 
    # Use EWS ConvertId function to convert from EntryId to EWS Id 
 
    $id = New-Object Microsoft.Exchange.WebServices.Data.AlternateId 
    $id.Mailbox = $SourceMailbox 
    $id.UniqueId = $entryId 
    $id.Format = [Microsoft.Exchange.WebServices.Data.IdFormat]::EntryId 
    $ewsId = $Null 
    try 
    { 
        $ewsId = $script:service.ConvertId($id, [Microsoft.Exchange.WebServices.Data.IdFormat]::EwsId) 
    } 
    catch {} 
    LogVerbose "EWS Id: $($ewsId.UniqueId)" 
    return $ewsId.UniqueId 
} 
 
function CreateService($targetMailbox) 
{ 
    # Creates and returns an ExchangeService object to be used to access mailboxes 
 
    # First of all check to see if we have a service object for this mailbox already 
    if ($script:services -eq $null) 
    { 
        $script:services = @{} 
    } 
    if ($script:services.ContainsKey($targetMailbox)) 
    { 
        return $script:services[$targetMailbox] 
    } 
 
    # Create new service 
    $exchangeService = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010_SP1) 
 
    # Set credentials if specified, or use logged on user. 
    if ($Credentials -ne $Null) 
    { 
        LogVerbose "Applying given credentials" 
        $exchangeService.Credentials = $Credentials.GetNetworkCredential() 
    } 
    elseif ($Username -and $Password) 
    { 
        LogVerbose "Applying given credentials for $Username" 
        if ($Domain) 
        { 
            $exchangeService.Credentials = New-Object  Microsoft.Exchange.WebServices.Data.WebCredentials($Username,$Password,$Domain) 
        } else { 
            $exchangeService.Credentials = New-Object  Microsoft.Exchange.WebServices.Data.WebCredentials($Username,$Password) 
        } 
    } 
    else 
    { 
        LogVerbose "Using default credentials" 
        $exchangeService.UseDefaultCredentials = $true 
    } 
 
    # Set EWS URL if specified, or use autodiscover if no URL specified. 
    if ($EwsUrl) 
    { 
        $exchangeService.URL = New-Object Uri($EwsUrl) 
    } 
    else 
    { 
        try 
        { 
            LogVerbose "Performing autodiscover for $targetMailbox" 
            if ( $AllowInsecureRedirection ) 
            { 
                $exchangeService.AutodiscoverUrl($targetMailbox, {$True}) 
            } 
            else 
            { 
                $exchangeService.AutodiscoverUrl($targetMailbox) 
            } 
            if ([string]::IsNullOrEmpty($exchangeService.Url)) 
            { 
                Log "$targetMailbox : autodiscover failed" Red 
                return $Null 
            } 
            LogVerbose "EWS Url found: $($exchangeService.Url)" 
        } 
        catch 
        { 
            Log "$targetMailbox : error occurred during autodiscover: $($Error[0])" Red 
            return $null 
        } 
    } 
 
    if ($exchangeService.URL.AbsoluteUri.ToLower().Equals("https://outlook.office365.com/ews/exchange.asmx")) 
    { 
        # This is Office 365, so we'll add a small delay to try and avoid throttling 
        if ($script:currentThrottlingDelay -lt 100) 
        { 
            $script:currentThrottlingDelay = 100 
            LogVerbose "Office 365 mailbox, throttling delay set to $($script:currentThrottlingDelay)ms" 
        } 
    } 
  
    if ($Impersonate) 
    { 
        $exchangeService.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $targetMailbox) 
        $exchangeService.HttpHeaders.Add("X-AnchorMailbox", $targetMailbox) 
    } 
 
    # We enable tracing so that we can retrieve the last response (and read any throttling information from it - this isn't exposed in the EWS Managed API) 
    CreateTraceListener $exchangeService 
    $exchangeService.TraceFlags = [Microsoft.Exchange.WebServices.Data.TraceFlags]::All 
    $exchangeService.TraceEnabled = $True 
 
    $script:services.Add($targetMailbox, $exchangeService) 
    return $exchangeService 
} 
 
function ProcessMailbox() 
{ 
    # Process the mailbox 
 
    $script:currentThrottlingDelay = $ThrottlingDelay # Reset throttling delay 
 
    Write-Host ([string]::Format("Processing mailbox {0}", $SourceMailbox)) -ForegroundColor Gray 
 
    if ( !([String]::IsNullOrEmpty($script:originalLogFile)) ) 
    { 
        $LogFile = $script:originalLogFile.Replace("%mailbox%", $SourceMailbox) 
    } 
 
    $script:service = CreateService($SourceMailbox) 
    if ($script:service -eq $Null) 
    { 
        Write-Host "Failed to create ExchangeService" -ForegroundColor Red 
    } 
     
    # Bind to source mailbox root folder 
    $sourceMbx = New-Object Microsoft.Exchange.WebServices.Data.Mailbox( $SourceMailbox ) 
    if ($SourceArchive) 
    { 
        $folderId = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::ArchiveMsgFolderRoot, $sourceMbx ) 
    } 
    else 
    { 
        $folderId = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot, $sourceMbx ) 
    } 
    $sourceMailboxRoot = ThrottledFolderBind $folderId 
 
    if ( $sourceMailboxRoot -eq $null ) 
    { 
        Write-Host "Failed to open source message store ($SourceMailbox)" -ForegroundColor Red 
        if ($Impersonate) 
        { 
            Write-Host "Please check that you have impersonation permissions" -ForegroundColor Red 
        } 
        return 
    } 
 
    # Bind to target mailbox root folder 
    if ([String]::IsNullOrEmpty($TargetMailbox)) 
    { 
        $TargetMailbox = $SourceMailbox 
    } 
    $targetMbx = New-Object Microsoft.Exchange.WebServices.Data.Mailbox( $TargetMailbox ) 
    if ($TargetArchive) 
    { 
        $folderId = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::ArchiveMsgFolderRoot, $targetMbx ) 
    } 
    else 
    { 
        $folderId = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot, $targetMbx ) 
    } 
 
    $targetMailboxRoot = ThrottledFolderBind $folderId 
    if ( $targetMailboxRoot -eq $null ) 
    { 
        Write-Host "Failed to open target message store ($TargetMailbox)" -ForegroundColor Red 
        return 
    } 
 
    if ($MergeFolderList -eq $Null) 
    { 
        # No folder list, this is a request to move the entire mailbox 
 
        MoveItems($sourceMailboxRoot, $targetMailboxRoot) 
        return 
    } 
 
 
    $MergeFolderList.GetEnumerator() | ForEach-Object { 
        $PrimaryFolder = $_.Name 
        LogVerbose "Target folder is $PrimaryFolder" 
        $SecondaryFolderList = $_.Value 
        LogVerbose "Source folder list is $SecondaryFolderList" 
 
        # Check we can bind to the source folder (if not, stop now) 
        $TargetFolderObject = $null 
        if ($ByFolderId) 
        { 
            $id = New-Object Microsoft.Exchange.WebServices.Data.FolderId($PrimaryFolder) 
            $TargetFolderObject = ThrottledFolderBind $id 
        } 
        elseif ($ByEntryId) 
        { 
            $PrimaryFolderId = ConvertId($PrimaryFolder) 
            $id = New-Object Microsoft.Exchange.WebServices.Data.FolderId($PrimaryFolderId) 
            $TargetFolderObject = ThrottledFolderBind $id 
        } 
        else 
        { 
            $TargetFolderObject = GetFolder($targetMailboxRoot, $PrimaryFolder, $CreateTargetFolder) 
        } 
 
        if ($TargetFolderObject) 
        { 
            # We have the target folder, now check we can get the source folder(s) 
            LogVerbose "Target folder located: $($TargetFolderObject.DisplayName)" 
 
            # Source folder could be a list of folders... 
            $SecondaryFolderList | ForEach-Object { 
                $SecondaryFolder = $_ 
                LogVerbose "Secondary folder is $SecondaryFolder" 
                $SourceFolderObject = $null 
                if ($ByFolderId) 
                { 
                    $id = New-Object Microsoft.Exchange.WebServices.Data.FolderId($SecondaryFolder) 
                    $SourceFolderObject = ThrottledFolderBind $id 
                } 
                elseif ($ByEntryId) 
                { 
                    $SecondaryFolderId = ConvertId($SecondaryFolder) 
                    $id = New-Object Microsoft.Exchange.WebServices.Data.FolderId($SecondaryFolderId) 
                    $SourceFolderObject = ThrottledFolderBind $id 
                } 
                else 
                { 
                    $SourceFolderObject = GetFolder($sourceMailboxRoot, $SecondaryFolder) 
                } 
                if ($SourceFolderObject) 
                { 
                    # Found source folder, now initiate move 
                    LogVerbose "Source folder located: $($SourceFolderObject.DisplayName)" 
                    MoveItems($SourceFolderObject, $TargetFolderObject) 
                } 
                else 
                { 
                    Write-Host "Merge parameters invalid: merge $SecondaryFolder into $PrimaryFolder" -ForegroundColor Red 
                } 
            } 
        } 
        else 
        { 
            Write-Host "Merge parameters invalid: merge $SecondaryFolder into $PrimaryFolder" -ForegroundColor Red 
        } 
    } 
} 
 
 
# The following is the main script 
 
if ($LogFile.Contains("%mailbox%")) 
{ 
    # We replace mailbox marker with the SMTP address of the mailbox - this gives us a log file per mailbox 
    $script:originalLogFile = $LogFile 
    $LogFile = $script:originalLogFile.Replace("%mailbox%", "Merge-MailboxFolder") 
} 
else 
{ 
    $script:originalLogFile = "" 
} 
 
if ( [string]::IsNullOrEmpty($SourceMailbox) ) 
{ 
    $SourceMailbox = CurrentUserPrimarySmtpAddress 
    if ( [string]::IsNullOrEmpty($SourceMailbox) ) 
    { 
        Write-Host "Source mailbox not specified.  Failed to determine current user's SMTP address." -ForegroundColor Red 
        Exit 
    } 
    else 
    { 
        Write-Host ([string]::Format("Current user's SMTP address is {0}", $SourceMailbox)) -ForegroundColor Green 
    } 
} 
 
# Check if we need to ignore any certificate errors 
# This needs to be done *before* the managed API is loaded, otherwise it doesn't work consistently (i.e. usually doesn't!) 
if ($IgnoreSSLCertificate) 
{ 
    Write-Host "WARNING: Ignoring any SSL certificate errors" -foregroundColor Yellow 
    TrustAllCerts 
} 
  
# Load EWS Managed API 
if (!(LoadEWSManagedAPI)) 
{ 
    Write-Host "Failed to locate EWS Managed API, cannot continue" -ForegroundColor Red 
    Exit 
} 
   
# Check we have valid credentials 
if ($Credentials -ne $Null) 
{ 
    If ($Username -or $Password) 
    { 
        Write-Host "Please specify *either* -Credentials *or* -Username and -Password" Red 
        Exit 
    } 
} 
 
# Check whether parameters make sense 
if ($Delete -and $Copy) 
{ 
    Write-Host "Cannot -Delete and -Copy, please use only one of these switches and try again." -ForegroundColor Red 
    exit 
} 
 
if ($MergeFolderList -eq $Null) 
{ 
    # No folder list, this is a request to move the entire mailbox 
    # Check -ProcessSubfolders and -CreateTargetFolder is set, otherwise we fail now (can't move a mailbox without processing subfolders!) 
    if (!$ProcessSubfolders) 
    { 
        Write-Host "Mailbox merge requested, but subfolder processing not specified.  Please retry using -ProcessSubfolders switch." -ForegroundColor Red 
        exit 
    } 
    if (!$CreateTargetFolder) 
    { 
        Write-Host "Mailbox merge requested, but folder creation not allowed.  Please retry using -CreateTargetFolder switch." -ForegroundColor Red 
        exit 
    } 
} 
 
Write-Host "" 
 
# Check whether we have a CSV file as input... 
$FileExists = Test-Path $SourceMailbox 
If ( $FileExists ) 
{ 
    # We have a CSV to process 
    LogVerbose "Reading mailboxes from CSV file" 
    $csv = Import-CSV $SourceMailbox -Header "PrimarySmtpAddress" 
    foreach ($entry in $csv) 
    { 
        LogVerbose $entry.PrimarySmtpAddress 
        if (![String]::IsNullOrEmpty($entry.PrimarySmtpAddress)) 
        { 
            if (!$entry.PrimarySmtpAddress.ToLower().Equals("primarysmtpaddress")) 
            { 
                $SourceMailbox = $entry.PrimarySmtpAddress 
                ProcessMailbox 
            } 
        } 
    } 
} 
Else 
{ 
    # Process as single mailbox 
    ProcessMailbox 
} 
 
if ($script:Tracer -ne $null) 
{ 
    $script:Tracer.Close() 
}