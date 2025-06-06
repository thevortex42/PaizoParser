function PaizoParser {
    #If you want to use the unrefined Session Data, set this to $false
    $EnrichSessions = $true
    
    <#
    If you have your credentials in a file, you can load them here.
    To create such a file, run the line which asks for credentials once (or create the $cred object in any other way),
    chose a location for the save file and then run the following code:
    $cred | Export-Clixml <PathAndFilenameToSaveto>
    The password will be saved as a SecureString in that file.
    That means it is only decryptable when running the script in the same account used to create it.
    THIS IS NOT 100% SECURE! USE AT YOUR OWN RISK!
    #>
    $credFile = $null
        
    if ($null -ne $credFile -and (Test-Path $credFile)) {
        $cred = Import-Clixml $credFile
    }

    if ($null -eq $cred) {
        $cred = Get-Credential -Message "Please enter your Paizo.com email and password"
    }

    $username = $cred.UserName
    $password = $cred.GetNetworkCredential().Password

    $sessionsURL = "https://paizo.com/organizedPlay/myAccount/allsessions"

    #(Down-)load the Assemblies for Browser and Selenium
    Load-BrowserDrivers Edge

    #Create the Options object to maximize the browser window
    $maximized = New-Object OpenQA.Selenium.Edge.EdgeOptions
    $maximized.AddArgument("start-maximized")

    #Start controlled Edge Browser
    $Global:Browser = Create-Browser Edge -options $maximized
    
    #Navigate to the Paizo Homepage
    $Global:Browser.Navigate().GotoURL($sessionsURL)

    #Create Actions object to enter username and password and click the login button
    $actions = [OpenQA.Selenium.Interactions.Actions]::new($Global:Browser)
    $userfield = $Global:Browser.FindElement([OpenQA.Selenium.By]::Name('e'))
    $actions.SendKeys($userfield,$username)
    $pwfield = $Global:Browser.FindElement([OpenQA.Selenium.By]::Name('zzz'))
    $actions.SendKeys($pwfield,$password)
    $actions.SendKeys($pwfield,[OpenQA.Selenium.Keys]::Enter)
    $actions.Build()
    
    #Actually execute those actions
    $actions.Perform()

    #Wait until the browser is done loading by checking if the html element with the ID "tabs" exists (Timeout: 20 seconds)
    $Global:Browser.Manage().Timeouts().ImplicitWait = New-TimeSpan -Seconds 20
    $null = $Global:Browser.FindElement([OpenQA.Selenium.By]::Id('tabs'))

    #Since we are sent back to the basic Organized Play anyways, start by parsing the registered characters
    #Find the area containing information about the characters
    $charPage = $Global:Browser.FindElement([OpenQA.Selenium.By]::className('tp-content'))
    $html = $charPage.getattribute('innerHTML')
    
    #Add headers for the table, since it doesn't contain any
    $Global:Characters = $html -replace '(<table border="0" cellpadding="6" cellspacing="0" width="100%">)','$1<thead><tr><th>CharacterID</th><th>System</th><th>Name</th><th>Reputation</th><th>EditButton></th><th>DeleteButton</th>' | Read-HtmlTable
    $Global:Characters = $Global:Characters | ? {$_.CharacterID -ne "Show Sessions"} | Select CharacterID, System, Name, Reputation
    #change reputation from multiline to comma seperated list
    $Global:Characters | % {$_.Reputation = $_.Reputation -replace "`r`n",", "}

    #Navigate to the site which contains the session list
    $Global:Browser.Navigate().GotoURL($sessionsURL)

    #find the actual list of sessions
    $results = $Global:Browser.FindElement([OpenQA.Selenium.By]::Id('results'))
    $html = $results.getattribute('innerHTML')

    if ($enrichSessions) {
        $SessionData = Parse-SessionData -SessionTable $html
    } else {
        #parse the table, while removing the shown time, which might by things like "today" and replacing them with the actual date, which is included in the html source code
        $foundSessions = $html -replace '<time datetime="(\d{4}-\d{2}-\d{2}).*<\/time>','$1' | Read-HtmlTable -TableIndex 1 | ? {$_.Session -ne $null -and $_.Session -ne "Show Seats"}
    }

    #Set timeout to 0 for faster detection of the last page
    $Global:Browser.Manage().Timeouts().ImplicitWait = 0

    #find out if this was the last (and only) page of sessions
    $ErrorActionPreference = "Stop"
    try {
        $Global:Browser.FindElement([OpenQA.Selenium.By]::LinkText("next >"))
        $nextIsLink = $true
    } catch {
        $nextIsLink = $false
    }
    
    #loop through all other session list pages by clicking "next >" every time
    while ($nextIsLink) {
        $Global:Browser.FindElement([OpenQA.Selenium.By]::LinkText("next >")).click()

        #Set timeout to 20 again to wait for the page to load
        $Global:Browser.Manage().Timeouts().ImplicitWait = New-TimeSpan -Seconds 20
        $results = $Global:Browser.FindElement([OpenQA.Selenium.By]::Id('results'))

        #Wait for a moment after loading. Without that, the data sometime is from the previous page
        Start-Sleep -Seconds 3
        $results = $Global:Browser.FindElement([OpenQA.Selenium.By]::Id('results'))
        $html = $results.getattribute('innerHTML')

        if ($enrichSessions) {
            $SessionData += Parse-SessionData -SessionTable $html
        } else {
            #parse the table, while removing the shown time, which might by things like "today" and replacing them with the actual date, which is included in the html source code
            $foundSessions += $html -replace '<time datetime="(\d{4}-\d{2}-\d{2}).*<\/time>','$1' | Read-HtmlTable -TableIndex 1 | ? {$_.Session -ne $null -and $_.Session -ne "Show Seats"}
        }

        #Set timeout to 0 (see above)
        $Global:Browser.Manage().Timeouts().ImplicitWait = 0
        try {
            $Global:Browser.FindElement([OpenQA.Selenium.By]::LinkText("next >"))
            $nextIsLink = $true
        } catch {
            $nextIsLink = $false
        }
    }

    #We are done parsing, so the browser can be closed
    $Global:Browser.Quit()

    if ($EnrichSessions) {
        $foundSessions = $SessionData
    } else {
        #The event is an array, which is bad for CSV - let us replace that with 'EventID - EventName'
        $foundSessions | % {$_.Event = $_.Event -join " - "}
        #The points (reputation, etc.) are in multiple lines - bad for CSV. Changing that to a comma seperated list
        $foundSessions | % {$_.Points = $_.Points -replace "(`r`n)+",", "}
        #There is some weird last column. We don't want that
        $foundSessions = $foundSessions | Select Date, GM, Scenario, Points, Event, Session, Player, Character, Faction, "Prest. / Rep.", Notes
    }

    #Show the results on screen
    $Global:Characters | Out-GridView
    $foundSessions | Out-GridView
    
    #Get the correct delimiter for the current location
    $delim = (Get-Culture).TextInfo.ListSeparator

    #Save the characters into a csv
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    $SaveCharFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $SaveCharFileDialog.filter = "CSV (*.csv)| *.csv"
    $SaveCharFileDialog.FileName = "OrgPlayCharacters.csv"
    $SaveCharFileDialog.ShowDialog() | Out-Null
    $Global:Characters | Export-Csv $SaveCharFileDialog.Filename -Delimiter $delim -NoTypeInformation -Force

    #remember the path, so we can use it for the next dialog again
    $path = Split-Path $SaveCharFileDialog.FileName -Parent

    #save the sessions into a csv
    $SaveSessionFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $SaveSessionFileDialog.filter = "CSV (*.csv)| *.csv"
    $SaveSessionFileDialog.FileName = "OrgPlaySessions.csv"
    $SaveSessionFileDialog.InitialDirectory = $path
    $SaveSessionFileDialog.ShowDialog() | Out-Null
    $foundSessions | Export-Csv $SaveSessionFileDialog.Filename -Delimiter $delim -NoTypeInformation -Force

}

function Parse-SessionData {
    param(
        [Parameter(mandatory=$true)]$SessionTable
    )

    $SessionData = @()

    $html = $SessionTable
    
    #Get Raw Data, while removing the shown time, which might by things like "today" and replacing them with the actual date, which is included in the html source code
    $RawData = $html -replace '<time datetime="(\d{4}-\d{2}-\d{2}).*<\/time>','$1' | Read-HtmlTable -TableIndex 1 | ? {$null -ne $_.Session -and $_.Session -ne "Show Seats"}


    Foreach ($session in $Rawdata) {
        #Extract the Session Link
        if ($html -match "<a href=""(.*)"" title="".*"">$([Regex]::Escape($session.Scenario))<\/a>") {
            $link = $Matches[1]
        } else {
            $link = $null
        }

        #Divide ID into PlayerID and CharacterID
        $IDs = $session.Player -split "-"

        $enrichedSession = [PSCustomObject]@{
            Date = $session.Date
            Scenario = $session.Scenario
            System = if ($null -ne $session.Character) {$Characters | Where-Object {$_.Name -eq $session.Character}  | Select-Object -ExpandProperty System} else {$null}
            Player = $IDs[0]
            CharacterId = if ($IDs.length -gt 1) {$IDs[1]} elseif ($null -ne $session.Character) {$Characters | Where-Object {$_.Name -eq $session.Character}  | % {($_.CharacterID -split "-")[1]}} else {$null}
            Character = $session.Character
            Faction = $session.Faction
            Reputation = $session.'Prest. / Rep.' -replace "\D",""
            PlayerOrGM = if ($session.'Prest. / Rep.' -match "GM") {'GM'} else {'Player'}
            GM = $session.GM
            Event = $session.Event -join " - "
            Session = $session.Session
            StoreURL = $link
        }
        $SessionData += $enrichedSession
    }

    return $SessionData

}

function Load-NugetAssembly {
    [CmdletBinding()]
    param(
        [string]$url,
        [string]$name,
        [string]$zipinternalpath,
        [switch]$downloadonly
    )
    if($psscriptroot -ne ''){      
        $localpath = join-path $psscriptroot $name
    }else{
        $localpath = join-path $env:TEMP $name
    }
    $tmp = "$env:TEMP\$([IO.Path]::GetRandomFileName())"      
    $zip = $null
    try{
        if(!(Test-Path $localpath)){
            Add-Type -A System.IO.Compression.FileSystem
            write-host "Downloading and extracting required library '$name' ... " -F Green -NoNewline      
            (New-Object System.Net.WebClient).DownloadFile($url, $tmp)
            $zip = [System.IO.Compression.ZipFile]::OpenRead($tmp)
            $zip.Entries | ?{$_.Fullname -eq $zipinternalpath} | %{
                [System.IO.Compression.ZipFileExtensions]::ExtractToFile($_,$localpath)
            }
            write-host "OK" -F Green  
        }
        if (Get-Item $localpath -Stream zone.identifier -ea SilentlyContinue){
            Unblock-File -Path $localpath
        }
        if(!$downloadonly.IsPresent){
            Add-Type -Path $localpath -EA Stop
        }
            
    }catch{
        throw "Error: $($_.Exception.Message)"      
    }finally{
        if ($zip){$zip.Dispose()}
        if(Test-Path $tmp){del $tmp -Force -EA 0}
    }
}

function Load-BrowserDrivers {
    param(
        [Parameter(mandatory=$true)][ValidateSet('Chrome','Edge','Firefox')][string]$Global:Browser,
        [Parameter(mandatory=$false)][string]$driverversion = ''
    )

    # Load Selenium Webdriver .NET Assembly and dependencies
    Load-NugetAssembly 'https://www.nuget.org/api/v2/package/Newtonsoft.Json' -name 'Newtonsoft.Json.dll' -zipinternalpath 'lib/net45/Newtonsoft.Json.dll' -EA Stop    
    Load-NugetAssembly 'https://www.nuget.org/api/v2/package/Selenium.WebDriver/4.23.0' -name 'WebDriver.dll' -zipinternalpath 'lib/netstandard2.0/WebDriver.dll' -EA Stop    
    
    switch($Global:Browser){
        'Chrome' {      
            $chrome = Get-Package -Name 'Google Chrome' -EA SilentlyContinue | select -F 1      
            if (!$chrome){
                throw "Google Chrome Browser not installed."      
                return
            }
            Load-NugetAssembly "https://www.nuget.org/api/v2/package/Selenium.WebDriver.ChromeDriver/$driverversion" -name 'chromedriver.exe' -zipinternalpath 'driver/win32/chromedriver.exe' -downloadonly -EA Stop      
        }
        'Edge' {      
            $edge = Get-Package -Name 'Microsoft Edge' -EA SilentlyContinue | select -F 1      
            if (!$edge){
                throw "Microsoft Edge Browser not installed."      
                return
            }
            Load-NugetAssembly "https://www.nuget.org/api/v2/package/Selenium.WebDriver.MSEdgeDriver.win32/$driverversion" -name 'msedgedriver.exe' -zipinternalpath 'driver/win32/msedgedriver.exe' -downloadonly -EA Stop      
        }
        'Firefox' {      
            $ff = Get-Package -Name "Mozilla Firefox*" -EA SilentlyContinue | select -F 1      
            if (!$ff){
                throw "Mozilla Firefox Browser not installed."      
                return
            }
            Load-NugetAssembly "https://www.nuget.org/api/v2/package/Selenium.WebDriver.GeckoDriver/$driverversion" -name 'geckodriver.exe' -zipinternalpath 'driver/win64/geckodriver.exe' -downloadonly -EA Stop      
        }
    }
}

function Create-Browser {
    param(
        [Parameter(mandatory=$true)][ValidateSet('Chrome','Edge','Firefox')][string]$Global:Browser,      
        [Parameter(mandatory=$false)][bool]$HideCommandPrompt = $true,
        [Parameter(mandatory=$false)][object]$options = $null
    )
    $driver = $null

    if($psscriptroot -ne ''){      
        $driverpath = $psscriptroot
    }else{
        $driverpath = $env:TEMP
    }
    switch($Global:Browser){
        'Chrome' {      
            # create driver service
            $dService = [OpenQA.Selenium.Chrome.ChromeDriverService]::CreateDefaultService($driverpath)
            # hide command prompt window
            $dService.HideCommandPromptWindow = $HideCommandPrompt
            # create driver object
            if ($options){
                $driver = New-Object OpenQA.Selenium.Chrome.ChromeDriver $dService,$options
            }else{
                $driver = New-Object OpenQA.Selenium.Chrome.ChromeDriver $dService
            }
        }
        'Edge' {      
            # create driver service
            $dService = [OpenQA.Selenium.Edge.EdgeDriverService]::CreateDefaultService($driverpath)
            # hide command prompt window
            $dService.HideCommandPromptWindow = $HideCommandPrompt
            # create driver object
            if ($options){
                $driver = New-Object OpenQA.Selenium.Edge.EdgeDriver $dService,$options
            }else{
                $driver = New-Object OpenQA.Selenium.Edge.EdgeDriver $dService
            }
        }
        'Firefox' {      
            # create driver service
            $dService = [OpenQA.Selenium.Firefox.FirefoxDriverService]::CreateDefaultService($driverpath)
            # hide command prompt window
            $dService.HideCommandPromptWindow = $HideCommandPrompt
            # create driver object
            if($options){
                $driver = New-Object OpenQA.Selenium.Firefox.FirefoxDriver $dService, $options
            }else{
                $driver = New-Object OpenQA.Selenium.Firefox.FirefoxDriver $dService
            }
        }
    }
    return $driver
}

function Read-HtmlTable {
    <#PSScriptInfo
    .VERSION 2.0.1
    .GUID 6ddb4b24-29bc-4268-a62f-402b3ee28e3d
    .AUTHOR iRon
    .COMPANYNAME
    .COPYRIGHT
    .TAGS Read Extract Scrape ConvertFrom Html Table
    .LICENSE https://github.com/iRon7/Read-HtmlTable/LICENSE
    .PROJECTURI https://github.com/iRon7/Read-HtmlTable
    .ICON https://raw.githubusercontent.com/iRon7/Read-HtmlTable/main/Read-HtmlTable.png
    .EXTERNALMODULEDEPENDENCIES
    .REQUIREDSCRIPTS
    .EXTERNALSCRIPTDEPENDENCIES
    .RELEASENOTES
    .PRIVATEDATA
    #>

    <#
    .SYNOPSIS
    Reads a HTML table

    .DESCRIPTION
    Scrapes (extracts) a HTML table from a string or the internet location

    .INPUTS
    String or Uri

    .OUTPUTS
    PSCustomObject[]

    .PARAMETER InputObject
        The html content (string) that contains a html table.

        If the string is less than 2048 characters and contains a valid uri protocol, the content is downloaded
        from the concerned location.

    .PARAMETER Uri
        A uri location referring to the html content that contains the html table

    .PARAMETER Header
        Specifies an alternate column header row for the imported string. The column header determines the property
        names of the objects created by ConvertFrom-Csv.

        Enter column headers as a comma-separated list. Do not enclose the header string in quotation marks.
        Enclose each column header in single quotation marks.

        If you enter fewer column headers than there are data columns, the remaining data columns are discarded.
        If you enter more column headers than there are data columns, the additional column headers are created
        with empty data columns.

        A $Null instead of a column name, will span the respective column with previous column.

        Note: To select specific columns or skip any data (or header) rows, use Select-Object cmdlet

    .PARAMETER TableIndex
        Specifies which tables should be selected from the html content (where 0 refers to the first table).
        By default, all tables are extracted from the content.

        Note: in case of multiple tables, the headers should be unified to properly output or display of each table.
        (see: https://github.com/PowerShell/PowerShell/issues/13906)

    .PARAMETER Separator
        Specifies the characters used to join a header with is spanned over multiple columns.
        (default: space character)

    .PARAMETER Delimiter
        Specifies the characters used to join a header with is spanned over multiple rows.
        (default: the newline characters used by the operating system)

    .PARAMETER NoTrim
        By default, all header - and data text is trimmed, to disable trimming, use the -NoTrim parameter.

    .EXAMPLE

        Read-HTMLTable https://github.com/iRon7/Read-HtmlTable

        Product            Invoice           Invoice    Invoice
        Item               Qauntity          @          Price
        -------------      ----------------- ---------- --------------
        Paperclips (Box)   100               1.15       115.00
        Paper (Case)       10                45.99      459.90
        Wastepaper Baskets 10                17.99      35.98
        Subtotal           Subtotal          Subtotal   610.88
        Tax                Tax               7%         42.76
        Total              Total             Total      653.64

    .LINK
        https://github.com/iRon7/Read-HtmlTable
    #>
    [CmdletBinding(DefaultParameterSetName='Html')][OutputType([Object[]])] param(
        [Parameter(ParameterSetName='Html', ValueFromPipeLine = $True, Mandatory = $True, Position = 0)][String]$InputObject,
        [Parameter(ParameterSetName='Uri', ValueFromPipeLine = $True, Mandatory = $True)][Uri]$Uri,
        [Object[]]$Header,
        [Int[]]$TableIndex,
        [String]$Separator = ' ',
        [String]$Delimiter = [System.Environment]::NewLine,
        [Switch]$NoTrim
    )
    Begin {
        function ParseHtml($String) {
            $Unicode = [System.Text.Encoding]::Unicode.GetBytes($String)
            $Html = New-Object -Com 'HTMLFile'
            if ($Html.PSObject.Methods.Name -Contains 'IHTMLDocument2_Write') { $Html.IHTMLDocument2_Write($Unicode) } else { $Html.write($Unicode) }
            $Html.Close()
            $Html
        }
        filter GetTopElement([String[]]$TagName) {
            if ($TagName -Contains $_.tagName) { $_}
            else { @($_.Children).Where{ $_ } | GetTopElement -TagName $TagName }
        }
        function GetUnit($Data, [int]$x, [int]$y) {
            if ($x -lt $Data.Count -and $y -lt $Data[$x].Count) { $Data[$x][$y] }
        }
        function SetUnit($Data, [int]$x, [int]$y, [HashTable]$Unit) {
            while ($x -ge $Data.Count) { $Data.Add([System.Collections.Generic.List[HashTable]]::new()) }
            while ($y -ge $Data[$x].Count) { $Data[$x].Add($Null) }
            $Data[$x][$y] = $Unit
        }
        function GetData([__ComObject[]]$TRs) {
            $Data = [System.Collections.Generic.List[System.Collections.Generic.List[HashTable]]]::new()
            $y = 0
            foreach($TR in $TRs) {
                $x = 0
                foreach($TD in ($TR |GetTopElement 'th', 'td')) {
                    while ($True) { # Skip any row spans
                        $Unit = GetUnit -Data $Data -x $x -y $y
                        if (!$Unit) { break }
                        $x++
                    }
                    $Text = if ($Null -ne $TD.innerText) { if ($NoTrim) { $TD.innerText } else { $TD.innerText.Trim() } }
                    for ($r = 0; $r -lt $TD.rowspan; $r++) {
                        $y1 = $y + $r
                        for ($c = 0; $c -lt $TD.colspan; $c++) {
                            $x1 = $x + $c
                            $Unit = GetUnit -Data $Data -x $x1 -y $y1
                            if ($Unit) { SetUnit -Data $Data -x $x1 -y $y1 -Unit @{ ColSpan = $True; Text = $Unit.Text, $Text } } # RowSpan/ColSpan overlap
                            else { SetUnit -Data $Data -x $x1 -y $y1 -Unit @{ ColSpan = $c -gt 0; RowSpan = $r -gt 0; Text = $Text } }
                        }
                    }
                    $x++
                }
                $y++
            }
            ,$Data
        }
    }
    Process {
        if (!$Uri -and $InputObject.Length -le 2048 -and ([Uri]$InputObject).AbsoluteUri) { $Uri = [Uri]$InputObject }
        $Response = if ($Uri -is [Uri] -and $Uri.AbsoluteUri) { Try { Invoke-WebRequest $Uri } Catch { Throw $_ } }
        $Html = if ($Response) { ParseHtml $Response.RawContent } else { ParseHtml $InputObject }
        $i = 0
        foreach($Table in ($Html.Body |GetTopElement 'table')) {
            if (!$PSBoundParameters.ContainsKey('TableIndex') -or $i++ -In $TableIndex) {
                $Rows = $Table |GetTopElement 'tr'
                if (!$Rows) { return }
                if ($PSBoundParameters.ContainsKey('Header')) {
                    $HeadRows = @()
                    $Data = GetData $Rows
                }
                else {
                    for ($i = 0; $i -lt $Rows.Count; $i++) { $Rows[$i].id = "id_$i" }
                    $THead = $Table |GetTopElement 'thead'
                    $HeadRows = @(
                        if ($THead) { $THead |GetTopElement 'tr' }
                        else { $Rows.Where({ !($_ |GetTopElement 'th') }, 'Until' ) }
                    )
                    if (!$HeadRows -or $HeadRows.Count -eq $Rows.Count) { $HeadRows = $Rows[0] }
                    $Head = GetData $HeadRows
                    $Data = GetData ($Rows.Where{ $_.id -notin $HeadRows.id })
                    $Header = @(
                        for ($x = 0; $x -lt $Head.Count; $x++) {
                            if ($Head[$x].Where({ !$_.ColSpan }, 'First') ) {
                                ,@($Head[$x].Where{ !$_.RowSpan }.ForEach{ $_.Text })
                            }
                            else { $Null } # aka spanned header column
                        }
                        for ($x = $Head.Count; $x -lt $Data.Count; $x++) {
                            if ($Null -ne $Data[$x].Where({ $_ -and !$_.ColSpan }, 'First') ) { '' }
                        }
                    )
                }
                $Header = $Header.ForEach{
                    if ($Null -eq $_) { $Null }
                    else {
                        $Name = [String[]]$_
                        $Name = if ($NoTrim) { $Name -Join $Delimiter }
                                else { (($Name.ForEach{ $_.Trim() }) -Join $Delimiter).Trim() }
                        if ($Name) { $Name } else { '1' }
                    }
                }
                $Unique = [System.Collections.Generic.HashSet[String]]::new([StringComparer]::InvariantCultureIgnoreCase)
                $Duplicates = @( for ($i = 0; $i -lt $Header.Count; $i++) { if ($Null -ne $Header[$i] -and !$Unique.Add($Header[$i])) { $i } } )
                $Duplicates.ForEach{
                    do {
                        $Name, $Number = ([Regex]::Match($Header[$_], '^([\s\S]*?)(\d*)$$')).Groups.Value[1, 2]
                        $Digits = '0' * $Number.Length
                        $Header[$_] = "$Name{0:$Digits}" -f (1 + $Number)
                    } while (!$Unique.Add($Header[$_]))
                }
                for ($y = 0; $y -lt ($Data |ForEach-Object Count |Measure-Object -Maximum).Maximum; $y++) {
                    $Name = $Null # (custom) -Header parameter started with a spanned ($Null) column
                    $Properties = [ordered]@{}
                    for ($x = 0; $x -lt $Header.Count; $x++) {
                        $Unit = GetUnit -Data $Data -x $x -y $y -Unit
                        if ($Null -ne $Header[$x]) {
                            $Name = $Header[$x]
                            $Properties[$Name] = if ($Unit) { $Unit.Text } # else $Null (align column overflow)
                        }
                        elseif ($Name -and !$Unit.ColSpan) {
                            $Properties[$Name] = $Properties[$Name], $Unit.Text
                        }
                    }
                    [pscustomobject]$Properties
                }
            }
        }
        $Null = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Html)
    }
}

#Declare Variables that will be used by more than one function
if (-not (Get-Variable Characters -Scope Global -ErrorAction Ignore)) {
    New-Variable -Name Characters -Scope Global
}
if (-not (Get-Variable Characters -Scope Global -ErrorAction Ignore)) {
    New-Variable -Name Browser -Scope Global
}

#Call the main function
& PaizoParser