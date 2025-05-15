function Create-Browser {
    param(
        [Parameter(mandatory=$true)][ValidateSet('Chrome','Edge','Firefox')][string]$browser,      
        [Parameter(mandatory=$false)][bool]$HideCommandPrompt = $true,
        [Parameter(mandatory=$false)][string]$driverversion = '',      
        [Parameter(mandatory=$false)][object]$options = $null
    )
    $driver = $null

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

    # Load Selenium Webdriver .NET Assembly and dependencies
    Load-NugetAssembly 'https://www.nuget.org/api/v2/package/Newtonsoft.Json' -name 'Newtonsoft.Json.dll' -zipinternalpath 'lib/net45/Newtonsoft.Json.dll' -EA Stop    
    Load-NugetAssembly 'https://www.nuget.org/api/v2/package/Selenium.WebDriver/4.23.0' -name 'WebDriver.dll' -zipinternalpath 'lib/netstandard2.0/WebDriver.dll' -EA Stop    
    
    if($psscriptroot -ne ''){      
        $driverpath = $psscriptroot
    }else{
        $driverpath = $env:TEMP
    }
    switch($browser){
        'Chrome' {      
            $chrome = Get-Package -Name 'Google Chrome' -EA SilentlyContinue | select -F 1      
            if (!$chrome){
                throw "Google Chrome Browser not installed."      
                return
            }
            Load-NugetAssembly "https://www.nuget.org/api/v2/package/Selenium.WebDriver.ChromeDriver/$driverversion" -name 'chromedriver.exe' -zipinternalpath 'driver/win32/chromedriver.exe' -downloadonly -EA Stop      
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
            $edge = Get-Package -Name 'Microsoft Edge' -EA SilentlyContinue | select -F 1      
            if (!$edge){
                throw "Microsoft Edge Browser not installed."      
                return
            }
            Load-NugetAssembly "https://www.nuget.org/api/v2/package/Selenium.WebDriver.MSEdgeDriver.win32/$driverversion" -name 'msedgedriver.exe' -zipinternalpath 'driver/win32/msedgedriver.exe' -downloadonly -EA Stop      
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
            $ff = Get-Package -Name "Mozilla Firefox*" -EA SilentlyContinue | select -F 1      
            if (!$ff){
                throw "Mozilla Firefox Browser not installed."      
                return
            }
            Load-NugetAssembly "https://www.nuget.org/api/v2/package/Selenium.WebDriver.GeckoDriver/$driverversion" -name 'geckodriver.exe' -zipinternalpath 'driver/win64/geckodriver.exe' -downloadonly -EA Stop      
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

$cred = Get-Credential -Message "Bitte Paizo E-Mail-Adresse und Kennwort eingeben"

$username = $cred.UserName
$password = $cred.GetNetworkCredential().Password

$baseURL = "https://paizo.com"
$sessionsURL = "https://paizo.com/organizedPlay/myAccount/allsessions"

#Einmal Browser starten, damit die Assemblies geladen werden
$browser = Create-Browser Edge
$browser.Close()

$maximized = New-Object OpenQA.Selenium.Edge.EdgeOptions
$maximized.AddArgument("start-maximized")

$browser = Create-Browser Edge -options $maximized
$browser.Navigate().GotoURL($sessionsURL)

$actions = [OpenQA.Selenium.Interactions.Actions]::new($browser)

$userfield = $browser.FindElement([OpenQA.Selenium.By]::Name('e'))
$actions.SendKeys($userfield,$username)
$pwfield = $browser.FindElement([OpenQA.Selenium.By]::Name('zzz'))
$actions.SendKeys($pwfield,$password)
$actions.SendKeys($pwfield,[OpenQA.Selenium.Keys]::Enter)
$actions.Build()
$actions.Perform()

#Warten, bis der Browser fertig geladen hat

$browser.Manage().Timeouts().ImplicitWait = New-TimeSpan -Seconds 20
$null = $browser.FindElement([OpenQA.Selenium.By]::Id('tabs'))

#Charaktere auslesen
$charPage = $browser.FindElement([OpenQA.Selenium.By]::className('tp-content'))
$html = $charPage.getattribute('innerHTML')
$characters = $html -replace '(<table border="0" cellpadding="6" cellspacing="0" width="100%">)','$1<thead><tr><th>CharacterID</th><th>System</th><th>Name</th><th>Reputation</th><th>EditButton></th><th>DeleteButton</th>' | Read-HtmlTable
$characters = $characters | ? {$_.CharacterID -ne "Show Sessions"} | Select CharacterID, System, Name, Reputation
$characters | % {$_.Reputation = $_.Reputation -replace "`r`n",", "}

#Wechseln auf die Session-Seite
$browser.Navigate().GotoURL($sessionsURL)

#die Ergebnisliste auslesen
$results = $browser.FindElement([OpenQA.Selenium.By]::Id('results'))
$html = $results.getattribute('innerHTML')

#Sessions in Variable speichern
$foundSessions = $html -replace '<time datetime="(\d{4}-\d{2}-\d{2}).*<\/time>','$1' | Read-HtmlTable -TableIndex 1 | ? {$_.Session -ne $null -and $_.Session -ne "Show Seats"}

#Timeout wieder auf 0 setzen
$browser.Manage().Timeouts().ImplicitWait = 0

#Weitere Seiten oeffnen
$ErrorActionPreference = "Stop"
try {
    $browser.FindElement([OpenQA.Selenium.By]::LinkText("next >"))
    $nextIsLink = $true
} catch {
    $nextIsLink = $false
}


while ($nextIsLink) {
    $browser.FindElement([OpenQA.Selenium.By]::LinkText("next >")).click()

    #Timeout wieder auf 20 Sekunden setzen und warten, bis Seite geladen ist
    $browser.Manage().Timeouts().ImplicitWait = New-TimeSpan -Seconds 20
    $results = $browser.FindElement([OpenQA.Selenium.By]::Id('results'))

    #Kurz warten, da sonst nicht die richtigen Daten ausgelesen werden
    Start-Sleep -Seconds 3
    $results = $browser.FindElement([OpenQA.Selenium.By]::Id('results'))
    $html = $results.getattribute('innerHTML')

    #Sessions in Variable speichern
    $foundSessions += $html -replace '<time datetime="(\d{4}-\d{2}-\d{2}).*<\/time>','$1' | Read-HtmlTable -TableIndex 1 | ? {$_.Session -ne $null -and $_.Session -ne "Show Seats"}

    #Timeout wieder auf 0 setzen
    $browser.Manage().Timeouts().ImplicitWait = 0
    try {
        $browser.FindElement([OpenQA.Selenium.By]::LinkText("next >"))
        $nextIsLink = $true
    } catch {
        $nextIsLink = $false
    }
}

$browser.Close()

#Event von Array zu Text konvertieren
$foundSessions | % {$_.Event = $_.Event -join " - "}
$foundSessions | % {$_.Points = $_.Points -replace "(`r`n)+",", "}
$foundSessions = $foundSessions | Select Date, GM, Scenario, Points, Event, Session, Player, Character, Faction, "Prest. / Rep.", Notes

#Anzeige der Ergebnisse am Bildschirm
$characters | Out-GridView
$foundSessions | Out-GridView

#Abspeichern #2 - Sessions
[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
$SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
$SaveFileDialog.filter = "CSV (*.csv)| *.csv"
$SaveFileDialog.FileName = "OrgPlayCharacters.csv"
$SaveFileDialog.ShowDialog() | Out-Null

$characters | Export-Csv $SaveFileDialog.Filename -Delimiter ";" -NoTypeInformation -Force

$path = Split-Path $SaveFileDialog.FileName -Parent

#Abspeichern #2 - Sessions
[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
$SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
$SaveFileDialog.filter = "CSV (*.csv)| *.csv"
$SaveFileDialog.FileName = "OrgPlaySessions.csv"
$SaveFileDialog.InitialDirectory = $path
$SaveFileDialog.ShowDialog() | Out-Null

$foundSessions | Export-Csv $SaveFileDialog.Filename -Delimiter ";" -NoTypeInformation -Force