<#
.SYNOPSIS
The Master FileNet Stored Search Toolkit (N-Dimensional & Modular Edition)
(Fixed: FileNet CPE <in>/<list> parsing error by using strict binary <or><eq> staircases for Broad Catch)
#>

#region 1. CONFIGURATIE & BRONDATA
$GlobalConfig = @{
    SearchName      = "ZO3 Elwin Master"          
    ObjectType      = "document"
    ClassName       = "document"           
    SubclassSymName = "XXXDocument"        
    RepositoryId    = "demorepo1"          
    OutputDir       = "$env:HOMEPATH\Desktop\SearchExport"
}

if (-not (Test-Path $GlobalConfig.OutputDir)) { New-Item -ItemType Directory -Path $GlobalConfig.OutputDir | Out-Null }

$PropertyDefs = @(
    @{ SymName = "BSN";                Name = "BSN";               DataType = "xs:string";    EngineType = "typestring" }
    @{ SymName = "DateCreated";        Name = "DateCreated";       DataType = "xs:timestamp"; EngineType = "typedate" }
    @{ SymName = "XXXAandachtsgebied"; Name = "Aandachtsgebied";   DataType = "xs:string";    EngineType = "typestring" }
    @{ SymName = "XXXOnderwerp";       Name = "Onderwerp";         DataType = "xs:string";    EngineType = "typestring" }
    @{ SymName = "XXXSubonderwerp";    Name = "Subonderwerp";      DataType = "xs:string";    EngineType = "typestring" }
)
$global:PropLookup = @{}
$PropertyDefs | ForEach-Object { $global:PropLookup[$_.SymName] = $_ }

# De Zichtbare Filters voor de UI (De basisgroep)
$GlobalFilters = @(
    @{ SymName = "BSN";         Operator = "EQUAL";          Values = @("", "") }
    @{ SymName = "DateCreated"; Operator = "GREATEROREQUAL"; Values = @("", "") }
    @{ SymName = "DateCreated"; Operator = "LESSOREQUAL";    Values = @("", "") }
)

# Definitie van de combinaties
$CombinationFields = @("XXXAandachtsgebied", "XXXOnderwerp", "XXXSubonderwerp")

$ExcelData = @"
WIA;Vaststellen recht, hoogte, duur;Wettelijke rente
WAZ;Vaststellen recht, hoogte, duur;Wettelijke rente
WAO;Vaststellen recht, hoogte, duur;Wettelijke rente
WAJONG;Aanvraag/Intake;Bezwaar en Beroep
WAO;Aanvraag/Intake;Bezwaar en Beroep
WIA;Aanvraag/Intake;Bezwaar en Beroep
"@
#endregion


#region 2. CALCULATOR (CARTESIAN BLACKLIST)
function Get-CartesianBlacklist {
    Write-Host "`n[*] Berekent N-Dimensionale Blacklist..." -ForegroundColor Cyan
    $rawRows = $ExcelData -split "`n" | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    
    $global:ValidCombos = New-Object System.Collections.Generic.HashSet[string]
    $global:UniqueValues = @{}
    foreach ($f in $CombinationFields) { $global:UniqueValues[$f] = @() }

    $global:ParsedRows = @()
    foreach ($row in $rawRows) {
        $parts = $row.Trim() -split ";"
        $keyParts = @(); $rowObj = @{}
        for ($i = 0; $i -lt $CombinationFields.Count; $i++) {
            $fName = $CombinationFields[$i]; $val = $parts[$i]
            $keyParts += $val; $rowObj[$fName] = $val
            if ($val -notin $global:UniqueValues[$fName]) { $global:UniqueValues[$fName] += $val }
        }
        $global:ParsedRows += $rowObj
        $global:ValidCombos.Add(($keyParts -join "||")) | Out-Null
    }

    $combinations = @( @{} )
    foreach ($fName in $CombinationFields) {
        $newCombos = @()
        foreach ($val in $global:UniqueValues[$fName]) {
            foreach ($c in $combinations) {
                $nc = $c.Clone(); $nc[$fName] = $val; $newCombos += $nc
            }
        }
        $combinations = $newCombos
    }

    $global:InvalidCombos = @()
    foreach ($c in $combinations) {
        $keyParts = @()
        foreach ($fName in $CombinationFields) { $keyParts += $c[$fName] }
        if (-not $global:ValidCombos.Contains(($keyParts -join "||"))) { $global:InvalidCombos += $c }
    }
    Write-Host "    Geldig: $($global:ValidCombos.Count) | Uitgesloten: $($global:InvalidCombos.Count)"
}
#endregion


#region 3. GENERATOR HELPERS (XML & JSON)
$global:xmlItemId = 1000 

function Build-XmlStaircase([array]$nodes, [string]$opTag) {
    if ($nodes.Count -eq 0) { return "" }; if ($nodes.Count -eq 1) { return $nodes[0] } 
    $first = $nodes[0]
    if ($nodes.Count -eq 2) { return "<$opTag>$first$($nodes[1])</$opTag>" }
    $rest = Build-XmlStaircase $nodes[1..($nodes.Count - 1)] $opTag
    return "<$opTag>$first$rest</$opTag>"
}

function Get-XmlNode($Sym, $Op, $Val, $isHidden) {
    $def = $global:PropLookup[$Sym]; $datatype = $def.EngineType; $objType = $GlobalConfig.ObjectType
    $editprop = if ($isHidden) { "hidden" } else { "editable" }
    
    $xmlOp = "eq"; if ($Op -eq "GREATEROREQUAL") { $xmlOp = "gte" }; if ($Op -eq "LESSOREQUAL") { $xmlOp = "lte" }; if ($Op -eq "NOTEQUAL") { $xmlOp = "neq" }
    if ($datatype -eq "typedate" -and $Val -match "^(\d{4}-\d{2}-\d{2})T.*([+-]\d{2}:\d{2})$") { $Val = $matches[1] + $matches[2] }
    
    $id = $global:xmlItemId++
    return "<$xmlOp><whereprop editproperty=`"$editprop`" name=`"$($def.Name)`" objecttype=`"$objType`" symname=`"$Sym`" itemid=`"$id`"><propdesc datatype=`"$datatype`" haschoices=`"false`" hasmarkings=`"false`" symname=`"$Sym`" /></whereprop><literal><![CDATA[$Val]]></literal></$xmlOp>"
}

function Get-JsonNode($Sym, $Op, $Val, $isHidden) {
    $def = $global:PropLookup[$Sym]
    return [ordered]@{
        itemId = ""; displayValues = @(); hasSelectedOperators = $false; hidden = $isHidden
        dataType = $def.DataType; values = @($Val, ""); name = $def.Name
        selectedOperator = $Op; readOnly = $false; id = $Sym; cardinality = "SINGLE"; required = $false
    }
}

function Build-SearchWrapperXml([string]$whereClause) {
    $osId = if ($null -ne $fnp8os) { $fnp8os.Id.ToString() } else { "{00000000-0000-0000-0000-000000000000}" }
    $osName = if ($null -ne $fnp8os) { $fnp8os.Name } else { "UNKNOWN_OS" }
    $scId = $global:xmlItemId++

    $raw = @"
<?xml version="1.0" encoding="UTF-8"?>
<storedsearch xmlns="http://filenet.com/namespaces/wcm/apps/1.0"><version dtd="3.0" searchobject="3" /><product name="Search Designer" version="3.0" /><searchspec versionselection="allversions"><searchtype><template showandorconditions="false" showmaxrecords="true" showoperators="true" /></searchtype><objectstores mergeoption="union"><objectstore id="$osId" name="$osName" /></objectstores><searchcriteria><folders/><searchclauses><searchclause join="and"><select><selectprops>
<selectprop alignment="left" itemid="1" name="DocumentTitle" objecttype="document" sortlevel="0" sortorder="none" symname="DocumentTitle" />
<selectprop alignment="left" itemid="2" name="DateCreated" objecttype="document" sortlevel="1" sortorder="descending" symname="DateCreated" />
<selectprop alignment="left" itemid="3" name="Id" objecttype="document" sortlevel="0" sortorder="none" symname="Id" />
</selectprops></select><from><class symname="$($GlobalConfig.ClassName)" /></from><where>$whereClause</where><subclasses><subclass editproperty="hidden" itemid="$scId" name="$($GlobalConfig.SubclassSymName)" objecttype="$($GlobalConfig.ObjectType)" symname="$($GlobalConfig.SubclassSymName)" includesubclasses="true"/></subclasses></searchclause></searchclauses></searchcriteria><options maxrecords="200" /></searchspec></storedsearch>
"@
    return $raw -replace '>\s+<', '><' -replace "`r|`n", ""
}
#endregion


#region 4. EXPORTERS (XT & ICN & SQL)
function Export-WorkplaceXTFiles {
    Write-Host "`n[*] Genereert XT Bestanden..." -ForegroundColor Cyan
    $global:xmlItemId = 1000 
    
    # UIContent (Alleen Globals)
    $uiBlocks = @()
    foreach ($g in $GlobalFilters) { $uiBlocks += Get-XmlNode $g.SymName $g.Operator $g.Values[0] $false }
    [System.IO.File]::WriteAllText("$($GlobalConfig.OutputDir)\UIContent.xml", (Build-SearchWrapperXml (Build-XmlStaircase $uiBlocks "and")), [System.Text.Encoding]::UTF8)

    # Content (Alles)
    $cBlocks = @()
    foreach ($g in $GlobalFilters) { $cBlocks += Get-XmlNode $g.SymName $g.Operator $g.Values[0] $false }
    
    # FIX: Broad Catch via Binaire OR/EQ traps in plaats van <in><list>
    foreach ($fName in $CombinationFields) {
        $broadCatchNodes = @()
        foreach ($v in $global:UniqueValues[$fName]) {
            $broadCatchNodes += Get-XmlNode $fName "EQUAL" $v $true
        }
        if ($broadCatchNodes.Count -gt 0) { $cBlocks += Build-XmlStaircase $broadCatchNodes "or" }
    }

    $blacklistXml = @()
    foreach ($inv in $global:InvalidCombos) {
        $blNodes = @()
        foreach ($fName in $CombinationFields) { $blNodes += Get-XmlNode $fName "NOTEQUAL" $inv[$fName] $true }
        $blacklistXml += Build-XmlStaircase $blNodes "or"
    }
    if ($blacklistXml.Count -gt 0) { $cBlocks += Build-XmlStaircase $blacklistXml "and" }
    
    [System.IO.File]::WriteAllText("$($GlobalConfig.OutputDir)\Content.xml", (Build-SearchWrapperXml (Build-XmlStaircase $cBlocks "and")), [System.Text.Encoding]::UTF8)
    Write-Host "    Gereed (Content.xml & UIContent.xml)."
}

function Export-ICNFiles {
    Write-Host "`n[*] Genereert ICN Split-Brain Bestanden..." -ForegroundColor Cyan
    $global:xmlItemId = 1000
    
    $jsonCriteria = @() # Alleen Globals voor UI stabiliteit
    $xmlBlocks = @()    # Volledige "Russian Doll" logica voor Backend

    # 1. Globals (Gaan naar UI en Backend)
    foreach ($g in $GlobalFilters) {
        $jsonCriteria += Get-JsonNode $g.SymName $g.Operator $g.Values[0] $false
        $xmlBlocks += Get-XmlNode $g.SymName $g.Operator $g.Values[0] $false
    }

    # 2. Broad Catch (Hidden, alleen naar Backend XML!)
    foreach ($fName in $CombinationFields) {
        $def = $global:PropLookup[$fName]
        
        # Voor de JSON houden we de mooie schone INANY lijst (zodat de UI code het snapt indien nodig)
        $jsonCriteria += [ordered]@{
            itemId = ""; displayValues = @(); hasSelectedOperators = $false; hidden = $true
            dataType = $def.DataType; values = @($global:UniqueValues[$fName]) + ""; name = $def.Name
            selectedOperator = "INANY"; readOnly = $false; id = $fName; cardinality = "SINGLE"; required = $false
        }
        
        # FIX: Voor de XML bouwen we wiskundige Binaire OR trappen ipv de falende <in><list> tag
        $broadCatchNodes = @()
        foreach ($v in $global:UniqueValues[$fName]) {
            $broadCatchNodes += Get-XmlNode $fName "EQUAL" $v $true
        }
        if ($broadCatchNodes.Count -gt 0) { $xmlBlocks += Build-XmlStaircase $broadCatchNodes "or" }
    }

    # 3. Blacklist (Hidden, alleen naar Backend XML!)
    $blacklistNodesXml = @()
    foreach ($inv in $global:InvalidCombos) {
        $flatXNodes = @()
        foreach ($fName in $CombinationFields) { $flatXNodes += Get-XmlNode $fName "NOTEQUAL" $inv[$fName] $true }
        $blacklistNodesXml += Build-XmlStaircase $flatXNodes "or"
    }
    if ($blacklistNodesXml.Count -gt 0) { $xmlBlocks += Build-XmlStaircase $blacklistNodesXml "and" }

    # JSON Opmaken (UI)
    $finalJsonObj = [ordered]@{
        includeAllProperties = $false; macros = [ordered]@{}; searchCriteria = $jsonCriteria
        resultsDisplay = [ordered]@{ columns = @("{NAME}", "DocumentTitle", "DateCreated", "Id"); honorNameProperty = $true; maxResults = 0; showChildComponentAttribute = $false; showContentSummary = $false; sortAsc = $false; sortBy = "DateCreated" }
    }
    [System.IO.File]::WriteAllText("$($GlobalConfig.OutputDir)\file1", ($finalJsonObj | ConvertTo-Json -Depth 10 -Compress), [System.Text.Encoding]::UTF8)
    
    # XML Opmaken (Backend)
    [System.IO.File]::WriteAllText("$($GlobalConfig.OutputDir)\file0", (Build-SearchWrapperXml (Build-XmlStaircase $xmlBlocks "and")), [System.Text.Encoding]::UTF8)
    Write-Host "    Gereed (file0 & file1)."
}

function Get-AcceSqlString {
    Write-Host "`n[*] Genereert direct uitvoerbare ACCE SQL Query..." -ForegroundColor Cyan
    $cls = $GlobalConfig.SubclassSymName
    $sql = "SELECT [this], DocumentTitle, Id, DateCreated FROM $cls WHERE (1=1)"

    foreach ($g in $GlobalFilters) {
        $op = "="; if ($g.Operator -eq "GREATEROREQUAL") { $op = ">=" }; if ($g.Operator -eq "LESSOREQUAL") { $op = "<=" }
        $val = if ($g.Values[0] -eq "") { "NULL" } else { "'$($g.Values[0])'" }
        if ($val -ne "NULL") { $sql += "`n  AND ($($g.SymName) $op $val)" }
    }
    foreach ($fName in $CombinationFields) {
        $inList = "'" + ($global:UniqueValues[$fName] -join "','") + "'"
        $sql += "`n  AND ($fName IN ($inList))"
    }
    foreach ($inv in $global:InvalidCombos) {
        $checks = @()
        foreach ($fName in $CombinationFields) { $checks += "$fName = '$($inv[$fName])'" }
        $inner = $checks -join " AND "
        $sql += "`n  AND NOT ($inner)"
    }
    [System.IO.File]::WriteAllText("$($GlobalConfig.OutputDir)\ACCE_Query.sql", $sql, [System.Text.Encoding]::UTF8)
    Write-Host "    Gereed (ACCE_Query.sql)."
}
#endregion


#region 5. DEPLOYMENT & TESTING (API)
function Update-IcnStoredSearchByName {
    param ([string]$SearchName, [ValidateSet("ICN_BOTH","ICN_XML_ONLY","ICN_JSON_ONLY","XT_BOTH")][string]$Mode)
    if ($null -eq $fnp8os) { Write-Host "Skipping deployment: FileNet niet verbonden." -ForegroundColor DarkYellow; return }

    Write-Host "`n[*] Pusht bestanden naar FileNet ($Mode) -> '$SearchName'..." -ForegroundColor Cyan
    $query = "SELECT Id FROM StoredSearch WHERE ([IsCurrentVersion] = True) AND (DocumentTitle = '$SearchName')"
    $enumerator = ([FileNet.Api.Query.SearchScope]::new($fnp8os)).FetchRows([FileNet.Api.Query.SearchSQL]::new($query), $null, $null, $true).GetEnumerator()

    if (-not $enumerator.MoveNext()) { Write-Host "    Fout: Search '$SearchName' niet gevonden!" -ForegroundColor Red; return }

    $storedSearch = [FileNet.Api.Core.Factory+StoredSearch]::FetchInstance($fnp8os, $enumerator.Current.Properties.GetIdValue("Id"), $null)
    $storedSearch.Checkout([FileNet.Api.Constants.ReservationType]::Exclusive, $null, $storedSearch.ClassName, $storedSearch.Properties)
    $storedSearch.Save([FileNet.Api.Constants.RefreshMode]::Refresh)

    $res = [FileNet.Api.Core.Factory+StoredSearch]::FetchInstance($fnp8os, $storedSearch.Reservation.Id, $null)
    $contentList = [FileNet.Api.Core.Factory+ContentElement]::CreateList()

    if ($Mode -match "ICN_BOTH|ICN_XML_ONLY") {
        $xStream = [System.IO.File]::OpenRead("$($GlobalConfig.OutputDir)\file0")
        $ctX = [FileNet.Api.Core.Factory+ContentTransfer]::CreateInstance(); $ctX.SetCaptureSource($xStream); $ctX.RetrievalName = "file0"; $ctX.ContentType = "application/x-filenet-searchtemplate"
        $contentList.Add($ctX)
    }
    if ($Mode -match "ICN_BOTH|ICN_JSON_ONLY") {
        $jStream = [System.IO.File]::OpenRead("$($GlobalConfig.OutputDir)\file1")
        $ctJ = [FileNet.Api.Core.Factory+ContentTransfer]::CreateInstance(); $ctJ.SetCaptureSource($jStream); $ctJ.RetrievalName = "file1"; $ctJ.ContentType = "application/json"
        $contentList.Add($ctJ)
    }
    
    $res.ContentElements = $contentList
    $res.Checkin([FileNet.Api.Constants.AutoClassify]::DO_NOT_AUTO_CLASSIFY, [FileNet.Api.Constants.CheckinType]::MAJOR_VERSION)
    $res.Save([FileNet.Api.Constants.RefreshMode]::Refresh)
    
    if ($null -ne $xStream) { $xStream.Close() }; if ($null -ne $jStream) { $jStream.Close() }
    Write-Host "    Succes! Stored Search geüpdatet." -ForegroundColor Green
}

function Test-IcnStoredSearch {
    param ([string]$SearchName)
    if ($null -eq $fnp8os) { return }
    Write-Host "`n[*] Test Backend API Executie voor '$SearchName'..." -ForegroundColor DarkYellow
    $query = "SELECT Id FROM StoredSearch WHERE ([IsCurrentVersion] = True) AND (DocumentTitle = '$SearchName')"
    $enumerator = ([FileNet.Api.Query.SearchScope]::new($fnp8os)).FetchRows([FileNet.Api.Query.SearchSQL]::new($query), $null, $null, $true).GetEnumerator()
    
    if ($enumerator.MoveNext()) {
        $storedSearch = [FileNet.Api.Core.Factory+StoredSearch]::FetchInstance($fnp8os, $enumerator.Current.Properties.GetIdValue("Id"), $null)
        try {
            $testResults = ([FileNet.Api.Query.SearchScope]::new($fnp8os)).FetchRows($storedSearch, $GlobalConfig.ClassName, $null, [System.Nullable[int]]50, $null, $true)
            $count = 0; $resultEnum = $testResults.GetEnumerator(); while ($resultEnum.MoveNext()) { $count++ }
            Write-Host "    API Succes! $count documenten gevonden zonder engine crash." -ForegroundColor Green
        } catch { Write-Host "    API FOUT: $($_.Exception.Message)" -ForegroundColor Red }
    }
}

function Generate-TestData {
    if ($null -eq $fnp8os) { return }
    Write-Host "`n[*] Genereert Test Documenten in FileNet..." -ForegroundColor Magenta
    $docCount = 0
    foreach ($row in $global:ParsedRows) {
        try {
            $doc = [FileNet.Api.Core.Factory+Document]::CreateInstance($fnp8os, $GlobalConfig.SubclassSymName)
            $doc.Properties["DocumentTitle"] = "TestDoc: $($row[$CombinationFields[0]])"
            foreach ($fName in $CombinationFields) { $doc.Properties[$fName] = $row[$fName] }
            $doc.Checkin([FileNet.Api.Constants.AutoClassify]::DO_NOT_AUTO_CLASSIFY, [FileNet.Api.Constants.CheckinType]::MAJOR_VERSION)
            $doc.Save([FileNet.Api.Constants.RefreshMode]::NO_REFRESH)
            $docCount++
        } catch { Write-Host "    Fout bij document creatie: $($_.Exception.Message)" -ForegroundColor Red }
    }
    Write-Host "    Klaar. $docCount documenten ingeschoten." -ForegroundColor Green
}
#endregion


#region 6. UTILITIES (DECOMPILE & REPAIR)
function Get-FileNetSearchXml {
    param ([string]$SearchName)
    if ($null -eq $fnp8os) { return $null }

    Write-Host "`n[*] Haalt XML op uit FileNet voor: '$SearchName'..." -ForegroundColor Cyan
    $query = "SELECT Id FROM StoredSearch WHERE ([IsCurrentVersion] = True) AND (DocumentTitle = '$SearchName')"
    $enumerator = ([FileNet.Api.Query.SearchScope]::new($fnp8os)).FetchRows([FileNet.Api.Query.SearchSQL]::new($query), $null, $null, $true).GetEnumerator()

    if (-not $enumerator.MoveNext()) { Write-Host "    Niet gevonden!" -ForegroundColor Red; return $null }

    $storedSearch = [FileNet.Api.Core.Factory+StoredSearch]::FetchInstance($fnp8os, $enumerator.Current.Properties.GetIdValue("Id"), $null)
    foreach ($ce in $storedSearch.ContentElements) {
        if ($ce.RetrievalName -eq "file0" -or $ce.RetrievalName -eq "Content.xml") {
            $stream = $ce.AccessContentStream()
            $reader = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::UTF8)
            $xmlString = $reader.ReadToEnd()
            $reader.Close()

            try { return [xml]$xmlString } 
            catch {
                Write-Host "    [!] KRITIEKE FOUT: XML corrupt (Buffer Boundary Bug in ICN of XT)!" -ForegroundColor Red
                Write-Host "    Melding: $($_.Exception.Message)" -ForegroundColor DarkGray
                return $null
            }
        }
    }
    return $null
}

function Convert-FileNetXmlToSql {
    param ([xml]$SearchXml)
    Write-Host "    Decompileert XML naar SQL..." -ForegroundColor Cyan
    $columns = @("[this]")
    foreach ($node in $SearchXml.GetElementsByTagName("selectprop")) {
        $sym = $node.GetAttribute("symname")
        if ($sym -and $sym -notin $columns) { $columns += $sym }
    }
    $selectClause = $columns -join ", "
    $fromClause = "document"
    $classNodes = $SearchXml.GetElementsByTagName("class")
    if ($classNodes.Count -gt 0) { $fromClause = $classNodes[0].GetAttribute("symname") }

    function Parse-XmlNode($node) {
        if ($null -eq $node -or $node.NodeType -ne [System.Xml.XmlNodeType]::Element) { return "" }
        $name = $node.LocalName.ToLower()

        if ($name -eq "where") {
            $parts = @(); foreach ($child in $node.ChildNodes) { $r = Parse-XmlNode $child; if ($r) { $parts+=$r } }
            if ($parts.Count -eq 0) { return "(1=1)" }; if ($parts.Count -eq 1) { return $parts[0] }
            return "(" + ($parts -join " AND ") + ")"
        }
        if ($name -eq "and" -or $name -eq "or") {
            $parts = @(); foreach ($child in $node.ChildNodes) { $r = Parse-XmlNode $child; if ($r) { $parts+=$r } }
            if ($parts.Count -eq 0) { return "" }; if ($parts.Count -eq 1) { return $parts[0] }
            $op = if ($name -eq "and") { " AND " } else { " OR " }
            return "(" + ($parts -join $op) + ")"
        }
        if ($name -match "^(eq|neq|gt|gte|lt|lte|like)$") {
            $prop = $node.whereprop; $lit = $node.literal; if (!$prop) { return "" }
            $sym = $prop.symname; $val = if ($lit) { $lit.InnerText } else { "" }
            $type = "typestring"; if ($prop.propdesc) { $type = $prop.propdesc.datatype }
            $needsQ = ($type -match "string|date|guid")
            $sqlOp = @{ "eq"="="; "neq"="<>"; "gt"=">"; "gte"=">="; "lt"="<"; "lte"="<="; "like"="LIKE" }[$name]

            if ([string]::IsNullOrEmpty($val)) {
                if ($sqlOp -eq "=") { return "($sym IS NULL)" }; if ($sqlOp -eq "<>") { return "($sym IS NOT NULL)" }
                return "($sym $sqlOp '')"
            }
            if ($needsQ) { $val = "'$val'" }; return "($sym $sqlOp $val)"
        }
        if ($name -eq "in") {
            $prop = $node.whereprop; $list = $node.list; if (!$prop -or !$list) { return "" }
            $sym = $prop.symname; $type = "typestring"; if ($prop.propdesc) { $type = $prop.propdesc.datatype }
            $needsQ = ($type -match "string|date|guid")
            $vals = @(); foreach ($lit in $list.literal) { $v = $lit.InnerText; if ($needsQ) { $v="'$v'" }; $vals+=$v }
            if ($vals.Count -eq 0) { return "" }; return "($sym IN (" + ($vals -join ", ") + "))"
        }
        return ""
    }

    $whereNode = $null; $whereList = $SearchXml.GetElementsByTagName("where")
    if ($whereList.Count -gt 0) { $whereNode = $whereList[0] }
    $whereClause = Parse-XmlNode $whereNode
    if ([string]::IsNullOrWhiteSpace($whereClause)) { $whereClause = "(1=1)" }

    $sql = "SELECT $selectClause `nFROM $fromClause `nWHERE $whereClause"
    return ($sql -replace "\) AND \(", ") `n  AND (")
}

function Get-IcnSearchAsSql {
    param ([string]$SearchName)
    $xml = Get-FileNetSearchXml -SearchName $SearchName
    if ($null -ne $xml) {
        $sql = Convert-FileNetXmlToSql -SearchXml $xml
        Write-Host "`n--- SQL RESULTAAT ---" -ForegroundColor Magenta
        Write-Host $sql -ForegroundColor White
    }
}
#endregion


#region 7. TEST SCENARIO'S (EXECUTION)
<# 
==================================================================================
  HIERONDER VIND JE DE BLOKKEN OM VERSCHILLENDE SCENARIO'S TE TESTEN.
  HAAL DE HAKJES '#' WEG BIJ HET SCENARIO DAT JE WILT DRAAIEN.
================================================================================== 
#>

# --- SCENARIO 1: Genereer lokaal bestanden (om te bekijken in VSCode) ---
# Get-CartesianBlacklist
# Export-ICNFiles
# Export-WorkplaceXTFiles
# Get-AcceSqlString


# --- SCENARIO 2: Upload perfecte Split-Brain files naar ICN en test of het niet crasht ---
# Get-CartesianBlacklist
# Export-ICNFiles
# Update-IcnStoredSearchByName -SearchName "ZO3 Elwin Master" -Mode "ICN_BOTH"
# Test-IcnStoredSearch -SearchName "ZO3 Elwin Master"


# --- SCENARIO 3: Haal XML op uit een bestaande zoekopdracht en print de SQL ---
# Get-IcnSearchAsSql -SearchName "Mijn Bestaande Zoekopdracht"


# --- SCENARIO 4: Genereer Test Data (Voor in ACCE/ICN) ---
# Get-CartesianBlacklist
# Generate-TestData

#endregion
