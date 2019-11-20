# # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # 
# (c) 2019 David Berglin 
# This file is part of the PowerBiVisibility project.
# PowerBiVisibility is free software: you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation, either version 3 of the License, or (at your option) any later version.
# PowerBiVisibility is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU General Public License for more details.
# You should have received a copy of the GNU General Public License along with PowerBiVisibility.  If not, see https://www.gnu.org/licenses/.
# # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # 
# 
#  ... if you see this: 
#          MyFileName.ps1 cannot be loaded because running scripts is disabled on this system. For more information, see about_Execution_Policies at http://go.microsoft.com/fwlink/?LinkID=135170.
#
#  ... short answer
#          set-executionpolicy -scope CurrentUser -executionPolicy RemoteSigned
#
#  ... long answer
#         NOTE:  In order to run Powershell scripts, you need to enable them... There is a serious security vulnerability if web-based scripts are downloaded and executed
#              1) Find "PowerShell ISE" or else just "PowerShell"
#              2) Run as administrator
#              3) Use a new window, not a saved file.... file based scripts are considered dangerous, so they can't run in some configurations
#              4) optional: use this command to view current setting: 
#                 Get-ExecutionPolicy
#              5) use this command to allow execution permissions: 
#                 set-executionpolicy -scope CurrentUser -executionPolicy RemoteSigned
#              6) execute it (F5 in ISE or Enter in command line mode)
#              7) use this command to return to the most secure configuration: 
#                 set-executionpolicy -scope CurrentUser -executionPolicy Restricted
#
# # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # 

# see https://docs.microsoft.com/en-us/powershell/power-bi/overview

cls #clear screen


""
"This script will:"
"...copy a PowerBI Desktop (pbix) file"
"...treat it like a .zip file and unzip it"
"...pull data into memory from the Layout file of the unzipped report folder"
"...delete the copied PowerBI file and the unzipped report folder"
"...convert nested JSON into powershell objects"
"...extract a List of Measures or data columns from the JSON object"
"...read all open PowerBI Desktop (pbix) files, find the selected one"
"...query the SSAS Tabular model server of the open file"
"...extract a List of Measures or data columns from the SSAS server"
"...optionally output dependencies of visuals as SQL scripts"
"...optionally output dependencies of visuals as Tab-Delimited data"
"...optionally save a formatted JSON file of the PBIX content (eg. for source control or comparison)"
""
"...note: to preserve tabs use Powershell ISE."
"...note: to avoid text wrap: full screen and zoom out."
""
""

$isIse = Test-Path variable:global:psISE # Is the script being run within the "Powershell ISE" app?
if (-not $isISE)
{
""
"Use 'Powershell ISE' instead of the command line."
"(because the ordinary powershell command line does not preserve tab characters)"
}


#----------------------------------------------
# This function is used to optionally output detailed logs to the PowerBI window
#----------------------------------------------
$doShowLogs = $false
Function log($message) 
{   
    if($doShowLogs){
        write-host $message
    }
}



#----------------------------------------------
# This function is used to present a file picker
#----------------------------------------------
Function Use-FilePicker($initialDirectory) # https://devblogs.microsoft.com/scripting/hey-scripting-guy-can-i-open-a-file-dialog-box-with-windows-powershell/
{   
    [System.Reflection.Assembly]::LoadWithPartialName(“System.windows.forms”) | Out-Null

    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = “All files (*.*)| *.*”
    $OpenFileDialog.ShowDialog() | Out-Null

    #return choice
    $OpenFileDialog.filename
}





#----------------------------------------------
# Get User Choices
#----------------------------------------------
"Enter 'c' to cancel (for any choice)"
"Enter 'p' for a file picker"

$sessionFile = $ExecutionContext.SessionState.PSVariable.GetValue("sessionFile") # get last used file name
$initialDirectory = "%userprofile%"
if("$sessionFile" -ne "") 
{ 
    $initialDirectory = [System.IO.Path]::GetDirectoryName($sessionFile)
    "(blank = `"$sessionFile`"" 
}

$choiceOfFile = Read-Host 'Paste the full path to the PowerBI file: C:\folder\example.pbix'

if($choiceOfFile -eq "p")
{
    $choiceOfFile = Use-FilePicker -initialDirectory $initialDirectory # present a file picker
}


#----------------------------------------------
# Validate
#----------------------------------------------
if($choiceOfFile -eq "" -and "$sessionFile" -ne ""){ $choiceOfFile = $sessionFile } # no new choice, use previous file
if($choiceOfFile -eq ""){ return } # no choice at all, so quit
if($choiceOfFile -eq "c"){ return } # c = Cancel
if("$PsScriptRoot" -eq ""){
    "This powershell script must be saved to -and executed from- a location where temporary files can be stored and manipulated"
    "stopping"
    ""
    return
}
if($choiceOfFile.ToLower().EndsWith(".pbix") -eq $false){
    "a .pbix file path was not supplied: $choiceOfFile"
    "stopping"
    ""
    return
}
if((test-path $choiceOfFile) -eq $false){
    "a file was not found at this location: $choiceOfFile"
    "stopping"
    ""
    return
}



#----------------------------------------------
# Interpret User Choices
#----------------------------------------------
$fileName = [System.IO.Path]::GetFileName($choiceOfFile)
$fileNameWithoutExtension    = [System.IO.Path]::GetFileNameWithoutExtension($choiceOfFile)

"Do you want:     SQL Insert script (s)   <-- blank choice uses this"
"            Tabbed Delimited Table (t)"
"                 JSON of PBIX file (j)"
"               Measure Expressions (m)"
"                     Detailed Logs (l)"
"                       combination (stjml)"
$choiceOfOutput = (Read-Host '.................Enter Your Choice').ToLower()
if($choiceOfOutput -eq "c"){return} #cancel
if($choiceOfOutput -eq ""){$choiceOfOutput = "s"} # apply the default

$doShowSqlScript        = $choiceOfOutput.Contains("s") 
$doShowTabbedTable      = $choiceOfOutput.Contains("t")
$doGenerateJson         = $choiceOfOutput.Contains("j")
$doShowExpression       = $choiceOfOutput.Contains("m")
$doShowLogs             = $choiceOfOutput.Contains("l")
$doResolveDependencies  = ($doShowSqlScript -or $doShowTabbedTable) # Do we need to find measures? Or just generate JSON?

$ExecutionContext.SessionState.PSVariable.Set("sessionFile", $choiceOfFile ) # now that we know we are doing something, lets remember the file location for next time.

$doIgnoreVisualQueryJSON = $true
$doIgnoreVisualDataTransformJSON = $true


#----------------------------------------------
#"...copy a PowerBI report file"
#----------------------------------------------
"Copying from $choiceOfFile"
$DestinationZipFile   = "$PsScriptRoot\Temp_File_Delete_Me.zip"
$DestinationZipFolder = "$PsScriptRoot\Temp_Folder_Delete_Me"
$DestinationJSONFile  = "$PsScriptRoot\Temp_Folder_Delete_Me\Report\Layout" # NOTE: No File Extension
"Copying to   $DestinationZipFile"

Copy-Item  -Path $choiceOfFile -Destination $DestinationZipFile
"Copied"


#----------------------------------------------
#"...treat it like a .zip file and unzip it"
#----------------------------------------------
Add-Type -AssemblyName System.IO.Compression.FileSystem
function Unzip
{
    param([string]$zipfile, [string]$outpath)

    [System.IO.Compression.ZipFile]::ExtractToDirectory($zipfile, $outpath)
}
Unzip $DestinationZipFile $DestinationZipFolder 
"Unzipped"


#----------------------------------------------
#"...pull in data from the Layout file of the unzipped report folder"
#----------------------------------------------
$json = Get-Content -Raw -Path $DestinationJSONFile -Encoding Unicode

$o = $json | ConvertFrom-Json # $o stands for object... the object represented by the JSON
$wrapper = [PSCustomObject]@{
        PowerBI_Visibility_Execution_UtcDate = (Get-Date).ToUniversalTime();
        Pbix = $o;
        DBs = [System.Collections.ArrayList]@();
    }

"JSON read ...transforming"


#----------------------------------------------
#"...delete the copied file and unzipped folder"
#----------------------------------------------
Remove-Item $DestinationZipFolder -recurse
Remove-Item $DestinationZipFile 
"Temporary Files Removed"



#----------------------------------------------
# set up variable to hold collections of data that is found
#----------------------------------------------
$script:sql = ""
$script:distinctData = [System.Collections.ArrayList]@() # Holds distinct measure and locations... helps prevent duplicates where table names are missing.
$script:ports = [System.Collections.ArrayList]@()
$script:measureExpressions = [System.Collections.ArrayList]@()

$visuals  = New-Object 'system.collections.generic.dictionary[string,object]'
$measures = New-Object 'system.collections.generic.dictionary[string,object]'
$columns  = New-Object 'system.collections.generic.dictionary[string,object]'
$tables   = New-Object 'system.collections.generic.dictionary[string,object]'



#----------------------------------------------
# This function creates a "base" object which may be extended for various dependency types (tables/columns/measures/etc)
#----------------------------------------------
function New-Dependency($table, $name, $address, $source)
{
    $o = [PSCustomObject]@{
        Table = $table;
        Name = $name;
        Address = $address;
        Source = $source;
    }
    return $o
}


$script:mostRecentDB = @{}
$script:mostRecentTable = @{}
#----------------------------------------------
# This function is used to Build a simple object from the complex SSAS Tabular database.
# Assumes that the DB is looping over db >> tables >> columns/measures
# The end-goal of this is to create a JSON string which summarizes the contents of the SSAS db.
#----------------------------------------------
Function log-db-as-JSON($type, $object)
{
    if($type -eq "db"){
        # DB is handled differently then other kinds of data, since it has a lot of machine-specific settings. So skip the processes below
        $script:mostRecentDB = [PSCustomObject]@{
            Tables = [System.Collections.ArrayList]@()
            Name = $object.Name # this is probably a GUID
            Description = $object.Description
            LastSchemaUpdate = $object.LastSchemaUpdate
            CompatibilityLevel = $object.CompatibilityLevel
        }
        $rownum = $wrapper.DBs.Add($script:mostRecentDB)
        return 
    }    

    $thing = @{} # eg. table or column or measure or some other object:Partitions, Hierarchies, Annotations, ExtendedProperties, etc
    # find all properties of this object.
    $object | Get-Member | Sort-Object -Property Name | Where-Object {$_.MemberType -ne "Method" } | ForEach-Object {
        if($_.Definition.StartsWith("Microsoft.")) { return } # This is a complex object, such as the "Microsoft.AnalysisServices.Tabular.Model" and will not be included in the JSON
        $value =  $object.($_.Name)
        if($value -eq $null) { $value = "(null)"}
            
        $thing[$_.Name] = $value.ToString()
    }

    # Place this in the db object structure
    if($type -eq "table"){
        $script:mostRecentTable = $thing
        $thing.Measures = [System.Collections.ArrayList]@()
        $thing.Columns = [System.Collections.ArrayList]@()
        $thing.Objects = [System.Collections.ArrayList]@()

        $rownum = $script:mostRecentDB.Tables.Add($script:mostRecentTable)
    }    
    elseif($type -eq "measure") {  
        $rownum = $script:mostRecentTable.Measures.Add($thing) 
    }
    elseif($type -eq "column") { 
        $rownum = $script:mostRecentTable.Columns.Add($thing) 
    }
    else { # anything else
        $thing."Object-Type" = $type
        $rownum = $script:mostRecentTable.Objects.Add($thing) 
    }
}



#----------------------------------------------
# This function is used to read, format, and store data found within the PowerBI file 
# (eg. compile a list of Measure and their locations)
#             $typeID                tracks where the data was found (useful fpr wjhat type of data it is, and to aid in research of the JSON)
#             $sectionDisplayName    The Tab name
#             $visualLabel           The individual Visual on the page (type + ID)
#             $queryRef              holds the Table & measure/column info. Some table data may be missing.
#----------------------------------------------
function AddVisualDependency($typeID, $sectionDisplayName, $visualLabel, $queryRef){

    $measure = "$queryRef".Trim() # note that table name might be missing
    if($measure -eq "") { return } # String.isnullorwhitespace
    $table = ""

    # convert "dimDate.Year" tp "dimDate" & "Year"
    $periodIndex = $measure.IndexOf(".")
    if($periodIndex -gt -1){
        $table = $measure.Substring(0,$periodIndex)
        $measure = $measure.Substring($periodIndex + 1)
    }

    # Extract Table and Measure names from PBIX reference. This happens when the queryRef uses: Sum(tbl.col) so we will extract 'tbl' and 'col'
    if($table.IndexOf("(") -gt -1 -and $measure.EndsWith(")"))
    { 
        $original = $table
        $table = $table.Substring($table.IndexOf("(") + 1)
        log ("in $queryRef ...this script intrepreted the table   '$original' as '$table' ")
        $original = $measure
        $measure = $measure.Substring(0, $measure.Length - 1)
        log ("in $queryRef ...this script intrepreted the measure '$original' as '$measure' ")
    }

    # resolve names and values
    $fullyQualifiedName = "'$table'[$measure]"
    $FullName   = "$fileName :: $sectionDisplayName :: $visualLabel"
    $typeName = 'visual'
    if($typeID -eq 4) {
    $typeName = 'Drill'
    }

    
    # track this visual dependency in memory
    $distinctText   = "$measure ...is.in... $sectionDisplayName ( $visualLabel )"
    if($script:distinctData.Contains($distinctText) -eq $false){
        $rownum = $script:distinctData.Add($distinctText);

        # build out the dependency object (each kind is a little different)
        $visualParent = New-Dependency $sectionDisplayName $visualLabel $FullName $fileName #table/name/address/source
        $visualParent | Add-Member -MemberType NoteProperty -Name "UsesMeasure"      -Value @() # empty array, this will hold a list of measures that are dependencies of this measure
    
        $visualChild  = New-Dependency $table $measure $fullyQualifiedName $fileName #table/name/address/source
        $visualParent.UsesMeasure += $visualChild
     
        $rownum = $visuals[$FullName] = $visualParent
    }
 }



#----------------------------------------------
# This function is used to extract the "Title" of visual if there is one.
#----------------------------------------------
function Get-Visual-Friendly-Name ($visual, $defaultIfMissing) {

    #"vcObjects": {
    #  "title": [
    #    {
    #      "properties": {
    #        "show": {
    #          "expr": {
    #            "Literal": {
    #              "Value": "true"      <--true seems to auto generate a Visual Title somehow, based on columns, if there is no explicit Title
    #        "text": {                      but this auto generated title is not recognized by this powershell script
    #          "expr": {
    #            "Literal": {
    #              "Value": "'Plan by Business Area'"  <--this explicit Title is recognized by this powershell script


    try {        
        #Some visual have friendly titles. If it exists, it is preferable as the label
        $visualTitle = $visual.vcObjects
        if($visualTitle -ne $null) {
            $friendlyName = $visualTitle.title[0].properties.text.expr.Literal.Value
            if($friendlyName -ne "''")
            {
                return $friendlyName
            }
        }
    } catch{ } # probably no custom label
    if($defaultIfMissing -eq $null) { return "" } # empty default
    return $defaultIfMissing 
}


#----------------------------------------------
# This function is used to collect data about nested objects and properties. Used in the process of researching an object that came from JSON
# The objects within a PowerBI report has changed, and will change, over time. 
# This code is used to read sections of the PowerBI report object, and discover what unique Property Names exists
# This is NOT used during normal execution, but is useful while coding this powershell script over time
# STEP: 1-of-2 Collect unique property names within this function
# usage
#       $visual.prototypeQuery | Skip-Null | Get-Member | ForEach-Object { AddUniquePropNames $_ }
#       |<------------------>|  ...Update that section for each object you are researching
#---------------------------------------------
$script:propertyNames = [System.Collections.ArrayList]@()
function AddUniquePropNames($propertyMember){
    if($propertyMember.MemberType -eq "Method"){ return } # ignore functions names
    $text = $propertyMember.Name
    if($script:propertyNames.Contains($text) -eq $false){
        $rownum = $script:propertyNames.Add($text);
    }
}

#---------------------------------------------
# this filter will skip null objects in a powershell pipeline, but also will skip the value of $false
# see... https://stackoverflow.com/questions/4356758/how-to-handle-null-in-the-pipeline
#---------------------------------------------
filter Skip-Null 
{ 
    if( $_ -ne $null ){ return $_ } else { return $false}
}






















"Evaluating the PBIX file"

#----------------------------------------------
# Convert Text sections into JSON
# Some properties contain a text string which also includes JSON. This loops over nested properties and converts these to JSON where they are known as objects. 
# This next section of code (and loop) traverses the JSON as an object: 
# it builds nested objects, standardizes some properties, adds reference points, and ultimiately finds Visual Dependencies
# The result is clean + consistent JSON (for saving as a file). This is also easier to traverse while looking for Measures and Column dependencies
#
# The code pattern below checks if the property is null. If it is not null, it converts the value to an object, assuming it is JSON. If it is null then the text "null" is inserted as its value. 
#       if($o.config -ne $null)         {$o.config         = $o.config | ConvertFrom-Json }          else { $o | Add-Member -NotePropertyName "config"         -NotePropertyValue "null" }
#
#----------------------------------------------
if($o.config -ne $null)         {$o.config         = $o.config | ConvertFrom-Json }          else { $o | Add-Member -NotePropertyName "config"         -NotePropertyValue "null" }
if($o.filters -ne $null)        {$o.filters        = $o.filters | ConvertFrom-Json }         else { $o | Add-Member -NotePropertyName "filters"        -NotePropertyValue "null" }

$o.pods | ForEach-Object { #what are pods? I dont know. But The script converts the JSON to an object anyhow
    $pod = $_
    if($pod.parameters -ne $null) {$pod.parameters = $pod.parameters | ConvertFrom-Json }
}

if($doIgnoreVisualQueryJSON -or $doIgnoreVisualDataTransformJSON) 
{
    "    (some JSON data is ignored)"
}

$o.sections | ForEach-Object { 
    $section = $_ # a Section is a PowerBI TAB

    if($section.config -ne $null)         {$section.config         = $section.config | ConvertFrom-Json }          else { $section | Add-Member -NotePropertyName "config"         -NotePropertyValue "null" }
    if($section.filters -ne $null)        {$section.filters        = $section.filters | ConvertFrom-Json }         else { $section | Add-Member -NotePropertyName "filters"        -NotePropertyValue "null" }

    #----------------------------------------------
    #"...convert nested JSON into powershell objects"
    # for JSON file output... convert major sections of JSON strings into the underlying objects
    #----------------------------------------------
    $section.visualContainers | ForEach-Object { 
        $container = $_

        if($container.config -ne $null)         {$container.config         = $container.config | ConvertFrom-Json }          else { $container | Add-Member -NotePropertyName "config"         -NotePropertyValue "null" }
        if($container.filters -ne $null)        {$container.filters        = $container.filters | ConvertFrom-Json }         else { $container | Add-Member -NotePropertyName "filters"        -NotePropertyValue "null" }
        if($container.query -ne $null)          {$container.query          = $container.query | ConvertFrom-Json }           else { $container | Add-Member -NotePropertyName "query"          -NotePropertyValue "null" }
        if($container.dataTransforms -ne $null) {$container.dataTransforms = $container.dataTransforms | ConvertFrom-Json }  else { $container | Add-Member -NotePropertyName "dataTransforms" -NotePropertyValue "null" }


        if($doIgnoreVisualQueryJSON) { $container.query = "ignored" }
        if($doIgnoreVisualDataTransformJSON) { $container.dataTransforms = "ignored" }

        #$container.config | Skip-Null | Get-Member | ForEach-Object { AddUniquePropNames $_ }

        $compareCounter = 10
        $container | Add-Member -NotePropertyName "compare_marker$compareCounter" -NotePropertyValue "Compare Markers are added to help file comparison software to align large complex JSON segments."
        For ($i=1; $i -le 10; $i++) 
        {
            # for JSON file output... By adding 10 rows of mostly identical data, it will help comparison software to recognize where one object ends, and another starts, since compare software only looks at lines of code.
            $compareCounter += 1; $container | Add-Member -NotePropertyName "compare_marker$compareCounter" -NotePropertyValue ($section.name + "-" +$container.config.name)
        }
        $compareCounter += 1; $container | Add-Member -NotePropertyName "compare_marker$compareCounter" -NotePropertyValue ("tab: " + $section.displayName )

        # Add friendly names of the visual, if it exists
        $container.config | ForEach-Object { 
            $visual = $_.singleVisual  
            $friendlyName = Get-Visual-Friendly-Name $visual "(No Title)" #custom function to extract "Title" of visual if there is one.
            $container.compare_marker20 = $container.compare_marker20 + " - " + $friendlyName # this updates the JSON file output so that the "compare_marker" is a little more user friendly, but this code is only hit when evaluating measures.
        }

        # This code places "compare_marker10" at the top of the JSON object by removing and re-adding all other properties.
        $namesToSortInMiddle = $container.PSObject.Properties | Where-Object { -not $_.Name.StartsWith("compare_marker") } | select -ExpandProperty Name 
        $namesToSortAtEnd = "config", "filters", "query", "dataTransforms"

        $container.PSObject.Properties | Sort-Object -Property Name | ForEach-Object {
            $name  = $_.Name
            if(-not $namesToSortInMiddle.Contains($name)) { return } # These go at the top, eg. compare_marker10
            if($namesToSortAtEnd.Contains($name)) { return } # These go at the end
            $value = $_.Value
            $container.PSObject.Properties.Remove($name)
            $container | Add-Member $name $value
        }
        $container.PSObject.Properties | Sort-Object -Property Name | ForEach-Object {
            $name  = $_.Name
            if(-not $namesToSortAtEnd.Contains($name)) { return } # These are already sorted
            $value = $_.Value
            $container.PSObject.Properties.Remove($name)

            $compareCounter += 1; $container | Add-Member -NotePropertyName "compare_marker$compareCounter" -NotePropertyValue ($section.name + "-" +$container.config.name+ "-" + $name)
            $container | Add-Member $name $value
            $compareCounter += 1; $container | Add-Member -NotePropertyName "compare_marker$compareCounter" -NotePropertyValue ($section.name + "-" +$container.config.name+ "-" + $name)
        }

    }

    #----------------------------------------------
    #"...extract a List of Measures or data columns from the JSON object"
    #----------------------------------------------
    if($doResolveDependencies -eq $true)
    {
        # $section.filters.value contain the DRILL filters
        $section.filters | ForEach-Object {
            $drill = $_.expression.Column
            if($drill -eq $null) {
                $drill = $_.expression.Measure
            }
            $drillTable   = ($drill.Expression.SourceRef.Entity)
            $drillMeasure = ($drill.Property)
            #write-host "Drill: '$drillTable'[$drillMeasure]  to " + $section.displayName
        
            AddVisualDependency 4 $section.displayName  "DrillFilter"  "$drillTable.$drillMeasure"  
        }

        $section.visualContainers | ForEach-Object { 
            $container = $_
            $container.config | ForEach-Object { 

                $visual = $_.singleVisual  
                $visualLabel = $visual.visualType + " " + $visual.name 
                $friendlyName = Get-Visual-Friendly-Name $visual #custom function to extract "Title" of visual if there is one.
                if($friendlyName -ne "") { $visualLabel = $visual.visualType + " " + $friendlyName }

                # This line builds a distinct lists the names of object properties
                #$visual.prototypeQuery | Skip-Null | Get-Member | ForEach-Object { AddUniquePropNames $_ }


                #$visual.prototypeQuery.From  
                #$visual.prototypeQuery.Select.Column.Expression   | ForEach-Object { AddVisualDependency 2 $section.displayName  $visualLabel  $_  }
                #$visual.prototypeQuery.Select.Column.Property     | ForEach-Object { AddVisualDependency 2 $section.displayName  $visualLabel  $_  }
                $visual.prototypeQuery.Select.Name                 | ForEach-Object { AddVisualDependency 2 $section.displayName  $visualLabel  $_  }
                $visual.prototypeQuery.Select.Measure.Property     | ForEach-Object { AddVisualDependency 3 $section.displayName  $visualLabel  $_  }
                #$visual.prototypeQuery.Select.Measure.Expression  | ForEach-Object { AddVisualDependency 2 $section.displayName  $visualLabel  $_  }
                #$visual.prototypeQuery.Version
                #$visual.prototypeQuery.OrderBy

                #"projections"  THESE MEASURES ARE APPARENTLY DUPLICATED IN THE prototypeQuery
                # $visual.projections.Category.queryRef   | ForEach-Object { AddVisualDependency 1 $section.displayName  $visualLabel  $_  }
                # $visual.projections.Series.queryRef     | ForEach-Object { AddVisualDependency 1 $section.displayName  $visualLabel  $_  }
                # $visual.projections.Y.queryRef          | ForEach-Object { AddVisualDependency 1 $section.displayName  $visualLabel  $_  }
                # $visual.projections.Values.queryRef     | ForEach-Object { AddVisualDependency 1 $section.displayName  $visualLabel  $_  }
                # $visual.projections.Y2.queryRef         | ForEach-Object { AddVisualDependency 1 $section.displayName  $visualLabel  $_  }
                # $visual.projections.Tooltips.queryRef   | ForEach-Object { AddVisualDependency 1 $section.displayName  $visualLabel  $_  }
                # $visual.projections.Rows.queryRef       | ForEach-Object { AddVisualDependency 1 $section.displayName  $visualLabel  $_  }
                # $visual.projections.Goal.queryRef       | ForEach-Object { AddVisualDependency 1 $section.displayName  $visualLabel  $_  }
                # $visual.projections.Indicator.queryRef  | ForEach-Object { AddVisualDependency 1 $section.displayName  $visualLabel  $_  }
                # $visual.projections.TrendLine.queryRef  | ForEach-Object { AddVisualDependency 1 $section.displayName  $visualLabel  $_  }
            }
        }
    } 

    $section.visualContainers = $section.visualContainers | Sort-Object -Property compare_marker10 # for JSON file output... Sort so that comparison of this collection is more likely to match
}

$o.sections = $o.sections | Sort-Object -Property name # for JSON file output... Sort so that comparison of this collection is more likely to match





















"Evaluating the Data Sources of open PBIX files"

$ProcessIdOfSelectedPowerBIFile = 0

#----------------------------------------------
# Connect to running SSAS Tabular Data Models in open PBIX files
# (The Data Model is encrypted, so it can only be read from a running instance.)
# (The running instance exposes a full SSAS server which can be queried like normal.)
# 1-of-2 FIND THE SERVER CONNECTION Data
#
# steps for connecting powershell to PowerBI Desktop has been adapted from these sources
#     https://www.biinsight.com/four-different-ways-to-find-your-power-bi-desktop-local-port-number/
#     https://community.powerbi.com/t5/Desktop/Powershell-to-Access-Power-BI-Desktop/td-p/569193
#     https://sysnetdevops.com/2017/04/24/exploring-the-powershell-alternative-to-netstat/
#     https://audministrator.wordpress.com/2018/11/18/powershell-accessing-power-bi-desktop-data-and-more/
#----------------------------------------------
Get-Process PBIDesktop | foreach-object{
    $report = $_
    $reportProcessId = $_.Id
    $title = $report.mainWindowTitle.ToString().Trim()
    if($title -eq "") {
        log ("Found empty report: Title: $title ...process:$reportProcessId" )
        return # The PowerBI process is running, but a report isn't open... eg. the splash screen
    }


    if(-not $title.Contains($fileNameWithoutExtension)) {
        log ("Found open report: Title: $title ...process:$reportProcessId" )
        return # The PowerBI report not for the selected file
    }
    
    "Found selected report: $fileNameWithoutExtension ...process ID:$reportProcessId" 
    $ProcessIdOfSelectedPowerBIFile = $reportProcessId


    #Find the running Child SSAS tabular model
    $Children = Get-WmiObject win32_process | where {$_.ParentProcessId -eq $reportProcessId}

    $Children | ForEach-Object{
        log ("  child: of " + $_.ParentProcessId + " is " + $_.ProcessId + " for " + $_.ProcessName )

        if($_.ProcessName -eq "msmdsrv.exe"){
            # This is the SSSAS tabular model running as a SSAS server in memory, from the selected report file
            $ProcessIdOfSelectedTabularModel = $_.ProcessId
            "Found selected report: Tabular Model SSAS Server ...process ID: $ProcessIdOfSelectedTabularModel" 


            #Handles  NPM(K)    PM(K)      WS(K)     CPU(s)     Id  SI ProcessName                                                                           
            #-------  ------    -----      -----     ------     --  -- -----------                                                                           
            #   1812      52   125412      33740       6.27   6724   0 msmdsrv                                                                               
            #   2137     163   185004     127140       3.34 150420   1 msmdsrv     <-- this is a PowerBI .pbix file that is open and running as an SSAS cube (behind the scenes)                                                                          

            #look for open ports on the SSAS tabular model process, so we can connect directly to the tabular model engine
            $found = $false
            Get-NetTCPConnection | ? { $_.State -eq "Listen" -and $_.OwningProcess -eq $ProcessIdOfSelectedTabularModel -and $_.LocalAddress -eq "127.0.0.1" } | ForEach-Object {
                $tcpConnection = $_
                "    found port: " + $tcpConnection.LocalPort
                $script:ports.Add( $tcpConnection.LocalPort ) | Out-Null
                $found = $true
            }

            if($found -eq $false) { "    no port found" }
        }
    }
}


#----------------------------------------------
# Connect to running SSAS Tabular Data Models in open PBIX files
# (The Data Model is encrypted, so it can only be read from a running instance.)
# (The running instance exposes a full SSAS server which can be queried like normal.)
# 2-of-2 READ THE SERVER DATA 
#----------------------------------------------
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.AnalysisServices.Tabular") | Out-Null;
$script:ports | ForEach-Object {
    $port = $_
    log ("    Checking Port: " + $port)

    $server = "localhost:$port"
    $as = New-Object Microsoft.AnalysisServices.Tabular.Server;
    $as.Connect($server);
    "Connected to $server  ...searching..."
    $as.Databases | ForEach-Object {
        $db = $_

        log-db-as-JSON "db" $db

        #---------------------------------------------------------------------
        #--  Read all Table and Measure data into memory
        #---------------------------------------------------------------------
        
        #Since this SSAS tabular model is linked to the open file, by ProcessID, it uses the same dependencySource
        $dependencySource = $fileName
        # old--> $dependencySource = ("" + $db.Server + " :: " + $db.Id)  # eg: localhost:59522 :: ProjDB-02_ALincoln_21a49199-db43-40a3-90ec-87d1614c7b30
        log ("              ID: " + $db.Id               )
        log ("          Server: " + $db.Server           )
        log ("      Connection: " + $server              )
        log ("LastSchemaUpdate: " + $db.LastSchemaUpdate )
        log ("dependencySource: " + $dependencySource    )
        log ""
        log ""
        log ""


        foreach($table in $db.Model.Tables | Sort-Object -Property Name) {
  
            $tableInQuotes    = "'" + $table.Name + "'" # 'dimDate'
            log ("  table:" + $tableInQuotes)
            
            log-db-as-JSON "table" $table

            $tableObject = New-Dependency $table.Name $table.Name $tableInQuotes $dependencySource #table/name/address/source
            $tableObject | Add-Member -MemberType NoteProperty -Name "LoweredName"        -Value $table.Name.ToLower()
            $tableObject | Add-Member -MemberType NoteProperty -Name "LoweredQuotedName"  -Value $tableInQuotes.ToLower()
            $tables[$table.Name] = $tableObject


            $table.Partitions | ForEach-Object { log-db-as-JSON "Partition" $_ }
            $table.Hierarchies | ForEach-Object { log-db-as-JSON "Hierarchy" $_ }
            $table.Annotations | ForEach-Object { log-db-as-JSON "Annotation" $_ }
            $table.ExtendedProperties | ForEach-Object { log-db-as-JSON "ExtendedProperty" $_ }

            foreach($sourceColumn in $table.Columns) {
                $columnInBrackets = "[" + $sourceColumn.Name + "]" # [Year]
                $address = $tableInQuotes + $columnInBrackets # eg. 'dimDate'[Year]
                log ("    col:" + $columnInBrackets)

                log-db-as-JSON "column" $sourceColumn

                $column = New-Dependency $table.Name $sourceColumn.Name $address $dependencySource #table/name/address/source
                $column | Add-Member -MemberType NoteProperty -Name "BracketedName"     -Value $columnInBrackets.ToLower()
                $column | Add-Member -MemberType NoteProperty -Name "UnBracketedName"   -Value ($table.Name + $columnInBrackets).ToLower() # table name has no apostraphe, so: myTable[myColumn] instead of 'myTable'[myColumn]
                $column | Add-Member -MemberType NoteProperty -Name "LoweredName"       -Value $address.ToLower()
                $column | Add-Member -MemberType NoteProperty -Name "Type"              -Value $sourceColumn.DataType
                $column | Add-Member -MemberType NoteProperty -Name "Length"            -Value $sourceColumn.Name.Length

                $columns[$address] = $column
            }

            foreach($sourceMeasure in $table.Measures) {

                log-db-as-JSON "measure" $sourceMeasure

                $measureInBrackets = "[" + $sourceMeasure.Name + "]" # [Year AVG]
                $address = $tableInQuotes + $measureInBrackets # eg. 'dimDate'[Year AVG]
                log ("    measure: " + $measureInBrackets )
                
                if($doShowExpression){
                    $fullExpression = "`r`n" + $sourceMeasure.Name + ":=`r`n" + $measure.Expression 
                    $script:measureExpressions.Add($fullExpression) | out-null
                    log ("      " + $fullExpression)
                }

                $measure = New-Dependency $table.Name $sourceMeasure.Name $address $dependencySource #table/name/address/source
                $measure | Add-Member -MemberType NoteProperty -Name "BracketedName"    -Value $measureInBrackets.ToLower()
                $measure | Add-Member -MemberType NoteProperty -Name "Expression"       -Value $sourceMeasure.Expression
                $measure | Add-Member -MemberType NoteProperty -Name "UsesMeasure"      -Value @() # empty array, this will hold a list of measures that are dependencies of this measure
                $measure | Add-Member -MemberType NoteProperty -Name "UsesColumn"       -Value @() # empty array, this will hold a list of columns  that are dependencies of this measure
                $measure | Add-Member -MemberType NoteProperty -Name "UsesTable"        -Value @() # empty array, this will hold a list of tables   that are dependencies of this measure 
                $measure | Add-Member -MemberType NoteProperty -Name "Length"           -Value $sourceMeasure.Name.Length

                $measures[$address] = $measure
            }# end of $table.Measures
        }# end of $db.Model.Tables
    }# end of db.Databases
    $as.Disconnect();
    "Disconnected $server"
}# end of $script:ports


#----------------------------------------------
# If the user asked to see the Measure Expressions, show them now...
#----------------------------------------------
if($doShowExpression){
    ""
    "#-----------------------"
    "# Measure Expressions   "
    "#-----------------------"
    ""
    $script:measureExpressions 
    ""
    ""
}





function Remove-Comments($code)
{
    # handle both types of comments...
    #       measure=
    #       /* 
    #           get rid of this since it might reference a 'table'[column]
    #       */ 
    #       sum(Logic-To-Keep)

    #       measure=
    #       //  get rid of this since it might reference a 'table'[column]
    #       sum(Logic-To-Keep)

    $start = $code.IndexOf("/*") # MULTI-Line
    while($start -gt -1) 
    { 
        
        $end = $code.IndexOf("*/", $start); 
        if($end -lt 0 -or $end -lt $start) { break } # malformed comment
        $textToRemove = $code.Substring($start, ($end - $start + 2)) # include both /* and */ in removal
        log "Removing comment: $textToRemove"
        $code = $code.Replace($textToRemove, "") 

        $start = $code.IndexOf("/*")
    }

    $start = $code.IndexOf("//") # SINGLE-Line
    while($start -gt -1) 
    { 
        $end = $code.IndexOf("`n", $start); 
        if($end -lt $start) { $end = $code.Length } # go to end of text (eg. you are on the final line of text)
        $textToRemove = $code.Substring($start, ($end - $start))
        log "Removing comment: $textToRemove"
        $code = $code.Replace($textToRemove, "") 

        $start = $code.IndexOf("//")
    }

    return $code
}

#---------------------------------------------------------------------
#--  Process each measure. (Discover its dependencies)
#               Find columns based on:            'tbl'[col]
#               Find columns based on:              tbl[col]
#               Find measures based on:                [meas]
#               Find columns based on:                 [col]  ...no table just [col]
#               Find table name in quotes:    'dimDate'       ...no column
#               Find table name (no quotes):    dimDate       ...no column
#---------------------------------------------------------------------
if($doResolveDependencies)
{
    "Finding dependencies"

    ForEach($measure in $measures.Values)
    {
        if($measure -eq $null) {continue}
        if($measure.Expression -eq $null) {continue}

        $foundTextReplacement = "DependencyWasFoundHere"

        $code = $measure.Expression.ToLower()
        if($code.StartsWith("`"") -and $code.EndsWith("`"")) { $code = "" } # This measure only returns "text". It does not actually reference columns or other measures

        $code = Remove-Comments $code #If a comment references another measure, it is NOT a dependecy, so remove all comments

        log ""
        log ("searching for dependencies in " + $measure.Address)

        #---------------------------------------------------------------------
        #--  Find Dependencies based on various textual conventions
        #---------------------------------------------------------------------
        $foundDependencies = $false
    

        # Find columns based on: 'tbl'[col]
        Foreach($column in $columns.Values | Where-Object { $code.Contains($_.LoweredName) } | Sort-Object -Property Length -Descending ) # Sort-Object as longest first so that "Total YTD" comes before "Total"
        {
            $foundDependencies = $true; log ("found dependency: col     " + $column.Name)
            if($measure.UsesColumn -notcontains $column){ $measure.UsesColumn += $column }
            $code = $code.Replace($column.LoweredName, $foundTextReplacement) # removed the text that was found so it isn't found again
        }
    
        # Find columns based on: tbl[col]
        Foreach($column in $columns.Values | Where-Object { $code.Contains($_.UnBracketedName) } | Sort-Object -Property Length -Descending )# Sort-Object as longest first so that "Total YTD" comes before "Total"
        {        
            $foundDependencies = $true; log ("found dependency: col     " + $column.Name)
            if($measure.UsesColumn -notcontains $column){ $measure.UsesColumn += $column }
            $code = $code.Replace($column.UnBracketedName, $foundTextReplacement) # removed the text that was found so it isn't found again
        }

        # Find measures based on: [meas]
        Foreach($dependency in $measures.Values | Where-Object { $code.Contains($_.BracketedName) } | Sort-Object -Property Length -Descending )# Sort-Object as longest first so that "Total YTD" comes before "Total"
        {
            $foundDependencies = $true; log ("found dependency: measure " + $dependency.Name)
            if($measure.UsesMeasure -notcontains $dependency){ $measure.UsesMeasure += $dependency }
            $code = $code.Replace($dependency.BracketedName, $foundTextReplacement) # removed the text that was found so it isn't found again
        }

        # Find columns based on: [col]  ...no table just [col]
        Foreach($column in $columns.Values | Where-Object { $code.Contains($_.BracketedName) } | Sort-Object -Property Length -Descending )# Sort-Object as longest first so that "Total YTD" comes before "Total"
        {        
            $foundDependencies = $true; log ("found dependency: col     " + $column.Name)
            if($measure.UsesColumn -notcontains $column){ $measure.UsesColumn += $column }
            $code = $code.Replace($column.BracketedName, $foundTextReplacement) # removed the text that was found so it isn't found again
        }

        if(-not $foundDependencies) 
        {
            # The measure formula does not contain a column or another measure. 
            # Maybe it only returns a number or static text
            # But maybe it references a table without referencing a specific column... like: COUNTROWS ( 'tableName' )
            # We do not want to search for all table text, since some 'filter' measures would duplicate... filter ( 'tableName', 'tableName'[Column] <> "some value" ), <-- no need to reference 'tableName' twice

            # Find dependencies based on table name in quotes: 'dimDate'
            Foreach($table in $tables.Values | Where-Object { $code.Contains($_.LoweredQuotedName) } | Sort-Object -Property Length -Descending )# Sort-Object as longest first so that "Total YTD" comes before "Total"
            {
                $foundDependencies = $true; log ("found dependency: table   " + $table.Name)
                if($measure.UsesTable -notcontains $table){ $measure.UsesTable += $table }
                $code = $code.Replace($table.LoweredQuotedName, $foundTextReplacement) # removed the text that was found so it isn't found again
            }
            # Find dependencies based on table name (no quotes): dimDate
            Foreach($table in $tables.Values | Where-Object { $code.Contains($_.LoweredName) } | Sort-Object -Property Length -Descending )# Sort-Object as longest first so that "Total YTD" comes before "Total"
            {        
                $foundDependencies = $true; log ("found dependency: table   " + $table.Name)
                if($measure.UsesTable -notcontains $table){ $measure.UsesTable += $table }
                $code = $code.Replace($table.LoweredName, $foundTextReplacement) # removed the text that was found so it isn't found again
            }
        }

        #---------------------------------------------------------------------
        #--  Output what was discovered to the screen
        #---------------------------------------------------------------------
        #TODO: When coding this Powershell Script, use this to look at progress
        #       if($foundDependencies)
        #       {
        #           #$measure.Expression
        #           #$measure.UsesColumn   | ForEach-Object { write-output ($measure.Address + "`t...uses...`tcolumn`t" + $_.Address ) }
        #           #$measure.UsesMeasure  | ForEach-Object { write-output ($measure.Address + "`t...uses...`tmeasure`t" + $_.Address ) }
        #       }
        #       else
        #       {
        #           $measure.Address + "`t...references...`tnull`tNo Dependencies?"
        #       }

        #---------------------------------------------------------------------
        #--  Output what has changed to the screen
        #---------------------------------------------------------------------
        #TODO: When coding this Powershell Script, use this to see how the dependencies have been replaced with "DependencyWasFoundHere" and to see what's left over
        #       ""
        #       ""
        #       "----searching dependencies in measure expression..."
        #       "----original"
        #       $measure.Expression
        #       "----result (with Dependencies removed)"
        #       $code
        #       ""


        if($code.Contains("["))
        {
            ""
            "DEPENDENCY PARSING WAS NOT ABLE TO RESOLVE ALL BRACKETTED TEXT: (eg. a reference to [column] or [measure] was unrecognized)" 
            "ERROR IN " + $measure.Address
            $code
        }

        log ($measure.Name + " was parsed into...`r`n" + $code)
    }
}


"done" # Discovering data is complete. Outputing data comes next
























#---------------------------------------------
# The objects within a PowerBI report has changed, and will change, over time. 
# This code is used to read sections of the PowerBI report object, and discover what unique Property Names exists
# This is NOT used during normal execution, but is useful while coding this powershell script over time
# STEP: 2-of-2 Show results of previous research
# usage
#       $visual.prototypeQuery | Skip-Null | Get-Member | ForEach-Object { AddUniquePropNames $_ }
#       |<------------------>|  ...Update that section for each object you are researching
#---------------------------------------------
if($script:propertyNames.Length -gt 0){
    "------------------------------"
    "These Property Names were discovered in the object, using AddUniquePropNames:"
    ""
    $script:propertyNames
    "------------------------------"
}


#----------------------------------------------
#"... save a formatted JSON file of the PBIX content"
#----------------------------------------------
if($doGenerateJson -eq $true)
{
    ""
    ""
    "Converting PBIX file to a formatted JSON file"

    $uglyJSON = $wrapper | ConvertTo-Json -Depth 100 

    #attempt to use JSON.net for formatting, if present
    $pathToNewtonsoftJsonNet = $PSScriptRoot + "\Newtonsoft.Json.dll" # $PSScriptRoot is a default variable to the folder containing the Powershell script (assuming it is saved)
    if((test-path -Path $pathToNewtonsoftJsonNet))
    {
        "loading Newtonsoft.Json.dll"
        #load NewtonSoft
        # https://stackoverflow.com/questions/12923074/how-to-load-assemblies-in-powershell/37468429#37468429
        # this locks the dll file...         Add-Type -Path $pathToNewtonsoftJsonNet
        $bytes = [System.IO.File]::ReadAllBytes($pathToNewtonsoftJsonNet)
        [System.Reflection.Assembly]::Load($bytes) | Out-Null

        # JSON.net has a prettier output than native Powershell JSON which can be unreadable
        $rawJSON = [Newtonsoft.Json.Linq.JToken]::Parse($uglyJSON).ToString()
    }
    else
    {
        "A DLL to NewtonSoft JSON.Net could not be found at $pathToNewtonsoftJsonNet"
        "This DLL produces a JSON format that is easier to read. Without it, the file will save using the Powershell Json format."
        "This DLL can be downloaded from here: https://www.newtonsoft.com/json    and copy the .dll from the unzipped bin, eg: \Bin\net45\Newtonsoft.Json.dll" 
        $rawJSON = $uglyJSON.Replace("  ", " ").Replace("  ", " ")
    }
    "JSON transformed"


    $newFileName = [System.IO.Path]::GetFileNameWithoutExtension($choiceOfFile) + ".json"
    $FinalJSONFile = $choiceOfFile.Replace($fileName, $newFileName) # The original path, but new file name
    $rawJSON | Out-File $FinalJSONFile

    "JSON File Saved to $FinalJSONFile"
    ""
}














#----------------------------------------------
# This function is used to build a SQL Insert Script for each measure/column/visual 
#----------------------------------------------
function AddSQL($parent, $parentType, $child, $childType, $content)
{
    if($doShowSqlScript -eq $false) {return}

    #note the $parent is an object of type: New-Dependency($table, $name, $address, $source)
    #note the $child  is an object of type: New-Dependency($table, $name, $address, $source)

    $DependencySource   = $parent.Source.Replace("'", "''")
    $ParentLocation     = $parent.Table.Replace("'", "''")
    $ParentName         = $parent.Name.Replace("'", "''")
    $ParentAddress      = $parent.Address.Replace("'", "''")
    $ChildLocation      = If ($child -ne $null) {$child.Table.Replace("'", "''")    } Else { $null }   
    $ChildName          = If ($child -ne $null) {$child.Name.Replace("'", "''")     } Else { $null }   
    $ChildAddress       = If ($child -ne $null) {$child.Address.Replace("'", "''") } Else { $null }   
    $content            = If ($content -ne $null) {"'" + $content.Replace("'", "''") + "'" } Else { 'null' }   


    # $script:sql  <-- syntax for referencing a variable in the scope of this powershell script is $script:variableName
    $script:sql += "INSERT INTO [dbo].[Dependencies] ([Source]`r`n"
    $script:sql += "                                 ,[ParentLocation],[ParentName],[ParentAddress],[ParentType]`r`n"
    $script:sql += "                                 ,[ChildLocation], [ChildName], [ChildAddress], [ChildType] `r`n"
    $script:sql += "                                 ,[Content])`r`n"
    $script:sql += "       VALUES ('$DependencySource'`r`n"
    $script:sql += "              ,'$ParentLocation'`r`n"
    $script:sql += "              ,'$ParentName'`r`n"
    $script:sql += "              ,'$ParentAddress'`r`n"
    $script:sql += "              ,'$parentType'`r`n"
    $script:sql += "              ,'$ChildLocation'`r`n" 
    $script:sql += "              ,'$ChildName'`r`n" 
    $script:sql += "              ,'$ChildAddress'`r`n" 
    $script:sql += "              ,'$childType'`r`n" 
    $script:sql += "              ,$content)`r`n"   
}


#----------------------------------------------
# This SQL removes old data from the Dependencies table for the same source (file or SSAS Cube)
#----------------------------------------------
function AddSqlPrefix($dependencySource){
    # Assume that the user is replacing all SQL objects for this data source
    $script:sql += "`r`n"
    $script:sql += "--The line below will delete all prior data related to SSAS dependencies `r`n"
    $script:sql += "delete from [dbo].[Dependencies] where [Source] = '$dependencySource'`r`n"
    $script:sql += "select 'rows deleted:' [compare], @@RowCount as [Deleted], '$dependencySource' [Source]`r`n`r`n"
    $script:sql += "`r`n"
}

#----------------------------------------------
# This SQL checks how many rows were inserted for the source (file or SSAS Cube)
#----------------------------------------------
function AddSqlSuffix($dependencySource, $expectedCount){
    $script:sql += "`r`n"
    $script:sql += "`r`n"
    $script:sql += "`r`n SELECT"
    $script:sql += "`r`n     'rows inserted:'     as [compare]"
    $script:sql += "`r`n    ,count(*)             as [actual]"
    $script:sql += "`r`n    ,$expectedCount       as [expected]"
    $script:sql += "`r`n    ,'$dependencySource'  as [source]"
    $script:sql += "`r`n FROM [dbo].[Dependencies] "
    $script:sql += "`r`n WHERE [Source] = '$dependencySource'"
    $script:sql += "`r`n"
    $script:sql += "--WARNING previous Dependencies are DELETED by default.`r`n"
    $script:sql += "--Scroll to the top to edit the `"Delete`" script.  `r`n"
    $script:sql += "`r`n"
}

#----------------------------------------------
#"...optionally output dependencies of visuals as SQL scripts"
#----------------------------------------------
if($doShowSqlScript -eq $true)
{

    $dependencySource = ""
    $expectedCount = 0

    ForEach($visual in $visuals.Values)
    {
    # ------------- 1-of-5 ------- Prefix ----------------
        # The SQL scripts have a prefix and suffix to remove prior data for the same dependencySource (file or SSAS Cube)
        $oldSource = $dependencySource
        $dependencySource = $visual.Source
        if($dependencySource -ne $oldSource) { 
            if($oldSource -ne ""){ 
                # Add suffix for previous Source
                AddSqlSuffix $dependencySource $expectedCount 
                $expectedCount = 0
            } 
            AddSqlPrefix $dependencySource
        }
        

    # ------------- 2-of-5 ------- Visuals from PBIX file ----------------
        $visual.UsesMeasure | ForEach-Object { 
            AddSQL $visual "visual" $_ "measure" $null
            $expectedCount += 1
        }
    }
                

    # ------------- 3-of-5 ------- Measures from SSAS Cube ----------------
    ForEach($measure in $measures.Values)
    {
        # The SQL scripts have a prefix and suffix to remove prior data for the same dependencySource (file or SSAS Cube)
        $oldSource = $dependencySource
        $dependencySource = $measure.Source
        if($dependencySource -ne $oldSource) { 

            log ("dependency source changed: [$oldSource] -ne [$dependencySource]")
            
            if($oldSource -ne ""){ 
                # Add suffix for previous Source
                AddSqlSuffix $dependencySource $expectedCount 
                $expectedCount = 0
            } 
            AddSqlPrefix $dependencySource
        }

        $measure.UsesTable   | ForEach-Object { 
            AddSQL $measure "measure" $_ "table" $null
            $expectedCount += 1
        }
        $measure.UsesColumn   | ForEach-Object { 
            AddSQL $measure "measure" $_ "column" $null
            $expectedCount += 1
        }
        $measure.UsesMeasure   | ForEach-Object { 
            AddSQL $measure "measure" $_ "measure" $null
            $expectedCount += 1
        }

        # Add DAX expression of measure to SQL data
        $dax = New-Dependency "" "DAX" "" #table/name/address
        $expression =  $measure.Name + ":=`r`n" + $measure.Expression
        AddSQL $measure "measure" $dax "DAX" $expression
        $expectedCount += 1

    }

    # ------------- 4-of-5 ------- Columns from SSAS Cube ----------------
    $columns.Values | ForEach-Object {
        AddSQL $_ "column" $null $null $null # no dependencies of columns... but what if they are calculated???? hmmm.
        $expectedCount += 1
    }

    # ------------- 5-of-5 ------- Suffix ----------------
    # The SQL scripts have a prefix and suffix to remove prior data for the same dependencySource (file or SSAS Cube)
    if($dependencySource -ne ""){ AddSqlSuffix $dependencySource $expectedCount }

    ""
    "The SQL script will now be copied to your clipboard and also displayed."
    Read-Host "press Enter to continue."

    cls

    $script:sql | clip # hold in clipboard
    $script:sql        # write to screen

 
    ""
    "The code above has already been placed on the clipboard. "
    "Copy the above code and review it before executing. "
    "Be aware of line breaks which may alter SQL code when copied from the Powershell window"
}



















#---------------------------------------------------------------------
#--  Output Table Rows/Columns
#---------------------------------------------------------------------
if($doShowTabbedTable) { 

    ""
    "The tabbed data will now be copied to your clipboard and also displayed."
    Read-Host "press Enter to continue."

    cls
    $tabbedData  = "`r`n`r`n"
    $tabbedData += "Table`tName`tAddress`tType`t...References...`tTable`tName`tAddress`tType`tSource`r`n"

    if($doFilter)
    {
        $columns.Values | Where-Object {$_.LoweredName -like $loweredFilter} | ForEach-Object {
            $tabbedData += "{0}`t{1}`t{2}`t{3}`t>>`t{4}`t{5}`t{6}`t{7}`t{8}`r`n" -f $_.Table, $_.Name, $_.Address, 'column', "Nothing", $null, $null, $null, $dbId
        }
    }
    else # all columns
    {
        $columns.Values | ForEach-Object {
           $tabbedData +=  "{0}`t{1}`t{2}`t{3}`t>>`t{4}`t{5}`t{6}`t{7}`t{8}`r`n" -f $_.Table, $_.Name, $_.Address, 'column', "Nothing", $null, $null, $null, $dbId
        }
    }


    ForEach($measure in $measures.Values)
    {
            $measure.UsesTable   | ForEach-Object { 
                $tabbedData += "{0}`t{1}`t{2}`t{3}`t>>`t{4}`t{5}`t{6}`t{7}`t{8}`r`n" -f $measure.Table, $measure.Name, $measure.Address, 'measure', $_.Table, $_.Name, $_.Address, 'table', $dbId
            }
            $measure.UsesColumn   | ForEach-Object { 
                $tabbedData += "{0}`t{1}`t{2}`t{3}`t>>`t{4}`t{5}`t{6}`t{7}`t{8}`r`n" -f $measure.Table, $measure.Name, $measure.Address, 'measure', $_.Table, $_.Name, $_.Address, 'column', $dbId
            }
            $measure.UsesMeasure   | ForEach-Object { 
                $tabbedData += "{0}`t{1}`t{2}`t{3}`t>>`t{4}`t{5}`t{6}`t{7}`t{8}`r`n" -f $measure.Table, $measure.Name, $measure.Address, 'measure', $_.Table, $_.Name, $_.Address, 'measure', $dbId
            }
    }
    

    $tabbedData | clip # hold in clipboard
    $tabbedData        # write to screen
    ""
    ""
    "The data above has already been placed on the clipboard. "

}