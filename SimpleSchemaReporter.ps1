<#
.SYNOPSIS

Reports on the Active Directory Domain Services (ADDS) schema for an ADDS forest.

.DESCRIPTION

The Simple Schema Reporter is designed to connect to the schema of the current ADDS forest
and report on the attributes of the class(es) specified by the ClassName parameter, or provide a
list of classes using the ListClasses parameter.

There is no requirement for the Active Directory cmdlets to be present on the machine running this
script as .Net is used directly.

Reports are available as HTML, XML or CSV files or can be output directly onto the Clipboard as an HTML table.

An option is available to immediately view the report(s) generated using the ViewOutput parameter.

The HTMLFile output contains a JavaScript function to sort the results of the table by any column heading.
Please be aware the the performance of the sorting script is poor, so please bear with it whilst it sorts.
Many of the schema entries have ~400 properties!

Find the current version on GitHub at: https://github.com/wightsci/SimpleSchemaReporter

.PARAMETER ListClasses
This switch parameter specifies that the report will be a list of all classes available,
rather than details about one class.

.PARAMETER ClassName
Specifies the Schema Class(es) to report on. The default class name is User. This parameter 
acccepts a comma-separated list. The list of class names is not validated before an attempt
is made to obtain schema information.

.PARAMETER ReportType
Specifies the type of report. This can be one or more of:
    HTMLFile
    HTMLClipboard
    CSVFile
    XMLFile

The default report type is HTMLFile. Multiple report formats should be comma-separated.

.PARAMETER ReportName
Specifies the file name to be used for the report. If no ReportName parameter is
provided, a system generated file name is used, based on the date, time and schema class.
If a ReportName parameter is provided then the schema class name is appended.

.PARAMETER ViewOutput
Specifies whether the report is displayed using the default application
for its file type. The default value is False.

.EXAMPLE
    SimpleSchemaReporter.ps1

Runs SimpleSchemaReporter with all defaults. An HTML report is generated for the User class
with a system generated name.

.EXAMPLE
    SimpleSchemaReporter.ps1 -ListClasses

Runs SimpleSchemaReporter to list classes available in ADDS. An HTML report is generated.

.EXAMPLE
    SimpleSchemaReporter.ps1 -ClassName Computer

An HTML report is generated for the Computer class with a system generated name.

.EXAMPLE
    SimpleSchemaReporter.ps1 -ClassName Computer -ViewOutput

An HTML report is generated for the Computer class with a system generated name, the report
is displayed in the user's default HTML viewer.

.EXAMPLE
    SimpleSchemaReporter.ps1 -ClassName Computer,User,Contact -ReportType HTML,CSV

An HTML report and a CSV report is generated for the Computer, User and Contact classes with system generated names.



.INPUTS

None. You cannot pipe objects to SimpleSchemaReporter.ps1

#>
Param (
    [Parameter(Mandatory=$False,ParameterSetName='ForClass')]
    [String[]]
    $ClassName = 'User',
    [Parameter(Mandatory=$True,ParameterSetName='ListClass')]
    [Switch]
    $ListClasses,
    [Parameter(Mandatory=$False,ParameterSetName='ForClass')]
    [Parameter(Mandatory=$False,ParameterSetName='ListClass')]
    [ValidateSet('HTMLFile','HTMLClipboard','XMLFile','CSVFile')]
    [String[]]
    $ReportType = 'HTMLFile',
    [Parameter(Mandatory=$False,ParameterSetName='ForClass')]
    [Parameter(Mandatory=$False,ParameterSetName='ListClass')]
    [String]
    $ReportName,
    [Parameter(Mandatory=$False,ParameterSetName='ForClass')]
    [Parameter(Mandatory=$False,ParameterSetName='ListClass')]
    [Switch]
    $ViewOutput = $False 
)
<#
Tested on Windows 10/Server 2019
#>
$stylesheet = @"
<style>
body {
    font-family: Calibri, Segoe, Arial, Sans-Serif;
    font-size: 9pt;
}
table {
    border-collapse: collapse;
    border-color: RoyalBlue;
    border-style: solid;
    width: 100%;
}
th, td {
    text-align: left;
    padding: 0.5em;
    border-collapse: collapse;
    border-color: RoyalBlue;
    border-style: solid;
    border-width: 1pt;
    width: 100%;
}
th {
    background-color: SteelBlue;
    color: White;
    cursor: pointer;
}
tr:nth-child(odd) {
    background-color: CornflowerBlue;
}
h1 {
    color: SteelBlue;
}
.reportfooter {
    font-size: 8pt;
    color: SteelBlue;
}
</style>
"@

#W3Schools Sort Function
$htmlScript = @"
//<![CDATA[
    function sortTable(n) {
        document.body.style.cursor = 'progress';
        var table, rows, switching, i, x, y, shouldSwitch, dir, switchcount = 0;
        table = document.getElementsByTagName("table")[0];
        switching = true;
        // Set the sorting direction to ascending:
        dir = "asc"; 
        /* Make a loop that will continue until
        no switching has been done: */
        while (switching) {
            // Start by saying: no switching is done:
            switching = false;
            rows = table.rows;
            /* Loop through all table rows (except the
            first, which contains table headers): */
            for (i = 1; i < (rows.length - 1); i++) {
                // Start by saying there should be no switching:
                shouldSwitch = false;
                /* Get the two elements you want to compare,
                one from current row and one from the next: */
                x = rows[i].getElementsByTagName("TD")[n];
                y = rows[i + 1].getElementsByTagName("TD")[n];
                /* Check if the two rows should switch place,
                based on the direction, asc or desc: */
                if (dir == "asc") {
                    if (x.innerHTML.toLowerCase() > y.innerHTML.toLowerCase()) {
                        // If so, mark as a switch and break the loop:
                        shouldSwitch = true;
                        break;
                    }
                } else if (dir == "desc") {
                    if (x.innerHTML.toLowerCase() < y.innerHTML.toLowerCase()) {
                        // If so, mark as a switch and break the loop:
                        shouldSwitch = true;
                        break;
                    }
                }
            }
            if (shouldSwitch) {
                /* If a switch has been marked, make the switch
                and mark that a switch has been done: */
                rows[i].parentNode.insertBefore(rows[i + 1], rows[i]);
                switching = true;
                // Each time a switch is done, increase this count by 1:
                switchcount ++; 
            } else {
                /* If no switching has been done AND the direction is "asc",
                set the direction to "desc" and run the while loop again. */
                if (switchcount == 0 && dir == "asc") {
                    dir = "desc";
                    switching = true;
                }
            }
        }
        document.body.style.cursor = 'default';
    }
//]]>
"@ 

function classprops ($classname) {
#Getting the properties of the passed-in class
    #Using a Generic.List to avoid += overheads
    $classproperties = New-Object System.Collections.Generic.List[System.Object]
    $class = $schema.FindClass($classname)
    foreach ($property in $class.MandatoryProperties) {
        $property | Add-Member -MemberType NoteProperty -Name "Mandatory" -Value $True
        if  (@($constrattrnames ) -contains $property.Name) {
            $property | Add-Member -MemberType NoteProperty -Name "Constructed" -Value $True
        }
        else {
            $property | Add-Member -MemberType NoteProperty -Name "Constructed" -Value $False
        }
        $classproperties.Add($property)   
    }
    foreach ($property in $class.OptionalProperties) {
        $property | Add-Member -MemberType NoteProperty -Name "Mandatory" -Value $False
        if  (@($constrattrnames ) -contains $property.Name) {
            $property | Add-Member -MemberType NoteProperty -Name "Constructed" -Value $True
        }
        else {
            $property | Add-Member -MemberType NoteProperty -Name "Constructed" -Value $False
        }
        $classproperties.Add($property)   
    }
    Return $classproperties
}

function findcontsrattr {
   #Finding constructed Attributes
   $searcher = New-Object System.DirectoryServices.DirectorySearcher
   $searcher.SearchRoot = $schema.GetDirectoryEntry()
   $searcher.Filter = '(&(systemFlags:1.2.840.113556.1.4.803:=4)(objectClass=attributeSchema))'
   $constrattrs = $searcher.FindAll().GetDirectoryEntry()
   Return ($constrattrs | Select-Object -Expand ldapDisplayName)
}

filter selectattributes {
    $_ | Select-Object Name,CommonName,OID,Syntax,Mandatory,Constructed,IsSingleValued,IsInAnr,RangeLower,RangeUpper,Link,LinkId
}

filter selectclasses {
    $_.FindAllClasses() | Select-Object Name,CommonName,subClassOf
}

function addjscript($html) {
    #Add a Javascript to allow table sorting
    $htmldata = [xml]($($html))
    $htmlTableHeaders = $htmldata.html.body.GetElementsByTagName("th")
    for ($i = 0; $i -lt $htmlTableHeaders.Count; $i++) {
        $newHtmlAttr = $htmldata.CreateAttribute("onclick")
        $newHtmlAttr.Value = "sortTable($($i))"
        [void]$htmlTableHeaders[$i].Attributes.SetNamedItem($newHtmlAttr)
    }
    $htmlHeadTag = $htmldata.html.head
    $newHtmlNode = $htmldata.CreateElement("script")
    $newHtmlNode.InnerXML = $htmlScript
    $newHtmlAttr = $htmldata.CreateAttribute("type")
    $newHtmlAttr.Value = "text/javascript"
    [void]$newHtmlNode.Attributes.SetNamedItem($newHtmlAttr)
    [void]$htmlHeadTag.AppendChild($newHtmlNode)
    Write-Verbose $htmlHeadTag.InnerXML
    return $htmldata
}

function generateclasslist($Type, $BaseName, $View, $schemaobject) {
    $htmlPre = @"
    <h1>Class Report</h1>
"@
    $htmlPost = @"
    <div class="reportfooter">Generated by Simple Schema Reporter.</div>
"@

$htmlHead = @"
<title>Class Report</title>
$stylesheet
"@

switch ($type) {
    'HTMLFile' {
        $outputfilename = "$($basename).html"
        $outputdata = $schemaobject | selectclasses | Sort-Object -Property Name | ConvertTo-Html -Head $htmlHead -PreContent $htmlPre -PostContent $htmlPost
        $htmlpage = addjscript $outputdata
        Out-File -FilePath $outputfilename -InputObject $htmlpage.html.OuterXml
    }
    'HTMLClipboard' {
        $outputdata = $schemaobject | selectclasses | Sort-Object -Property Name | ConvertTo-Html -Fragment -PreContent $htmlPre -PostContent $htmlPost
        Set-Clipboard -Value $outputdata
    }
    'XMLFile' {
        $outputfilename = "$($basename).xml"
        $outputdata = $schemaobject | selectclasses | Sort-Object -Property Name | ConvertTo-XML -As String -NoTypeInformation
        Out-File -FilePath $outputfilename -InputObject $outputdata
    }
    'CSVFile' {
        $outputfilename = "$($basename).csv"
        $outputdata = $schemaobject | selectclasses | Sort-Object -Property Name
        $outputdata | Export-CSV -NoTypeInformation -Path $outputfilename
    }
}
Write-Verbose "$type Class Report created"
if ($View -and ($type -ne 'HTMLClipboard')) { Start-Process $outputfilename }   
}

function generatereport ($Schema, $Type, $BaseName, $View) {
$htmlPre = @"
        <h1>Schema Report for $schemaToReport Class</h1>
"@
        $htmlPost = @"
        <div class="reportfooter">Generated by Simple Schema Reporter.</div>
"@

$htmlHead = @"
<title>Schema Report for $schemaToReport Class</title>
$stylesheet
"@
switch ($type) {
    'HTMLFile' {
        $outputfilename = "$($basename).html"
        $outputdata = $Schema | selectattributes | Sort-Object -Property Name | ConvertTo-Html -Head $htmlHead -PreContent $htmlPre -PostContent $htmlPost
        #Add a Javascript to allow table sorting
        $htmldata = addjscript $outputdata
        Out-File -FilePath $outputfilename -InputObject $htmldata.html.OuterXml
    }
    'HTMLClipboard' {
        $outputdata = $Schema | selectattributes | Sort-Object -Property Name | ConvertTo-Html -Fragment -PreContent $htmlPre -PostContent $htmlPost
        Set-Clipboard -Value $outputdata
    }
    'XMLFile' {
        $outputfilename = "$($basename).xml"
        $outputdata = $Schema | selectattributes | Sort-Object -Property Name | ConvertTo-XML -As String -NoTypeInformation
        Out-File -FilePath $outputfilename -InputObject $outputdata
    }
    'CSVFile' {
        $outputfilename = "$($basename).csv"
        $outputdata = $Schema | selectattributes | Sort-Object -Property Name
        $outputdata | Export-CSV -NoTypeInformation -Path $outputfilename
    }
}
Write-Verbose "$type Report created for $schemaToReport Class"
if ($View -and ($type -ne 'HTMLClipboard')) { Start-Process $outputfilename }
}

#region main code
#Getting the schema
$schema = [System.DirectoryServices.ActiveDirectory.ActiveDirectorySchema]::GetCurrentSchema()

if ($ListClasses.IsPresent) {
    ForEach ($reportTypetoRun in $ReportType) {
        if (($Null -eq $ReportName) -or ('' -eq $ReportName)) {
            $ReportDisplayName = "SimpleSchema-Class-List-$((Get-Date).ToString('yyyy-MM-dd-HH-mm-ss'))"
            }
        else {
        Write-Verbose "***$ReportName***"
            $ReportDisplayName = "$ReportName-Class-List"
        }
        generateclasslist -Type $reportTypetoRun -View $ViewOutput -BaseName $ReportDisplayName -SchemaObject $schema
    }
}
else {
    #Getting Constructed Attributes
    $constrattrnames = findcontsrattr
    #Creating the reports
    ForEach ($schemaToReport in $ClassName) {
        $classSchema = classprops $schemaToReport
        ForEach ($reportTypetoRun in $ReportType) {
            if (($Null -eq $ReportName) -or ('' -eq $ReportName)) {
                $ReportDisplayName = "SimpleSchema-$($schemaToReport)-$((Get-Date).ToString('yyyy-MM-dd-HH-mm-ss'))"
                }
            else {
            Write-Verbose "***$ReportName***"
                $ReportDisplayName = "$ReportName-$($schemaToReport)"
            }
            generatereport -Schema $classSchema -Type $reportTypeToRun -BaseName $ReportDisplayName -View $ViewOutput
        }
    }
}
#endregion

#SimpleSchemaReporter.ps1