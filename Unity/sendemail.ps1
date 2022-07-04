param(
    [Parameter(Mandatory=$True, Position=0, ValueFromPipeline=$false)]
    [System.String]
    $CsvPartSuggestion,

    [Parameter(Mandatory=$True, Position=0, ValueFromPipeline=$false)]
    [System.String]
    $FaeFsrEmail,

    [Parameter(Mandatory=$True, Position=0, ValueFromPipeline=$false)]
    [System.String]
    $FsrFaeName,

    [Parameter(Mandatory=$True, Position=0, ValueFromPipeline=$false)]
    [System.String]
    $Icc3MainPart
)

#Parameters
$SiteURL = "https://arrowelectronics.sharepoint.com/sites/DCDQuickChat"

$CSVPath  = ".\test.csv"
$EmailTo = $FaeFsrEmail #$EmailTo = "marcm@crescent.com", "victor@crescent.com"    or    @("marcm@crescent.com", "victor@crescent.com")
$NameTo = $FsrFaeName
$MainPart = $Icc3MainPart
   
#Connect to PnP Online
Connect-PnPOnline -Url $SiteURL -Interactive

#Import Data from CSV File
$CSVFile = Import-Csv $CSVPath

#Define CSS Styles
$HeadTag = @"
<style type="text/css">
table {
 font-size:11px; color:#333333; border-width: 1px; border-color: #a9c6c9;
 border: b1a0c7 0.5pt solid; border-spacing: 1px; border-collapse: separate;  
}
  
th {
border-width: 1px; padding: 5px; background-color:#8064a2; font-size: 14pt; font-weight: 600;
border: #b1a0c7 0.5pt solid; font-family: Calibri; height: 15pt; color: white;
}
  
td {
 border: #b1a0c7 0.5pt solid; font-family: Calibri; height: 15pt; color: black;
 font-size: 11pt; font-weight: 400; text-decoration: none; padding:5px; 
}
  
tr:nth-child(even) { background-color: #e4dfec; }
</style>
"@

#HTML Template
$EmailPreContent = @"
<h2> Part Suggestion </h2>
<p><b> Hi $NameTo </b></p>
<p> Please review the part suggestion to the $MainPart part </p>
"@

$EmailPostContent = @"
<p> Please do not respond to this email </p>
<p> Regards </p>
<h2> ARROW </h2>
<p><b> Five Years Out </b></p>
"@

#Frame Email Body
$EmailBody = $CSVFile | ConvertTo-Html -Head $HeadTag -PreContent $EmailPreContent -PostContent $EmailPostContent | Out-String
 
#Send Email
Send-PnPMail -To $EmailTo -Subject "Part Suggestion" -Body $EmailBody #-From arrow@arrow.com -Password "password**"