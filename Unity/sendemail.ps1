param(
    [Parameter(Mandatory=$True, Position=0, ValueFromPipeline=$false)]
    [System.String]
    $CsvPartSuggestion,

    [Parameter(Mandatory=$True, Position=1, ValueFromPipeline=$false)]
    [System.String]
    $FaeFsrEmail,

    [Parameter(Mandatory=$True, Position=2, ValueFromPipeline=$false)]
    [System.String]
    $FsrFaeName,

    [Parameter(Mandatory=$True, Position=3, ValueFromPipeline=$false)]
    [System.String]
    $YearQx,

    [Parameter(Mandatory=$True, Position=4, ValueFromPipeline=$false)]
    [System.String]
    $YearR
)

#Parameters
$SiteURL = "https://arrowelectronics.sharepoint.com/sites/DCDQuickChat"

$CSVPath  = $CsvPartSuggestion
$EmailTo = $FaeFsrEmail
#$EmailTo = "marcm@crescent.com", "victor@crescent.com"    or    @("marcm@crescent.com", "victor@crescent.com")
$NameTo = $FsrFaeName
   
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
<p><b> Hi $NameTo </b></p>
<p>Nice to e-meet you, I hope this email finds you well. I am contacting you as part of the Artificial Intelligence ESC system. </p>
<p>We are currently working on increase the number of registrations with attachments. Let me explain to you in more detail how this works: </p>
<ul>
<li>This information was pulled out from the DW Dashboard and corresponds to $YearQx of $YearR. </li>
</ul>
<p> Below you will find our recommendations: </p>
"@

$EmailPostContent = @"
<p> I hope this information is relevant and can help you win more sockets with your actual customers.  I would love to circle back this information in case you or your customer would like to have a call to further talk about these solutions. <p>
<p> Also, if you have any feedback or comments regarding this activity, please reply to this email. <p>
<p> Thank you in advance for your attention, I am looking forward to work together. </p>
<p> Best Regards, </p>
<h2> ARROW </h2>
<p><b> Five Years Out </b></p>
"@

#Frame Email Body
$EmailBody = $CSVFile | ConvertTo-Html -Head $HeadTag -PreContent $EmailPreContent -PostContent $EmailPostContent | Out-String
 
#Send Email
Send-PnPMail -To $EmailTo -Subject "Attachment Strategy - Product Proposal" -Body $EmailBody #-From "dcdcoordinator@arrow.com" -Password "password**"