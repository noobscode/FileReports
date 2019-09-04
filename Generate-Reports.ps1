###################################
#
#    Author: Alexander Nordbø
#    
#    https://github.com/noobscode
#
#    Generate neat file reports.
#
###################################

cls

############################# USAGE SECTION #############################

# Location Where reports are saved
$report = "C:\Temp\Report.html"

# Filesize in bytes. Use google to convert MB/GB to bytes and change the variable to whatever you'd like
# Example: 419430400 bytes is 400MB
$FileSize = 4

# Path to Folder from where you wish to generate the report
$FolderPath = "C:\Temp"

# File type to search for, you can use wildcards here.
# Examples: *.upf*, *.jpeg, *
$FileType = "*"

############################# END USAGE SECTION ##########################

# Convert input size from MegaBytes to Bytes
#$a = “$([math]::Round($FileSize / 1))”

# Convert Bytes to Megabytes
$result = “$([math]::Round(($FileSize * 1MB),2))”


$ConvertParams = @{
    PreContent = "<p><H1>Object Report</H1></P>
    <p>Tip: Open script file and edit values in the usage section to fit your needs</p>
	<p>You can find this report at: $report</p>
    <p>The list bellow contains a list of all $FileType files bigger than $FileSize MB</p>"
    PostContent = "<p class='footer'>$(get-date)</p>"
    head = @"
<Title> File Size Report </Title>
<style>
table a:link {
	color: #666;
	font-weight: bold;
	text-decoration:none;
}
table a:visited {
	color: #999999;
	font-weight:bold;
	text-decoration:none;
}
table a:active,
table a:hover {
	color: #bd5a35;
	text-decoration:underline;
}
table {
	font-family:Arial, Helvetica, sans-serif;
	color:#666;
	font-size:12px;
	text-shadow: 1px 1px 0px #fff;
	background:#eaebec;
	margin:20px;
	border:#ccc 1px solid;

	-moz-border-radius:3px;
	-webkit-border-radius:3px;
	border-radius:3px;

	-moz-box-shadow: 0 1px 2px #d1d1d1;
	-webkit-box-shadow: 0 1px 2px #d1d1d1;
	box-shadow: 0 1px 2px #d1d1d1;
}
table th {
	padding:21px 25px 22px 25px;
	border-top:1px solid #fafafa;
	border-bottom:1px solid #e0e0e0;

	background: #ededed;
	background: -webkit-gradient(linear, left top, left bottom, from(#ededed), to(#ebebeb));
	background: -moz-linear-gradient(top,  #ededed,  #ebebeb);
}
table th:first-child {
	text-align: left;
	padding-left:20px;
}
table tr:first-child th:first-child {
	-moz-border-radius-topleft:3px;
	-webkit-border-top-left-radius:3px;
	border-top-left-radius:3px;
}
table tr:first-child th:last-child {
	-moz-border-radius-topright:3px;
	-webkit-border-top-right-radius:3px;
	border-top-right-radius:3px;
}
table tr {
	text-align: center;
	padding-left:20px;
}
table td:first-child {
	text-align: left;
	padding-left:20px;
	border-left: 0;
}
table td {
	padding:18px;
	border-top: 1px solid #ffffff;
	border-bottom:1px solid #e0e0e0;
	border-left: 1px solid #e0e0e0;

	background: #fafafa;
	background: -webkit-gradient(linear, left top, left bottom, from(#fbfbfb), to(#fafafa));
	background: -moz-linear-gradient(top,  #fbfbfb,  #fafafa);
}
table tr.even td {
	background: #f6f6f6;
	background: -webkit-gradient(linear, left top, left bottom, from(#f8f8f8), to(#f6f6f6));
	background: -moz-linear-gradient(top,  #f8f8f8,  #f6f6f6);
}
table tr:last-child td {
	border-bottom:0;
}
table tr:last-child td:first-child {
	-moz-border-radius-bottomleft:3px;
	-webkit-border-bottom-left-radius:3px;
	border-bottom-left-radius:3px;
}
table tr:last-child td:last-child {
	-moz-border-radius-bottomright:3px;
	-webkit-border-bottom-right-radius:3px;
	border-bottom-right-radius:3px;
}
table tr:hover td {
	background: #f2f2f2;
	background: -webkit-gradient(linear, left top, left bottom, from(#f2f2f2), to(#f0f0f0));
	background: -moz-linear-gradient(top,  #f2f2f2,  #f0f0f0);	
}
</style>
"@
}

$($PSScriptRoot)

Write-Host "Generating report..."
Get-ChildItem -Recurse -Force -ErrorAction SilentlyContinue -Path $FolderPath -include $FileType |  ## optionally -include *.txt,*.bak to match any combined set of file masks 

 
where Length -gt $result |  ## use to filter by other properties   ## Note: Length is the file size 

ForEach-Object {$_ | add-member -name "Owner" -membertype noteproperty -value (get-acl $_.fullname).owner -passthru} | 
Select-Object FullName, @{Name="Size in MB";Expression={[math]::Truncate($_.Length / 1MB)}}, CreationTime, LastAccessTime, LastWriteTime, Extension, Owner | 
Sort-Object FullName |

#choose which output type you want - only one of the following 

#Export-CSV -Path c:\Temp\Report.csv -Force -NoTypeInformation 

#Out-GridView

# Output As HTML file
ConvertTo-HTML @ConvertParams | Out-File $report

Write-Host "Done!, Report saved to $report"

# Show report in browser.
Invoke-Item $report
