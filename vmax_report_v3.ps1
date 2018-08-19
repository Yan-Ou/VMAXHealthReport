$smtpServer = "" # SMTP server to send email out
$from = "" # Sender's email
$to = @("","") # Recipients' emails
#$to = "yano@datacom.co.nz"

$systems = cmd /c "symcfg list" | select-string "Local"

$arrays = @{"298700272" = "AKL VMAX 10K"; "495700095" = "AKL VMAX 40K"; "298700207" = "HLZ VMAX 10K"; "498700042" = "WLG VMAX 10K"}

foreach ($system in $systems){

$system = $system.tostring().trim() -split("\s+")
$sid = $system[0].trim("0")

$table = New-Object system.Data.DataTable("Thin Pool")
[void]$table.Columns.Add("Pool Name")
[void]$table.Columns.Add("Total in Pool (GB)")
#[void]$table.Columns.Add("Usable in Pool(GB)")
[void]$table.Columns.Add("Free in Pool (GB)")
[void]$table.Columns.Add("Used in Pool (GB)")
[void]$table.Columns.Add("Full %")
[void]$table.Columns.Add("Sub %")
[void]$table.Columns.Add("Bound TDAT")
[void]$table.Columns.Add("Enabled TDAT")
[void]$table.Columns.Add("Unbound TDAT")
[void]$table.Columns.Add("Free Capacity on the Array (GB)")
[void]$table.Columns.Add("TDAT Consumed %")

$pool_info = cmd /c "symcfg -sid $sid list -pool -thin -gb -detail" | select-string "EFD_|FC_|SATA_"
$pool_rawinfo = cmd /c "symcfg -sid $sid list -pool -thin -gb -detail"
$symid = $pool_rawinfo[1].tostring().trim()
$pool_total = (cmd /c "symcfg -sid $sid list -pool -thin -gb -detail" | select-string ^"GBs").tostring().trim() -split("\s+")
$total_average = ([double]($pool_total[1])+[double]($pool_total[4]))/2

$HTMLHeader =@"
<!DOCTYPE html>
<html>
<head>
<style>
BODY{font-family: Arial; font-size: 9pt; text-aligh: center;}
TABLE{border: 1px solid black; border-collapse: collapse;}
TH{border: 1px solid black; background: #dddddd; padding: 5px;}
TD{border: 1px solid black; padding: 5px;}
</style>
</head>
<body>
"@

function Convert-DatatablHtml($dt)
{

$html = "
<div>
<h3><p>$symid</p></h3>
<table>"
$hmtl +="<tr>"
for($i = 0;$i -lt $dt.Columns.Count;$i++)
{
$html += "<td>"+$dt.Columns[$i].ColumnName+"</td>"
}
$html +="</tr>"

for($i=0;$i -lt $dt.Rows.Count; $i++)
{
$hmtl +="<tr>"
 for($j=0; $j -lt $dt.Columns.Count; $j++)
 {

  $html += "<td>"+$dt.Rows[$i][$j].ToString()+"</td>"
 }
 $html +="</tr>"
}

$html += "</table></div>
<div>
<table>
"

return $html
}

$HTMLChart = ""

$ListOfAttachments = @()
foreach ($pool in $pool_info){

$tech_fast = @{"EFD"=4;"FC"=30;"SATA"=70}

$pool = $pool.tostring().trim() -split("\s+")
$row = $table.NewRow()
$row."Pool Name" = $pool[0]
$pname = $pool[0]
$row."Total in Pool (GB)" = $pool[3]
#$row."Usable (GB)" = $pool[4]
$row."Free in Pool (GB)" = $pool[5]
$row."Used in Pool (GB)" = $pool[6]
$row."Full %" = $pool[7]
$row."Sub %" = $pool[8]
$tech = (($row."Pool Name").tostring().trim() -split("_"))[0]
$bound_tdat = cmd /c "symcfg show -sid $sid -pool $pname -thin"
$unbound_tdat = cmd /c "symdev list -sid $sid -datadev -tech $tech -nonpooled" | select-string "RW"
$row."Bound TDAT" = (($bound_tdat|select-string "of Devices in Pool").tostring() -split("\s+"))[-1]
$row."Enabled TDAT" = (($bound_tdat|select-string "Enabled Devices in Pool").tostring() -split("\s+"))[-1]
$row."Unbound TDAT" = $unbound_tdat.Count
$row."Free Capacity on the Array (GB)" = [Double](((cmd /c "symdev list -sid $sid -datadev -tech $tech -gb" | select-string "RW")[0].tostring() -split("\s+"))[8])*$row."Unbound TDAT"
$row."TDAT Consumed %" = [math]::Truncate((([Double]$row."Total in Pool (GB)")/([Double]$row."Total in Pool (GB)"+$row."Free Capacity on the Array (GB)"))*100)
$table.Rows.Add($row)

$pool_average = ([double]($row."Total in Pool (GB)")+[double]($row."Used in Pool (GB)"))/2
$current_alloc = [math]::Truncate($pool_average/$total_average*100)
#write-host $current_alloc

[void][Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms.DataVisualization")
$PoolUsageChart = New-Object System.Windows.Forms.DataVisualization.Charting.Chart
$PoolUsageChart.Width = 500
$PoolUsageChart.Height = 250
$PoolUsageChart.Left = 10
$PoolUsageChart.BackColor = [System.Drawing.Color]::White
$PoolUsageChart.BorderColor = 'Black'
$PoolUsageChart.BorderDashStyle = 'Solid'

[void]$PoolUsageChart.Titles.Add($row."Pool Name"+" Pool Usage:")
$PoolUsageChart.Titles[0].Font = "segoeuilight,9pt"
$PoolUsageChart.Titles[0].Alignment = "topleft"
$chartarea = New-Object System.Windows.Forms.DataVisualization.Charting.ChartArea
$chartarea.Name = "ChartArea1"
$PoolUsageChart.ChartAreas.Add($chartarea)

[void]$PoolUsageChart.Series.Add("data1")
$PoolUsageChart.Series["data1"].ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::Pie
$PoolCheckList = @("Used in Pool (GB)", "Free in Pool (GB)")
$PoolValueList = @($row."Used in Pool (GB)",$row."Free in Pool (GB)")
$PoolUsageChart.Series["data1"].Points.DataBindXY($PoolCheckList,$PoolValueList)

#$PoolUsageChart.Series["data1"]['PieLabelStyle'] = 'Disabled'
$PoolUsageChart.Series["data1"].Label = "#PERCENT{P2}"
$PoolLegend = New-Object System.Windows.Forms.DataVisualization.Charting.Legend
$PoolLegend.IsEquallySpacedItems = $True
$PoolLegend.BorderColor = "Black"
$PoolUsageChart.Legends.Add($PoolLegend)
$PoolUsageChart.Series["data1"].LegendText = "#VALX (#VALY)"
#$maxValue = $PoolUsageChart.Series["data1"].Points.FindMaxByValue()
#$maxValue.Color = "#FF7F50"
$PoolUsageChart.Series["data1"].Points[0].Color = "#DC143C"
$PoolUsageChart.Series["data1"].Points[1].Color = "#4682B4"

$PoolUsageChart.SaveImage((Get-Location).Path+"\"+$sid+"_"+$row."Pool Name"+"_Usage.png","png")
$image = (Get-Location).Path+"\"+$sid+"_"+$row."Pool Name"+"_Usage.png"
$attach = $sid+"_"+$row."Pool Name"+"_Usage.png"
$ListOfAttachments += $image

[void][Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms.DataVisualization")
$FASTChart = New-Object System.Windows.Forms.DataVisualization.Charting.Chart
$FASTChart.Width = 500
$FASTChart.Height = 250
$FASTChart.Left = 10
$FASTChart.BackColor = [System.Drawing.Color]::White
$FASTChart.BorderColor = 'Black'
$FASTChart.BorderDashStyle = 'Solid'

[void]$FASTChart.Titles.Add("FAST Policy Status: "+$tech)
$FASTChart.Titles[0].Font = "segoeuilight,9pt"
$FASTChart.Titles[0].Alignment = "topleft"
$chartarea2 = New-Object System.Windows.Forms.DataVisualization.Charting.ChartArea
$chartarea2.Name = "ChartArea2"
$chartarea2.AxisY.Title = "FAST Allocation (%)"
$chartarea2.AxisX.Title = ""
$chartarea2.AxisY.Interval = 10
$chartarea2.AxisX.Interval = 1
$FASTChart.ChartAreas.Add($chartarea2)

[void]$FASTChart.Series.Add("data2")
$FASTChart.Series["data2"].ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::Column
$rec_alloc = $tech_fast.Get_Item($tech)
$FASTCheckList = @("Recommended Allocatoion", "Current Allocatoion")
$FASTValueList = @($rec_alloc,$current_alloc)
$FASTChart.Series["data2"].Points.DataBindXY($FASTCheckList,$FASTValueList)
$FASTChart.Series["data2"].Points[0].Color = "#2E8B57"
$FASTChart.Series["data2"].Points[1].Color = "#FFA500"
$FASTChart.Series["data2"].Label = "#VALY %"

$FASTChart.SaveImage((Get-Location).Path+"\"+$sid+"_"+$tech+"_FAST.png","png")
$image2 = (Get-Location).Path+"\"+$sid+"_"+$tech+"_FAST.png"
$attach2 = $sid+"_"+$tech+"_FAST.png"
$ListOfAttachments += $image2

$HTMLChart += "<tr>
<td><IMG SRC=$attach></td>
<td><IMG SRC=$attach2></td></tr>
"
}
$subject = "VMAX Thin Pool Daily Usage Report for "
$subject += $arrays.Get_Item($sid)
$HTMLTable = Convert-DatatablHtml -dt $table
$HTMLEnd = "</table></div>"
$HTMLMessage = $HTMLHeader + $HTMLTable + $HTMLChart + $HTMLEnd
#HTMLMessage | out-file "c:\Scripts\test.html"
Send-MailMessage  -to $to -SmtpServer $smtpServer -Subject $subject -Attachments $ListOfAttachments -BodyAsHtml -body $HTMLMessage -From $from
}
