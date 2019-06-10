#########################
# EPSON Printer Monitor #
####################v2.2#

# Configuration
#   Lower Tolerance to trigger alerts at: (equal to or below)
$tol = 1

# Processing
Function Main {
    $SNMP = New-Object -ComObject olePrn.OleSNMP
    cd 'C:\Repository\EPSON Printer Monitor\'
    Import-Csv -Path '.\EPM_data.csv' | ForEach-Object {
        $snmp.open($_.Hostname,"public",2,3000)
        $_.KEY = $snmp.get("43.11.1.1.9.1.1")
        $_.CYAN = $snmp.get("43.11.1.1.9.1.2")
        $_.MAGENTA = $snmp.get("43.11.1.1.9.1.3")
        $_.YELLOW = $snmp.get("43.11.1.1.9.1.4")
        If ($_.CYAN -le $tol -or $_.MAGENTA -le $tol -or $_.YELLOW -le $tol -or $_.KEY -le $tol) {
            $_.Low = '1'
        } else {
            $_.Low = '0'
            $_.Alerted = '0'
        }
        If ($_.Low -eq '1' -and $_.Alerted -eq '0') {
            BeamItUp -H $_.Hostname -C $_.CYAN -M $_.MAGENTA -Y $_.YELLOW -K $_.KEY
            $_.Alerted = '1'
        }
        $_
    } | Export-Csv -Path '.\EPM_data.csv.tmp' -NoTypeInformation
    Remove-Item -Path '.\EPM_data.csv'
    Rename-Item -Path '.\EPM_data.csv.tmp' -NewName '.\EPM_data.csv'
}

Function BeamItUp {
    Param ([string]$H,[string]$C,[string]$M,[string]$Y,[string]$K)
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $APIKey = '###'
    $EncodedCredentials = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(("{0}:{1}" -f $APIKey,$null)))
    $HTTPHeaders = @{}
    $HTTPHeaders.Add('Authorization', ("Basic {0}" -f $EncodedCredentials))
    $HTTPHeaders.Add('Content-Type', 'application/json')
    $Description = '<!doctype html><html><head><meta charset="utf-8"><style>h1,p{font-family: Segoe, "Segoe UI", "DejaVu Sans", "Trebuchet MS", Verdana, "sans-serif"; margin:0px;}</style></head><body><p><b>A Epson printer has reached 1% of remaining ink and is in need of Ink Bag Replacement<br>Please replace the ink bag in this printer as soon as possible</b><br><br><b>Hostname:</b> ' + $H + '<br><b>Cyan:</b> ' + $C + '%<br><b>Magenta:</b> ' + $M + '%<br><b>Yellow:</b> ' + $Y + '%<br><b>Key:</b> ' + $K + '%</p></body></html>'
    $Subject = $H + ' is in need of inkbag replacement'
    $URL = 'https://###.freshservice.com/helpdesk/tickets.json'
    $TicketAttributes = @{}
    $TicketAttributes.Add('description_html', $Description)
    $TicketAttributes.Add('subject' , $Subject)
    $TicketAttributes.Add('email' , '###')
    $TicketAttributes.Add('priority' , '3')
    $TicketAttributes.Add('status' , '2')
    $TicketAttributes.Add('source' , '2')
    $TicketAttributes.Add('ticket_type' , 'Incident')
    $TicketAttributes = @{'helpdesk_ticket' = $TicketAttributes}
    $JSON = $TicketAttributes | ConvertTo-Json
    Invoke-RestMethod -Method Post -Uri $URL -Headers $HTTPHeaders -Body $JSON
}

Main #Starts Processing now that Functions are in Memory
