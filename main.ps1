###############
# Freshservice Printer Monitor
###############
# Creates a Freshservice ticket when a printer is running low on ink
# Created by Caleb Fraser (github.com/yuniq-neko)

###############
# Configuration
###############
# Tolerance : This is the level of ink we want the ticket generated at. (Equal to or Below)
$tol = 1
# Freshservice URL : This is the URL of your Freshservice Helpdesk (https://<NAME>.freshservice.com/)
$FreshURL = 'https://<helpdesk-name>.freshservice.com/'
# Freshservice API Key : This is the API key that is needed to authorise creating tickets
$FreshAPI = ''

###############
# Processing : Load all functions into memory; and then run them! 😁
###############
Function Main { # Main Processing Function; Gets data from Printers, and saves them into CSV.
    $SNMP = New-Object -ComObject olePrn.OleSNMP # Import SNMP Libraries
    Set-Location $(Split-Path $MyInvocation.MyCommand.Path) # Move the Powershell session to our location
    Import-Csv -Path '.\EPM_data.csv' | ForEach-Object { # Import CSV; Run the following for each printer in CSV:
        $snmp.open($_.Hostname,"public",2,3000) # Open a SNMP connection to the printer
        $_.KEY = $snmp.get("43.11.1.1.9.1.1") # Get Value of SNMP Code for Black Ink, Save it to the variable from CSV
        $_.CYAN = $snmp.get("43.11.1.1.9.1.2") # Get Value of SNMP Code for Cyan Ink, Save it to the variable from CSV
        $_.MAGENTA = $snmp.get("43.11.1.1.9.1.3") # Get Value of SNMP Code for Magenta Ink, Save it to the variable from CSV
        $_.YELLOW = $snmp.get("43.11.1.1.9.1.4") # Get Value of SNMP Code for Yellow Ink, Save it to the variable from CSV
        If ($_.CYAN -le $tol -or $_.MAGENTA -le $tol -or $_.YELLOW -le $tol -or $_.KEY -le $tol) { # Below Tolerance? If so; Set Low to 1
            $_.Low = '1' 
        } else { # No longer low; Reset all the tags.
            $_.Low = '0'
            $_.Alerted = '0'
        }
        If ($_.Low -eq '1' -and $_.Alerted -eq '0') { # Low and have'nt reported it already?
            BeamItUp -H $_.Hostname -C $_.CYAN -M $_.MAGENTA -Y $_.YELLOW -K $_.KEY # "BeamItUp" to Scotty! (Or well... Send a Freshservice Ticket... 😁)
            $_.Alerted = '1' # Flip the tag for Alerted; As we have now reported it
        }
        $_ # Output results back into Console in preperation for re-wrapping into CSV
    } | Export-Csv -Path '.\EPM_data.csv.tmp' -NoTypeInformation # Create Temporary CSV with all up to date data
    Remove-Item -Path '.\EPM_data.csv' # Remove and Drop the Original CSV File
    Rename-Item -Path '.\EPM_data.csv.tmp' -NewName '.\EPM_data.csv' # And make the Temporary CSV into a not-so-temporary CSV
}

Function BeamItUp {
    Param ([string]$H,[string]$C,[string]$M,[string]$Y,[string]$K)
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $EncodedCredentials = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(("{0}:{1}" -f $FreshAPI,$null)))
    $HTTPHeaders = @{}
    $HTTPHeaders.Add('Authorization', ("Basic {0}" -f $EncodedCredentials))
    $HTTPHeaders.Add('Content-Type', 'application/json')
    $Description = '<!doctype html><html><head><meta charset="utf-8"><style>h1,p{font-family: Segoe, "Segoe UI", "DejaVu Sans", "Trebuchet MS", Verdana, "sans-serif"; margin:0px;}</style></head><body><p><b>A Epson printer has reached 1% of remaining ink and is in need of Ink Bag Replacement<br>Please replace the ink bag in this printer as soon as possible</b><br><br><b>Hostname:</b> ' + $H + '<br><b>Cyan:</b> ' + $C + '%<br><b>Magenta:</b> ' + $M + '%<br><b>Yellow:</b> ' + $Y + '%<br><b>Key:</b> ' + $K + '%</p></body></html>'
    $Subject = $H + ' is in need of inkbag replacement'
    $URL = $FreshURL + 'helpdesk/tickets.json'
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
