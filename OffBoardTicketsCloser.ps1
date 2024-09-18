# This script changes batches of 20 CW tickets to a new target status (e.g. Closed) for any status="Off Board" tickets
# which are older than 1st of last month and are not merged into another ticket 

# Define credentials for authentication
$ClientID = "<your clientID>"
$PubKey = "<your public api key>"
$PrivateKey = "<your private api key>"

# Define Server for connectivity
$Server = "<your.server.fqdn>"
$Company = "<your api users company>"

# Target ticket status
$TargetStatus = ">Closed"

# Connection variables
$Connection = @{
    Server = $Server
    Company = $Company
    PubKey = $PubKey
    PrivateKey = $PrivateKey
    ClientID = $ClientID
}

# Connect
try {
    Write-Host "`nConnecting to CW Manage..."
    Connect-CWM @Connection
    Write-Host "Done`n"
}
catch {
    Write-Host "`nAn error occurred"
    Write-Host $_.Exception.Message -ForegroundColor Red
}

# Define off board ticket search conditions
# - Created before 1st of last month
# - Status = Off Board
# - Not merged to a parent ticket
$Year = (Get-Date).year
$LastMonth = (Get-Date).month - 1
$ClosedDate = "[" + $Year + "-" + $LastMonth + "-01T00:00:00Z]"
$Condition = 'status/name="Off Board" AND mergedParentTicket/id=null AND closedDate<=' + $ClosedDate
$Fields = "id"
$PageSize = 20

# Get matching tickets
try {
    Write-Host "`nGetting off board tickets..."
    $Tickets = Get-CWMTicket -condition $Condition -fields $Fields -pageSize $PageSize
    Write-Host "Done`n"
}
catch {
    Write-Host "`nAn error occurred"
    Write-Host $_.Exception.Message -ForegroundColor Red
}

# Loop on tickets and set to target status
foreach ($Ticket in $Tickets) {
    # Build update parameters
    $UpdateParam = @{
        ID = $Ticket.id
        Operation = 'replace'
        Path = 'status/name'
        Value = $TargetStatus
    }
    # Update ticket
    try {
        Write-Host "`nUpdating/closing ticket" $Ticket.id "..."
        $Update = Update-CWMTicket @UpdateParam
        Write-Host "Done`n"
    }
    catch {
        Write-Host "`nAn error occurred"
        Write-Host $_.Exception.Message -ForegroundColor Red
    }
}

# Disconnect
try {
    Write-Host "`nDisconnecting from CW Manage..."
    Disconnect-CWM
    Write-Host "Done`n"
}
catch {
    Write-Host "`nAn error occurred"
    Write-Host $_.Exception.Message -ForegroundColor Red
}
