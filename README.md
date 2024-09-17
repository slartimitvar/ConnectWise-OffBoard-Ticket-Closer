# ConnectWise-OffBoard-Ticket-Closer
Change status=Off Board tickets to status = >Closed

Set API auth properties and server properties in the first 11 lines or so.

Will return and process all tickets matching
  - status/name="Off Board"
  - AND mergedParentTicket/id=null
  - AND closedDate<=1st day of previous to current month

Any ticket with a mergedParentTicket ID will be ignored

Works in batches of 20
