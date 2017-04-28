Attribute VB_Name = "SmallMarket"
Sub SmallMarket()
'Created by Michael Welch - 3/24/16
'Removes all large market trips
'V.1 - Removes large markets by vendor
'Removes the Following:
'NYC, Teterboro, Miami, Orlando, Dallas, Houston, Palm Peach, Chicago, Los Angeles, San Francisco
'Boston, Phoenix, London, Minneapolis, Las Vegas, D.C., Groundworks and I/Os

Dim endRow As Integer

endRow = Cells(Rows.Count, "A").End(xlUp).Row

For i = 2 To endRow

'NYC

If Cells(i, 20).Value = "Limo Systems" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "Classic Limousine Service" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "MLL Limousine Services, Inc" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "NY Global" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "CJ Limousine" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "NY Limo Pros, Inc" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "Fortune Limousine" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "Orbit Limousine" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "Dominicks Limousine Service" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "Gardella's Elite Limousine" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "Klein Transportation" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "Landmark Transportation" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "Executive Transportation Group" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "Stout Transportation Services" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "US Limousine Service" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "Twin Forks Limo" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "H&B Super Express" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "Allstar International Chauffered Transportation" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "I/O Farooq - Limo Systems" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "PlannerNet" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "Bearcom- Nextels" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "GW James Carroll, NY" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "GW (Frankie) Frank Ingenito, NY" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "GW (Jim) James Roy, NY" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "I/O Albert Kataro" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "I/O Bhaskar Nair" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "I/O Miguel DeLaCruz" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "I/O Carlos DeLaCruz" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "I/O George Dandakis" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "I/O Alexander Curtinhas" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "GW Chris Condon NYC" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "GW Anthony (Tony) Perosi, NYC" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "I/O Arthur Ibragimov" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "Savoya Onsite Coordinator" Then Cells(i, 20).Value = ""

'Miami

If Cells(i, 20).Value = "Adventure Limousine" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "Chauffeured Miami" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "Associated Limousine Services" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "Royal Limousine" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "Academy/Endeavor Buslines" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "Travel by Bus" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "All Points Limousine" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "Onsite Coordinator" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "PlannerNet" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "Contract Greeter" Then Cells(i, 20).Value = ""

' Dallas

If Cells(i, 20).Value = "AJL International" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "Wynne Sedan & Motorcoach" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "Concierge Limousine" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "CitiSedan" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "A Plus Limo" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "Road Runner Charters" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "Dan Dipert" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "Gotta Go Trailways / Echo Transportation" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "GW Chris Countryman DAL" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "GW Tyler Young DAL" Then Cells(i, 20).Value = ""

'Houston

If Cells(i, 20).Value = "Global Transportation" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "GHL" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "First Class Tours" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "GW Ryan Thomason HOU" Then Cells(i, 20).Value = ""

'Palm Beach

If Cells(i, 20).Value = "Signature Limousine Services" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "Klassy Koach" Then Cells(i, 20).Value = ""

'Chicago

If Cells(i, 20).Value = "A Higher Caliber Livery" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "Isaac Alan Xport Inc" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "Travelers Limousine" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "Limo Corp of Chicago" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "Chicago Classic Coach" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "Masters Transportation" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "Midwest Motor Coach" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "Windy City Limousine" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "Groundwork Chicago - Ken Diedrich" Then Cells(i, 20).Value = ""

'Las Angeles

If Cells(i, 20).Value = "Strack Chauffeured Transportation, Inc." Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "Exclusive Livery Service, Inc." Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "Lynx Transportation" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "Signature Executive Transportation" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "AAA Limousine Service" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "DL Limousine Company" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "Southcoast Transportation" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "VIP Limousines (LA)" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "Best Limousines & Transportation" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "Fleetwood Limousine, Ltd." Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "Lux Bus America" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "Vogue Limousine" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "Horizon Limousine, Inc." Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "GW Chris Ring SAN/SNA" Then Cells(i, 20).Value = ""

'San Francisco

If Cells(i, 20).Value = "AAA Corporate Transportation" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "Black Pearl Transportation, Inc" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "Universal Limo" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "Alpha Limousine" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "Premiere Sedan & Limousine Service, Inc." Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "SF Pinnacle Limousine" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "United Car Services" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "Bali Limousine, Inc." Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "Coach 21" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "Royal Coach Tours" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "California Wine Tours" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "American Stage Tours" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "AITS-MRY" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "Peninsula Tours" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "Professional Charter Services" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "Napa Valley Tours & Transportation" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "GW - Javier" Then Cells(i, 20).Value = ""

'Boston

If Cells(i, 20).Value = "Black Tie Limousine" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "Axis Coach" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "Orient Express Limo" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "All Luxury Limousine" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "DPV Transportation" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "A Yankee Line" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "New England Coach" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "Peter Pan Bus Lines" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "Cavalier Coach" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "Boston Chauffeur" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "GW (Mike) Michael Domnarski, BOS" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "GW Jay Cronin BOS" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "GW (Ron) Ronald Legros" Then Cells(i, 20).Value = ""

'Phoenix

If Cells(i, 20).Value = "City Lights Luxury Transportation" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "Arizona Limousines" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "Jet Limousine" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "Divine Transportation" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "All Aboard America" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "Arrow Stage Lines" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "American Explorer Motorcoach" Then Cells(i, 20).Value = ""

'London

If Cells(i, 20).Value = "Gerrard Chauffeur Drive" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "TBR Global - London" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "GW Vladimir Perevezentsev LHR" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "GW Guy Batchelor LHR" Then Cells(i, 20).Value = ""

'Minneapolis


If Cells(i, 20).Value = "Total Luxury Limo" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "Star Transportation" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "Cadillac Chauffeur Services" Then Cells(i, 20).Value = ""

'Las Vegas

If Cells(i, 20).Value = "Lucky Limousine" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "Omni Limousine" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "Earth Limos" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "Executive Las Vegas" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "Arrow Stage Lines Las Vegas" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "Sweetours" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "Triple J Tours" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "Earth Buses LLC" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "Showtime Tours" Then Cells(i, 20).Value = ""

'Orlando

If Cells(i, 20).Value = "Wheelers Luxury Transportation" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "Fuego Executive Transportation" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "Wheels to Wings / Bellwood Norther LLC" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "Precision Limousine" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "Small World Tours & Cruises" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "Express Transportation" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "Escot Bus Lines" Then Cells(i, 20).Value = ""

'DC

If Cells(i, 20).Value = "Alpha Limousine DC" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "DC Livery" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "DC Trails" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "GW John McCarthy, DC" Then Cells(i, 20).Value = ""

'Groundworks and I/Os

If Cells(i, 20).Value = "GW Ben Frazier AUS/SAT" Then Cells(i, 20).Value = ""

If Cells(i, 20).Value = "I/O Thomas Hendrick" Then Cells(i, 20).Value = ""

Next i

Range("t1:t5000").Select
Selection.SpecialCells(xlCellTypeBlanks).EntireRow.Delete

End Sub
