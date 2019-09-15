"section Section1;

shared Input = let
    Source = Excel.CurrentWorkbook(){[Name=""Input""]}[Content],
    #""Changed Type"" = Table.TransformColumnTypes(Source,{{""Name"", type text}, {""Value"", type text}}),
    FromTable = Record.FromTable(#""Changed Type""),
    #""Change types"" = Record.TransformFields(FromTable, {{""stars"", Number.FromText}, {""OnTimeCoverage"", Number.FromText}})
in
    #""Change types"";

shared Airlines = let
    Source = Table.FromRecords(Json.Document(File.Contents(Input[Airlines]))),
    #""Changed Type"" = Table.TransformColumnTypes(Source,{{""Airline"", type text}, {""IATA"", type text}, {""ICAO"", type text}})
in
    #""Changed Type"";

shared Airports = let
    Source = Table.FromRecords(Json.Document(File.Contents(Input[Airports]))),
    #""Changed Type"" = Table.TransformColumnTypes(Source,{{""IATA"", type text}, {""ICAO"", type text}, {""TimeZone"", type text}})
in
    #""Changed Type"";

shared TimeZones = let
    Source = Table.FromRecords(Json.Document(File.Contents(Input[TimeZones]))),
    #""Changed Type"" = Table.TransformColumnTypes(Source,{{""TimeZone"", type text}, {""S"", Int64.Type}, {""D"", Int64.Type}})
in
    #""Changed Type"";

shared AllFlights = let
    Source = Table.FromRecords(Json.Document(File.Contents(Input[Flights]))),
    #""Changed Type"" = Table.TransformColumnTypes(Source,{{""Date"", type date}, {""Airline"", type text}, {""From"", type text}, {""To"", type text}, {""Departure"", type time}, {""Flight Number"", Int64.Type}}),
    #""Sorted Rows"" = Table.Sort(#""Changed Type"",{{""Date"", Order.Ascending}})
in
    #""Sorted Rows"";

shared OnTimePerformance = let
    AllFlights = Table.SelectColumns(Table.ExpandTableColumn(Table.NestedJoin(AllFlights,{""Airline""},Airlines,{""Airline""},""Airlines"",JoinKind.LeftOuter), ""Airlines"", {""IATA""}, {""IATA""}),{""IATA"", ""Flight Number"", ""From""}),
    SearchResults = Table.TransformColumnTypes(Table.RenameColumns(Table.SelectRows(Table.FromRecords(List.Distinct(List.Transform(List.Combine(List.Transform(AllSearchResults, each Json.Document(_)[route])), each Record.SelectFields(_, {""operating_carrier"", ""operating_flight_no"", ""flyFrom""})))), each not(List.Contains(Record.ToList(_), """"))),{{""flyFrom"", ""From""}, {""operating_flight_no"", ""Flight Number""}, {""operating_carrier"", ""IATA""}}),{{""IATA"", type text}, {""Flight Number"", Int64.Type}, {""From"", type text}}),
    AllResults = Table.Distinct(Table.Combine({AllFlights, SearchResults})),
    #""Merged FlightBlacklist"" = Table.NestedJoin(AllResults,{""IATA"", ""Flight Number"", ""From""},FlightBlacklist,{""IATA"", ""Flight Number"", ""From""},""FlightBlacklist"",JoinKind.LeftAnti),
    #""Removed FlightBlacklist"" = Table.RemoveColumns(#""Merged FlightBlacklist"",{""FlightBlacklist""}),
    #""Merged OnTimePerformanceHC"" = Table.NestedJoin(#""Removed FlightBlacklist"",{""From"", ""Flight Number"", ""IATA""},OnTimePerformanceHC,{""From"", ""Flight Number"", ""IATA""},""OnTimePerformanceHC"",JoinKind.LeftAnti),
    #""Removed OnTimePerformance"" = Table.RemoveColumns(#""Merged OnTimePerformanceHC"",{""OnTimePerformanceHC""}),
    #""Added OnTimePerformance"" = Table.AddColumn(#""Removed OnTimePerformance"", ""OnTimePerformance"", each Text.FromBinary(Json.FromValue(List.Single(Json.Document(Web.Contents(""https://www.flightstats.com/v2/api/on-time-performance/""&Text.Combine({[IATA], Number.ToText([#""Flight Number""]), [From]}, ""/"") & ""?rqid=""&Input[rqid]))))), type text),
    #""Appended OnTimePerformanceHC"" = Table.Combine({#""Added OnTimePerformance"", OnTimePerformanceHC})
in
    #""Appended OnTimePerformanceHC"";

shared OnTimePerformanceHC = let
    Source = Excel.CurrentWorkbook(){[Name=""OnTimePerformance""]}[Content],
    #""Changed Type"" = Table.TransformColumnTypes(Source,{{""From"", type text}, {""Flight Number"", Int64.Type}, {""IATA"", type text}, {""OnTimePerformance"", type text}}),
    #""Filtered Rows"" = Table.SelectRows(#""Changed Type"", each not([OnTimePerformance] = null or Record.HasFields(Json.Document([OnTimePerformance]), {""errorCode""})))
in
    #""Filtered Rows"";

shared FlightStatsHeaders = let
    Source = [Headers=[Cookie=Input[Cookie], #""User-Agent""=Input[#""User-Agent""]]]
in
    Source;

shared HistoricalFlightDetail = let
    Source = HistoricalFlightSummary,
    #""Parsed JSON"" = Table.FromRecords(Table.TransformRows(Source, each _ & [HistoricalFlight = List.Single(List.Select(Json.Document([HistoricalFlight]), (x) => x[departureAirport][fs] = [From] and x[arrivalAirport][fs] = [To]))])),
    #""Expanded HistoricalFlight"" = Table.ExpandRecordColumn(#""Parsed JSON"", ""HistoricalFlight"", {""url""}, {""url""}),
    #""Changed Type"" = Table.TransformColumnTypes(#""Expanded HistoricalFlight"",{{""Date"", type date}, {""Airline"", type text}, {""From"", type text}, {""To"", type text}, {""Departure"", type time}, {""Flight Number"", Int64.Type}, {""IATA"", type text}, {""url"", type any}}),
    #""Merged HistoricalFlightHC"" = Table.NestedJoin(#""Changed Type"",{""Date"", ""Airline"", ""From"", ""To"", ""Departure"", ""Flight Number"", ""IATA""},HistoricalFlightDetailHC,{""Date"", ""Airline"", ""From"", ""To"", ""Departure"", ""Flight Number"", ""IATA""},""HistoricalFlightHC"",JoinKind.LeftAnti),
    #""Removed HistoricalFlightHC"" = Table.RemoveColumns(#""Merged HistoricalFlightHC"",{""HistoricalFlightHC""}),
    #""Kept First 5 Rows"" = Table.FirstN(#""Removed HistoricalFlightHC"",5),
    #""Added HistoricalFlight"" = Table.AddColumn(#""Kept First 5 Rows"", ""HistoricalFlight"", each Text.FromBinary(Web.Contents(""https://www.flightstats.com/v2/api"" &[url]&""?rqid=""&Input[rqid], FlightStatsHeaders)), type text),
    #""Appended HistoricalFlightHC"" = Table.Combine({#""Added HistoricalFlight"", HistoricalFlightDetailHC})
in
    #""Appended HistoricalFlightHC"";

shared HistoricalFlightDetailHC = let
    Source = Excel.CurrentWorkbook(){[Name=""HistoricalFlightDetail""]}[Content],
    #""Changed Type"" = Table.TransformColumnTypes(Source,{{""Date"", type date}, {""Airline"", type text}, {""From"", type text}, {""To"", type text}, {""Departure"", type time}, {""Flight Number"", Int64.Type}, {""IATA"", type text}, {""url"", type text}, {""HistoricalFlight"", type text}}),
    #""Filtered Rows"" = Table.SelectRows(#""Changed Type"", each ([HistoricalFlight] <> null))
in
    #""Filtered Rows"";

shared HistoricalFlightSummary = let
    Source = AllFlights,
    #""Filtered Rows"" = Table.SelectRows(Source, each [Date] <= Date.From(DateTime.LocalNow())),
    #""Merged Airlines"" = Table.NestedJoin(#""Filtered Rows"",{""Airline""},Airlines,{""Airline""},""Airlines"",JoinKind.LeftOuter),
    #""Expanded Airlines"" = Table.ExpandTableColumn(#""Merged Airlines"", ""Airlines"", {""IATA""}, {""IATA""}),
    #""Merged HistoricalFlightsHC"" = Table.NestedJoin(#""Expanded Airlines"",{""Date"", ""Airline"", ""From"", ""To"", ""Departure"", ""Flight Number"", ""IATA""},HistoricalFlightSummaryHC,{""Date"", ""Airline"", ""From"", ""To"", ""Departure"", ""Flight Number"", ""IATA""},""HistoricalFlightsHC"",JoinKind.LeftAnti),
    #""Removed HistoricalFlightsHC"" = Table.RemoveColumns(#""Merged HistoricalFlightsHC"",{""HistoricalFlightsHC""}),
    #""Added HistoricalFlight"" = Table.AddColumn(#""Removed HistoricalFlightsHC"", ""HistoricalFlight"", each Text.FromBinary(Web.Contents(""https://www.flightstats.com/v2/api/historical-flight/""&Text.Combine({[IATA], Number.ToText([#""Flight Number""]), Date.ToText([Date], ""yyyy/MM/dd"")}, ""/"")&""?rqid=""&Input[rqid], FlightStatsHeaders)), type text),
    #""Appended HistoricalFlightHC"" = Table.Combine({#""Added HistoricalFlight"", HistoricalFlightSummaryHC})
in
    #""Appended HistoricalFlightHC"";

shared HistoricalFlightSummaryHC = let
    Source = Excel.CurrentWorkbook(){[Name=""HistoricalFlightSummary""]}[Content],
    #""Changed Type"" = Table.TransformColumnTypes(Source,{{""Date"", type date}, {""Airline"", type text}, {""From"", type text}, {""To"", type text}, {""Departure"", type time}, {""Flight Number"", Int64.Type}, {""IATA"", type text}, {""HistoricalFlight"", type text}}),
    #""Filtered Rows"" = Table.SelectRows(#""Changed Type"", each ([HistoricalFlight] <> null))
in
    #""Filtered Rows"";

shared QueryString = let
    Source = Excel.CurrentWorkbook(){[Name=""Table14""]}[Content],
    #""Changed Type"" = Record.FromTable(Table.TransformColumnTypes(Source,{{""Name"", type text}, {""Value"", type any}})),
    parseDates = #""Changed Type""&[dateFrom=Date.ToText(Date.From(#""Changed Type""[dateFrom]), ""d/M/yyyy""), returnFrom=Date.ToText(Date.From(#""Changed Type""[returnFrom]), ""d/M/yyyy""), dateTo=Date.ToText(Date.From(#""Changed Type""[dateTo]), ""d/M/yyyy""), returnTo=Date.ToText(Date.From(#""Changed Type""[returnTo]), ""d/M/yyyy"")],
    BuildQueryString = Uri.BuildQueryString(parseDates)
in
    BuildQueryString;

shared AllSearchResults = let
    Source = Json.Document(Web.Contents(""https://api.skypicker.com/flights?""&QueryString)),
    data = Table.FromRecords(Source[data]),
    #""Removed Columns"" = Table.RemoveColumns(data,{""routes"", ""countryFrom"", ""countryTo"", ""dTimeUTC"", ""aTimeUTC"", ""cityFrom"", ""cityTo"", ""mapIdfrom"", ""mapIdto"", ""nightsInDest"", ""virtual_interlining"", ""fly_duration"", ""return_duration"", ""facilitated_booking_available"", ""type_flights"", ""found_on"", ""conversion"", ""booking_token""}),
    #""Parse Columns"" = Table.TransformColumns(#""Removed Columns"", {{""route"", Table.FromRecords}, {""bags_price"", each if _ = [] then null else Record.FieldValues(_)}, {""baglimit"", each if _ = [] then null else _}, {""transfers"", each if List.IsEmpty(_) then null else _}, {""duration"", each Record.TransformFields(_, List.Transform(Record.FieldNames(_), (y) => {y, (x) => x / 3600 / 24}))}}),
    #""Expanded duration"" = Table.ExpandRecordColumn(#""Parse Columns"", ""duration"", {""departure"", ""return"", ""total""}, {""fly_duration"", ""return_duration"", ""total_duration""}),
    #""Changed duration type"" = Table.TransformColumnTypes(#""Expanded duration"",{{""fly_duration"", type duration}, {""return_duration"", type duration}, {""total_duration"", type duration}}),
    #""Expanded availability"" = Table.ExpandRecordColumn(#""Changed duration type"", ""availability"", {""seats""}, {""seats""}),
    ToJson = List.Transform(List.Transform(Table.ToRecords(#""Expanded availability""), Json.FromValue), Text.FromBinary)
in
    ToJson;

shared SearchResults = let
    Source = Table.FromRecords(List.Transform(AllSearchResults, Json.Document)),
    #""route as Table"" = Table.TransformColumns(Source, {""route"", Table.FromRecords}),
    #""Added OnTimePerformance to route"" = Table.TransformColumns(#""route as Table"", {""route"", each Table.FromRecords(Table.TransformRows(_, (y) => y&[OnTimePerformance=((Table.First(Table.SelectRows(OnTimePerformanceHC, (x) => x[From] = y[flyFrom] and x[#""Flight Number""] = Number.FromText(y[operating_flight_no]) and x[IATA] = y[operating_carrier]), null)))]))}),
    #""Parse OnTimePerformance"" = Table.TransformColumns(#""Added OnTimePerformance to route"", {""route"", each Table.TransformColumns(_, {""OnTimePerformance"", each if _ = null then null else Json.Document([OnTimePerformance])})}),
    ToJson = List.Transform(List.Transform(Table.ToRecords(#""Parse OnTimePerformance""), Json.FromValue), Text.FromBinary)
in
    ToJson ;

shared SearchResultsHC = let
    Source = Excel.CurrentWorkbook(){[Name=""SearchResults""]}[Content][SearchResults]
in
    Source;

shared SearchResultsWithOnTimePerformance = let
    Source = Table.FromRecords(List.Transform(SearchResultsHC, Json.Document)),
    #""route as Table"" = Table.TransformColumns(Source, {{""route"", Table.FromRecords}, {""aTime"", dateTimeFromUnixTime}, {""dTime"", dateTimeFromUnixTime}}),
    route = Table.SelectRows(#""route as Table"", each List.AllTrue(List.Transform(List.Select([route][OnTimePerformance], each _ <> null), each _[details][overall][stars] >= Input[stars]))),
    #""Select route columns"" = Table.TransformColumns(route, {""route"", each Table.TransformColumns(Table.SelectColumns(_, {""aTime"", ""dTime"", ""flyTo"", ""flyFrom"", ""airline"", ""flight_no"", ""fare_classes"", ""OnTimePerformance""}), {""OnTimePerformance"", each if _ = null then _ else Record.SelectFields(_, {""details""})})}),
    #""Required OnTimePerformance"" = Table.SelectRows(#""Select route columns"", each List.Count(List.Select([route][OnTimePerformance], each _ <> null)) / List.Count([route][OnTimePerformance]) >= Input[OnTimeCoverage]),
    #""Convert route to json"" = Table.TransformColumns(#""Required OnTimePerformance"", {""route"", each Text.FromBinary(Json.FromValue(_))}),
    #""Removed Other Columns"" = Table.SelectColumns(#""Convert route to json"",{""id"", ""price"", ""dTime"", ""aTime"", ""fly_duration"", ""return_duration"", ""total_duration"", ""quality"", ""route""}),
    #""Changed Type"" = Table.TransformColumnTypes(#""Removed Other Columns"",{{""price"", Int64.Type}, {""dTime"", type datetime}, {""aTime"", type datetime}, {""fly_duration"", type duration}, {""return_duration"", type duration}, {""total_duration"", type duration}, {""quality"", type number}, {""route"", type text}})
in
    #""Changed Type"";

shared FlightBlacklist = let
    Source = Table.FromRecords(Json.Document(File.Contents(Input[FlightBlacklist]))),
    #""Changed Type"" = Table.TransformColumnTypes(Source,{{""IATA"", type text}, {""Flight Number"", Int64.Type}, {""From"", type text}})
in
    #""Changed Type"";

shared SearchResultDetail = let
    Source = SearchResultsWithOnTimePerformance,
    #""Filtered Rows"" = Table.SelectRows(Source, each ([id] = Input[id])),
    #""Removed Other Columns"" = Table.SelectColumns(#""Filtered Rows"",{""price"", ""quality"", ""route""}),
    #""Route as Json"" = Table.TransformColumns(#""Removed Other Columns"", {""route"", Json.Document}),
    #""Expanded route"" = Table.ExpandListColumn(#""Route as Json"", ""route""),
    #""Expanded route1"" = Table.ExpandRecordColumn(#""Expanded route"", ""route"", {""aTime"", ""dTime"", ""flyTo"", ""flyFrom"", ""flight_no"", ""fare_classes"", ""OnTimePerformance""}, {""aTime"", ""dTime"", ""flyTo"", ""flyFrom"", ""flight_no"", ""fare_classes"", ""OnTimePerformance""}),
    #""Expanded OnTimePerformance"" = Table.ExpandRecordColumn(#""Expanded route1"", ""OnTimePerformance"", {""details""}, {""details""}),
    #""Expanded details"" = Table.ExpandRecordColumn(#""Expanded OnTimePerformance"", ""details"", {""overall"", ""otp"", ""delayPerformance""}, {""overall"", ""otp"", ""delayPerformance""}),
    #""Expanded overall"" = Table.ExpandRecordColumn(#""Expanded details"", ""overall"", {""stars"", ""ontimePercent"", ""delayMean""}, {""stars"", ""ontimePercent"", ""delayMean""}),
    #""Expanded otp"" = Table.ExpandRecordColumn(#""Expanded overall"", ""otp"", {""stars""}, {""ontimeStars""}),
    #""Expanded delayPerformance"" = Table.ExpandRecordColumn(#""Expanded otp"", ""delayPerformance"", {""stars""}, {""stars.1""}),
    #""Renamed Columns"" = Table.RenameColumns(#""Expanded delayPerformance"",{{""stars.1"", ""delayStars""}}),
    #""Parse UnixTime"" = Table.TransformColumns(#""Renamed Columns"", {{""aTime"", dateTimeFromUnixTime}, {""dTime"", dateTimeFromUnixTime}}),
    #""Changed Type"" = Table.TransformColumnTypes(#""Parse UnixTime"",{{""aTime"", type datetime}, {""dTime"", type datetime}, {""flyTo"", type text}, {""flyFrom"", type text}, {""flight_no"", Int64.Type}, {""fare_classes"", type text}, {""stars"", type number}, {""ontimePercent"", Int64.Type}, {""delayMean"", Int64.Type}, {""ontimeStars"", type number}, {""delayStars"", type number}})
in
    #""Changed Type"";

shared dateTimeFromUnixTime = let
    Source = (unixTime) => #datetime(1970, 1, 1, 0, 0, 0) + #duration(0, 0, 0, unixTime)
in
    Source;"
