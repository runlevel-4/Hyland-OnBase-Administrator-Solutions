# Begin Log
Start-Transcript -Path "\\path\to\log\file.log"

# You will need to uncomment this if you do not have the Powershell SQL Server module installed
#install-module -Name SqlServer

# create friendly name tags from the xml data to make it readable in the csv output
$tag_names = @{

    # add your xml fields here from your XML file
    "fname" = "First Name"
    "lname" = "Last Name"
    "emailadd" = "Email"
    "contractid" = "Contract Number"
    #...you get the idea

}

# get SQL records for each contract doument and its subsequent documents and store it in an array @()
# NOTE: you can swap out the contract doc type and joined kw type columns for whatever doc type/kw type you are trying to extract xml data for (in this case it was legal contracts)

$query = @"

set transaction isolation level read uncommitted;

-- We added a CTE because we needed to filter for in-progress contracts later on while keeping other data intact (you may not need that depending again on your use case)
with all_contracts as (
	select distinct
		id.itemnum [doc_handle]
-- CASE statement because have a lot of ugly characters in our document names
		,case
		    when dt.itemtypename = 'Contract File' then  rtrim(replace(rtrim(replace(rtrim(replace(rtrim(replace(rtrim(replace(rtrim(replace(rtrim(replace(rtrim(replace(rtrim(replace(rtrim(replace(rtrim(replace(rtrim(replace(rtrim(replace(rtrim(replace(rtrim(replace(ki###.keyvaluechar, '/', '')), '(', '')), ')', '')), '.', '')), ' ', '')), '+', '')), '-', '')), ',', '')), '_', '')), ':', '')), '<B>', '')), '''', '')), '"', '')), '*', '')), '?', ''))
		else rtrim(replace(rtrim(replace(rtrim(replace(rtrim(replace(rtrim(replace(rtrim(replace(rtrim(replace(rtrim(replace(rtrim(replace(rtrim(replace(rtrim(replace(rtrim(replace(rtrim(replace(rtrim(replace(rtrim(replace(id.itemname, '/', '')), '(', '')), ')', '')), '.', '')), ' ', '')), '+', '')), '-', '')), ',', '')), '_', '')), ':', '')), '<B>', '')), '''', '')), '"', '')), '*', '')), '?', ''))
		end [doc_name]
		,dt.itemtypename [doc_type]
		,id.itemdate [doc_date]
		,pp.lastuseddrive [parent_path]
		,idp.filepath [sub_path_file]
		,concat(rtrim(pp.lastuseddrive),rtrim(idp.filepath)) [full_path]
		,ki###.keyvaluebig [contract_number]
		--,ki###.keyvaluechar [related_contract_id]
		,ki###.keyvaluechar [status]
	from hsi.itemdata id
	inner join hsi.doctype dt
		on dt.itemtypenum = id.itemtypenum
		and dt.itemtypenum in (###,###,###) -- add doc type IDs here
	left join hsi.keyitem### ki### -- Contract ID\Amendment Number, but again you can use whatever kw type num you want to join on for whatver doc type you are after
		on ki###.itemnum = id.itemnum
	left join hsi.keyitem### ki### -- Related Contract ID on additional docs
		on ki###.keyvaluechar = convert(varchar,ki###.keyvaluebig)
	left join hsi.keyitem### ki### -- Contract Status
		on ki###.itemnum = id.itemnum
			and ki###.keyvaluechar in ('EXECUTED') --('PENDING','EXECUTED')
	left join hsi.keyitem### ki### -- Contract Title
		on ki###.itemnum = id.itemnum
	inner join hsi.itemdatapage idp
		on idp.itemnum = id.itemnum
			and id.maxdocrev = idp.docrevnum
			and idp.filetypenum <> 1
	inner join hsi.diskgroup dg
		on dg.diskgroupnum = idp.diskgroupnum
	inner join hsi.physicalplatter pp
		on pp.diskgroupnum = dg.diskgroupnum
	where id.status <> 16 -- exclude deleted docs
    --and id.itemdate between '####-##-##' and '####-##-##'

)

select *
from all_contracts
where contract_number in (
							select contract_number
							from all_contracts
							where all_contracts.status in ('EXECUTED') -- ('PENDING','EXECUTED')
						)

and (
    -- add additional SQL filters here
)

order by contract_number desc, doc_type asc;

"@

$sql_results = @(Invoke-SqlCmd -ServerInstance "DBINSTANCENAME" -Database "DBNAME" -Query $query -TrustServerCertificate)

Write-host "Grabbed the file records from the database...`n"

# left this here in case you want to pause after the SQL query runs or errors occur
#pause

$output_dir = '\\path\to\xml_data' # main csv output containing xml data
$output_dir_files = '\\path\to\additional\attached\documents' # other non-xml files tied to the document
$xml_full_records = @()
$htm_full_records = @()

# check for output directory existence
if (-not(Test-Path -Path "$($output_dir_files)")) {
        
        Write-Host "$($output_dir_files) does not exist.`n" -ForegroundColor Yellow
            
        try {
            New-Item -Path "$($output_dir_files)" -ItemType Directory > $null -Force
            Write-Host "Created directory: $($output_dir_files)`n" -ForegroundColor Green
        }
        catch {
            Write-Host "Could not create directory!`n" -ForegroundColor Red
            break
        }
}

# evaluate each SQL record for its filepath, filename, file extension and contract number
# anywhere that mentions the $sql_contract_num variable, you can replace it with the SQL column for the doc type/unique kw type that you want
foreach ($row in $sql_results) {

    $sql_doc_name = $row.doc_name
    $sql_file_path = $row.full_path
    $sql_file_name = Split-Path -Path $row.full_path -Leaf
    $sql_file_ext = (Get-Item $sql_file_path).Extension
    $sql_contract_num = $row.contract_number

    # check filetypes
    # if the file is not a .xml file then just copy it to the additional output directory
    if ($sql_file_ext -ne ".xml") {

        Write-Host "No conversion needed...copying $($sql_file_name) to $($output_dir_files)\$($sql_contract_num)_$($sql_doc_name)_$($sql_file_name)`n" -ForegroundColor Cyan
        Copy-Item -Path $sql_file_path -Destination "$($output_dir_files)\$($sql_contract_num)_$($sql_doc_name)_$($sql_file_name)"
    
    }

    #----- XML FILES -----#

    # if it is a .xml file then use magic to prep it for csv conversion...we love magic
    elseif ($sql_file_ext -eq ".xml") {

        Write-Host "Converting $($sql_file_name) to the csv...`n" -ForegroundColor Magenta

        # get xml file contents
        [xml]$xml_file = Get-Content -Path $sql_file_path

        # create arrays to store xml contents for csv prep
        $xml_nodes = @()
        $xml_nodes += @($xml_file.form.field)
        $xml_nodes += @($xml_file.form.repeater.repeateritem.field)

        # create an ordered dictionary to hold the pivoted data
        $csv_object = [ordered]@{}

        # create a dictonary to store xml <name> tags
        $countTable = @{}

        # loop through each xml node
        foreach ($node in $xml_nodes) {

            # check for a valid <name> tag
            if ($node -and $node.name) {

                # loop through each xml <name> tag to translate them to user-friendly tags
                $xml_names_cleanup = foreach ($name in $node.name) {
                    if ($tag_names.ContainsKey($name)) {
                        $new_node_name = $tag_names[$name]
                        $new_node_name
                    }
                    #catch any <name> tags that might not be in the $tag_names @{} dictionary so you can go back and add them to give them prettier names
                    else {
                        $name
                    }
                }

                # rename repeating <name> tags for uniqueness using the $countTable @{} dictionary from earlier, by appending numbers to the end of the repeaters so that Excel is happy with unique header names
                $xml_name_rename = foreach ($n in $xml_names_cleanup) {
                    if ($countTable.ContainsKey($n)) {
                        $countTable[$n]++
                        "${n}_$($countTable[$n])"
                    }
                    else {
                        $countTable[$n] = 1
                        $n
                    }
                }

                # get the inner text/string value of <name> and <value> tags
                $csv_object[$xml_name_rename] = $node.value
            }
        }

        # gather the xml data into a memory object to gather all header info for the final csv
        $xml_full_records += [PSCustomObject]$csv_object
    }
}

# final csv packaging after all xml tags have been gathered in memory
if ($xml_full_records) {

    Write-Host "Adding data to $($output_dir)\xml_contracts.csv" -ForegroundColor Cyan

    # gather every unique property name found across all processed files
    $xml_headers = $xml_full_records | ForEach-Object { $_.PSObject.Properties.Name } | Select-Object -Unique

    # stuff the xml data into the csv file
    $xml_full_records | Select-Object -Property $xml_headers | Export-Csv -Path "$($output_dir)\xml_contracts.csv" -NoTypeInformation -Force
    
    Write-Host "XML Extract Complete!" -ForegroundColor Green
}

# End Log
Stop-Transcript