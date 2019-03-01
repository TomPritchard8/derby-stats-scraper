param([string]$team="") #Must be the first statement in the script

# Excel Column Mapping
$numberColumn = 1;
$nameColumn = 2;
#----
$MisconductColumn = 4;
$HighBlockColumn = 5;
$BackBlockColumn = 6;
$LowBlockColumn = 7;
$LegBlockColumn = 8;
$ForearmsColumn = 9;
$HeadBlockColumn = 10;
$MultiplayerColumn = 11;
$IllegalContactColumn = 12;
$DirectionColumn = 13;
$IllegalPositionColumn = 14;
$CutColumn = 15;
$InterferenceColumn = 16;
$IllegalProcedureColumn = 17;
#----
$penaltyColumn = 21;
$ExpulsionColumn = 22;
$jamsColumn = 23;

# Initialization
$skaterNumber = @{}
$skaterName = @{}
$skaterPenalties = @{}
$skaterJams = @{}
$gamePenalties = @{}
$statbooks = gci $strPath | Where-Object {$_.FullName -like "*xlsx*"};
$objExcel = New-Object -ComObject Excel.Application
$objExcel.Visible = $false

# Initialize Hashes
#$skaterPenaltyHashtable = @{}
$teamPenaltyData = @{}

# Find the right row for the team stats
function Get-StartRow ($worksheet) { # Set the starting cell based on the team name
	if ($worksheet.cells.item(2,1).value2.ToLower().Contains($team)) # References are in (row,col) format
	{
		return 4;
	}
	else
	{
		return 33;
	}
}


# Check there are Stats Books!
if ($statbooks.length -eq 0)
{
	Write-Host "No Stats Books found!";
	continue;
}

# Look in each Stats Book
foreach ($statbook in $statbooks)
{
	$workbook = $objExcel.Workbooks.Open($statbook.FullName)
	$worksheet = $workbook.sheets.item("Penalty Summary")
	$startRow = Get-StartRow($worksheet)
	
	# Loop through each row
	for ($row = $startRow; $row -le $startRow + 19; $row++)
	{
		# Get and clean skater number
		$skaterNumber = $worksheet.cells.item($row, $numberColumn).value2.ToString().Trim();  
		if (($skaterNumber -eq $null) -or ($skaterNumber -eq ""))
        {
            continue;
        }
        $skaterNumber = $skaterNumber -replace '\*',''; # Clean up the skater number
        
        # Get skater's game summary data		
        $skaterName = $worksheet.cells.item($row, $nameColumn).value2
		$penaltiesThisGame = $worksheet.cells.item($row, $penaltyColumn).value2
		$jamsThisGame = $worksheet.cells.item($row, $jamsColumn).value2
		
		# Get skater's detailed penalty data for this game
		for ($i=0; $i -lt 15; $i++)
		{
			$gamePenalties[$i] = $worksheet.cells.item($row, $i+4).value2;
		}	
				
		# Add skater's penalty info to their overall record
		if ($teamPenaltyData.ContainsKey($skaterNumber))
		{		
			$skaterPenaltyHashtable = $teamPenaltyData.Get_Item($skaterNumber);
			$skaterPenaltyHashtable.'Mis' += $gamePenalties[0];
			$skaterPenaltyHashtable.'HiB' += $gamePenalties[1];
			$skaterPenaltyHashtable.'BaB' += $gamePenalties[2];
			$skaterPenaltyHashtable.'LoB' += $gamePenalties[3];
			$skaterPenaltyHashtable.'LeB' += $gamePenalties[4];
			$skaterPenaltyHashtable.'FAr' += $gamePenalties[5];
			$skaterPenaltyHashtable.'HeB' += $gamePenalties[6];
			$skaterPenaltyHashtable.'MPB' += $gamePenalties[7];
			$skaterPenaltyHashtable.'ICo' += $gamePenalties[8];
			$skaterPenaltyHashtable.'Drn' += $gamePenalties[9];
			$skaterPenaltyHashtable.'IPo' += $gamePenalties[10];
			$skaterPenaltyHashtable.'Cut' += $gamePenalties[11];
			$skaterPenaltyHashtable.'Int' += $gamePenalties[12];
			$skaterPenaltyHashtable.'IPr' += $gamePenalties[13];
			$skaterPenaltyHashtable.'PTot' += $penaltiesThisGame;
			$skaterPenaltyHashtable.'JTot' += $jamsThisGame;
		}		
					
		else # Initialize if this is the first time we've seen the skater listed
		{			
			$newSkaterPenaltyHashtable = @{			
				'Mis' = $gamePenalties[0];
				'HiB' = $gamePenalties[1];
				'BaB' = $gamePenalties[2];
				'LoB' = $gamePenalties[3];
				'LeB' = $gamePenalties[4];
				'FAr' = $gamePenalties[5];
				'HeB' = $gamePenalties[6];
				'MPB' = $gamePenalties[7];
				'ICo' = $gamePenalties[8];
				'Drn' = $gamePenalties[9];
				'IPo' = $gamePenalties[10];
				'Cut' = $gamePenalties[11];
				'Int' = $gamePenalties[12];
				'IPr' = $gamePenalties[13];
				'PTot' = $penaltiesThisGame;
				'JTot' = $jamsThisGame
			};					
			$penaltyHashtable = New-Object -TypeName PSObject -Prop $newSkaterPenaltyHashtable;
			$teamPenaltyData.Add($skaterNumber, $penaltyHashtable)
		}
	}
	
    $workbook.close($False);
	    
}

$objExcel.quit();

# Show something for all our work
$teamPenaltyData.GetEnumerator() | Select-Object -Property Name, @{Name="JTot";Expression={$_.Value.JTot}}, @{Name="PTot";Expression={$_.Value.PTot}}, 
@{Name="PenPJ";Expression={[math]::Round(($_.Value.PTot / $_.Value.JTot),2)}},
@{Name="Mis";Expression={$_.Value.Mis}}, @{Name="HiB";Expression={$_.Value.HiB}}, 
@{Name="BaB";Expression={$_.Value.BaB}}, @{Name="LoB";Expression={$_.Value.LoB}}, 
@{Name="LeB";Expression={$_.Value.LeB}}, @{Name="For";Expression={$_.Value.FAr}}, 
@{Name="HeB";Expression={$_.Value.HeB}}, @{Name="MPB";Expression={$_.Value.MPB}}, 
@{Name="ICo";Expression={$_.Value.ICo}}, @{Name="Dir";Expression={$_.Value.Drn}},
@{Name="IPo";Expression={$_.Value.IPo}}, @{Name="Cut";Expression={$_.Value.Cut}}, 
@{Name="Int";Expression={$_.Value.Int}}, @{Name="IPr";Expression={$_.Value.IPr}} | Sort-Object -Property PenPJ -descending | Format-Table *

# Export the results to CSV too
$outputPath = ".\" + $team + "PenaltySummary" + ".csv"
$teamPenaltyData.GetEnumerator() | Select-Object -Property Name, @{Name="JTot";Expression={$_.Value.JTot}}, @{Name="PTot";Expression={$_.Value.PTot}}, 
@{Name="PenPJ";Expression={[math]::Round(($_.Value.PTot / $_.Value.JTot),3)}},
@{Name="Mis";Expression={$_.Value.Mis}}, @{Name="HiB";Expression={$_.Value.HiB}}, 
@{Name="BaB";Expression={$_.Value.BaB}}, @{Name="LoB";Expression={$_.Value.LoB}}, 
@{Name="LeB";Expression={$_.Value.LeB}}, @{Name="For";Expression={$_.Value.FAr}}, 
@{Name="HeB";Expression={$_.Value.HeB}}, @{Name="MPB";Expression={$_.Value.MPB}}, 
@{Name="ICo";Expression={$_.Value.ICo}}, @{Name="Dir";Expression={$_.Value.Drn}},
@{Name="IPo";Expression={$_.Value.IPo}}, @{Name="Cut";Expression={$_.Value.Cut}},  
@{Name="Int";Expression={$_.Value.Int}}, @{Name="IPr";Expression={$_.Value.IPr}} | Sort-Object -Property PenPJ -descending | export-csv -Path $outputPath -NoTypeInformation
Write-Host "Write output to $outputPath";
