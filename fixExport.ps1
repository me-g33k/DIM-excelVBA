function fixExport() {
	
	###
	### Main Function
	###

	param ( [string]$curFile )
	$armorHeader="Name, Tag, Tier, Type, Equippable, Light, Owner, % Leveled, Locked, Equipped, Year, % Quality, % IntQ, % DiscQ, % StrQ, Int, Disc, Str, Notes, Perk01, Perk02, Perk03, Perk04, Perk05, Perk06"
	$weaponsHeader="Name, Tag, Tier, Type, Light, Dmg, Owner, % Leveled, Locked, Equipped, Year,AA, Impact, Range, Stability, ROF, Reload, Mag, Equip, Notes, Nodes, Node01, Node02, Node03, Node04, Node05, Node06, Node07, Node08, Node09, Node10, Node11, Node12, Node13, Node14, Node15, Node16"
	# get-content $curFile | where {$_.readcount -lt 2} > sorted.csv
	
	switch ($curFile) {
		"destinyWeapons.csv" {
			$weaponsHeader | Out-File sorted.csv
			(Get-Content $curFile).replace('Ã¤', 'ä') | Set-Content $curFile
		}
		"destinyArmor.csv" {
			$armorHeader | Out-File sorted.csv
		}
	}
	
	get-content $curFile | where {$_.readcount -gt 1} | sort >> sorted.csv
	# gci $curFile | rename-item -Newname "oldFile.csv"
	remove-item $curFile
	gci sorted.csv | rename-item -Newname $curFile
}
