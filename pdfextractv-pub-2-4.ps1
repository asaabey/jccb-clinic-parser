

# Known Bugs
# DOB not detected if its in second line and preceded by string-FIXED 

$iTextSharpDllPath="iTextSharp.dll"

Unblock-File -Path $iTextSharpDllPath

Add-Type -Path $iTextSharpDllPath
$script:dllPath = $MyInvocation.MyCommand.Path


#Environment variables

$wd="C:\Users\asabe\Desktop\clinic-pilot"
$outputDir="C:\Users\asabe\Desktop\clinic-pilot\"
$CompositeFile=$outputDir + 'Composite.csv'
$dtc=New-Object System.Data.DataTable("composite")
$cols=@("HRN","DOB","NAME","LOCALITY","CLINIC","COMPONENT","DATE","VALUE")
    foreach ($col in $cols) {
	    $dtc.Columns.Add($col) | Out-Null
    }

# Counters
$RecCntDist=0
$RecCnt=0
$RecDiscard=0
$PgCnt=0
$PgDiscard=0

# Anchors
$idAnchor=0


function Get-PDFContent2 {param([Parameter(Mandatory=$true,Position=1)]$pdfFile,[Parameter(Mandatory=$true,Position=2)][int]$pageNumber)    
    
    $reader = New-Object iTextSharp.text.pdf.PdfReader $pdfFile
    $strategy = New-Object iTextSharp.text.pdf.parser.SimpleTextExtractionStrategy
    [iTextSharp.text.pdf.parser.PdfTextExtractor]::GetTextFromPage($reader, $pageNumber, $strategy)
    $reader.Close()
}

function Format-Name([string]$t){
    return $t.replace(","," ").Trim().ToUpper()
}

function Parse-Text ([string] $txtFile) {    
    #param([Parameter(Mandatory=$true)] $txtFile) 
    #if(Test-Path $txtFile){
    if($txtFile -ne ""){
            $dt=New-Object System.Data.DataTable("clients")
            $dt2=New-Object System.Data.DataTable("locality")
            $dt3=New-Object System.Data.DataTable("OrphanedDOB")
            $dt4=New-Object System.Data.DataTable("OrphanedClinic")
            $dt5=New-Object System.Data.DataTable("OrphanedValue")
    
            $cols=@("HRN","DOB","NAME","LOCALITY","CLINIC","COMPONENT","DATE","VALUE")
            foreach ($col in $cols) {
	            $dt.Columns.Add($col) | Out-Null
            }
            $cols=@("HRN","LOCALITY")
            foreach ($col in $cols) {
	            $dt2.Columns.Add($col) | Out-Null
            }
            $cols=@("HRN","DOB","NAME_PART")
            foreach ($col in $cols) {
	            $dt3.Columns.Add($col) | Out-Null
            }
            $cols=@("HRN","CLINIC")
            foreach ($col in $cols) {
	            $dt4.Columns.Add($col) | Out-Null
            }
            $cols=@("HRN","VALUE")
            foreach ($col in $cols) {
	            $dt5.Columns.Add($col) | Out-Null
            }
    
            $rgxHRN="^[0-9]{7}"
            $rgxDate="^\d{2}\/\d{2}\/\d{4}"
            $rgxDate0="\d{2}\/\d{2}\/\d{4}"
            $rgxPostCode="\s\d{4}.$"
            #$rgxTime="^([0-2][0-9]):([0-5][0-9])$/gm"
            $rgxTime="^(([0-2][0-9]):([0-5][0-9]))"
            
            
            #$t=Get-content $txtFile
            [string[]]$t=$txtFile -split '[\r\n]'

            if($t.Length -lt 10){
                #write-host "PDF parsed but no text " -ForegroundColor DarkYellow
                $global:PgDiscard++
            } else {
            $global:PgCnt++
            $ClinDate=$t[0].Substring(4).Split(",")
    
            # Get profile from file header
            [string]$ClinProfile=$t[3].Trim()       
            if($ClinProfile -like "CLINICIAN*"){
                [string]$ClinVal=$t[1]
                if($t[2] -ne "for all Session Types"){
                    [string]$ClinLoc=$t[2]
                } else {
                    [string]$ClinLoc=""
                }
                
                
            } Elseif($ClinProfile -like "LOCATION*"){
                [string]$ClinLoc=$t[1]
                [string]$ClinVal=""
            } 

    

            
    
            $ClinDate4=Get-Date ($ClinDate[1]) -Format 'dd/MM/yyyy'
    
            $ClinLoc2=$ClinLoc.Split(",")
            $ClinLoc3=$ClinLoc2[0].Trim()
            [string]$IndexHRN=$null 
            [int]$counter=0
            [int] $c2=0
            [int] $lineCount=0
            
            


            # Parse lines
            foreach ($line in $t)
            {
                
                $global:IdAnchor++
                #Detect Identity line
                if ($line -match $rgxHRN)
                {
                    [string]$HRN=$line.Substring(0,7)
                    [string]$DOBTemp=$line.Substring($line.Length-10,10)
                    #Write-Host $DOBTemp
                    if( $DOBTemp -match $rgxDate){
                        [DateTime]$DOBTemp1=[DateTime]::Parse($DOBTemp)
                        [string]$DOB=$DOBTemp1.ToString("dd/MM/yyyy")
                        
                    } else {
                        [string]$DOB=$null
                    }
            
                    # check if DOB is in same line
                    if($DOB -eq ""){
                        [int]$rTrm=8
                    } else {
                        [int]$rTrm=18
                    }
                    
                    [string]$NAME=$line.Substring(8,($line.Length-$rTrm))
            
                    $row=$dt.NewRow()   
                    $row["HRN"]=$HRN
                    $IndexHRN=$HRN
                    $row["NAME"]=$NAME
                    $row["DOB"]=$DOB
            
                    $row["CLINIC"]=$ClinLoc3    

                    if($ClinLoc3 -like "TELE%"){
                        $row["COMPONENT"]="BOOKING_TELEHELATH"
                    } else {
                        $row["COMPONENT"]="BOOKING_CLINIC"
                    }

            
            
            
                    #$row["DATE"]=$ClinDate4.ToShortDateString()
                    $row["DATE"]=$ClinDate4 

                    $row["VALUE"]=$ClinVal
                    $dt.Rows.Add($row)
                    $counter++

                    $global:IdAnchor=1
            
                }


                #Detect Locality
                if($line.Length -gt 10)
                {
                    $cond2=$line.Substring($line.Length-6)

                    if ($cond2 -like ',*08*')
                    {
                        $address=$line.Split(",")

                        $addressPostCode=$address[$address.GetUpperBound(0)]
                        [string]$addressSuburb=$address[$address.GetUpperBound(0)-1]
                        $row=$dt2.NewRow()
                        $row["HRN"]=$IndexHRN
                        $row["LOCALITY"]=$addressSuburb.Trim()
                        $dt2.Rows.Add($row)
                
                    }
        
                }

                #Detect orphaned DOB
                
                if($line -match $rgxDate){
                    $row=$dt3.NewRow()
                    $row["HRN"]=$IndexHRN
                    $row["DOB"]=$line.Substring(0,10)
                    $dt3.Rows.Add($row)
                
                }

                #Detect orphaned DOB -2nd variant
                #part of name is at the front and this will be extracted as well

                if($line -match $rgxDate0 -and $global:IdAnchor -eq 2 ){
                    $row=$dt3.NewRow()
                    $row["HRN"]=$IndexHRN
                    $row["DOB"]=$line.Substring(($line.length-10),10)
                    $row["NAME_PART"]=$line.Substring(0,($line.length-10))
                    $dt3.Rows.Add($row)
                
                }



                #Detect orphaned clinic , when clinician list parsed
                if($line -like "112*"){
                    $row=$dt4.NewRow()
                    $row["HRN"]=$IndexHRN
                    $row["CLINIC"]="TELEHEALTH"
                    $dt4.Rows.Add($row)
                }

                #Detect new record by time signature
                if($line -match $rgxTime){
                    #Detect orphaned value
                    #using fixed negative offset; will not work for last record
                    if($lineCount -gt 0){
                        $s=$t[($lineCount-1)]  
                        $junk="Time Patient Services Other Resources Att"
                        
                        $row=$dt5.NewRow()
                        $row["HRN"]=$IndexHRN                   
                            
                        if($s -ne $junk){
                            $row["VALUE"]=$s    
                        } else {
                            $row["VALUE"]=""
                        }
                        $dt5.Rows.Add($row)
                    }
                    $c2++
                }
                $lineCount++    
    
    }
    
    $global:RecCnt+=$counter

    

    

    #Join orphaned datasets

    for($i=0;$i -lt $dt.Rows.Count;$i++){
        for($j=0;$j -lt $dt2.Rows.Count;$j++){
            if($dt.Rows[$i]["HRN"] -eq $dt2.Rows[$j]["HRN"]){
                 $dt.rows[$i]["LOCALITY"]=$dt2.rows[$j]["LOCALITY"].ToString().ToUpper().Replace(",","")
            } 
        }
        # use interpolation for missing locality
        if($dt.rows[$i]["LOCALITY"].ToString() -eq ""){
            if($dt2.rows.Count -gt 0){
                $dt.rows[$i]["LOCALITY"]=$dt2.rows[0]["LOCALITY"].ToString().ToUpper().Replace(",","")    
            }
        }
        # join orphaned DOB
        if($dt.rows[$i]["DOB"].ToString() -eq ""){
            for($j=0;$j -lt $dt3.Rows.Count;$j++){
                if($dt.Rows[$i]["HRN"] -eq $dt3.Rows[$j]["HRN"]){
                 [DateTime]$DOBTemp0=[DateTime]::Parse($dt3.rows[$j]["DOB"]) 
                 [string]$DOB=$DOBTemp1.ToString("dd/MM/yyyy")
                 if($dt3.Rows[$j]["NAME_PART"] -ne ""){
                    $dt.Rows[$i]["NAME"]=$dt.Rows[$i]["NAME"]+", "+$dt3.Rows[$j]["NAME_PART"]
                 }
                 $dt.rows[$i]["DOB"]=$DOB
                } 
            }            
        }
        # join orphaned clinic
        # Clinic imputation method 1
        if($dt.rows[$i]["CLINIC"].ToString() -eq ""){
            for($j=0;$j -lt $dt4.Rows.Count;$j++){
                if($dt.Rows[$i]["HRN"] -eq $dt4.Rows[$j]["HRN"]){
                 
                 $dt.rows[$i]["CLINIC"]=$dt4.Rows[$j]["CLINIC"].ToString()
                } 
            }            
        }
        # join orphaned value
        # value imputation 
        if($dt.rows[$i]["VALUE"].ToString() -eq ""){
            for($j=0;$j -lt $dt5.Rows.Count;$j++){
                if($dt.Rows[$i]["HRN"] -eq $dt5.Rows[$j]["HRN"]){
                    $s=$dt5.Rows[$j]["VALUE"].ToString()
                    if($s -ne ""){
                        $dt.rows[$i]["VALUE"]=$s
                    } else {
                        $dt.rows[$i]["VALUE"]=$dt5.Rows[($j-1)]["VALUE"].ToString()       
                    } 
                 
                } else {
                # no match, usually the last line 
                    if($i -gt 0){
                        #impute from previous value
                        $dt.rows[$i]["VALUE"]=$dt.rows[($i-1)]["VALUE"]
                    }
                
                }
            }            
        }
        # final formatting name
        $dt.Rows[$i]["NAME"]=Format-Name($dt.Rows[$i]["NAME"])

        # final validation
        if($dt.Rows[$i]["DOB"].ToString().Length -lt 4){
            $dt.Rows[$i].Delete()
            $global:RecDiscard++

        }

    }


 

    # Clinic imputation 2
    
        if($dtfmain.Rows.Count -gt 1){
            # Empty clinic
            if($dtfmain.Rows[0]["CLINIC"] -eq ""){
                
                [string[]] $ClnArr = @()
                for($i=0;$i -lt $dtfmain.Rows.Count;$i++){
                    $ClnArr += $dtfmain.Rows[$i]["LOCALITY"].ToString()
                }
                
                # Find mode
                $ClinStr=$ClnArr | Group-Object | sort -Descending count | select-object -first 1 
                
                for($i=0;$i -lt $dtfmain.Rows.Count;$i++){
                    $dtfmain.Rows[$i]["CLINIC"]=$ClinStr.Name
                }
            
            }

            # Telehealth 
            if($dtfmain.Rows[0]["CLINIC"] -eq "TELEHEALTH"){
                
                [string[]] $ClnArr = @()
                for($i=0;$i -lt $dtfmain.Rows.Count;$i++){
                    $ClnArr += $dtfmain.Rows[$i]["LOCALITY"].ToString()
                }
                # Find mode
                $ClinStr=$ClnArr | Group-Object | sort -Descending count | select-object -first 1 
                
                for($i=0;$i -lt $dtfmain.Rows.Count;$i++){
                    $dtfmain.Rows[$i]["CLINIC"]="TELE-"+$ClinStr.Name
                }
            
            }


    [bool] $distinct="TRUE"
    $dtf=$dt.DefaultView.ToTable($distinct)

    $global:RecCntDist+=$dtf.Rows.Count
    
  

    $fileoutputComposite=$global:outputDir+"composite.csv"
    $dtf | Export-csv $fileoutputComposite -Append -NoTypeInformation
    
    }
    } 
    
  
   
}
}
cls
function Get-FileListing($remoteRootDir){
    $ens=Get-Childitem $remoteRootDir -Recurse -Filter *.pdf
    $fileTotalCount=($ens | Measure-Object).Count
    $fileCount=0
    foreach($e in $ens){
        $fileCount++
        
        $r=New-Object iTextSharp.text.pdf.PdfReader $e.FullName
        #Write-Host "reader pages:" $r.NumberOfPages

        for ($i = 1; $i -le $r.NumberOfPages; $i++) { 
            $outFile=$global:outputDir  + $e.BaseName +"-i"+$i.ToString()+".txt"
            #Get-PDFContent2 $e.FullName $i |Set-Content $outfile
            [string]$txt=Get-PDFContent2 $e.FullName $i 
            
            $pbAct="Parsing "+$e.basename +"("+'File ' `
                            + $fileCount + ' of ' + $fileTotalCount+")" `
                            + "Parsed records:" + $global:RecCnt `
                            + " / Unique encounters:" + $global:RecCntDist `
                            + " / record discards:"+$global:RecDiscard `
                            + " / Page discards:"+$global:PgDiscard; 
            $ipct=(($i/$r.NumberOfPages)*100)
             
            Write-Progress -Activity $pbAct -Status 'Progress..' -PercentComplete $ipct
            
            $txt | Set-Content $outfile
            Parse-Text  $txt
            
            Remove-Item $outfile           
        }
        
        $r.Close()
    }
}
cls
if(Test-Path $CompositeFile){
    Remove-Item ($CompositeFile)
}

$timeStart=(Get-Date)
Get-FileListing($wd)

# upload to webAPI
# $datafile="C:\Users\asabe\Desktop\clinic-pilot\composite.csv"
$data=import-csv $CompositeFile | ConvertTo-Json
# $uri="https://enpoint/"

$headers=New-Object "System.Collections.Generic.Dictionary[[string],[string]]"
$headers.Add("AuthToken","token here") 
# token needs to be addedd
$headers.Add("grant_type","password")



Invoke-RestMethod -Uri $uri -Body $data -Method Post -Headers $headers -ContentType 'application/json'



$timeStop=(Get-Date)
Write-Host "Completed Execution :"  (New-TimeSpan -Start $timeStart -End $timeStop).TotalSeconds.ToString()  "s"
