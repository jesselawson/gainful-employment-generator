$scriptWelcome = "Hello and welcome to The Automatic Regurgitator of Gainful Employment disclosure Templates (TARGET). We'll be starting shortly."

write-host $scriptWelcome

# Declare the variables we're going to populate from our data file
$opeId=""
$currentYear=""
$prevYear=""
$geHtmlName=""
$customProgramName=""
$cipCode=""
$awardLevel=""
$tutionAndFees=""
$booksAndSupplies=""
$roomAndBoard="0"
$roomAndBoardNotOffered="true"
$additionalFeeAndExpenses=""
$urlForProgramCost=""
$numberOfStudentsCompletedProgram=""
$numberOfStudentsCompletedProgramWithDebt=""
$isLessThanTenGraduatesReceivedLoan=
$title4MedianDebt=""
$privateMedianDebt=""
$financingPlanDebt=""
$normalTimeToCompleteProgram=""
$durationTypeInWeeksMonthsYears=""
$numberOfStudentsCompletedProgramInNormalTime=""
$additionalNotes=""
$dateCreated=""

# Load up the csv file that should be in the directory. We're looking for a csv file named 'target.csv'
$csv = import-csv -path .\target.csv

# Loop through the csv data object and start pulling out our variables
$i = 0
foreach($item in $csv) {

    #write-host "Working on #$i..."

    $randomSeed=Get-Random
    $opeId=$item.opeId
    $currentYear=$item.currentYear
    $prevYear=$item.prevYear
    $geHtmlName=$item.geHtmlName -replace "\.0", "0"
    $geHtmlName="$geHtmlName.html"
    #write-host $geHtmlName
    $cipCode=$item.cipCode
    write-host "$geHtmlName $cipCode"

    #"Connecting to ED.GOV for Official program name by CIP Code... "
    $string = Invoke-RestMethod -uri "http://ope.ed.gov/GainfulEmployment/BGP.aspx?action=searchprogramnamebycipcode&cipcode=$cipCode" -Method Get

    #"Cleaning results..."
    $string = $string -replace "success\|", ""

    # Remove the last two characters, which is a floating (erroneous) '|#'
    $string.Substring(0,$string.Length-2)

    $programName=$string

    $customProgramName=$item.customProgramName
    $awardLevel=$item.awardLevel
    $tutionAndFees=$item.tutionAndFees
    $booksAndSupplies=$item.booksAndSupplies
    $roomAndBoard="0"
    $roomAndBoardNotOffered="true";
    $numberOfStudentsCompletedProgram=$item.numberOfStudentsCompletedProgram
    $numberOfStudentsCompletedProgramWithDebt=$item.numberOfStudentsCompletedProgramWithDebt
    $isLessThanTenGraduatesReceivedLoan='no'
    if($numberOfStudentsCompletedProgramWithDebt > 9){
        $isLessThanTenGraduatesReceivedLoan='yes'
    }
    $title4MedianDebt=$item.title4MedianDebt
    $privateMedianDebt=$item.privateMedianDebt
    $financingPlanDebt=$item.financingPlanDebt
    $normalTimeToCompleteProgram=$item.normalTimeToCompleteProgram
    $durationTypeInWeeksMonthsYears=$item.durationTypeInWeeksMonthsYears
    $numberOfStudentsCompletedProgramInNormalTime=$item.numberOfStudentsCompletedProgramInNormalTime
    $dateCreated="1/26/2015" #Get-Date -format M/d/yyyy

    #"Connecting to ED.GOV for related programs by CIP code... "
    $string = Invoke-RestMethod -uri "http://ope.ed.gov/GainfulEmployment/BGP.aspx?action=getrelatedoccupations&cipCode=$cipCode" -Method Get

    #"Cleaning results..."
    $string = $string -replace "list\|", ""

    # Remove the first character, which is a floating (erroneous) '@'
    $string = $string.substring(1)

    # Split the results on the delimeters. OPE separates each job by '@', then from each of those strings it separates the SOC and Name by '###'
    $listofjobs = $string -split '@'

    # Build URLs of each of the related jobs
    #"Building related jobs list..."
    $relatedJobs = ""
    foreach($job in $listofjobs) {
      $parts = $job.split('###')
      $socCode = $parts[0]
      $jobName = $parts[3]
      $html = "dict[randomSeed].relatedOccupations.push({ oneTitle: ' "+$jobName+"', oneCode: '"+$socCode+"' });"
      $relatedJobs = $relatedJobs + $html
    }

    # Create a new HTML file in the gainful-employment directory
    $original_file = 'template.html'
    $destination_file =  "gainful-employment\$geHtmlName"
    (Get-Content $original_file) | Foreach-Object {
    $_ -replace '{RANDOM_SEED}', $randomSeed `
        -replace '{OPE_ID}', $opeId `
        -replace '{CURRENT_YEAR}', $currentYear `
        -replace '{PREV_YEAR}', $prevYear `
        -replace '{PROGRAM_NAME}', $programName `
        -replace '{CUSTOM_PROGRAM_NAME}', $customProgramName `
        -replace '{CIP_CODE}', $cipCode `
        -replace '{AWARD_LEVEL}', $awardLevel `
        -replace '{TUITION_AND_FEES}', $tutionAndFees `
        -replace '{BOOKS_AND_SUPPLIES}', $booksAndSupplies `
        -replace '{NUMBER_OF_STUDENTS_COMPLETED_PROGRAM}', $numberOfStudentsCompletedProgram `
        -replace '{NUMBER_OF_STUDENTS_COMPLETED_PROGRAM_WITH_DEBT}', $numberOfStudentsCompletedProgramWithDebt `
        -replace '{IS_LESS_THAN_TEN_GRADUATES_RECEIVED_LOAN}', $isLessThanTenGraduatesReceivedLoan `
        -replace '{TITLE_4_MEDIAN_DEBT}', $title4MedianDebt `
        -replace '{PRIVATE_MEDIAN_DEBT}', $privateMedianDebt `
        -replace '{FINANCING_PLAN_DEBT}', $financingPlanDebt `
        -replace '{NORMAL_TIME_TO_COMPLETE_PROGRAM}', $normalTimeToCompleteProgram `
        -replace '{NUMBER_OF_STUDENTS_COMPLETED_PROGRAM_IN_NORMAL_TIME}', $numberOfStudentsCompletedProgramInNormalTime `
        -replace '{DATE_CREATED}', $dateCreated `
        -replace '{RELATED_OCCUPATIONS_LIST}', $relatedJobs
    } | Set-Content $destination_file

    #write-host "Finished $customProgramName!"

    $i++
}

"That's all, folks!"
