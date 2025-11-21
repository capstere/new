Param(
    [string]$BatchId,
    [int]$JobId
)

# region Configuration
# Scheduled via Task Scheduler (default: no parameters) to process SharePoint jobs every ~2 hours.
# Manual run: optional -BatchId <Lot> or -JobId <Queue Item Id> to target a specific job.
$script:AutoMappscriptRoot = "\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\Skiftspecifika dokument\Skift 1\Mahdi\powerpoint\AutoMappscript"
$script:LogDirectory       = "$script:AutoMappscriptRoot\Log"

$script:SharePointSettings = @{
    Tenant              = "danaher.onmicrosoft.com"
    ClientId            = "Insert myself"
    Certificate         = "Insert myself"
    ProductionSiteUrl   = "https://danaher.sharepoint.com/sites/CEP-Sweden-Production-Management"
    DocumentSiteUrl     = "https://danaher.sharepoint.com/sites/CEP_CWS"
    ProductionListName  = "Cepheid | Production orders"
    JobListName         = "AutomaticMappscriptJobs"
    PendingJobStatuses  = @("New", "Retry")
}

$script:JobListConnection = $null
$script:ExitCode = 0

function Write-AutoLog {
    param(
        [string]$Message,
        [string]$FileName = "Log.csv"
    )

    $target = Join-Path -Path $script:LogDirectory -ChildPath $FileName

    if(Test-Path -Path $target){
        $Message | Add-Content -Path $target
    }
}
# endregion

# region Material lookup
function get-matvariables ($material) {

    [hashtable]$return = @{}

    switch ($material){
        
        #SARS
        {($_ -eq 'XPRSARS-COV2-10') -or ($_ -eq 'D39525')}{
            $ASSAY = "XPRSARS COV2"
            $highpos = $false
            $matrev = 'D39525'
            $mail = “700-9246 SARS COV2”
            $product = “XPRSARSCOV210”
            $sealassay = "SARS CoV2"
            $partnr = "700-9246"
            $ultra = $false
            $path = '\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\1. IPT - KOMMANDE TESTER\Mappscript\XPRSARS-COV2-10'
            $na = "N/A"
        }
        {($_ -eq 'XP3SARS-COV2-10') -or ($_ -eq 'D48538')}{
            $ASSAY = "XP3 SARS COV2"
            $highpos = $false
            $matrev = 'D48538'
            $mail = “700-7425 XP3SARSCOV210”
            $product = “XP3SARSCOV210”
            $sealassay = "Xpress CoV-2 plus"
            $partnr = "700-7425"
            $ultra = $false
            $path = '\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\1. IPT - KOMMANDE TESTER\Mappscript\XP3 SARS COV2'
            $na = "N/A"
        }
        {($_ -eq 'XPCOV2/FLU/RSV-10 deaad') -or ($_ -eq 'D41929')}{
            $ASSAY = "XPCOV2 FLU RSV"
            $highpos = $false
            $matrev = 'D41929'
            $mail = “700-6990 XPCOV2/FLU/RSV-10”
            $product = “XPCOV2FLURSV10”
            $sealassay = "SARS CoV2 Flu/RSV"
            $partnr = "700-6990"
            $ultra = $false
            $path = '\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\1. IPT - KOMMANDE TESTER\Mappscript\XPRSARS-COV2 FLU RSV'
            $na = "N/A"
        }
        {($_ -eq 'XP3COV2/FLU/RSV-10') -or ($_ -eq 'D47377')}{
            $ASSAY = "XP3 COV2 FLU RSV"
            $highpos = $false
            $matrev = 'D47377'
            $mail = “700-7493 XP3COV2FLURSV10”
            $product = “XP3COV2FLURSV10”
            $sealassay = "Xpress CoV-2/Flu/RSV plus"
            $partnr = "700-7493"
            $ultra = $false
            $path = '\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\1. IPT - KOMMANDE TESTER\Mappscript\XP3 SARS-COV2 FLU RSV plus'
            #$path = '\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\Skiftspecifika dokument\Skift 1\Mahdi\powerpoint\AutoMappscript\Fluvid+'
            $na = "N/A"
        }
        #MTB
        {($_ -eq 'GXMTB/RIF-ULTRA-50') -or ($_ -eq 'D25862')}{
            $ASSAY = “GXMTB RIF ULTRA”
            $ultra = $true
            $highpos = $true
            $ASSAYMTB = 1
            $matrev = 'D25862'
            $mail =  “700-5702 MTB Ultra”
            $product = “GXMTBRIFULTRA50”
            $ASSAYFAM = "MTB"
            $sealassay = "MTB"
            $partnr = "700-5702"
            $path = '\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\1. IPT - KOMMANDE TESTER\Mappscript\GXMTB RIF-ULTRA'
            $na = "NA"
        }
        {($_ -eq 'GXMTB/RIF-ULTRA-10') -or ($_ -eq 'D25862')}{
            $ASSAY = “GXMTB RIF ULTRA”
            $ultra = $true
            $highpos = $true
            $ASSAYMTB = 2
            $matrev = 'D25862'
            $mail = “700-5702 MTB Ultra”
            $product = “GXMTBRIFULTRA10”
            $ASSAYFAM = "MTB"
            $sealassay = "MTB"
            $partnr = "700-5702"
            $path = '\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\1. IPT - KOMMANDE TESTER\Mappscript\GXMTB RIF-ULTRA'
            $na = "NA"
        }
        {($_ -eq 'GXMTBRIF-ULT-CN-10') -or ($_ -eq 'D25862')}{
            $ASSAY = “GXMTB RIF ULTRA CN”
            $ultra = $true
            $highpos = $true
            $ASSAYMTB = 2
            $matrev = 'D25862'
            #$mail = “700-5702 MTB Ultra”
            $product = “GXMTBRIF-ULT-CN-10”
            $ASSAYFAM = "MTB"
            $sealassay = "MTB"
            $partnr = "700-8749"
            $path = '\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\1. IPT - KOMMANDE TESTER\Mappscript\GXMTB RIF-ULTRA'
            $na = "NA"
        }
        {($_ -eq 'GXMTBRIF-ULT-CN-50') -or ($_ -eq 'D25862')}{
            $ASSAY = “GXMTB RIF ULTRA CN”
            $ultra = $true
            $highpos = $true
            $ASSAYMTB = 2
            $matrev = 'D25862'
            #$mail = “700-5702 MTB Ultra”
            $product = “GXMTBRIF-ULT-CN-50”
            $ASSAYFAM = "MTB"
            $sealassay = "MTB"
            $partnr = "700-8749"
            $path = '\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\1. IPT - KOMMANDE TESTER\Mappscript\GXMTB RIF-ULTRA'
            $na = "NA"
        }
        {($_ -eq "GXMTB/XDR-10 död") -or ($_ -eq 'D37339')}{
            $ASSAY = “MTB XDR”
            $highpos = $false
            $ASSAYMTB = 5
            $matrev = 'D37339'
            $product = “GXMTBXDR10”
            $ASSAYFAM = "MTB"
            $sealassay = "MTB"
            $partnr = "700-6737"
            $doublesampling = $true
            $ultra = $true
            $path = '\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\1. IPT - KOMMANDE TESTER\Mappscript\MTB XDR'
            $na = "NA"
        }
        {($_ -eq 'CGXMTB/RIF-50') -or ($_ -eq 'D31503')}{
            $ASSAY = “MTB RIF”
            $highpos = $false
            $ASSAYMTB = 3
            $matrev = 'D31503'
            $mail = “700-4006 MTB-RIF”
            $product = “CGXMTBRIF50”
            $ASSAYFAM = "MTB"
            $sealassay = "MTB"
            $partnr = "700-4006"
            $ultra = $false
            $path = '\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\1. IPT - KOMMANDE TESTER\Mappscript\MTB RIF'
            $na = "N/A"
        }
        {($_ -eq 'GXMTB/RIF-50') -or ($_ -eq 'D31503')}{
            $ASSAY = “MTB RIF”
            $highpos = $false
            $ASSAYMTB = 3
            $matrev = 'D31503'
            $mail = “700-4006 MTB-RIF”
            $product = “GXMTBRIF50”
            $ASSAYFAM = "MTB"
            $sealassay = "MTB"
            $partnr = "700-4006"
            $ultra = $false
            $path = '\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\1. IPT - KOMMANDE TESTER\Mappscript\MTB RIF'
            $na = "N/A"
        }
        {($_ -eq "700-6082 ASSY,10 CART POUCH,MTB/RIF,HB, INDIA") -or ($_ -eq 'D31503')}{
            $ASSAY = “MTB MII”
            $highpos = $false
            $ASSAYMTB = 4
            $matrev = 'D31503'
            $mail = “700-4006 700-6082 MTB MII”
            $product = “700-6082”
            $ASSAYFAM = "MTB"
            $sealassay = "MTB"
            $partnr = "700-4006"
            $ultra = $false
            $path = '\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\1. IPT - KOMMANDE TESTER\Mappscript\MTB RIF'
            $na = "N/A"
        }
        {($_ -eq "GXMTB/RIF-10") -or ($_ -eq 'D31503')}{
            $ASSAY = “MTB RIF”
            $highpos = $false
            $ASSAYMTB = 5
            $matrev = 'D31503'
            $mail = “700-4006 MTB-RIF”
            $product = “GXMTBRIF10”
            $ASSAYFAM = "MTB"
            $sealassay = "MTB"
            $partnr = "700-4006"
            $ultra = $false
            $path = '\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\1. IPT - KOMMANDE TESTER\Mappscript\MTB RIF'
            $na = "N/A"
        }
        {($_ -eq "CGXMTB/RIF-10") -or ($_ -eq 'D31503')}{
            $ASSAY = “MTB RIF”
            $highpos = $false
            $ASSAYMTB = 5
            $matrev = 'D31503'
            $mail = “700-4006 MTB-RIF”
            $product = “CGXMTBRIF10”
            $ASSAYFAM = "MTB"
            $sealassay = "MTB"
            $partnr = "700-4006"
            $ultra = $false
            $path = '\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\1. IPT - KOMMANDE TESTER\Mappscript\MTB RIF'
            $na = "N/A"
        }
        {($_ -eq "GXMTB/RIF-CN-50") -or ($_ -eq 'D31503')}{
            $ASSAY = “MTB RIF CN”
            $highpos = $false
            $ASSAYMTB = 5
            $matrev = 'D31503'
            $mail = “700-4006 MTB-RIF”
            $product = “GXMTBRIFCN50”
            $ASSAYFAM = "MTB"
            $sealassay = "MTB"
            $partnr = "700-9074"
            $ultra = $false
            $path = '\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\1. IPT - KOMMANDE TESTER\Mappscript\MTB RIF'
            $na = "N/A"
        }
        {($_ -eq "GXMTB/RIF-CN-10") -or ($_ -eq 'D31503')}{
            $ASSAY = “MTB RIF CN”
            $highpos = $false
            $ASSAYMTB = 5
            $matrev = 'D31503'
            $mail = “700-4006 MTB-RIF”
            $product = “GXMTBRIFCN10”
            $ASSAYFAM = "MTB"
            $sealassay = "MTB"
            $partnr = "700-9074"
            $ultra = $false
            $path = '\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\1. IPT - KOMMANDE TESTER\Mappscript\MTB RIF'
            $na = "N/A"
        }
        {($_ -eq "700-7773") -or ($_ -eq 'D31503')}{
            $ASSAY = “MTB RIF”
            $highpos = $false
            $ASSAYMTB = 5
            $matrev = 'D31503'
            $mail = “700-4006 MTB-RIF”
            $product = “700-7773”
            $ASSAYFAM = "MTB"
            $sealassay = "MTB"
            $partnr = "700-4006"
            $ultra = $false
            $path = '\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\1. IPT - KOMMANDE TESTER\Mappscript\MTB RIF'
            $na = "N/A"
        }
        #CT/NG
        {($_ -eq "GXCT/NG-10") -or ($_ -eq 'D16904')}{
            $ASSAY = “CT NG”
            $highpos = $false
            $ASSAYCTNG = 1
            $matrev = 'D16904'
            $mail = “700-5097 CT/NG”
            $product = “GXCTNG10”
            $ASSAYFAM = "CTNG"
            $sealassay = "CTNG and Xpress CT/NG"
            $partnr = "700-5097"
            $ultra = $false
            $path = '\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\1. IPT - KOMMANDE TESTER\Mappscript\CT NG'
            $na = "N/A"
        }
        {($_ -eq "GXCT/NG-CE-10") -or ($_ -eq 'D16904')}{
            $ASSAY = “CT NG”
            $highpos = $false
            $ASSAYCTNG = 2
            $matrev = 'D16904'
            $mail = “700-5097 CT/NG”
            $product = “GXCTNGCE10” 
            $ASSAYFAM = "CTNG"
            $sealassay = "CTNG and Xpress CT/NG"
            $partnr = "700-5097"
            $ultra = $false
            $path = '\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\1. IPT - KOMMANDE TESTER\Mappscript\CT NG'
            $na = "N/A"
        }
        {($_ -eq "GXCT/NGX-CE-10") -or ($_ -eq 'D16904')}{
            $ASSAY = “CT NGX”
            $highpos = $false
            $ASSAYCTNG = 3
            $matrev = 'D16904'
            $mail = “700-5097 CT/NG”
            $product = “GXCTNGXCE10”
            $ASSAYFAM = "CTNG"
            $sealassay = "CTNG and Xpress CT/NG"
            $partnr = "700-5097"
            $ultra = $false
            $path = '\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\1. IPT - KOMMANDE TESTER\Mappscript\CT NG'
            $na = "N/A"
        }
        {($_ -eq "GXCT/NG-CN-10") -or ($_ -eq 'D16904')}{
            $ASSAY = “CTNG CN”
            $highpos = $false
            $ASSAYCTNG = 4
            $matrev = 'D16904'
            $mail = “700-9076 CT/NG”
            $product = “GXCTNGCN10”
            $ASSAYFAM = "CTNG"
            $sealassay = "CTNG and Xpress CT/NG"
            $partnr = "700-9076"
            $ultra = $false
            $path = '\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\1. IPT - KOMMANDE TESTER\Mappscript\CT NG'
            $na = "N/A"
        }
        {($_ -eq "GXCT/NG-CE-120") -or ($_ -eq 'D16904')}{
            $ASSAY = “CTNG”
            $highpos = $false
            $ASSAYCTNG = 4
            $matrev = 'D16904'
            $mail = “700-5097 CT/NG”
            $product = “GXCTNGCE120”
            $ASSAYFAM = "CTNG"
            $sealassay = "CTNG and Xpress CT/NG"
            $partnr = "700-5097"
            $ultra = $false
            $path = '\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\1. IPT - KOMMANDE TESTER\Mappscript\CT NG'
            $na = "N/A"
        }
        #MRSA NXG
        {($_ -eq "GXMRSA-NXG-CE-10") -or ($_ -eq 'D23916')}{
            $ASSAY = "MRSA NXG"
            $highpos = $true
            $matrev = 'D23916'
            $mail = “700-4881 MRSA NxG”
            $product = “GXMRSANXGCE10”
            $sealassay = "MRSA"
            $partnr = "700-4881"
            $ultra = $false
            $path = '\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\1. IPT - KOMMANDE TESTER\Mappscript\GXMRSA NXG'
            $na = "NA"

        }
        #MRSA/SA
        {($_ -eq "GXSACOMP-CN-10") -or ($_ -eq 'D36872')}{
            $ASSAY = "SA COMP CN"
            $matrev = 'D36872'
            #$mail = “700-4881 MRSA NxG”
            $product = “GXSACOMP-CN-10”
            $sealassay = "SA Comp"
            $partnr = "700-9132"
            $ultra = $false
            $path = '\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\1. IPT - KOMMANDE TESTER\Mappscript\MRSA SA'
            $na = "NA"

        }
        {($_ -eq "GXSACOMP-CE-10") -or ($_ -eq 'D36872')}{
            $ASSAY = "SA COMP"
            $matrev = 'D36872'
            #$mail = “700-4881 MRSA NxG”
            $product = “GXSACOMP-CE-10”
            $sealassay = "SA Comp"
            $partnr = "700-4023"
            $ultra = $false
            $path = '\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\1. IPT - KOMMANDE TESTER\Mappscript\MRSA SA'
            $na = "NA"

        }
        {($_ -eq "GXMRSA/SA-SSTI-CE") -or ($_ -eq 'D36872')}{
            $ASSAY = "MRSA SA SSTI"
            $matrev = 'D36872'
            #$mail = “700-4881 MRSA NxG”
            $product = “GXMRSA/SA-SSTI-CE”
            $sealassay = "MRSA"
            $partnr = "700-3874"
            $ultra = $false
            $path = '\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\1. IPT - KOMMANDE TESTER\Mappscript\MRSA SA'
            $na = "NA"

        }
        {($_ -eq "GXMRSA/SA-BC-CE-10") -or ($_ -eq 'D36872')}{
            $ASSAY = "MRSA SA BC"
            $matrev = 'D36872'
            #$mail = “700-4881 MRSA NxG”
            $product = “GXMRSA/SA-BC-CE-10”
            $sealassay = "MRSA"
            $partnr = "700-3875"
            $ultra = $false
            $path = '\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\1. IPT - KOMMANDE TESTER\Mappscript\MRSA SA'
            $na = "NA"

        }
        #C.DIFF
        {($_ -eq "GXCDIFFBT-CE-10") -or ($_ -eq 'D37468')}{
            $ASSAY = “C. DIFF BT”
            $highpos = $false
            $matrev = 'D37468'
            $mail = “700-5178 C.Diff BT”
            $product = “GXCDIFFBTCE10”
            $sealassay = "C. Diff"
            $partnr = "700-5178"
            $ultra = $false
            $path = '\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\1. IPT - KOMMANDE TESTER\Mappscript\C. Difficile'
            $na = "NA"
        }
        {($_ -eq "GXCDIFFICILE-CE-10") -or ($_ -eq 'D37468')}{
            $ASSAY = “C. DIFF”
            $highpos = $false
            $matrev = 'D37468'
            $mail = “700-5102 C. difficile”
            $product = “GXCDIFFICILECE10”
            $sealassay = "C. Diff"
            $partnr = "700-5102"
            $ultra = $false
            $path = '\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\1. IPT - KOMMANDE TESTER\Mappscript\C. Difficile'
            $na = "NA"
        }
        {($_ -eq "GXCDIFFICILE-CN-10") -or ($_ -eq 'D37468')}{
            $ASSAY = “C. DIFF”
            $highpos = $false
            $matrev = 'D37468'
            $mail = “700-5102 C. difficile”
            $product = “GXCDIFFICILECN10”
            $sealassay = "C. Diff"
            $partnr = "700-5102"
            $ultra = $false
            $path = '\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\1. IPT - KOMMANDE TESTER\Mappscript\C. Difficile'
            $na = "NA"
        }
        #HIV
        {($_ -eq "GXHIV-VL-CE-10 FUNKAR EJ") -or ($_ -eq 'D51110')}{
            $ASSAY = “HIV VL”
            $highpos = $true
            $ASSAYHIV = 1
            $matrev = 'D51110'
            $mail = “700-4370 HIV VL”
            $product = “GXHIV-VL-CE-10”
            $ASSAYFAM = "HIV"
            $sealassay = "HIV"
            $partnr = "700-4370"
            $ultra = $false
            $path = '\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\1. IPT - KOMMANDE TESTER\Mappscript\HIV VL'
            $na = "N/A"
        }
        {($_ -eq "GXHIV-VL-CN-10 FUNKAR EJ") -or ($_ -eq 'D51110')}{
            $ASSAY = “HIV VL CN”
            $highpos = $true
            $ASSAYHIV = 2
            $matrev = 'D51110'
            $mail = “700-4370 HIV VL”
            $product = “GXHIV-VL-CN-10”
            $ASSAYFAM = "HIV"
            $sealassay = "HIV"
            $partnr = "700-4370"
            $ultra = $false
            $path = '\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\1. IPT - KOMMANDE TESTER\Mappscript\HIV VL'
            $na = "N/A"
        }
        {($_ -eq "GXHIV-VL-IN-10 FUNKAR EJ") -or ($_ -eq 'D51110')}{
            $ASSAY = “HIV VL IN”
            $highpos = $true
            $ASSAYHIV = 3
            $matrev = 'D51110'
            $mail = “700-4370 HIV VL”
            $product = “GXHIV-VL-IN-10”
            $ASSAYFAM = "HIV"
            $sealassay = "HIV"
            $partnr = "700-4370"
            $ultra = $false
            $path = '\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\1. IPT - KOMMANDE TESTER\Mappscript\HIV VL'
            $na = "N/A"
        }
        {($_ -eq "GXHIV-VL-XC-CE-10") -or ($_ -eq 'D51111')}{
            $ASSAY = “HIV VL XC”
            $highpos = $true
            $ASSAYHIV = 4
            $matrev = 'D51111'
            $mail = “700-6646 HIV VL XC”
            $product = “GXHIV-VL-XC-CE-10”
            $ASSAYFAM = "HIV"
            $sealassay = "HIV"
            $partnr = "700-6646"
            $ultra = $false
            $path = '\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\1. IPT - KOMMANDE TESTER\Mappscript\HIV VL XC'
            $na = "NA"
        }
        {($_ -eq "GXHIV-VL-XC-CA-10") -or ($_ -eq 'D79377')}{
            $ASSAY = “HIV VL XC CA”
            $highpos = $true
            $ASSAYHIV = 4
            $matrev = 'D79377'
            $mail = “700-6646 HIV VL XC”
            $product = “GXHIV-VL-XC-CA-10”
            $ASSAYFAM = "HIV"
            $sealassay = "HIV"
            $partnr = "700-6646"
            $ultra = $false
            $path = '\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\1. IPT - KOMMANDE TESTER\Mappscript\HIV VL XC CA'
            $na = "NA"
        }
        {($_ -eq "GXHIV-QA-CE-10") -or ($_ -eq "D61353")}{
            $ASSAY = “HIV QA”
            $highpos = $true
            $ASSAYHIV = 5
            $matrev = "D61353"
            $mail = “700-4369 HIV QA”
            $product = “GXHIV-QA-CE-10”
            $ASSAYFAM = "HIV"
            $sealassay = "HIV"
            $partnr = "700-4369"
            $ultra = $false
            $path = '\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\1. IPT - KOMMANDE TESTER\Mappscript\HIV QA'
            $na = "N/A"
        }
        {($_ -eq "GXHIV-QA-XC-CE-10") -or ($_ -eq 'D36985')}{
            $ASSAY = “HIV QA XC”
            $highpos = $false
            $ASSAYHIV = 6
            $matrev = 'D36985'
            $mail = “700-6793 HIV QA XC”
            $product = “700-6793/ HIV-1 QUAL XC, CE-IVD”
            $ASSAYFAM = "HIV"
            $sealassay = "HIV"
            $partnr = "700-6793"
            $ultra = $false
            $path = '\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\1. IPT - KOMMANDE TESTER\Mappscript\HIV QA XC'
            $na = "NA"
        }
        #OTHER QUANT ASSAYS
        {($_ -eq "GXHBV-VL-CE-10") -or ($_ -eq 'D51114')}{
            $ASSAY = “HBV VL”
            $highpos = $true
            $matrev = 'D51114'
            $mail = “700-5720 HBV VL” 
            $product = “GXHBV-VL-CE-10”
            $sealassay = "HBV"
            $partnr = "700-5720"
            $ultra = $false
            $path = '\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\1. IPT - KOMMANDE TESTER\Mappscript\HBV VL'
            $na = "N/A"
        }
        {($_ -eq "GXHBV-VL-CN-10") -or ($_ -eq 'D51114')}{
            $ASSAY = “HBV VL”
            $highpos = $true
            $matrev = 'D51114'
            $mail = “700-5720 HBV VL” 
            $product = “GXHBV-VL-CN-10”
            $sealassay = "HBV"
            $partnr = "700-9093"
            $ultra = $false
            $path = '\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\1. IPT - KOMMANDE TESTER\Mappscript\HBV VL'
            $na = "N/A"
        }
        {($_ -eq "GXHPV-CE-10") -or ($_ -eq 'D16546')}{
            $ASSAY = “HPV”
            $highpos = $false
            $matrev = 'D16546'
            $mail = “700-4148 HPV”
            $product = “GXHPV-CE-10”
            $sealassay = "HPV"
            $partnr = "700-4148"
            $ultra = $false
            $hpv = $true
            $path = '\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\1. IPT - KOMMANDE TESTER\Mappscript\HPV'
            $na = "N/A"
        }
        {($_ -eq "CGXHPV-CE-10") -or ($_ -eq 'D16546')}{
            $ASSAY = “HPV”
            $highpos = $false
            $matrev = 'D16546'
            $mail = “700-4148 HPV”
            $product = “CGXHPV-CE-10”
            $sealassay = "HPV"
            $partnr = "700-4148"
            $ultra = $false
            $hpv = $true
            $path = '\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\1. IPT - KOMMANDE TESTER\Mappscript\HPV'
            $na = "N/A"
        }
        {($_ -eq "GXHPV2-CE-10") -or ($_ -eq 'D16546')}{
            $ASSAY = “HPV v2”
            $highpos = $false
            $matrev = 'D16546'
            $mail = “700-8500 HPV V2”
            $product = “GXHPV2-CE-10”
            $sealassay = "HPV"
            $partnr = "700-8500"
            $ultra = $false
            $hpv = $true
            $path = '\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\1. IPT - KOMMANDE TESTER\Mappscript\HPV'
            $na = "N/A"
        }
        {($_ -eq "GXHCV-VL-CE-10") -or ($_ -eq 'D51112')}{
            $ASSAY = “HCV VL”
            $highpos = $true
            $matrev = 'D51112'
            $mail = “700-4581 HCV VL”
            $product = “GXHCV-VL-CE-10”
            $sealassay = "HCV"
            $partnr = "700-4581"
            $ultra = $false
            $hpv = $false
            $path = '\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\1. IPT - KOMMANDE TESTER\Mappscript\HCV VL'
            $na = "N/A"
        }
        {($_ -eq "GXHCV-VL-CN-10") -or ($_ -eq 'D51112')}{
            $ASSAY = “HCV VL”
            $highpos = $true
            $matrev = 'D51112'
            $mail = “700-4581 HCV VL”
            $product = “GXHCV-VL-CN-10”
            $sealassay = "HCV"
            $partnr = "700-9113"
            $ultra = $false
            $hpv = $false
            $path = '\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\1. IPT - KOMMANDE TESTER\Mappscript\HCV VL'
            $na = "N/A"
        }
        {($_ -eq "GXHCV-VL-IN-10") -or ($_ -eq 'D51112')}{
            $ASSAY = “HCV VL”
            $highpos = $true
            $matrev = 'D51112'
            $mail = “700-4581 HCV VL”
            $product = “GXHCV-VL-IN-10”
            $sealassay = "HCV"
            $partnr = "700-4581"
            $ultra = $false
            $hpv = $false
            $path = '\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\1. IPT - KOMMANDE TESTER\Mappscript\HCV VL'
            $na = "N/A"
        }
        {($_ -eq "GXHCV-FS-CE-10") -or ($_ -eq 'D51113')}{
            $ASSAY = “HCV FS”
            $highpos = $true
            $matrev = 'D51113'
            $product = “GXHCV-FS-CE-10”
            $sealassay = "HCV"
            $partnr = "700-5634"
            $ultra = $false
            $hpv = $false
            $path = '\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\1. IPT - KOMMANDE TESTER\Mappscript\HCV FS'
            $na = "N/A"
        }
        #CARBA
        {($_ -eq "GXCARBARP-CE-10") -or ($_ -eq 'D18272')}{
            $ASSAY = “CARBA”
            $carba = $true
            $matrev = 'D18272'
            $mail = “700-5014 CARBA-RP”
            $product = “GXCARBARPCE10”
            $sealassay = "Carba-R"
            $partnr = "700-5014"
            $ultra = $false
            $path = '\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\1. IPT - KOMMANDE TESTER\Mappscript\CARBA-R'
            $na = "N/A"        
        }
        {($_ -eq "GXCARBARP-CE-120") -or ($_ -eq 'D18272')}{
            $ASSAY = “CARBA”
            $carba = $true
            $matrev = 'D18272'
            $mail = “700-5014 CARBA-RP”
            $product = “GXCARBARPCE120”
            $sealassay = "Carba-R"
            $partnr = "700-5014"
            $ultra = $false
            $path = '\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\1. IPT - KOMMANDE TESTER\Mappscript\CARBA-R'
            $na = "N/A"        
        
        }
        {($_ -eq "GXCARBAR-CE-10") -or ($_ -eq 'D18272')}{
            $ASSAY = “CARBA”
            $carba = $true
            $matrev = 'D18272'
            $mail = “700-5014 CARBA-RP”
            $product = “GXCARBARCE10”
            $sealassay = "Carba-R"
            $partnr = "700-4186"
            $ultra = $false
            $path = '\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\1. IPT - KOMMANDE TESTER\Mappscript\CARBA-R'
            $na = "N/A"        
        
        }
        #VANAB
        {($_ -eq "GXVANA/B-CE-10") -or ($_ -eq 'D55782')}{
            $ASSAY = “VAN AB”
            #$carba = $true
            $matrev = 'D55782'
            #$mail = “700-5014 CARBA-RP”
            $product = “GXVANA/B-CE-10”
            $sealassay = "Van A"
            $partnr = "700-3882"
            $ultra = $false
            $path = '\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\1. IPT - KOMMANDE TESTER\Mappscript\VANA-VANB'
            $na = "N/A"        
        
        }
        #STREP A
        {($_ -eq "XPRSTREPA-CE-10") -or ($_ -eq 'D27089')}{
            $ASSAY = “STREP A”
            $highpos = $true
            $matrev = 'D27089'
            $product = “XPRSTREPA-CE-10”
            $sealassay = "Strep A"
            $partnr = "700-5360"
            $ultra = $false
            $path = '\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\1. IPT - KOMMANDE TESTER\Mappscript\STREP A'
            $na = "N/A"        
        
        }
        {($_ -eq "XPRSTREPA-10") -or ($_ -eq 'D27089')}{
            $ASSAY = “STREP A”
            $highpos = $true
            $matrev = 'D27089'
            $product = “XPRSTREPA-10”
            $sealassay = "Strep A"
            $partnr = "700-5420"
            $ultra = $false
            $path = '\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\1. IPT - KOMMANDE TESTER\Mappscript\STREP A'
            $na = "N/A"        
        
        }
        #FLURSV/FLU
        {($_ -eq "XPRSFLU/RSV-CE-10") -or ($_ -eq 'D26120')}{
            $ASSAY = “FLURSV”
            $highpos = $true
            $matrev = 'D26120'
            $product = “XPRSFLURSVCE10”
            $sealassay = "Flu"
            $partnr = "700-5164"
            $ultra = $false
            $path = '\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\1. IPT - KOMMANDE TESTER\Mappscript\FLU'
            $na = "NA"        
        
        }
        {($_ -eq "XPRSFLU-10") -or ($_ -eq 'D26120')}{
            $ASSAY = “FLU”
            $highpos = $true
            $matrev = 'D26120'
            $product = “XPRSFLU10”
            $sealassay = "Flu"
            $partnr = "700-5260"
            $ultra = $false
            $path = '\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\1. IPT - KOMMANDE TESTER\Mappscript\FLU'
            $na = "NA"        
        
        }
        {($_ -eq "XPRSFLU/RSV-10") -or ($_ -eq 'D26120')}{
            $ASSAY = “FLURSV”
            $highpos = $true
            $matrev = 'D26120'
            $product = “XPRSFLURSV10”
            $sealassay = "Flu"
            $partnr = "700-5252"
            $ultra = $false
            $path = '\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\1. IPT - KOMMANDE TESTER\Mappscript\FLU'
            $na = "NA"        
        
        }
        #NORO
        {($_ -eq "GXNOV-CE-10") -or ($_ -eq 'D17716')}{
            $ASSAY = “NORO”
            $highpos = $true
            $matrev = 'D17716'
            $product = “GXNOV-CE-10”
            $sealassay = "Norovirus"
            $partnr = "700-4147"
            $ultra = $false
            $path = '\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\1. IPT - KOMMANDE TESTER\Mappscript\Xpert Norovirus'
            $na = "N/A"        
        
        }
        {($_ -eq "GXNOV-10") -or ($_ -eq 'D17716')}{
            $ASSAY = “NORO”
            $highpos = $true
            $matrev = 'D17716'
            $product = “GXNOV-10”
            $sealassay = "Norovirus"
            $partnr = "700-4147"
            $ultra = $false
            $path = '\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\1. IPT - KOMMANDE TESTER\Mappscript\Xpert Norovirus'
            $na = "N/A"        
        
        }
        #EBOLA
        {($_ -eq "GXEBOLA-CE-10") -or ($_ -eq 'D21938')}{
            $ASSAY = “EBOLA”
            $highpos = $true
            $matrev = 'D21938'
            $product = “GXEBOLA-CE-10”
            $sealassay = "Ebola"
            $partnr = "700-4731"
            $ultra = $false
            $path = '\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\1. IPT - KOMMANDE TESTER\Mappscript\Ebola'
            $na = "N/A"        
        
        }
        {($_ -eq "GXEBOLA-CE-50") -or ($_ -eq 'D21938')}{
            $ASSAY = “EBOLA”
            $highpos = $true
            $matrev = 'D21938'
            $product = “GXEBOLA-CE-50”
            $sealassay = "Ebola"
            $partnr = "700-4731"
            $ultra = $false
            $path = '\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\1. IPT - KOMMANDE TESTER\Mappscript\Ebola'
            $na = "N/A"        
        
        }
        {($_ -eq "GXEBOLA-10") -or ($_ -eq 'D21938')}{
            $ASSAY = “EBOLA”
            $highpos = $true
            $matrev = 'D21938'
            $product = “GXEBOLA-10”
            $sealassay = "Ebola"
            $partnr = "700-4671"
            $ultra = $false
            $path = '\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\1. IPT - KOMMANDE TESTER\Mappscript\Ebola'
            $na = "N/A"        
        
        }
        #JAPAN ASSAYS
        {($_ -eq "GXMTB/RIF-JP-10") -or ($_ -eq 'D66612')}{
            $ASSAY = “MTB RIF JP”
            $highpos = $true
            $matrev = 'D66612'
            $product = “GXMTB/RIF-JP-10”
            $sealassay = "MTB"
            $partnr = "700-4676"
            $ultra = $false
            $path = '\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\1. IPT - KOMMANDE TESTER\Mappscript\MTB JP'
            $na = "N/A"
            $japan = "MTB"
        }
        {($_ -eq "GXCDIFF-JP-10") -or ($_ -eq 'D66613')}{
            $ASSAY = “C.DIFF JP”
            $highpos = $true
            $matrev = 'D66613'
            $product = “GXCDIFF-JP-10”
            $sealassay = "C. Diff"
            $partnr = "700-5102"
            $ultra = $false
            $path = '\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\1. IPT - KOMMANDE TESTER\Mappscript\C.Diff JP'
            $na = "N/A"
            $japan = "CDIFF"
        }
        {($_ -eq "GXSACOMP-JP-10") -or ($_ -eq 'D66614')}{
            $ASSAY = “SA COMP JP”
            $highpos = $true
            $matrev = 'D66614'
            $product = “GXSACOMP-JP-10”
            $sealassay = "SA Comp"
            $partnr = "700-4023"
            $ultra = $false
            $path = '\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\1. IPT - KOMMANDE TESTER\Mappscript\MRSA JP'
            $na = "N/A"
            $japan = "MRSA"
        }
        {($_ -eq "GXCT/NG-JP-10") -or ($_ -eq 'D66615')}{
            $ASSAY = “CTNG JP”
            $highpos = $true
            $matrev = 'D66615'
            $product = “GXCT/NG-JP-10”
            $sealassay = "CTNG and Xpress CT/NG"
            $partnr = "700-5097"
            $ultra = $false
            $path = '\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\1. IPT - KOMMANDE TESTER\Mappscript\CTNG JP'
            $na = "N/A"
            $japan = "CTNG"
        }



        #ej matchar
        default{
            $ASSAY = 0
            $matrev = 'N/A'
            $path = '\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\Skiftspecifika dokument\Skift 1\Mahdi\powerpoint\AutoMappscript\NotYet'
        }

    }

    $return.assay = $ASSAY
    $return.assaymtb = $ASSAYMTB
    $return.assayctng = $ASSAYCTNG
    $return.assayhiv = $ASSAYHIV
    $return.matrev = $matrev
    $return.highpos = $highpos
    $return.ultra = $ultra
    $return.doublesampling = $doublesampling
    $return.hpv = $hpv
    $return.carba = $carba
    $return.mail = $mail
    $return.product = $product
    $return.assayfam = $ASSAYFAM
    $return.sealassay = $sealassay
    $return.partnr = $partnr
    $return.path = $path
    $return.na = $na
    $return.japan = $japan



    #$ASSAY = $returnmat.assay
    #$ASSAYMTB = $returnmat.assaymtb
    #$ASSAYHIV = $returnmat.assayhiv
    #$ASSAYCTNG = $returnmat.assayctng


    return $return


}

# endregion


#////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


function create-folder ($returnmat, $robalnr, $ordernr, $batchnr, $lsp, $samplereagent, $prodtime, $version, $robalitem) {

    #$excel = New-Object -ComObject Excel.Application
    #$excel.ScreenUpdating = $false
    #$excel.Visible = $False
    #$excel.DisplayStatusBar = $false
    #$excel.EnableEvents = $false


    #$book = $excel.Workbooks.Open('\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\Skiftspecifika dokument\Skift 1\Mahdi\mappscript\Mappscript\Mailtemplat.xlsm', $false, $false)

    #create-folder $returnmat $robalitem.robalnr $robalitem.ordernr $robalitem.batchnr $robalitem.lsp $robalitem.samplereagent $robalitem.prodtime $robalitem.orderamount $username


    #TODO
    #EPPLUS till kallibreringsfilen

    $equipment = Import-Clixml -LiteralPath "\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\Skiftspecifika dokument\Skift 1\Mahdi\powerpoint\AutoMappscript\equipment.xml"

    #/////////////////////////////////////////////////////////////////
    
    $assay = $returnmat.assay
    $document = $returnmat.matrev
    $highpos = $returnmat.highpos
    $ultra = $returnmat.ultra
    $doublesampling = $returnmat.doublesampling
    $carba = $returnmat.carba
    $hpv = $returnmat.hpv
    $mail = $returnmat.mail
    $product = $returnmat.product
    $returnmat.assaymtb, $returnmat.assayctng, $returnmat.assayhiv | ForEach-Object{if($_ -ne $null){$ASSAYNUM = $_}}
    $ASSAYFAM = $returnmat.ASSAYFAM
    $sealassay = $returnmat.sealassay
    $partnr = $returnmat.partnr

    $robal = $robalitem.robalnr
    $POnr = $robalitem.ordernr
    $batchnr = $robalitem.batchnr
    $LSP = $robalitem.lsp
    $samplereagent = $robalitem.samplereagent #not finished
    $DATUM = $robalitem.prodtime
    $orderamount = $robalitem.orderamount -replace ' ',''
    $orderamount = $orderamount -replace ',',''
    $orderamount = if([int]$orderamount.Length -eq 5){($orderamount.Insert(2,','))}elseif([int]$orderamount.Length -eq 4){($orderamount.Insert(1,','))}else{$orderamount}
    #$order = $orderamount -replace ' ',''
    

    #$username = "(Script)"
    
    $LINA = “ROBAL “ + “$robal”

    #if($robal -ge 10){ $sealtestrobal = “ROBAL-“ + "0" + “$robal”}else{ $sealtestrobal = “ROBAL-“ + "00" + “$robal”}
    #if($robal -lt 10){ $sealtestrobal = "0" + “$robal”}else{ $sealtestrobal = $robal}
    [string]$sealtestrobal = if(($robal.ToString().length) -gt 1){"ROBAL-0"+$robal}else{"ROBAL-00"+$robal}


    if($robalitem.kundsr -eq "700-6052"){

        $MAPPNAMN = “R” + “$robal” + “ - “  + “$assay #” + “$lsp” + “ “ + “- PQC” + ” (” + ”$orderamount” + ”)” + "  KUNDSR"

    }else{

        $MAPPNAMN = “R” + “$robal” + “ - “  + “$assay #” + “$lsp” + “ “ + “- PQC” + ” (” + ”$orderamount” + ”)”

    }

    $createdPath = "\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\1. IPT - KOMMANDE TESTER\$LINA\$MAPPNAMN"

    #START
    #\\SE.CEPHEID.PRI\Cepheid Sweden
    #Set-Location \
    #Set-Location -path '\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\1. IPT - KOMMANDE TESTER\'
    New-item -path "\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\1. IPT - KOMMANDE TESTER\$LINA\$MAPPNAMN" -ItemType Directory
    #Set-location \

    Copy-Item -Path “\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\1. IPT - KOMMANDE TESTER\Mappscript\D10552 Cartridge Seal Integrity Vacuum Test Worksheet.xlsx” -Destination “\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\1. IPT - KOMMANDE TESTER\$LINA\$MAPPNAMN\$DATUM $assay $LSP Seal Test Neg.xlsx”

    $sealtestnegpath = “\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\1. IPT - KOMMANDE TESTER\$LINA\$MAPPNAMN\$DATUM $assay $LSP Seal Test Neg.xlsx”

    Copy-Item -Path “\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\1. IPT - KOMMANDE TESTER\Mappscript\D10552 Cartridge Seal Integrity Vacuum Test Worksheet.xlsx” -Destination “\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\1. IPT - KOMMANDE TESTER\$LINA\$MAPPNAMN\$DATUM $assay $LSP Seal Test Pos.xlsx”

    $sealtestpospath = “\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\1. IPT - KOMMANDE TESTER\$LINA\$MAPPNAMN\$DATUM $assay $LSP Seal Test Pos.xlsx”

    $worksheetpath = create-dir -ASSAYFAM $ASSAYFAM -assay $assay -LSP $LSP -DATUM $DATUM -assaynum $ASSAYNUM -LINA $LINA -MAPPNAMN $MAPPNAMN -returnmat $returnmat -document $document

    Set-Location -path ‘\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\1. IPT - KOMMANDE TESTER\’
    Set-location -path "\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\1. IPT - KOMMANDE TESTER\$LINA\$MAPPNAMN\"
    New-item “\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\1. IPT - KOMMANDE TESTER\$LINA\$MAPPNAMN\$DATUM $assay #$LSP GX files” -ItemType Directory
    Set-Location .. 
    Set-Location ..
    #Copy-Item “IPT PQC Mailtemplat.xlsm” -Destination “$LINA\$MAPPNAMN\$assay #$LSP Mailtemplat.xlsm” #REMOVED
    #Set-Location \

    $mailtemplatpath = "\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\1. IPT - KOMMANDE TESTER\$LINA\$MAPPNAMN\$assay #$LSP Mailtemplat.xlsm”


    #Set-Location \

    #sealtest -excel $excel -ultra $ultra -highpos $highpos -equipment $equipment -batchnr $batchnr -LSP $LSP -POnr $POnr -sealtestrobal $sealtestrobal -sealtestpospath $sealtestpospath -sealtestnegpath $sealtestnegpath -partnr $partnr

    #$excel. ScreenUpdating = $false

    sealtestdatapos -excel $excel -DATUM $DATUM -ultra $ultra -highpos $highpos -equipment $equipment -batchnr $batchnr -LSP $LSP -POnr $POnr -sealtestrobal $sealtestrobal -sealtestpospath $sealtestpospath -partnr $partnr -sealassay $sealassay -carba $carba -hpv $hpv -returnmat $returnmat
   
    sealtestdataneg -excel $excel -DATUM $DATUM -equipment $equipment -batchnr $batchnr -LSP $LSP -POnr $POnr -sealtestrobal $sealtestrobal -sealtestnegpath $sealtestnegpath -partnr $partnr -sealassay $sealassay -carba $carba -hpv $hpv -ultra $ultra -returnmat $returnmat

    worksheet -excel $excel -worksheetpath $worksheetpath -product $product -DATUM $DATUM -LSP $LSP -batchnr $batchnr -partnr $partnr -ultra $ultra -version $version -material $returnmat

    #$excel. ScreenUpdating = $true

    #mailtemplat -excel $excel -mailtemplatpath $mailtemplatpath -mail $mail -LINA $LINA -batchnr $batchnr -LSP $LSP #REMOVED

    #Set-Location \

    #$MES = @(2, 6, 8, 9, 10, 11)

    #if($robal -notin $MES){

    #alla ROBAL har nu blivit flyttad till MES. Signaturlistan är onödig.

    signlist -batchnr $batchnr -POnr $POnr -LSP $LSP -LINA $LINA -MAPPNAMN $MAPPNAMN -assay $assay -document $document -partnr $partnr
    
    #}

    Set-Location "\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\Skiftspecifika dokument\Skift 1\Mahdi\powerpoint\AutoMappscript"

    return $createdPath

}
    
function Create-Dir($ASSAYFAM, $assay, $LSP, $DATUM, $assaynum, $LINA, $MAPPNAMN, $document, $returnmat){

        $matpath = $returnmat.path

        $matpath | Get-ChildItem | ForEach-Object{

            if($_ -like "*Worksheet*"){

                Copy-Item -Path $_.Fullname -Destination “\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\1. IPT - KOMMANDE TESTER\$LINA\$MAPPNAMN\$DATUM $assay $LSP Worksheet $document.xlsx”

            }else{

                Copy-Item -Path $_.Fullname -Destination “\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\1. IPT - KOMMANDE TESTER\$LINA\$MAPPNAMN\”

            }


        }
        
    return “\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\1. IPT - KOMMANDE TESTER\$LINA\$MAPPNAMN\$DATUM $assay $LSP Worksheet $document.xlsx”

    }
function Sealtestdatapos($excel, $DATUM, $ultra, $highpos, $carba, $hpv, $equipment, $batchnr, $LSP, $POnr, $sealtestrobal, $sealtestpospath, $partnr, $sealassay, $returnmat){

    #$book = $excel.Workbooks.Open($sealtestpospath, $false, $false)

    #$file = '\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\Skiftspecifika dokument\Skift 1\Mahdi\powerpoint\AutoMappscript\dev\D10552 Cartridge Seal Integrity Vacuum Test Worksheet.xlsx'

    $excel = New-Object -TypeName OfficeOpenXml.ExcelPackage $sealtestpospath

    $book = $excel.Workbook

    $checkbox = $book.Worksheets.Drawings | ?{$_.As.Control.CheckBox}

    $checkbox | %{$_.Checked = 1}

    $instrumentlist = @("Balance ID Number", "Vacuum Oven ID Number", "Timer ID Number", "OC Mold(s)", "Part Number", "Cartridge Number (LSP)", "PO Number")
    

    foreach($sheet in $book.Worksheets){

        #write-host $sheet.Name

        if($sheet.index -eq 0){
            
            continue

        }
        #elseif($sheet.Name -eq "Replacement Cartridge (11)"){

         #   break

        #}
        elseif($sheet.Name -eq "Datasheet (1)"){

        #write-host "first"

        foreach($instrument in $instrumentlist){

        

            #$cellvalue = $sheet.Cells.Find($instrument)
            $cellvalue = $sheet.Cells| ?{$_.Text -like "*$instrument*"}

            $celladdress = $cellvalue.Address

            $cellrow, $cellcolumn = $sheet.Cells[$celladdress].start.Row, $sheet.Cells[$celladdress].start.column

            #$cellrow, $cellcolumn = $cellvalue.row, $cellvalue.column

            switch($instrument){

                "Part Number"{
                #write-host "part"
                    $sheet.cells[($cellrow + 1), $cellcolumn].value = $partnr 
                    $sheet.cells[($cellrow + 1), ($cellcolumn + 2)].value = $batchnr
                    $sheet.cells[($cellrow + 1), ($cellcolumn + 4)].value = $sealtestrobal

                }
                "Cartridge Number (LSP)"{

                    $sheet.cells[($cellrow + 1), $cellcolumn].value = $lsp                     
                        
                }
                "PO Number"{

                    $sheet.cells[($cellrow + 1), $cellcolumn].value = $POnr
                    $sheet.cells[($cellrow + 1), ($cellcolumn + 2)].value = $sealassay                        

                }
                "Balance ID Number"{

                    $sheet.cells[($cellrow + 1), $cellcolumn].value = $equipment.scalespos
                    $sheet.cells[($cellrow + 1), ($cellcolumn + 2)].value = $equipment.CAL_D_SCALES

                }
                "Vacuum Oven ID Number"{

                    $sheet.cells[($cellrow + 1), $cellcolumn].value = $equipment.OVENSpos
                    $sheet.cells[($cellrow + 1), ($cellcolumn + 2)].value = $equipment.CAL_D_OVENS                     
                        
                }
                "Timer ID Number"{

                    $sheet.cells[($cellrow + 1), $cellcolumn].value = $equipment.TIMERSpos
                    $sheet.cells[($cellrow + 1), ($cellcolumn + 2)].value = $equipment.CAL_D_TIMERSpos                      

                }
                #"Signature"{

                    #$sheet.cells[($cellrow + 1), ($cellcolumn + 2)].value = $DATUM                        

                #}
                "OC Mold(s)"{


                    $sheet.cells[($cellrow + 5), ($cellcolumn + 2)].value = "N/A"
                    $sheet.cells[($cellrow + 5), $cellcolumn].value = "N/A"
                    $sheet.cells[($cellrow + 3), ($cellcolumn + 2)].value = "N/A"
                    $sheet.cells[($cellrow + 3), $cellcolumn].value = "N/A"
                    $sheet.cells[($cellrow + 1), ($cellcolumn + 2)].value = "N/A"
                    $sheet.cells[($cellrow + 1), $cellcolumn].value = "N/A"

                }

            }

        }
            

        }
        else{

        #write-host "else"

            #$sheet.checkboxes() = 1

            foreach($instrument in $instrumentlist){

                #$cellvalue = $sheet.Cells.Find($instrument)
                #$cellrow, $cellcolumn = $cellvalue.row, $cellvalue.column

                $cellvalue = $sheet.Cells| ?{$_.Text -like "*$instrument*"}

                $celladdress = $cellvalue.Address

                $cellrow, $cellcolumn = $sheet.Cells[$celladdress].start.Row, $sheet.Cells[$celladdress].start.column

                switch($instrument){

                    "Balance ID Number"{

                    #write-host "balance"

                        $sheet.cells[($cellrow + 1), $cellcolumn].value = $equipment.scalespos
                        $sheet.cells[($cellrow + 1), ($cellcolumn + 2)].value = $equipment.CAL_D_SCALES

                    }
                    "Vacuum Oven ID Number"{

                        $sheet.cells[($cellrow + 1), $cellcolumn].value = $equipment.OVENSpos
                        $sheet.cells[($cellrow + 1), ($cellcolumn + 2)].value = $equipment.CAL_D_OVENS                        
                        
                    }
                    "Timer ID Number"{

                        $sheet.cells[($cellrow + 1), $cellcolumn].value = $equipment.TIMERSpos
                        $sheet.cells[($cellrow + 1), ($cellcolumn + 2)].value = $equipment.CAL_D_TIMERSpos                       

                    }
                    #"Signature"{

                     #   $sheet.cells[($cellrow + 1), ($cellcolumn + 2)].value = $DATUM                       

                    #}
                    "OC Mold(s)"{


                            $sheet.cells[($cellrow + 5), ($cellcolumn + 2)].value = "N/A"
                            $sheet.cells[($cellrow + 5), $cellcolumn].value = "N/A"
                            $sheet.cells[($cellrow + 3), ($cellcolumn + 2)].value = "N/A"
                            $sheet.cells[($cellrow + 3), $cellcolumn].value = "N/A"
                            $sheet.cells[($cellrow + 1), ($cellcolumn + 2)].value = "N/A"
                            $sheet.cells[($cellrow + 1), $cellcolumn].value = "N/A"

                    }

                }

            }

        }   

    }


    if((-not $ultra) -and (-not $returnmat.doublesampling) -and (-not $highpos) -and (-not $carba) -and (-not $hpv) -and ($returnmat.japan -eq $null)){

        For ($count = 1 ; $count -le 5 ; $count++){

            #$cellvalue = $sheet.Cells.Find("Cartridge ID")
            #$cellrow, $cellcolumn = $cellvalue.row, $cellvalue.column

            $oddbagnumber = $count * 2 – 1
            $evenbagnumber = $count * 2 
            $sheet = $book.worksheets[($count)]

            $cellvalue = $sheet.Cells| ?{$_.Text -like "*Cartridge ID*"}

            $celladdress = $cellvalue.Address

            $cellrow, $cellcolumn = $sheet.Cells[$celladdress].start.Row, $sheet.Cells[$celladdress].start.column    

            If ($count -lt 5){

                For ($incount = 01 ; $incount -le 10 ; $incount++){

                    If ($incount -lt 10){

                        $sheet.cells[($cellrow + $incount * 2), $cellcolumn].value = “0$oddbagnumber” + “_1_” + “1$incount”
                        $sheet.cells[(($cellrow + 20) + $incount * 2), $cellcolumn].value = “0$evenbagnumber” + “_1_” + “1$incount”
                    }
                    Else {

                        $sheet.cells[($cellrow + $incount * 2), $cellcolumn].value = “0$oddbagnumber” + “_1_” + “20”
                        $sheet.cells[(($cellrow + 20) + $incount * 2), $cellcolumn].value = “0$evenbagnumber” + “_1_20”
                    }
                }
            }

            Else {

                For ($incount = 01 ; $incount -le 10 ; $incount++){

                    If ($incount -lt 10){

                        $sheet.cells[($cellrow + $incount * 2), $cellcolumn].value = “0$oddbagnumber” + “_1_” + “1$incount”
                        $sheet.cells[(($cellrow + 20) + $incount * 2), $cellcolumn].value = “$evenbagnumber” + “_1_” + “1$incount”
                    }
                    Else {

                        $sheet.cells[($cellrow + $incount * 2), $cellcolumn].value = “0$oddbagnumber” + “_1_” + “20”
                        $sheet.cells[(($cellrow + 20) + $incount * 2), $cellcolumn].value = “$evenbagnumber” + “_1_20”
                    }
                }
            }

        }
    }

        #MTB ULTRA
    
    elseif($returnmat.doublesampling -eq $true){

        For ($count = 1 ; $count -le 5 ; $count++){

            $oddbagnumber = $count * 2 – 1
            $evenbagnumber = $count * 2 
            #$sheet = $book.worksheets.item(1 + $count)
            #$cellvalue = $sheet.Cells.Find("Cartridge ID")
            #$cellrow, $cellcolumn = $cellvalue.row, $cellvalue.column 

            $sheet = $book.worksheets[($count)]

            $cellvalue = $sheet.Cells| ?{$_.Text -like "*Cartridge ID*"}

            $celladdress = $cellvalue.Address

            $cellrow, $cellcolumn = $sheet.Cells[$celladdress].start.Row, $sheet.Cells[$celladdress].start.column


            If ($count -lt 5){

                For ($incount = 01 ; $incount -le 5 ; $incount++){

                        If ($incount -lt 5){ 

                        $sheet.cells[($cellrow + $incount * 2), $cellcolumn].value = “0$oddbagnumber” + “_0_” + “1$incount” + “X”
                        $sheet.cells[(($cellrow + 20) + $incount * 2), $cellcolumn].value = “0$evenbagnumber” + “_0_” + “1$incount” + “X”
                    }

                        Else { 

                        $sheet.cells[($cellrow + $incount * 2), $cellcolumn].value = “0$oddbagnumber” + “_0_” + “1$incount” + “X”
                        $sheet.cells[($cellrow + 30), $cellcolumn].value = “0$evenbagnumber” + “_0_15X”
                        $sheet.cells[($cellrow + 40), $cellcolumn].value = “0$evenbagnumber” + “_0_20X”
                    }
                }
                For ($incount = 01 ; $incount -le 5 ; $incount++){ 

                    $incountsecondhalf = $incount + 5

                    If ($incount -lt 5){ 

                        $sheet.cells[(($cellrow + 10) + $incount * 2), $cellcolumn].value = “0$oddbagnumber” + “_0_” + “1$incountsecondhalf” + “+”
                        $sheet.cells[(($cellrow + 30) + $incount * 2), $cellcolumn].value = “0$evenbagnumber” + “_0_” + “1$incountsecondhalf” + “+”
                    }
                    Else { 

                        $sheet.cells[(($cellrow + 10) + $incount * 2), $cellcolumn].value = “0$oddbagnumber” + “_0_” + “20” + “+”
                        $sheet.cells[(($cellrow + 30) + $incount * 2), $cellcolumn].value = “0$evenbagnumber” + “_0_20” + “+”
                    }

                }
            }

            Else {

                For ($incount = 01 ; $incount -le 5 ; $incount++){

                    If ($incount -lt 5){

                        $sheet.cells[($cellrow + $incount * 2), $cellcolumn].value = “0$oddbagnumber” + “_0_” + “1$incount” + “X”
                        $sheet.cells[(($cellrow + 20) + $incount * 2), $cellcolumn].value = “$evenbagnumber” + “_0_” + “1$incount” + “X”
                    }
                    Else {

                        $sheet.cells[($cellrow + $incount * 2), $cellcolumn].value = “0$oddbagnumber” + “_0_” + “$incount” + “X”
                        $sheet.cells[($cellrow + 30), $cellcolumn].value = “0$evenbagnumber” + “_0_15X”
                        $sheet.cells[(($cellrow + 20) + $incount * 2), $cellcolumn].value = “$evenbagnumber” + “_0_20X”

                    }
                }
            }
            For ($incount = 01 ; $incount -le 5 ; $incount++){
     
                $sheet = $book.worksheets[5]
                $incountsecondhalf = $incount + 5
                If ($incount -lt 5){ 

                    $sheet.cells[(($cellrow + 10) + $incount * 2), $cellcolumn].value = “0$oddbagnumber” + “_0_” + “1$incountsecondhalf” + “+”
                    $sheet.cells[(($cellrow + 30) + $incount * 2), $cellcolumn].value = “$evenbagnumber” + “_0_” + “1$incountsecondhalf” + “+”
                }
                Else { 

                    $sheet.cells[(($cellrow + 10) + $incount * 2), $cellcolumn].value = “0$oddbagnumber” + “_0_” + “20” + “+”
                    $sheet.cells[($cellrow + 10), $cellcolumn].value = “0$oddbagnumber” + “_0_15X”
                    $sheet.cells[($cellrow + 30), $cellcolumn].value = “$evenbagnumber” + “_0_15X”
                    $sheet.cells[(($cellrow + 30) + $incount * 2), $cellcolumn].value = “$evenbagnumber” + “_0_20” + “+”
                }
            }


        }
    }

    elseif(($highpos -eq $true) -and ($returnmat.japan -eq $null)){

        #$cellvalue = $sheet.Cells.Find("Cartridge ID")
        #$cellrow, $cellcolumn = $cellvalue.row, $cellvalue.column       

        For ($count = 1 ; $count -le 5 ; $count++){
            $sheet = $book.worksheets[($count)]

            $cellvalue = $sheet.Cells| ?{$_.Text -like "*Cartridge ID*"}

            $celladdress = $cellvalue.Address

            $cellrow, $cellcolumn = $sheet.Cells[$celladdress].start.Row, $sheet.Cells[$celladdress].start.column

            If ($count -le 5){
                # Determine what bags should be entered and add leading zeros if necessary
                $FIRSTBAG = "0" + (2 * $count - 1).ToString()
                $SECONDBAG = If ( (2 * $count) -lt 10) { "0" + (2 * $count).ToString() } Else { 2 * $count }
                $BAGS = $FIRSTBAG, $SECONDBAG

                # Fill in low pos replicates
                $ROW = $cellrow + 2
                Foreach ($i in $BAGS) {
                    For ($REPLICATE = 11; $REPLICATE -le 18; $REPLICATE++) {
                        $sheet.cells[$ROW,$cellcolumn].value = "${i}_1_${REPLICATE}"
                        $ROW++ 
                        $ROW++
                    }

                    # Fill in last two replicates
                    $sheet.cells[$ROW,$cellcolumn].value = If ($HIGHPOS) { "${i}_2_19" } Else { "${i}_1_19" }
                    $ROW++
                    $ROW++
                    $sheet.cells[$ROW,$cellcolumn].value = If ($HIGHPOS) { "${i}_2_20" } Else { "${i}_1_20" }
                    $ROW++
                    $ROW++
                    }
            }
        }

    }

        #JAPAN ASSAYS

    elseif($returnmat.japan -eq "MTB" ){
    
        For ($count = 1 ; $count -le 5 ; $count++){
            $sheet = $book.worksheets[($count)]

            $cellvalue = $sheet.Cells| ?{$_.Text -like "*Cartridge ID*"}

            $celladdress = $cellvalue.Address

            $cellrow, $cellcolumn = $sheet.Cells[$celladdress].start.Row, $sheet.Cells[$celladdress].start.column

            If ($count -le 5){
                # Determine what bags should be entered and add leading zeros if necessary
                $FIRSTBAG = "0" + (2 * $count - 1).ToString()

                #write-host $FIRSTBAG

                $SECONDBAG = If ( (2 * $count) -lt 10) { "0" + (2 * $count).ToString() } Else { 2 * $count }

                #write-host $secondbag

                $BAGS = $FIRSTBAG, $SECONDBAG

                # Fill in low pos replicates
                $ROW = $cellrow + 2
                Foreach ($i in $BAGS) {

                    For ($REPLICATE = 11; $REPLICATE -le 16; $REPLICATE++) {

                        $sheet.cells[$ROW,$cellcolumn].value = "${i}_1_${REPLICATE}"
                        $ROW++ 
                        $ROW++
                    }

                    # Fill in last two replicates
                    $sheet.cells[$ROW,$cellcolumn].value = If ($HIGHPOS) { "${i}_2_17" } Else { "${i}_1_17" }
                    $ROW++
                    $ROW++
                    $sheet.cells[$ROW,$cellcolumn].value = If ($HIGHPOS) { "${i}_3_18" } Else { "${i}_1_18" }
                    $ROW++
                    $ROW++
                    $sheet.cells[$ROW,$cellcolumn].value = If ($HIGHPOS) { "${i}_4_19" } Else { "${i}_1_19" }
                    $ROW++
                    $ROW++
                    $sheet.cells[$ROW,$cellcolumn].value = If ($HIGHPOS) { "${i}_5_20" } Else { "${i}_1_20" }
                    $ROW++
                    $ROW++
                    }
            }

        }    
    
    
    
    }
    elseif($returnmat.japan -eq "CTNG"){

        For ($count = 1 ; $count -le 5 ; $count++){
            $sheet = $book.worksheets[($count)]

            $cellvalue = $sheet.Cells| ?{$_.Text -like "*Cartridge ID*"}

            $celladdress = $cellvalue.Address

            $cellrow, $cellcolumn = $sheet.Cells[$celladdress].start.Row, $sheet.Cells[$celladdress].start.column


            If ($count -le 5){
                # Determine what bags should be entered and add leading zeros if necessary
                $FIRSTBAG = "0" + (2 * $count - 1).ToString()

                #write-host $FIRSTBAG

                $SECONDBAG = If ( (2 * $count) -lt 10) { "0" + (2 * $count).ToString() } Else { 2 * $count }

                #write-host $secondbag

                $BAGS = $FIRSTBAG, $SECONDBAG

                # Fill in low pos replicates
                $ROW = $cellrow + 2
                Foreach ($i in $BAGS) {

                    For ($REPLICATE = 11; $REPLICATE -le 18; $REPLICATE++) {

                        $sheet.cells[$ROW,$cellcolumn].value = "${i}_1_${REPLICATE}"
                        $ROW++ 
                        $ROW++
                    }

                    # Fill in last two replicates
                    $sheet.cells[$ROW,$cellcolumn].value = If ($HIGHPOS) { "${i}_2_19" } Else { "${i}_1_19" }
                    $ROW++
                    $ROW++
                    $sheet.cells[$ROW,$cellcolumn].value = If ($HIGHPOS) { "${i}_3_20" } Else { "${i}_1_20" }
                    $ROW++
                    $ROW++
                    }
            }




        }

    }
    elseif(($returnmat.japan -eq "CDIFF") -or ($returnmat.japan -eq "MRSA")){

        For ($count = 1 ; $count -le 5 ; $count++){
            $sheet = $book.worksheets[($count)]

            $cellvalue = $sheet.Cells| ?{$_.Text -like "*Cartridge ID*"}

            $celladdress = $cellvalue.Address

            $cellrow, $cellcolumn = $sheet.Cells[$celladdress].start.Row, $sheet.Cells[$celladdress].start.column

            If ($count -le 5){
                # Determine what bags should be entered and add leading zeros if necessary
                $FIRSTBAG = "0" + (2 * $count - 1).ToString()

                #write-host $FIRSTBAG

                $SECONDBAG = If ( (2 * $count) -lt 10) { "0" + (2 * $count).ToString() } Else { 2 * $count }

                #write-host $secondbag

                $BAGS = $FIRSTBAG, $SECONDBAG

                # Fill in low pos replicates
                $ROW = $cellrow + 2
                Foreach ($i in $BAGS) {

                    For ($REPLICATE = 11; $REPLICATE -le 17; $REPLICATE++) {

                        $sheet.cells[$ROW,$cellcolumn].value = "${i}_1_${REPLICATE}"
                        $ROW++ 
                        $ROW++
                    }

                    # Fill in last two replicates
                    $sheet.cells[$ROW,$cellcolumn].value = If ($HIGHPOS) { "${i}_2_18" } Else { "${i}_1_17" }
                    $ROW++
                    $ROW++
                    $sheet.cells[$ROW,$cellcolumn].value = If ($HIGHPOS) { "${i}_2_19" } Else { "${i}_1_18" }
                    $ROW++
                    $ROW++
                    $sheet.cells[$ROW,$cellcolumn].value = If ($HIGHPOS) { "${i}_3_20" } Else { "${i}_1_19" }
                    $ROW++
                    $ROW++
                    }
            }




        }

    }


        #CARBA
    if($carba){
    
        #write-host "if entered"

        #$cellvalue = $sheet.Cells.Find("Cartridge ID")
        #$cellrow, $cellcolumn = $cellvalue.row, $cellvalue.column
        $HIGHPOS = $true

        For ($count = 1 ; $count -le 4 ; $count++){

            $sheet = $book.worksheets[($count)]

            $cellvalue = $sheet.Cells| ?{$_.Text -like "*Cartridge ID*"}

            $celladdress = $cellvalue.Address

            $cellrow, $cellcolumn = $sheet.Cells[$celladdress].start.Row, $sheet.Cells[$celladdress].start.column

            #tog din kod niklas, den var skitgrym

            If ($count -le 4){

                # Determine what bags should be entered and add leading zeros if necessary
                $FIRSTBAG = If ( (3 * $count - 2) -lt 10 ) { "0" + (3 * $count - 2).ToString() } Else { 3 * $count - 2 }
                $SECONDBAG = If ( (3 * $count - 1) -lt 10) { "0" + (3 * $count - 1).ToString() }
                $THIRDBAG = If ( (2 * $count) -lt 10) { "0" + (3 * $count).ToString() }
                $BAGS = $FIRSTBAG, $SECONDBAG, $THIRDBAG

                # Fill in low pos replicates

                $ROW = $cellrow + 2

                Foreach ($i in $BAGS) {
                    For ($REPLICATE = 15; $REPLICATE -le 18; $REPLICATE++) {
                        $sheet.cells[$ROW, $cellcolumn].Value = "${i}_1_${REPLICATE}"
                        $ROW++ 
                        $ROW++
                    }

                    # Fill in last two replicates
                    $sheet.cells[$ROW,$cellcolumn].Value = If ($HIGHPOS) { "${i}_2_19" } Else { "${i}_1_19" }
                    $ROW++
                    $ROW++
                    $sheet.cells[$ROW,$cellcolumn].Value = If ($HIGHPOS) { "${i}_2_20" } Else { "${i}_1_20" }
                    $ROW++
                    $ROW++

                    If ($i -eq 10) { Break }
                }

            }
        }    
    
    
          

    }

        #HPV
    if($hpv){

        #$cellvalue = $sheet.Cells.Find("Cartridge ID")
        #$cellrow, $cellcolumn = $cellvalue.row, $cellvalue.column
        For ($count = 1 ; $count -le 10 ; $count++){
            $oddbagnumber = $count * 2 – 1
            $evenbagnumber = $count * 2 
            $sheet = $book.worksheets[($count)]

            $cellvalue = $sheet.Cells| ?{$_.Text -like "*Cartridge ID*"}

            $celladdress = $cellvalue.Address

            $cellrow, $cellcolumn = $sheet.Cells[$celladdress].start.Row, $sheet.Cells[$celladdress].start.column

            If ($count -lt 10){
            For ($incount = 01 ; $incount -le 14 ; $incount++){
            $newcount = $incount + 6
            If ($incount -lt 14){ If ($incount -lt 4){ $sheet.cells[($cellrow + $incount * 2), $cellcolumn].value = “0$count” + “_1_” + “0$newcount”} Else {$sheet.cells[($cellrow + $incount * 2), $cellcolumn].value = “0$count” + “_1_” + “$newcount”}}
            Else {
            $sheet.cells[($incount * 2 - 1), $cellcolumn].value = “0$count” + “_2_19”  
            $sheet.cells[($cellrow + $incount * 2), $cellcolumn].value = “0$count” + “_2_20” }}}
            
        
            #bag 10
            Else{ For ($incount = 01 ; $incount -le 14 ; $incount++){ 
            $newcount = $incount + 6
            
            If ($incount -lt 14){ If ($incount -lt 4){ $sheet.cells[($cellrow + $incount * 2), $cellcolumn].value = “$count” + “_1_” + “0$newcount”} Else {$sheet.cells[($cellrow + $incount * 2), $cellcolumn].value = “$count” + “_1_” + “$newcount”}}
            Else {
            $sheet.cells[(($cellrow - 1) + $incount * 2 - 1), $cellcolumn].value = “$count” + “_2_19”  
            $sheet.cells[($cellrow + $incount * 2), $cellcolumn].value = “$count” + “_2_20” }
            }}

        }

    }

        #N/A FILL

    foreach($sheet in $book.Worksheets){

        if($sheet.Index -eq 0){continue}#elseif($sheet.Name -eq "Replacement Cartridge (11)"){break}

        #$cellvalue = $sheet.Cells.Find("Cartridge ID")
        #$cellrow, $cellcolumn = $cellvalue.row, $cellvalue.column

        $cellvalue = $sheet.Cells| ?{$_.Text -like "*Cartridge ID*"}

        $celladdress = $cellvalue.Address

        $cellrow, $cellcolumn = $sheet.Cells[$celladdress].start.Row, $sheet.Cells[$celladdress].start.column

        #if(($cellvalue + 1) -ne ''){
        
            #$cellvalue = $sheet.Cells.Find("Name Of Tester")
            #$cellrow, $cellcolumn = $cellvalue.row, $cellvalue.column
            #$cellcolumn = $cellcolumn + $sheet.cells[$cellrow, $cellcolumn).MergeArea.Columns.Count   

            $cellvalue = $sheet.Cells| ?{$_.Text -like "*Name of Tester*"}

           if($cellvalue.Merge){


                $id = $sheet.GetMergeCellId($cellvalue.start.Row, $cellvalue.start.Column)

                $range = $sheet.MergedCells[$id-1]

                if($sheet.Cells[$range].Start.Row -ne $sheet.Cells[$range].End.Row){
                
                    $cellrow, $cellcolumn = $sheet.Cells[$range].Start.Row, $sheet.Cells[$range].End.Column

                }else{

                    $cellrow, $cellcolumn = $sheet.Cells[$range].End.Row, $sheet.Cells[$range].End.Column
                }

           }else{

           $celladdress = $cellvalue.Address

           $cellrow, $cellcolumn = $sheet.Cells[$celladdress].End.Row, $sheet.Cells[$celladdress].End.column
           }

           $cellcolumn++

            #$celladdress = $cellvalue.Address

            #$cellrow, $cellcolumn = $sheet.Cells[$celladdress].start.Row, $sheet.Cells[$celladdress].start.column

            while($true){

            #write-host "entered"

                if(($sheet.cells[($cellrow + 3), ($cellcolumn + 2)].Text) -ne ''){break}
        
                $sheet.cells[($cellrow + 3), ($cellcolumn + 2)].value = “N/A"
                $sheet.cells[($cellrow + 3), ($cellcolumn + 3)].value = “N/A"
                $sheet.cells[($cellrow + 3), ($cellcolumn + 4)].value = “N/A"

                $cellrow = $cellrow - 2

            }        
      
        #}


    }

    $excel.Save()
    $excel.Dispose()

    }

function Sealtestdataneg($excel, $DATUM, $carba, $hpv, $equipment, $batchnr, $LSP, $POnr, $sealtestrobal, $sealtestnegpath, $partnr, $sealassay, $ultra, $returnmat){

    #$file = '\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\Skiftspecifika dokument\Skift 1\Mahdi\powerpoint\AutoMappscript\dev\D10552 Cartridge Seal Integrity Vacuum Test Worksheet.xlsx'

    #$excel = New-Object -TypeName OfficeOpenXml.ExcelPackage $sealtestnegpath

    #$book = $excel.Workbook
                        
    #$instrumentlist = @("Balance ID Number", "Vacuum Oven ID Number", "Timer ID Number", "Signature", "OC Mold(s)", "Part Number", "Cartridge Number (LSP)", "PO Number")

    $excel = New-Object -TypeName OfficeOpenXml.ExcelPackage $sealtestnegpath

    $book = $excel.Workbook

    $checkbox = $book.Worksheets.Drawings | ?{$_.As.Control.CheckBox}

    $checkbox | %{$_.Checked = 1}

    $instrumentlist = @("Balance ID Number", "Vacuum Oven ID Number", "Timer ID Number", "OC Mold(s)", "Part Number", "Cartridge Number (LSP)", "PO Number")

    foreach($sheet in $book.Worksheets){

        #write-host $sheet.Name

        if($sheet.index -eq 0){
            
            continue

        }
        #elseif($sheet.Name -eq "Replacement Cartridge (11)"){

         #   break

        #}
        elseif($sheet.Name -eq "Datasheet (1)"){

        #write-host "first"

        foreach($instrument in $instrumentlist){

        

            #$cellvalue = $sheet.Cells.Find($instrument)
            $cellvalue = $sheet.Cells| ?{$_.Text -like "*$instrument*"}

            $celladdress = $cellvalue.Address

            $cellrow, $cellcolumn = $sheet.Cells[$celladdress].start.Row, $sheet.Cells[$celladdress].start.column

            #$cellrow, $cellcolumn = $cellvalue.row, $cellvalue.column

            switch($instrument){

                "Part Number"{
                #write-host "part"
                    $sheet.cells[($cellrow + 1), $cellcolumn].value = $partnr
                    $sheet.cells[($cellrow + 1), ($cellcolumn + 2)].value = $batchnr
                    $sheet.cells[($cellrow + 1), ($cellcolumn + 4)].value = $sealtestrobal

                }
                "Cartridge Number (LSP)"{

                    $sheet.cells[($cellrow + 1), $cellcolumn].value = $lsp                     
                        
                }
                "PO Number"{

                    $sheet.cells[($cellrow + 1), $cellcolumn].value = $POnr
                    $sheet.cells[($cellrow + 1), ($cellcolumn + 2)].value = $sealassay                        

                }
                "Balance ID Number"{

                    $sheet.cells[($cellrow + 1), $cellcolumn].value = $equipment.SCALESNEG
                    $sheet.cells[($cellrow + 1), ($cellcolumn + 2)].value = $equipment.CAL_D_SCALES

                }
                "Vacuum Oven ID Number"{

                    $sheet.cells[($cellrow + 1), $cellcolumn].value = $equipment.OVENSNEG
                    $sheet.cells[($cellrow + 1), ($cellcolumn + 2)].value = $equipment.CAL_D_OVENS                     
                        
                }
                "Timer ID Number"{

                    $sheet.cells[($cellrow + 1), $cellcolumn].value = $equipment.TIMERSNEG
                    $sheet.cells[($cellrow + 1), ($cellcolumn + 2)].value = $equipment.CAL_D_TIMERSNEG                      

                }
                #"Signature"{

                    #$sheet.cells[($cellrow + 1), ($cellcolumn + 2)].value = $DATUM                        

                #}
                "OC Mold(s)"{


                    $sheet.cells[($cellrow + 5), ($cellcolumn + 2)].value = "N/A"
                    $sheet.cells[($cellrow + 5), $cellcolumn].value = "N/A"
                    $sheet.cells[($cellrow + 3), ($cellcolumn + 2)].value = "N/A"
                    $sheet.cells[($cellrow + 3), $cellcolumn].value = "N/A"
                    $sheet.cells[($cellrow + 1), ($cellcolumn + 2)].value = "N/A"
                    $sheet.cells[($cellrow + 1), $cellcolumn].value = "N/A"

                }

            }

        }
            

        }
        else{

        #write-host "else"

            #$sheet.checkboxes() = 1

            foreach($instrument in $instrumentlist){

                #$cellvalue = $sheet.Cells.Find($instrument)
                #$cellrow, $cellcolumn = $cellvalue.row, $cellvalue.column

                $cellvalue = $sheet.Cells| ?{$_.Text -like "*$instrument*"}

                $celladdress = $cellvalue.Address

                $cellrow, $cellcolumn = $sheet.Cells[$celladdress].start.Row, $sheet.Cells[$celladdress].start.column

                switch($instrument){

                    "Balance ID Number"{

                    #write-host "balance"

                        $sheet.cells[($cellrow + 1), $cellcolumn].value = $equipment.SCALESNEG
                        $sheet.cells[($cellrow + 1), ($cellcolumn + 2)].value = $equipment.CAL_D_SCALES

                    }
                    "Vacuum Oven ID Number"{

                        $sheet.cells[($cellrow + 1), $cellcolumn].value = $equipment.OVENSNEG
                        $sheet.cells[($cellrow + 1), ($cellcolumn + 2)].value = $equipment.CAL_D_OVENS
                        
                    }
                    "Timer ID Number"{

                        $sheet.cells[($cellrow + 1), $cellcolumn].value = $equipment.TIMERSneg
                        $sheet.cells[($cellrow + 1), ($cellcolumn + 2)].value = $equipment.CAL_D_TIMERSneg

                    }
                    #"Signature"{

                     #   $sheet.cells[($cellrow + 1), ($cellcolumn + 2)].value = $DATUM                       

                    #}
                    "OC Mold(s)"{


                            $sheet.cells[($cellrow + 5), ($cellcolumn + 2)].value = "N/A"
                            $sheet.cells[($cellrow + 5), $cellcolumn].value = "N/A"
                            $sheet.cells[($cellrow + 3), ($cellcolumn + 2)].value = "N/A"
                            $sheet.cells[($cellrow + 3), $cellcolumn].value = "N/A"
                            $sheet.cells[($cellrow + 1), ($cellcolumn + 2)].value = "N/A"
                            $sheet.cells[($cellrow + 1), $cellcolumn].value = "N/A"

                    }

                }

            }

        }   

    }
        
    if((-not $carba) -and (-not $hpv)){
        
        For ($count = 1 ; $count -le 5 ; $count++){
       
        #$cellvalue = $sheet.Cells.Find("Cartridge ID")
        #$cellrow, $cellcolumn = $cellvalue.row, $cellvalue.column   

        $oddbagnumber = $count * 2 – 1
        $evenbagnumber = $count * 2 
        $sheet = $book.Worksheets[($count)]
         
        $cellvalue = $sheet.Cells| ?{$_.Text -like "*Cartridge ID*"}

        $celladdress = $cellvalue.Address

        $cellrow, $cellcolumn = $sheet.Cells[$celladdress].start.Row, $sheet.Cells[$celladdress].start.column    

        If ($count -lt 5){

            For ($incount = 01 ; $incount -le 10 ; $incount++){

                If ($incount -lt 10){

                    $sheet.cells[($cellrow + $incount * 2), $cellcolumn].value = “0$oddbagnumber” + “_0_” + “0$incount”
                    $sheet.cells[(($cellrow + 20) + $incount * 2), $cellcolumn].value = “0$evenbagnumber” + “_0_” + “0$incount”
                }
                Else {

                    $sheet.cells[($cellrow + $incount * 2), $cellcolumn].value = “0$oddbagnumber” + “_0_” + “10”
                    $sheet.cells[(($cellrow + 20) + $incount * 2), $cellcolumn].value = “0$evenbagnumber” + “_0_10”
                }
            }
        }

        Else {

            For ($incount = 01 ; $incount -le 10 ; $incount++){

                If ($incount -lt 10){

                    $sheet.cells[($cellrow + $incount * 2), $cellcolumn].value = “0$oddbagnumber” + “_0_” + “0$incount”
                    $sheet.cells[(($cellrow + 20) + $incount * 2), $cellcolumn].value = “$evenbagnumber” + “_0_” + “0$incount”
                }
                Else {

                    $sheet.cells[($cellrow + $incount * 2), $cellcolumn].value = “0$oddbagnumber” + “_0_” + “10”
                    $sheet.cells[(($cellrow + 20) + $incount * 2), $cellcolumn].value = “$evenbagnumber” + “_0_10”
                }
            }
        }

    }
    }
        #CARBA
    if($carba){

        #$cellvalue = $sheet.Cells.Find("Cartridge ID")
        #$cellrow, $cellcolumn = $cellvalue.row, $cellvalue.column

        For ($count = 1 ; $count -le 10 ; $count++){
            
            $sheet = $book.Worksheets[($count)]

            $cellvalue = $sheet.Cells | ?{$_.Text -like "*Cartridge ID*"}

            $celladdress = $cellvalue.Address

            $cellrow, $cellcolumn = $sheet.Cells[$celladdress].start.Row, $sheet.Cells[$celladdress].start.column

            $cellrow++
            $cellrow++

        If ($count -le 10){
            # Determine if leading zeros are necessary
            $BAG = If ( $count -lt 10 ) { "0" + ( $count ).ToString() } Else { $count }

            # Fill in replicates
            $ROW = $cellrow
            For ($REPLICATE = 1; $REPLICATE -le 14; $REPLICATE++) {
                $sheet.cells[$ROW,$cellcolumn].value = If ( $REPLICATE -lt 10 ) { "${BAG}_0_0${REPLICATE}" } Else { "${BAG}_0_${REPLICATE}" }
                $ROW++
                $ROW++
            }
        }
        }
    }
        #ULTRA
    elseif($returnmat.doublesampling -eq $true){

        For ($count = 1 ; $count -le 5 ; $count++){

            $oddbagnumber = $count * 2 – 1
            $evenbagnumber = $count * 2 
            #$sheet = $book.worksheets.item(1 + $count)
            #$cellvalue = $sheet.Cells.Find("Cartridge ID")
            #$cellrow, $cellcolumn = $cellvalue.row, $cellvalue.column 

            $sheet = $book.worksheets[($count)]

            $cellvalue = $sheet.Cells| ?{$_.Text -like "*Cartridge ID*"}

            $celladdress = $cellvalue.Address

            $cellrow, $cellcolumn = $sheet.Cells[$celladdress].start.Row, $sheet.Cells[$celladdress].start.column


            If ($count -lt 5){

                For ($incount = 01 ; $incount -le 5 ; $incount++){

                        If ($incount -lt 5){ 

                        $sheet.cells[($cellrow + $incount * 2), $cellcolumn].value = “0$oddbagnumber” + “_0_” + “0$incount” + “X”
                        $sheet.cells[(($cellrow + 20) + $incount * 2), $cellcolumn].value = “0$evenbagnumber” + “_0_” + “0$incount” + “X”
                    }

                        Else { 

                        $sheet.cells[($cellrow + $incount * 2), $cellcolumn].value = “0$oddbagnumber” + “_0_” + “0$incount” + “X”
                        $sheet.cells[($cellrow + 30), $cellcolumn].value = “0$evenbagnumber” + “_0_05X”
                        $sheet.cells[($cellrow + 40), $cellcolumn].value = “0$evenbagnumber” + “_0_10X”
                    }
                }
                For ($incount = 01 ; $incount -le 5 ; $incount++){ 

                    $incountsecondhalf = $incount + 5

                    If ($incount -lt 5){ 

                        $sheet.cells[(($cellrow + 10) + $incount * 2), $cellcolumn].value = “0$oddbagnumber” + “_0_” + “0$incountsecondhalf” + “+”
                        $sheet.cells[(($cellrow + 30) + $incount * 2), $cellcolumn].value = “0$evenbagnumber” + “_0_” + “0$incountsecondhalf” + “+”
                    }
                    Else { 

                        $sheet.cells[(($cellrow + 10) + $incount * 2), $cellcolumn].value = “0$oddbagnumber” + “_0_” + “10” + “+”
                        $sheet.cells[(($cellrow + 30) + $incount * 2), $cellcolumn].value = “0$evenbagnumber” + “_0_10” + “+”
                    }

                }
            }

            Else {

                For ($incount = 01 ; $incount -le 5 ; $incount++){

                    If ($incount -lt 5){

                        $sheet.cells[($cellrow + $incount * 2), $cellcolumn].value = “0$oddbagnumber” + “_0_” + “0$incount” + “X”
                        $sheet.cells[(($cellrow + 20) + $incount * 2), $cellcolumn].value = “$evenbagnumber” + “_0_” + “0$incount” + “X”
                    }
                    Else {

                        $sheet.cells[($cellrow + $incount * 2), $cellcolumn].value = “0$oddbagnumber” + “_0_” + “$incount” + “X”
                        $sheet.cells[($cellrow + 30), $cellcolumn].value = “0$evenbagnumber” + “_0_05X”
                        $sheet.cells[(($cellrow + 20) + $incount * 2), $cellcolumn].value = “$evenbagnumber” + “_0_10X”

                    }
                }
            }
            For ($incount = 01 ; $incount -le 5 ; $incount++){
     
                $sheet = $book.worksheets[5]
                $incountsecondhalf = $incount + 5
                If ($incount -lt 5){ 

                    $sheet.cells[(($cellrow + 10) + $incount * 2), $cellcolumn].value = “0$oddbagnumber” + “_0_” + “0$incountsecondhalf” + “+”
                    $sheet.cells[(($cellrow + 30) + $incount * 2), $cellcolumn].value = “$evenbagnumber” + “_0_” + “0$incountsecondhalf” + “+”
                }
                Else { 

                    $sheet.cells[(($cellrow + 10) + $incount * 2), $cellcolumn].value = “0$oddbagnumber” + “_0_” + “10” + “+”
                    $sheet.cells[($cellrow + 10), $cellcolumn].value = “0$oddbagnumber” + “_0_05X”
                    $sheet.cells[($cellrow + 30), $cellcolumn].value = “$evenbagnumber” + “_0_05X”
                    $sheet.cells[(($cellrow + 30) + $incount * 2), $cellcolumn].value = “$evenbagnumber” + “_0_10” + “+”
                }
            }


        }
    }

        #HPV

    if($hpv){
    
        #$cellvalue = $sheet.Cells.Find("Cartridge ID")
        #$cellrow, $cellcolumn = $cellvalue.row, $cellvalue.column

        For ($count = 1 ; $count -le 4 ; $count++){

        $sheet = $book.worksheets[($count)]

        $cellvalue = $sheet.Cells | ?{$_.Text -like "*Cartridge ID*"}

        $celladdress = $cellvalue.Address

        $cellrow, $cellcolumn = $sheet.Cells[$celladdress].start.Row, $sheet.Cells[$celladdress].start.column

        If ($count -le 4){

            # Determine what bags should be entered and add leading zeros if necessary
            $FIRSTBAG = If ( (3 * $count - 2) -lt 10 ) { "0" + (3 * $count - 2).ToString() } Else { 3 * $count - 2 }
            $SECONDBAG = If ( (3 * $count - 1) -lt 10) { "0" + (3 * $count - 1).ToString() }
            $THIRDBAG = If ( (2 * $count) -lt 10) { "0" + (3 * $count).ToString() }
            $BAGS = $FIRSTBAG, $SECONDBAG, $THIRDBAG

            # Fill in neg replicates
            $ROW = $cellrow + 2
            Foreach ($i in $BAGS) {
                For ($REPLICATE = 1; $REPLICATE -le 6; $REPLICATE++) {
                    $sheet.cells[$ROW,$cellcolumn].value = "${i}_0_0${REPLICATE}"
                    $ROW++ 
                    $ROW++
                }


                If ($i -eq 10) { Break }
            }

        }
        }    


    
    }


                #N/A FILL

    foreach($sheet in $book.Worksheets){

        if($sheet.Index -eq 0){continue}#elseif($sheet.Name -eq "Replacement Cartridge (11)"){break}

        #$cellvalue = $sheet.Cells.Find("Cartridge ID")
        #$cellrow, $cellcolumn = $cellvalue.row, $cellvalue.column

        $cellvalue = $sheet.Cells| ?{$_.Text -like "*Cartridge ID*"}

        $celladdress = $cellvalue.Address

        $cellrow, $cellcolumn = $sheet.Cells[$celladdress].start.Row, $sheet.Cells[$celladdress].start.column

        #if(($cellvalue + 1) -ne ''){
        
            #$cellvalue = $sheet.Cells.Find("Name Of Tester")
            #$cellrow, $cellcolumn = $cellvalue.row, $cellvalue.column
            #$cellcolumn = $cellcolumn + $sheet.cells[$cellrow, $cellcolumn).MergeArea.Columns.Count   

            $cellvalue = $sheet.Cells| ?{$_.Text -like "*Name of Tester*"}

           if($cellvalue.Merge){


                $id = $sheet.GetMergeCellId($cellvalue.start.Row, $cellvalue.start.Column)

                $range = $sheet.MergedCells[$id-1]

                if($sheet.Cells[$range].Start.Row -ne $sheet.Cells[$range].End.Row){
                
                    $cellrow, $cellcolumn = $sheet.Cells[$range].Start.Row, $sheet.Cells[$range].End.Column

                }else{

                    $cellrow, $cellcolumn = $sheet.Cells[$range].End.Row, $sheet.Cells[$range].End.Column
                }

           }else{

           $celladdress = $cellvalue.Address

           $cellrow, $cellcolumn = $sheet.Cells[$celladdress].End.Row, $sheet.Cells[$celladdress].End.column
           }

           $cellcolumn++

            #$celladdress = $cellvalue.Address

            #$cellrow, $cellcolumn = $sheet.Cells[$celladdress].start.Row, $sheet.Cells[$celladdress].start.column

            while($true){

            #write-host "entered"

                if(($sheet.cells[($cellrow + 3), ($cellcolumn + 2)].Text) -ne ''){break}
        
                $sheet.cells[($cellrow + 3), ($cellcolumn + 2)].value = “N/A"
                $sheet.cells[($cellrow + 3), ($cellcolumn + 3)].value = “N/A"
                $sheet.cells[($cellrow + 3), ($cellcolumn + 4)].value = “N/A"

                $cellrow = $cellrow - 2

            }        
      
        #}


    }

    $excel.Save()
    $excel.Dispose()

}

function cellconv($sheet, $row, $column){

        $merged = ($sheet.cell($row, $column)).MergedRange().RangeAdress
     
        if($merged.Columnspan -gt 1){

            #write-host "merged cell"

            #$column = $column + $sheet.cells.item($row, $column).MergeArea.Columns.Count

            $column = $column + $merged.ColumnSpan

        }else{

            $column = $column + 1

        }

        $range = $column
        #$range = $sheet.Cells($row, $column).Address($false, $false)

        #write-host $range
      

        #$cell = $sheet.cell($row,$column)

        #if($cell.IsMerged()){
        
         #   $range = $cell.MergedRange()
        
        #}else{
        
        

        #}


        return $range
    }

function worksheet($excel, $worksheetpath, $product, $DATUM, $LSP, $batchnr, $partnr, $ultra, $version, $material){

        #$book = $excel.Workbooks.Open($worksheetpath, $false, $false)

        #$file = "\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\1. IPT - KOMMANDE TESTER\Mappscript\HIV QA XC\D36985 Attachment 1 Xpert HIV-1 Qual XC Product QC Testing Worksheet RevB.xlsx"

        $excel = New-Object -TypeName OfficeOpenXml.ExcelPackage $worksheetpath

        $book = $excel.Workbook

        #$_.Protection.IsProtected = $false; avstängd vid förfrågan

        $book.Worksheets | %{
        
        if($_.Name -notlike "*QC Data*"){$_.Protection.IsProtected = $false}
        
        $_.HeaderFooter.OddHeader.LeftAlignedText = “Part No.: $partnr

Batch No(s).: $batchnr

Cartridge No. (LSP): $LSP “}

        if($ultra){

            worksheet-vacuum -book $book -DATUM $DATUM

        }else{

            worksheet-seal -book $book -DATUM $DATUM
        }
        worksheet-test -book $book -DATUM $DATUM -LSP $LSP -product $product -version $version -material $material

        worksheet-datasummary -book $book -DATUM $DATUM

        worksheet-stats -book $book -DATUM $DATUM

        $excel.save()
        Start-Sleep 2
        $excel.save()
        $excel.Dispose()
    }

function worksheet-test($book, $DATUM, $LSP, $product, $version, $material){

        $sheet = $book.worksheets['Test Summary']

        $nalist = @("Type of Resample", "Bag Numbers Tested Using Infinity",
                    "Minor Visual Failure", "Major Visual Failure",
                    "NC#", "LFI", "Comments", "NCR", "Bag #s Tested on Inifinity",
                    "Extra test due to", "Retest due to False Positive(s)", "Complementary Resample",
                    "TBWT (Sample Type 1) Retest","TBMDR1 Low (Sample Type 2) Retest", "TBMDR1 High (Sample Type 3) Retest", "TBMDR2 Low (Sample Type 4) Retest"
                    "TBMDR2 High (Sample Type 5) Retest", "Medium Positive (Sample Type 1) Retest", "Low Positive (Sample Type 2) Retest","High Positive (Sample Type 3) Retest"
                    "Low Positive (Sample Type 1) Retest", "Medium Positive (Sample Type 2) Retest", "High Positive (Sample Type 3) Retest")

        $namelist = @("Product Catalog/Part No.", "Cartridge No. (LSP)", "Recorded By:")

        foreach($name in $nalist){

           $cellvalue = $sheet.Cells| ?{$_.Text -like "*$name*"}

           if($cellvalue -eq $null){continue}


           if($cellvalue.Merge){

                $id = $sheet.GetMergeCellId($cellvalue.start.Row, $cellvalue.start.Column)

                $range = $sheet.MergedCells[$id-1]

                if($sheet.Cells[$range].Start.Row -ne $sheet.Cells[$range].End.Row){
                
                    $cellrow, $cellcolumn = $sheet.Cells[$range].Start.Row, $sheet.Cells[$range].End.Column

                }else{

                    $cellrow, $cellcolumn = $sheet.Cells[$range].End.Row, $sheet.Cells[$range].End.Column
                }

           }else{

           $celladdress = $cellvalue.Address

           $cellrow, $cellcolumn = $sheet.Cells[$celladdress].End.Row, $sheet.Cells[$celladdress].End.column
           }
       
            #if($cellvalue.rownumber -eq $null){

                #write-host $name " Not Found"
            
               # continue

            #}
        
            if($cellvalue.Text -eq "Comments"){
                
                #$id = $sheet.GetMergeCellId($cellrow, $cellcolumn)

                #$range = $sheet.MergedCells[$id-1]
                
                #$cellrow, $cellcolumn = $sheet.Cells[$range].Start.Row, $sheet.Cells[$range].Start.Column           

                if($cellvalue.Merge){


                    $id = $sheet.GetMergeCellId($cellvalue.start.Row, $cellvalue.start.Column)

                    $range = $sheet.MergedCells[$id-1]

                    if($sheet.Cells[$range].Start.Row -ne $sheet.Cells[$range].End.Row){
                
                        $cellrow, $cellcolumn = $sheet.Cells[$range].Start.Row, $sheet.Cells[$range].Start.Column

                    }else{

                        $cellrow, $cellcolumn = $sheet.Cells[$range].End.Row, $sheet.Cells[$range].Start.Column
                    }

                }else{

                $celladdress = $cellvalue.Address

                $cellrow, $cellcolumn = $sheet.Cells[$celladdress].End.Row, $sheet.Cells[$celladdress].End.column
                }

                $cellrow++

                $sheet.Cells[$cellrow, $cellcolumn].Value = "N/A"


            }elseif($cellvalue.Text -eq "Major Visual Failure"){

                $delrow,$delcolumn = ($cellrow + 1), ($cellcolumn + 1)

                $sheet.Cells[$cellrow,($cellcolumn + 1)].Value = "N/A" #Major Visual Failure

                $sheet.Cells[$delrow,$delcolumn].Value = "N/A" #Delamination
        
            }elseif($cellvalue.Text -like "*Type Of Resample*"){ 
        
                $cellcolumn = $cellcolumn + 1

                $sheet.Cells[$cellrow,$cellcolumn].value = $material.na
                
            }else{
                
                $cellcolumn = $cellcolumn + 1

                $sheet.Cells[$cellrow,$cellcolumn].value = "N/A"
            }
        
            #start-sleep 2
        }

        foreach($name in $namelist){

            if($name -eq "Date:"){

                $cellvalue = $sheet.Cells| ?{$_.Text -like "Date:"} | Select-Object -First 1 
            }else{
            
                $cellvalue = $sheet.Cells| ?{$_.Text -like "*$name*"}

            }

            if($cellvalue -eq $null){continue}

            if($cellvalue.Merge){

                $id = $sheet.GetMergeCellId($cellvalue.start.Row, $cellvalue.start.Column)

                write-host $name

                $range = $sheet.MergedCells[$id-1]

                if($sheet.Cells[$range].Start.Row -ne $sheet.Cells[$range].End.Row){
                
                    $cellrow, $cellcolumn = $sheet.Cells[$range].Start.Row, $sheet.Cells[$range].End.Column

                }else{

                    $cellrow, $cellcolumn = $sheet.Cells[$range].End.Row, $sheet.Cells[$range].End.Column
                }

            }else{

            $celladdress = $cellvalue.Address

            $cellrow, $cellcolumn = $sheet.Cells[$celladdress].End.Row, $sheet.Cells[$celladdress].End.column
            }

            switch($name){

                "Product Catalog/Part No."{

                    $cellcolumn++

                    $sheet.cells[$cellrow,$cellcolumn].value = $product
                }

                "Cartridge No. (LSP)"{
                    $cellcolumn++

                    $sheet.cells[$cellrow,$cellcolumn].value = [int]$LSP

                }

                "Recorded By:"{
                    
                    $cellcolumn++

                    $sheet.cells[$cellrow,$cellcolumn].value = $version + "Ge feedback :)"

                    #$cellvalue = $sheet.cell($cellrow,$newcolumn)
                    #$celladdress = $cellvalue.address
                    #$cellrow, $cellcolumn = $celladdress.rownumber, ($celladdress.columnnumber + $cellvalue.MergedRange().RangeAddress.ColumnSpan)

                    while($true){
                    
                        if($sheet.Cells[$cellrow,$cellcolumn].Text -eq "Date:"){
       
                            $cellcolumn++
                            $sheet.cells[$cellrow,$cellcolumn].value = $DATUM

                            break

                        }else{
                        
                        
                            $cellcolumn++
                       
                        }
                    
                    }

                }

                #"Date:"{

                    #($sheet.cell($cellrow,$newcolumn)).mergedrange().value = $DATUM

                #}



            }
            #start-sleep 2
        }


    }

function worksheet-seal($book, $DATUM){

        $sheet = $book.worksheets[($book.Worksheets | ?{$_.Name -like "*Seal Test*" -or $_.Name -like "*STF*"}).name]

        $cells = $sheet.Cells["A1:S45"]

        $cellrange = $cells | ?{$_.text -ne "" -and $_.text -ne "N/A"}

        $cellvalue = $cellrange | ?{$_.text -like "*of subgroup*"}

        $comments = $cellrange | ?{$_.text -like "*comments*"}

        $date = $cellrange | ?{$_.text -like "date:"}

        $celladdress = $cellvalue.Address

        $cellrow, $cellcolumn = $sheet.Cells[$celladdress].End.Row, $sheet.Cells[$celladdress].End.column

        $bagrow = $Cellrow + 1

        $i = 1

        while($true){

            $cellvalue = $sheet.Cells[$bagrow, $CellColumn].text

            if($cellvalue -eq 0){

                #write-host 'found at row '$bagrow

                $bagrow = $bagrow - 1

                break

            }elseif($i -le 10){

               $sheet.cells[$bagrow, $CellColumn].Value = 20
               $sheet.cells[$bagrow, ($CellColumn + 1)].Value = 0

            }else{
        
               $sheet.cells[$bagrow, $CellColumn].Value = 'N/A'
               $sheet.cells[$bagrow, ($CellColumn + 1)].Value = 'N/A'

            }

            $bagrow++
            $i++
        }

        $namelist = @("Comments", "Date")

        foreach($name in $namelist){

            if($name -eq "Comments"){

                $celladdress = $comments.Address

                $namerow, $namecolumn = $sheet.Cells[$celladdress].End.Row, $sheet.Cells[$celladdress].End.column

                $namerow = $namerow + 1

                $sheet.cells[$namerow, $namecolumn].Value = 'N/A'

            }elseif($name -eq "Date"){

                $celladdress = $date.Address

                $namerow, $namecolumn = $sheet.Cells[$celladdress].End.Row, $sheet.Cells[$celladdress].End.column

                $namecolumn = $namecolumn + 1

                $namerow = $namerow - 1

                $sheet.cells[$namerow, $namecolumn].Value = $DATUM

            }


        }
    }

function worksheet-vacuum($book, $DATUM){

        $sheet = $book.worksheets['Vacuum Seal Data']

        #$cellvalue = $sheet.Cells.Find("Cartridge ID")

        #$cellvalue = ($sheet.CellsUsed() | where{$_.Getstring() -eq "Cartridge ID"}).Address

        $cellvalue = $sheet.Cells| ?{$_.Text -like "*Cartridge ID*"}

        $celladdress = $cellvalue.Address

        $cellrow, $cellcolumn = $sheet.Cells[$celladdress].End.Row, $sheet.Cells[$celladdress].End.column

        $cellrow++

        $ROW = $cellrow

        For ($count = 1 ; $count -le 10 ; $count++){

            #If ($count -le 10){
                # Determine if leading zeros are necessary
                $BAG = If ( $count -lt 10 ) { "0" + ( $count ).ToString() } Else { $count }

                # Fill in replicates
                
                For ($REPLICATE = 1; $REPLICATE -le 20; $REPLICATE++) {
                    #$sheet.cells[$ROW,$cellcolumn].value = If ( $REPLICATE -lt 10 ) { "${BAG}_0_0${REPLICATE}" } Else { "${BAG}_0_${REPLICATE}" }
                
                    #write-host "${BAG}_0_0${REPLICATE}"
                

                    if($REPLICATE -lt 10){$sheet.cells[$ROW,$cellcolumn].value = "${BAG}_0_0${REPLICATE}"}
                    elseif($REPLICATE -eq 10){$sheet.cells[$ROW,$cellcolumn].value = "${BAG}_0_${REPLICATE}"}
                    elseif($REPLICATE -gt 10 -and $REPLICATE -lt 19){$sheet.cells[$ROW,$cellcolumn].value = "${BAG}_1_${REPLICATE}"}
                    else{$sheet.cells[$ROW,$cellcolumn].value = "${BAG}_2_${REPLICATE}"}

                    $sheet.cells[$ROW,($cellcolumn + 1)].value = "P"


                    #write-host $ROW

                    $ROW++

                }
            #}
        }
        

        $cellvalue = $sheet.Cells| ?{$_.Text -eq "10_2_20"}

        #$cellrow, $cellcolumn = $cellvalue.RowNumber, $cellvalue.ColumnNumber
        $celladdress = $cellvalue.Address

        $cellrow, $cellcolumn = $sheet.Cells[$celladdress].End.Row, $sheet.Cells[$celladdress].End.column

        $bagrow = $Cellrow + 1

        $i = 1

        while($true){

            $cellvalue = $sheet.Cells[$bagrow, $CellColumn].text

            $lastcell = $sheet.Cells[$bagrow, ($CellColumn - 1)].text

            if($lastcell.Trim() -eq 250){

                #write-host 'found at row '$bagrow

                #$bagrow = $bagrow - 1

               $sheet.cells[$bagrow, $CellColumn].Value = "N/A"
               $sheet.cells[$bagrow, ($CellColumn + 1)].Value = "N"

                break

            }elseif($i -le 50){

               $sheet.cells[$bagrow, $CellColumn].Value = "N/A"
               $sheet.cells[$bagrow, ($CellColumn + 1)].Value = "N"

            }

            $bagrow++
            $i++
        }


        #$cell = $sheet.cell(258,1)

        #$test = $cell.IsMerged()

        #$range = $cell.MergedRange()

        #$range.Value = "Test2"

        $namelist = @("Comments:", "Date:")

        foreach($name in $namelist){

            #$namecell = ($sheet.Cellsused() | where{$_.Getstring() -eq $name}).Address

            #$namerow, $namecolumn = $namecell.rownumber, $namecell.columnnumber

            $namecell = $sheet.Cells| ?{$_.Text -eq $name}

            $celladdress = $namecell.Address

            $namerow, $namecolumn = $sheet.Cells[$celladdress].End.Row, $sheet.Cells[$celladdress].End.column

            if($name -eq "Comments:"){

                $namerow = $namerow + 1

                $sheet.cells[$namerow, $namecolumn].Value = 'N/A'

            }elseif($name -eq "Date:"){

                $namecolumn = $namecolumn + 1

                #$namerow = $namerow - 1

                $sheet.cells[$namerow, $namecolumn].Value = $DATUM

            }


        }

    }

function worksheet-stats($book, $DATUM){

        $sheet = $book.Worksheets['Statistical Process Control']

        $cells = $sheet.Cells["A1:U54"]

        $cellrange = $cells | ?{$_.text -ne "" -and $_.text -ne "N/A"}

        $cellvalue = $cellrange | ?{$_.text -like "*of subgroup*" -and $_.text -notlike "*input*"}

        $celladdress = $cellvalue.Address

        $cellrow, $cellcolumn = $sheet.Cells[$celladdress].End.Row, $sheet.Cells[$celladdress].End.column

        $comments = $cellrange | ?{$_.text -like "*comments*"}

        $date = $cellrange | ?{$_.text -like "date:"}


        $bagrow = $Cellrow + 1

        $i = 1

        while($true){

            $cellvalue = $sheet.Cells[$bagrow, $CellColumn].text

            if($cellvalue -eq 0){

                #write-host 'found at row '$bagrow

                $bagrow = $bagrow - 1

                break

            }elseif($i -le 10){

               $sheet.Cells[$bagrow, $CellColumn].Value = 20
               $sheet.Cells[$bagrow, ($CellColumn + 1)].Value = 0

            }else{
        
               $sheet.Cells[$bagrow, $CellColumn].Value = 'N/A'
               $sheet.Cells[$bagrow, ($CellColumn + 1)].Value = 'N/A'

            }

            $bagrow++
            $i++
        }

        $namelist = @("Comments", "Date")

        foreach($name in $namelist){


            if($name -eq "Comments"){

                $celladdress = $comments.Address

                $namerow, $namecolumn = $sheet.Cells[$celladdress].End.Row, $sheet.Cells[$celladdress].End.column

                $namerow = $namerow + 1

                $sheet.Cells[$namerow, $namecolumn].Value = 'N/A'

            }elseif($name -eq "Date"){

                $celladdress = $date.Address

                $namerow, $namecolumn = $sheet.Cells[$celladdress].End.Row, $sheet.Cells[$celladdress].End.column

                $namecolumn = $namecolumn + 1

                $namerow = $namerow - 1

                $sheet.Cells[$namerow, $namecolumn].Value = $DATUM

            }


        }

    }

function worksheet-datasummary($book, $DATUM){

        $sheet = $book.Worksheets['Data Summary']
        
        $validation = $sheet.DataValidations

        #$validation.as.Listvalidation

        $cells = $sheet.Cells["A1:H22"]

        $cellrange = $cells | ?{$_.text -ne "" -and $_.text -ne "N/A"}

        $cellvalue = $cellrange | ?{$_.text -like "*Comment*"}

        $nacells = @()

        foreach($val in $validation){
        
            
            $list = $val.As.ListValidation

            if($list -ne $null){
            
                #$sheet.Cells[$list.Address.End.Row]

                $rangearray = ($list.Address.Address).Split(",")

                $rangearray | %{
                
                    $nacells += $sheet.cells[$_] | ?{$_.Style.Border.Bottom.Style -eq "Thin"}
                
                }
            
            }
        
        }

        #$nacells | %{$_.localaddress}

        $nacells | %{$sheet.Cells[$_.Localaddress].value = "N/A"}

        #$sheet.Cells['E50'].Style.Border.Bottom.Style

        $cells = $sheet.Cells["A1:H370"]

        $cellrange = $cells | ?{$_.text -ne "" -and $_.text -ne "N/A"}
        
        $date = $cellrange | ?{$_.Text -eq "Date:"}
        
        $celladdress = $date.Address

        $namerow, $namecolumn = $sheet.Cells[$celladdress].End.Row, $sheet.Cells[$celladdress].End.column

        $namecolumn = $namecolumn + 1

        $sheet.Cells[$namerow, $namecolumn].Value = $DATUM       
    }

function signlist($partnr, $batchnr, $POnr, $LSP, $document, $LINA, $MAPPNAMN, $assay){

        #Copy-Item -Path “QC\QC-1\IPT\1. IPT - KOMMANDE TESTER\Mappscript\D22580 In Process Testing Signature List.docx” -Destination “QC\QC-1\IPT\1. IPT - KOMMANDE TESTER\$LINA\$MAPPNAMN\”

        $path = “\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\1. IPT - KOMMANDE TESTER\Mappscript\”

        #$signpath = “QC\QC-1\IPT\1. IPT - KOMMANDE TESTER\$LINA\$MAPPNAMN\”

        $path | Get-ChildItem | ForEach-Object{if($_ -like "*Signature*"){$path = $_.Fullname; Copy-Item -Path $_.fullname -Destination “\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\1. IPT - KOMMANDE TESTER\$LINA\$MAPPNAMN\”}}
        
        [string]$signpath = (“\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\1. IPT - KOMMANDE TESTER\$LINA\$MAPPNAMN\” | Get-ChildItem | Where-Object{$_ -like "*Signature*"}).FullName

        #$document = Get-WordDocument -FilePath $signpath
        $doc = Get-WordDocument -FilePath $signpath
        
        #'\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\Skiftspecifika dokument\Skift 1\Mahdi\powerpoint\AutoMappscript\dev\D22580 In Process Testing Signature List.docx'

        $doc.ReplaceText("Part No:","Part No: $partnr")

        $doc.ReplaceText("Batch No(s)*:","Batch No(s)*: $batchnr")

        $doc.ReplaceText("Production Order:","Production Order: $POnr")

        $doc.ReplaceText("Cartridge No. (LSP):","Cartridge No. (LSP): $LSP")

        $doc.ReplaceText("Document No.:","Document No.: D12547, $document")

        $doc.Save()

        $doc.dispose()

        $doc = $null

    }


#////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

function get-folderrev ($rowarray){

    $revrefresh = 1
    
    # read-host 'Hämta revisionen? 1. Ja | 2. Nej'

    if($revrefresh -eq 1){

        write-host 'Hämtar alla revisioner i mapparna....'

        #få fram alla revisioner från alla produktmappar

        $revlist = @{}
        $indivmaterialfolder = @()
        $folder = "\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\1. IPT - KOMMANDE TESTER\Mappscript"
        #får fram alla mappar i huvudmappen
        $materialfolders = Get-ChildItem $folder | Where-Object {($_.PSisContainer) -and ($_ -notlike "*mappscript*")`
        -and ($_ -notlike "*GBS*") -and ($_ -notlike "*XP3 SARS-COV2 FLU RSV plus*")} #ser till att mappen "mappscript" inte är inkluderad, Fluvid+ (korrupt worksheet)-
        #-om den skapas från denna mapp för någon anledning och GBS (Ej implementerad).
        #-and ($_ -notlike "*XP3 SARS-COV2 FLU RSV plus*")
        #$materialfolders += New-Object System.IO.FileInfo("\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\Skiftspecifika dokument\Skift 1\Mahdi\powerpoint\AutoMappscript\Fluvid+")


        foreach ($materialfolder in $materialfolders){

            #$materialfolderpath = Join-Path $folder $materialfolder.Name

            $indivmaterialfolder += Get-ChildItem $materialfolder.FullName -Recurse -Include *.xlsx | Select-Object -First 1

        }

        $docfiles = @()

        foreach($file in $indivmaterialfolder){

               if($file.Name.IndexOf('Worksheet') -gt 1){

                $docfiles += $file

                }
        }

        $doclength = $docfiles.count

        $count = 0

        foreach ($indivfile in $docfiles) {
            
            $count++

            $tried = $False

            $percentagecomplete = ($count / $doclength) * 100

            $path = [string]$indivfile

            try{
            
                $spreadsheet = [DocumentFormat.OpenXml.Packaging.SpreadsheetDocument]::Open($path, $true)
                $tried =  $True

            }

            catch{
            
                write-host "Failed opening worksheet " $indivfile
                $tried = $False
                $revname = "N/A"
                $matdoc = $indivfile.Name.Substring(0,6)

                if(-not ($revlist.ContainsKey($matdoc))){$revlist.Add($matdoc, $revname)}
            }

            if($tried){

                $rawdata = ($spreadsheet.WorkbookPart.GetPartById('rId1')).Worksheet.Innertext
                $headertrimmed = $rawdata.replace(' ', '')
                $headertrimmed1 = $headertrimmed.Substring(0, $headertrimmed.IndexOf('Effective'))
                $rev = $headertrimmed1.Substring($headertrimmed1.IndexOf('Rev'))
                $revuntrim = $rev -replace 'Rev',''
                $revname = ($revuntrim.Substring(1)).trim()
                $matdoc = $indivfile.Name.Substring(0,6)
                #write-host $matdoc $rev
                if(-not ($revlist.ContainsKey($matdoc))){$revlist.Add($matdoc, $revname)}
                $spreadsheet.Close()
                Start-Sleep 0.5
            }


            Write-Progress -Activity 'Hämtar alla revisioner i mapparna....' -Status $matdoc -PercentComplete $percentagecomplete

        }

        Write-Progress -Activity 'Hämtar alla revisioner i mapparna....' -Status 'Klart!' -Completed 

        $spreadsheet.Dispose()

        #Sätt i revisionerna i respektive körning

        foreach($item in $rowarray){

            $material = $item.material

            $returnrev = get-matvariables $material

            if($returnrev.matrev -ne "N/A"){$item.productrev = $revlist[$returnrev.matrev]}else{$item.productrev = $returnrev.matrev}

            $item.matrev = $returnrev.matrev

        }

    }

    $return = @($rowarray, $revlist)

    return $rowarray, $revlist
}

function sortcheck-revision ($rowarray, $revarray, $revlist, $agile){

#Sorterar och checkar om revisionen är nuvarande för varje körning
     Write-host "Verifierar revisionerna"
       
    #$updrevisionlist = @()

    #$updrevisionlist.GetType()

    $updrevisionlist = @()

    foreach($rev in $revarray){

        if(!$revlist[$rev.documentname]){continue}

        if($rev.documentrev -ne $revlist[$rev.documentname]){
       
            $updrevisionlist += [ordered]@{
            
            "documentname"  = $rev.documentname
            "documentrev"   = $rev.documentrev
            "date"          = $rev.date
            }
                        
        }    
    
    }


    if($agile -eq "Sharepoint"){    
        $done = @()

        foreach($document in $updrevisionlist){
            
            if($document['documentname'] -notin $done){

                write-host "Sharepoint revision does not match with the current revision of $($document['documentname']). Its either that the non-draft revision has not been uploaded to sharepoint yet or the new revision of $($document['documentname']) does exist in sharepoint. Updating" -ForegroundColor DarkYellow

                "Sharepoint revision does not match with the current revision of $($document['documentname']). Its either that the non-draft revision has not been uploaded to sharepoint yet or the new revision of $($document['documentname']) does exist in sharepoint. Updating,$(((Get-Date -Format "yyyy-MM-dd HH:mm:ss").ToString()))" | Add-Content -Path "\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\Skiftspecifika dokument\Skift 1\Mahdi\powerpoint\AutoMappscript\Log\Log.csv"

                #write-host $document['documentname']

                $updaterevision = sharepointupdate -docname $document['documentname'] -docrev $document['documentrev'] -date $document['date']  -agile $agile
                $done += $document['documentname']


                if($updaterevision){

                    foreach($item in $rowarray){
            
                        if($item['matrev'] -eq $document['documentname']){
                
                            $item.productrev = $document['documentrev']
                
                        }

            
                    }                    

                }


            }else{continue}

        }
    }
    #sharepointupdate -docname $item['matrev'] -docrev $revdoc['documentrev']

    return $rowarray
}

function running-lotcheck ($rowarray){

    $rowarrayloop = @()

    foreach($item in $rowarray){
    
        if($item.robalnr -ne $null){
        
            $rowarrayloop += $item

        }

    }

    $rowarray = $rowarrayloop

    #få fram alla LSPs från alla körningar som har en mapp i kommande tester

    write-host 'Hämtar alla körningar som redan har en mapp....'

    $indivrobal = @()
    $folder = "\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\1. IPT - KOMMANDE TESTER"
    $robalfolders = Get-ChildItem $folder | Select-Object | Where-Object {$_.PSisContainer -and $_ -like "*ROBAL*"}

    foreach ($robalfolder in $robalfolders){

        $robalfolderPath = Join-Path $folder $robalfolder.Name
        $indivrobal += Get-ChildItem $robalfolderPath | Where-Object {$_.PSisContainer}

    }

    $indivrobal += '\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\2. IPT - PÅGÅENDE KÖRNINGAR' | Get-ChildItem | Select-Object | Where-Object {$_.PSisContainer -and $_ -like "*#*"}

    #får tag på alla LSP:en i kommande mapparna och lägger in den i en lista. 
    $mapplsp = $indivrobal | ForEach-Object {$_.Name.Substring($_.Name.IndexOf("#") + 1, 5)}


    foreach($item in $rowarray){

        foreach($mapp in $mapplsp){

            if($mapp -eq $item.lsp){

                $item.mapcreated = "Yes"

            }
        
        }

        $returnmat =  get-matvariables $item.material

        if($returnmat.assay -eq 0){
                    
            if($item.mapcreated -ne "Yes"){
                write-host $item.material "finns ej i mappscript"

                "Produkt finns ej i mappscript för $($item.material),$(((Get-Date -Format "yyyy-MM-dd HH:mm:ss").ToString()))" | Add-Content -Path "\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\Skiftspecifika dokument\Skift 1\Mahdi\powerpoint\AutoMappscript\Log\Log.csv"

                $item.mapcreated = "Yes"

                #TODO: Logg skrivning KLART
                #TODO: Filtrera bort alla material som kan ej skapas mapp KLART
            }
        }

    }

    return $rowarray

    Write-host " "
}

# [NEW] SharePoint job queue bridge helpers
# region Job queue helpers

function Get-PendingBatchJobs{
    param(
        [int]$JobId,
        [string]$BatchId
    )

    if($BatchId){$BatchId = $BatchId.Trim()}

    if(-not $sharepoint){
        return @()
    }

    $errorLogPath = Join-Path -Path $script:LogDirectory -ChildPath 'Errorlog.csv'

    try{
        $script:JobListConnection = Connect-PnPOnline -Url $script:SharePointSettings.ProductionSiteUrl -Tenant $script:SharePointSettings.Tenant -ClientId $script:SharePointSettings.ClientId -CertificateBase64Encoded $script:SharePointSettings.Certificate -ReturnConnection
    }catch{
        Write-Host "Kunde inte ansluta till jobbkön i SharePoint." -BackgroundColor DarkRed
        "Could not connect to SharePoint job queue,$((($_) -replace ",",";")),$($_.InvocationInfo.ScriptLineNumber),$(((Get-Date -Format "yyyy-MM-dd HH:mm:ss").ToString()))" | Add-Content -Path $errorLogPath
        return @()
    }

    try{
        $list = Get-PnPList -Identity $script:SharePointSettings.JobListName -Connection $script:JobListConnection
        $items = Get-PnPListItem -List $list -PageSize 200 -Connection $script:JobListConnection

        $pending = $items | Where-Object { $_.FieldValues.Status -in $script:SharePointSettings.PendingJobStatuses }

        if($JobId){
            $pending = $pending | Where-Object { $_.Id -eq $JobId }
        }

        if($BatchId){
            $pending = $pending | Where-Object { ($_.FieldValues.BatchId) -eq $BatchId }
        }

        $jobs = @()

        foreach($item in $pending){
            $fields = $item.FieldValues
            $jobs += [pscustomobject]@{
                ItemId       = $item.Id
                BatchId      = if($fields.ContainsKey("BatchId")){($fields["BatchId"]).ToString().Trim()}else{$null}
                ProductCode  = if($fields.ContainsKey("ProductCode")){($fields["ProductCode"]).ToString().Trim()}else{$null}
                RobalId      = if($fields.ContainsKey("RobalId")){($fields["RobalId"]).ToString().Trim()}else{$null}
                Status       = $fields.Status
                NPath        = if($fields.ContainsKey("NPath")){($fields["NPath"]).ToString().Trim()}else{$null}
                ErrorMessage = if($fields.ContainsKey("ErrorMessage")){($fields["ErrorMessage"]).ToString().Trim()}else{$null}
                Title        = $fields.Title
            }
        }

        return $jobs
    }catch{
        Write-Host "Kunde inte läsa jobbkön i SharePoint." -BackgroundColor DarkRed
        "Could not read SharePoint job queue,$((($_) -replace ",",";")),$($_.InvocationInfo.ScriptLineNumber),$(((Get-Date -Format "yyyy-MM-dd HH:mm:ss").ToString()))" | Add-Content -Path $errorLogPath
        return @()
    }
}

function Set-BatchJobStatus{
    param(
        [int]$ItemId,
        [string]$Status,
        [string]$ErrorMessage,
        [string]$NPath
    )

    if(-not $ItemId){
        return
    }

    $errorLogPath = Join-Path -Path $script:LogDirectory -ChildPath 'Errorlog.csv'

    $values = @{Status = $Status}

    if($PSBoundParameters.ContainsKey('ErrorMessage')){$values['ErrorMessage'] = $ErrorMessage}
    if($PSBoundParameters.ContainsKey('NPath')){$values['NPath'] = $NPath}

    try{
        if(-not $script:JobListConnection){
            $script:JobListConnection = Connect-PnPOnline -Url $script:SharePointSettings.ProductionSiteUrl -Tenant $script:SharePointSettings.Tenant -ClientId $script:SharePointSettings.ClientId -CertificateBase64Encoded $script:SharePointSettings.Certificate -ReturnConnection
        }

        Set-PnPListItem -List $script:SharePointSettings.JobListName -Identity $ItemId -Values $values -Connection $script:JobListConnection | Out-Null
    }catch{
        "Could not update SharePoint job $ItemId,$((($_) -replace ",",";")),$($_.InvocationInfo.ScriptLineNumber),$(((Get-Date -Format "yyyy-MM-dd HH:mm:ss").ToString()))" | Add-Content -Path $errorLogPath
    }
}

function Find-RobalItemForJob{
    param(
        [array]$RowArray,
        [array]$Lots,
        [pscustomobject]$Job
    )

    $searchSets = @()

    if($Lots){$searchSets += ,$Lots}
    if($RowArray){$searchSets += ,$RowArray}

    foreach($set in $searchSets){

        if(-not $set){continue}

        if($Job.RobalId){
            $match = $set | Where-Object { $_.itemid -eq $Job.RobalId } | Select-Object -First 1
            if($match){return $match}
        }

        if($Job.BatchId){
            $match = $set | Where-Object {
                $batches = ($_.batchnr).ToString().Replace(" ","").Split(",")
                $Job.BatchId -in $batches
            } | Select-Object -First 1

            if($match){return $match}
        }

        if($Job.ProductCode){
            $match = $set | Where-Object { $_.material -eq $Job.ProductCode } | Select-Object -First 1
            if($match){return $match}
        }
    }

    return $null
}

function Process-RobalItem{
    param(
        $robalitem,
        $version
    )

    $folderLogPath = Join-Path -Path $script:LogDirectory -ChildPath 'Folderlog.csv'
    $errorLogPath = Join-Path -Path $script:LogDirectory -ChildPath 'Errorlog.csv'
    $logPath = Join-Path -Path $script:LogDirectory -ChildPath 'Log.csv'

    $returnmat =  get-matvariables $robalitem.material

    if($returnmat.assay -eq 0){
                    
        if($robalitem.mapcreated -ne "Yes"){
            write-host "Produkt finns ej i mappscript"

            "Produkt finns ej i mappscript för $($robalitem.material),$(((Get-Date -Format "yyyy-MM-dd HH:mm:ss").ToString()))" | Add-Content -Path $logPath

            $robalitem.mapcreated = "Yes"
        }

        return @{Success = $false; Message = "Produkt finns ej i mappscript."}
    }

    try{
        $folderPath = create-folder $returnmat $robalitem.robalnr $robalitem.ordernr $robalitem.batchnr $robalitem.lsp $robalitem.samplereagent $robalitem.prodtime -version $version -robalitem $robalitem
        "$(((Get-Date -Format "yyyy-MM-dd HH:mm:ss").ToString())),$($robalitem.robalnr),$($robalitem.lsp),$($robalitem.material),$(($robalitem.batchnr) -replace ',','/'),$($robalitem.ordernr),$(($robalitem.samplereagent) -replace ",","/"),$(($robalitem.orderamount -replace ',','')),$($robalitem.prodtime),$($robalitem.productrev)" | Add-Content -Path $folderLogPath
        $robalitem.mapcreated = "Yes"
        return @{Success = $true; Path = $folderPath}
    }catch{

        Write-Host "An error has occured at function create-folder while trying to build folder for $($robalitem.material) at line: $($_.Exception.InvocationInfo.ScriptLineNumber): " $_ -BackgroundColor DarkRed

        "An error has occured at function create-folder while trying to build folder for $($robalitem.material),$((($_) -replace ",",";")),$($_.InvocationInfo.ScriptLineNumber),$(((Get-Date -Format "yyyy-MM-dd HH:mm:ss").ToString()))" | Add-Content -Path $errorLogPath

        return @{Success = $false; Message = $_.Exception.Message}
    }
}

function Get-LotData{
    param(
        $agile,
        $objexcel = $null,
        $wb = $null,
        $sheet = $null,
        $revsheet = $null
    )

    $errorLogPath = Join-Path -Path $script:LogDirectory -ChildPath 'Errorlog.csv'
    $logPath = Join-Path -Path $script:LogDirectory -ChildPath 'Log.csv'

    $rowarray = @()
    $revarray = @()
    $revlist = @{}
    $revcheck = @()
    $lotsleft = @()

    try{

        try{

            $rowarray, $revarray = refresh -objexcel $objexcel -wb $wb -sheet $sheet -revsheet $revsheet -agile $agile

        }catch{
        
            $rowarray, $revarray = refresh -objexcel $objexcel -wb $wb -sheet $sheet -revsheet $revsheet -agile $agile


        }
        
    }catch{
        
        Write-Host "An error has occured at function refresh at line: ($($_.InvocationInfo.ScriptLineNumber)): " $_ -BackgroundColor DarkRed
                          "An error has occured at function refresh,$((($_) -replace ",",";")),$($_.InvocationInfo.ScriptLineNumber),$(((Get-Date -Format "yyyy-MM-dd HH:mm:ss").ToString()))" | Add-Content -Path $errorLogPath
                          
    }

    if($agile -eq "Agile"){
        try{

            if((Receive-Job 'MailCheck') -gt 0){import -revarray $revarray}
            }catch{Write-Host "An error has occured at function import at line: ($($_.InvocationInfo.ScriptLineNumber)): " $_ -BackgroundColor DarkRed
                              "An error has occured at function import,$((($_) -replace ",",";")),$($_.InvocationInfo.ScriptLineNumber),$(((Get-Date -Format "yyyy-MM-dd HH:mm:ss").ToString()))" | Add-Content -Path $errorLogPath}
    }
    try{

        $rowarray, $revlist = get-folderrev -rowarray $rowarray
        }catch{Write-Host "An error has occured at function get-folderrev at line: ($($_.InvocationInfo.ScriptLineNumber)): " $_ -BackgroundColor DarkRed
                          "An error has occured at function get-folderrev,$((($_) -replace ",",";")),$($_.InvocationInfo.ScriptLineNumber),$(((Get-Date -Format "yyyy-MM-dd HH:mm:ss").ToString()))" | Add-Content -Path $errorLogPath}
    try{

        $revcheck = sortcheck-revision -rowarray $rowarray -revarray $revarray -agile $agile -revlist $revlist
        }catch{Write-Host "An error has occured at function sortcheck-revision at line: ($($_.InvocationInfo.ScriptLineNumber)): " $_ -BackgroundColor DarkRed
                          "An error has occured at function sortcheck-revision,$((($_) -replace ",",";")),$($_.InvocationInfo.ScriptLineNumber),$(((Get-Date -Format "yyyy-MM-dd HH:mm:ss").ToString()))" | Add-Content -Path $errorLogPath}

    try{

        $rowarray = running-lotcheck -rowarray $revcheck
        }catch{Write-Host "An error has occured at function running-lotcheck at line: ($($_.InvocationInfo.ScriptLineNumber)): " $_ -BackgroundColor DarkRed
                          "An error has occured at function running-lotcheck,$((($_) -replace ",",";")),$($_.InvocationInfo.ScriptLineNumber),$(((Get-Date -Format "yyyy-MM-dd HH:mm:ss").ToString()))" | Add-Content -Path $errorLogPath}

    try{

        $lotsleft = totallots -rowarray $rowarray
        }catch{Write-Host "An error has occured at function lotsleft at line: ($($_.InvocationInfo.ScriptLineNumber)): " $_ -BackgroundColor DarkRed
                          "An error has occured at function lotsleft,$((($_) -replace ",",";")),$($_.InvocationInfo.ScriptLineNumber),$(((Get-Date -Format "yyyy-MM-dd HH:mm:ss").ToString()))" | Add-Content -Path $errorLogPath}

    if($lotsleft.count -eq 0){
            
        write-host "No lots remaining to be created, exiting"
        "No lots remaining to be created exiting,$(((Get-Date -Format "yyyy-MM-dd HH:mm:ss").ToString()))" | Add-Content -Path $logPath

    }

    return [pscustomobject]@{
        RowArray = $rowarray
        RevArray = $revarray
        RevList  = $revlist
        LotsLeft = $lotsleft
    }
}

function Invoke-JobQueueProcessing{
    param(
        $agile,
        $version,
        [string]$BatchId,
        [int]$JobId,
        $objexcel = $null,
        $wb = $null,
        $sheet = $null,
        $revsheet = $null
    )

    $logPath = Join-Path -Path $script:LogDirectory -ChildPath 'Log.csv'
    $errorLogPath = Join-Path -Path $script:LogDirectory -ChildPath 'Errorlog.csv'

    $data = Get-LotData -agile $agile -objexcel $objexcel -wb $wb -sheet $sheet -revsheet $revsheet

    $jobs = Get-PendingBatchJobs -JobId $JobId -BatchId $BatchId

    if((!$jobs -or $jobs.Count -eq 0) -and $BatchId){
        $jobs = @([pscustomobject]@{
            ItemId = $null
            BatchId = $BatchId
            ProductCode = $null
            RobalId = $null
            Status = "Manual"
            Title = "Manual-$BatchId"
        })
    }

    if(!$jobs -or $jobs.Count -eq 0){
        Write-Host "Inga jobb att processa." -ForegroundColor DarkYellow
        "No jobs to process,$(((Get-Date -Format "yyyy-MM-dd HH:mm:ss").ToString()))" | Add-Content -Path $logPath
        return @{Success = $true; Processed = 0}
    }

    $processed = 0

    foreach($job in $jobs){

        $processed++

        if($job.ItemId){
            Set-BatchJobStatus -ItemId $job.ItemId -Status "InProgress"
        }

        $robalitem = Find-RobalItemForJob -RowArray $data.RowArray -Lots $data.LotsLeft -Job $job

        if(-not $robalitem){
            $message = "Kunde inte hitta jobb $($job.ItemId) ($($job.BatchId)) i produktionslistan."
            Write-Host $message -BackgroundColor DarkRed
            if($job.ItemId){
                Set-BatchJobStatus -ItemId $job.ItemId -Status "Error" -ErrorMessage $message
            }
            continue
        }

        if($robalitem.mapcreated -eq "Yes"){
            $comment = "Mapp finns redan för batch $($job.BatchId)."
            Write-Host $comment -ForegroundColor DarkYellow
            if($job.ItemId){
                Set-BatchJobStatus -ItemId $job.ItemId -Status "Done"
            }
            continue
        }

        $result = Process-RobalItem -robalitem $robalitem -version $version

        if($result.Success){
            if($job.ItemId){
                Set-BatchJobStatus -ItemId $job.ItemId -Status "Done" -NPath $result.Path
            }
        }else{
            $message = if($result.Message){$result.Message}else{"Okänt fel vid mappskapande."}
            if($job.ItemId){
                Set-BatchJobStatus -ItemId $job.ItemId -Status "Error" -ErrorMessage $message
            }
        }
    }

    return @{Success = $true; Processed = $processed}
}

# endregion

function manualinput($rowarray){
    
    Do{

    #array list
    $robal2 = @()
    $robal6 = @()
    $robal8 = @()
    $robal9 = @()
    $robal10 = @()
    $robal11 = @()
    $robal12 = @()
    $robalall = @()



    foreach($item in $rowarray){

        if($item["mapcreated"] -ne "Yes"){

            $robalall += $item

            switch ($item.robalnr){

                2{$robal2 += $item}
                6{$robal6 += $item}
                8{$robal8 += $item}
                9{$robal9 += $item}
                10{$robal10 += $item}
                11{$robal11 += $item}
                12{$robal12 += $item}

            }

        }

    }


    #User väljer robal

    Write-host " "

    $inputrobal = Read-Host 'Vilken robal (skriv bara numret ex. 9 för Robal 9 eller skriv "A" för att få upp alla Eller "3" för att Exit)?'
    Write-Host ''
    $i = 1

    switch ($inputrobal){
    
        2{$roballoop = $robal2}
        6{$roballoop = $robal6}
        8{$roballoop = $robal8}
        9{$roballoop = $robal9}
        10{$roballoop = $robal10}
        11{$roballoop = $robal11}
        12{$roballoop = $robal12}
        A{$roballoop = $robalall}
    }

    

    #Visa alla ej skapade körningar i en lista beroende på om User valde alla robal eller enskild robal
    if ($roballoop.Count -gt 0){

        foreach($item in $roballoop){
             
    
            Write-Host $i':' "ROBAL"$item.robalnr 'LSP:' $item.lsp 'Material:' $item.material 'Batch(s):' $item.batchnr 'Ordernr:' $item.ordernr 'Sample Reagent:' $item.samplereagent 'Order Amount:' $item.orderamount 'Production time:' $item.prodtime 'Rev: '-NoNewline; if($item.productrev -eq 'FEL REV'){Write-Host $item.productrev -ForegroundColor Red}else{Write-Host $item.productrev -ForegroundColor Magenta}
            Write-Host '--------------------------------------------------------------------------------------------------------------------------------------------------------------'

            $i = $i + 1
        }

    }elseif($inputrobal -eq 3){

        break

    }else{
    
        write-host 'Det finns inga kommande körningar'
        Write-host ''

        Continue

    }

    $maxrobalcount = $roballoop.Count


    $Userinput = Read-Host 'Välj körning som du vill skapa mapp för eller skriv "A" för att skapa mapp för alla: '

    $Userinput = $Userinput.ToUpper()

    if($Userinput -ne 'A'){

        $Userinput = [int]$Userinput

    }

    #Detta gäller om user väljer en mapp att skapa

    if($Userinput -le $maxrobalcount){

        $chosenitem = $Userinput - 1
        $robalitem = $roballoop[$chosenitem]

        write-host ''
        Write-Host "ROBAL"$robalitem.robalnr 'LSP:' $robalitem.lsp 'Material:' $robalitem.material 'Batch(s):' $robalitem.batchnr 'Ordernr:' $robalitem.ordernr 'Sample Reagent:' $robalitem.samplereagent 'Order Amount:' $robalitem.orderamount 'Production time:' $robalitem.prodtime 'Rev: '-NoNewline; Write-Host $robalitem.productrev -ForegroundColor Magenta

        $returnmat =  get-matvariables $robalitem.material

        if($returnmat.assay -eq 0){
        
            write-host "Produkt finns ej i mappscript"

        }else{


            try{

            create-folder $returnmat $robalitem.robalnr $robalitem.ordernr $robalitem.batchnr $robalitem.lsp $robalitem.samplereagent $robalitem.prodtime -version $version
            "$(((Get-Date -Format "yyy-MM-dd HH:mm:ss").ToString())),$($robalitem.robalnr),$($robalitem.lsp),$($robalitem.material),$($robalitem.batchnr),$($robalitem.ordernr),$($robalitem.samplereagent),$(($robalitem.orderamount -replace ',','')),$($robalitem.prodtime),$($robalitem.productrev)" | Add-Content -Path "\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\Skiftspecifika dokument\Skift 1\Mahdi\powerpoint\AutoMappscript\Log\Folderlog.csv"
            $robalitem.mapcreated = 'Yes'
            }catch{
                
                Write-Host "An error has occured at function create-folder while trying to build folder for $($robalitem.material) at line: $($_.InvocationInfo.ScriptLineNumber): " $_ -BackgroundColor DarkRed

                "An error has occured at function create-folder while trying to build folder for $($robalitem.material),$($_),$($_.Exception.InvocationInfo.ScriptLineNumber),$(((Get-Date -Format "yyy-MM-dd HH:mm:ss").ToString()))" | Add-Content -Path "\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\Skiftspecifika dokument\Skift 1\Mahdi\powerpoint\AutoMappscript\Log\Errorlog.csv"
            }

            #write-host "$(((Get-Date -Format "yyy/mm/dd HH:mm").ToString())),$($robalitem.robalnr),$($robalitem.lsp),$($robalitem.material),$($robalitem.batchnr),$($robalitem.ordernr),$($robalitem.samplereagent),$($robalitem.orderamount),$($robalitem.prodtime),$($robalitem.productrev)"



            $inputloop = Read-Host "Klart! Vill du skapa en till mapp? 1. Ja | 2. Nej | 3. Exit "
        }


    #Detta gäller om user vill skapa mapp för allt i listan
    }elseif($userinput -eq 'A'){

        Write-Host 'Du valde alla körningar'

        foreach ($robalitem in $roballoop){

            write-host ''
            Write-Host "ROBAL"$robalitem.robalnr 'LSP:' $robalitem.lsp 'Material:' $robalitem.material 'Batch(s):' $robalitem.batchnr 'Ordernr:' $robalitem.ordernr 'Sample Reagent:' $robalitem.samplereagent 'Order Amount:' $robalitem.orderamount 'Production time:' $robalitem.prodtime 'Rev: '-NoNewline; Write-Host $robalitem.productrev -ForegroundColor Magenta
            Write-host ""
            $returnmat =  get-matvariables $robalitem.material

            if($returnmat.assay -eq 0){
            
                write-host ''
                write-host "Produkt finns ej i mappscript"

            }else{

                try{

                #create-folder $returnmat $robalitem.robalnr $robalitem.ordernr $robalitem.batchnr $robalitem.lsp $robalitem.samplereagent $robalitem.prodtime -version $version
                "$(((Get-Date -Format "yyy-MM-dd HH:mm:ss").ToString())),$($robalitem.robalnr),$($robalitem.lsp),$($robalitem.material),$($robalitem.batchnr),$($robalitem.ordernr),$($robalitem.samplereagent),$(($robalitem.orderamount -replace ',','')),$($robalitem.prodtime),$($robalitem.productrev)" | Add-Content -Path "\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\Skiftspecifika dokument\Skift 1\Mahdi\powerpoint\AutoMappscript\Log\Folderlog.csv"
                $robalitem.mapcreated = 'Yes'
                }catch{
                
                    Write-Host "An error has occured at function create-folder while trying to build folder for $($robalitem.material) at line: $($_.Exception.InvocationInfo.ScriptLineNumber): " $_ -BackgroundColor DarkRed

                    "An error has occured at function create-folder while trying to build folder for $($robalitem.material),$($_),$($_.Exception.InvocationInfo.ScriptLineNumber),$(((Get-Date -Format "yyy-MM-dd HH:mm").ToString()))" | Add-Content -Path "\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\Skiftspecifika dokument\Skift 1\Mahdi\powerpoint\AutoMappscript\Log\Errorlog.csv"
                }

            }

        }

       $inputloop = Read-Host "Klart! Vill du skapa en till mapp? 1. Ja | 2. Nej | 3. Exit " 


    #Om user skrev in ett nummer som inte är med på listan av alla kommande körningar i listan
    }elseif($Userinput -gt $maxrobalcount){
    
        write-host ''
        write-host 'Körning finns ej, skrev du fel? Försök igen'

    }

    }until ($inputloop -eq 3 )
    #stop-process -Id $PID

    return 1
} #ONÖDIG

function import($comms, $revarray){

    write-host "Revision change detected"

    

    $objWord = New-Object -ComObject word.application
    $objWord.Visible = $False
    $objWord.Screenupdating = $false

    $path = '\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\Skiftspecifika dokument\Skift 1\Mahdi\powerpoint\AutoMappscript\PAFiles'
    $temp = '\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\Skiftspecifika dokument\Skift 1\Mahdi\powerpoint\AutoMappscript\temp'
    $archivepath = '\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\Skiftspecifika dokument\Skift 1\Mahdi\powerpoint\AutoMappscript\Archive'

    $path | Get-ChildItem | ForEach-Object{

        $ext = [IO.Path]::GetExtension($_)
        $filepath = $_.FullName
        $filename = $_.BaseName
        if($ext -eq ".zip"){
           
           # TODO: Denna del av kod är specifikt för produktmetoderna som IPT använder dvs att det kommer alltid med en fil som heter innehåller "Testing Procedure". Detta funkar ej för andra som har flera dokument som senare packas ihop till
           # en .zip fil. Måste checka om filen inte är produktmetod specifikt innehåll. 

            $matpath = (get-matvariables $filename).path

            $matpath | Get-ChildItem | ForEach-Object{Remove-Item $_.fullname -Force}

            start-sleep 5

            Expand-Archive -Path $filepath -DestinationPath $matpath
            Expand-Archive -Path $filepath -DestinationPath $temp

            Start-Sleep 1

            $docfile  = ($temp | Get-ChildItem | Select-Object | Where-Object{ $_ -like "*Testing Procedure*" }).fullname
            $docfilename  = ($temp | Get-ChildItem | Select-Object | Where-Object{ $_ -like "*Testing Procedure*" }).basename
            $matdoc = $docfilename.Substring(0,6)

            try{

                $objDoc = $objWord.Documents.Open([string]$docfile, $false, $true)
                $tried =  $True

            }
            catch{

                write-host "Failed opening document and updating revisionlist for" $docfile.BaseName
                $tried = $false
            }

            if($tried){

                $Section = $objDoc.Sections.Item(1)
                $header = $Section.Headers.Item(1)
                $headertext = $header.range.text
                $headertrimmed = $headertext.replace(' ', '')
                $headertrimmed1 = $headertrimmed.Substring(0, $headertrimmed.IndexOf('Effective'))
                $headercleaned = $headertrimmed1.substring(26)
                $headercleanedtrimmed = $headercleaned.trim()
                $revname = $headercleanedtrimmed

                #$revlist.Add($matdoc, $revname)

                $objDoc.Close()
                Start-Sleep 0.5
            }

            foreach($hashtable in $revarray){

                if($matdoc -eq $hashtable['Documentname']){
                
                    $hashtable['Documentrev'] = $revname

                }

            }

            #Set-location $temppath

            #Move-Item -Path ($archivepath + '\*') -Destination $matpath


        }else{

            $matpath = (get-matvariables ($filename.substring(0, 6))).path

            Copy-Item -Path $filepath -Destination $matpath

            #TODO: Ta bort gamla revision filen, nu den kopierar bara till där filen ska vara men om det redan finns en fil med samma eller liknande namn så kommer det antingen orsaka en error eller så kommer det finnas 2+ kopior


            #foreach($document in $revarray){

                #if($_.Contains($document['Documentname'])){
                    
                    #matpath = (get-matvariables $document['Documentname']).path

                    #Copy-Item -Path $filepath -Destination $archivepath

                    #Move-Item -Path $filepath -Destination $matpath

                #}

            #}

        }

        Document-Archive -DocNrPath $filepath -DocNrName $filename -DocRev $revname

        Remove-Item $_.FullName -Force

        $temp | Get-ChildItem | ForEach-Object{Remove-Item $_.fullname -Force}

    }

    $objWord.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($objWord) | Out-Null

    Write-Host 'Updating Revision List'
    
    $revarray | Export-Clixml -path .\Revisionlist.xml

    write-host "Continuing"

    Stop-Job *

    Remove-Job *

    $start = Start-Job -Name "MailCheck" -ScriptBlock {
        
        $path = '\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\Skiftspecifika dokument\Skift 1\Mahdi\powerpoint\AutoMappscript\PAFiles\*'

        while(($count = ($path | Get-ChildItem).count) -ge 0){

            $count

        }

        #start-sleep 3
    }

    start-sleep 5

} #AGILE FUNKTION GAMMAL EXCEL METOD

function Document-Archive($DocNrPath,$DocNrName, $DocRev, $agile){

    $DocNrName = $DocNrName.Substring(0, 6)

    $docfoldername = $DocNrName + " Rev " + $DocRev

    $archivepath = '\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\Skiftspecifika dokument\Skift 1\Mahdi\powerpoint\AutoMappscript\Archive\'

    $docpathcheck = $archivepath + $DocNrName
    
    $docpath = $archivepath + $DocNrName + "\"

    $DocRevPath = $docpath + $docfoldername

    #write-host $DocRevPath

    if(Test-Path -Path $docpathcheck){

        if(-not (Test-Path -Path $DocRevPath)){

            New-Item -Path $docpath -Name $docfoldername -ItemType "Directory" | Out-Null

            start-sleep 0.5

            if($agile -eq "Agile"){

            Expand-Archive -Path $DocNrPath -DestinationPath $DocRevPath

            }else{
            
                $DocNrPath | Get-ChildItem | ForEach-Object{Copy-Item $_.fullname -Destination $DocRevPath}
                
                #Copy-Item -Path $DocNrPath -Destination $DocRevPath
            
            }

        }else{

            Write-Host "Folder exists, continuing"

        }

    }elseif(-not (Test-Path -Path $docpathcheck)){

        New-Item -Path $archivepath -Name $DocNrName -ItemType "Directory" | Out-Null

        start-sleep 0.5

        New-Item -Path $docpath -Name $docfoldername -ItemType "Directory" | Out-Null


        if($agile -eq "Agile"){

            Expand-Archive -Path $DocNrPath -DestinationPath $DocRevPath

        }else{
        
         $DocNrPath | Get-ChildItem | ForEach-Object{Copy-Item $_.fullname -Destination $DocRevPath}

            #Copy-Item -Path $DocNrPath -Destination $DocRevPath
        }
    }

    

}

function refresh($objexcel, $wb, $sheet, $revsheet, $agile){

    #Write-Host "Hämtar information..."

    #$username = "Labuser_ipt@cepheid.com"
    #$password = "Spring2023!!123"

    #$secpass = ConvertTo-SecureString -String $password -AsPlainText -Force

    #[PSCredential]$creds = New-Object System.Management.Automation.PSCredential -ArgumentList $username, $secpass

    #$creds = Import-Clixml -Path "\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\Skiftspecifika dokument\Skift 1\Mahdi\powerpoint\AutoMappscript\cred.cred"

    #$secpass = Import-Clixml -Path "\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\Skiftspecifika dokument\Skift 1\Mahdi\powerpoint\AutoMappscript\cred.cred"

    #$key = Get-Content "\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\Skiftspecifika dokument\Skift 1\Mahdi\powerpoint\AutoMappscript\key.key"

    #[PSCredential]$creds = New-Object System.Management.Automation.PSCredential -ArgumentList $username, ($secpass | ConvertTo-SecureString -Key $key)

    #$thumbprint = "4C711F9531F29BBE90A3531E489AA574A102F29C"

    Write-Host "Hämtar körningar"

    $clientid = $script:SharePointSettings.ClientId

    $tenant = $script:SharePointSettings.Tenant

    $certificateprivatekey = $script:SharePointSettings.Certificate

    $productionSiteUrl = $script:SharePointSettings.ProductionSiteUrl

    Connect-PnPOnline -Url $productionSiteUrl -Tenant $tenant -ClientId $clientid -CertificateBase64Encoded $certificateprivatekey

    $list = Get-PnPList -Identity $script:SharePointSettings.ProductionListName

    $robalitems = Get-PnPListItem -List $list -PageSize 500

    $robal = $robalitems | Where{$_.FieldValues.Work_x0020_Center -like "*ROBAL*" -and $_.FieldValues.Actual_x0020_startdate_x002f__x0 -eq $null `
     -and ($_.FieldValues.Title -match '^\d+$') -and $_.FieldValues.Production_x0020_Day_x002f__x002 -ge ((Get-Date).ToShortDateString())}

    $rowarray = @()

    foreach($item in $robal){
    
        $itemId = $item.Id
        $item = $item.FieldValues
        #(($item.Work_x0020_Center -replace "ROBAL ", "") -replace "0", "").Trim()
        $robalnr = if(($item.Work_x0020_Center -replace "ROBAL ", "").StartsWith("0")){($item.Work_x0020_Center -replace "ROBAL ", "").Replace("0","")}else{($item.Work_x0020_Center -replace "ROBAL ", "")}
        $ordernr = ($item.Title).Replace(" ","")
        $material = ($item.Material).Replace(" ","")
        $batchnr = if($item.SAP_x0020_Batch_x0023__x0020_2 -ne $null){(($item.Batch_x0023_ + "," + $item.SAP_x0020_Batch_x0023__x0020_2).Replace(" ",""))}else{($item.Batch_x0023_).Replace(" ","")} 
        $lsp = ($item.LSP)
        $samplereagent = if($item.PAL_x0020__x002d__x0020_Sample_x){if(($item.PAL_x0020__x002d__x0020_Sample_x).Length -gt 10){(($item.PAL_x0020__x002d__x0020_Sample_x).split("`n")) -join "/"}else{$item.PAL_x0020__x002d__x0020_Sample_x}}
        $orderamount = $item.Order_x0020_quantity
        $prodtime = $item.Production_x0020_Day_x002f__x002.ToString("yyyy MM dd")
        $kundsr = if(($item.Sample_x0020_Reagent_x0020_P_x00) -ne $null){(($item.Sample_x0020_Reagent_x0020_P_x00).replace(" ",""))}else{($item.Sample_x0020_Reagent_x0020_P_x00)}
        $actualsd = $null
        $mapcreated = $null
        $productrev = $null

        if($lsp -ne $null){

            $hashtable = [ordered]@{
                "itemid" = $itemId
                "robalnr" = $robalnr 
                "ordernr" = $ordernr
                "material" = $material
                "batchnr" = $batchnr
                "lsp" = $lsp
                "samplereagent" = $samplereagent
                "orderamount" = $orderamount
                "prodtime" = $prodtime
                "kundsr" = $kundsr
                "actualsd" = $actualsd
                "mapcreated" = $mapcreated
                "productrev" = $productrev
                "matrev" = $null
            }

            $rowarray += $hashtable
        }
    }


    $revarray = Import-Clixml -Path .\Revisionlist.xml 

    $reference = @(
    "D51114"
    "D51113"
    "D51112"
    "D51111"
    "D51110"
    "D48538"
    "D47377"
    "D41929"
    "D39525"
    "D37468"
    "D36985"
    "D31503"
    "D25862"
    "D23916"
    "D22580"
    "D19537"
    "D18272"
    "D17716"
    "D16904"
    "D16898"
    "D16546"
    "D13152"
    "D12547"
    "D10552"
    "D55782"
    "D27089"
    'D66612'
    'D66613'
    'D66614'
    'D66615'
    'D26120'
    'D17716'
    'D21938'
    'D68620'
    'D68621'
    'D37988'
    'D37339'
    'D79377'
    )

    #$reference | %{$_}

    Connect-PnPOnline -Url $script:SharePointSettings.DocumentSiteUrl -Tenant $tenant -ClientId $clientid -CertificateBase64Encoded $certificateprivatekey

    Write-Host "Hämtar revisioner"

    $items = Get-PnPListItem -List "Cepheid" -PageSize 4000
    $items = $items | ?{$_.FileSystemObjectType -eq "Folder"}

    #$items | ?{$_.FieldValues.FileLeafRef -like "*D12547*"}

    $sharepointarray = @()

    foreach($document in $reference){

        $sorthashtable = @{}

        $docfolders = $items | ?{$_.FieldValues.FileLeafRef -like "*$document*"}

        $docfolders | %{$sorthashtable[$_.FieldValues.FileLeafRef] = $_.FieldValues.Last_x0020_Modified}

        $LatestRev = (($sorthashtable.GetEnumerator() | Sort-Object -Property Value -Descending | Select-Object -First 1).Name).split("_")

        $date = [string]($sorthashtable.GetEnumerator() | Sort-Object -Property Value -Descending | Select-Object -First 1).Value

        $hashtable = [ordered]@{
            "documentname" = ($LatestRev[0]).Replace(" ","") 
            #"datenr" = $datenr
            "documentrev" = ($LatestRev[1]).Replace(" ","")
            "date" = ($date).Replace(" ","")
        }

        $sharepointarray += $hashtable
    }

   #$revarray = $sharepointarray

   $revarray = if($agile -eq "Agile"){$revarray}else{$sharepointarray}

   return $rowarray, $revarray
}

function totallots($rowarray){

    
    $cleanarray = @()

    foreach($lot in $rowarray){

        if($lot.mapcreated -ne "Yes"){

            $cleanarray += $lot

        }

    }

    return $cleanarray

}

function main($objexcel, $wb, $sheet, $revsheet, $main, $agile, $version, [string]$BatchId, [int]$JobId){

    # [NEW] Automatic mode now processes SharePoint job queue items instead of free-running all lots.

    if($main -eq $True){

        $result = Invoke-JobQueueProcessing -agile $agile -version $version -BatchId $BatchId -JobId $JobId -objexcel $objexcel -wb $wb -sheet $sheet -revsheet $revsheet

        return $result

    }else{
    
       $data =  Get-LotData -agile $agile -objexcel $objexcel -wb $wb -sheet $sheet -revsheet $revsheet
       $endmanual =  manualinput -rowarray $data.RowArray

       return $endmanual

    }
}

function menu($comms, $version, $agile){

    $endprogram = $false

    

    Do{

    $endmanual = 0

     Clear-Host

     Write-host " "
     Write-Host $version -BackgroundColor DarkGray -ForegroundColor Black
     Write-host " "
     if($agile -eq "Disabled"){write-host "Agile/Sharepoint is disabled" -BackgroundColor DarkRed; Write-host " "}elseif(($agile -eq "Agile") -or ($agile -eq "Sharepoint")){
     Write-Host "Revision Interaction Mode: " -ForegroundColor DarkGray -NoNewline; Write-Host $agile -ForegroundColor DarkYellow; Write-Host " "}

        $userinput = Read-Host "Select Mode:
    
      1. Automatic Mode

      2. Manual Mode

      3. Exit
    
"
        #return $userinput
    
        $continue = $true

        Do{

            switch ($userinput){

                1{

                    $main = $True

                    main -objexcel $objexcel -wb $wb -sheet $sheet -revsheet $revsheet -main $main -agile $agile -version $version -BatchId $BatchId -JobId $JobId

                   
                 }

                2{
                    $main = $False

                    $endmanual = main -objexcel $objexcel -wb $wb -sheet $sheet -revsheet $revsheet -main $main -agile $agile -version $version -BatchId $BatchId -JobId $JobId

                }
                
                3{

                Write-host "Ending script....."

                $endmanual = 1

                $endprogram = $True

                }
                #4: SETTINGS
            }
            
            if($userinput -gt 3 -or $userinput -eq 0){break}

            if ([console]::KeyAvailable){

                $key = [system.console]::readkey($true)

                if ($key.Key -eq 'Q'){

                    Write-host "Key pressed, Discountining...."

                    $continue = $false
                }
            }   

        }until(!$continue -or $endmanual -eq 1)


    }until($endprogram)

} #ONÖDIG

function sharepointupdate($docname, $docrev, $date, $agile){

    $matpath = (get-matvariables $docname).path

    $folderpath = '\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\Skiftspecifika dokument\Skift 1\Mahdi\powerpoint\AutoMappscript\Folder'
    
    $folderpath | Get-ChildItem | ForEach-Object{Remove-Item $_.fullname -Force}

    $sharepointpath = $script:SharePointSettings.DocumentSiteUrl

    $site = "/Cepheid/$($docname)_$($docrev)"

    $fullpath = $sharepointpath + $site + '/'

    $username = "Labuser_ipt@cepheid.com"
    #$password = "Spring2023!!123"

    #$secpass = ConvertTo-SecureString -String $password -AsPlainText -Force

    #$creds = Import-Clixml -Path "\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\Skiftspecifika dokument\Skift 1\Mahdi\powerpoint\AutoMappscript\cred.cred"

    #[PSCredential]$creds = New-Object System.Management.Automation.PSCredential -ArgumentList $username, $secpass

    #$secpass = Import-Clixml -Path "\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\Skiftspecifika dokument\Skift 1\Mahdi\powerpoint\AutoMappscript\cred.cred"

    #$key = Get-Content "\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\Skiftspecifika dokument\Skift 1\Mahdi\powerpoint\AutoMappscript\key.key"

    #[PSCredential]$creds = New-Object System.Management.Automation.PSCredential -ArgumentList $username, ($secpass | ConvertTo-SecureString -Key $key)

    #draft check

    $clientid = $script:SharePointSettings.ClientId

    $tenant = $script:SharePointSettings.Tenant

    $certificateprivatekey = $script:SharePointSettings.Certificate


    #$file = "N:\QC\QC-1\IPT\Skiftspecifika dokument\Skift 1\Mahdi\powerpoint\AutoMappscript\dev\draft"
    
    $downloadrevision = $false
    #$ignorerev = $false

    if(Test-Path -Path "\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\Skiftspecifika dokument\Skift 1\Mahdi\powerpoint\AutoMappscript\draft\$($docname)\$($docrev)\$($docrev).txt"){
    
        #checks to see if the script has previously detected a revision change and created a draft for that specific revision.

        [string]$draftdate = Get-Content -Path "\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\Skiftspecifika dokument\Skift 1\Mahdi\powerpoint\AutoMappscript\draft\$($docname)\$($docrev)\$($docrev).txt"

        if($draftdate -eq $date){

            Write-host "Revision $($docname) $($docrev) from sharepoint is a draft. Continuing without downloading" -ForegroundColor DarkYellow

            "Revision $($docname)_$($docrev) from sharepoint is a draft. Continuing without downloading,$(((Get-Date -Format "yyyy-MM-dd HH:mm:ss").ToString()))" | Add-Content -Path "\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\Skiftspecifika dokument\Skift 1\Mahdi\powerpoint\AutoMappscript\Log\Log.csv"
     
            $updaterevision = $false

        }else{$downloadrevision = $true}

    }else{
    
        #draft doesnt exist, download revision

        "Revision $($docname)_$($docrev) draft does not exist in draft folder. Downloading,$(((Get-Date -Format "yyyy-MM-dd HH:mm:ss").ToString()))" | Add-Content -Path "\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\Skiftspecifika dokument\Skift 1\Mahdi\powerpoint\AutoMappscript\Log\Log.csv"

        $downloadrevision = $true

    }
    
    if($downloadrevision){

        try{Get-PnPAppAuthAccessToken -Connection $connection | Out-Null}catch{

            write-host "Logging into sharepoint"

            $connection = Connect-PnPOnline -Url $sharepointpath -Tenant $tenant -ClientId $clientid -CertificateBase64Encoded $certificateprivatekey
            
            #-Credentials $creds -Verbose
        }

        Write-Host "Getting documents of $($docname)"

        Get-PnPFolderItem -FolderSiteRelativeUrl $site -Connection $connection | ForEach-Object{Get-PnPFile -Url ($site + '/' + $_.name) -Path '\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\Skiftspecifika dokument\Skift 1\Mahdi\powerpoint\AutoMappscript\Folder' -Filename $_.name -AsFile -Connection $connection}

        Start-Sleep 2

        $workpath = '\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\Skiftspecifika dokument\Skift 1\Mahdi\powerpoint\AutoMappscript\Folder' | Get-ChildItem | ?{$_.Fullname -like "*Worksheet*"}

        #$workpath = $ | Get-ChildItem | ?{$_.fullname -like "*Worksheet*"}

        $excel = New-Object -TypeName OfficeOpenXml.ExcelPackage $workpath

        $righttext = $excel.Workbook.Worksheets[0].HeaderFooter.OddHeader.RightAlignedText
    
        $trimmed = $righttext.Replace(" ","")


        try{

            $trimmed1 = $trimmed.Substring(0, $trimmed.IndexOf("p."))
        }
        catch{

            $trimmed1 = $trimmed.Substring(0, $trimmed.IndexOf("Page:"))

        }


        $draft = ($trimmed1.Substring($trimmed1.IndexOf("Effective:")).Replace("Effective:","").Replace(" ","")).trim()

        $excel.Dispose()

        if($draft -eq "draft"){
        
            Write-Host "Revision" "$($docname)_$($docrev)" "is currently a draft, archiving to draft folder" -ForegroundColor DarkYellow

            "Revision $($docname)_$($docrev) is currently a draft. Archiving to draft folder, $(((Get-Date -Format "yyyy-MM-dd HH:mm:ss").ToString()))" | Add-Content -Path "\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\Skiftspecifika dokument\Skift 1\Mahdi\powerpoint\AutoMappscript\Log\Log.csv"

            $draftpath = New-Item -Path "\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\Skiftspecifika dokument\Skift 1\Mahdi\powerpoint\AutoMappscript\draft\$($docname)\$($docrev)" -ItemType Directory

            $date | Out-File -FilePath "\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\Skiftspecifika dokument\Skift 1\Mahdi\powerpoint\AutoMappscript\draft\$($docname)\$($docrev)\$($docrev).txt"

            $folderpath | Get-ChildItem | ForEach-Object{Copy-Item -Path $_.Fullname -Destination $draftpath}

            $updaterevision = $false
        
        }else{

            "Revision $($docname)_$($docrev) is not a draft. Replacing the old revision in the respective product folder and archiving the new revision to Archive, $(((Get-Date -Format "yyyy-MM-dd HH:mm:ss").ToString()))" | Add-Content -Path "\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\Skiftspecifika dokument\Skift 1\Mahdi\powerpoint\AutoMappscript\Log\Log.csv"
        
            $matpath | Get-ChildItem | ForEach-Object{Remove-Item $_.fullname -Force}

            $folderpath | Get-ChildItem | ForEach-Object{Copy-Item -Path $_.Fullname -Destination $matpath}

            Document-Archive -DocNrPath $folderpath -DocNrName $docname -DocRev $docrev -agile $agile
            
            $updaterevision = $true                    
        
        }

        Disconnect-PnPOnline
    }


    return $updaterevision

    write-host "Done"
}

#///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

Set-Location "\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\Skiftspecifika dokument\Skift 1\Mahdi\powerpoint\AutoMappscript"

cls


$status = Get-Content -Path "\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\Skiftspecifika dokument\Skift 1\Mahdi\powerpoint\AutoMappscript\status.txt"


if($status -eq "Disabled"){Write-Host "Automappscript is disabled" -ForegroundColor DarkYellow}else{

    $version = "Automatic Mappscript v5.6.4:20250505 "

    $env:PNPPOWERSHELL_UPDATECHECK = "Off"

    write-host "..."

    Write-Host $version

    Write-host "..."

    Start-sleep 1

    #Load necessary assembly for word processing

    try{
        Write-host "Loading OpenXML Assembly" -ForegroundColor DarkYellow
        $base64 = Get-Content -Path "\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\Skiftspecifika dokument\Skift 1\Mahdi\powerpoint\AutoMappscript\Modules\OpenXMLAssembly.txt"
        $bits = [Convert]::FromBase64String($base64)
        [System.Reflection.Assembly]::Load($bits) | Out-Null

    }catch{
    Write-Host 'Could not load essential assembly Documentfortmat.OpenXML for .xlsx interaction' -BackgroundColor DarkRed
    "Could not load essential assembly Documentfortmat.OpenXML for .xlsx interaction,$((($_) -replace ",",";")),$($_.InvocationInfo.ScriptLineNumber),$(((Get-Date -Format "yyyy-MM-dd HH:mm:ss").ToString()))" | Add-Content -Path "\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\Skiftspecifika dokument\Skift 1\Mahdi\powerpoint\AutoMappscript\Log\Errorlog.csv"
    
    
    }

    #öppna excel, uppdatera informationen och skapa globala variabler


    try{
        #Import-Module PnP.PowerShell
        Import-Module -Name PnP.PowerShell -ErrorAction Stop
        $sharepoint = $true
        $pnp = $true
    }
    catch{
        Write-Host 'Could not load necessary module PnP Powershell for Sharepoint interaction. Module could be uninstalled from local modules path. Retrying with local path..' -ForegroundColor DarkYellow
        try{Import-Module -Name '.\Modules\PnP.PowerShell' -ErrorAction Stop; $sharepoint = $true}
        catch{
        Write-Host 'Could not load essential module PnP Powershell for sharepoint interaction. Sharepoint is disabled' -BackgroundColor DarkRed; $sharepoint = $false; $pnp = $false; $script:ExitCode = 1
        "Could not load essential module PnP Powershell for sharepoint interaction. AutoMappscript is disabled,$((($_) -replace ",",";")),$($_.InvocationInfo.ScriptLineNumber),$(((Get-Date -Format "yyyy-MM-dd HH:mm:ss").ToString()))" | Add-Content -Path "\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\Skiftspecifika dokument\Skift 1\Mahdi\powerpoint\AutoMappscript\Log\Errorlog.csv"
        }
    }
    #finally{Write-Host 'Could not load PnP Powershell. Sharepoint feature is disabled.' -BackgroundColor Red; $sharepoint = $false}

    #Set-ExecutionPolicy -Scope "CurrentUser" -ExecutionPolicy "RemoteSigned"

    #install-module pswriteoffice -scope CurrentUser

    #Import-Module PSWriteOffice

    #install-module pswriteword -scope CurrentUser
    try{
        import-module pswriteword -ErrorAction Stop}
    catch{
    
        Write-Host "Could not load PSWriteWord for signature list creation, part of folder creation function will not work." -BackgroundColor DarkRed
        "Could not load PSWriteWord for signature list creation; part of folder creation function will not work,$((($_) -replace ",",";")),$($_.InvocationInfo.ScriptLineNumber),$(((Get-Date -Format "yyyy-MM-dd HH:mm:ss").ToString()))" | Add-Content -Path "\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\Skiftspecifika dokument\Skift 1\Mahdi\powerpoint\AutoMappscript\Log\Errorlog.csv"

        }

    #Install-Module pnp.powershell -MaximumVersion 1.12.0 -Scope CurrentUser

    try{
   
        write-host "Loading EPPlus assembly" -ForegroundColor DarkYellow      
        $epplus = @('\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\Skiftspecifika dokument\Skift 1\Mahdi\powerpoint\AutoMappscript\Modules\EPPlus\EPPlus.6.2.7\lib\net35\EPPlus.dll', '\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\Skiftspecifika dokument\Skift 1\Mahdi\powerpoint\AutoMappscript\Modules\EPPlus\EPPlus.Interfaces.6.1.1\lib\net35\EPPlus.Interfaces.dll',
        '\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\Skiftspecifika dokument\Skift 1\Mahdi\powerpoint\AutoMappscript\Modules\EPPlus\EPPlus.System.Drawing.6.1.1\lib\net35\EPPlus.System.Drawing.dll')

        foreach($assembly in $epplus){

            $epplus64 = [System.Convert]::ToBase64String([System.IO.File]::ReadAllBytes($assembly))

            $bits = [Convert]::FromBase64String($epplus64)

            [System.Reflection.Assembly]::Load($bits) | Out-Null
        }

        $epplus = $true
    }catch{

        Write-Host "Could not load necessary libraries EPPlus 6.2.7, EPPlus Interfaces 6.1.1 and EPPlus System Drawing 6.1.1 required for folder creation. Check if its in in the Modules folder and restart the script. Exiting script..." -BackgroundColor DarkRed
        "Could not load necessary libraries EPPlus 6.2.7; EPPlus Interfaces 6.1.1 and EPPlus System Drawing 6.1.1 required for folder creation, Error Loading EPPlus, 3599,$(((Get-Date -Format "yyyy-MM-dd HH:mm:ss").ToString()))" | Add-Content -Path "\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\Skiftspecifika dokument\Skift 1\Mahdi\powerpoint\AutoMappscript\Log\Errorlog.csv"
        $epplus = $false
        $script:ExitCode = 1

    }

    if(($epplus) -and ($pnp)){

        #main

        If(-not (Test-Path .\Log)){
    
            Write-Host 'Log folder does not exist, creating a new one' -ForegroundColor DarkYellow

            $path = ((New-Item -Name 'Log' -ItemType 'Directory').FullName + '\')

            'Errorlog.csv','Log.csv', 'Folderlog.csv' | %{if(-not (Test-Path -Path ($path + $_))){(New-Item -Path $path -Name $_ -ItemType File) | Out-Null}}

            'Error Message,Error Exception,Line,Date' | Add-Content -Path ($path + 'Errorlog.csv'); '' | Add-Content -Path ($path + 'Errorlog.csv')
            'Date Created,ROBAL,LSP,Material,Batch(s),Order,Sample Reagent,Order Amount,Production time,Revision' | Add-Content -Path ($path + 'Folderlog.csv'); '' | Add-Content -Path ($path + 'folderlog.csv')
            'General Message,Date' | Add-Content -Path ($path + 'Log.csv'); '' | Add-Content -Path ($path + 'log.csv')
        }


        $agile = if(-not $sharepoint){Write-Host 'Sharepoint is disabled' -BackgroundColor DarkRed; ("Disabled")}else{("Sharepoint")} #if((Read-Host '1. Agile | 2. Sharepoint') -eq 1){"Agile"}else{"Sharepoint"}}

        #if($agile -eq "Disabled"){if((Read-Host 'Type 1 to enable Agile (Power Automate script needs to be running) else press Enter') -eq 1){$agile = "Agile"}else{$agile = "Disabled"}}

        if($agile -eq "Agile"){

            Start-Job -Name "MailCheck" -ScriptBlock {
        
                    $path = '\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\Skiftspecifika dokument\Skift 1\Mahdi\powerpoint\AutoMappscript\PAFiles\*'

                            while(($count = ($path | Get-ChildItem).count) -ge 0){

                    $count

                }

                    #start-sleep 3
                } | Out-Null
        }

        #menu -version $version -agile $agile

        $main = $True

        $runResult = main -objexcel $objexcel -wb $wb -sheet $sheet -revsheet $revsheet -main $main -agile $agile -version $version -BatchId $BatchId -JobId $JobId

        if(($runResult -is [hashtable]) -and ($runResult.ContainsKey('Success'))){
            if(-not $runResult.Success){$script:ExitCode = 1}
        }elseif($runResult -is [int]){
            $script:ExitCode = $runResult
        }


        Start-Sleep 0.5

        #[System.Runtime.Interopservices.Marshal]::ReleaseComObject($objexcel) | Out-Null

    }else{
        if(-not $script:ExitCode){$script:ExitCode = 1}
    }
}

[Environment]::Exit($script:ExitCode)
#v5.6.4:20250505
