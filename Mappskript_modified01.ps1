function get-matvariables ($material) {

    [hashtable]$return = @{}

    switch ($material){
        
        #SARS
        {($_ -eq 'XPRSARS-COV2-10') -or ($_ -eq 'D39525')}{
            $ASSAY = "XPRSARS COV2"
            $highpos = $false
            $matrev = 'D39525'
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
            $product = “XP3SARSCOV210”
            $sealassay = "Xpress CoV-2 plus"
            $partnr = "700-7425"
            $ultra = $false
            $path = '\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\1. IPT - KOMMANDE TESTER\Mappscript\XP3 SARS COV2'
            $na = "N/A"
        }
        {($_ -eq 'XPRS-COV2-10') -or ($_ -eq 'D48538')}{
            $ASSAY = "XPRS COV2"
            $highpos = $false
            $matrev = 'D48538'
            $product = “XPRSCOV210”
            $sealassay = "Xpress CoV-2 plus"
            $partnr = "700-8085"
            $ultra = $false
            $path = '\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\1. IPT - KOMMANDE TESTER\Mappscript\XP3 SARS COV2'
            $na = "N/A"
        }
        {($_ -eq 'XPRS-COV2-CE-10') -or ($_ -eq 'D48538')}{
            $ASSAY = "XPRS COV2"
            $highpos = $false
            $matrev = 'D48538'
            $product = “XPRSCOV2CE10”
            $sealassay = "Xpress CoV-2 plus"
            $partnr = "700-8086"
            $ultra = $false
            $path = '\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\1. IPT - KOMMANDE TESTER\Mappscript\XP3 SARS COV2'
            $na = "N/A"
        }
        {($_ -eq 'XPCOV2/FLU/RSV-10 deaad') -or ($_ -eq 'D41929')}{
            $ASSAY = "XPCOV2 FLU RSV"
            $highpos = $false
            $matrev = 'D41929'
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
            $product = “XP3COV2FLURSV10”
            $sealassay = "Xpress CoV-2/Flu/RSV plus"
            $partnr = "700-7493"
            $ultra = $false
            $path = '\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\1. IPT - KOMMANDE TESTER\Mappscript\XP3 SARS-COV2 FLU RSV plus'
            $na = "N/A"
        }
        {($_ -eq 'XPRS4PLEX-10') -or ($_ -eq 'D47377')}{
            $ASSAY = "XPRS COV2 PLEX"
            $highpos = $false
            $matrev = 'D47377'
            $product = “XPRS4PLEX10”
            $sealassay = "Xpress CoV-2/Flu/RSV plus"
            $partnr = "700-7906"
            $ultra = $false
            $path = '\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\1. IPT - KOMMANDE TESTER\Mappscript\XP3 SARS-COV2 FLU RSV plus'
            $na = "N/A"
        }
        {($_ -eq 'XPRS4PLEX-CE-10') -or ($_ -eq 'D47377')}{
            $ASSAY = "XPRS COV2 PLEX"
            $highpos = $false
            $matrev = 'D47377'
            $product = “XPRS4PLEXCE10”
            $sealassay = "Xpress CoV-2/Flu/RSV plus"
            $partnr = "700-7903"
            $ultra = $false
            $path = '\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\1. IPT - KOMMANDE TESTER\Mappscript\XP3 SARS-COV2 FLU RSV plus'
            $na = "N/A"
        }
        #MTB
        {($_ -eq 'GXMTB/RIF-ULTRA-50') -or ($_ -eq 'D25862')}{
            $ASSAY = “GXMTB RIF ULTRA”
            $ultra = $true
            $highpos = $true
            $matrev = 'D25862'
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
            $matrev = 'D25862'
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
            $matrev = 'D25862'
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
            $matrev = 'D25862'
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
            $matrev = 'D31503'
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
            $matrev = 'D31503'
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
            $matrev = 'D31503'
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
            $matrev = 'D31503'
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
            $matrev = 'D31503'
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
            $matrev = 'D31503'
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
            $matrev = 'D31503'
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
            $matrev = 'D31503'
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
            $matrev = 'D16904'
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
            $matrev = 'D16904'
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
            $matrev = 'D16904'
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
            $matrev = 'D16904'
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
            $matrev = 'D16904'
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
            $matrev = 'D51110'
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
            $matrev = 'D51110'
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
            $matrev = 'D51110'
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
            $matrev = 'D51111'
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
            $matrev = 'D79377'
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
            $matrev = "D61353"
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
            $matrev = 'D36985'
            $product = “700-6793/ HIV-1 QUAL XC, CE-IVD”
            $ASSAYFAM = "HIV"
            $sealassay = "HIV"
            $partnr = "700-6793"
            $ultra = $false
            $path = '\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\1. IPT - KOMMANDE TESTER\Mappscript\HIV QA XC'
            $na = "NA"
        }
        {($_ -eq "GXHIV-QA-XC-CN-10") -or ($_ -eq 'D83056')}{
            $ASSAY = “HIV QA XC CN”
            $highpos = $false
            $matrev = 'D83056'
            $product = “700-9658/ HIV-1 QUAL XC, CN ”
            $ASSAYFAM = "HIV"
            $sealassay = "HIV"
            $partnr = "700-9658"
            $ultra = $false
            $path = '\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\1. IPT - KOMMANDE TESTER\Mappscript\HIV QA XC CN'
            $na = "NA"
        }
        #OTHER QUANT ASSAYS
        {($_ -eq "GXHBV-VL-CE-10") -or ($_ -eq 'D51114')}{
            $ASSAY = “HBV VL”
            $highpos = $true
            $matrev = 'D51114'
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
            $matrev = 'D55782'
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
        {($_ -eq "GXNOV-GB-10") -or ($_ -eq 'D17716')}{
            $ASSAY = “NORO”
            $highpos = $true
            $matrev = 'D17716'
            $product = “GXNOV-GB-10”
            $sealassay = "Norovirus"
            $partnr = "700-9547"
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


        #GI PANEL

        {($_ -eq "GXGI-10") -or ($_ -eq 'D55989')}{
            $ASSAY = “GI PANEL”
            $highpos = $false
            $matrev = 'D55989'
            $product = “GXGI-10”
            $sealassay = "Xpert GI Panel"
            $partnr = "700-9454"
            $ultra = $false
            $doublesampling = $true
            $path = '\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\1. IPT - KOMMANDE TESTER\Mappscript\GI PANEL'
            $na = "NA"
        
        }

        {($_ -eq "GXGI-CE-10") -or ($_ -eq 'D55989')}{
            $ASSAY = “GI PANEL”
            $highpos = $false
            $matrev = 'D55989'
            $product = “GXGI-CE-10”
            $sealassay = "Xpert GI Panel"
            $partnr = "700-9467"
            $ultra = $true
            $doublesampling = $true
            $path = '\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\1. IPT - KOMMANDE TESTER\Mappscript\GI PANEL'
            $na = "NA"
        
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

    return $return

}

function GetMergedCellRange ($Sheet, $Cell){

    #This checks if a cell is merged and returns the range eg. H3:J3. If the cell is not a merged cell then it just returns the address eg. H3.

    if($Cell.Merge){

        $idx = $Sheet.GetMergeCellId($Cell.Start.Row, $Cell.Start.Column)

        $Range = $Sheet.MergedCells[$idx-1]

        return $Range

    }else{

        return $Cell.Address
        

    }

}
#////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


function create-folder ($returnmat, $robalnr, $ordernr, $batchnr, $lsp, $samplereagent, $prodtime, $version, $robalitem) {

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
    
    $LINA = “ROBAL “ + “$robal”

    [string]$sealtestrobal = if(($robal.ToString().length) -gt 1){"ROBAL-0"+$robal}else{"ROBAL-00"+$robal}


    $BulkOBR = @(

                  "700-5208"
                  "700-5288"
                  "700-6609"
                  "700-9287"
                 
                 )


    if(($robalitem.kundsr -notin $BulkOBR) -and ($robalitem.kundsr -match '\d')){

        $MAPPNAMN = “R” + “$robal” + “ - “  + “$assay #” + “$lsp” + “ “ + “- PQC” + ” (” + ”$orderamount” + ”)” + "  OBR"

    }else{

        $MAPPNAMN = “R” + “$robal” + “ - “  + “$assay #” + “$lsp” + “ “ + “- PQC” + ” (” + ”$orderamount” + ”)”

    }



    New-item -path "\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\1. IPT - KOMMANDE TESTER\$LINA\$MAPPNAMN" -ItemType Directory

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

    $mailtemplatpath = "\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\1. IPT - KOMMANDE TESTER\$LINA\$MAPPNAMN\$assay #$LSP Mailtemplat.xlsm”

    sealtestdatapos -excel $excel -DATUM $DATUM -ultra $ultra -highpos $highpos -equipment $equipment -batchnr $batchnr -LSP $LSP -POnr $POnr -sealtestrobal $sealtestrobal -sealtestpospath $sealtestpospath -partnr $partnr -sealassay $sealassay -carba $carba -hpv $hpv -returnmat $returnmat
   
    sealtestdataneg -excel $excel -DATUM $DATUM -equipment $equipment -batchnr $batchnr -LSP $LSP -POnr $POnr -sealtestrobal $sealtestrobal -sealtestnegpath $sealtestnegpath -partnr $partnr -sealassay $sealassay -carba $carba -hpv $hpv -ultra $ultra -returnmat $returnmat

    worksheet -excel $excel -worksheetpath $worksheetpath -product $product -DATUM $DATUM -LSP $LSP -batchnr $batchnr -partnr $partnr -ultra $ultra -version $version -material $returnmat
       
    signlist -batchnr $batchnr -POnr $POnr -LSP $LSP -LINA $LINA -MAPPNAMN $MAPPNAMN -assay $assay -document $document -partnr $partnr

    Set-Location "\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\Skiftspecifika dokument\Skift 1\Mahdi\powerpoint\AutoMappscript"

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

    $excel = New-Object -TypeName OfficeOpenXml.ExcelPackage $sealtestpospath

    $book = $excel.Workbook

    $checkbox = $book.Worksheets.Drawings | ?{$_.As.Control.CheckBox}

    $checkbox | %{$_.Checked = 1}

    $instrumentlist = @("Balance ID Number", "Vacuum Oven ID Number", "Timer ID Number", "OC Mold(s)", "Part Number", "Cartridge Number (LSP)", "PO Number")
    

    foreach($sheet in $book.Worksheets){


        if($sheet.index -eq 0){
            
            continue

        }

        elseif($sheet.Name -eq "Datasheet (1)"){


        foreach($instrument in $instrumentlist){

            $cellvalue = $sheet.Cells| ?{$_.Text -like "*$instrument*"}

            $celladdress = $cellvalue.Address

            $cellrow, $cellcolumn = $sheet.Cells[$celladdress].start.Row, $sheet.Cells[$celladdress].start.column

            switch($instrument){

                "Part Number"{

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

            foreach($instrument in $instrumentlist){

                $cellvalue = $sheet.Cells| ?{$_.Text -like "*$instrument*"}

                $celladdress = $cellvalue.Address

                $cellrow, $cellcolumn = $sheet.Cells[$celladdress].start.Row, $sheet.Cells[$celladdress].start.column

                switch($instrument){

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

                $SECONDBAG = If ( (2 * $count) -lt 10) { "0" + (2 * $count).ToString() } Else { 2 * $count }

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

                $SECONDBAG = If ( (2 * $count) -lt 10) { "0" + (2 * $count).ToString() } Else { 2 * $count }

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

        if($sheet.Index -eq 0){continue}

        $cellvalue = $sheet.Cells| ?{$_.Text -like "*Cartridge ID*"}

        $celladdress = $cellvalue.Address

        $cellrow, $cellcolumn = $sheet.Cells[$celladdress].start.Row, $sheet.Cells[$celladdress].start.column

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


            while($true){

                if(($sheet.cells[($cellrow + 3), ($cellcolumn + 2)].Text) -ne ''){break}
        
                $sheet.cells[($cellrow + 3), ($cellcolumn + 2)].value = “N/A"
                $sheet.cells[($cellrow + 3), ($cellcolumn + 3)].value = “N/A"
                $sheet.cells[($cellrow + 3), ($cellcolumn + 4)].value = “N/A"

                $cellrow = $cellrow - 2

            }        
      



    }

    $excel.Save()
    $excel.Dispose()

    }

function Sealtestdataneg($excel, $DATUM, $carba, $hpv, $equipment, $batchnr, $LSP, $POnr, $sealtestrobal, $sealtestnegpath, $partnr, $sealassay, $ultra, $returnmat){


    $excel = New-Object -TypeName OfficeOpenXml.ExcelPackage $sealtestnegpath

    $book = $excel.Workbook

    $checkbox = $book.Worksheets.Drawings | ?{$_.As.Control.CheckBox}

    $checkbox | %{$_.Checked = 1}

    $instrumentlist = @("Balance ID Number", "Vacuum Oven ID Number", "Timer ID Number", "OC Mold(s)", "Part Number", "Cartridge Number (LSP)", "PO Number")

    foreach($sheet in $book.Worksheets){


        if($sheet.index -eq 0){
            
            continue

        }

        elseif($sheet.Name -eq "Datasheet (1)"){

        foreach($instrument in $instrumentlist){

        
            $cellvalue = $sheet.Cells| ?{$_.Text -like "*$instrument*"}

            $celladdress = $cellvalue.Address

            $cellrow, $cellcolumn = $sheet.Cells[$celladdress].start.Row, $sheet.Cells[$celladdress].start.column

            switch($instrument){

                "Part Number"{

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

            foreach($instrument in $instrumentlist){

                $cellvalue = $sheet.Cells| ?{$_.Text -like "*$instrument*"}

                $celladdress = $cellvalue.Address

                $cellrow, $cellcolumn = $sheet.Cells[$celladdress].start.Row, $sheet.Cells[$celladdress].start.column

                switch($instrument){

                    "Balance ID Number"{


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

        if($sheet.Index -eq 0){continue}

        $cellvalue = $sheet.Cells| ?{$_.Text -like "*Cartridge ID*"}

        $celladdress = $cellvalue.Address

        $cellrow, $cellcolumn = $sheet.Cells[$celladdress].start.Row, $sheet.Cells[$celladdress].start.column

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

            while($true){

                if(($sheet.cells[($cellrow + 3), ($cellcolumn + 2)].Text) -ne ''){break}
        
                $sheet.cells[($cellrow + 3), ($cellcolumn + 2)].value = “N/A"
                $sheet.cells[($cellrow + 3), ($cellcolumn + 3)].value = “N/A"
                $sheet.cells[($cellrow + 3), ($cellcolumn + 4)].value = “N/A"

                $cellrow = $cellrow - 2

            }        

    }

    $excel.Save()
    $excel.Dispose()

}

function worksheet($excel, $worksheetpath, $product, $DATUM, $LSP, $batchnr, $partnr, $ultra, $version, $material){

        $excel = New-Object -TypeName OfficeOpenXml.ExcelPackage $worksheetpath

        $book = $excel.Workbook

        $book.Worksheets | %{
            # Allow editing on all sheets except the data sheet
            if($_.Name -notlike "*QC Data*"){$_.Protection.IsProtected = $false}

            # Preserve any existing image in the odd header.  EPPlus stores header pictures
            # in the Pictures collection on the header.  If a picture exists we need to
            # include the image placeholder (&G) at the beginning of the header text so
            # the logo remains visible【906359467953244†L113-L121】.
            $imgPrefix = ""
            try {
                $pictures = $_.HeaderFooter.OddHeader.Pictures
                if ($null -ne $pictures -and $pictures.Count -gt 0) {
                    $imgPrefix = [OfficeOpenXml.Drawing.HeaderFooter.ExcelHeaderFooter]::Image + "`n"
                }
            } catch {
                # Older versions of EPPlus expose the picture via LeftHeaderPicture
                if ($_.HeaderFooter.LeftHeaderPicture -and $_.HeaderFooter.LeftHeaderPicture.Filename -ne "") {
                    $imgPrefix = [OfficeOpenXml.Drawing.HeaderFooter.ExcelHeaderFooter]::Image + "`n"
                }
            }

            # Compose the left header text.  Use backtick-n (`n) for newlines so that PowerShell
            # correctly inserts a line break inside the string.  The image code will be empty
            # when there is no picture in the header.
            $headerText = $imgPrefix + "Part No.: $partnr`nBatch No(s).: $batchnr`nCartridge No. (LSP): $LSP"
            $_.HeaderFooter.OddHeader.LeftAlignedText = $headerText
        }

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


        $namelist = @("Product Catalog/Part No.", "Cartridge No. (LSP)", "Recorded By:")
        

        #This code accounts for the current cell being merged or not. Otherwise the cell next (column) to the first cell in the merge collection will have wrong info in it. 
        $TypeOfResampleCell = $sheet.Cells| ?{$_.Text -like "*Type Of Resample*"} 
        $Range = GetMergedCellRange -Sheet $sheet -Cell $TypeOfResampleCell
        $TypeOfResampleCellEndRow, $TypeOfResampleCellEndColumn = $sheet.Cells[$Range].End.Row, $sheet.Cells[$Range].End.Column

        
        #Enters NA or N/A to the next column depending on product worksheet
        $sheet.Cells[$TypeOfResampleCellEndRow, ($TypeOfResampleCellEndColumn + 1)].Value = $material.na


        foreach($name in $namelist){

                $FoundCell = $sheet.Cells| ?{$_.Text -like "*$name*"} 
                $Range = GetMergedCellRange -Sheet $sheet -Cell $FoundCell
                $cellrow, $cellcolumn = $sheet.Cells[$Range].End.Row, $sheet.Cells[$Range].End.Column

                #Enters the relevant info on the next column

            switch($name){

                "Product Catalog/Part No."{

                    $sheet.cells[$cellrow, ($cellcolumn + 1) ].value = $product
                }

                "Cartridge No. (LSP)"{

                    $sheet.cells[$cellrow, ($cellcolumn + 1) ].value = [int]$LSP

                }

                "Recorded By:"{

                    #Finds "Recorded By:" cell and adds Automappscript version in the next column and keeps going to the next cell until it finds the cell with "Date:" value or a cell with specific green color.
                    #"Recorded By:" cell are in the same row as "Date:" cell.
                    #The reason why we loop until we find Date is because some product worksheets dont have a cell with "Date:" value. 


                    #Enters the info to the next cell column
                    $sheet.cells[$cellrow, ($cellcolumn + 1)].value = $version + "Ge feedback :)"


                    #Sets the cellrow and cellcolumn variables as the last cell in the merged cells.
                    #Otherwise the code later on that checks for a green cell will be triggered in the wrong cell.
                    $Range = GetMergedCellRange -Sheet $sheet -Cell ($sheet.Cells[$cellrow, ($cellcolumn + 1)])
                    $cellrow, $cellcolumn = $sheet.Cells[$Range].End.Row, ($sheet.Cells[$Range].End.Column + 1)

                    while($true){
                    
                        #If the current cell has the date text, move right and enter Date value

                        if($sheet.Cells[$cellrow,$cellcolumn].Text -like "*Date*"){

                            $cellcolumn++
                            $sheet.cells[$cellrow,$cellcolumn].value = $DATUM

                            break

                         #If the current cell has a green background color, enter Date value. The indexed and RGB value corresponds to the specific green color used by worksheets

                        }elseif($sheet.Cells[$cellrow, $cellcolumn].Style.Fill.BackgroundColor.Indexed -eq 42 -or $sheet.Cells[$cellrow, $cellcolumn].Style.Fill.BackgroundColor.Rgb -eq 'FFCCFFCC' ){

                            $sheet.cells[$cellrow,$cellcolumn].value = $DATUM

                            break                        
                       
                        }else{$cellcolumn++}
                    
                    }
                }
            }
        }

        #Per Intended Use, information can be entered either manually or electronically on cells with a specific green background color. Blue must be entered electronically.
        #The code below uses the RBG or Indexed value of that specified green/blue background to find cells that needs to be entered. The cells that require special info will not be affected by this
        #as the code also checks if the cell has no text value since the special cells will be handled first.
        # Green has RBG of FFCCFFCC or Index 42
        # Blue has RGB of FFCCFFFF or Index 41
                                                                                    #Green                                             #Blue                                                  #Green                                         #Blue
        $TestSummaryInputCells = $sheet.Cells | ?{($_.Style.Fill.BackgroundColor.Indexed -eq 42) -or ($_.Style.Fill.BackgroundColor.Indexed -eq 41) -or ($_.Style.Fill.BackgroundColor.Rgb -eq 'FFCCFFCC') -or ($_.Style.Fill.BackgroundColor.Rgb -eq 'FFCCFFFF')}

        $TestSummaryInputCells | %{if($_.Text -eq ""){$_.Value = "N/A"}}


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

        $celladdress = $cellvalue.Address

        $cellrow, $cellcolumn = $sheet.Cells[$celladdress].End.Row, $sheet.Cells[$celladdress].End.column

        $bagrow = $Cellrow + 1

        $i = 1

        while($true){

            $cellvalue = $sheet.Cells[$bagrow, $CellColumn].text

            $lastcell = $sheet.Cells[$bagrow, ($CellColumn - 1)].text

            if($lastcell.Trim() -eq 250){


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

        $namelist = @("Comments:", "Date:")

        foreach($name in $namelist){

            $namecell = $sheet.Cells| ?{$_.Text -eq $name}

            $celladdress = $namecell.Address

            $namerow, $namecolumn = $sheet.Cells[$celladdress].End.Row, $sheet.Cells[$celladdress].End.column

            if($name -eq "Comments:"){

                $namerow = $namerow + 1

                $sheet.cells[$namerow, $namecolumn].Value = 'N/A'

            }elseif($name -eq "Date:"){

                $namecolumn = $namecolumn + 1

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

        $cells = $sheet.Cells["A1:H22"]

        $cellrange = $cells | ?{$_.text -ne "" -and $_.text -ne "N/A"}

        $cellvalue = $cellrange | ?{$_.text -like "*Comment*"}

        $nacells = @()

        foreach($val in $validation){
        
            
            $list = $val.As.ListValidation

            if($list -ne $null){

                $rangearray = ($list.Address.Address).Split(",")

                $rangearray | %{
                
                    $nacells += $sheet.cells[$_] | ?{$_.Style.Border.Bottom.Style -eq "Thin"}
                
                }
            
            }
        
        }

        $nacells | %{$sheet.Cells[$_.Localaddress].value = "N/A"}


        $cells = $sheet.Cells["A1:H370"]

        $cellrange = $cells | ?{$_.text -ne "" -and $_.text -ne "N/A"}
        
        $date = $cellrange | ?{$_.Text -eq "Date:"}
        
        $celladdress = $date.Address

        $namerow, $namecolumn = $sheet.Cells[$celladdress].End.Row, $sheet.Cells[$celladdress].End.column

        $namecolumn = $namecolumn + 1

        $sheet.Cells[$namerow, $namecolumn].Value = $DATUM       
       
    }

function signlist($partnr, $batchnr, $POnr, $LSP, $document, $LINA, $MAPPNAMN, $assay){

        $path = “\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\1. IPT - KOMMANDE TESTER\Mappscript\”

        $path | Get-ChildItem | ForEach-Object{if($_ -like "*Signature*"){$path = $_.Fullname; Copy-Item -Path $_.fullname -Destination “\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\1. IPT - KOMMANDE TESTER\$LINA\$MAPPNAMN\”}}
        
        [string]$signpath = (“\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\1. IPT - KOMMANDE TESTER\$LINA\$MAPPNAMN\” | Get-ChildItem | Where-Object{$_ -like "*Signature*"}).FullName

        $doc = Get-WordDocument -FilePath $signpath

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
    


    if($revrefresh -eq 1){

        write-host 'Hämtar alla revisioner i mapparna....'

        #få fram alla revisioner från alla produktmappar

        $revlist = @{}
        $indivmaterialfolder = @()
        $folder = "\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\1. IPT - KOMMANDE TESTER\Mappscript"
        #får fram alla mappar i huvudmappen
        $materialfolders = Get-ChildItem $folder | Where-Object {($_.PSisContainer) -and ($_ -notlike "*mappscript*")`
        -and ($_ -notlike "*GBS*")}
        #ser till att mappen "mappscript" inte är inkluderad, Fluvid+ (korrupt worksheet)-
        #-om den skapas från denna mapp för någon anledning och GBS (Ej implementerad).


        foreach ($materialfolder in $materialfolders){


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
            }
        }

    }

    return $rowarray

    Write-host " "
}

function Document-Archive($DocNrPath,$DocNrName, $DocRev, $agile){

    $DocNrName = $DocNrName.Substring(0, 6)

    $docfoldername = $DocNrName + " Rev " + $DocRev

    $archivepath = '\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\Skiftspecifika dokument\Skift 1\Mahdi\powerpoint\AutoMappscript\Archive\'

    $docpathcheck = $archivepath + $DocNrName
    
    $docpath = $archivepath + $DocNrName + "\"

    $DocRevPath = $docpath + $docfoldername


    if(Test-Path -Path $docpathcheck){

        if(-not (Test-Path -Path $DocRevPath)){

            New-Item -Path $docpath -Name $docfoldername -ItemType "Directory" | Out-Null

            start-sleep 0.5

            if($agile -eq "Agile"){

            Expand-Archive -Path $DocNrPath -DestinationPath $DocRevPath

            }else{
            
                $DocNrPath | Get-ChildItem | ForEach-Object{Copy-Item $_.fullname -Destination $DocRevPath}
                
            
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

        }
    }

    

}

function refresh($objexcel, $wb, $sheet, $revsheet, $agile){

    Write-Host "Hämtar körningar"

    $clientid = "23715695-a9a6-4f32-af7b-4cd164e0f1f9"

    $tenant = "danaher.onmicrosoft.com"

    $certificateprivatekey = "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"

    Connect-PnPOnline -Url "https://danaher.sharepoint.com/sites/CEP-Sweden-Production-Management" -Tenant $tenant -ClientId $clientid -CertificateBase64Encoded $certificateprivatekey

    $list = Get-PnPList -Identity "Cepheid | Production orders"

    $robalitems = Get-PnPListItem -List $list -PageSize 500

    $robal = $robalitems | Where{$_.FieldValues.Work_x0020_Center -like "*ROBAL*" -and $_.FieldValues.Actual_x0020_startdate_x002f__x0 -eq $null `
     -and ($_.FieldValues.Title -match '^\d+$') -and $_.FieldValues.Production_x0020_Day_x002f__x002 -ge ((Get-Date).ToShortDateString())}

    $rowarray = @()

    foreach($item in $robal){
    
        $item = $item.FieldValues
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
    'D83056'
    'D55989'
    )


    Connect-PnPOnline -Url 'https://danaher.sharepoint.com/sites/CEP_CWS' -Tenant $tenant -ClientId $clientid -CertificateBase64Encoded $certificateprivatekey

    Write-Host "Hämtar revisioner"

    $items = Get-PnPListItem -List "Cepheid" -PageSize 4000
    $items = $items | ?{$_.FileSystemObjectType -eq "Folder"}

    $sharepointarray = @()

    foreach($document in $reference){

        $sorthashtable = @{}

        $docfolders = $items | ?{$_.FieldValues.FileLeafRef -like "*$document*"}

        $docfolders | %{$sorthashtable[$_.FieldValues.FileLeafRef] = $_.FieldValues.Last_x0020_Modified}

        $LatestRev = (($sorthashtable.GetEnumerator() | Sort-Object -Property Value -Descending | Select-Object -First 1).Name).split("_")

        $date = [string]($sorthashtable.GetEnumerator() | Sort-Object -Property Value -Descending | Select-Object -First 1).Value

        $hashtable = [ordered]@{
            "documentname" = ($LatestRev[0]).Replace(" ","") 
            "documentrev" = ($LatestRev[1]).Replace(" ","")
            "date" = ($date).Replace(" ","")
        }

        $sharepointarray += $hashtable
    }

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

function main($objexcel, $wb, $sheet, $revsheet, $main, $agile, $version){

    try{

        try{

            $rowarray, $revarray = refresh -objexcel $objexcel -wb $wb -sheet $sheet -revsheet $revsheet -agile $agile

        }catch{
        
            $rowarray, $revarray = refresh -objexcel $objexcel -wb $wb -sheet $sheet -revsheet $revsheet -agile $agile


        }
        
    }catch{
        
        Write-Host "An error has occured at function refresh at line: ($($_.InvocationInfo.ScriptLineNumber)): " $_ -BackgroundColor DarkRed
                          "An error has occured at function refresh,$((($_) -replace ",",";")),$($_.InvocationInfo.ScriptLineNumber),$(((Get-Date -Format "yyyy-MM-dd HH:mm:ss").ToString()))" | Add-Content -Path "\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\Skiftspecifika dokument\Skift 1\Mahdi\powerpoint\AutoMappscript\Log\Errorlog.csv"
                          
    }

    try{

        $rowarray, $revlist = get-folderrev -rowarray $rowarray
        }catch{Write-Host "An error has occured at function get-folderrev at line: ($($_.InvocationInfo.ScriptLineNumber)): " $_ -BackgroundColor DarkRed
                          "An error has occured at function get-folderrev,$((($_) -replace ",",";")),$($_.InvocationInfo.ScriptLineNumber),$(((Get-Date -Format "yyyy-MM-dd HH:mm:ss").ToString()))" | Add-Content -Path "\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\Skiftspecifika dokument\Skift 1\Mahdi\powerpoint\AutoMappscript\Log\Errorlog.csv"}
    try{

        $revcheck = sortcheck-revision -rowarray $rowarray -revarray $revarray -agile $agile -revlist $revlist
        }catch{Write-Host "An error has occured at function sortcheck-revision at line: ($($_.InvocationInfo.ScriptLineNumber)): " $_ -BackgroundColor DarkRed
                          "An error has occured at function sortcheck-revision,$((($_) -replace ",",";")),$($_.InvocationInfo.ScriptLineNumber),$(((Get-Date -Format "yyyy-MM-dd HH:mm:ss").ToString()))" | Add-Content -Path "\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\Skiftspecifika dokument\Skift 1\Mahdi\powerpoint\AutoMappscript\Log\Errorlog.csv"}

    try{

        $rowarray = running-lotcheck -rowarray $revcheck
        }catch{Write-Host "An error has occured at function running-lotcheck at line: ($($_.InvocationInfo.ScriptLineNumber)): " $_ -BackgroundColor DarkRed
                          "An error has occured at function running-lotcheck,$((($_) -replace ",",";")),$($_.InvocationInfo.ScriptLineNumber),$(((Get-Date -Format "yyyy-MM-dd HH:mm:ss").ToString()))" | Add-Content -Path "\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\Skiftspecifika dokument\Skift 1\Mahdi\powerpoint\AutoMappscript\Log\Errorlog.csv"}

    try{

        $lotsleft = totallots -rowarray $rowarray
        }catch{Write-Host "An error has occured at function lotsleft at line: ($($_.InvocationInfo.ScriptLineNumber)): " $_ -BackgroundColor DarkRed
                          "An error has occured at function lotsleft,$((($_) -replace ",",";")),$($_.InvocationInfo.ScriptLineNumber),$(((Get-Date -Format "yyyy-MM-dd HH:mm:ss").ToString()))" | Add-Content -Path "\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\Skiftspecifika dokument\Skift 1\Mahdi\powerpoint\AutoMappscript\Log\Errorlog.csv"}

    if($main -eq $True){

        if($lotsleft.count -eq 0){
            
            write-host "No lots remaining to be created, exiting"
            "No lots remaining to be created exiting,$(((Get-Date -Format "yyyy-MM-dd HH:mm:ss").ToString()))" | Add-Content -Path "\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\Skiftspecifika dokument\Skift 1\Mahdi\powerpoint\AutoMappscript\Log\Log.csv"

        }else{

            foreach($robalitem in $lotsleft){

                $returnmat =  get-matvariables $robalitem.material

                if($returnmat.assay -eq 0){
                    
                    if($robalitem.mapcreated -ne "Yes"){
                        write-host "Produkt finns ej i mappscript"

                        "Produkt finns ej i mappscript för $($robalitem.material),$(((Get-Date -Format "yyyy-MM-dd HH:mm:ss").ToString()))" | Add-Content -Path "\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\Skiftspecifika dokument\Skift 1\Mahdi\powerpoint\AutoMappscript\Log\Log.csv"

                        $robalitem.mapcreated = "Yes"

                        Continue
                    }
                }

            write-host ''
            Write-Host "ROBAL"$robalitem.robalnr 'LSP:' $robalitem.lsp 'Material:' $robalitem.material 'Batch(s):' $robalitem.batchnr 'Ordernr:' $robalitem.ordernr 'Sample Reagent:' $robalitem.samplereagent 'Order Amount:' $robalitem.orderamount 'Production time:' $robalitem.prodtime 'Rev: '-NoNewline; Write-Host $robalitem.productrev -ForegroundColor Magenta


            try{
                if($robalitem.mapcreated -ne "Yes"){
                    
                        create-folder $returnmat $robalitem.robalnr $robalitem.ordernr $robalitem.batchnr $robalitem.lsp $robalitem.samplereagent $robalitem.prodtime -version $version -robalitem $robalitem
                        "$(((Get-Date -Format "yyyy-MM-dd HH:mm:ss").ToString())),$($robalitem.robalnr),$($robalitem.lsp),$($robalitem.material),$(($robalitem.batchnr) -replace ',','/'),$($robalitem.ordernr),$(($robalitem.samplereagent) -replace ",","/"),$(($robalitem.orderamount -replace ',','')),$($robalitem.prodtime),$($robalitem.productrev)" | Add-Content -Path "\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\Skiftspecifika dokument\Skift 1\Mahdi\powerpoint\AutoMappscript\Log\Folderlog.csv"
                        $robalitem.mapcreated = "Yes"                    
                   


                }
            }catch{

                    Write-Host "An error has occured at function create-folder while trying to build folder for $($robalitem.material) at line: $($_.Exception.InvocationInfo.ScriptLineNumber): " $_ -BackgroundColor DarkRed

                    "An error has occured at function create-folder while trying to build folder for $($robalitem.material),$((($_) -replace ",",";")),$($_.InvocationInfo.ScriptLineNumber),$(((Get-Date -Format "yyyy-MM-dd HH:mm:ss").ToString()))" | Add-Content -Path "\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\Skiftspecifika dokument\Skift 1\Mahdi\powerpoint\AutoMappscript\Log\Errorlog.csv"


                }

            }
        }
    }else{
    
       $endmanual =  manualinput -rowarray $rowarray

       return $endmanual

    }
}

function sharepointupdate($docname, $docrev, $date, $agile){

    $matpath = (get-matvariables $docname).path

    $folderpath = '\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\Skiftspecifika dokument\Skift 1\Mahdi\powerpoint\AutoMappscript\Folder'
    
    $folderpath | Get-ChildItem | ForEach-Object{Remove-Item $_.fullname -Force}

    $sharepointpath = 'https://danaher.sharepoint.com/sites/CEP_CWS'

    $site = "/Cepheid/$($docname)_$($docrev)"

    $fullpath = $sharepointpath + $site + '/'

    $clientid = "23715695-a9a6-4f32-af7b-4cd164e0f1f9"

    $tenant = "danaher.onmicrosoft.com"

    $certificateprivatekey = "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"

    
    $downloadrevision = $false

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
            
        }

        Write-Host "Getting documents of $($docname)"

        Get-PnPFolderItem -FolderSiteRelativeUrl $site -Connection $connection | ForEach-Object{Get-PnPFile -Url ($site + '/' + $_.name) -Path '\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\Skiftspecifika dokument\Skift 1\Mahdi\powerpoint\AutoMappscript\Folder' -Filename $_.name -AsFile -Connection $connection}

        Start-Sleep 2

        $workpath = '\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\Skiftspecifika dokument\Skift 1\Mahdi\powerpoint\AutoMappscript\Folder' | Get-ChildItem | ?{$_.Fullname -like "*Worksheet*"}

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


    try{
        Import-Module -Name PnP.PowerShell -ErrorAction Stop
        $sharepoint = $true
        $pnp = $true
    }
    catch{
        Write-Host 'Could not load necessary module PnP Powershell for Sharepoint interaction. Module could be uninstalled from local modules path. Retrying with local path..' -ForegroundColor DarkYellow
        try{Import-Module -Name '.\Modules\PnP.PowerShell' -ErrorAction Stop; $sharepoint = $true}
        catch{
        Write-Host 'Could not load essential module PnP Powershell for sharepoint interaction. Sharepoint is disabled' -BackgroundColor DarkRed; $sharepoint = $false; $pnp = $false
        "Could not load essential module PnP Powershell for sharepoint interaction. AutoMappscript is disabled,$((($_) -replace ",",";")),$($_.InvocationInfo.ScriptLineNumber),$(((Get-Date -Format "yyyy-MM-dd HH:mm:ss").ToString()))" | Add-Content -Path "\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\Skiftspecifika dokument\Skift 1\Mahdi\powerpoint\AutoMappscript\Log\Errorlog.csv"
        }
    }

    try{
        import-module pswriteword -ErrorAction Stop}
    catch{
    
        Write-Host "Could not load PSWriteWord for signature list creation, part of folder creation function will not work." -BackgroundColor DarkRed
        "Could not load PSWriteWord for signature list creation; part of folder creation function will not work,$((($_) -replace ",",";")),$($_.InvocationInfo.ScriptLineNumber),$(((Get-Date -Format "yyyy-MM-dd HH:mm:ss").ToString()))" | Add-Content -Path "\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\Skiftspecifika dokument\Skift 1\Mahdi\powerpoint\AutoMappscript\Log\Errorlog.csv"

        }

    try{
   
        write-host "Loading EPPlus Assembly" -ForegroundColor DarkYellow      
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


        $agile = if(-not $sharepoint){Write-Host 'Sharepoint is disabled' -BackgroundColor DarkRed; ("Disabled")}else{("Sharepoint")}

        if($agile -eq "Agile"){

            Start-Job -Name "MailCheck" -ScriptBlock {
        
                    $path = '\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\Skiftspecifika dokument\Skift 1\Mahdi\powerpoint\AutoMappscript\PAFiles\*'

                            while(($count = ($path | Get-ChildItem).count) -ge 0){

                    $count

                }

                } | Out-Null
        }


        $main = $True

        main -objexcel $objexcel -wb $wb -sheet $sheet -revsheet $revsheet -main $main -agile $agile -version $version


        Start-Sleep 0.5

    }
}

#Start-Sleep 10

#stop-process -Id $PID

#v5.6.4:20250505