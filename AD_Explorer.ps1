$ErrorActionPreference = "SilentlyContinue"

function txt {
    $user = $TextBoxID.text
    $username = Get-ADUser $user -Properties Name, DisplayName -ErrorVariable erro2 | Select-Object Name, Displayname

    if ($erro2) {
        Start-Sleep -s 1
        $LabelError.Visible = $true
        Start-Sleep -s 2
        $LabelError.Visible = $false
    }
    
    else { 
        function txt_data {
            $LabelProcessing.Visible = $true ; Start-Sleep -s 2
            $file = (Get-ADUser $user -Properties MemberOf -ErrorVariable erro2).MemberOf | Get-ADGroup -Properties Name, Description | Select-Object Name, Description
            $username | Out-File "$Env:USERPROFILE\Documents\AD_$user.txt"
            $file | Add-Content "$Env:USERPROFILE\Documents\AD_$user.txt"
            $date = get-date -format "dddd, dd/MM/yyyy, HH:mm:ss"
            (Get-Content "$Env:USERPROFILE\Documents\AD_$user.txt" -Raw) -replace '@{', '' | Out-file "$Env:USERPROFILE\Documents\AD_$user.txt"
            (Get-Content "$Env:USERPROFILE\Documents\AD_$user.txt" -Raw) -replace '}', '' | Out-file "$Env:USERPROFILE\Documents\AD_$user.txt"
            $date | Add-Content "$Env:USERPROFILE\Documents\AD_$user.txt" 
            Invoke-Expression "$Env:USERPROFILE\Documents\AD_$user.txt"
            $LabelProcessing.Visible = $false
            Start-Sleep -s 1 }  
        }
        txt_data
}

function grid {

    $user = $TextBoxID.text
    $date = get-date -format "dddd, dd/MM/yyyy, HH:mm:ss"
    $file = (Get-ADUser $user -Properties MemberOf -ErrorVariable erro2).MemberOf | Get-ADGroup -Properties Name, Description | Select-Object Name, Description
    $username = (Get-ADUser $user -Properties Name, DisplayName).DisplayName

    if ($erro2) {
        $LabelError.Visible = $true
        Start-Sleep -s 2
        $LabelError.Visible = $false
    }
    
    else {
        function grid_data {
            $file | Out-gridview -Title "Grupos de Acesso de $user ($username) -> $date"
            Start-Sleep -s 1
        }
        do {Start-Export} while (grid_data)
        $LabelProcessing.Visible = $false
    }
}

function Save-CSVasExcel { #Credits to github.com/gangstanthony
    param (
        [string]$CSVFile = $(Throw 'No file provided.')
    )
    
    BEGIN {
        function Resolve-FullPath ([string]$Path) {    
            if ( -not ([System.IO.Path]::IsPathRooted($Path)) ) {
                # $Path = Join-Path (Get-Location) $Path
                $Path = "$PWD\$Path"
            }
            [IO.Path]::GetFullPath($Path)
        }

        function Release-Ref ($ref) {
            ([System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$ref) -gt 0)
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
        }
        
        $CSVFile = Resolve-FullPath $CSVFile
        $xl = New-Object -ComObject Excel.Application
    }

    PROCESS {
        $wb = $xl.workbooks.open($CSVFile)
        $xlOut = $CSVFile -replace '\.csv$', '.xlsx'
        
        # can comment out this part if you don't care to have the columns autosized
        $ws = $wb.Worksheets.Item(1)
        $range = $ws.UsedRange 
        [void]$range.EntireColumn.Autofit()

        $num = 1
        $dir = Split-Path $xlOut
        $base = $(Split-Path $xlOut -Leaf) -replace '\.xlsx$'
        $nextname = $xlOut
        while (Test-Path $nextname) {
            $nextname = Join-Path $dir $($base + "-$num" + '.xlsx')
            $num++
        }

        $wb.SaveAs($nextname, 51)
    }

    END {
        $xl.Quit()
    
        $null = $ws, $wb, $xl | ForEach-Object {Release-Ref $_}

        # del $CSVFile
    }
}

function csv {

    $user = $TextBoxID.text
    $username = Get-ADUser $user -Properties Name, DisplayName -ErrorVariable erro2 | Select-Object Name, Displayname
    $file = (Get-ADUser $user -Properties MemberOf -ErrorVariable erro2).MemberOf | Get-ADGroup -Properties Name, Description | Select-Object Name, Description

    if ($erro2) {
        $LabelError.Visible = $true
        Start-Sleep -s 2
        $LabelError.Visible = $false
    }

    else {
        function csv_data {
        $file | Export-CSV -Path "$Env:USERPROFILE\Documents\AD_$user.csv" -NoTypeInformation | Format-Table
        $date = get-date -format "dddd, dd/MM/yyyy, HH:mm:ss" ; $date | Add-Content "$Env:USERPROFILE\Documents\AD_$user.csv"
        $username | Add-Content "$Env:USERPROFILE\Documents\AD_$user.csv"
        Save-CSVasExcel "$Env:USERPROFILE\Documents\AD_$user.csv"
        Invoke-Expression "$Env:USERPROFILE\Documents\AD_$user.xlsx"
        Remove-Item "$Env:USERPROFILE\Documents\AD_$user.csv"
        Start-Sleep -s 1
        }
        do {Start-Export} while (csv_data)
        $LabelProcessing.Visible = $false
    }
    
}



############################################# 
#                   GUI
############################################

function Start-Export {
    $LabelError.Visible = $false
    $LabelProcessing.Visible = $true
    Start-Sleep -s 3
}

function Start-AD {
    
    Add-Type -AssemblyName System.Windows.Forms    
 
     $FormProgram                  = New-Object System.Windows.Forms.Form
     $FormProgram.Text             = "Mensagem de AD Explorer"
     $FormProgram.Size             = New-Object System.Drawing.Size(290,85)
     $FormProgram.StartPosition    = 'CenterScreen'
     $FormProgram.MaximizeBox      = $false
     $FormProgram.MinimizeBox      = $false
     $FormProgram.CloseBox         = $false
     $FormProgram.TopMost          = $false
 
     $Labelmsg                     = New-Object System.Windows.Forms.Label
     $Labelmsg.Location            = New-Object System.Drawing.Size(33,10) 
     $Labelmsg.width               = 25
     $Labelmsg.height              = 10
     $Labelmsg.Autosize            = $true
     $Labelmsg.Text                = "Carregando o Active Directory. Aguarde..."
     $FormProgram.Controls.Add($Labelmsg)
 
     $FormProgram.Show()| Out-Null
 
     Start-Sleep -Seconds 10
 
     $FormProgram.Close() | Out-Null
}


Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$FormAdExplorer                  = New-Object system.Windows.Forms.Form
$FormAdExplorer.ClientSize       = New-Object System.Drawing.Point(555,291)
$FormAdExplorer.text             = "AD Explorer: uma forma simples de consultar dados do AD"
$FormAdExplorer.TopMost          = $false
$FormAdExplorer.StartPosition    = 'CenterScreen'
$FormAdExplorer.FormBorderStyle  = 'Fixed3D'
$FormAdExplorer.MaximizeBox      = $false

$iconBase64                      = "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
AABAREQAQERECkBERA5AREQKQEREAAAAAAAAAAAAAAAAAAAAAAAAAAAAQEREAEBERAhAREQIQEREBAAAAAAAAAAAQEREAkBERAZAREQGQEREAAAAAAAAAAAAAAAAAAAAAAAAAAAAQEREAEBERApAREQOQERECkBERAAAAAAAAAAAAAAAAAAAAAAAAAAAAEBERABAREQKQEREDlZVPAzzvwA++cUA6PG9AP/yvgD/8r4A//K+AP/xvQD/6bgD/+e3BP/puAP/8b0A//K+AP/yvgD/8r4A//S/APz4wwDA7boBFFNRPQxAREQOQERECkBERAAAAAAAAAAAAAAAAAAAAAAAAAAAAEBERABAREQKQEREDkBERApAREQAAAAAAAAAAAAAAAAAAAAAAAAAAABAREQAQEREBkBERARAREQAAAAAAAAAAABAREQEQERECEBERAhAREQAAAAAAAAAAAAAAAAAAAAAAAAAAABAREQAQERECkBERA5AREQKQEREAAAAAAAAAAAAAAAAAAAAAAAAAAAAQEREAFRTPAzdrwlw770D+PG9APzyvgD/8r4A//K+AP/yvgD/8r4A//G9AP/puAP/57cE/+m4A//xvQD/8r4A//K+AP/yvgD/8r4A/PjDANjxvgAkS09ADEBERA5AREQKQEREAAAAAAAAAAAAAAAAAAAAAAAAAAAAQEREAEBERApAREQOQERECkBERAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABAREQCQEREAgAAAAAAAAAAAAAAAEBERAZAREQMQERECkBERAAAAAAAAAAAAAAAAAAAAAAAAAAAAEBERABAREQKQEREDkBERApAREQAAAAAAAAAAAAAAAAA/78AAO60AAT4xACO7LsD/+a2BP/puAP/8b0A//K+AP/yvgD/8r4A//K+AP/yvgD/8b0A/+m4A//ntwT/6bgD//G9AP/yvgD/8r4A//K+AP/xvQD8+cUA6PG8ADxYUzsMQEREDkBERApAREQAAAAAAAAA
AAAAAAAAAAAAAAAAAABAREQAQERECkBERA5AREQKQEREAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAEBERAAAAAAAAAAAAAAAAABAREQAQERECkBERA5AREQKQEREAAAAAAAAAAAAAAAAAAAAAAAAAAAAQEREAEBERApAREQOQERECkBERAAAAAAA/78AAPC7ABL4wwC29MAA/PG9AP/puAP/57cE/+m4A//xvQD/8r4A//K+AP/yvgD/8r4A//K+AP/xvQD/6bgD/+e3BP/puAP/8b0A//K+AP/yvgD/8r4A//G9APz4wwD2878AXltWOgxDQkIOQERECkBERAAAAAAAAAAAAAAAAAAAAAAAAAAAAEBERABAREQKQEREDkBERApAREQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAEBERABAREQKQEREDkBERApAREQAAAAAAAAAAAAAAAAAAAAAAAAAAABAREQAQERECkBERA5AREQKyX8TAvS/ACT3wwDY8r4A/PK+AP/yvgD/8b0A/+m4A//ntwT/6bgD//G9AP/yvgD/8r4A//K+AP/yvgD/8r4A//G9AP/puAP/57cE/+m4A//xvQD/8r4A//K+AP/yvgD/8b0A//TAAP/4xAB+S09ADEVJQg5AREQKQEREAAAAAAAAAAAAAAAAAAAAAAAAAAAAQEREAEBERApAREQOQERECkBERAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAQEREAEBERApAREQOQERECkBERAAAAAAAAAAAAAAAAAAAAAAAAAAAAEBERABAREQKUVE8ENmtCUr5xQDq8b0A//K+AP/yvgD/8r4A//K+AP/xvQD/6bgD/+e3BP/puAP/8b0A//K+AP/yvgD/8r4A//K+AP/yvgD/8b0A/+m4A//ntwT/6bgD//G9AP/yvgD/8r4A//K+AP/xvQD/9MEA/PjEAKCGcigSS0tADkBE
RApAREQAAAAAAAAAAAAAAAAAAAAAAAAAAABAREQAQERECkBERA5AREQKQEREAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABAREQAQERECkBERA5AREQKQEREAAAAAAAAAAAAAAAAAAAAAAD/fwAAkpQYAuO0BnTsugT66bgD/PG9AP/yvgD/8r4A//K+AP/yvgD/8r4A//G9AP/puAP/57cE/+m4A//xvQD88b0A//G9AP/xvQD88b0A//K+AP/xvQD/6bgD/+e3BP/puAP/8b0A//K+AP/yvgD/8r4A//K+AP/0vwD898IAwrOTGCBNTT8QQERECkBERAAAAAAAAAAAAAAAAAAAAAAAAAAAAEBERABAREQKQEREDkBERApAREQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAEBERABAREQKQEREDkBERApAREQAAAAAAAAAAAD/vwAA57EABvnFAJL0wAD86bgD/+e3BP/puAP/8b0A//K+AP/yvgD/8r4A//K+AP/yvgD/8b0A/+m4A//rugT87r0D1vbBAJ70wACY+sUAxPbBAPzxvQD/8r4A//G9AP/puAP/57cE/+m4A//xvQD/8r4A//K+AP/yvgD/8r4A//G9AP/3wwDayaEQLkhMQQ5AREQKQEREAAAAAAAAAAAAAAAAAAAAAAAAAAAAQEREAEBERApAREQOQERECAAAAAAAAAAAAAAAAEBERABAREQCAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAQEREAEBERApAREQOQERECkBERAD/qgAA7rsAFPjDALz0wAD88r4A//G9AP/puAP/57cE/+m4A//xvQD/8r4A//K+AP/yvgD/8r4A//G9AP/4xADs3bAIWFNPPRBSUT0MvKIXAv+4AAL0vwAy98IA2PG9AP/yvgD/8b0A/+m4A//ntwT/6bgD//G9AP/yvgD/8r4A//K+AP/yvgD/8b0A/PnF
AOjYrQpGU1I9EEBERApAREQAAAAAAAAAAAAAAAAAAAAAAAAAAABAREQAQERECkBERApAREQGAAAAAAAAAABAREQAQEREAkBERAQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABAREQAQERECkBERA5STD0M8b8AKPjEANryvgD88r4A//K+AP/yvgD/8b0A/+m4A//ntwT/6bgD//G9AP/yvgD/8r4A//G9AP/2wQD68r8AOr2LDgJAREQKQEREDkBERApAREQA3bIAAvK8ABr4wwDm8b0A//K+AP/xvQD/6bgD/+e3BP/puAP/8b0A//K+AP/yvgD/8r4A//K+AP/xvQD898IA9uG0B2ZTUj0QQ0JCDEBERAAAAAAAAAAAAAAAAAAAAAAAAAAAAEBERABAREQIQERECEBERAQAAAAAAAAAAEBERAJAREQGQEREBkBERAAAAAAAAAAAAAAAAAAAAAAAAAAAAEBERABbVToM1KkMUvC+A+7xvQD88r4A//K+AP/yvgD/8r4A//K+AP/xvQD/6bgD/+e3BP/puAP/8b0A//K+AP/xvgD898MApv/UAAIAAAAAQEREAEBERApAREQOQERECkBERAD/1AAA8r8AbvXBAPzyvgD/8r4A//G9AP/puAP/57cE/+m4A//xvQD/8r4A//K+AP/yvgD/8r4A//G9AP/zwAD/6roFhkhMQQ5HSkEMQEREAAAAAAAAAAAAAAAAAAAAAAAAAAAAQEREAEBERAZAREQEQEREAAAAAAAAAAAAQEREBEBERAhAREQIQEREAAAAAAAAAAAAAAAAAP//AAD/qgAA9sIAcu67A/rmtgT86bgD//G9AP/yvgD/8r4A//K+AP/yvgD/8r4A//G9AP/puAP/57cE/+m4A//xvQD/98IA/PC9AGD/qgAAAAAAAAAAAABAREQAQERECkBERA5AREQKz5sRAvG+ACz3wgD88r4A//K+AP/yvgD/8b0A/+m4A//ntwT/6bgD//G9AP/yvgD/8r4A//K+
AP/yvgD/8b0A//TAAPztvASmfWwsFk5NPwxAREQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAQEREAkBERAIAAAAAAAAAAAAAAABAREQGQEREDEBERApAREQAAAAAAP+/AADksgAG+cUAmPTAAPzxvQD/6bgD/+e3BP/puAP/8b0A//K+AP/yvgD/8r4A//K+AP/yvgD/8b0A/+m4A//ntwT/6bgD//fCAPzxvQBc//8AAAAAAAAAAAAAAAAAAEBERABAREQKQEREDlZROwzvuwAm9sIA/PK+AP/yvgD/8r4A//K+AP/xvQD/6bgD/+e3BP/puAP/8b0A//K+AP/yvgD/8r4A//K+AP/yvgD/878A/O27BMaoihwiTFA9DEBERAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABAREQAAAAAAAAAAAAAAAAAQEREAEBERApAREQOQERECr2UFwDwvwAW+MMAwvTAAPzyvgD/8r4A//G9AP/puAP/57cE/+m4A//xvQD/8r4A//K+AP/yvgD/8r4A//K+AP/xvQD/6bgD/+e3BP/quAP/9sEAlL+/AAAAAAAAAAAAAAAAAAAAAAAAQEREAEBERApQUD4Q3bAHZPbCAPzyvgD/8r4A//K+AP/yvgD/8r4A//G9AP/puAP/57cE/+m4A//xvQD/8r4A//K+AP/yvgD/8r4A//K+AP/xvQD/7r0D3L+bFDBPTj4MQEREAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABAREQAQERECk5NPxDNpQ80+MQA3vG+AP/yvgD/8r4A//K+AP/yvgD/8b0A/+m4A//ntwT/6bgD//G9AP/yvgD/8r4A//K+AP/yvgD/8r4A//G9AP/puAP/57cE/+6+A/DqugAe/78AAAAAAAAAAAAAAAAAAAAAAABAREQAgXAqEOy6BNjpuAP88b0A//K+AP/yvgD/8r4A//K+AP/yvgD/8b0A/+m4A//ntwT/6bgD//G9
AP/yvgD/8r4A//K+AP/yvgD/8r4A//G9APzwvwPqz6YNSFhXOwxAREQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAANalDgLcrwhW7bwE8Om4A/zxvQD/8r4A//K+AP/yvgD/8r4A//K+AP/xvQD/6bgD/+e3BP/puAP/8b0A//K+AP/yvgD/8r4A//K+AP/yvgD/8b0A/+y7A/zouAamlH0iFM+IEQLUqQAC//8AAP9/AAD/1AAA/78AAOGgCgLesAde7bwE6um4A//xvQD/8r4A//K+AP/yvgD/8r4A//K+AP/xvQD/6bgD/+e3BP/puAP/8b0A//K+AP/yvgD/8r4A//K+AP/yvgD/8b0A/O+9A/jbrgloWFc7DHwvLwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA//8AAP//AAD4wwB49MAA/Om4A//ntwT/6bgD//G9AP/yvgD/8r4A//K+AP/yvgD/8r4A//G9AP/puAP/6rkE/++9A/z2wQD8+MMA/PG9AP/yvgD/8b0A//XAAP/3xACEV1E7DLqZFS7suwSq8r4ARvS/ABLwwAAM8r4ANvnGAJr1wQBcz7ARAryZFCTqugW27LsD/PG9AP/yvgD/8r4A//K+AP/xvgD88b0A/PG9AP/puAP/57cE/+m4A//xvQD/8r4A//K+AP/yvgD/8r4A//K+AP/xvQD/7LoD/+S3B4hLT0AKg4UZAAAAAAAAAAAAAAAAAAAAAAAAAAAA/78AAO+9AAj5xQCe9cEA/PG9AP/xvQD/6bgD/+e3BP/puAP/8b0A//K+AP/yvgD/8r4A//K+AP/yvgD/98IA4uKzBnbFnxE8yJ8QLPO+AEr4xACs878A//XAAPr0wQBq/8IAAvG+ADLwvwPm5rYE/O+9A/zzvgD6878A+PfDAPzyvgD89cEA//nEAKrxuwEWUlE9DNyvCG7wvgPw8L0A//G+AP/1wQD898IA5PnEANb0wADw9sIA/Om4
A//ntwT/6bgD//G9AP/yvgD/8r4A//K+AP/yvgD/8r4A//G9AP/suwP86bkGqJd8IhDQoREAAAAAAAAAAABAREQAeGUvBO+8ABr2wgDG9L8A/PK+AP/yvgD/8r4A//G9AP/puAP/57cE/+m4A//xvQD/8r4A//K+AP/xvQD/9sEAxPS9ABLJfxMCV1I7DFBQPRBSUTsMzY8RAvXAAGL0vgBO8LQABPXAAEz5wwDw8b0A/Om4A//ntwT/6bgD//G9AP/yvgD/8r4A//K+AP/yvgD8+cQA5PS/AE5aWTkMupkVLu27BMb3wgC29MAAMP+/AAL/1AAA8boACPS/AFj4xADi6bgD/+e3BP/puAP/8b0A//K+AP/yvgD/8r4A//K+AP/yvgD/8b0A/+y6AvztuwPGyqMPGt2pDAAAAAAA1bYPAue3BDD1wgHg8b0A//K+AP/yvgD/8r4A//K+AP/yvgD/8b0A/+m4A//ntwT/6bgD//G9AP/xvgD/98IA8PG6ABjUqgACAAAAAEBERABAREQKQEREDkBERAqjiBYC/7oAAvXBAGr1wAD68b0A//K+AP/xvQD/6bgD/+e3BP/puAP/8b0A//K+AP/yvgD/8r4A//G9APz1wQD8+cUAmqiKGxxfWDcSVlE6DM+wEQIAAAAAAAAAAP9/AADUqQAC8r8AHvjDAOLpuAP/57cE/+m4A//xvQD/8r4A//K+AP/yvgD/8r4A//K+AP/xvQD/67oC//G/AtzkswYolJQAAPS/AFb3wgDy7boC/O26Af/xvQD/8r4A//K+AP/yvgD/8r4A//K+AP/xvQD/6bgD/+e3BP/puAP/8b0A/PfCAKD/qQACAAAAAAAAAAAAAAAAQEREAEBERApAREQOQkZDCuarCAT3wgDo8b0A/PK+AP/yvgD/8r4A//G9AP/puAP/57cE/+m4A//xvQD/8r4A//K+AP/yvgD/8r4A//O/APz3wwDKVVQ8DENHQg5AREQKQEREAAAAAAAAAAAAAAAAAOK4
AALyvgBY9cIA/Om4A//ntwT/6bgD//G9AP/yvgD/8r4A//K+AP/yvgD/8r4A//G9AP/tugH89cIB6u68AUL3xQCK9MEA/O67Af/ruQL/67kC//G9AP/yvgD/8r4A//K+AP/yvgD/8r4A//G9AP/puAP/57cE/+y7A/zxvQB2AAAAAAAAAAAAAAAAAAAAAAAAAABAREQAQERECkBERA5aVToM+cUAyPG9APzyvgD/8r4A//K+AP/yvgD/8b0A/+m4A//ntwT/6bgD//G9AP/yvgD/8r4A//K+AP/yvQD/9cEAmMyOEgJAREQKQEREDkBERApAREQAAAAAAAAAAAD//wAA8sAACvS/APTxvQD/6bgD/+e3BP/puAP/8b0A//K+AP/yvgD/8r4A//K+AP/yvgD/8b0A//PAAPz0vwFw//8AAvjDAKbzvwD87LoC/+i3A//puAP/8b0A//K+AP/yvgD/8r4A//K+AP/yvgD/8b0A/+m4A//ntwT/5rUFkr2UFwAAAAAAAAAAAAAAAAAAAAAAAAAAAEBERAA+QkIKVFA7EO68A+LxvQD88r4A//K+AP/yvgD/8r4A//K+AP/xvQD/6bgD/+e3BP/puAP/8b0A//K+AP/yvgD/878A//G9AIAAAAAAQEREAEBERApAREQOQERECkBERAAAAAAAAAAAAH9/AAD2wQDo8b0A/PG9AP/puAP/57cE/+m4A//xvQD/8r4A//K+AP/yvgD/8r4A//TAAPz3wwCSrpgcAv//AADvvwAI98IAwPK9APzruQL/57cE/+m4A//xvQD/8r4A//K+AP/yvgD/8r4A//K+AP/xvQD/6bgD/Oq5BN6DcSwQm54jAAAAAAAAAAAAAAAAAAAAAAAAAAAA17cOAtWqCkjrugT86bgD//G9AP/yvgD/8r4A//K+AP/yvgD/8r4A//G9AP/puAP/57cE/+m4A//xvQD/8r4A//G9APz4wwCm/9QAAgAAAABAREQAQERECkBERA5AREQKQEREAP+/
AADuvAAY9cAA+vG9AP/yvgD/8b0A/+m4A//ntwT/6bgD//G9AP/yvgD/8r4A//K+APz3wgC28sAABP9/AAAAAAAA/78AAO25ABL4wwDW8b0A/Om4A//ntwT/6bgD//G9AP/yvgD/8r4A//K+AP/yvgD/8r4A//G9AP/sugP85bYGklVPPAzMjhICAAAAAAAAAAD/fwAA//8AAO66ABb3wgDa6bgD/+e3BP/puAP/8b0A//K+AP/yvgD/8r4A//K+AP/yvgD/8b0A/+m4A//ntwT/6bgD//G9AP/xvQD/98MA8vK+ABr/zAAAAAAAAEBERABAREQKQEREDkBERArcrgwC9MAAfvXAAPzyvgD/8r4A//K+AP/xvQD/6bgD/+e3BP/puAP/8b0A//G9APz3wwDS7rkAEP+qAAAAAAAAAAAAAAAAAADMzAAA8LwAIPjFAOTxvQD86bgD/+e3BP/puAP/8b0A//K+AP/yvgD/8r4A//K+AP/yvgD/8b0A/+y7A/zouAaqy6QPMJ93IgD/qgAA77cACP+4AAL5xQB69cEA/PG9AP/puAP/57cE/+m4A//xvQD/8r4A//K+AP/yvgD/8r4A//K+AP/xvQD/6bgD/+e3BP/puAP/8r4A/PfDAMLvwQAM67AAAv//AAD/uAACv4gWAllXOgxQUD4Q3bAIVPXBAPjxvgD/8r4A//K+AP/yvgD/8r4A//G9AP/puAP/57cE/+m4A/z4xQDk8L0AIP+/AAAAAAAAAAAAAAAAAAAAAAAAAAAAAP+UAAD0vwAw+MMA8PG9AP/puAP/57cE/+m4A//xvQD/8r4A//K+AP/yvgD/8r4A//K+AP/xvQD/6rgD/+q5BPzvvAPg+cQA1PfCAPD0wABi/88AAvjDALDyvwD88b0A/+m4A//ntwT/6bgD//G9AP/yvgD/8r4A//K+AP/yvgD/8r4A//G9AP/puAP/5rYE/Oy7A9TmtQES9L4APvrEAOD0vwB28r0AMPC9ACDxvgBE7LwEpOm4
BPzpuAP/8b0A//K+AP/yvgD/8r4A//K+AP/yvgD/8b0A/+m4A//suwT00agLPMN7FQIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA/9QAAPO+AEr1wQD68b0A/+m4A//ntwT/6bgD//G9AP/yvgD/8r4A//K+AP/yvgD/8r4A//G9AP/puAP/5rYE/Om4A/zxvQD898IA+PK/ADzxugAO98IA0PG9APzxvQD/6bgD/+e3BP/puAP/8b0A//K+AP/yvgD/8r4A//K+AP/yvgD/8b0A/O++A9yvkBkmzqcNNvjEAOzxvQD89cAA//fDAPz2wQD8+MMA/PG+AP/puAP/57cE/+m4A//xvQD/8r4A//K+AP/yvgD/8r4A//G9AP/0wAD837IIXk9PPRBAREQIAAAAAAAAAAAAAAAAQEREAEBERAIAAAAAAAAAAAAAAADlrAAC9cEAZvXBAPzxvQD/6bgD/+e3BP/puAP/8b0A//K+AP/yvgD/8r4A//K+AP/yvgD/8b0A/+m4A//ntwT/6bgD//G9APz4xQDo77oAIu28ACD5xQDk8b0A/PG9AP/puAP/57cE/+m4A//xvQD/8r4A//K+AP/yvgD/8b0A/PjFAOTvuwAkx6APLOy9BObpuAP88b0A//K+AP/yvgD/8r4A//K+AP/yvgD/8b0A/+m4A//ntwT/6bgD//G9AP/yvgD/8r4A//K+AP/0wQD898IAetalDgJAREQKQERECkBERAYAAAAAAAAAAEBERABAREQCQEREBAAAAAAAAAAAAAAAAOWsAAL4wwCC9MAA/PG9AP/puAP/57cE/+m4A//xvQD/8r4A//K+AP/yvgD/8r4A//K+AP/xvQD/6bgD/+e3BP/puAP/8b0A/PfDANTyvwAQ8r4ANPfDAPTxvQD/8b0A/+m4A//ntwT/6bgD//G9AP/yvgD/8b0A/PnEAOzyvgAu77sAGvjFANzpuAP857cE/+m4A//xvQD/8r4A//K+AP/yvgD/8r4A//K+
AP/xvQD/6bgD/+e3BP/puAP/8b0A//K+AP/0vwD898MAnt+fAAIAAAAAQEREAEBERAhAREQIQEREBAAAAAAAAAAAQEREAkBERAZAREQGQEREAAAAAAAAAAAA/40AAvfDAKD0wAD88b0A/+m4A//ntwT/6bgD//G9AP/yvgD/8r4A//K+AP/yvgD/8r4A//G9AP/puAP/57cE/+m4A//yvgD898MAtP/IAAL0wABW9cEA/PG9AP/xvQD/6bcD/Oa2BPzpuAP/8b0A//jDAPTzvwA+77sAEPbCANLxvQD/8b0A/+m4A//ntwT/6bgD//G9AP/yvgD/8r4A//K+AP/yvgD/8r4A//G9AP/puAP/57cE/+m4A//xvQD898IAwPS/AAb//wAAAAAAAAAAAABAREQAQEREBkBERARAREQAAAAAAAAAAABAREQEQERECEBERAhAREQAAAAAAP//AADqswAG98MAvPK+APzxvQD/6bgD/+e3BP/puAP/8b0A//K+AP/yvgD/8r4A//K+AP/yvgD/8b0A/+m4A//ntwT/6bgD//TAAPz4xACK/8oABPfCAHz3wwD89cAA9PnGAMrvvwPI6bgE7vC+A/rzvwBQ6roACPfCAMLyvgD88r4A//K+AP/xvQD/6bgD/+e3BP/puAP/8b0A//K+AP/yvgD/8r4A//K+AP/yvgD/8b0A/+m4A//mtgT87r0D3Oq4ART/vwAAAAAAAAAAAAAAAAAAAAAAAAAAAABAREQCQEREAgAAAAAAAAAAAAAAAEBERAZAREQMQERECkBERAAAAAAA/78AAO+7ABD3wwDS8b0A/PG9AP/puAP/57cE/+m4A//xvQD/8r4A//K+AP/yvgD/8r4A//K+AP/xvQD/6bgD/+e3BP/puAP/9cAA/PS/AGTjyAAC878AWPbDAA7/uAACv6YNAol1JxLLpQ8+dmguEPfCAK7zvwD88r4A//K+AP/yvgD/8r4A//G9AP/puAP/57cE/+m4A//xvQD/8r4A//K+
AP/yvgD/8r4A//K+AP/xvQD88L4D6ryaFDJRUD0MQEREAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAEBERAAAAAAAAAAAAAAAAABAREQAQERECkBERA5AREQKQEREAAAAAADMzAAA8b0AHvnFAOLyvgD88b0A/+m4A//ntwT/6bgD//G9AP/yvgD/8r4A//K+AP/yvgD/8r4A//G9AP/puAP/57cE/+m4A//2wwD277wAKv+fAAKqqgAAAAAAAAAAAACbYCMAYFs4DM+mDUzuvQP88b0A//K+AP/yvgD/8r4A//K+AP/yvgD/8b0A/+m4A//ntwT/6bgD//G9AP/yvgD/8r4A//K+AP/xvQD/9sIA+PK+AEBNUD0MQEREDkBERApAREQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAEBERABAREQKQEREDkBERApAREQAAAAAAP+qAADyvgAs+MQA7vG9AP/xvQD/6bgD/+e3BP/puAP/8b0A//K+AP/yvgD/8r4A//K+AP/yvgD/8b0A/+m4A//mtgT87rwDzumxBwQAAAAAAAAAAAAAAAAAAAAAAAAAAEBERABZTjoM5rcGqOq4A//xvQD/8r4A//K+AP/yvgD/8r4A//K+AP/xvQD/6bgD/+e3BP/puAP/8b0A//K+AP/xvQD/9cEA/PPAAGL/1AACQEREAEBERApAREQOQERECkBERAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAQEREAEBERApAREQOQERECkBERAAAAAAA/9QAAPK+AET2wgD48b0A//G9AP/puAP/57cE/+m4A//xvQD/8r4A//K+AP/yvgD/8r4A//K+AP/xvQD/7bsD/9utCHhKSUAMQEREAAAAAAAAAAAAAAAAAAAAAAAAAAAA158OAtarCUrtvAT86bgD//G9AP/yvgD/8r4A//K+AP/yvgD/8r4A//G9AP/puAP/57cE/+m4
A//xvQD/9MAA/PfEAIT/vwACAAAAAAAAAABAREQAQERECkBERA5AREQKQEREAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABAREQAQERECkBERA5AREQKQEREAAAAAAD/1AAC88AAYvTBAPzxvQD/8b0A/+m4A//ntwT/6bgD//G9AP/yvgD/8r4A//K+AP/yvgD/8r4A//jDAPzbrQdcT1I9EEBERApAREQAAAAAAAAAAAAAAAAAAAAAAP+qAADuuwEm770D/Oe3BP/puAP/8b0A//K+AP/yvgD/8r4A//K+AP/yvgD/8b0A/+m4A//ntwT/67kD/PjDAKj/qQACAAAAAAAAAAAAAAAAAAAAAEBERABAREQKQEREDkBERApAREQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAEBERABAREQKQEREDkBERApAREQAAAAAAOWsAAL3wwB+9cEA/PK+AP/xvQD/6bgD/+e3BP/puAP/8b0A//K+AP/yvgD/8r4A//K+AP/1wQD/8LwAcEpOQApAREQOQERECkBERAAAAAAAAAAAAAAAAAD/uAAC8b0ARPfDAPzpuAP/57cE/+m4A//xvQD/8r4A//K+AP/yvgD/8r4A//K+AP/xvQD/6bgD/Om4BM6YfiIUn3ciAAAAAAAAAAAAAAAAAAAAAAAAAAAAQEREAEBERApAREQOQERECAAAAAAAAAAAAAAAAEBERABAREQCAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAQEREAEBERApAREQOQERECkBERAAAAAAA/9QAAvfEAJr0wAD88r4A//G9AP/puAP/57cE/+m4A//xvQD/8r4A//K+AP/yvgD/8b0A/PjDAM7XugkER0ZBDEBERA5AREQKQEREAAAAAAAAAAAA5ZgAAvfDAKTyvgD/8b0A/+m4A//ntwT/6bgD//G9AP/yvgD/8r4A//K+AP/yvgD/8b0A/PjF
AN66mBQiS0tADkBERApAREQAAAAAAAAAAAAAAAAAAAAAAAAAAABAREQAQERECkBERApAREQGAAAAAAAAAABAREQAQEREAkBERAQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABAREQAQERECkBERA5AREQKQEREAP9/AADctgAE+MIAtvK/APzyvgD/8b0A/+m4A//ntwT/6bgD//G9AP/yvgD/8r4A//K+AP/0wAD8+MMAismNEwJUUzoMSEhBDkdLQQzXnw4C1KoAAvXBAGD0wAD88r4A//K+AP/xvQD/6bgD/+e3BP/puAP/8b0A//K+AP/yvgD/8b0A//jEAO7yvwAs048PAkBERApAREQOQERECkBERAAAAAAAAAAAAAAAAAAAAAAAAAAAAEBERABAREQIQERECEBERAQAAAAAAAAAAEBERAJAREQGQEREBkBERAAAAAAAAAAAAAAAAAAAAAAAAAAAAEBERABAREQKQEREDkBERApAREQA/6oAAO25AA73wQDQ8b0A/PK+AP/xvQD/6bgD/+e3BP/puAP/8b0A//K+AP/yvgD/8r4A//XAAPz4wwC28r0AQKKEHxiFcSkY0agNOvjDAJrzvwD/8b0A//K+AP/yvgD/8r4A//G9AP/puAP/57cE/+m4A//xvQD/8b0A//XBAPrzvgBK/9QAAAAAAABAREQAQERECkBERA5AREQKQEREAAAAAAAAAAAAAAAAAAAAAAAAAAAAQEREAEBERAZAREQEQEREAAAAAAAAAAAAQEREBEBERAhAREQIQEREAAAAAAAAAAAAAAAAAAAAAAAAAAAAQEREAEBERApAREQOQERECkBERADMmQAA8LsAHPnFAN7yvgD88r4A//G9AP/puAP/57cE/+m4A//xvQD/8r4A//K+AP/yvgD/8b4A//jDAPzyvgD467kD9uu6BPzquQP88b0A//K+AP/yvgD/8r4A//K+AP/yvgD/8b0A/+m4A//ntwT/6bgD//XBAPz0wABu/78AAgAA
AAAAAAAAAAAAAEBERABAREQKQEREDkBERApAREQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAQEREAkBERAIAAAAAAAAAAAAAAABAREQGQEREDEBERApAREQAAAAAAAAAAAAAAAAAAAAAAAAAAABAREQAQERECkBERA5AREQKQEREAP+/AADvvQAq+MQA7PG9APzyvgD/8b0A/+m4A//ntwT/6bgD//G9AP/yvgD/8r4A//K+AP/yvgD/8b0A//G9AP/puAP/57cE/+m4A//xvQD/8r4A//K+AP/yvgD/8r4A//K+AP/xvQD/6bgD/+m5BPzpugWWvaUOAgAAAAAAAAAAAAAAAAAAAAAAAAAAQEREAEBERApAREQOQERECkBERAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABAREQAAAAAAAAAAAAAAAAAQEREAEBERApAREQOQERECkBERAAAAAAAAAAAAAAAAAAAAAAAAAAAAEBERABAREQKQEREDkBERApAREQA/9QAAPG8AED3wgD28b0A//K+AP/xvQD/6bgD/+e3BP/puAP/8b0A//K+AP/yvgD/8r4A//K+AP/yvgD/8b0A/+m4A//ntwT/6bgD//G9AP/yvgD/8r4A//K+AP/yvgD/8r4A//K+APzsuwS4Yl81EkJGQQxAREQAAAAAAAAAAAAAAAAAAAAAAAAAAABAREQAQERECkBERA5AREQKQEREAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABAREQAQERECkBERA5AREQKQEREAAAAAAAAAAAAAAAAAAAAAAAAAAAAQEREAEBERApAREQOQERECkBERAD/uAAC9MAAXPTAAPzxvQD/8r4A//G9AP/puAP/57cE/+m4A//xvQD/8r4A//K+AP/yvgD/8r4A//K+AP/xvQD/6bgD/+e3BP/puAP/8b0A//K+AP/yvgD/8r4A//G9APz3wgDS7LcCEEtKPwxAREQOQERECkBE
RAAAAAAAAAAAAAAAAAAAAAAAAAAAAEBERABAREQKQEREDkBERApAREQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAEBERABAREQKQEREDkBERApAREQAAAAAAAAAAAAAAAAAAAAAAAAAAABAREQAQERECkBERA5AREQKQEREAP+/AAL3wgB49cEA/PK+AP/yvgD/8b0A/+m4A//ntwT/6bgD//G9AP/yvgD/8r4A//K+AP/yvgD/8r4A//G9AP/puAP/57cE/+m4A//xvQD/8r4A//G9APz4xQDk8LwAIP//AABAREQAQERECkBERA5AREQKQEREAAAAAAAAAAAAAAAAAAAAAAAAAAAAQEREAEBERApAREQOQERECkBERAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAQEREAEBERApAREQOQERECkBERAAAAAAAAAAAAAAAAAAAAAAAAAAAAEBERABAREQKQEREDkBERApAREQA/7gAAvjDAJT0wAD88r4A//K+AP/xvQD/6bgD/+e3BP/puAP/8b0A//K+AP/yvgD/8r4A//K+AP/yvgD/8b0A/+m4A//ntwT/6bgD//G9AP/4wwD08r0ANP+UAAAAAAAAAAAAAEBERABAREQKQEREDkBERApAREQAAAAAAAAAAAAAAAAAAAAAAAAAAABAREQAQERECkBERA5AREQKQEREAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABAREQAQERECkBERA5AREQKQEREAAAAAAAAAAAAAAAAAAAAAAAAAAAAQEREAEBERApAREQOQERECl9hIgDtuwAE98IAsvO/APzyvgD/8r4A//G9AP/puAP/57cE/+m4A//xvQD/8r4A//K+AP/yvgD/8r4A//K+AP/xvQD/6bgD/+a2BP/tuwP8878AVri4AAIAAAAAAAAAAAAAAAAAAAAAQEREAEBE
RApAREQOQERECkBERAAAAAAAAAAAAAAAAAAAAAAAAAAAAEBERABAREQKQEREDkBERAgAAAAAAAAAAAAAAABAREQAQEREAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAEBERABAREQKQEREDkBERApAREQAAAAAAAAAAAAAAAAAAAAAAAAAAABAREQAQERECkBERA5AREQKn3ciAO+4AAz2wgDM8r4A/PK+AP/yvgD/8b0A/+m4A//ntwT/6bgD//G9AP/yvgD/8r4A//K+AP/yvgD/8r4A//G9AP/tuwP84rQHgFtVOgxAREQAAAAAAAAAAAAAAAAAAAAAAAAAAABAREQAQERECkBERA5AREQKQEREAAAAAAAAAAAAAAAAAAAAAAAAAAAAQEREAEBERApAREQKQEREBgAAAAAAAAAAQEREAEBERAJAREQEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAQEREAEBERApAREQOQERECkBERAAAAAAAAAAAAAAAAAAAAAAAAAAAAEBERABAREQKQEREDkBERAq/lhYC8bsAGPnFANzxvQD88r4A//K+AP/xvQD/6bgD/+e3BP/puAP/8b0A//K+AP/yvgD/8r4A//K+AP/0vwD898MAnFhXOwxDR0IOQERECkBERAAAAAAAAAAAAAAAAAAAAAAAAAAAAEBERABAREQKQEREDkBERApAREQAAAAAAAAAAAAAAAAAAAAAAAAAAABAREQAQERECEBERAhAREQEAAAAAAAAAABAREQCQEREBkBERAZAREQAAAAAAAAAAAAAAAAAAAAAAAAAAABAREQAQERECkBERA5AREQKQEREAAAAAAAAAAAAAAAAAAAAAAAAAAAAQEREAEBERApAREQOQERECr+WFgLxvQAm+MUA6vG9APzyvgD/8r4A//G9AP/puAP/57cE/+m4A//xvQD/8r4A//K+AP/yvQD89sIAwPS9AAafYSIAQERECkBERA5AREQKQEREAAAAAAAAAAAAAAAAAAAA
AAAAAAAAQEREAEBERApAREQOQERECkBERAAAAAAAAAAAAAAAAAAAAAAAAAAAAEBERABAREQGQEREBEBERAAAAAAAAAAAAEBERARAREQIQERECEBERAAAAAAAAAAAAAAAAAAAAAAAAAAAAEBERABAREQKQEREDkBERApAREQAAAAAAAAAAAAAAAAAAAAAAAAAAABAREQAQERECkBERA5AREQKlGsWAvK+ADr3wwD28b0A//K+AP/yvgD/8b0A/+m4A//ntwT/6bgD//G9AP/xvQD8+MQA2u68ABT/vwAAAAAAAEBERABAREQKQEREDkBERApAREQAAAAAAAAAAAAAAAAAAAAAAAAAAABAREQAQERECkBERA5AREQKQEREAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAEBERAJAREQCAAAAAAAAAAAAAAAAQEREBkBERAxAREQKQEREAAAAAAAAAAAAAAAAAAAAAAAAAAAAQEREAEBERApAREQOQERECkBERAAAAAAAAAAAAAAAAAAAAAAAAAAAAEBERABAREQKQEREDkBERAqvmxEC9MAAVvTAAPzxvQD/8r4A//K+AP/xvQD/6bgD/+e3BP/puAP8+MUA6vC8ACT/vwAAAAAAAAAAAAAAAAAAQEREAEBERApAREQOQERECkBERAAAAAAAAAAAAAAAAAAAAAAAAAAAAEBERABAREQKQEREDkBERApAREQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAQEREAAAAAAAAAAAAAAAAAEBERABAREQIQEREDEBERAhAREQAAAAAAAAAAAAAAAAAAAAAAAAAAABAREQAQERECkBERAxAREQIQEREAAAAAAAAAAAAAAAAAAAAAAAAAAAAQEREAEBERApAREQMQERECNimDQL2wQB09cEA/PK+AP/yvgD/8r4A//G9AP/quQL/7bwD+NyvCUSmqCAAAAAAAAAAAAAAAAAAAAAAAAAAAABAREQAQERECkBERAxAREQIQEREAAAA
AAAAAAAAAAAAAAAAAAAAAAAAQEREAEBERApAREQMQERECEBERAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAQEREAEBERAZAREQIQEREBgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAEBERABAREQGQERECEBERAYAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABAREQAQEREBkBERAhAREQG5K8JAvjEAI70wAD88r4A//K+AP/xvgD/9cEA/Oe2BWZcWjoKQEREBgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAEBERABAREQGQERECEBERAYAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABAREQAQEREBkBERAhAREQGAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAQEREBEBERAZAREQEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAEBERARAREQGQEREBAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABAREQEQEREBkpOQATgtQMC+MMArPO/APzyvgD/9MAA/PfEAITRnwcCQEREBEBERAZAREQEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAEBERARAREQGQEREBAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABAREQEQEREBkBERAQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABAREQCQEREAkBERAIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAQEREAkBERAJAREQCAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAEBERAJAREQCdGAxAuq/AAr2wgDG9L8A/PfDAKj/mwACAAAAAAAAAABAREQCQEREAkBERAIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAQEREAkBE
RAJAREQCAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAEBERAJAREQCQEREAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAEBERABAREQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABAREQAQEREAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAQEREAEBERAD1uQMA77sAFvfCAKzyuwAKqqoAAAAAAAAAAAAAAAAAAEBERABAREQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABAREQAQEREAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAQEREAEBERAAAAAAAAAAAAAAAAAAAAAAA/////3/////////+P/////////wf////////8A/////////gB////////8AB////////gAD///////8AAH///////AAAP//////4AAAf//////AAAA//////4AAAA//////AAAAB/////wAAAAD////+AAfgAH////wAD/AAP///+AAP+AAP///wAB/4AAf//8AAH/gAA///gAAP+AAB//8AAA/wAAD//gAAD/gAAH/8AAAb3AAAH/AAPHAHAAAP4AB/4AOfAAfAAP/AAP+AA4AA/4AAf8ABAAH/gAB/wAGAAP+AAH/AAcAA/4AAf8AD4AB/AAB/wAfwAD8AAH+AD/gAAwAA3gAf/AABgAGAAD/+AADAAwAAf/4AAGAGAAB//wAAMAwAAP//gAAYGAAB///AAB/wAAP//+AAD/AAB///8AAP8AAP///4AB/4AA////wAH/gAH////gAf+AA////+AA/wAH////8AB/AA/////4ADwAH/////wAAAA//////gAAAD//////AAAAf/////+AAAD//////8AAAf//////wAAD///////gAAf///////AAB///////+AAP//////
/8AB////////4AP////////wB/////////gP////////+B/////////8H/////////4//////////3////8="
$iconBytes                       = [Convert]::FromBase64String($iconBase64)
$stream                          = New-Object IO.MemoryStream($iconBytes, 0, $iconBytes.Length)
$stream.Write($iconBytes, 0, $iconBytes.Length);
$iconImage                       = [System.Drawing.Image]::FromStream($stream, $true)
$FormAdExplorer.Icon             = [System.Drawing.Icon]::FromHandle((New-Object System.Drawing.Bitmap -Argument $stream).GetHIcon())
$FormAdExplorer.KeyPreview       = $true
$FormAdExplorer.Add_KeyDown({
    if ($_.KeyCode -eq "Enter") {
        $ButtonConfirm.PerformClick()
    }

    elseif ($_.KeyCode -eq "Escape") {
        $ButtonQuit.PerformClick()
    }
})

$ButtonQuit                      = New-Object system.Windows.Forms.Button
$ButtonQuit.text                 = "Sair"
$ButtonQuit.width                = 60
$ButtonQuit.height               = 29
$ButtonQuit.location             = New-Object System.Drawing.Point(433,213)
$ButtonQuit.Font                 = New-Object System.Drawing.Font('Rage',10)
$ButtonQuit.ForeColor            = [System.Drawing.ColorTranslator]::FromHtml("#080808")
$ButtonQuit.BackColor            = [System.Drawing.ColorTranslator]::FromHtml("#b0b0b0")
$ButtonQuit.DialogResult         = [System.Windows.Forms.DialogResult]::Cancel

$ButtonConfirm                   = New-Object system.Windows.Forms.Button
$ButtonConfirm.text              = "Confirmar"
$ButtonConfirm.width             = 78
$ButtonConfirm.height            = 30
$ButtonConfirm.location          = New-Object System.Drawing.Point(426,158)
$ButtonConfirm.Font              = New-Object System.Drawing.Font('Rage',10)
$ButtonConfirm.ForeColor         = [System.Drawing.ColorTranslator]::FromHtml("#0d0d0d")
$ButtonConfirm.BackColor         = [System.Drawing.ColorTranslator]::FromHtml("#b0b0b0")

$GroupboxAction                  = New-Object system.Windows.Forms.Groupbox
$GroupboxAction.height           = 121
$GroupboxAction.width            = 183
$GroupboxAction.text             = "Exporte em"
$GroupboxAction.location         = New-Object System.Drawing.Point(201,125)

$TextBoxID                       = New-Object system.Windows.Forms.TextBox
$TextBoxID.TabIndex              = 0
$TextBoxID.multiline             = $false
$TextBoxID.width                 = 86
$TextBoxID.height                = 20
$TextBoxID.location              = New-Object System.Drawing.Point(61,180)
$TextBoxID.Font                  = New-Object System.Drawing.Font('Microsoft Sans Serif',10)
$TextBoxID.text                  = $user
$FormAdExplorer.Controls.Add($TextBoxID)

$RadioButtonTxt                  = New-Object system.Windows.Forms.RadioButton
$RadioButtonTxt.text             = ".txt"
$RadioButtonTxt.AutoSize         = $true
$RadioButtonTxt.width            = 104
$RadioButtonTxt.height           = 20
$RadioButtonTxt.location         = New-Object System.Drawing.Point(55,25)
$RadioButtonTxt.Font             = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$RadioButtonCsv                  = New-Object system.Windows.Forms.RadioButton
$RadioButtonCsv.text             = ".xlsx"
$RadioButtonCsv.AutoSize         = $true
$RadioButtonCsv.width            = 104
$RadioButtonCsv.height           = 20
$RadioButtonCsv.location         = New-Object System.Drawing.Point(55,58)
$RadioButtonCsv.Font             = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$RadioButtonGrid                 = New-Object system.Windows.Forms.RadioButton
$RadioButtonGrid.text            = "Tabela interativa"
$RadioButtonGrid.AutoSize        = $true
$RadioButtonGrid.width           = 104
$RadioButtonGrid.height          = 20
$RadioButtonGrid.location        = New-Object System.Drawing.Point(55,90)
$RadioButtonGrid.Font            = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$LabelExplorer                   = New-Object system.Windows.Forms.Label
$LabelExplorer.text              = "AD EXPLORER"
$LabelExplorer.AutoSize          = $true
$LabelExplorer.width             = 25
$LabelExplorer.height            = 10
$LabelExplorer.location          = New-Object System.Drawing.Point(211,30)
$LabelExplorer.Font              = New-Object System.Drawing.Font('Microsoft YaHei UI',12)

$LabelUser                       = New-Object system.Windows.Forms.Label
$LabelUser.text                  = "Digite a chave do usuário"
$LabelUser.AutoSize              = $true
$LabelUser.width                 = 25
$LabelUser.height                = 10
$LabelUser.location              = New-Object System.Drawing.Point(26,150)
$LabelUser.Font                  = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$LabelInfo                       = New-Object system.Windows.Forms.Label
$LabelInfo.text                  = "Este programa realiza a busca apenas de grupos e usuários do domínio local.`nOs arquivos de relatório são salvos na pasta Documentos (ou Documents)."
$LabelInfo.AutoSize              = $true
$LabelInfo.width                 = 25
$LabelInfo.height                = 10
$LabelInfo.location              = New-Object System.Drawing.Point(45,65)
$LabelInfo.Font                  = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$LabelProcessing                 = New-Object system.Windows.Forms.Label
$LabelProcessing.text            = "Gerando visualização de dados. Aguarde..."
$LabelProcessing.AutoSize        = $true
$LabelProcessing.width           = 25
$LabelProcessing.height          = 10
$LabelProcessing.location        = New-Object System.Drawing.Point(134,255)
$LabelProcessing.Font            = New-Object System.Drawing.Font('Microsoft Sans Serif',10,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
$LabelProcessing.ForeColor       = [System.Drawing.ColorTranslator]::FromHtml("#0c5151")


$LabelError                      = New-Object system.Windows.Forms.Label
$LabelError.text                 = "Usuário não encontrado no domínio."
$LabelError.AutoSize             = $true
$LabelError.width                = 25
$LabelError.height               = 10
$LabelError.location             = New-Object System.Drawing.Point(134,255)
$LabelError.Font                 = New-Object System.Drawing.Font('Microsoft Sans Serif',10,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold -bor [System.Drawing.FontStyle]::Underline))
$LabelError.ForeColor            = [System.Drawing.ColorTranslator]::FromHtml("#cb0909")

$MainStripMenu = New-Object System.Windows.Forms.MenuStrip
$MainStripMenu.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#c0c0c0")

$AboutStripMenuItem             = New-Object System.Windows.Forms.ToolStripMenuItem
$AboutStripMenuItem.Name        = "AboutStripMenuItem"
$AboutStripMenuItem.Size        = New-Object System.Drawing.Size(51, 20)
$AboutStripMenuItem.Text        = "Sobre"

$MainStripMenu.Items.AddRange(@($AboutStripMenuItem))
$FormAdExplorer.controls.AddRange(@($ButtonQuit,$ButtonConfirm,$GroupboxAction,$TextBoxID,$LabelExplorer,$LabelUser,$LabelInfo, $LabelProcessing, $LabelError, $MainStripMenu))
$GroupboxAction.controls.AddRange(@($RadioButtonTxt,$RadioButtonCsv,$RadioButtonGrid))

$LabelProcessing.Visible = $false
$LabelError.Visible = $false


$FormAdExplorer.Add_Shown({
    $FormAdExplorer.Activate()
    HideConsole
})

$ButtonConfirm.Add_Click({
    if ($RadioButtonTxt.Checked -and $TextBoxID.text){
        txt
     }
 
     elseif ($RadioButtonCsv.Checked -and $TextBoxID.text) {
         csv
     }
 
     elseif ($RadioButtonGrid.Checked -and $TextBoxID.text) {
         grid
     }
 
     elseif ($RadioButtonTxt.Checked -or $RadioButtonCsv.Checked -or $RadioButtonGrid.Checked -and $TextBoxID.text -eq "") {
         $aviso = New-Object -ComObject Wscript.Shell ; $aviso.popup("Informe o ID do usuário.",0,"Atenção!",0x0)
     }
 
     else { $aviso = New-Object -ComObject Wscript.Shell ; $aviso.popup("Selecione uma opção válida.",0,"Atenção!",0x0) } 
})


$AboutStripMenuItem.Add_Click({
    $sobre = New-Object -ComObject Wscript.Shell
    $sobre.Popup("Versão: 2.2`nDesenvolvido por: gitlab.com/vpess",0,"Sobre AD Explorer",0x0)
})


Add-Type -Name Window -Namespace Console -MemberDefinition '
[DllImport("Kernel32.dll")]
public static extern IntPtr GetConsoleWindow();

[DllImport("user32.dll")]
public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
'

function HideConsole {
    $consolePtr = [Console.Window]::GetConsoleWindow()
    [Console.Window]::ShowWindow($consolePtr, 0)
}

function AD_Explorer {

if (Get-Module -ListAvailable -Name ActiveDirectory) {
    Import-Module ActiveDirectory
    do {Start-AD} until (Import-Module ActiveDirectory == $true)
    $MainStripMenu.ShowDialog()
    $FormAdExplorer.ShowDialog()
    $GroupboxAction.ShowDialog()
} 
else {
    $warning_net = New-Object -ComObject Wscript.Shell
    $warning_net.Popup("O computador não possui o módulo Active Directory para Powershell, ou não possui o Active Directory instalado.",0,"Aviso!",0x0)
}
}
AD_Explorer
