$Host.UI.RawUI.WindowTitle = "Active Directory Explorer"
$ErrorActionPreference = "SilentlyContinue"

Write-Host "Esse script efetua a busca de dados apenas dos usuários e grupos do domínio petrobras.biz.`n" -ForegroundColor Red

Write-Host "Carregando o módulo do Active Directory...`n" ; Import-Module ActiveDirectory -ErrorVariable erro

if ($erro) {
    Write-Warning -Message "`nNão foi possível carregar o Active Directory.`n"
}

else{

function txt {

    $user = Read-Host "`nDigite a chave do usuário"
    Write-Host "`nExportando dados. Aguarde..." -ForegroundColor Green
    $username = Get-ADUser $user -Properties Name, DisplayName -ErrorVariable erro2 | Select-Object Name, Displayname
    $file = (Get-ADUser $user -Properties MemberOf -ErrorVariable erro2).MemberOf | Get-ADGroup -Properties Name, Description | Select-Object Name, Description

    if ($erro2) {
        Write-Warning -Message "`nChave não encontrada no AD. Tente novamente.`n"
        Start-Sleep -s 2
        return txt
    }
    
    else {
        $date = get-date -format "dddd, dd/MM/yyyy, HH:mm:ss"
        $username | Out-File "$Env:USERPROFILE\desktop\AD_$user.txt"
        $file | Add-Content "$Env:USERPROFILE\desktop\AD_$user.txt"
        (Get-Content "$Env:USERPROFILE\desktop\AD_$user.txt" -Raw) -replace '@{', '' | Out-file "$Env:USERPROFILE\desktop\AD_$user.txt"
        (Get-Content "$Env:USERPROFILE\desktop\AD_$user.txt" -Raw) -replace '}', '' | Out-file "$Env:USERPROFILE\desktop\AD_$user.txt"
        (Get-Content "$Env:USERPROFILE\desktop\AD_$user.txt" -Raw) -replace ';', ' <—————> ' | Out-file "$Env:USERPROFILE\desktop\AD_$user.txt" 
        $date | Add-Content "$Env:USERPROFILE\desktop\AD_$user.txt" 
        Write-Host "`nArquivo exportado para a Área de Trabalho." -ForegroundColor DarkCyan
        Invoke-Expression "$Env:USERPROFILE\desktop\AD_$user.txt"
        Start-Sleep -s 1
        return selec
    }

}

function grid {

    $user = Read-Host "`nDigite a chave do usuário"
    Write-Host "`nGerando tabela. Aguarde..." -ForegroundColor Green
    $file = (Get-ADUser $user -Properties MemberOf -ErrorVariable erro2).MemberOf | Get-ADGroup -Properties Name, Description | Select-Object Name, Description

    if ($erro2) {
        Write-Warning -Message "`nChave não encontrada no AD. Tente novamente.`n"
        Start-Sleep -s 2
        return grid
    }
    
    else {
        $file | Out-gridview -Title "Grupos de Acesso de $user"
        Write-Host "`nTabela gerada com sucesso." -ForegroundColor DarkCyan
        Start-Sleep -s 1
        return selec
    }
}

function csv {

    $user = Read-Host "`nDigite a chave do usuário"
    Write-Host "`nExportando dados. Aguarde..." -ForegroundColor Green
    $username = Get-ADUser $user -Properties Name, DisplayName -ErrorVariable erro2 | Select-Object Name, Displayname
    $file = (Get-ADUser $user -Properties MemberOf -ErrorVariable erro2).MemberOf | Get-ADGroup -Properties Name, Description | Select-Object Name, Description

    if ($erro2) {
        Write-Warning -Message "`nChave não encontrada no AD. Tente novamente.`n"
        Start-Sleep -s 2
        return csv
    }

    else {
        $file | Export-CSV -Delimiter ';' -Path "$Env:USERPROFILE\desktop\AD_$user.csv"
        $date = get-date -format "dddd, dd/MM/yyyy, HH:mm:ss" ; $date | Add-Content "$Env:USERPROFILE\desktop\AD_$user.csv"
        $username | Add-Content "$Env:USERPROFILE\desktop\AD_$user.csv"
        Write-Host "`nArquivo exportado para a Área de Trabalho." -ForegroundColor DarkCyan
        Invoke-Expression "$Env:USERPROFILE\desktop\AD_$user.csv"
        Start-Sleep -s 1
        return selec
    }
    
}

function selec {

param (
[string]$Titulo = 'Menu'
)

Write-Host "`n============================ AD Explorer ============================`n" -ForegroundColor DarkYellow

Write-Host "	[1] exportar os grupos do AD em .txt" -ForegroundColor DarkGray
Write-Host "	[2] exportar os grupos do AD em .csv" -ForegroundColor DarkGreen
Write-Host "	[3] exibir os grupos em tabela interativa" -ForegroundColor DarkCyan 
Write-Host "	[q] para fechar o script" -ForegroundColor Red

Write-Host "`n=====================================================================" -ForegroundColor DarkYellow

 $selection = Read-Host "`nSelecione uma das opções acima"
 switch ($selection)
 {

     '1' {txt}

     '2' {csv}

     '3' {grid}

     'q'{
        Write-Host "`nSaindo..." -ForegroundColor Red
        Start-sleep -s 2
        return
     }

     default {

        if ($selection -ige 4 -or $selection -ne 'q'){
             Write-Host "`n>>> Selecione apenas opções que estejam no menu!" -ForegroundColor Red
             Start-Sleep -s 2
             return selec
             }
        }
 }

}

selec

}

<#iae
Quem fez foi BJBD. 
Ele fez com o objetivo de facilitar a busca de grupos de acesso a pastas de rede.
Com um arquivo que contém as informações do AD, podemos simplesmente buscar o grupo de acesso com um CTRL+F.
#>