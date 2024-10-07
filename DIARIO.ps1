# Defina o caminho base
$base = "G:\Meu Drive\DIARIO"

# Mostre os meses
$meses = @("JANEIRO", "FEVEREIRO", "MARCO", "ABRIL", "MAIO", "JUNHO", "JULHO", "AGOSTO", "SETEMBRO", "OUTUBRO", "NOVEMBRO", "DEZEMBRO")

# Obtém a data atual
$dataAtual = Get-Date
$mesInt = $dataAtual.Month  # Obtém o mês atual (1-12)
$ano = $dataAtual.Year       # Obtém o ano atual

# Defina as pastas do ano e do mês
$pastaAno = Join-Path -Path $base -ChildPath $ano
$nomeMes = $meses[$mesInt - 1]
$pastaMes = Join-Path -Path $pastaAno -ChildPath $nomeMes

# Verifique e crie a pasta do ano
if (-not (Test-Path -Path $pastaAno)) {
    New-Item -ItemType Directory -Path $pastaAno | Out-Null
    Write-Host "Pasta do ano criada: $pastaAno"
}

# Verifique e crie a pasta do mês
if (-not (Test-Path -Path $pastaMes)) {
    New-Item -ItemType Directory -Path $pastaMes | Out-Null
    Write-Host "Pasta do mês criada: $pastaMes"
} else {
    Write-Host "A pasta do mês $nomeMes do ano $ano já existe."
    Read-Host "Pressione Enter para continuar..."  # Pausa antes de fechar
    return  # Encerra o script se a pasta já existir
}

# Defina o número de dias no mês
$dias = [DateTime]::DaysInMonth($ano, $mesInt)

# Crie os arquivos para cada dia do mês
for ($dia = 1; $dia -le $dias; $dia++) {
    $diaFormatado = "{0:D2}" -f $dia
    $anoCur = $ano.Substring(2, 2)  # Obtém os últimos dois dígitos do ano
    $arquivo = Join-Path -Path $pastaMes -ChildPath "$diaFormatado $mes $anoCur.doc"

    # Cria o arquivo
    New-Item -ItemType File -Path $arquivo -Force | Out-Null
    Write-Host "Arquivo criado: $arquivo"
}

Write-Host "Todos os arquivos foram criados na pasta $pastaMes."
Read-Host "Pressione Enter para continuar..."  # Pausa antes de fechar
