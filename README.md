# Simulador Web MVP

Primeira versao web local do Simulador de Cartas.

O app agora usa um banco local SQLite quando existe base baixada da Play. A planilha Excel fica apenas como fallback inicial.

## Como rodar

Abra o PowerShell nesta pasta e execute:

```powershell
python app.py
```

Depois acesse:

```text
http://127.0.0.1:8765
```

## Usar uma planilha especifica

O app tenta encontrar automaticamente a planilha mais recente na pasta `Play Consorcios`.
Se quiser apontar manualmente:

```powershell
python app.py --workbook "C:\caminho\para\Simulador de Cartas_v2.2.xlsm"
```

## Testar a leitura da planilha

```powershell
python app.py --check
```

Depois de usar o botao `Atualizar Dados`, o `--check` deve mostrar a fonte como `Banco local da Play`.

## Atualizar dados da Play

No topo da tela, clique em `Atualizar Dados`.

Senha padrao local para teste:

```text
play123
```

Para trocar a senha, defina a variavel de ambiente:

```powershell
$env:SIMULADOR_UPDATE_PASSWORD="sua-senha"
python app.py
```

O app acessa a tabela publica em:

```text
https://playcontempladas.com.br/?segmento=todos
```

E grava os dados em:

```text
data/cartas.sqlite3
```

Ao abrir novamente, o app carrega primeiro esse banco local. Se o banco nao existir, tenta usar a planilha Excel.

## O que esta nesta versao

- Leitura da aba `Dados`.
- Filtro por busca, categoria, administradora, credito, entrada, parcela e custo.
- Filtro por multiplas administradoras em dropdown com checkboxes.
- Filtros visuais aplicados sobre o conjunto atual, sem refazer a busca principal.
- Motor Python inicial para gerar propostas por credito alvo, categoria, entrada maxima, parcela maxima e regra de saldo devedor.
- Geracao de cartas individuais e pacotes por administradora seguindo as margens principais do VBA.
- Motor de propostas otimizado com poda de combinacoes impossiveis e resposta mais leve para cenarios de credito alto.
- Busca com teto de entrada/parcela prioriza cartas mais eficientes para evitar travamento em metas altas.
- Custo calculado somente no total da proposta, sem custo por carta individual.
- Tabela sem coluna visual de status.
- Renderizacao da tabela em lote, com limite visual de 5.000 linhas por filtro para manter a tela rapida.
- Selecao de cartas filtradas ou da proposta inteira quando a tabela estiver mostrando propostas calculadas.
- Geracao de texto de proposta com totais e faixas de parcelas.
- Texto de proposta mostra `Parcelas` como titulo simples, sem listar todos os prazos na mesma linha.
- Botao para atualizar dados diretamente do site da Play.
- Botao `Atualizar Dados` protegido por senha simples para testes.
- Favicon com a logo LJ e titulo `Simulador de Cartas` no navegador.
- Botao para recarregar a base sem reiniciar o app.

## Proximos passos sugeridos

- Salvar propostas geradas em banco SQLite.
- Criar login/usuarios.
- Publicar em nuvem com atualizacao agendada da Play.
- Trocar a senha simples por login quando sair do teste interno.
- Conferir a calculadora Python contra novos cenarios reais da versao 2.2.
- Exportar proposta em PDF.
