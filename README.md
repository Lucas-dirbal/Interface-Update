# Interface Update

Ferramenta em Python para analisar a base atual da Interface, isolar o banco para atualizacao e executar os instaladores mais novos em ordem mensal, com seguranca e rastreabilidade.

## EXE pronto

Executavel gerado:

- `dist\Interface Update.exe`

Ao rodar como EXE, ele tenta localizar a pasta de instaladores ao lado do executavel.

## O que ele faz

- le a versao atual da base pelo campo `T000_TS_COMPILACAO`;
- encontra os instaladores na pasta `SETUP INTERFACE ATUALIZACAO`;
- monta a fila de atualizacao mes a mes;
- renomeia o banco para `BASE EM ATUALIZACAO.FDB` antes de atualizar;
- registra em `mapeamento_base_em_atualizacao.txt` qual era o nome original antes da troca;
- troca automaticamente o nome da base nos arquivos de configuracao;
- automatiza o instalador com confirmacoes, conclusao e abertura da Interface;
- espera a Interface terminar a rotina pos-instalacao e clica em `Cancelar` na tela de login;
- restaura automaticamente o nome original da base ao final;
- aplica automaticamente a correcao da grid de localizacao de produtos ao final;
- cria backup dos `.ini` antes de qualquer alteracao;
- permite restaurar o nome original depois.

## Inicio rapido

### 1. Instale a dependencia

```powershell
pip install fdb psutil pywinauto pywin32
```

### 2. Abra a interface grafica

```powershell
python Leitor.py
```

Ou abra o executavel:

```powershell
.\dist\Interface Update.exe
```

### 3. Siga o fluxo recomendado

1. Abra o programa como Administrador.
2. Clique em `Analisar`.
3. Confira a versao atual e a fila de setups.
4. Clique em `Executar Atualizacao`.
5. O app faz sozinho: preparo da base, instalador, confirmacoes, abertura da Interface e fechamento para seguir para a proxima versao.
6. Se quiser voltar ao nome original da base ao final, clique em `Restaurar Nome Original`.

## Pastas padrao

O programa usa estes caminhos por padrao:

- Interface: `C:\interface`
- Setups: `.\SETUP INTERFACE ATUALIZACAO`

Exemplo de instalador esperado:

```text
Interface 1.0 170420261020.exe
```

## Requisitos

- Python 3
- biblioteca `fdb`
- `psutil`
- `pywinauto`
- `pywin32`
- `tkinter` habilitado no Python
- `fbclient.dll` dentro da pasta da Interface

## Interface grafica

Ao abrir com `python Leitor.py`, a tela mostra:

- pasta da Interface;
- pasta dos setups;
- banco principal encontrado;
- versao atual do banco;
- tabela onde a compilacao foi localizada;
- indicador se a base ja esta isolada;
- fila de atualizacao mes a mes;
- log completo da execucao.

Para a automacao completa dos instaladores, o app deve rodar com permissao de Administrador.

## Fluxo automatico da atualizacao

Quando voce clicar em `Executar Atualizacao`, o programa faz:

1. prepara a base automaticamente;
2. abre o setup da versao;
3. confirma o idioma;
4. envia 5 vezes `Enter` para avancar o instalador;
5. tenta garantir a opcao de executar a Interface ao final;
6. conclui o setup;
7. espera a `InterfaceSi.exe` abrir;
8. aguarda a tela de acesso aparecer;
9. clica em `Cancelar` na tela de login da Interface;
10. confirma que a Interface foi encerrada por completo;
11. rele a versao do banco e passa para a proxima versao da fila;
12. ao final da fila, restaura o nome original da base;
13. aplica a correcao da grid de localizacao de produtos.

## Fluxo de seguranca

Antes da atualizacao, o programa pode preparar o ambiente automaticamente:

1. encontra a base principal;
2. prioriza a base `.FDB` referenciada nos arquivos `.ini`;
3. registra o nome original em `mapeamento_base_em_atualizacao.txt`;
4. renomeia a base encontrada para `BASE EM ATUALIZACAO.FDB`;
5. cria backup dos arquivos `.ini`;
6. atualiza os caminhos que apontam para o banco;
7. valida se o computador continua acessando a base corretamente.

Arquivos normalmente afetados:

- `C:\interface\Bd\Config.ini`
- `C:\interface\Bd\Config_Inspecoes.ini`
- `C:\interface\dbxconnections.ini`
- `C:\interface\Servidor ILoja\ServidorILoja.ini`

Os backups sao salvos com extensao `.bak` e timestamp no mesmo local do arquivo original.

## Como a fila de atualizacao e montada

O programa analisa todos os setups encontrados e escolhe apenas o mais recente de cada mes acima da versao atual da base.

Exemplo:

- se houver dois setups em fevereiro, ele usa apenas o mais novo de fevereiro;
- depois faz o mesmo para marco, abril e assim por diante;
- apos cada instalador, ele rele a versao do banco e recalcula a fila.

Isso evita executar etapas desnecessarias caso um setup ja atualize a base para uma versao mais avancada.

## Comandos disponiveis

### Abrir a interface

```powershell
python Leitor.py
```

### Ver status da base

```powershell
python Leitor.py status
```

### Preparar a base

```powershell
python Leitor.py preparar
```

### Executar a atualizacao

```powershell
python Leitor.py atualizar
```

### Restaurar o nome original da base

```powershell
python Leitor.py restaurar
```

## Parametros uteis

### Usar outra pasta da Interface

```powershell
python Leitor.py status --interface "D:\Interface"
```

### Usar outra pasta de instaladores

```powershell
python Leitor.py status --setups "D:\Instaladores\SETUP INTERFACE ATUALIZACAO"
```

### Passar argumentos para os instaladores

```powershell
python Leitor.py atualizar --args-installer "/VERYSILENT /SUPPRESSMSGBOXES"
```

### Atualizar sem preparar antes

Use apenas quando a base ja estiver renomeada e os arquivos de configuracao ja tiverem sido ajustados.

```powershell
python Leitor.py atualizar --nao-preparar
```

## Saida do comando `status`

O comando `status` informa:

- pasta analisada;
- banco principal encontrado;
- origem da configuracao utilizada;
- tabela onde esta `T000_TS_COMPILACAO`;
- data e hora da versao atual;
- se a base esta isolada;
- fila de atualizacao pendente.

## Cuidados importantes

- nao rode a atualizacao com usuarios trabalhando no sistema;
- nao apague os arquivos `.bak` antes de validar tudo;
- se ja existir `BASE EM ATUALIZACAO.FDB`, a preparacao sera interrompida;
- o programa nao força instalacao silenciosa por padrao;
- se algum setup exigir interacao, acompanhe a execucao.

## Exemplo de uso completo no terminal

```powershell
python Leitor.py status
python Leitor.py preparar
python Leitor.py atualizar
python Leitor.py restaurar
```

## Arquivo principal

Codigo-fonte principal:

- `Leitor.py`

## Resumo

Se a ideia for usar do jeito mais simples, basta executar:

```powershell
python Leitor.py
```

Depois disso, o fluxo normal e:

`Analisar` -> `Preparar Base` -> `Executar Atualizacao` -> `Restaurar Nome Original`
