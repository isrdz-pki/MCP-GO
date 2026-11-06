# Agente GO

Aplicativo Windows com janela para o ERP Gestão Online (GO): baixa o relatório de vendas e consulta pedidos pela API, sem navegador. É distribuído como executável, então quem usa não precisa instalar Python nem bibliotecas.

## Funcionalidades

| Aba | O que faz |
|---|---|
| Token (topo) | Colar o `access_token` do GO, salvar no Cofre de Credenciais do Windows, testar e acompanhar a validade (12 h) |
| Relatório de vendas | Relatório "CSV - Vendas por unidade" (report 44) por período e unidade, com atalhos de data |
| Pedido | Detalhes de um pedido e download dos anexos (NF/DANFE, XML) |
| CPF / CNPJ | Compras e dados cadastrais de um cliente, NF de cada compra, exportação para .xlsx |
| Lote (planilha) | Consulta em massa a partir de uma planilha com números de pedido |

Também roda sem janela, para o Agendador de Tarefas, usando o token salvo:

```
"Agente GO.exe" --relatorio 01/10/2026 31/10/2026 --unidade 52 [--saida arquivo.xlsx]
```

O log fica em `%LOCALAPPDATA%\Agente GO\agente_go.log`. Códigos de saída: 0 = ok, 1 = erro, 2 = sem token, 3 = token vencido.

## Para quem só vai usar

Baixe o `.zip` na página de **Releases**, extraia e abra o `Agente GO.exe`. O passo a passo completo, incluindo como pegar o token, está em [LEIA-ME.txt](LEIA-ME.txt).

## Autenticação

O GO não oferece renovação automática do token: o `grant_type=refresh_token` responde `unsupported_grant_type`, e o login por senha exige reCAPTCHA. Por isso, cada usuário cola o próprio `access_token`, que pega no GO aberto no navegador. O token vale cerca de 12 horas e fica guardado só no cofre do Windows daquele usuário (keyring, serviço `agente-go`). Nenhum token ou segredo fica no código.

## Desenvolvimento

Requer Python 3.13 no Windows.

```
py -m pip install -r requirements.txt
py src\app.py                 # roda a partir do código
gerar_executavel.bat          # gera dist\Agente GO\ (PyInstaller, modo pasta)
py empacotar.py               # gera "Agente GO vX.Y.zip" na pasta acima do projeto
```

### Estrutura

```
src/app.py           janela principal (abas, token, relatório, modo sem janela)
src/relatorio.py     relatório de vendas e token (sem interface)
src/consulta_gui.py  consultas de pedido, CPF/CNPJ, NF e lote
src/icone.ico        ícone do aplicativo
```

O executável é gerado em modo pasta (onedir), e não como arquivo único. Assim ele abre mais rápido e é menos barrado por antivírus.
