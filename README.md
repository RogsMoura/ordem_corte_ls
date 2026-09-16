# Gerador de Ordem de Corte - Ls Pijamas

Aplicativo desktop em Python para transformar pedidos em PDF em uma ordem de
corte organizada em Excel, com tabelas de CAMISA e CALCA em uma única aba.

## Recursos

- Interface leve em Tkinter com a identidade visual da loja.
- Leitura de PDFs com um ou vários pedidos de clientes.
- Separação automática por número do pedido.
- Geração de um `.xlsx` com o mesmo nome do PDF selecionado.
- Aba única `Ordem de Corte`.
- Fundo alternado por pedido para facilitar a leitura na produção.
- Ícone e logo incorporados no executável Windows.

## Requisitos

- Windows.
- Python 3.10 ou superior para executar o código-fonte.
- `tkinter`, normalmente incluído na instalação oficial do Python.
- `pdfplumber` e `openpyxl`, listados em `requirements.txt`.

## Instalação para desenvolvimento

No PowerShell, dentro da pasta do projeto:

```powershell
py -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
pip install -r requirements.txt
```

## Executar pelo código-fonte

```powershell
python gerador_ordem_corte.py
```

Selecione o PDF e clique em **Gerar Planilha de Corte**. O Excel será salvo na
mesma pasta do PDF, trocando apenas a extensão. Por exemplo:

```text
pedido.pdf -> pedido.xlsx
```

O programa foi desenvolvido para PDFs com texto selecionável. PDFs formados
somente por imagens precisam passar por OCR antes da leitura.

## Regras de produção

- Sem tecido especificado no produto, o padrão é `Gabardine`.
- Tecidos reconhecidos: `Gabardine`, `Oxfordine`, `Cedromix`, `Melange` e `Premium`.
- Camisas estampadas mostram `Estampado`; calças estampadas usam `Gabardine`.
- `Classic` só é usado quando existe `Tamanho Camisa Classic:`. O gênero fica `Feminino Classic` ou `Masculino Classic`.
- Sem Classic, o gênero fica somente `Feminino` ou `Masculino`.
- `FRISOS / DETALHES` recebe a cor de `COR DO FRISO` quando os frisos forem `Sim`.
- Itens com `Cores Smile:` recebem `Smile`; os demais recebem `X`.
- Friso e Smile juntos aparecem separados por ` | `.
- `OBSERVAÇÕES` mostra o campo `Observações:` ou `X` quando ausente.
- A grade de CALCA contém pedido, cor, tamanho, barra e tecido.
- Em pedidos de futebol, o time é combinado com a cor nas grades, como `Palmeiras - Verde musgo`.
- `Conjunto liso + 1 Estampa` gera um conjunto liso e uma camisa estampada sem calça.
- Informações de bordado não são exibidas na ordem de corte.

## Criar o executável Windows

Instale o PyInstaller:

```powershell
pip install pyinstaller
```

Gere o executável com a logo atual:

```powershell
pyinstaller --noconfirm --clean --onefile --windowed `
	--name "Ordem Corte LS" `
	--icon=logo-ls.ico `
	--add-data "logo-ls.png;." `
	--add-data "logo-ls.ico;." `
	gerador_ordem_corte.py
```

O arquivo será criado em `dist\Ordem Corte LS.exe`. Ele já contém Python,
dependências, logo e ícone. Para levar a outro computador, copie somente o
`.exe`; não é necessário instalar Python ou bibliotecas.

## Arquivos necessários no projeto

- `gerador_ordem_corte.py`: código da aplicação.
- `logo-ls.png`: logo exibida na janela.
- `logo-ls.ico`: ícone usado pelo executável.
- `requirements.txt`: dependências Python.
- `README.md`: documentação.

Arquivos de teste, PDFs de clientes, planilhas geradas, ambientes virtuais,
cache e pastas de build não devem ser publicados no repositório.