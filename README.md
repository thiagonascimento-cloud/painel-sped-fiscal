# Painel SPED Fiscal – Compras e Vendas

Ferramenta para ler arquivos **SPED Fiscal (EFD ICMS/IPI)** e mostrar as compras e vendas do período: totais por natureza, CFOPs, clientes e fornecedores, movimento por dia, produtos das notas e um relatório resumido para impressão.

O conteúdo do SPED é processado **só no computador de quem usa**. Nada é enviado para servidores.

## Formas de usar

### 1. Página (`Painel SPED.html`)

Abra o arquivo no navegador (Chrome ou Edge).

1. Abra o `.txt` do SPED, copie tudo (Ctrl+A, Ctrl+C) e cole no campo **"Cole o SPED aqui"**. Também dá para arrastar um ou mais arquivos `.txt` para a página.
2. Veja as abas **Painel**, **Detalhado**, **Movimento Diário** e **Tabela CFOP**.
3. Use **Imprimir relatório** para um resumo de uma página A4, ou **Baixar Excel** para a planilha.

A página precisa de internet só para as fontes e para gerar o Excel.

### 2. Planilha pelo script (`Gerar Painel SPED.bat`)

Requer [Python 3](https://www.python.org/downloads/). A biblioteca `openpyxl` é instalada automaticamente na primeira execução.

- Arraste a **pasta do mês** ou os arquivos SPED `.txt` para cima do `Gerar Painel SPED.bat`, ou
- dê dois cliques nele e escolha os arquivos na janela.

A planilha é salva na mesma pasta dos SPEDs. Pela linha de comando:

```bash
python sped_painel.py "caminho/da/pasta" -o "saida.xlsx"
```

## Planilha gerada

| Aba | Conteúdo |
|---|---|
| **Resumo** | Dados da empresa e valores consolidados; aparece só o que tem valor |
| **Notas fiscais** | Uma linha por nota x CFOP, com cliente/fornecedor, CFOP, natureza, valores e chave de acesso |
| **Produtos** | Itens das notas (registro C170): código, descrição, NCM, quantidade e valores |
| **Clientes e fornecedores** | Totais por participante |

Colunas e linhas sem valor não são geradas, e os zeros ficam em branco.

> O SPED traz os itens das notas (C170) apenas para os documentos recebidos de terceiros. Nas NF-e emitidas pela própria empresa os itens não são informados, por isso as vendas não aparecem na aba Produtos.

## Registros lidos

`0000` (empresa e período), `0150` (participantes), `0200` (produtos), `C100/C170/C190` (NF-e/NFC-e), `C500/C590` (energia), `D100/D190` (CT-e) e `D500/D590` (comunicação).

## Tabela de CFOP

`cfop_oficial.json` traz os 619 CFOPs da tabela oficial (IT 2023.002 v2.10 – Ajuste SINIEF 07/2005 e alterações), com descrição, nota explicativa e natureza usada no painel (Venda, Compra para revenda, Devolução, Transferência, Remessa/Retorno etc.).

- **revisada**: natureza conferida manualmente.
- **automática**: natureza definida pelo grupo do CFOP e pelos indicadores oficiais de devolução, retorno e remessa. Na página, a natureza pode ser ajustada na aba Tabela CFOP.

A base de CFOPs foi extraída de [consulta-cfop](https://github.com/silvioalbqrq/consulta-cfop). Em caso de divergência, prevalece a legislação vigente.

## Arquivos

| Arquivo | Função |
|---|---|
| `Painel SPED.html` | Página completa (HTML, CSS e JS em um arquivo, com a tabela de CFOP embutida) |
| `sped_painel.py` | Script que gera a planilha |
| `Gerar Painel SPED.bat` | Atalho do script para Windows |
| `cfop_oficial.json` | Tabela oficial de CFOP usada pelo script |

## Logo (opcional)

Para exibir um logo no cabeçalho, crie um arquivo `logo.js` na mesma pasta da página com o conteúdo:

```js
window.LOGO_SVG = '<svg ...>...</svg>';
```

Sem esse arquivo, a página funciona normalmente, sem logo. O `logo.js` está no `.gitignore` e não é enviado ao repositório.
