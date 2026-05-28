# Scripts Node

Utilitários pequenos em Node.js para tarefas rápidas e pontuais.

## `merge-excels.js`

### Mescla vários arquivos Excel em uma única pasta de trabalho.

O que ele faz:
- Lê a primeira aba de cada arquivo por padrão
- Combina todas as linhas em uma única aba de saída
- Aceita arquivos `.xlsx` individuais ou uma pasta com arquivos `.xlsx`
- Pode remover duplicados por uma coluna escolhida
- Pode ordenar por uma coluna escolhida
- Aceita nome da coluna ou índice da coluna começando em 1 para dedupe e ordenação

### Como usar

Instalar dependências:

```bash
npm install
```

Executar direto:

```bash
node merge-excels.js file1.xlsx file2.xlsx -o merged.xlsx
```

Você também pode passar uma pasta e o script vai coletar todos os arquivos `.xlsx` dentro dela de forma recursiva:

```bash
node merge-excels.js ./excels -o merged.xlsx
```

Ou usar o script do npm:

```bash
npm run merge-excel -- file1.xlsx file2.xlsx -o merged.xlsx
```

### Opções

- `-o, --output <arquivo>`: caminho do arquivo de saída. O padrão é `merged.xlsx`
- A entrada pode ser um ou mais arquivos `.xlsx`, ou uma ou mais pastas contendo arquivos `.xlsx`
- `--dedupe <coluna>`: remove linhas duplicadas usando o nome de uma coluna ou o índice começando em 1
- `--sort <coluna>`: ordena por uma coluna usando o nome ou o índice começando em 1
- `--desc`: ordena em ordem decrescente
- `--sheet <nome|índice>`: usa uma aba específica de cada arquivo. O padrão é a primeira aba
- `-h, --help`: mostra a ajuda

### Exemplos

Mesclar três arquivos e salvar em uma pasta personalizada:

```bash
node merge-excels.js sales-1.xlsx sales-2.xlsx sales-3.xlsx -o output/merged.xlsx
```

Remover linhas duplicadas por e-mail:

```bash
node merge-excels.js *.xlsx --dedupe Email -o merged.xlsx
```

Ordenar por uma coluna de data:

```bash
node merge-excels.js *.xlsx --sort CreatedAt --desc -o merged.xlsx
```

Usar a segunda aba de cada arquivo:

```bash
node merge-excels.js *.xlsx --sheet 2 -o merged.xlsx
```

## Adicionando novos scripts

Coloque novos scripts Node.js nesta pasta e documente-os aqui com:
- O que o script faz
- Flags ou entradas obrigatórias
- Exemplos de comandos
- Qualquer dependência necessária

