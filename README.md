# Sobral Programação Industrial

Primeira versão funcional do gerador de programação de produção.

## Fluxo
1. Envie Pedido.
2. Envie Usinagem.
3. Informe data e turno.
4. Clique em Gerar programação.
5. Revise a grade.
6. Consulte Quantidades líderes e Apontamentos.
7. Exporte o Excel.

## Já faz
- Lê XLSX, XLS e CSV.
- Reconhece Pedido, Item, OF, Descrição/Produto, Qtd e outros aliases comuns.
- Identifica automaticamente tipo da peça, medida, rebaixo, material, acabamento, cor, lado e máquina.
- Junta dados de Pedido e Usinagem por OF ou Pedido + Item.
- Consolida quantidades iguais para os líderes.
- Mantém controle de apontamentos.
- Exporta Excel com:
  - CONTROLE GERENTE
  - APONTAMENTOS
  - QUANTIDADES LIDERES
  - uma aba por máquina/processo.

## Rodar
npm install
npm run dev

## Deploy
Suba para GitHub e importe na Vercel.

## Próxima etapa
Ajustar o parser aos arquivos reais da Famossul, inclusive PDF, para a leitura ficar fiel ao layout real.
