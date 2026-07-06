# word_macros
Macros para facilitar o dia a dia no trabalho na elaboração de documentos jurídicos.

Word Valor por Extenso

Macro em VBA para Microsoft Word que converte valores monetários para sua representação por extenso, preservando exatamente a quantidade de casas decimais informadas, sem arredondamento.

Foi desenvolvida para documentos jurídicos, contratos, laudos, escrituras, pareceres e outros textos em que valores monetários precisam ser escritos simultaneamente em algarismos e por extenso.

Principais recursos
Converte valores monetários para o formato por extenso.

Formata automaticamente o valor com separador de milhares.

Não arredonda as casas decimais.

Preserva exatamente a quantidade de casas decimais digitadas (até 5 casas).

Suporta números inteiros muito grandes (milhares, milhões, bilhões e trilhões).

Utiliza a denominação correta para frações monetárias:

Exemplos

Entrada:
1250,35

Saída:
R$ 1.250,35 (mil duzentos e cinquenta reais e trinta e cinco centavos)

Entrada:
1000000

Saída:
R$ 1.000.000 (um milhão de reais)

Entrada:
1000001

Saída:
R$ 1.000.001 (um milhão e um reais)

Entrada:
0,5

Saída:
R$ 0,5 (cinco décimos)

Entrada:
0,001

Saída:
R$ 0,001 (um milésimo)

Entrada:
0,00035

Saída:
R$ 0,00035 (trinta e cinco centésimos de milésimo)

Como utilizar

Digite o valor no documento do Word utilizando vírgula como separador decimal.

A macro substituirá o valor digitado pelo valor formatado e acrescentará sua representação por extenso entre parênteses.

Regras de funcionamento

Não utilize o símbolo R$ antes do valor.

Não utilize pontos como separadores de milhares, ele coloca automaticamente.

Utilize vírgula como separador decimal.

São aceitas até 5 casas decimais, ele não realiza arredondamentos.

O cursor deve estar imediatamente após o número.
