# Excel-Converter
Uma Ferramenta para converter arquivos xls e xlsx em txt. A saída é um documento txt com as colunas separadas, respeitando o espaçamento de 16 caracters de espaçamento entre uma coluna e outra. Permitindo assim uma facil importação para sistemas que utilizem documentos txt assim como uma fácil vizualização dos dados dentro do documento.

Primeiramente deve-se escolher uma pasta em seu computador para a saída dos arquivos. Feito isso é possível escolher três opções:

1 - Criação de um padrão de conversão: Com o padrão de conversão é possível personalizar o número de colunas que o arquivo txt deve ter, assim como escolher quais colunas do arquivo xls, xlsx devem ser repassadas para o documento de sáida.

Por exemplo. Suponha que o arquivo xls contenha 10 colunas, contudo não é de seu interesse as duas primeiras colunas. Ao criar o padrão de saída deve-se então inserir no segundo campo de preenchimento o número total de colunas (neste caso 8), sendo gerado em sequencia o seguinte padrão de preenchimento:

Total de Colunas: 8
Coluna 1: 
Coluna 2: 
Coluna 3: 
Coluna 4: 
Coluna 5: 
Coluna 6: 
Coluna 7: 
Coluna 8:


Cada Coluna deste padrão representa as colunas do arquivo de saída txt. Como queremos ignorar as duas primeiras colunas, a primeira coluna do documento txt deve ser a de número 3 do arquivo xls,e assim sucessivamente. O preenchimento para configurar este padrão de saída então deve ser:

Total de Colunas: 8
Coluna 1: 3
Coluna 2: 4
Coluna 3: 5
Coluna 4: 6
Coluna 5: 7
Coluna 6: 8
Coluna 7: 9
Coluna 8: 10


Resumidamente, para se criar um padrão, deve-se inserir o número de colunas do arquivo final, e preencher cada coluna gerada com alguma coluna existente dentro do documento xls. Sendo possível também inserir o valor 0, por exemplo: Coluna 1: 0. Neste caso a coluna 1 do documento final será totalmente preenchida com o caracter (X), simulando uma ausência de informação.

Existe um padrão denominado ManterPadrao, e como o próprio nome sugere irá converter o arquivo para txt respeitando a ordem e o número de colunas presentes no arquivo xls.



2 - Conversão em massa: Irá converter todos os arquivos xls e xlsx contidos na pasta selecionada. A conversão em massa é realizada selecionando uma pasta de saída, um padrão de conversão e a opção de inserir se alguma coluna do documento xls contém uma informação de data. Caso isso aconteça é necessário preencher este campo com o número das colunas que contém uma informação de data, caso o contrário, seus valores não serão preservados.

Após esse procedimento é aberta uma barra de progresso, a mesma é incrementada seguindo a divisão do arquivo convertido pelo número total, atigindo 100% quando o valor da divisão final é 1 . A função desta barra é para dar um feedback de que os arquivos estão sim sendo convertidos, mas não é precisa, visto que não leva em consideração os diferentes tamanhos de arquivos que podem exister na pasta. 


3 - Converter um arquivo: Irá converter um arquivo xls ou xlsx selecionado. Para a conversão deve-se selecionar uma pasta de saída para o arquivo txt, um padrão de conversão, e assim como descrito na Conversão em massa, também deve-se selecionar quais colunas no documento xls ou xlsx contém uma informação de data. Diferente da Conversão em massa essa opção ainda não possui uma barra de progresso, o que pode ser um pouco incomodo caso o arquivo a ser convertido seja demasiado grande (exemplo: 12 colunas e 16000 linhas).

