// # Importação de arquivo TXT vivo para PLANILHA automatizado
// Tratamento de dados cadastrais em Planilhas - Power Query - Power BI
//
// Esta planilha serve para automatizar a importação de dados cadastrais da vivo em arquivos .txt para uma planilha. 
// Muito útil quando há uma série de DC dentro de um mesmo arquivo .txt
// É possível a importação de vários arquivos de dentro de uma pasta, mas é necessário a modificação do código. Me contate para saber mais.
//
// Clque no arquivo Tratamento dados cadastrais VIVO.xlsx, faça o download siga as instruções contidas na planilha ou logo abaixo:
//
// Instruções de uso Tratamento de DADOS CADASTRAIS VIVO
// 1º Aperte as teclas - Alt + S + P + N + F (Alt + SPNF) - Atalho para configurar arquivo txt fonte dos dados
// 2º Clique em Alterar Fonte... no canto inferior esquerdo:
// 3º Clique em Procurar... e logo após selecione o arquivo txt a ser importado / tratado e aperte abrir
// 4° Após selecionar clique em Importar...
// 5º Clique em Fechar / OK
// 6º Aperte Ctrl + Alt + F5 e Verifique a planilha Dados Vivo tratados
// 7º Salve as atualizações na pasta e com o nome desejados - "Alt A A" - Atalho para salvar como
//
// Disclaimer / Aviso legal: 
// Este desenvolvedor NÃO se responsabiliza por eventuais erros no tratamento de dados ou análises realizadas pelo analista. É de responsabilidade do USUÁRIO verificar  se os DADOS TRATADOS (PLANILHA) estão de acordo com o arquivo original.
// Em caso de erro favor notificar joao.barbosa@pc.pi.gov.br ou via github
// Obs: Planilha desenvolvida no Power Query - Microsoft Excel 365, alguns recursos podem estar indisponíveis em versões mais antigas (2010, entre outras) e necessitar adicionar o Power Query para o devido funcionamento

//Segue o código na Linguagem M - Power QUery
let
    Fonte = Csv.Document(File.Contents("C:\Users\joao_\OneDrive\Documentos FTSP\APC - João Filho\ANALISES\Tratamento de dados\Arquivos de testes\DadosCadastraisCustomizado.txt"),[Delimiter=",", Columns=2, Encoding=65001, QuoteStyle=QuoteStyle.None]),
    Linhas_filtradas = Table.SelectRows(Fonte, each ([Column1] <> null and [Column1] <> "" and [Column1] <> "*                                                                              *" and [Column1] <> "*                           PARÂMETRO(S) DE CONSULTA                           *" and [Column1] <> "* ---------------------------------------------------------------------------- *" and [Column1] <> "* ............................................................................ *" and [Column1] <> "********************************************************************************")),
    Valor_substituido_ponto = Table.ReplaceValue(Linhas_filtradas,".","",Replacer.ReplaceText,{"Column1"}),
    Valor_substituído_asterisco = Table.ReplaceValue(Valor_substituido_ponto,"*","",Replacer.ReplaceText,{"Column1"}),
    Dividir_coluna_por_delimitador_doispontos = Table.SplitColumn(Valor_substituído_asterisco, "Column1", Splitter.SplitTextByEachDelimiter({":"}, QuoteStyle.Csv, false), {"Column1.1", "Column1.2"}),
    Texto_aparado = Table.TransformColumns(Dividir_coluna_por_delimitador_doispontos,{{"Column1.1", Text.Trim, type text}, {"Column1.2", Text.Trim, type text}}),
    Texto_limpo = Table.TransformColumns(Texto_aparado,{{"Column1.1", Text.Clean, type text}, {"Column1.2", Text.Clean, type text}}),
    Coluna_indice_adicionada = Table.AddIndexColumn(Texto_limpo, "Índice", 1, 1, Int64.Type),
    Coluna_em_pivo_dinamica = Table.Pivot(Coluna_indice_adicionada, List.Distinct(Coluna_indice_adicionada[Column1.1]), "Column1.1", "Column1.2"),
    Preenchido_acima = Table.FillUp(Coluna_em_pivo_dinamica,{"Hora", "RELATÓRIO DE PESQUISA", "NÚMERO DA LINHA", "CLIENTE", "CPF", "ENDEREÇO", "BAIRRO", "CEP", "MUNICÍPIO", "ESTADO", "MODALIDADE", "SITUAÇÃO", "DATA HABILITAÇÃO"}),
    Tipo_alterado = Table.TransformColumnTypes(Preenchido_acima,{{"Data", type date}, {"Hora", type time}, {"DATA HABILITAÇÃO", type date}}),
    Linhas_nulas_filtradas_data = Table.SelectRows(Tipo_alterado, each [Data] <> null and [Data] <> ""),
    Coluna_removida_relatorio_pesquisa = Table.RemoveColumns(Linhas_nulas_filtradas_data,{"RELATÓRIO DE PESQUISA"}),
    Criar_coluna_extraido_DDD = Table.AddColumn(Coluna_removida_relatorio_pesquisa, "DDD", each Text.BetweenDelimiters([NÚMERO DA LINHA], "(", ")"), type text),
    Coluna_CPF_retirar_traco = Table.ReplaceValue(Criar_coluna_extraido_DDD,"-","",Replacer.ReplaceText,{"CPF"}),
    Colunas_renomeadas = Table.RenameColumns(Coluna_CPF_retirar_traco,{{"Data", "Data_DC"}, {"Hora", "Hora_DC"}, {"NÚMERO DA LINHA", "NUMERO"}}),
    Colunas_reordenadas = Table.ReorderColumns(Colunas_renomeadas,{"Column2", "Índice", "Data_DC", "Hora_DC", "DDD", "NUMERO", "DATA HABILITAÇÃO", "CLIENTE", "CPF", "ENDEREÇO", "BAIRRO", "CEP", "MUNICÍPIO", "ESTADO", "MODALIDADE", "SITUAÇÃO"}),
    Outras_colunas_removidas = Table.SelectColumns(Colunas_reordenadas,{"Data_DC", "Hora_DC", "DDD", "NUMERO", "DATA HABILITAÇÃO", "CLIENTE", "CPF", "ENDEREÇO", "BAIRRO", "CEP", "MUNICÍPIO", "ESTADO", "MODALIDADE", "SITUAÇÃO"})
in
    Outras_colunas_removidas
