Attribute VB_Name = "modConstantes"
Option Explicit

Public Const sfProjudi As String = "Projudi", sfPJe As String = "PJe"
Public Const sfTJBA As String = "TJ/BA", sfTRT5 As String = "TRT 5"
Public Const sfURLBuscaProjudiParte As String = "https://projudi.tjba.jus.br/projudi/buscas/ProcessosParte"
Public Const sfURLBuscaProjudiAdvogado As String = "https://projudi.tjba.jus.br/projudi/buscas/ProcessosQualquerAdvogado"
Public Const sfURLBuscaPJe As String = "https://pje.tjba.jus.br/pje-web/Processo/ConsultaProcesso/listView.seam"
Public Const sfAdvsEmbasa As String = "ANA PAULA AMORIM CORTES," & _
                                    "ANALYZ PESSOA BRAZ DE OLIVEIRA," & _
                                    "ANANDA ATMAN AZEVEDO DOS SANTOS," & _
                                    "ANGELA MOISES FARIA LANTYER," & _
                                    "CARLOS HENRIQUE MARTINS JUNIOR," & _
                                    "CESAR BRAGA LINS BAMBERG RODRIGUEZ," & _
                                    "ELISANGELA DE QUEIROZ FERNANDES BRITO," & _
                                    "FABIO JUNIO SOUZA OLIVEIRA," & _
                                    "FERNANDA BARRETO MOTA," & _
                                    "," & _
                                    "IZABELA RIOS LEITE," & _
                                    "," & _
                                    "JAIRO BRAGA LIMA," & _
                                    "," & _
                                    "JORGE KIDELMIR NASCIMENTO DE OLIVEIRA FILHO," & _
                                    "JULIANA CARDOSO NASCIMENTO," & _
                                    "LIVIA MOURA MARQUES DE OLIVEIRA," & _
                                    "MARIA QUINTAS RADEL," & _
                                    "MARIANA BRASIL NOGUEIRA LIMA," & _
                                    "MILA LEITE NASCIMENTO," & _
                                    "PEDRO CAMERA PACHECO," & _
                                    "," & _
                                    "TANIA MARIA REBOUCAS," 'Agentes que serão ignorados ao contar digitalizações

Public Const sfAgentesAutomaticosProjudi As String = "ECT,SISTEMA CNJ," 'Agentes que serão ignorados ao contar digitalizações
Public Const sfQtdEventosPossivelExecucao As Byte = 60 'Quantidade de eventos até a qual o sistema presume que não é possível ser um RI de execução.

