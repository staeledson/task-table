import pandas as pd

# Extraindo tabelas.
main_table = pd.read_excel('./tabelas/itens_vituz.xlsx')
basic_table = pd.read_excel('./tabelas/codigos_ws_basico.xls')
specialized_table = pd.read_excel('./tabelas/codigos_ws_especializado.xls')
strategic_table = pd.read_excel('./tabelas/codigos_ws_estrategico.xls')

class Itens(object):
    def __init__(self, nome, categoria, grupo, unidade_medida, categoria_horus, cod_horus):
        self.nome = nome
        self.categoria = categoria
        self.grupo = grupo
        self.unidade_medida = unidade_medida
        self.categoria_horus = categoria_horus
        self.cod_horus = cod_horus

tabela_de_itens = []
for index, linha in enumerate(main_table['nome']):
    new_iten = Itens(
        nome = main_table['nome'][index], 
        categoria = main_table['categoria'][index], 
        grupo = main_table['grupo'][index],
        unidade_medida = main_table['unidade_medida'][index], 
        categoria_horus = '', 
        cod_horus = main_table['cod_horus'][index],
        )
    tabela_de_itens.append(new_iten)

# Inserindo Código horus nos itens já existentes da tabela 'itens_vituz'.
def insertCodeHorus(table, cat_horus):
    for index, line in enumerate(table['Descrição']):
        for item in tabela_de_itens:
            if(item.nome in table['Descrição'][index]):
                item.categoria_horus = cat_horus
                item.cod_horus = basic_table['Código'][index]

insertCodeHorus(basic_table, 'BÁSICO')
insertCodeHorus(specialized_table, 'ESPECIALIZADO')
insertCodeHorus(strategic_table, 'ESTRATÉGICO')

# Função para verificar e evitar a duplicidade de itens
def verifyItem (n):
    for index, item in enumerate(tabela_de_itens):
        if (n in item.nome):
            return True
        
# Função para inserir itens das tabelas de códigos que ainda não exitem na tabela 'itens_vituz'.
def insertItem(table, grupo, categoria_horus):
    for index, description in enumerate(table['Descrição']):
        #print(description)
        if(not verifyItem(description)):
            new_item = Itens(
                nome = table['Descrição'][index], 
                categoria = 'Material Médico Hospitalar', 
                grupo = grupo,
                unidade_medida = '', 
                categoria_horus = categoria_horus, 
                cod_horus = table['Código'][index],
                )
            tabela_de_itens.append(new_item)

insertItem(basic_table, 'MEDICAMENTOS BÁSICOS', 'BÁSICO')
insertItem(specialized_table, 'MEDICAMENTOS ESPECIALIZADOS', 'ESPECIALIZADO')
insertItem(strategic_table, 'MEDICAMENTOS ESTRATÉGICOS', 'ESTRATÉGICO')


# Preparando dados para criar nova planilha.
arr_nome = []
arr_categoria = []
arr_grupo = []
arr_unidade_medida = []
arr_categoria_horus = []
arr_cod_horus = []

dados = []
for description in tabela_de_itens:
    dados.append([
        description.nome,
        description.categoria,
        description.grupo,
        description.unidade_medida,
        description.categoria_horus,
        description.cod_horus,
    ])
    
# Ordenando dados e criando nova planilha.
dados.sort()
df = pd.DataFrame(
   dados,
   columns=['nome', 'categoria', 'grupo', 'unidade_medida', 'categoria_horus', 'cod_horus']
)
df.to_excel('./tabelas/tabela_final.xlsx')
