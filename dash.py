#programa feito por Moon Hyuk Kang, 19209988
#bibliotecas usadas
import pandas as pd                    
import matplotlib.pyplot as plt
import dataframe_image as dfi
from tkinter import *
from tkinter import filedialog
import plotly.express as px
import plotly.graph_objects as go
import dash
import dash_core_components as dcc
import dash_html_components as html
import webbrowser
#-------------------------------------------------------------------------------------------------------------------------------
#interface de usuário
window = Tk()
window.title('DashPLOT')  
window.config(background = "#a4dfd1") 
window.geometry('700x500')

#-------------------------------------------------------------------------------------------------------------------------------
#função para abrir o explorador de arquivo e escolher o arquivo excel
def abrir_arquivo():
    global tabela_vendas
    #procurar arquivo pelo explorador de arquivo
    filepath = filedialog.askopenfilename(filetypes=(("csv", "*.csv"), ("all files", "*.*"))) 
    label = Label(window, text=filepath) 
    label.pack()
    label.place(x=200, y=80)
    #salvar o dataframe dentro da varíavel tabela_vendas
    tabela_vendas = pd.read_csv(filepath)
    
#função para filtrar e processar dados do arquivo excel
def filtrar():
    global tabela_faturamento
    global tabela_quantidade
    global ticket_medio
    global top_produto
    global tabela_produto
    
    tabela_faturamento = tabela_vendas[["ID Loja", "Valor Final"]].groupby("ID Loja").sum()   # agrupar por ID Loja e fazer a soma do valor final
    tabela_faturamento = tabela_faturamento.sort_values(by = "ID Loja", ascending= True) #ascendente por ordem alfabética de ID Loja

    tabela_quantidade = tabela_vendas[["ID Loja", "Quantidade"]].groupby("ID Loja").sum().sort_values(by='Quantidade', ascending=False) 

    ticket_medio = (tabela_faturamento['Valor Final'] / tabela_quantidade['Quantidade']).to_frame() #converter para tabela 
    ticket_medio = ticket_medio.rename(columns={0:"Ticket Medio"})  #função para reconemar coluna
    
    #dataframe de tabela de quantidade vendido por cada produto.
    #dataframe para dashboard online
    tabela_produto = tabela_vendas.groupby("Produto", as_index=False).sum().sort_values("Quantidade", ascending=False) 
    tabela_produto = tabela_produto.iloc[0:100]

    #dataframe para a função img_produto
    top_produto = tabela_vendas[['Produto', 'Quantidade']].groupby('Produto').sum('Quantidade').head(20).sort_values(by='Quantidade', ascending=False)

#------------------------------------------------------------------------------------------------------------------------
#funções para converter tabelas em imagens png e exportar (pasta padrão é onde o executavel estiver localizado)
#tabela de quantidade, vai no botão bt1    
def img_quantidade():
    df_styled1 = tabela_quantidade.style.background_gradient()  
    dfi.export(df_styled1,"quantidade.png")
#tabela de faturamento, vai no botão bt2
def img_faturamento():
    df_styled2 = tabela_faturamento.style.background_gradient()
    dfi.export(df_styled2,"faturamento.png")
#tabela de ticket medio, vai no botão bt3
def img_ticket_medio():
    df_styled3 = ticket_medio.style.background_gradient()
    dfi.export(df_styled3,"ticket_medio.png")
#tabela de produtos mais vendidos, vai no botão bt4    
def img_produto():
    df_styled4 = top_produto.style.background_gradient()
    dfi.export(df_styled4,"top_produtos.png")



#------------------------------------------------------------------------------------------------------------------------
#começar servidor 
def start_server():
    app = dash.Dash(__name__)

    #criar gráficos
    #gráfico em torta (pie) da tabela de faturamento
    fig_fatura = px.pie(data_frame=tabela_vendas,
                        names='ID Loja',
                        values = 'Valor Final',
                        labels={'ID Loja':'Loja', 'Valor Final':'Fatura em R$'},
                        height=600,
                        hole=0.3,
                        title='Gráfico de Faturamento')

    #gráfico em histograma da tabela de vendas
    fig_vendas = px.histogram(tabela_vendas,
                              y='Quantidade',
                              x='ID Loja',
                              labels={'ID Loja':'Loja', 'Quantidade':'Quantidade'},
                              color_discrete_sequence=['#95c551'],
                              title='Histograma de Produtos vendidos por Loja')

    #tabela dos 100 produtos mais vendidos
    fig_produto = go.Figure(data=go.Table(
        header=dict(values=list(tabela_produto[['Produto','Quantidade']].columns),
            fill_color = 'paleturquoise',
            align='center'),
        cells=dict(dict(values=[tabela_produto['Produto'], tabela_produto['Quantidade']],
                        fill_color = '#e3d1e3'),
                        align='center')
        ))
                                          
    #----------------------------------------------------------------------------------------------------------------
    #fazer layout do dashboard
    #primeiro layout 
    app.layout = html.Div([ 
        html.H1('FATURAMENTO',
                style={'backgroundColor':'#faefde',   #cor do background em código HEX
                       'textAlign': 'center',         #centralizar o texto
                       'font-family':'Courier New',   #fonte do texto
                       'color': 'red',                #cor do texto
                       'fontSize': 40}),              #tamanho da fonte
        #texto do primeiro layout
        html.P('Este gráfico mostra o faturamento total de cada uma das lojas',
               style={'textAlign': 'center',
                      'font-family':'Courier New',
                      'fontSize': 20}),
        #gráfico do primeiro layer
        dcc.Graph(id='grafico_pie',figure=fig_fatura),
        
            #segundo layer
            html.H2('QUANTIDADE', style={'backgroundColor':'#faefde',   
                                         'textAlign': 'center',
                                         'color': 'blue',
                                         'font-family':'Courier New',
                                         'fontSize': 40}),
            #texto do segundo layer
            html.P('Este histograma mostra a quantidade produtos vendidos por cada loja.',
                                   style={'textAlign': 'center',
                                          'color': 'blue',
                                          'font-family':'Courier New',
                                          'fontSize': 20}),
            #gráfico do segundo layer
            dcc.Graph(figure=fig_vendas),

            #terceiro layer
            html.H3('PRODUTOS', style={'backgroundColor':'#faefde',
                                       'textAlign': 'center',
                                       'color': 'orange',
                                       'font-family':'Courier New',
                                       'fontSize': 40}),
            #texto do segundo layer
            html.P('Esta tabela mostra os 100 produtos mais vendidos.',
                                   style={'textAlign': 'center',
                                          'color': 'blue',
                                          'font-family':'Courier New',
                                          'fontSize': 20}),
            #gráfico do terceiro layer
            dcc.Graph(figure=fig_produto),

        ])
                                          
    webbrowser.open('http://127.0.0.1:8050')    #vai abrir o link do dash automaticamente no navegador de internet padrão
    
    if __name__ == '__main__':
        app.run_server(debug=False)

#--------------------------------------------------------------------------------------------------------------------------
#botão para abrir arquivo excel        
button_choose_file = Button(window, width = 20, text='Escolha o arquivo', font=("Helvetica", "9", "bold"), bg = '#ffbb42', command=abrir_arquivo)
button_choose_file.pack()
button_choose_file.place(x=200, y=50)

#filtrar e processar dados
button_start = Button(window, width = 10, text = "OK", font=("Helvetica", "9", "bold"), bg = '#ffbb42', command = filtrar)
button_start.pack()
button_start.place(x=250, y=120)

#botão para iniciar servidor
bStartServer = Button(window, width = 20, text='Começar servidor', font=("Helvetica", "9", "bold"), bg = '#ffbb42', command=start_server)
bStartServer.pack()
bStartServer.place(x=200, y=190)


#---------------------------------------------------------------------------------------------------------------------------
# botão para exportar as tabelas pandas como imagens png

#vai receber a função img_quantidade
bt1 = Button(window, width = 20, text='Quantidade', bg = '#ffbb42', font=("Helvetica", "9", "bold"), command=img_quantidade)
bt1.place(x=400, y=400)

#vai receber a função img_faturamento
bt2 = Button(window, width = 20, text='Faturamento', bg = '#ffbb42', font=("Helvetica", "9", "bold"), command=img_faturamento)
bt2.place(x=400, y=350)

#vai receber a função img_ticket_medio
bt3 = Button(window, width = 20, text='Ticket Médio', bg = '#ffbb42', font=("Helvetica", "9", "bold"), command=img_ticket_medio)
bt3.place(x=400, y=300)

#vai receber a função img_produto
bt4 = Button(window, width = 20, text='Produto', bg = '#ffbb42', font=("Helvetica", "9", "bold"), command=img_produto)
bt4.place(x=400, y=450) 

#deixar o programa rodando em loop
window.mainloop()
