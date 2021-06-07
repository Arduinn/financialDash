# Dashboard packages
import dash
import dash_bootstrap_components as dbc
import dash_core_components as dcc
from dash_html_components.Div import Div
import dash_table as dtt
import dash_html_components as html
from dash.dependencies import Input, Output, State
import datetime

# Finance Packages
import plotly.express as px
import pandas as pd
import plotly.graph_objs as go
import plotly.io as pio
import plotly.offline as py
from plotly.offline import plot, iplot

# Importing Files
external_stylesheets = ['./financialStyle.css']


# Mining Data
# import financialDash


# APP
moedas = pd.read_excel("data.xlsx", sheet_name='moedas', engine='openpyxl')
acoes = pd.read_excel("data.xlsx", sheet_name='acoes', engine='openpyxl')
serieTemp = pd.read_excel("data.xlsx", sheet_name='serieTemp', engine='openpyxl')
dfTweets = pd.read_excel("data.xlsx", sheet_name='twiComment', engine='openpyxl')

# Creating DASH
app = dash.Dash(__name__)

app.layout = html.Div([
    
    html.Div(
        className="app-header",
        children=[
            html.H1('Finance Dashboard', className="app-header--title"),
            html.H2('Made by Victor Arduin', className="app-header--subtitle"),
            html.H3(datetime.datetime.now().strftime('%Y-%m-%d'), style={'fontSize': '25px', 'textAlign': 'center'})
        ],
    ),
    
    html.Div(
        html.Div(children=[
            html.Div(children=[
                html.H3('Currencies'),
                dtt.DataTable( 
                    data=moedas.to_dict('records'),
                    columns=[{'id': c, 'name': c} for c in moedas.columns],
                    style_cell=dict(textAlign='left'),
                    style_header=dict(backgroundColor="#4675B3", textAlign='center'),
                    style_data=dict(backgroundColor="lavender", textAlign='center'),
                    style_cell_conditional=[
                        {'if': {'column_id': 'Moeda'},
                        'width': '200px'},
                        ])
                ], className='app-body--tables'),
            
            html.Div(children=[
                html.H3('Commodities'),
                dtt.DataTable( 
                    data=moedas.to_dict('records'),
                    columns=[{'id': c, 'name': c} for c in moedas.columns],
                    style_cell=dict(textAlign='left'),
                    style_header=dict(backgroundColor="#4675B3", textAlign='center'),
                    style_data=dict(backgroundColor="lavender", textAlign='center'),
                    style_cell_conditional=[
                        {'if': {'column_id': 'Moeda'},
                        'width': '200px'},
                        ])
                ], className='app-body--tables'),
            
        ])
    ),
    
    html.Div(
        children=[
                html.H3('Stock Info'),
                dtt.DataTable( 
                    data=acoes.to_dict('records'),
                    columns=[{'id': c, 'name': c} for c in acoes.columns],
                    style_as_list_view=True,
                    style_table={'maxHeight':'600px',
                                 'overflowY':'scroll'},
                    sort_action='native',
                    filter_action='native',
                    # fixed_columns={'headers': True, 'data': 1},
                    style_header={'border': '1px solid black',
                                  'height':'auto',
                                  'backgroundColor':"#4675B3", 
                                  'textAlign':'center',
                                  'whiteSpace':'normal'},
                    style_cell={ 'border': '1px solid grey',
                                'width': '200px'},
                    style_data={
                        'whiteSpace': 'normal',
                        'height':'auto'
                    }
                )], className='.app-body--acoes'
        ),
    
    html.Div(children=[html.H3('Time Serie'),
        dcc.Dropdown(
            id="ticker",
            options=[{"label": x, "value": x} 
                    for x in serieTemp.Empresas.unique()],
            value=serieTemp.Empresas.unique()[1],
            clearable=False,
        ),
        dcc.Graph(id="time-series-chart-lines"),
        dcc.Graph(id="time-series-chart"),
        dcc.Graph(id="time-series-chart-avg"),
        html.Div(id='table-container',  className='tableDiv'),
    ])
    
])



@app.callback(
    Output("time-series-chart-lines", "figure"), 
    [Input("ticker", "value")])
    
def display_time_series_lines(ticker):
    df = serieTemp.loc[serieTemp['Empresas'] == ticker]
    
    Close = go.Scatter(
    x=df.Date,
    y=df.Close,
    name = ticker,
    line = dict(color = '#17BECF'),
    opacity = 0.8)

    data = [Close]

    layout = dict(
    title="Série com Rangeslider e Botoes",
    title_x=0.5,
    xaxis=dict(
    rangeselector=dict(
    buttons=list([
    dict(count=1,
    label='1m',
    step='month',
    stepmode='backward'),
    dict(count=6,
    label='6m',
    step='month',
    stepmode='backward'),
    dict(step='all')
    ])
    ),
    rangeslider=dict(
    visible = True
    ),
    type='date'
    )
    )

    fig = dict(data=data, layout=layout)
    return fig 

@app.callback(
    Output("time-series-chart", "figure"), 
    [Input("ticker", "value")])

def display_time_series(ticker):
    
    df = serieTemp.loc[serieTemp['Empresas'] == ticker]

    trace = go.Candlestick(x=df['Date'],
    open=df['Open'],
    high=df['High'],
    low=df['Low'],
    close=df['Close'])
    data = [trace]

    layout = {
    'title': ticker,
    'title_x': 0.5,
    'yaxis': {'title': 'Precos'},
    'annotations': [{
    'x': '2017-05-17', 
    'y': 15, 
    'xref': 'x', 
    'yref': 'y',
    'showarrow': True,
    'font':dict(
    family="Courier New, monospace",
    size=12
    ),
    'text': 'Audio Joesley',
    'align':"center",
    'arrowhead':2,
    'arrowsize':1,
    'arrowwidth':2,
    'bordercolor':"#c7c7c7",
    'borderwidth':2,
    'borderpad':4
    },
    {
    'x': '2016-01-01', 
    'y': 7, 
    'xref': 'x', 
    'yref': 'y',
    'showarrow': True,
    'font':dict(
    family="Courier New, monospace",
    size=12
    ),
    'text': 'Impeachment Dilma',
    'align':"center",
    'arrowhead':2,
    'arrowsize':1,
    'arrowwidth':2,
    'bordercolor':"#c7c7c7",
    'borderwidth':2,
    'borderpad':4
    },
    {
    'x': '2018-05-27', 
    'y': 27, 
    'xref': 'x', 
    'yref': 'y',
    'showarrow': True,
    'font':dict(
    family="Courier New, monospace",
    size=12
    ),
    'text': 'Greve dos Caminhoneiros',
    'align':"center",
    'arrowhead':2,
    'arrowsize':1,
    'arrowwidth':2,
    'bordercolor':"#c7c7c7",
    'borderwidth':2,
    'borderpad':4
    }]
    }
    
    fig = dict(data=data, layout=layout)
    return fig

@app.callback(
    Output("time-series-chart-avg", "figure"), 
    [Input("ticker", "value")])

def display_time_series_avg(ticker):
    
    df = serieTemp.loc[serieTemp['Empresas'] == ticker]

    # Média simples de 3 dias
    df['MM_3'] = df.Close.rolling(window=3).mean()

    # Média simples de 17 dias
    df['MM_17'] = df.Close.rolling(window=17).mean()

    close = go.Scatter(
    x=df.Date,
    y=df.Close,
    name = ticker,
    line = dict(color = '#00E875'),
    opacity = 0.8)

    MM_3 = go.Scatter(
    x=df.Date,
    y=df['MM_3'],
    name = "Média Móvel 3 Períodos",
    line = dict(color = '#DBFC29'),
    opacity = 0.8)

    MM_17 = go.Scatter(
    x=df.Date,
    y=df['MM_17'],
    name = "Média Móvel 17 Períodos",
    line = dict(color = '#C6003B'),
    opacity = 0.8)

    data = [close, MM_3, MM_17]

    fig = dict(data=data)
    return fig

@app.callback(
    dash.dependencies.Output('table-container','children'),
    [dash.dependencies.Input("ticker", "value")])

def display_tweety_comments(ticker):
    df = dfTweets[dfTweets['Empresas'] == ticker]

    return html.Div([
        dtt.DataTable(
       id='table',
       columns=[{"name": i, "id": i} for i in df.columns],
       data=df.to_dict('records'),
       editable=True
   )
])


if __name__ == "__main__":
    app.run_server(debug=True)
