from dash import Dash, dcc, html, Input, Output, callback
import pandas as pd
import numpy as np
from dash import dash_table as dt



external_stylesheets = ['https://codepen.io/chriddyp/pen/bWLwgP.css']

app = Dash(__name__, external_stylesheets=external_stylesheets)


data=np.array([["Exito",2,3],["Cencosub",4,5]])

df3=pd.DataFrame(data,index=["Exito","Cencosub"],columns=["Cliente","ZVMI","FIRME"])


def Create_Table(Table,Name_id):
    """
    - Table: Dataframe referencia a tabla
    - Name_id: Id de la tabla
    - Alias_libreria: Libreria Dash dash_table as dt
    """ 
    Tabla=dt.DataTable(
            id=Name_id,
            columns=[{"name": i, "id": i} for i in Table.columns],
            data=Table.to_dict("records"),
            style_data={
            'fontSize':'11px'
            },
            style_table={
                'margin': '0 auto',
                'border': '1px solid black',
                'borderCollapse': 'collapse'
            },
            style_header={
                'fontSize':'11px',
                'backgroundColor': '#4074D5',
                'fontWeight': 'bold',
                'border': '1px solid black'
            },
            style_cell={
                'textAlign': 'center',
                'border': '1px solid black',
                'padding': '5px',
                'width': '20px'
            },
            )
    return(Tabla)


def Create_Dropdown(Column_Table,Estados,Name_id):
        Dropdown=dcc.Dropdown(
                        id=Name_id,
                        options=[{"label": st, "value": st} for st in Estados],
                        placeholder="-Select a State-",
                        multi=True,
                        value=Column_Table.unique())
        return(Dropdown)


def estructura(data):
    return html.Div([
            html.H3('Tab content 1'),
            dcc.Graph(
                figure={
                    'data': data
                }
            )
        ])

lista=df3['Cliente'].unique().tolist()
print(lista)
app.layout = html.Div([
    html.H1('Dash Tabs component demo'),
    dcc.Tabs(id="tabs-example-graph", value='tab-1-example-graph', children=[
        dcc.Tab(label='Tab One', value='tab-1-example-graph', children=[
                Create_Dropdown(df3['Cliente'],lista,"filter_dropdown_Clase_Pedido_CEN"),
                Create_Table(df3,"Tabla_Agenda_Exito"),          
    ]),
        dcc.Tab(label='Tab Two', value='tab-2-example-graph'),
    ]),
])

@callback(Output("Tabla_Agenda_Exito", "data"),
          Input('tabs-example-graph', 'value'),
          Input("filter_dropdown_Clase_Pedido_CEN", "value"),)



def render_content(tab,estados):
    if tab == 'tab-1-example-graph':
        dff2 = df3[df3['Cliente'].isin(estados)]
        return (dff2.to_dict("records"))
    elif tab == 'tab-2-example-graph':
        dff2 = df3[df3['Cliente'].isin(estados)]
        return (dff2.to_dict("records"))
    

app.run(host='0.0.0.0', port=8000, debug=False)

