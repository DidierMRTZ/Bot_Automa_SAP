from dash import Dash, dcc, html
from dash.dependencies import Input, Output
from dash import dash_table as dt
from dash import dcc
from dash import html

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
    """
    - Column_Table: Columnas de estatos de interes
    - Estados: Los estados unicos
    - Name_id: Nombre del container dropdown 
    """   
    Dropdown=dcc.Dropdown(
                    id=Name_id,
                    options=[{"label": st, "value": st} for st in Estados],
                    placeholder="-Select a State-",
                    multi=True,
                    value=Column_Table.unique())
    return(Dropdown)
