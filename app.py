import base64
import datetime
import io

import dash
import dash_table
import pandas as pd
import dash_core_components as dcc
import dash_html_components as html

from dash.dependencies import Input, Output, State
from tplreader import get_tpls




external_stylesheets = ['https://codepen.io/chriddyp/pen/bWLwgP.css']

app = dash.Dash(__name__, external_stylesheets=external_stylesheets)

upload_style = {
    'width': '100%',
    'height': '60px',
    'lineHeight': '60px',
    'borderWidth': '1px',
    'borderStyle': 'dashed',
    'borderRadius': '5px',
    'textAlign': 'center',
    'margin': '10px'
}

table_titles = ["模板", "下载配置文件", "上传修改配置文件", "生成报告", "下载报告"]
table_data = []
for filename in get_tpls("outputs"):
    table_data.append([filename, "下载配置文件", "上传修改配置文件", "生成报告", "下载报告"])
df = pd.DataFrame(table_data, columns = table_titles) 

app.layout = html.Div([
    dcc.Upload(
        id='upload-data',
        children=html.Div([
            'Word 模板设置',
            html.A('Select Files')
        ]),
        style=upload_style,
        multiple=True
    ),
    dash_table.DataTable(
        id="tlp_table",
        data=df.to_dict('records'),
        columns=[{'name': i, 'id': i} for i in df.columns]
    )
])


@app.callback(Output('tlp_table', 'data'),
              Input('upload-data', 'contents'),
              State('upload-data', 'filename'),
              State('upload-data', 'last_modified'))
def update_output(list_of_contents, list_of_names, list_of_dates):
    if list_of_contents is not None:
        for i, contents in enumerate(list_of_contents):
            content_type, content_string = contents.split(',')
            try:
                with open(f"outputs/{list_of_names[i]}", "wb") as outfile:
                    decoded = base64.b64decode(content_string)
                    outfile.write(io.BytesIO(decoded).getbuffer())
                    outfile.write(contents)
            except Exception as e:
                print(e)

        table_data = []
        for filename in get_tpls("outputs"):
            table_data.append([filename, "下载配置文件", "上传修改配置文件", "生成报告", "下载报告"])
        df = pd.DataFrame(table_data, columns = table_titles) 
        return df.to_dict('records')
    else:
        table_data = []
        for filename in get_tpls("outputs"):
            table_data.append([filename, "下载配置文件", "上传修改配置文件", "生成报告", "下载报告"])
        df = pd.DataFrame(table_data, columns = table_titles) 
        return df.to_dict('records')



if __name__ == '__main__':
    app.run_server(debug=True)