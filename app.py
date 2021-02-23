import base64
import datetime
import io

import dash
import dash_table
import pandas as pd
import dash_core_components as dcc
import dash_html_components as html

from dash.dependencies import Input, Output, State
from tplreader import get_tpls, get_tpls_status, tpl_reader


external_stylesheets = ['https://codepen.io/chriddyp/pen/bWLwgP.css']

app = dash.Dash(__name__, external_stylesheets=external_stylesheets)

upload_style = {
    'width': '98%',
    'height': '60px',
    'lineHeight': '60px',
    'borderWidth': '1px',
    'borderStyle': 'dashed',
    'borderRadius': '5px',
    'textAlign': 'center',
    'margin': '10px',
    'columnCount': 4
}

table_titles = ["模板", "下载配置文件", "上传修改配置文件", "生成报告", "下载报告"]
table_data = []
for filename in get_tpls("outputs"):
    table_data.append([filename, "下载配置文件", "上传修改配置文件", "生成报告", "下载报告"])
df = pd.DataFrame(table_data, columns = table_titles) 

app.layout = html.Div([
    html.H6("智能报告生成器", style={"width":"100%","textAlign":"center"}),
    html.Div([
    dcc.Upload(
        id='upload-data',
        children=html.Div([
            "已上传",
            html.Button('更新Word模板')
        ])
    ),
    html.Div(
        id="tpl_config",
        children=[]
    ),
    dcc.Upload(
        id='upload-config',
        children=html.Div([
            html.Button('上传配置-生成报告')
        ])
    ),
    html.Button(
        id="gen_report",
        children=html.A("下载报告"),
    )
    ], style=upload_style),
    html.Div(id="report_container",style=upload_style)
])


@app.callback(Output('tpl_config', 'children'),
              Input('upload-data', 'contents'))
def update_output(contents):
    if contents is not None:
        content_type, content_string = contents.split(',')
        try:
            with open(f"outputs/tpl.docx", "wb") as outfile:
                decoded = base64.b64decode(content_string)
                outfile.write(io.BytesIO(decoded).getbuffer())
                outfile.write(contents)
        except Exception as e:
            print(e)

    tpl_status, config_status = get_tpls_status("outputs")
    if config_status:
        return [
            "已生成",
            html.A('下载配置文件', href="/assets/config.xlsx",target="blank")
        ]
    else:
        return [
            "请先上传Word模板"
        ]
    

def gen_report():
    config = tpl_reader("outputs/tpl.docx")

    return config.get("tables", [])

@app.callback(Output('report_container', 'children'),
              Input('upload-config', 'contents'))
def update_config(contents):
    if contents is not None:
        content_type, content_string = contents.split(',')
        try:
            with open(f"outputs/tpl.docx.xlsx", "wb") as outfile:
                decoded = base64.b64decode(content_string)
                outfile.write(io.BytesIO(decoded).getbuffer())
                outfile.write(contents)
        except Exception as e:
            print(e)
        
    return gen_report()

if __name__ == '__main__':
    app.run_server(debug=True)