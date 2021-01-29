import json
with open('config.json', 'r') as f:
    config = json.load(f)
STYLE = config['style']
def prepare_body(text):
    body = """
    <html>
    <head>
        <meta  content="png"/>
        <meta  content="Landscape"/>
        <meta charset="UTF-8">
        <link rel="stylesheet" href="{}" type="text/css">
    </head>
    <table>
        {}
    </table>
    </html>
    """
    return body.format(STYLE,text)
