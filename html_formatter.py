import os
STYLE = os.path.dirname(os.path.abspath(__file__)) + '/style.css'
def prepare_body(text):
    body = """
    <html>
    <head>
        <meta  content="png"/>
        <meta  content="Landscape"/>
        <meta charset="UTF-8">
        <link rel="stylesheet" href="{}">
    </head>
    <table>
        {}
    </table>
    </html>
    """
    return body.format(STYLE,text)
