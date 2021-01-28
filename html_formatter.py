
def prepare_body(text):
    body = """
    <html>
    <head>
        <meta  content="png"/>
        <meta  content="Landscape"/>
        <meta charset="UTF-8">
        <link rel="stylesheet" href="/home/sharofiddin/practice/Python/style.css">
    </head>
    <table>
        {}
    </table>
    </html>
    """
    return body.format(text)
