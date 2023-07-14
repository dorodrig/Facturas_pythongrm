from flask import Flask, send_file
import execnet


app = Flask(__name__)

@app.route('/')
def llamar_archivo():
    resultado = ejecutar_codigo()
    return send_file('index.html')

def ejecutar_codigo():
    with open('XmlFacturaFasecolda3.py', 'r') as archivo:
        codigo = archivo.read()

    namespace = {}
    exec(codigo, namespace)
    
    return namespace

if __name__ == '__main__':
    app.run()
