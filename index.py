from flask import Flask, send_file
from subprocess import Popen

app = Flask(__name__)

@app.route('/')
def llamar_archivo():
    # Lógica para ejecutar el archivo .py
    p = Popen(['python', 'XmlFacturaFasecolda3 (1).py'])
    p.wait()
    
    # Devolver el archivo HTML directamente
    return send_file('index.html')

if __name__ == '__main__':
    app.run()
