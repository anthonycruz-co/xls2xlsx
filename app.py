import os
from time import sleep
import subprocess

from flask import Flask
from flask import request
from flask import send_file, make_response


app = Flask(__name__)


@app.route('/')
def hello_world():
    return 'Hello World!'

@app.route('/upload', methods=['POST'])
def upload_file():
    path = os.environ.get('FLASK_TMP_DIR', '/tmp')
    if request.method == 'POST':
        f = request.files['file']
        filename = f.filename
        fullpath = os.path.join(path, filename)
        f.save(fullpath)
        new_filename = xls2xlsx(path, filename)
        try:
            if new_filename:
                print "Making response object"
                fullpath = os.path.join(path, new_filename)
                with open(fullpath, 'rb') as fp:
                    response = make_response(fp.read())
                    response.headers['Content-Type'] = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                    response.headers['Content-Disposition'] = 'attachment; filename={0}'.format(new_filename)
                    print response
                    return response
        except Exception as e:
            return str(e)
    return "Error"


def xls2xlsx(path, filename):
    fullpath = os.path.join(path, filename)
    if os.path.isfile(fullpath):
        path = path.replace(filename, '')
        cmd = u"libreoffice --convert-to xlsx --headless --outdir {0} {1}".format(
            path,
            fullpath
        )
        new_filename = filename.replace('.xls', '.xlsx')
        fullpath = os.path.join(path, new_filename)
        output = subprocess.call(cmd.split())
        sleep(int(os.environ.get('XLSX_COMMAND_WAIT', 1)))
        if output == 0 and os.path.isfile(fullpath):
            # fullpath = os.path.join(path, new_filename)
            return new_filename

    return False


if __name__ == '__main__':
    port = int(os.environ.get('FLASK_PORT', 5000))
    host = os.environ.get('FLASK_HOST_ADDRESS', '127.0.0.1')
    app.run(host=host, port=port)
