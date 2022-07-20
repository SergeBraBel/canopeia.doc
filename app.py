from flask import Flask, render_template, url_for, request,  send_from_directory
import MTCOrderExcel, UPSTrack

app = Flask(__name__)
my_dict = {}
UPLOAD_FOLDER = r'C:\Users\sdemi\PycharmProjects\flaskProject2\uploads'

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        my_dict[request.form.get('id')] = request.form.get('message')
        order = MTCOrderExcel.ExelFileCreator(request.form.get('nfNumber'), request.form.get('orderNumber'))
        order.create()
        excel_file_name = f"{request.form.get('orderNumber')}_pedido(nf{request.form.get('nfNumber')}).xlsx"
        my_dict['path'] = f"{request.form.get('orderNumber')}_pedido(nf{request.form.get('nfNumber')}).xlsx"
        return render_template('index.html', message=my_dict, file_to_download=excel_file_name, title='DocForm')
    else:
        return render_template('index.html', message=my_dict, title='ExcelForm')



@app.route('/lable', methods=['GET', 'POST'])
def lable():
    if request.method == 'POST':
        #print(type('order_number'))
        order_number = int(request.form.get('order_number'))
        height = int(request.form.get('height'))
        width = int(request.form.get('width'))
        length = int(request.form.get('length'))
        total_gross_weight = float(request.form.get('total_gross_weight'))
        exchange_rate = float(request.form.get('exchange_rate'))
        freight = int(request.form.get('freight'))
        sel1 = (request.form.get('sel1'))
        if (request.form.get('sel1')) == "real":
            flag = True
        else:
            flag = False
        lableFileName = f"{str(order_number)}_lable.gif"
        invoice = f"{str(order_number)}_invoice.docx"
        PL = f"{str(order_number)}_PL.docx"

        lable = UPSTrack.UPS(order_number, height, width, length, total_gross_weight,exchange_rate,freight,flag)
        lable.create()
        trackNum = lable.track
        my_dict['lable_path'] = lableFileName
        print(lableFileName)
        return render_template('lable.html', trackNum=trackNum, file_to_download=lableFileName, invoice=invoice, PL=PL, title='lable', order_number=order_number, height=height, width=width, length=length, total_gross_weight=total_gross_weight, exchange_rate=exchange_rate, freight=freight, sel1=sel1, flag=flag )
    else:
        return render_template('lable.html', title='lable')

"""@app.route('/uploads/<path:filename>', methods=['GET', 'POST'])
def download(filename):
    return send_from_directory(UPLOAD_FOLDER, filename, as_attachment=True)"""

@app.route('/uploads/<path:filename>', methods=['GET', 'POST'])
def download(filename):
    return send_from_directory(UPLOAD_FOLDER, filename, as_attachment=True)


if __name__ == '__main__':
    app.run(debug=True)


