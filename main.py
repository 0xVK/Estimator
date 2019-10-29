import os
import random
import string
import pdfkit
import pyexcel
from flask import Flask, render_template, request, render_template_string, Response

app = Flask(__name__, template_folder='site', static_folder='site/static', static_url_path='/static')
available_objects = ('Tables', 'Page', 'Reports', 'Codeunits', 'Dataports', 'XMLPorts', 'Queries', 'Forms', 'Menus')


@app.route('/pdf/<name>')
def get_output_file(name):

    base_dir = os.path.dirname(os.path.abspath(__file__))
    pdf_dir = os.path.join(base_dir, 'pdf')
    file_name = os.path.join(pdf_dir, name)

    if not os.path.isfile(file_name):
        return 'Bad request'

    with open(file_name, 'rb') as f:
        resp = Response(f.read())

    resp.headers["Content-Disposition"] = "attachment; filename={0}".format(name)
    resp.headers["Content-type"] = "application/pdf"
    return resp


def save_to_pdf(context={}):

    base_dir = os.path.dirname(os.path.abspath(__file__))
    pdf_dir = os.path.join(base_dir, 'pdf')

    if not os.path.exists(pdf_dir):
        os.makedirs(pdf_dir)

    while (True):
        rand_name = ''.join(random.choice(string.ascii_lowercase +
                                          string.ascii_uppercase +
                                          string.digits)
                            for x in range(4))
        rand_name = rand_name + '_estimate.pdf'
        if not os.path.exists(os.path.join(pdf_dir, rand_name)):
            break

    rendered_html = render_template('pdf_template.html', data=context)
    pdfkit.from_string(rendered_html, os.path.join(pdf_dir, rand_name))

    return rand_name


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/ajax_estimate', methods=['POST'])
def ajax_estimate():

    def validate(n=''):
        return int(n) if n.isdigit() else 0

    data = {'total_time': 0}

    for available_object in available_objects:
        data[available_object] = {}

    data['Tables']['from'] = validate(request.form.get('from_tables'))
    data['Tables']['to'] = validate(request.form.get('to_tables'))
    data['Page']['from'] = validate(request.form.get('from_pages'))
    data['Page']['to'] = validate(request.form.get('to_pages'))
    data['Reports']['from'] = validate(request.form.get('from_reports'))
    data['Reports']['to'] = validate(request.form.get('from_reports'))
    data['Codeunits']['from'] = validate(request.form.get('from_codeunits'))
    data['Codeunits']['to'] = validate(request.form.get('to_codeunits'))
    data['Dataports']['from'] = validate(request.form.get('from_dataports'))
    data['Dataports']['to'] = validate(request.form.get('to_dataports'))
    data['XMLPorts']['from'] = validate(request.form.get('from_dataports'))
    data['XMLPorts']['to'] = validate(request.form.get('to_dataports'))
    data['Queries']['from'] = validate(request.form.get('from_queries'))
    data['Queries']['to'] = validate(request.form.get('to_queries'))
    data['Forms']['from'] = validate(request.form.get('from_forms'))
    data['Forms']['to'] = validate(request.form.get('to_forms'))
    data['Menus']['from'] = validate(request.form.get('from_menus'))
    data['Menus']['to'] = validate(request.form.get('to_menus'))

    data['fname'] = request.form.get('fname')
    data['lname'] = request.form.get('lname')
    data['email'] = request.form.get('email')
    data['phone'] = request.form.get('phone')
    data['notes'] = request.form.get('notes')

    calculated_data = calculate(read_xls('Estimates_Statistics.xlsx'))

    for available_object in available_objects:
        total_object_time = get_time_to_object(available_object,
                                               calculated_data,
                                               data[available_object].get('from'),
                                               data[available_object].get('to'))
        data[available_object]['total_time'] = total_object_time[1]
        data['total_time'] += total_object_time[0]

    data['file_name'] = 'pdf/' + save_to_pdf(data)
    
    return render_template('ajax_estimate.html', data=data)


@app.route('/estimate', methods=['POST', 'GET'])
def estimate():
    
	# TODO видалити цю ф-ю. 
    def validate(n=''):
        return int(n) if n.isdigit() else 0

    data = {'total_time': 0}

    for available_object in available_objects:
        data[available_object] = {}

    data['Tables']['from'] = validate(request.form.get('from_tables'))
    data['Tables']['to'] = validate(request.form.get('to_tables'))
    data['Page']['from'] = validate(request.form.get('from_pages'))
    data['Page']['to'] = validate(request.form.get('to_pages'))
    data['Reports']['from'] = validate(request.form.get('from_reports'))
    data['Reports']['to'] = validate(request.form.get('from_reports'))
    data['Codeunits']['from'] = validate(request.form.get('from_codeunits'))
    data['Codeunits']['to'] = validate(request.form.get('to_codeunits'))
    data['Dataports']['from'] = validate(request.form.get('from_dataports'))
    data['Dataports']['to'] = validate(request.form.get('to_dataports'))
    data['XMLPorts']['from'] = validate(request.form.get('from_dataports'))
    data['XMLPorts']['to'] = validate(request.form.get('to_dataports'))
    data['Queries']['from'] = validate(request.form.get('from_queries'))
    data['Queries']['to'] = validate(request.form.get('to_queries'))
    data['Forms']['from'] = validate(request.form.get('from_forms'))
    data['Forms']['to'] = validate(request.form.get('to_forms'))
    data['Menus']['from'] = validate(request.form.get('from_menus'))
    data['Menus']['to'] = validate(request.form.get('to_menus'))

    data['fname'] = request.form.get('fname')
    data['lname'] = request.form.get('lname')
    data['email'] = request.form.get('email')
    data['phone'] = request.form.get('phone')
    data['notes'] = request.form.get('notes')

    calculated_data = calculate(read_xls('Estimates_Statistics.xlsx'))

    for available_object in available_objects:
        total_object_time = get_time_to_object(available_object,
                                               calculated_data,
                                               data[available_object].get('from'),
                                               data[available_object].get('to'))
        data[available_object]['total_time'] = total_object_time[1]
        data['total_time'] += total_object_time[0]

    data['file_name'] = 'pdf/' + save_to_pdf(data)
    return render_template('estimate.html', data=data)


def read_xls(filename_xls=''):

    table = pyexcel.get_array(file_name=filename_xls)

    data = {}

    for available_object in available_objects:
        data[available_object] = []

    for row in table[3:]:
        name = row[0]
        row_dict = {
                'x1': row[1],
                'x2': row[2],
                'y11': row[4],
                'y12': row[5],
                'y21': row[7],
                'y22': row[8],
            }
        if name in available_objects and row[1] and row[2] and row_dict not in data[name]:
            data[name].append(row_dict)

    return data


def calculate(data={}):
    calculated_data = {}

    for available_object in available_objects:
        calculated_data[available_object] = {
            'sum_x1': 0,
            'sum_x2': 0,
            'sum_y11': 0,
            'sum_y12': 0,
            'sum_y21': 0,
            'sum_y22': 0,
            'sqr_x1': 0,
            'sqr_x2': 0,
            'mult_x1y11': 0,
            'mult_x1y12': 0,
            'mult_x2y21': 0,
            'mult_x2y22': 0,
        }

    for object_name in available_objects:
        calculated_data[object_name]['count'] = len(data[object_name])
        for row in data[object_name]:
            calculated_data[object_name]['sum_x1'] += row['x1']
            calculated_data[object_name]['sum_x2'] += row['x2']
            calculated_data[object_name]['sum_y11'] += row['y11']
            calculated_data[object_name]['sum_y12'] += row['y12']
            calculated_data[object_name]['sum_y21'] += row['y21']
            calculated_data[object_name]['sum_y22'] += row['y22']

            calculated_data[object_name]['sqr_x1'] += row['x1'] * row['x1']
            calculated_data[object_name]['sqr_x2'] += row['x2'] * row['x2']

            calculated_data[object_name]['mult_x1y11'] += row['x1'] * row['y11']
            calculated_data[object_name]['mult_x1y12'] += row['x1'] * row['y12']
            calculated_data[object_name]['mult_x2y21'] += row['x2'] * row['y21']
            calculated_data[object_name]['mult_x2y22'] += row['x2'] * row['y22']

    return calculated_data


def get_time_to_object(object_name='', calculated_data={}, new_obj=0, mod_obj=0):

    def fun(a=0, b=0, c=0, d=0, n=0):

        try:
            l = (b / a - d / c) / (n / a - a / c)
            m = (b / a - n / a * l)
        except ZeroDivisionError:
            l = m = 0

        return l, m

    l, m = fun(calculated_data[object_name]['sum_x1'], calculated_data[object_name]['sum_y11'],
               calculated_data[object_name]['sqr_x1'], calculated_data[object_name]['mult_x1y11'],
               calculated_data[object_name]['count'])

    time_to_create_dev = m * new_obj + l if new_obj else 0

    l, m = fun(calculated_data[object_name]['sum_x1'], calculated_data[object_name]['sum_y12'],
               calculated_data[object_name]['sqr_x1'], calculated_data[object_name]['mult_x1y12'],
               calculated_data[object_name]['count'])

    time_to_create_qa = m * new_obj + l if new_obj else 0

    l, m = fun(calculated_data[object_name]['sum_x2'], calculated_data[object_name]['sum_y21'],
               calculated_data[object_name]['sqr_x2'], calculated_data[object_name]['mult_x2y21'],
               calculated_data[object_name]['count'])

    time_to_mod_dev = m * mod_obj + l if mod_obj else 0

    l, m = fun(calculated_data[object_name]['sum_x2'], calculated_data[object_name]['sum_y22'],
               calculated_data[object_name]['sqr_x2'], calculated_data[object_name]['mult_x2y22'],
               calculated_data[object_name]['count'])

    time_to_mod_qa = m * mod_obj + l if mod_obj else 0

    total_time = time_to_create_dev + time_to_create_qa + time_to_mod_dev + time_to_mod_qa

    return round(total_time, 2), '{} - {}'.format(round(total_time * 0.8, 2), round(total_time * 1.2, 2))


if __name__ == "__main__":
    app.run(debug=True, port='8080')
    #app.run(debug=True)

