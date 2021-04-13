from coordinador import app, mail, db, bcrypt
from coordinador.forms import RegistrationForm, LoginForm, SubscribeForm
from coordinador.models import User, Subscription
from flask import render_template, jsonify, redirect, url_for
from flask_mail import Message
from flask_login import login_user, current_user, logout_user
from datetime import date
import requests, zipfile, os
from openpyxl import load_workbook

# get file url
def get_url():
    myDate = date.today()
    year = myDate.year
    month1 = '0{month}'.format(month=myDate.month) if myDate.month < 10 else myDate.month
    if int(myDate.day) == 1:
        month1 = '0{month}'.format(month=myDate.month - 1) if myDate.month - 1 < 10 else myDate.month - 1
    month2 = '0{month}'.format(month=myDate.month) if myDate.month < 10 else myDate.month
    day = '0{day}'.format(day=myDate.day) if myDate.day < 10 else myDate.day
    return 'https://www.coordinador.cl/wp-content/uploads/{year}/{month1}/PROGRAMA{year}{month2}{day}.zip'.format(year = year, month1 = month1, month2 = month2, day = day)

# download zip file from coordinators site
def download_url(url, save_path, chunk_size=128):
    r = requests.get(url, stream=True)
    with open(save_path, 'wb') as fd:
        for chunk in r.iter_content(chunk_size=chunk_size):
            fd.write(chunk)

# extact content from zip file
def unzip_file(path):
    with zipfile.ZipFile(path, 'r') as zip_ref:
        zip_ref.extractall('files')

# remove unnecessary files
def remove_files(path):
    for f in os.listdir(path):
        os.remove(os.path.join(path, f))

# gets data from excel
def extract_data(path):
    subscriptions = {s.name for s in Subscription.query.all()}
    ps = load_workbook(path)
    sheet = ps['PROGRAMA']
    data = {}
    for row in range(2, sheet.max_row + 1):
        if str(sheet['C'+ str(row)].value) in subscriptions:
            print('ROW: ', row)
            print('VALUE: ', str(sheet['C'+ str(row)].value))
            data[str(sheet['C'+ str(row)].value)] = {
                '1': str(sheet['E'+ str(row)].value),
                '2': str(sheet['F'+ str(row)].value),
                '3': str(sheet['G'+ str(row)].value),
                '4': str(sheet['H'+ str(row)].value),
                '5': str(sheet['I'+ str(row)].value),
                '6': str(sheet['J'+ str(row)].value),
                '7': str(sheet['K'+ str(row)].value),
                '8': str(sheet['L'+ str(row)].value),
                '9': str(sheet['M'+ str(row)].value),
                '10': str(sheet['N'+ str(row)].value),
                '11': str(sheet['O'+ str(row)].value),
                '12': str(sheet['P'+ str(row)].value),
                '13': str(sheet['Q'+ str(row)].value),
                '14': str(sheet['R'+ str(row)].value),
                '15': str(sheet['S'+ str(row)].value),
                '16': str(sheet['T'+ str(row)].value),
                '17': str(sheet['U'+ str(row)].value),
                '18': str(sheet['V'+ str(row)].value),
                '19': str(sheet['W'+ str(row)].value),
                '20': str(sheet['X'+ str(row)].value),
                '21': str(sheet['Y'+ str(row)].value),
                '22': str(sheet['Z'+ str(row)].value),
                '23': str(sheet['AA'+ str(row)].value),
                '24': str(sheet['AB'+ str(row)].value),
                'Total': str(sheet['AC'+ str(row)].value),
                }
    return data

# creates registration form errors dictionary
def get_errors(form):
    errors = {}
    if form.username.errors:
        errors['username'] = form.username.errors
    if form.email.errors:
        errors['email'] = form.email.errors
    if form.confirm_password.errors:
        errors['confirm_password'] = form.confirm_password.errors
    return errors

@app.route('/')
def hello_world():
    msg = Message("[Coordinador Electrico]",
                  sender="c.jouanne.g@gmail.com",
                  recipients=["c.jouanne.g@gmail.com"])
    msg.html = render_template('mail.html')
    mail.connect()
    mail.send(msg)

    return '<h1>Hello, World!<h1>'


@app.route('/get_report')
def get_report():
    url = get_url()
    download_url(url,"./files/coordinador.zip")
    unzip_file('./files/coordinador.zip')
    myDate = date.today()
    year = myDate.year
    month = '0{month}'.format(month=myDate.month) if myDate.month < 10 else myDate.month
    day = '0{day}'.format(day=myDate.day) if myDate.day < 10 else myDate.day
    path = './files/PRG{year}{month}{day}.xlsx'.format(year = str(year)[2:], month = month, day = day)
    data = extract_data(path)
    print(data)
    users = User.query.all()
    for u in users:
        subscriptions = u.subscriptions
        if subscriptions and len(subscriptions) > 0:
            msg = Message("[Coordinador Electrico]",
                          sender="c.jouanne.g@gmail.com",
                          recipients=[u.email])
            msg.html = render_template('mail.html', user=u, data=data, subscriptions=subscriptions)
            mail.connect()
            mail.send(msg)
    remove_files('./files')
    return 'about router'

@app.route('/register', methods=['GET', 'POST'])
def register():
    if current_user.is_authenticated:
        return redirect(url_for('hello_world'))
    form = RegistrationForm()
    if form.validate_on_submit():
        hashed_password = bcrypt.generate_password_hash(form.password.data).decode('utf-8')
        user = User(username=form.username.data, email=form.email.data, password=hashed_password)
        db.session.add(user)
        db.session.commit()
        response = { 'status': 200 }
        return jsonify(response)
    return render_template('register.html', form=form)

@app.route('/login', methods=['GET', 'POST'])
def login():
    if current_user.is_authenticated:
        return redirect(url_for('hello_world'))
    form = LoginForm()
    if form.validate_on_submit():
        user = User.query.filter_by(email=form.email.data).first()
        if user and bcrypt.check_password_hash(user.password, form.password.data):
            login_user(user, remember=form.remember.data)
            response = { 'status': 200, 'currentUser': str(current_user) }
            return jsonify(response)
        return jsonify({ 'status': 400, 'error': 'Login Unsuccessful. Please check Email and Password' })
    return render_template('login.html', form=form)

@app.route('/logout')
def logout():
    logout_user()
    return redirect(url_for('get_report'))

@app.route('/subscribe', methods=['GET', 'POST'])
def subscribe():
    if current_user.is_authenticated:
        form = SubscribeForm()
        if form.validate_on_submit():
            subscription = Subscription(name=form.name.data, user_id=current_user.id)
            db.session.add(subscription)
            db.session.commit()
            response = { 'status': 200, 'subscriptions': str(current_user.subscriptions) }
            return jsonify(response)
        print(form.hidden_tag())
        return render_template('subscribe.html', form=form)
    return jsonify({ 'error': 'user not authenticated'})
