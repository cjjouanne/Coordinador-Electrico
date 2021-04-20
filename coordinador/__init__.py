from flask import Flask, render_template
from flask_sqlalchemy import SQLAlchemy
from flask_bcrypt import Bcrypt
from flask_login import LoginManager


app = Flask(__name__)

app.config['SECRET_KEY'] = 'SECRET_KEY'
app.config['SQLALCHEMY_DATABASE_URI'] = 'postgresql://myuser:mypassword@localhost:5432/coordinador'

db = SQLAlchemy(app)

bcrypt = Bcrypt(app)
login_manager = LoginManager(app)

from coordinador import routes
