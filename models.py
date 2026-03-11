from flask_sqlalchemy import SQLAlchemy
from datetime import datetime
import secrets

db = SQLAlchemy()

class Admin(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(50), unique=True, nullable=False)
    password_hash = db.Column(db.String(200), nullable=False)  # хранить хеш!

class Employee(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    first_name = db.Column(db.String(50), nullable=False)
    last_name = db.Column(db.String(50), nullable=False)
    middle_name = db.Column(db.String(50))
    phone = db.Column(db.String(20), nullable=False)
    token = db.Column(db.String(50), unique=True, nullable=False, default=lambda: secrets.token_urlsafe(32))
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    intervals = db.relationship('WorkInterval', backref='employee', lazy=True, cascade='all, delete-orphan')
    adjustments = db.relationship('DailyAdjustment', back_populates='employee', cascade='all, delete-orphan')

class WorkInterval(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    employee_id = db.Column(db.Integer, db.ForeignKey('employee.id'), nullable=False)
    start_time = db.Column(db.DateTime, nullable=False, default=datetime.utcnow)
    end_time = db.Column(db.DateTime, nullable=True)
    type = db.Column(db.Enum('work', 'pause', name='interval_types'), nullable=False)

class DailyAdjustment(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    employee_id = db.Column(db.Integer, db.ForeignKey('employee.id'), nullable=False)
    date = db.Column(db.Date, nullable=False)               # дата корректировки
    delta_minutes = db.Column(db.Integer, nullable=False)    # изменение в минутах (может быть отрицательным)
    comment = db.Column(db.Text, nullable=True)              # комментарий
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    employee = db.relationship('Employee', back_populates='adjustments')